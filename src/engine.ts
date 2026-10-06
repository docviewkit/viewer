import { assertInputSize, Core, type CoreDocument, encodePdfPassword, resolveLimits } from "./core.js";
import { copyByteSource, isArrayBuffer, isUint8Array } from "./bytes.js";
import { SceneRenderer, sheetFormulaViewportBounds } from "./render.js";
import { fontMetricLayoutObjects, measureSceneFontMetrics } from "./font-metrics.js";
import {
  copyPreparedFonts,
  collectFontRequests,
  documentFontRuns,
  FontProviderCache,
  fontAssetBytes,
  prepareFontAssets,
  prepareEmbeddedFontAssets,
  registerFonts,
  verifyFontAssetIntegrity,
  type PreparedFontAsset,
  type RegisteredFonts,
} from "./font.js";
import { documentSpaceBounds, sceneTextLayout, type SceneObject } from "./scene.js";
import type {
  Diagnostic,
  DocumentObject,
  EngineOptions,
  HitResult,
  HitTestRequest,
  ObjectListRequest,
  OfficeDocument,
  OfficeEngine,
  OperationOptions,
  OpenOptions,
  RenderOptions,
  RenderRequest,
  RenderResult,
  TextSearchRequest,
  TextSearchResult,
} from "./types.js";
import {
  immutableDiagnostics,
  immutableDocumentInfo,
  immutableDocumentObject,
  immutableHitResults,
  immutableRenderResult,
  listDocumentObjects,
  OfficeEngineError,
  searchDocumentText,
  validateRenderOptions,
  validateOperationOptions,
} from "./types.js";
import { WorkerEngine } from "./worker-engine.js";
import { calculateSpreadsheet } from "./core.js";

function snapshotEngineOptions(options: EngineOptions): EngineOptions {
  const wasm = isArrayBuffer(options.wasm) || isUint8Array(options.wasm)
    ? copyByteSource(options.wasm)
    : options.wasm instanceof URL
      ? new URL(options.wasm.href)
      : options.wasm;
  const workerUrl = options.workerUrl instanceof URL
    ? new URL(options.workerUrl.href)
    : options.workerUrl;
  const calculationWasm = isArrayBuffer(options.calculationWasm) || isUint8Array(options.calculationWasm)
    ? copyByteSource(options.calculationWasm)
    : options.calculationWasm instanceof URL
      ? new URL(options.calculationWasm.href)
      : options.calculationWasm;
  const limits = options.limits === undefined
    ? undefined
    : Object.freeze({ ...options.limits });
  const fonts = options.fonts === undefined
    ? undefined
    : Object.freeze(options.fonts.map((font) => Object.freeze({
        ...font,
      bytes: copyByteSource(font.bytes),
    })));
  const license = options.license === undefined
    ? undefined
    : Object.freeze({ ...options.license });
  return Object.freeze({
    ...options,
    ...(wasm === undefined ? {} : { wasm }),
    ...(workerUrl === undefined ? {} : { workerUrl }),
    ...(calculationWasm === undefined ? {} : { calculationWasm }),
    ...(limits === undefined ? {} : { limits }),
    ...(fonts === undefined ? {} : { fonts }),
    ...(license === undefined ? {} : { license }),
  });
}

function documentObject(
  object: SceneObject,
  objectsByNumericId: ReadonlyMap<number, SceneObject>,
  fonts: RegisteredFonts,
  includeObjectTransform: boolean,
): DocumentObject {
  const fontRuns = documentFontRuns(object, fonts);
  return immutableDocumentObject({
    ...object,
    bounds: documentSpaceBounds(object, objectsByNumericId, includeObjectTransform),
    ...(fontRuns === undefined ? {} : { fontRuns }),
    ...(object.source.format === "xlsx" ? { wrapText: sceneTextLayout(object.visual)?.wrap ?? false } : {}),
  });
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new OfficeEngineError("OPERATION_ABORTED", "The operation was aborted", { cause: signal.reason });
  }
}

function validateInlineOperation(options: OperationOptions): void {
  validateOperationOptions(options);
  throwIfAborted(options.signal);
  if (options.timeoutMs !== undefined) {
    throw new OfficeEngineError(
      "INLINE_TIMEOUT_UNSUPPORTED",
      "timeoutMs requires Worker execution because inline document operations are synchronous",
    );
  }
}

class InlineDocument implements OfficeDocument {
  #coreDocument: CoreDocument | undefined;
  readonly #objectsByNumericId = new Map<number, SceneObject>();
  readonly #objectsById = new Map<string, DocumentObject>();
  #renderer: SceneRenderer | undefined;
  #fonts: RegisteredFonts | undefined;
  #onClose: ((document: InlineDocument) => void) | undefined;
  readonly info;
  #diagnostics: readonly Diagnostic[];
  readonly #additionalDiagnostics: readonly Diagnostic[];
  readonly #localFonts: boolean;
  readonly #unitLoads = new Map<number, Promise<void>>();
  #allLoad: Promise<void> | undefined;
  #fontDiagnosticCount = 0;
  #closed = false;

  constructor(
    coreDocument: CoreDocument,
    fonts: RegisteredFonts,
    onClose: (document: InlineDocument) => void,
    additionalDiagnostics: readonly Diagnostic[] = [],
    localFonts = false,
  ) {
    const { scene } = coreDocument;
    if (scene.info === undefined) {
      fonts.close();
      coreDocument.close();
      throw new OfficeEngineError("INVALID_DOCUMENT", "The core did not identify a supported document", {
        diagnostics: scene.diagnostics,
      });
    }
    this.#coreDocument = coreDocument;
    this.#fonts = fonts;
    this.#onClose = onClose;
    this.info = immutableDocumentInfo(scene.info);
    this.#additionalDiagnostics = additionalDiagnostics;
    this.#localFonts = localFonts;
    this.#diagnostics = [];
    this.#replaceScene(scene);
  }

  diagnostics(): readonly Diagnostic[] {
    this.#assertOpen();
    return this.#diagnostics;
  }

  async render(request: RenderRequest, options: RenderOptions = {}): Promise<RenderResult> {
    this.#assertOpen();
    validateRenderOptions(options);
    validateInlineOperation(options);
    const unit = this.info.units[request.unitIndex];
    if (unit === undefined || unit.index !== request.unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${request.unitIndex}`);
    }
    await this.#ensureUnit(
      request.unitIndex,
      sheetFormulaViewportBounds(unit, request),
    );
    return immutableRenderResult(await this.#renderer!.render(unit, request, {
      isCancelled: () => options.signal?.aborted === true,
    }));
  }

  async hitTest(request: HitTestRequest, options: OperationOptions = {}): Promise<readonly HitResult[]> {
    this.#assertOpen();
    validateInlineOperation(options);
    const unit = this.info.units[request.unitIndex];
    if (!Number.isInteger(request.unitIndex)
      || request.unitIndex < 0
      || unit === undefined
      || unit.index !== request.unitIndex) {
      throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${request.unitIndex}`);
    }
    if (!Number.isFinite(request.x) || !Number.isFinite(request.y)) {
      throw new OfficeEngineError("INVALID_HIT_TEST", "Hit-test coordinates must be finite numbers");
    }
    const limit = request.limit ?? 16;
    if (!Number.isInteger(limit) || limit < 1 || limit > 256) {
      throw new OfficeEngineError("INVALID_HIT_TEST", "Hit-test limit must be an integer from 1 through 256");
    }
    await this.#ensureUnit(request.unitIndex);
    const numericIds = this.#coreDocument!.hitTest(
      request.unitIndex,
      request.x,
      request.y,
      limit,
    );
    const results: HitResult[] = [];
    for (const numericId of numericIds) {
      const internal = this.#objectsByNumericId.get(numericId);
      if (internal === undefined) continue;
      const object = this.#objectsById.get(internal.id);
      if (object === undefined) continue;
      const ancestors: DocumentObject[] = [];
      const visited = new Set<number>([numericId]);
      let parent = internal.parentNumericId;
      while (parent !== undefined && ancestors.length < 128 && !visited.has(parent)) {
        visited.add(parent);
        const ancestor = this.#objectsByNumericId.get(parent);
        if (ancestor === undefined) break;
        const publicAncestor = this.#objectsById.get(ancestor.id);
        if (publicAncestor !== undefined) ancestors.unshift(publicAncestor);
        parent = ancestor.parentNumericId;
      }
      results.push({ object, ancestors });
    }
    return immutableHitResults(results);
  }

  async getObject(id: string, options: OperationOptions = {}): Promise<DocumentObject | undefined> {
    this.#assertOpen();
    validateInlineOperation(options);
    const object = this.#objectsById.get(id);
    if (object !== undefined) return object;
    await this.#ensureAll();
    return this.#objectsById.get(id);
  }

  async searchText(request: TextSearchRequest, options: OperationOptions = {}): Promise<readonly TextSearchResult[]> {
    this.#assertOpen();
    validateInlineOperation(options);
    if (request.unitIndex === undefined) await this.#ensureAll();
    else await this.#ensureUnit(request.unitIndex);
    return searchDocumentText(this.#objectsById.values(), this.info.units, request);
  }

  async listObjects(request: ObjectListRequest = {}, options: OperationOptions = {}): Promise<readonly DocumentObject[]> {
    this.#assertOpen();
    validateInlineOperation(options);
    if (request.unitIndex === undefined) await this.#ensureAll();
    else {
      const unit = this.info.units[request.unitIndex];
      await this.#ensureUnit(
        request.unitIndex,
        unit === undefined ? undefined : sheetFormulaViewportBounds(unit, request),
      );
    }
    return listDocumentObjects(this.#objectsById.values(), this.info.units, request);
  }

  async #ensureUnit(
    unitIndex: number,
    formulaBounds?: { readonly maxRow: number; readonly maxColumn: number },
  ): Promise<void> {
    const existing = this.#unitLoads.get(unitIndex);
    if (existing !== undefined) return existing;
    const loaded = formulaBounds === undefined
      ? this.#coreDocument!.loadUnit(unitIndex)
      : this.#coreDocument!.loadUnitRegion(unitIndex, formulaBounds.maxRow, formulaBounds.maxColumn);
    if (loaded === undefined) return;
    const { scene, replace } = loaded;
    const pending = this.#fonts!
      .addEmbeddedAssets(prepareEmbeddedFontAssets(scene.embeddedFonts))
      .then(() => this.#localFonts ? this.#fonts!.addLocalRequests(collectFontRequests(scene.objects)) : undefined)
      .then(() => {
        if (this.#closed) return;
        if (replace) this.#replaceScene(scene);
        else this.#appendScene(scene);
      })
      .finally(() => this.#unitLoads.delete(unitIndex));
    this.#unitLoads.set(unitIndex, pending);
    await pending;
  }

  async #ensureAll(): Promise<void> {
    if (this.#allLoad !== undefined) return this.#allLoad;
    const pending = (async () => {
      await Promise.all(this.#unitLoads.values());
      if (this.#closed) return;
      const scene = this.#coreDocument!.loadAll();
      if (scene === undefined) return;
      await this.#fonts!.addEmbeddedAssets(prepareEmbeddedFontAssets(scene.embeddedFonts));
      if (this.#localFonts) await this.#fonts!.addLocalRequests(collectFontRequests(scene.objects));
      if (!this.#closed) this.#replaceScene(scene);
    })().finally(() => {
      this.#allLoad = undefined;
    });
    this.#allLoad = pending;
    await pending;
  }

  #replaceScene(scene: CoreDocument["scene"]): void {
    const fonts = this.#fonts!;
    const fontDiagnostics = fonts.diagnostics();
    this.#objectsByNumericId.clear();
    this.#objectsById.clear();
    for (const object of scene.objects) this.#objectsByNumericId.set(object.numericId, object);
    for (const object of scene.objects) {
      this.#objectsById.set(
        object.id,
        documentObject(object, this.#objectsByNumericId, fonts, scene.info?.format === "odp"),
      );
    }
    if (this.#renderer === undefined) {
      this.#renderer = new SceneRenderer(
        scene.objects,
        this.#coreDocument!.limits,
        fonts,
        fonts.imageCodecFonts(),
      );
    } else {
      this.#renderer.replaceObjects(scene.objects, fonts.imageCodecFonts());
    }
    this.#fontDiagnosticCount = fontDiagnostics.length;
    this.#diagnostics = immutableDiagnostics([
      ...scene.diagnostics,
      ...this.#additionalDiagnostics,
      ...fontDiagnostics,
    ]);
  }

  #appendScene(scene: CoreDocument["scene"]): void {
    const fonts = this.#fonts!;
    const fontDiagnostics = fonts.diagnostics();
    for (const object of scene.objects) this.#objectsByNumericId.set(object.numericId, object);
    for (const object of scene.objects) {
      this.#objectsById.set(
        object.id,
        documentObject(object, this.#objectsByNumericId, fonts, this.info.format === "odp"),
      );
    }
    this.#renderer!.appendObjects(scene.objects);
    this.#diagnostics = immutableDiagnostics([
      ...this.#diagnostics,
      ...scene.diagnostics,
      ...fontDiagnostics.slice(this.#fontDiagnosticCount),
    ]);
    this.#fontDiagnosticCount = fontDiagnostics.length;
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    const renderer = this.#renderer;
    const fonts = this.#fonts;
    const coreDocument = this.#coreDocument;
    const onClose = this.#onClose;
    this.#renderer = undefined;
    this.#fonts = undefined;
    this.#coreDocument = undefined;
    this.#onClose = undefined;
    this.#objectsByNumericId.clear();
    this.#objectsById.clear();
    renderer?.close();
    fonts?.close();
    coreDocument?.close();
    onClose?.(this);
  }

  #assertOpen(): void {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
  }
}

class InlineEngine implements OfficeEngine {
  #core: Core | undefined;
  #fonts: readonly PreparedFontAsset[];
  readonly #fontProvider: FontProviderCache | undefined;
  readonly #fontPolicy: "deterministic" | "local-first";
  readonly #calculationWasm: EngineOptions["calculationWasm"];
  readonly #documents = new Set<InlineDocument>();
  #closed = false;

  constructor(
    core: Core,
    fonts: readonly PreparedFontAsset[],
    fontProvider: FontProviderCache | undefined,
    fontPolicy: "deterministic" | "local-first",
    calculationWasm: EngineOptions["calculationWasm"],
  ) {
    this.#core = core;
    this.#fonts = fonts;
    this.#fontProvider = fontProvider;
    this.#fontPolicy = fontPolicy;
    this.#calculationWasm = calculationWasm;
  }

  async open(bytes: ArrayBuffer | Uint8Array, options: OpenOptions = {}): Promise<OfficeDocument> {
    this.#assertOpen();
    throwIfAborted(options.signal);
    if (options.timeoutMs !== undefined) {
      if (!Number.isFinite(options.timeoutMs) || options.timeoutMs <= 0) {
        throw new OfficeEngineError("INVALID_TIMEOUT", "timeoutMs must be a finite positive number");
      }
      throw new OfficeEngineError(
        "INLINE_TIMEOUT_UNSUPPORTED",
        "timeoutMs requires Worker execution because inline parsing is synchronous",
      );
    }
    const core = this.#core;
    if (core === undefined) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
    const fonts = this.#fonts;
    assertInputSize(bytes.byteLength, core.limits.inputBytes);
    const input = options.transferInput === true
      ? isUint8Array(bytes)
        ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
        : new Uint8Array(bytes)
      : new Uint8Array(copyByteSource(bytes));
    const password = encodePdfPassword(options.password);
    let coreDocument: CoreDocument;
    try {
      coreDocument = core.open(input, undefined, password);
    } finally {
      password?.fill(0);
    }
    if (coreDocument.requiresCalculation() && this.#calculationWasm !== false) {
      try {
        coreDocument.applyCalculation(
          await calculateSpreadsheet(this.#calculationWasm, input, core.limits),
        );
      } catch {
        // Calculation is optional; the Core diagnostic preserves best-effort rendering.
      }
    }
    try {
      throwIfAborted(options.signal);
    } catch (cause) {
      coreDocument.close();
      throw cause;
    }
    if (coreDocument.scene.fatal) {
      const diagnostics = coreDocument.scene.diagnostics;
      coreDocument.close();
      throw new OfficeEngineError(
        diagnostics[0]?.code ?? "INVALID_DOCUMENT",
        diagnostics[0]?.message ?? "The document is invalid or unsupported",
        { diagnostics },
      );
    }
    let registeredFonts: RegisteredFonts | undefined;
    let preflightFonts: RegisteredFonts | undefined;
    const additionalDiagnostics: Diagnostic[] = [];
    try {
      const requests = collectFontRequests(coreDocument.scene.objects, coreDocument.scene.fontAlternateNames);
      const embeddedFonts = prepareEmbeddedFontAssets(coreDocument.scene.embeddedFonts);
      const hostFonts = copyPreparedFonts(fonts);
      let providerRequests = requests;
      if (this.#fontProvider !== undefined) {
        preflightFonts = await registerFonts(
          hostFonts,
          undefined,
          undefined,
          options.signal,
          embeddedFonts,
          [],
          {
            policy: this.#fontPolicy,
            fontAlternateNames: coreDocument.scene.fontAlternateNames,
            requestedFaces: this.#fontPolicy === "local-first" ? requests : [],
          },
        );
        providerRequests = requests.filter((request) => !preflightFonts!.hasExactFace(request));
      }
      const providerResult = this.#fontProvider === undefined || providerRequests.length === 0
        ? { assets: [], diagnostics: [] }
        : await this.#fontProvider.resolve(providerRequests, options.signal);
      additionalDiagnostics.push(...providerResult.diagnostics);
      const providerBudget = Math.max(0, core.limits.fontBytes - fontAssetBytes(embeddedFonts));
      const providerFonts = prepareFontAssets(providerResult.assets, providerBudget);
      if (this.#fontPolicy === "local-first" && providerFonts.length === 0 && preflightFonts !== undefined) {
        registeredFonts = preflightFonts;
        preflightFonts = undefined;
      } else {
        preflightFonts?.close();
        preflightFonts = undefined;
        registeredFonts = await registerFonts(
          hostFonts,
          undefined,
          undefined,
          options.signal,
          embeddedFonts,
          [],
          {
            providerAssets: providerFonts,
            fontAlternateNames: coreDocument.scene.fontAlternateNames,
            policy: this.#fontPolicy,
            requestedFaces: requests,
          },
        );
      }
      throwIfAborted(options.signal);
      this.#assertOpen();
      const measuredLayoutFormat = coreDocument.scene.info?.format;
      const metricObjects = measuredLayoutFormat === undefined
        ? []
        : fontMetricLayoutObjects(coreDocument.scene.objects, measuredLayoutFormat);
      if (metricObjects.length > 0 && measuredLayoutFormat !== undefined) {
        const metrics = measureSceneFontMetrics(metricObjects, registeredFonts, {
          maxBytes: core.limits.fontBytes,
        });
        additionalDiagnostics.push(...metrics.diagnostics);
        if (metrics.metricCount !== 0 && metrics.table.byteLength <= core.limits.fontBytes) {
          const provisional = coreDocument;
          const relayoutPassword = encodePdfPassword(options.password);
          try {
            const measured = core.open(input, metrics.table, relayoutPassword);
            if (measured.scene.fatal || measured.scene.info === undefined) {
              const diagnostics = measured.scene.diagnostics;
              measured.close();
              throw new OfficeEngineError(
                diagnostics[0]?.code ?? "INVALID_DOCUMENT",
                diagnostics[0]?.message ?? "The measured document layout is invalid",
                { diagnostics },
              );
            }
            provisional.close();
            coreDocument = measured;
            additionalDiagnostics.push({
              code: "FONT_METRICS_APPLIED",
              severity: "info",
              fidelity: "approximate",
              phase: "layout",
              message: measuredLayoutFormat === "pptx"
                ? "PPTX table layout used browser-measured font advances"
                : `${measuredLayoutFormat.toUpperCase()} pagination used browser-measured font advances`,
              details: { faceCount: metrics.faceCount, metricCount: metrics.metricCount },
            });
          } catch (cause) {
            coreDocument = provisional;
            additionalDiagnostics.push({
              code: "FONT_METRICS_RELAYOUT_FAILED",
              severity: "warning",
              fidelity: "approximate",
              phase: "layout",
              message: `${measuredLayoutFormat.toUpperCase()} font-aware relayout failed; the bounded approximate layout was retained`,
              details: { error: cause instanceof Error ? cause.message : "Unknown relayout failure" },
            });
          } finally {
            relayoutPassword?.fill(0);
          }
        } else if (metrics.metricCount !== 0) {
          additionalDiagnostics.push({
            code: "FONT_METRICS_LIMIT",
            severity: "warning",
            fidelity: "approximate",
            phase: "layout",
            message: "Browser-measured font advances exceeded the document font byte limit",
            details: { bytes: metrics.table.byteLength, limit: core.limits.fontBytes },
          });
        }
      }
      throwIfAborted(options.signal);
      this.#assertOpen();
    } catch (cause) {
      preflightFonts?.close();
      registeredFonts?.close();
      coreDocument.close();
      throw cause;
    }
    const document = new InlineDocument(
      coreDocument,
      registeredFonts,
      (closed) => this.#documents.delete(closed),
      additionalDiagnostics,
      this.#fontPolicy === "local-first",
    );
    this.#documents.add(document);
    return document;
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    const core = this.#core;
    this.#core = undefined;
    this.#fonts = [];
    for (const document of [...this.#documents]) document.close();
    this.#fontProvider?.close();
    core?.close();
  }

  #assertOpen(): void {
    if (this.#closed) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
  }
}

async function createBaseOfficeEngine(options: EngineOptions): Promise<OfficeEngine> {
  const fontPolicy = options.fontPolicy ?? "local-first";
  if (fontPolicy !== "deterministic" && fontPolicy !== "local-first") {
    throw new OfficeEngineError("INVALID_FONT_POLICY", `Unsupported font policy ${String(fontPolicy)}`);
  }
  const limits = resolveLimits(options.limits);
  const fonts = prepareFontAssets(options.fonts, limits.fontBytes);
  await verifyFontAssetIntegrity(fonts);
  const documentLimits = Object.freeze({
    ...limits,
    fontBytes: limits.fontBytes - fontAssetBytes(fonts),
  });
  const fontProvider = options.fontProvider === undefined
    ? undefined
    : new FontProviderCache(options.fontProvider, {
        maxBytes: documentLimits.fontBytes,
        policy: fontPolicy,
        timeoutMs: options.fontProviderTimeoutMs,
      });
  if (options.execution === "inline") {
    return new InlineEngine(
      await Core.create(options.wasm, documentLimits),
      fonts,
      fontProvider,
      fontPolicy,
      options.calculationWasm,
    );
  }
  return WorkerEngine.create({ ...options, fontPolicy }, documentLimits, fonts, fontProvider);
}

export async function createOfficeEngine(options: EngineOptions = {}): Promise<OfficeEngine> {
  if (options.formatPack !== undefined
    && options.formatPack !== false
    && typeof options.formatPack !== "function") {
    throw new OfficeEngineError("INVALID_FORMAT_PACK", "formatPack must be false or an async pack factory");
  }
  const stableOptions = snapshotEngineOptions(options);
  const packInputBytes = stableOptions.formatPack === undefined || stableOptions.formatPack === false
    ? undefined
    : resolveLimits(stableOptions.limits).inputBytes;
  const FormatPackRuntimeClass = stableOptions.formatPack === undefined || stableOptions.formatPack === false
    ? undefined
    : (await import("./format-pack.js")).FormatPackRuntime;
  const nativeEngine = await createBaseOfficeEngine(stableOptions);
  if (FormatPackRuntimeClass === undefined
    || stableOptions.formatPack === undefined
    || stableOptions.formatPack === false) return nativeEngine;
  const { wasm: _nativeWasm, formatPack, ...packEngineOptions } = stableOptions;
  return new FormatPackRuntimeClass(
    nativeEngine,
    formatPack,
    (_candidate, wasm) => createBaseOfficeEngine(
      snapshotEngineOptions({ ...packEngineOptions, wasm, formatPack: false }),
    ),
    packInputBytes,
    stableOptions.execution === "inline" ? "inline" : "worker",
  );
}
