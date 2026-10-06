import { Core, type CoreDocument } from "./core.js";
import { calculateSpreadsheet } from "./core.js";
import { SceneRenderer, sheetFormulaViewportBounds } from "./render.js";
import { fontMetricLayoutObjects, measureSceneFontMetrics } from "./font-metrics.js";
import { documentSpaceBounds, sceneTextLayout, type SceneObject } from "./scene.js";
import type { Diagnostic, DocumentObject, HitResult } from "./types.js";
import { listDocumentObjects, OfficeEngineError, searchDocumentText } from "./types.js";
import {
  collectFontRequests,
  documentFontRuns,
  fontAssetBytes,
  prepareEmbeddedFontAssets,
  registerFonts,
  type RegisteredFonts,
} from "./font.js";
import type { FontProviderDiagnostic } from "./types.js";
import type { PreparedFontAsset } from "./font.js";
import type { SerializedOfficeError, WorkerRequest, WorkerResponse } from "./worker-protocol.js";

interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerResponse, transfer?: Transferable[]): void;
}

const scope = globalThis as unknown as WorkerScope;
let document: CoreDocument | undefined;
let renderer: SceneRenderer | undefined;
let fonts: RegisteredFonts | undefined;
let fontPolicy: "deterministic" | "local-first" = "deterministic";
let additionalDocumentDiagnostics: readonly Diagnostic[] = [];
let objectsById = new Map<string, SceneObject>();
let objectsByNumericId = new Map<number, SceneObject>();
const publicObjects = new Map<number, DocumentObject>();
const pendingFontResponses = new Map<number, {
  readonly resolve: (value: {
    readonly fonts: readonly PreparedFontAsset[];
    readonly diagnostics: readonly FontProviderDiagnostic[];
  }) => void;
}>();
let documentOperation = Promise.resolve();

function enqueueDocumentOperation<T>(operation: () => Promise<T> | T): Promise<T> {
  const scheduled = documentOperation.then(operation, operation);
  documentOperation = scheduled.then(() => undefined, () => undefined);
  return scheduled;
}

function requestProviderFonts(
  id: number,
  requests: ReturnType<typeof collectFontRequests>,
): Promise<{
  readonly fonts: readonly PreparedFontAsset[];
  readonly diagnostics: readonly FontProviderDiagnostic[];
}> {
  if (pendingFontResponses.has(id)) {
    return Promise.reject(new OfficeEngineError("WORKER_STATE_INVALID", "A font request is already pending"));
  }
  return new Promise((resolve) => {
    pendingFontResponses.set(id, { resolve });
    scope.postMessage({ id, ok: true, type: "font-request", requests });
  });
}

function publicObject(object: SceneObject, registeredFonts: RegisteredFonts): DocumentObject {
  const cached = publicObjects.get(object.numericId);
  if (cached !== undefined) {
    publicObjects.delete(object.numericId);
    publicObjects.set(object.numericId, cached);
    return cached;
  }
  const fontRuns = documentFontRuns(object, registeredFonts);
  const value: DocumentObject = {
    id: object.id,
    type: object.type,
    unitIndex: object.unitIndex,
    bounds: documentSpaceBounds(object, objectsByNumericId, document?.scene.info?.format === "odp"),
    hidden: object.hidden,
    ...(object.parentId === undefined ? {} : { parentId: object.parentId }),
    ...(object.text === undefined ? {} : { text: object.text }),
    ...(fontRuns === undefined ? {} : { fontRuns }),
    ...(object.source.format === "xlsx" ? { wrapText: sceneTextLayout(object.visual)?.wrap ?? false } : {}),
    ...(object.name === undefined ? {} : { name: object.name }),
    ...(object.title === undefined ? {} : { title: object.title }),
    ...(object.description === undefined ? {} : { description: object.description }),
    ...(object.actions === undefined ? {} : { actions: object.actions }),
    source: object.source,
  };
  publicObjects.set(object.numericId, value);
  while (publicObjects.size > 50_000) {
    const oldest = publicObjects.keys().next().value;
    if (oldest === undefined) break;
    publicObjects.delete(oldest);
  }
  return value;
}

function requireOpenDocument(): {
  readonly document: CoreDocument;
  readonly renderer: SceneRenderer;
  readonly fonts: RegisteredFonts;
} {
  if (document === undefined || renderer === undefined || fonts === undefined || document.scene.info === undefined) {
    throw new OfficeEngineError("WORKER_STATE_INVALID", "Worker has no open document");
  }
  return { document, renderer, fonts };
}

function currentDocumentDiagnostics(state: ReturnType<typeof requireOpenDocument>): readonly Diagnostic[] {
  return [
    ...state.document.scene.diagnostics,
    ...additionalDocumentDiagnostics,
    ...state.fonts.diagnostics(),
  ];
}

function replaceScene(scene: CoreDocument["scene"], registeredFonts: RegisteredFonts): void {
  const nextRenderer = new SceneRenderer(
    scene.objects,
    document!.limits,
    registeredFonts,
    registeredFonts.imageCodecFonts(),
  );
  renderer?.close();
  renderer = nextRenderer;
  objectsById = new Map(scene.objects.map((object) => [object.id, object]));
  objectsByNumericId = new Map(scene.objects.map((object) => [object.numericId, object]));
  publicObjects.clear();
}

function appendScene(scene: CoreDocument["scene"]): void {
  renderer!.appendObjects(scene.objects);
  for (const object of scene.objects) {
    objectsById.set(object.id, object);
    objectsByNumericId.set(object.numericId, object);
  }
}

async function ensureUnitLoaded(
  unitIndex: number,
  formulaBounds?: { readonly maxRow: number; readonly maxColumn: number },
): Promise<ReturnType<typeof requireOpenDocument>> {
  const state = requireOpenDocument();
  const loaded = formulaBounds === undefined
    ? state.document.loadUnit(unitIndex)
    : state.document.loadUnitRegion(unitIndex, formulaBounds.maxRow, formulaBounds.maxColumn);
  if (loaded === undefined) return state;
  const { scene, replace } = loaded;
  await state.fonts.addEmbeddedAssets(prepareEmbeddedFontAssets(scene.embeddedFonts));
  if (fontPolicy === "local-first") await state.fonts.addLocalRequests(collectFontRequests(scene.objects));
  if (replace) replaceScene(scene, state.fonts);
  else appendScene(scene);
  return requireOpenDocument();
}

async function ensureAllLoaded(): Promise<ReturnType<typeof requireOpenDocument>> {
  const state = requireOpenDocument();
  const scene = state.document.loadAll();
  if (scene === undefined) return state;
  await state.fonts.addEmbeddedAssets(prepareEmbeddedFontAssets(scene.embeddedFonts));
  if (fontPolicy === "local-first") await state.fonts.addLocalRequests(collectFontRequests(scene.objects));
  replaceScene(scene, state.fonts);
  return requireOpenDocument();
}

function hitResults(numericIds: readonly number[], registeredFonts: RegisteredFonts): readonly HitResult[] {
  const results: HitResult[] = [];
  for (const numericId of numericIds) {
    const object = objectsByNumericId.get(numericId);
    if (object === undefined) continue;
    const ancestors: DocumentObject[] = [];
    const visited = new Set<number>([numericId]);
    let parent = object.parentNumericId;
    while (parent !== undefined && ancestors.length < 128 && !visited.has(parent)) {
      visited.add(parent);
      const ancestor = objectsByNumericId.get(parent);
      if (ancestor === undefined) break;
      ancestors.unshift(publicObject(ancestor, registeredFonts));
      parent = ancestor.parentNumericId;
    }
    results.push({ object: publicObject(object, registeredFonts), ancestors });
  }
  return results;
}

function serializeError(cause: unknown): SerializedOfficeError {
  if (cause instanceof OfficeEngineError) {
    return { code: cause.code, message: cause.message, diagnostics: cause.diagnostics };
  }
  if (cause instanceof Error) {
    return { code: "WORKER_INTERNAL_ERROR", message: cause.message, diagnostics: [] };
  }
  return { code: "WORKER_INTERNAL_ERROR", message: "Unknown worker failure", diagnostics: [] };
}

async function handle(request: WorkerRequest): Promise<void> {
  try {
    if (request.type === "font-response") {
      const pending = pendingFontResponses.get(request.id);
      if (pending === undefined) {
        throw new OfficeEngineError("WORKER_STATE_INVALID", "Worker received an unexpected font response");
      }
      pendingFontResponses.delete(request.id);
      pending.resolve({ fonts: request.fonts, diagnostics: request.diagnostics });
      return;
    }
    if (request.type === "open") {
      if (document !== undefined) throw new OfficeEngineError("WORKER_STATE_INVALID", "Worker already owns a document");
      const core = await Core.create(request.module, request.limits);
      const input = new Uint8Array(request.bytes);
      const password = request.password === undefined ? undefined : new Uint8Array(request.password);
      let opened;
      try {
        opened = core.open(input, undefined, password);
      } catch (cause) {
        password?.fill(0);
        throw cause;
      }
      if (opened.requiresCalculation() && request.calculationWasm !== false) {
        try {
          opened.applyCalculation(
            await calculateSpreadsheet(request.calculationWasm, input, request.limits),
          );
        } catch {
          // Calculation is optional; the Core diagnostic preserves best-effort rendering.
        }
      }
      let loadedFonts: RegisteredFonts | undefined;
      let preflightFonts: RegisteredFonts | undefined;
      try {
        const provisionalScene = opened.scene;
        if (provisionalScene.fatal || provisionalScene.info === undefined) {
          throw new OfficeEngineError(
            provisionalScene.diagnostics[0]?.code ?? "INVALID_DOCUMENT",
            provisionalScene.diagnostics[0]?.message ?? "The document is invalid or unsupported",
            { diagnostics: provisionalScene.diagnostics },
          );
        }
        const fontRequests = collectFontRequests(provisionalScene.objects, provisionalScene.fontAlternateNames);
        const embeddedFonts = prepareEmbeddedFontAssets(provisionalScene.embeddedFonts);
        preflightFonts = await registerFonts(
          request.fonts,
          undefined,
          undefined,
          undefined,
          embeddedFonts,
          [],
          {
            policy: request.fontPolicy,
            fontAlternateNames: provisionalScene.fontAlternateNames,
            requestedFaces: request.fontPolicy === "local-first" ? fontRequests : [],
          },
        );
        const providerRequests = fontRequests.filter((fontRequest) => !preflightFonts!.hasExactFace(fontRequest));
        const provider = providerRequests.length === 0
          ? { fonts: [], diagnostics: [] }
          : await requestProviderFonts(request.id, providerRequests);
        if (fontAssetBytes(embeddedFonts) + fontAssetBytes(provider.fonts) > request.limits.fontBytes) {
          throw new OfficeEngineError(
            "FONT_BYTES_LIMIT",
            "Document-embedded and provider fonts exceed the configured font byte limit",
          );
        }
        if (request.fontPolicy === "local-first" && provider.fonts.length === 0) {
          loadedFonts = preflightFonts;
          preflightFonts = undefined;
        } else {
          preflightFonts.close();
          preflightFonts = undefined;
          loadedFonts = await registerFonts(
            request.fonts,
            undefined,
            undefined,
            undefined,
            embeddedFonts,
            [],
            {
              providerAssets: provider.fonts,
              fontAlternateNames: provisionalScene.fontAlternateNames,
              policy: request.fontPolicy,
              requestedFaces: fontRequests,
            },
          );
        }
        const additionalDiagnostics: Diagnostic[] = [...provider.diagnostics];
        const measuredLayoutFormat = provisionalScene.info.format;
        const metricObjects = fontMetricLayoutObjects(provisionalScene.objects, measuredLayoutFormat);
        if (metricObjects.length > 0) {
          const metrics = measureSceneFontMetrics(metricObjects, loadedFonts, {
            maxBytes: request.limits.fontBytes,
          });
          additionalDiagnostics.push(...metrics.diagnostics);
          if (metrics.metricCount !== 0 && metrics.table.byteLength <= request.limits.fontBytes) {
            try {
              const measured = core.open(input, metrics.table, password);
              if (measured.scene.fatal || measured.scene.info === undefined) {
                const diagnostics = measured.scene.diagnostics;
                measured.close();
                throw new OfficeEngineError(
                  diagnostics[0]?.code ?? "INVALID_DOCUMENT",
                  diagnostics[0]?.message ?? "The measured document layout is invalid",
                  { diagnostics },
                );
              }
              opened.close();
              opened = measured;
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
              additionalDiagnostics.push({
                code: "FONT_METRICS_RELAYOUT_FAILED",
                severity: "warning",
                fidelity: "approximate",
                phase: "layout",
                message: `${measuredLayoutFormat.toUpperCase()} font-aware relayout failed; the bounded approximate layout was retained`,
                details: { error: cause instanceof Error ? cause.message : "Unknown relayout failure" },
              });
            }
          } else if (metrics.metricCount !== 0) {
            additionalDiagnostics.push({
              code: "FONT_METRICS_LIMIT",
              severity: "warning",
              fidelity: "approximate",
              phase: "layout",
              message: "Browser-measured font advances exceeded the document font byte limit",
              details: { bytes: metrics.table.byteLength, limit: request.limits.fontBytes },
            });
          }
        }
        const scene = opened.scene;
        if (scene.info === undefined) {
          throw new OfficeEngineError("INVALID_DOCUMENT", "The final document has no document information");
        }
        const loadedRenderer = new SceneRenderer(
          scene.objects,
          request.limits,
          loadedFonts,
          loadedFonts.imageCodecFonts(),
        );
        scope.postMessage({
          id: request.id,
          ok: true,
          type: "open",
          snapshot: {
            info: scene.info,
            diagnostics: [
              ...scene.diagnostics,
              ...additionalDiagnostics,
              ...loadedFonts.diagnostics(),
            ],
          },
        });
        document = opened;
        renderer = loadedRenderer;
        fonts = loadedFonts;
        fontPolicy = request.fontPolicy;
        additionalDocumentDiagnostics = additionalDiagnostics;
        objectsById = new Map(scene.objects.map((object) => [object.id, object]));
        objectsByNumericId = new Map(scene.objects.map((object) => [object.numericId, object]));
        publicObjects.clear();
      } catch (cause) {
        preflightFonts?.close();
        loadedFonts?.close();
        opened.close();
        core.close();
        throw cause;
      } finally {
        password?.fill(0);
      }
      return;
    }

    let state = requireOpenDocument();
    if (request.type === "hit-test") {
      state = await ensureUnitLoaded(request.unitIndex);
      const numericIds = state.document.hitTest(request.unitIndex, request.x, request.y, request.limit);
      scope.postMessage({
        id: request.id,
        ok: true,
        type: "hit-test",
        results: hitResults(numericIds, state.fonts),
        documentDiagnostics: currentDocumentDiagnostics(state),
      });
      return;
    }
    if (request.type === "get-object") {
      let object = objectsById.get(request.objectId);
      if (object === undefined) {
        state = await ensureAllLoaded();
        object = objectsById.get(request.objectId);
      }
      scope.postMessage({
        id: request.id,
        ok: true,
        type: "get-object",
        ...(object === undefined ? {} : { object: publicObject(object, state.fonts) }),
        documentDiagnostics: currentDocumentDiagnostics(state),
      });
      return;
    }
    if (request.type === "list-objects") {
      if (request.request.unitIndex === undefined) {
        state = await ensureAllLoaded();
      } else {
        const unit = state.document.scene.info!.units[request.request.unitIndex];
        state = await ensureUnitLoaded(
          request.request.unitIndex,
          unit === undefined ? undefined : sheetFormulaViewportBounds(unit, request.request),
        );
      }
      const selected = listDocumentObjects(state.document.scene.objects, state.document.scene.info!.units, request.request);
      scope.postMessage({
        id: request.id,
        ok: true,
        type: "list-objects",
        objects: selected.map((object) => publicObject(objectsById.get(object.id)!, state.fonts)),
        documentDiagnostics: currentDocumentDiagnostics(state),
      });
      return;
    }
    if (request.type === "search-text") {
      state = request.request.unitIndex === undefined
        ? await ensureAllLoaded()
        : await ensureUnitLoaded(request.request.unitIndex);
      const selected = searchDocumentText(state.document.scene.objects, state.document.scene.info!.units, request.request);
      scope.postMessage({
        id: request.id,
        ok: true,
        type: "search-text",
        results: selected.map((result) => ({
          object: publicObject(objectsById.get(result.object.id)!, state.fonts),
          ranges: result.ranges,
        })),
        documentDiagnostics: currentDocumentDiagnostics(state),
      });
      return;
    }
    throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", `Unsupported worker request: ${request.type}`);
  } catch (cause) {
    pendingFontResponses.delete(request.id);
    scope.postMessage({ id: request.id, ok: false, error: serializeError(cause) });
  }
}

type RenderWorkerRequest = Extract<WorkerRequest, { type: "render" }>;

interface ScheduledRender {
  readonly request: RenderWorkerRequest;
  readonly sequence: number;
  cancelled: boolean;
}

const renderQueue: ScheduledRender[] = [];
let activeRender: ScheduledRender | undefined;
let renderSequence = 0;

function renderPriority(priority: RenderWorkerRequest["priority"]): number {
  if (priority === "interactive") return 2;
  if (priority === "visible") return 1;
  return 0;
}

function rejectRender(item: ScheduledRender, message: string): void {
  scope.postMessage({
    id: item.request.id,
    ok: false,
    error: serializeError(new OfficeEngineError("OPERATION_ABORTED", message)),
  });
}

function cancelQueuedRender(targetId: number, message: string): boolean {
  const index = renderQueue.findIndex((item) => item.request.id === targetId);
  if (index === -1) return false;
  const [item] = renderQueue.splice(index, 1);
  if (item !== undefined) rejectRender(item, message);
  return true;
}

function cancelRender(targetId: number): void {
  if (activeRender?.request.id === targetId) {
    activeRender.cancelled = true;
    return;
  }
  cancelQueuedRender(targetId, "Rendering was cancelled");
}

function enqueueRender(request: RenderWorkerRequest): void {
  if (request.supersedeKey !== undefined) {
    for (let index = renderQueue.length - 1; index >= 0; index -= 1) {
      const item = renderQueue[index];
      if (item?.request.supersedeKey !== request.supersedeKey) continue;
      renderQueue.splice(index, 1);
      rejectRender(item, "Rendering was superseded by a newer request");
    }
    if (activeRender?.request.supersedeKey === request.supersedeKey) activeRender.cancelled = true;
  }
  if (activeRender !== undefined
    && renderPriority(request.priority) > renderPriority(activeRender.request.priority)) {
    activeRender.cancelled = true;
  }
  renderQueue.push({ request, sequence: renderSequence++, cancelled: false });
  renderQueue.sort((left, right) => {
    const byPriority = renderPriority(right.request.priority) - renderPriority(left.request.priority);
    return byPriority === 0 ? left.sequence - right.sequence : byPriority;
  });
  void drainRenderQueue();
}

async function drainRenderQueue(): Promise<void> {
  if (activeRender !== undefined) return;
  while (renderQueue.length > 0) {
    const item = renderQueue.shift();
    if (item === undefined) return;
    activeRender = item;
    try {
      await enqueueDocumentOperation(async () => {
        const requestedUnit = requireOpenDocument().document.scene.info!.units[item.request.request.unitIndex];
        const state = await ensureUnitLoaded(
          item.request.request.unitIndex,
          requestedUnit === undefined ? undefined : sheetFormulaViewportBounds(requestedUnit, item.request.request),
        );
        const unit = state.document.scene.info!.units[item.request.request.unitIndex];
        if (unit === undefined || unit.index !== item.request.request.unitIndex) {
          throw new OfficeEngineError("INVALID_UNIT", `Document has no unit at index ${item.request.request.unitIndex}`);
        }
        const result = await state.renderer.render(unit, item.request.request, {
          isCancelled: () => item.cancelled,
          yieldEvery: item.request.priority === "prefetch"
            ? 128
            : state.document.scene.info!.format === "pdf" ? 8_192 : 1_024,
        });
        if (item.cancelled) {
          result.bitmap.close();
          throw new OfficeEngineError("OPERATION_ABORTED", "Rendering was superseded or cancelled");
        }
        scope.postMessage({
          id: item.request.id,
          ok: true,
          type: "render",
          result,
          documentDiagnostics: currentDocumentDiagnostics(state),
        }, [result.bitmap]);
      });
    } catch (cause) {
      scope.postMessage({ id: item.request.id, ok: false, error: serializeError(cause) });
    } finally {
      activeRender = undefined;
    }
  }
}

scope.onmessage = (event) => {
  const request = event.data;
  if (request.type === "font-response") {
    void handle(request);
    return;
  }
  if (request.type === "cancel-render") {
    cancelRender(request.targetId);
    return;
  }
  if (request.type === "render") {
    enqueueRender(request);
    return;
  }
  void enqueueDocumentOperation(() => handle(request));
};
