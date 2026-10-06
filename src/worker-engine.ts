import { assertInputSize, compileCore, encodePdfPassword } from "./core.js";
import { copyByteSource, isArrayBuffer } from "./bytes.js";
import type {
  Diagnostic,
  DocumentObject,
  EngineOptions,
  FontProviderDiagnostic,
  FontRequest,
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
  ResourceLimits,
  TextSearchRequest,
  TextSearchResult,
} from "./types.js";
import {
  immutableDiagnostics,
  immutableDocumentInfo,
  immutableDocumentObject,
  immutableHitResults,
  immutableRenderResult,
  OfficeEngineError,
  validateOperationOptions,
  validateRenderOptions,
} from "./types.js";
import type { WorkerRequest, WorkerResponse, WorkerSnapshot } from "./worker-protocol.js";
import {
  copyPreparedFonts,
  FontProviderCache,
  prepareFontAssets,
  type PreparedFontAsset,
} from "./font.js";

type WorkerCommand =
  | Omit<Extract<WorkerRequest, { type: "open" }>, "id">
  | Omit<Extract<WorkerRequest, { type: "render" }>, "id">
  | Omit<Extract<WorkerRequest, { type: "hit-test" }>, "id">
  | Omit<Extract<WorkerRequest, { type: "get-object" }>, "id">
  | Omit<Extract<WorkerRequest, { type: "list-objects" }>, "id">
  | Omit<Extract<WorkerRequest, { type: "search-text" }>, "id">;

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted === true) {
    throw new OfficeEngineError("OPERATION_ABORTED", "The operation was aborted", { cause: signal.reason });
  }
}

interface PendingRequest {
  readonly resolve: (response: WorkerResponse) => void;
  readonly reject: (error: unknown) => void;
  readonly cleanup: () => void;
  readonly onFontRequest?: (requests: readonly FontRequest[]) => Promise<{
    readonly fonts: readonly PreparedFontAsset[];
    readonly diagnostics: readonly FontProviderDiagnostic[];
  }>;
}

class WorkerClient {
  readonly #worker: Worker;
  readonly #pending = new Map<number, PendingRequest>();
  readonly #fontReplies = new Set<number>();
  #nextId = 1;
  #closed = false;
  #terminalError: OfficeEngineError | undefined;

  constructor(worker: Worker) {
    this.#worker = worker;
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const response = event.data as Partial<WorkerResponse> | null;
      if (response === null
        || typeof response !== "object"
        || !Number.isSafeInteger(response.id)
        || typeof response.ok !== "boolean") {
        this.close(new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Document worker returned an invalid response"));
        return;
      }
      const pending = this.#pending.get(response.id as number);
      if (pending === undefined) {
        this.close(new OfficeEngineError(
          "WORKER_PROTOCOL_ERROR",
          "Document worker returned an unknown request id",
        ));
        return;
      }
      if (response.ok === true && response.type === "font-request") {
        if (!Array.isArray(response.requests)
          || pending.onFontRequest === undefined
          || this.#fontReplies.has(response.id as number)) {
          this.close(new OfficeEngineError(
            "WORKER_PROTOCOL_ERROR",
            "Document worker returned an invalid font request",
          ));
          return;
        }
        const id = response.id as number;
        this.#fontReplies.add(id);
        void pending.onFontRequest(response.requests).then((reply) => {
          if (!this.#pending.has(id) || this.#closed) return;
          const fonts = copyPreparedFonts(reply.fonts);
          try {
            this.#worker.postMessage({
              id,
              type: "font-response",
              fonts,
              diagnostics: reply.diagnostics,
            } satisfies WorkerRequest, fonts.map((font) => font.bytes));
          } catch (cause) {
            this.close(new OfficeEngineError(
              "WORKER_PROTOCOL_ERROR",
              "Could not return provider fonts to the document worker",
              { cause },
            ));
          }
        }, (cause) => {
          this.close(cause instanceof OfficeEngineError
            ? cause
            : new OfficeEngineError("FONT_PROVIDER_FAILED", "Font provider resolution failed", { cause }));
        }).finally(() => this.#fontReplies.delete(id));
        return;
      }
      if (response.ok === false) {
        const error = response.error;
        if (error === undefined
          || typeof error.code !== "string"
          || typeof error.message !== "string"
          || !Array.isArray(error.diagnostics)) {
          this.close(new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Document worker returned an invalid error"));
          return;
        }
        this.#pending.delete(response.id as number);
        pending.cleanup();
        try {
          pending.reject(new OfficeEngineError(error.code, error.message, {
            diagnostics: error.diagnostics,
          }));
        } catch (cause) {
          const protocolError = new OfficeEngineError(
            "WORKER_PROTOCOL_ERROR",
            "Document worker returned malformed diagnostics",
            { cause },
          );
          pending.reject(protocolError);
          this.close(protocolError);
        }
      } else {
        this.#pending.delete(response.id as number);
        pending.cleanup();
        pending.resolve(response as WorkerResponse);
      }
    };
    worker.onerror = (event) => {
      this.close(new OfficeEngineError("WORKER_FAILED", event.message || "Document worker failed"));
    };
    worker.onmessageerror = () => {
      this.close(new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Document worker returned an unreadable message"));
    };
  }

  request(
    command: WorkerCommand,
    transfer: Transferable[] = [],
    options: {
      readonly signal?: AbortSignal;
      readonly timeoutMs?: number;
      readonly cancelRender?: boolean;
    } = {},
    onFontRequest?: PendingRequest["onFontRequest"],
  ): Promise<WorkerResponse> {
    if (this.#closed) {
      return Promise.reject(this.#terminalError
        ?? new OfficeEngineError("DOCUMENT_CLOSED", "The document worker is closed"));
    }
    const id = this.#nextId++;
    return new Promise((resolve, reject) => {
      let timer: ReturnType<typeof setTimeout> | undefined;
      const abort = () => {
        if (options.cancelRender === true) {
          try {
            this.#worker.postMessage({ id, type: "cancel-render", targetId: id } satisfies WorkerRequest);
          } catch (cause) {
            this.close(new OfficeEngineError(
              "WORKER_PROTOCOL_ERROR",
              "Could not cancel document rendering",
              { cause },
            ));
          }
          return;
        }
        this.close(new OfficeEngineError(
          "OPERATION_ABORTED",
          "The operation was aborted",
          { cause: options.signal?.reason },
        ));
      };
      const cleanup = () => {
        if (timer !== undefined) clearTimeout(timer);
        options.signal?.removeEventListener("abort", abort);
      };
      this.#pending.set(id, {
        resolve,
        reject,
        cleanup,
        ...(onFontRequest === undefined ? {} : { onFontRequest }),
      });
      if (options.signal?.aborted === true) {
        abort();
        return;
      }
      options.signal?.addEventListener("abort", abort, { once: true });
      if (options.timeoutMs !== undefined) {
        timer = setTimeout(() => {
          this.close(new OfficeEngineError("OPERATION_TIMEOUT", `Operation exceeded ${options.timeoutMs}ms`));
        }, options.timeoutMs);
      }
      try {
        this.#worker.postMessage({ ...command, id } as WorkerRequest, transfer);
      } catch (cause) {
        this.close(new OfficeEngineError(
          "WORKER_PROTOCOL_ERROR",
          "Could not send a command to the document worker",
          { cause },
        ));
      }
    });
  }

  assertActive(): void {
    if (this.#closed) {
      throw this.#terminalError
        ?? new OfficeEngineError("DOCUMENT_CLOSED", "The document worker is closed");
    }
  }

  close(error = new OfficeEngineError("DOCUMENT_CLOSED", "The document worker was closed")): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#terminalError = error;
    this.#worker.onmessage = null;
    this.#worker.onerror = null;
    this.#worker.onmessageerror = null;
    try {
      this.#worker.terminate();
    } catch {
      // Logical closure and pending-request cleanup must not depend on a custom factory.
    }
    for (const pending of this.#pending.values()) {
      pending.cleanup();
      pending.reject(error);
    }
    this.#pending.clear();
    this.#fontReplies.clear();
  }
}

class WorkerDocument implements OfficeDocument {
  readonly #client: WorkerClient;
  readonly #onClose: (document: WorkerDocument) => void;
  readonly info;
  #diagnostics: readonly Diagnostic[];
  #closed = false;

  constructor(client: WorkerClient, snapshot: WorkerSnapshot, onClose: (document: WorkerDocument) => void) {
    this.#client = client;
    this.#onClose = onClose;
    this.info = immutableDocumentInfo(snapshot.info);
    this.#diagnostics = immutableDiagnostics(snapshot.diagnostics);
  }

  diagnostics() {
    this.#assertOpen();
    return this.#diagnostics;
  }

  async render(request: RenderRequest, options: RenderOptions = {}): Promise<RenderResult> {
    this.#assertOpen();
    validateRenderOptions(options);
    throwIfAborted(options.signal);
    const response = await this.#client.request({
      type: "render",
      request,
      priority: options.priority ?? "visible",
      ...(options.supersedeKey === undefined ? {} : { supersedeKey: options.supersedeKey }),
    }, [], {
      ...(options.signal === undefined ? {} : { signal: options.signal }),
      ...(options.timeoutMs === undefined ? {} : { timeoutMs: options.timeoutMs }),
      cancelRender: true,
    });
    if (!response.ok || response.type !== "render") {
      throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected render response");
    }
    this.#updateDiagnostics(response.documentDiagnostics);
    return immutableRenderResult(response.result);
  }

  async hitTest(request: HitTestRequest, options: OperationOptions = {}): Promise<readonly HitResult[]> {
    this.#assertOpen();
    validateOperationOptions(options);
    throwIfAborted(options.signal);
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
    const response = await this.#client.request({
      type: "hit-test",
      unitIndex: request.unitIndex,
      x: request.x,
      y: request.y,
      limit,
    }, [], options);
    if (!response.ok || response.type !== "hit-test") {
      throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected hit-test response");
    }
    this.#updateDiagnostics(response.documentDiagnostics);
    return immutableHitResults(response.results.map((result) => ({
      object: immutableDocumentObject(result.object),
      ancestors: result.ancestors.map(immutableDocumentObject),
    })));
  }

  async getObject(id: string, options: OperationOptions = {}): Promise<DocumentObject | undefined> {
    this.#assertOpen();
    validateOperationOptions(options);
    throwIfAborted(options.signal);
    const response = await this.#client.request({ type: "get-object", objectId: id }, [], options);
    if (!response.ok || response.type !== "get-object") {
      throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected object response");
    }
    this.#updateDiagnostics(response.documentDiagnostics);
    return response.object === undefined ? undefined : immutableDocumentObject(response.object);
  }

  async searchText(request: TextSearchRequest, options: OperationOptions = {}): Promise<readonly TextSearchResult[]> {
    this.#assertOpen();
    validateOperationOptions(options);
    throwIfAborted(options.signal);
    const response = await this.#client.request({ type: "search-text", request }, [], options);
    if (!response.ok || response.type !== "search-text") {
      throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected text search response");
    }
    this.#updateDiagnostics(response.documentDiagnostics);
    return Object.freeze(response.results.map((result) => Object.freeze({
      object: immutableDocumentObject(result.object),
      ranges: Object.freeze(result.ranges.map((range) => Object.freeze([...range] as [number, number]))),
    })));
  }

  async listObjects(request: ObjectListRequest = {}, options: OperationOptions = {}): Promise<readonly DocumentObject[]> {
    this.#assertOpen();
    validateOperationOptions(options);
    throwIfAborted(options.signal);
    const response = await this.#client.request({ type: "list-objects", request }, [], options);
    if (!response.ok || response.type !== "list-objects") {
      throw new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected object list response");
    }
    this.#updateDiagnostics(response.documentDiagnostics);
    return Object.freeze(response.objects.map(immutableDocumentObject));
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#client.close();
    this.#onClose(this);
  }

  #assertOpen(): void {
    if (this.#closed) throw new OfficeEngineError("DOCUMENT_CLOSED", "The document is closed");
    this.#client.assertActive();
  }

  #updateDiagnostics(diagnostics: readonly Diagnostic[] | undefined): void {
    if (diagnostics !== undefined) this.#diagnostics = immutableDiagnostics(diagnostics);
  }
}

export class WorkerEngine implements OfficeEngine {
  #module: WebAssembly.Module | undefined;
  readonly #limits: ResourceLimits;
  readonly #url: URL | string;
  readonly #factory: (url: URL | string) => Worker;
  readonly #fontProvider: FontProviderCache | undefined;
  readonly #fontPolicy: "deterministic" | "local-first";
  readonly #prewarmWorkers: boolean;
  readonly #calculationWasm: EngineOptions["calculationWasm"];
  #fonts: readonly PreparedFontAsset[];
  #idleClient: WorkerClient | undefined;
  readonly #documents = new Set<WorkerDocument>();
  readonly #pendingClients = new Set<WorkerClient>();
  #closed = false;

  private constructor(
    module: WebAssembly.Module,
    limits: ResourceLimits,
    url: URL | string,
    factory: (url: URL | string) => Worker,
    fonts: readonly PreparedFontAsset[],
    fontProvider: FontProviderCache | undefined,
    fontPolicy: "deterministic" | "local-first",
    prewarmWorkers: boolean,
    calculationWasm: EngineOptions["calculationWasm"],
  ) {
    this.#module = module;
    this.#limits = limits;
    this.#url = url;
    this.#factory = factory;
    this.#fonts = fonts;
    this.#fontProvider = fontProvider;
    this.#fontPolicy = fontPolicy;
    this.#prewarmWorkers = prewarmWorkers;
    this.#calculationWasm = calculationWasm instanceof URL ? calculationWasm.href : calculationWasm;
    this.#prepareIdleClient();
  }

  static async create(
    options: EngineOptions,
    limits: ResourceLimits,
    fonts: readonly PreparedFontAsset[],
    fontProvider: FontProviderCache | undefined,
  ): Promise<WorkerEngine> {
    if (typeof Worker === "undefined" && options.workerFactory === undefined) {
      throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "Module Worker support is required");
    }
    const module = await compileCore(options.wasm);
    const url = options.workerUrl ?? new URL("./worker.js", import.meta.url);
    const factory = options.workerFactory ?? ((workerUrl) => new Worker(workerUrl, {
      type: "module",
      name: "office-viewer-document",
    }));
    return new WorkerEngine(
      module,
      limits,
      url,
      factory,
      fonts,
      fontProvider,
      options.fontPolicy ?? "local-first",
      options.workerFactory === undefined,
      options.calculationWasm,
    );
  }

  async open(bytes: ArrayBuffer | Uint8Array, options: OpenOptions = {}): Promise<OfficeDocument> {
    if (this.#closed) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
    if (options.timeoutMs !== undefined && (!Number.isFinite(options.timeoutMs) || options.timeoutMs <= 0)) {
      throw new OfficeEngineError("INVALID_TIMEOUT", "timeoutMs must be a finite positive number");
    }
    throwIfAborted(options.signal);
    const module = this.#module;
    if (module === undefined) throw new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
    assertInputSize(bytes.byteLength, this.#limits.inputBytes);
    const input = options.transferInput === true && isArrayBuffer(bytes)
      ? bytes
      : copyByteSource(bytes);
    const encodedPassword = encodePdfPassword(options.password);
    const password = encodedPassword === undefined
      ? undefined
      : new Uint8Array(encodedPassword).buffer;
    encodedPassword?.fill(0);
    const fonts = copyPreparedFonts(this.#fonts);
    let client: WorkerClient;
    try {
      client = this.#idleClient ?? new WorkerClient(this.#factory(this.#url));
      this.#idleClient = undefined;
      this.#prepareIdleClient();
    } catch (cause) {
      if (password !== undefined && password.byteLength !== 0) new Uint8Array(password).fill(0);
      throw new OfficeEngineError("WORKER_START_FAILED", "Could not start the document Worker; check worker-src CSP", { cause });
    }
    if (this.#closed) {
      if (password !== undefined && password.byteLength !== 0) new Uint8Array(password).fill(0);
      const error = new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
      client.close(error);
      throw error;
    }
    this.#pendingClients.add(client);
    let response: WorkerResponse;
    try {
      response = await client.request(
        {
          type: "open",
          module,
          ...(this.#calculationWasm === undefined ? {} : { calculationWasm: this.#calculationWasm }),
          bytes: input,
          ...(password === undefined ? {} : { password }),
          limits: this.#limits,
          fonts,
          fontPolicy: this.#fontPolicy,
        },
        [input, ...(password === undefined ? [] : [password]), ...fonts.map((font) => font.bytes)],
        {
          ...(options.signal === undefined ? {} : { signal: options.signal }),
          timeoutMs: options.timeoutMs ?? 60_000,
        },
        async (requests) => {
          if (this.#fontProvider === undefined) return { fonts: [], diagnostics: [] };
          const result = await this.#fontProvider.resolve(requests, options.signal);
          return {
            fonts: prepareFontAssets(result.assets, this.#limits.fontBytes),
            diagnostics: result.diagnostics,
          };
        },
      );
    } catch (cause) {
      this.#pendingClients.delete(client);
      client.close(cause instanceof OfficeEngineError ? cause : undefined);
      throw cause;
    } finally {
      if (password !== undefined && password.byteLength !== 0) new Uint8Array(password).fill(0);
    }
    this.#pendingClients.delete(client);
    if (this.#closed) {
      const error = new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
      client.close(error);
      throw error;
    }
    if (!response.ok || response.type !== "open") {
      const error = new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Unexpected open response");
      client.close(error);
      throw error;
    }
    let document: WorkerDocument;
    try {
      document = new WorkerDocument(client, response.snapshot, (closed) => this.#documents.delete(closed));
    } catch (cause) {
      const error = cause instanceof OfficeEngineError
        ? cause
        : new OfficeEngineError("WORKER_PROTOCOL_ERROR", "Document worker returned an invalid snapshot", { cause });
      client.close(error);
      throw error;
    }
    this.#documents.add(document);
    return document;
  }

  close(): void {
    if (this.#closed) return;
    this.#closed = true;
    this.#module = undefined;
    this.#fonts = [];
    const error = new OfficeEngineError("ENGINE_CLOSED", "The engine is closed");
    for (const client of this.#pendingClients) client.close(error);
    this.#pendingClients.clear();
    for (const document of [...this.#documents]) document.close();
    this.#idleClient?.close(error);
    this.#idleClient = undefined;
    this.#fontProvider?.close();
  }

  #prepareIdleClient(): void {
    if (this.#closed || !this.#prewarmWorkers || this.#idleClient !== undefined) return;
    try {
      this.#idleClient = new WorkerClient(this.#factory(this.#url));
    } catch {
      // Preserve the existing open-time WORKER_START_FAILED contract when the
      // environment cannot create a Worker during speculative prewarming.
    }
  }
}
