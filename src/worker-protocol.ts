import type {
  Diagnostic,
  DocumentInfo,
  DocumentObject,
  FontPolicy,
  FontProviderDiagnostic,
  FontRequest,
  HitResult,
  ObjectListRequest,
  RenderPriority,
  RenderRequest,
  RenderResult,
  ResourceLimits,
  TextSearchRequest,
  TextSearchResult,
  WasmSource,
} from "./types.js";
import type { PreparedFontAsset } from "./font.js";

export interface WorkerSnapshot {
  readonly info: DocumentInfo;
  readonly diagnostics: readonly Diagnostic[];
}

export type WorkerRequest =
  | {
      readonly id: number;
      readonly type: "open";
      readonly module: WebAssembly.Module;
      readonly calculationWasm?: WasmSource | false;
      readonly bytes: ArrayBuffer;
      readonly password?: ArrayBuffer;
      readonly limits: ResourceLimits;
      readonly fonts: readonly PreparedFontAsset[];
      readonly fontPolicy: FontPolicy;
    }
  | {
      readonly id: number;
      readonly type: "font-response";
      readonly fonts: readonly PreparedFontAsset[];
      readonly diagnostics: readonly FontProviderDiagnostic[];
    }
  | {
      readonly id: number;
      readonly type: "render";
      readonly request: RenderRequest;
      readonly priority: RenderPriority;
      readonly supersedeKey?: string;
    }
  | { readonly id: number; readonly type: "cancel-render"; readonly targetId: number }
  | {
      readonly id: number;
      readonly type: "hit-test";
      readonly unitIndex: number;
      readonly x: number;
      readonly y: number;
      readonly limit: number;
    }
  | { readonly id: number; readonly type: "get-object"; readonly objectId: string }
  | { readonly id: number; readonly type: "list-objects"; readonly request: ObjectListRequest }
  | { readonly id: number; readonly type: "search-text"; readonly request: TextSearchRequest };

export interface SerializedOfficeError {
  readonly code: string;
  readonly message: string;
  readonly diagnostics: readonly Diagnostic[];
}

interface WorkerDocumentDiagnosticUpdate {
  readonly documentDiagnostics?: readonly Diagnostic[];
}

export type WorkerResponse =
  | { readonly id: number; readonly ok: true; readonly type: "font-request"; readonly requests: readonly FontRequest[] }
  | { readonly id: number; readonly ok: true; readonly type: "open"; readonly snapshot: WorkerSnapshot }
  | ({ readonly id: number; readonly ok: true; readonly type: "render"; readonly result: RenderResult }
    & WorkerDocumentDiagnosticUpdate)
  | ({ readonly id: number; readonly ok: true; readonly type: "hit-test"; readonly results: readonly HitResult[] }
    & WorkerDocumentDiagnosticUpdate)
  | ({ readonly id: number; readonly ok: true; readonly type: "get-object"; readonly object?: DocumentObject }
    & WorkerDocumentDiagnosticUpdate)
  | ({ readonly id: number; readonly ok: true; readonly type: "list-objects"; readonly objects: readonly DocumentObject[] }
    & WorkerDocumentDiagnosticUpdate)
  | ({ readonly id: number; readonly ok: true; readonly type: "search-text"; readonly results: readonly TextSearchResult[] }
    & WorkerDocumentDiagnosticUpdate)
  | { readonly id: number; readonly ok: false; readonly error: SerializedOfficeError };
