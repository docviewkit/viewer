import type { FormatPack, FormatPackCandidate } from "./format-pack.js";
import type { OfficeEngine, WasmSource } from "./types.js";

function wasmUrl(candidate: FormatPackCandidate): URL {
  switch (candidate) {
    case "odf":
      return new URL("./office-viewer-odf.wasm", import.meta.url);
    case "iwork":
      return new URL("./office-viewer-pdf.wasm", import.meta.url);
    case "legacy-office":
      return new URL("./office-viewer-legacy-office.wasm", import.meta.url);
    case "wps":
      return new URL("./office-viewer-legacy-office.wasm", import.meta.url);
    case "pdf":
      return new URL("./office-viewer-pdf.wasm", import.meta.url);
    case "xps":
      return new URL("./office-viewer-xps.wasm", import.meta.url);
    case "ofd":
      return new URL("./office-viewer-ofd.wasm", import.meta.url);
  }
}

/**
 * Optional subpath payload for native format Wasm:
 * `formatPack: () => import("@docviewkit/sdk/extended-formats")
 *   .then(({ extendedFormatPack }) => extendedFormatPack)`.
 *
 * Resolving a URL does not fetch or instantiate any Wasm module. The runtime
 * calls load only after content dispatch selects one candidate.
 */
export const extendedFormatPack: FormatPack = Object.freeze({
  async load(candidate: FormatPackCandidate): Promise<WasmSource | OfficeEngine> {
    return wasmUrl(candidate);
  },
});
