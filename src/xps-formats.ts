import type { FormatPack, FormatPackCandidate } from "./format-pack.js";
import type { WasmSource } from "./types.js";
import { OfficeEngineError } from "./types.js";

/** Optional XPS/OpenXPS parser and renderer payload. */
export const xpsFormatPack: FormatPack = Object.freeze({
  async load(candidate: FormatPackCandidate): Promise<WasmSource> {
    if (candidate !== "xps") {
      throw new OfficeEngineError(
        "UNSUPPORTED_FORMAT",
        `The XPS FormatPack cannot load the ${candidate} format family`,
      );
    }
    return new URL("./office-viewer-xps.wasm", import.meta.url);
  },
});
