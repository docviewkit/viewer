import type { FormatPack, FormatPackCandidate } from "./format-pack.js";
import { OfficeEngineError } from "./types.js";
import type { WasmSource } from "./types.js";

/**
 * Optional WPS-only FormatPack for Writer (`.wps`), Spreadsheets (`.et`), and
 * Presentation (`.dps`) OLE/CFB documents.
 */
export const wpsFormatPack: FormatPack = Object.freeze({
  async load(candidate: FormatPackCandidate): Promise<WasmSource> {
    if (candidate !== "wps") {
      throw new OfficeEngineError(
        "UNSUPPORTED_FORMAT",
        `The WPS FormatPack cannot load the ${candidate} format family`,
      );
    }
    return new URL("./office-viewer-legacy-office.wasm", import.meta.url);
  },
});
