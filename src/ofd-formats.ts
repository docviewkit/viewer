import type { FormatPack, FormatPackCandidate } from "./format-pack.js";
import type { WasmSource } from "./types.js";
import { OfficeEngineError } from "./types.js";

/** Optional OFD fixed-document parser and renderer payload. */
export const ofdFormatPack: FormatPack = Object.freeze({
  async load(candidate: FormatPackCandidate): Promise<WasmSource> {
    if (candidate !== "ofd") {
      throw new OfficeEngineError(
        "UNSUPPORTED_FORMAT",
        `The OFD FormatPack cannot load the ${candidate} format family`,
      );
    }
    return new URL("./office-viewer-ofd.wasm", import.meta.url);
  },
});
