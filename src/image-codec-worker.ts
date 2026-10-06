import { decodeOfficeImagePayload } from "./image-codecs.js";
import type { OfficeImageCodecFont, OfficeImageCodecInput, OfficeImageCodecOutput } from "./image-codecs.js";
import { OfficeEngineError } from "./types.js";

interface CodecRequest extends OfficeImageCodecInput {
  readonly id: number;
}

const MAX_CODEC_FONT_BYTES = 256 * 1024 * 1024;

interface WorkerScope {
  onmessage: ((event: MessageEvent<unknown>) => void) | null;
  postMessage(message: unknown, transfer: Transferable[]): void;
}

function codecFonts(value: unknown): readonly OfficeImageCodecFont[] {
  if (!Array.isArray(value)) {
    throw new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec fonts are invalid");
  }
  let totalBytes = 0;
  return value.map((font) => {
    if (typeof font !== "object" || font === null) {
      throw new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec font is invalid");
    }
    const candidate = font as Partial<OfficeImageCodecFont>;
    if (typeof candidate.family !== "string"
      || candidate.family.length === 0
      || candidate.family.length > 256
      || /[\0-\x1f\x7f]/u.test(candidate.family)
      || !(candidate.bytes instanceof ArrayBuffer)
      || candidate.bytes.byteLength === 0) {
      throw new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec font is invalid");
    }
    totalBytes += candidate.bytes.byteLength;
    if (!Number.isSafeInteger(totalBytes) || totalBytes > MAX_CODEC_FONT_BYTES) {
      throw new OfficeEngineError(
        "FONT_BYTES_LIMIT",
        `Image codec fonts require ${totalBytes} bytes; hard limit is ${MAX_CODEC_FONT_BYTES}`,
      );
    }
    return { family: candidate.family, bytes: candidate.bytes };
  });
}

function request(value: unknown): CodecRequest {
  if (typeof value !== "object" || value === null) {
    throw new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec request is not an object");
  }
  const candidate = value as Partial<CodecRequest>;
  if (candidate.id !== 1
    || typeof candidate.format !== "string"
    || typeof candidate.mediaType !== "string"
    || !(candidate.bytes instanceof ArrayBuffer)
    || typeof candidate.maxPixels !== "number"
    || typeof candidate.maxBytes !== "number"
    || typeof candidate.maxCompressionRatio !== "number"
    || (candidate.targetWidth !== undefined && typeof candidate.targetWidth !== "number")
    || (candidate.targetHeight !== undefined && typeof candidate.targetHeight !== "number")) {
    throw new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec request is invalid");
  }
  return { ...candidate, fonts: codecFonts(candidate.fonts) } as CodecRequest;
}

async function imageBitmap(output: OfficeImageCodecOutput): Promise<ImageBitmap> {
  if (typeof createImageBitmap !== "function") {
    throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "createImageBitmap is required in the image codec Worker");
  }
  if (output.kind === "rgba") {
    const pixels = new Uint8ClampedArray(output.data.length);
    pixels.set(output.data);
    return createImageBitmap(new ImageData(pixels, output.width, output.height));
  }
  return createImageBitmap(new Blob([output.data.slice().buffer], { type: output.mediaType }));
}

const scope = globalThis as unknown as WorkerScope;
scope.onmessage = (event) => {
  void (async () => {
    try {
      const input = request(event.data);
      const output = await decodeOfficeImagePayload(input);
      const bitmap = await imageBitmap(output);
      const pixels = bitmap.width * bitmap.height;
      if (bitmap.width !== output.width || bitmap.height !== output.height
        || !Number.isSafeInteger(pixels) || pixels > input.maxPixels) {
        bitmap.close();
        throw new OfficeEngineError("IMAGE_DIMENSION_LIMIT", "Decoded bitmap dimensions failed codec validation");
      }
      scope.postMessage({ id: input.id, ok: true, bitmap, approximate: output.approximate }, [bitmap]);
    } catch (cause) {
      const error = cause instanceof OfficeEngineError
        ? cause
        : new OfficeEngineError("IMAGE_DECODE_FAILED", "Office image codec failed", { cause });
      scope.postMessage({ id: 1, ok: false, code: error.code, message: error.message }, []);
    }
  })();
};
