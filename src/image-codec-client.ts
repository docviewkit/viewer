import type { IdentifiedImage, OfficeImageMediaType } from "./image.js";
import type { OfficeImageCodecFont } from "./image-codecs.js";
import type { ResourceLimits } from "./types.js";
import { OfficeEngineError } from "./types.js";

const CODEC_TIMEOUT_MS = 9_000;
const CODEC_WORKER_IDLE_MS = 30_000;
const CODEC_WORKER_LIMIT = 8;

export interface DecodedOfficeImage {
  readonly bitmap: ImageBitmap;
  readonly approximate: boolean;
}

interface CodecRequest {
  readonly id: number;
  readonly format: IdentifiedImage["format"];
  readonly mediaType: OfficeImageMediaType;
  readonly bytes: ArrayBuffer;
  readonly maxPixels: number;
  readonly maxBytes: number;
  readonly maxCompressionRatio: number;
  readonly targetWidth?: number;
  readonly targetHeight?: number;
  readonly fonts: readonly OfficeImageCodecFont[];
}

interface CodecSuccess {
  readonly id: number;
  readonly ok: true;
  readonly bitmap: ImageBitmap;
  readonly approximate: boolean;
}

interface CodecFailure {
  readonly id: number;
  readonly ok: false;
  readonly code: string;
  readonly message: string;
}

type CodecResponse = CodecSuccess | CodecFailure;

function isCodecResponse(value: unknown): value is CodecResponse {
  if (typeof value !== "object" || value === null) return false;
  const candidate = value as Partial<CodecResponse>;
  if (candidate.id !== 1 || typeof candidate.ok !== "boolean") return false;
  if (candidate.ok) {
    const bitmap = candidate.bitmap as Partial<ImageBitmap> | undefined;
    return bitmap !== undefined
      && typeof bitmap.width === "number"
      && Number.isSafeInteger(bitmap.width)
      && bitmap.width > 0
      && typeof bitmap.height === "number"
      && Number.isSafeInteger(bitmap.height)
      && bitmap.height > 0
      && typeof bitmap.close === "function"
      && typeof candidate.approximate === "boolean";
  }
  const failure = candidate as Partial<CodecFailure>;
  return typeof failure.code === "string" && typeof failure.message === "string";
}

function nativeFormat(format: IdentifiedImage["format"]): boolean {
  return format === "png" || format === "jpeg" || format === "gif" || format === "webp" || format === "bmp";
}

interface IdleCodecWorker {
  readonly worker: Worker;
  readonly timer: ReturnType<typeof setTimeout>;
}

interface CodecWorkerWaiter {
  readonly resolve: (worker: Worker) => void;
  readonly reject: (cause: unknown) => void;
  readonly timer: ReturnType<typeof setTimeout>;
}

const idleCodecWorkers: IdleCodecWorker[] = [];
const codecWorkerWaiters: CodecWorkerWaiter[] = [];
const codecWorkerConstructors = new WeakMap<Worker, typeof Worker>();
let codecWorkerCount = 0;

function unrefTimer(timer: ReturnType<typeof setTimeout>): void {
  (timer as unknown as { unref?: () => void }).unref?.();
}

function codecTimeout(format: IdentifiedImage["format"]): OfficeEngineError {
  return new OfficeEngineError(
    "IMAGE_DECODE_TIMEOUT",
    `Office ${format.toUpperCase()} decode exceeded ${CODEC_TIMEOUT_MS}ms`,
  );
}

function currentWorkerConstructor(): typeof Worker | undefined {
  return typeof Worker === "undefined" ? undefined : Worker;
}

function createCodecWorker(): Worker {
  const worker = new Worker(new URL("./image-codec-worker.js", import.meta.url), {
    type: "module",
    name: "office-image-codec",
  });
  codecWorkerConstructors.set(worker, Worker);
  codecWorkerCount += 1;
  return worker;
}

function takeCodecWorkerWaiter(): CodecWorkerWaiter | undefined {
  const waiter = codecWorkerWaiters.shift();
  if (waiter !== undefined) clearTimeout(waiter.timer);
  return waiter;
}

function acquireCodecWorker(
  deadline: number,
  format: IdentifiedImage["format"],
): Promise<Worker> {
  while (idleCodecWorkers.length > 0) {
    const idle = idleCodecWorkers.pop()!;
    clearTimeout(idle.timer);
    if (codecWorkerConstructors.get(idle.worker) === currentWorkerConstructor()) {
      return Promise.resolve(idle.worker);
    }
    codecWorkerCount -= 1;
    idle.worker.terminate();
  }
  if (codecWorkerCount < CODEC_WORKER_LIMIT) {
    try {
      return Promise.resolve(createCodecWorker());
    } catch (cause) {
      return Promise.reject(cause);
    }
  }
  const remaining = deadline - Date.now();
  if (remaining <= 0) return Promise.reject(codecTimeout(format));
  return new Promise<Worker>((resolve, reject) => {
    let waiter: CodecWorkerWaiter;
    const timer = setTimeout(() => {
      const index = codecWorkerWaiters.indexOf(waiter);
      if (index >= 0) codecWorkerWaiters.splice(index, 1);
      reject(codecTimeout(format));
    }, remaining);
    waiter = { resolve, reject, timer };
    codecWorkerWaiters.push(waiter);
  });
}

function releaseCodecWorker(worker: Worker): void {
  worker.onmessage = null;
  worker.onerror = null;
  worker.onmessageerror = null;
  if (codecWorkerConstructors.get(worker) !== currentWorkerConstructor()) {
    discardCodecWorker(worker);
    return;
  }
  const waiter = takeCodecWorkerWaiter();
  if (waiter !== undefined) {
    waiter.resolve(worker);
    return;
  }
  const timer = setTimeout(() => {
    const index = idleCodecWorkers.findIndex((candidate) => candidate.worker === worker);
    if (index < 0) return;
    idleCodecWorkers.splice(index, 1);
    codecWorkerCount -= 1;
    worker.terminate();
  }, CODEC_WORKER_IDLE_MS);
  unrefTimer(timer);
  idleCodecWorkers.push({ worker, timer });
}

function discardCodecWorker(worker: Worker): void {
  worker.onmessage = null;
  worker.onerror = null;
  worker.onmessageerror = null;
  worker.terminate();
  codecWorkerCount -= 1;
  for (let waiter = takeCodecWorkerWaiter(); waiter !== undefined; waiter = takeCodecWorkerWaiter()) {
    try {
      waiter.resolve(createCodecWorker());
      return;
    } catch (cause) {
      waiter.reject(cause);
    }
  }
}

async function nativeDecode(bytes: Uint8Array, mediaType: string): Promise<ImageBitmap> {
  if (typeof createImageBitmap === "undefined") {
    throw new OfficeEngineError("UNSUPPORTED_ENVIRONMENT", "createImageBitmap is required for image rendering");
  }
  return createImageBitmap(new Blob([bytes.slice().buffer], { type: mediaType }));
}

function codecDecode(
  bytes: Uint8Array,
  image: IdentifiedImage,
  limits: ResourceLimits,
  fonts: readonly OfficeImageCodecFont[],
  target?: { readonly width: number; readonly height: number },
): Promise<DecodedOfficeImage> {
  if (typeof Worker === "undefined") {
    return Promise.reject(new OfficeEngineError(
      "UNSUPPORTED_ENVIRONMENT",
      `A module Worker is required to decode Office ${image.format.toUpperCase()} images safely`,
    ));
  }
  const deadline = Date.now() + CODEC_TIMEOUT_MS;
  return acquireCodecWorker(deadline, image.format).then((worker) => new Promise<DecodedOfficeImage>((resolve, reject) => {
    const buffer = bytes.slice().buffer;
    const request: CodecRequest = {
      id: 1,
      format: image.format,
      mediaType: image.mediaType,
      bytes: buffer,
      maxPixels: limits.imagePixels,
      maxBytes: limits.entryBytes,
      maxCompressionRatio: limits.compressionRatio,
      ...(target === undefined ? {} : {
        targetWidth: target.width,
        targetHeight: target.height,
      }),
      fonts: image.format === "svg"
        ? fonts.map(({ family, bytes: fontBytes }) => ({ family, bytes: fontBytes }))
        : [],
    };
    let settled = false;
    const finish = (reuse: boolean, action: () => void) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (reuse) releaseCodecWorker(worker);
      else discardCodecWorker(worker);
      action();
    };
    const timer = setTimeout(
      () => finish(false, () => reject(codecTimeout(image.format))),
      Math.max(0, deadline - Date.now()),
    );
    worker.onmessage = (event: MessageEvent<unknown>) => {
      const response = event.data;
      if (!isCodecResponse(response)) {
        finish(false, () => reject(new OfficeEngineError("IMAGE_CODEC_PROTOCOL_INVALID", "Office image codec returned an invalid response")));
      } else if (!response.ok) {
        finish(true, () => reject(new OfficeEngineError(response.code, response.message)));
      } else {
        finish(true, () => resolve({ bitmap: response.bitmap, approximate: response.approximate }));
      }
    };
    worker.onerror = (event: ErrorEvent) => {
      event.preventDefault();
      finish(false, () => reject(new OfficeEngineError(
        "IMAGE_CODEC_WORKER_FAILED",
        event.message || "Office image codec Worker failed",
        { cause: event.error },
      )));
    };
    worker.onmessageerror = () => finish(false, () => reject(new OfficeEngineError(
      "IMAGE_CODEC_PROTOCOL_INVALID",
      "Office image codec response could not be deserialized",
    )));
    try {
      worker.postMessage(request, [buffer]);
    } catch (cause) {
      finish(false, () => reject(new OfficeEngineError("IMAGE_CODEC_WORKER_FAILED", "Could not send image bytes to the codec Worker", { cause })));
    }
  }), (cause) => Promise.reject(cause instanceof OfficeEngineError
    ? cause
    : new OfficeEngineError(
      "IMAGE_CODEC_WORKER_FAILED",
      "Could not start the Office image codec Worker",
      { cause },
    )));
}

export async function decodeOfficeImage(
  bytes: Uint8Array,
  image: IdentifiedImage,
  limits: ResourceLimits,
  fonts: readonly OfficeImageCodecFont[] = [],
  target?: { readonly width: number; readonly height: number },
): Promise<DecodedOfficeImage> {
  if (nativeFormat(image.format)) {
    try {
      return { bitmap: await nativeDecode(bytes, image.mediaType), approximate: false };
    } catch (cause) {
      if (image.format !== "bmp") throw cause;
      return codecDecode(bytes, image, limits, fonts, target);
    }
  }
  return codecDecode(bytes, image, limits, fonts, target);
}
