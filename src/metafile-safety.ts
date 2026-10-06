import { identifyImage } from "./image.js";
import type { OfficeImageMediaType } from "./image.js";
import { OfficeEngineError } from "./types.js";

const EMR_COMMENT = 70;
const EMR_BITBLT = 76;
const EMR_STRETCHDIBITS = 81;
const EMR_SETWINDOWEXTEX = 9;
const EMR_SETWINDOWORGEX = 10;
const EMR_SETVIEWPORTEXTEX = 11;
const EMR_SETVIEWPORTORGEX = 12;
const WMF_PLACEABLE_SIGNATURE = 0x9ac6_cdd7;
const META_SETWINDOWORG = 523;
const META_SETWINDOWEXT = 524;
const META_DIBCREATEPATTERNBRUSH = 0x0142;
const META_DIBSTRETCHBLT = 0x0b41;
const META_STRETCHDIB = 0x0f43;
const EMF_PLUS_SIGNATURE = 0x2b46_4d45;
const EMF_PLUS_OBJECT = 0x4008;
const EMF_PLUS_OBJECT_TYPE_IMAGE = 5;
const EMF_PLUS_IMAGE_TYPE_BITMAP = 1;
const EMF_PLUS_IMAGE_TYPE_METAFILE = 2;
const EMF_PLUS_BITMAP_TYPE_PIXELS = 1;
const EMF_PLUS_BITMAP_TYPE_COMPRESSED = 2;
const MAX_METAFILE_RECORDS = 200_000;
const MAX_EMF_PLUS_CONTINUATION_BYTES = 64 * 1024 * 1024;

interface MetafileBudget {
  readonly maxBytes: number;
  readonly maxPixels: number;
  bytes: number;
  pixels: number;
}

interface ContinuationSegment {
  readonly targetOffset: number;
  readonly sourceOffset: number;
  readonly length: number;
}

interface ObjectContinuation {
  readonly id: number;
  readonly objectType: number;
  readonly data: Uint8Array | undefined;
  readonly totalSize: number;
  readonly segments: ContinuationSegment[];
  offset: number;
}

function failure(code: string, message: string): never {
  throw new OfficeEngineError(code, message);
}

function u16le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 2 > bytes.length) failure("IMAGE_HEADER_INVALID", "Metafile image record is truncated");
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

function u32le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > bytes.length) failure("IMAGE_HEADER_INVALID", "Metafile image record is truncated");
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, true);
}

function i32le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > bytes.length) failure("IMAGE_HEADER_INVALID", "Metafile image record is truncated");
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getInt32(0, true);
}

function i16le(bytes: Uint8Array, offset: number): number {
  const value = u16le(bytes, offset);
  return value & 0x8000 ? value - 0x1_0000 : value;
}

function addWmfPlaceableBounds(bytes: Uint8Array, maxBytes: number): Uint8Array {
  if (bytes.length < 18) return bytes;
  const placeable = u32le(bytes, 0) === WMF_PLACEABLE_SIGNATURE;
  let source = bytes;
  let unitsPerInch = 1_440;
  if (placeable) {
    if (bytes.length < 40 || bytes.length > maxBytes) return bytes;
    let checksum = 0;
    for (let offset = 0; offset < 20; offset += 2) checksum ^= u16le(bytes, offset);
    if (checksum !== u16le(bytes, 20) || u16le(bytes, 14) === 0) return bytes;
    if (i16le(bytes, 6) !== i16le(bytes, 10) && i16le(bytes, 8) !== i16le(bytes, 12)) return bytes;
    unitsPerInch = u16le(bytes, 14);
    source = bytes.subarray(22);
  } else if (bytes.length > maxBytes - 22) {
    return bytes;
  }
  const type = u16le(source, 0);
  if ((type !== 1 && type !== 2) || u16le(source, 2) < 9) return bytes;
  let origin: readonly [number, number] = [0, 0];
  let extent: readonly [number, number] | undefined;
  let offset = u16le(source, 2) * 2;
  let records = 0;
  while (offset + 6 <= source.length && records < MAX_METAFILE_RECORDS) {
    const size = u32le(source, offset) * 2;
    if (size < 6 || offset + size > source.length) break;
    records += 1;
    const kind = u16le(source, offset + 4);
    if (kind === META_SETWINDOWORG && size >= 10) {
      origin = [i16le(source, offset + 8), i16le(source, offset + 6)];
    } else if (kind === META_SETWINDOWEXT && size >= 10) {
      const candidate = [i16le(source, offset + 8), i16le(source, offset + 6)] as const;
      if (candidate[0] > 0 && candidate[1] > 0) extent = candidate;
    }
    if (kind === 0) break;
    offset += size;
  }
  if (!extent) return bytes;
  const right = origin[0] + extent[0];
  const bottom = origin[1] + extent[1];
  if (right > 0x7fff || bottom > 0x7fff || origin[0] < -0x8000 || origin[1] < -0x8000) {
    return bytes;
  }
  const output = new Uint8Array(source.length + 22);
  const view = new DataView(output.buffer);
  view.setUint32(0, WMF_PLACEABLE_SIGNATURE, true);
  view.setInt16(6, origin[0], true);
  view.setInt16(8, origin[1], true);
  view.setInt16(10, right, true);
  view.setInt16(12, bottom, true);
  view.setUint16(14, unitsPerInch, true);
  let checksum = 0;
  for (let word = 0; word < 10; word += 1) checksum ^= view.getUint16(word * 2, true);
  view.setUint16(20, checksum, true);
  output.set(source, 22);
  return output;
}

function inspectWmfBitmapRecords(bytes: Uint8Array, budget: MetafileBudget): void {
  const headerOffset = u32le(bytes, 0) === WMF_PLACEABLE_SIGNATURE ? 22 : 0;
  if (bytes.length < headerOffset + 18) return;
  let offset = headerOffset + u16le(bytes, headerOffset + 2) * 2;
  for (let records = 0; offset + 6 <= bytes.length && records < MAX_METAFILE_RECORDS; records += 1) {
    const size = u32le(bytes, offset) * 2;
    if (size < 6 || offset + size > bytes.length) break;
    const kind = u16le(bytes, offset + 4);
    if (kind === META_DIBSTRETCHBLT || kind === META_STRETCHDIB || kind === META_DIBCREATEPATTERNBRUSH) {
      const dibOffset = offset + (kind === META_DIBCREATEPATTERNBRUSH ? 10 : kind === META_DIBSTRETCHBLT ? 26 : 28);
      if (dibOffset >= offset + size) failure("IMAGE_HEADER_INVALID", "Embedded WMF bitmap is truncated");
      const dib = bytes.subarray(dibOffset, offset + size);
      const identified = identifyImage(dib, "image/bmp", budget.maxPixels);
      chargeBudget(budget, dib.length, identified.pixels);
    }
    if (kind === 0) break;
    offset += size;
  }
}

function ascii(bytes: Uint8Array, offset: number, length: number): string {
  return String.fromCharCode(...bytes.subarray(offset, offset + length));
}

function embeddedRasterMediaType(bytes: Uint8Array): OfficeImageMediaType | undefined {
  if (bytes.length >= 8
    && bytes[0] === 0x89 && ascii(bytes, 1, 3) === "PNG"
    && bytes[4] === 0x0d && bytes[5] === 0x0a && bytes[6] === 0x1a && bytes[7] === 0x0a) {
    return "image/png";
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
  if (bytes.length >= 6 && (ascii(bytes, 0, 6) === "GIF87a" || ascii(bytes, 0, 6) === "GIF89a")) return "image/gif";
  if (bytes.length >= 12 && ascii(bytes, 0, 4) === "RIFF" && ascii(bytes, 8, 4) === "WEBP") return "image/webp";
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
  if (bytes.length >= 4 && [12, 16, 40, 52, 56, 64, 108, 124].includes(u32le(bytes, 0))) return "image/bmp";
  if (bytes.length >= 4
    && ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0)
      || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a))) {
    return "image/tiff";
  }
  if (bytes.length >= 4 && bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 1 && bytes[3] === 0) return "image/x-icon";
  if (bytes.length >= 4 && bytes[0] === 0x0a && bytes[2] === 1) return "image/x-pcx";
  if (bytes.length >= 12 && u32le(bytes, 0) === 0x0c00_0000 && ascii(bytes, 4, 4) === "jP  ") return "image/jp2";
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0x4f && bytes[2] === 0xff && bytes[3] === 0x51) return "image/j2k";
  return undefined;
}

function chargeBudget(budget: MetafileBudget, bytes: number, pixels: number): void {
  if (!Number.isSafeInteger(bytes) || bytes < 0 || budget.bytes + bytes > budget.maxBytes) {
    failure("IMAGE_SIZE_LIMIT", `Embedded metafile images exceed the cumulative ${budget.maxBytes}-byte limit`);
  }
  if (!Number.isSafeInteger(pixels) || pixels < 0 || budget.pixels + pixels > budget.maxPixels) {
    failure("IMAGE_DIMENSION_LIMIT", `Embedded metafile images exceed the cumulative ${budget.maxPixels}-pixel limit`);
  }
  budget.bytes += bytes;
  budget.pixels += pixels;
}

function inspectImageObject(data: Uint8Array, budget: MetafileBudget): boolean {
  if (data.length < 8) return false;
  const imageType = u32le(data, 4);
  if (imageType === EMF_PLUS_IMAGE_TYPE_METAFILE) return true;
  if (imageType !== EMF_PLUS_IMAGE_TYPE_BITMAP || data.length < 28) return false;

  const bitmapType = u32le(data, 24);
  if (bitmapType === EMF_PLUS_BITMAP_TYPE_PIXELS) {
    const width = i32le(data, 8);
    const height = i32le(data, 12);
    const stride = Math.abs(i32le(data, 16));
    if (width <= 0 || height <= 0 || stride === 0) return false;
    const pixels = width * height;
    const requiredBytes = stride * height;
    if (!Number.isSafeInteger(requiredBytes) || requiredBytes > data.length - 28) {
      failure("IMAGE_HEADER_INVALID", "Embedded EMF+ pixel bitmap exceeds its image object");
    }
    chargeBudget(budget, requiredBytes, pixels);
    return false;
  }
  if (bitmapType !== EMF_PLUS_BITMAP_TYPE_COMPRESSED) return false;

  const image = data.subarray(28);
  if (image.length === 0) return false;
  const mediaType = embeddedRasterMediaType(image);
  if (mediaType === undefined) {
    failure("IMAGE_FORMAT_UNSUPPORTED", "Embedded EMF+ compressed bitmap is not a supported raster image");
  }
  const identified = identifyImage(image, mediaType, budget.maxPixels);
  chargeBudget(budget, image.length, identified.pixels);
  return false;
}

function appendContinuation(
  continuation: ObjectContinuation,
  bytes: Uint8Array,
  sourceOffset: number,
  sourceLength: number,
): void {
  const length = Math.min(sourceLength, continuation.totalSize - continuation.offset);
  if (length <= 0) return;
  if (continuation.data !== undefined) {
    continuation.data.set(bytes.subarray(sourceOffset, sourceOffset + length), continuation.offset);
  }
  continuation.segments.push({ targetOffset: continuation.offset, sourceOffset, length });
  continuation.offset += length;
}

function sourceOffsetFor(continuation: ObjectContinuation, targetOffset: number): number | undefined {
  const segment = continuation.segments.find(({ targetOffset: start, length }) => (
    targetOffset >= start && targetOffset < start + length
  ));
  return segment === undefined ? undefined : segment.sourceOffset + targetOffset - segment.targetOffset;
}

function inspectEmfPlusRecords(
  bytes: Uint8Array,
  start: number,
  length: number,
  budget: MetafileBudget,
  disableNested: (imageTypeOffsets: readonly number[]) => void,
): void {
  const end = start + length;
  let offset = start;
  let recordCount = 0;
  let continuation: ObjectContinuation | undefined;
  while (offset + 12 <= end && recordCount < MAX_METAFILE_RECORDS) {
    const recordType = u16le(bytes, offset);
    const flags = u16le(bytes, offset + 2);
    const recordSize = u32le(bytes, offset + 4);
    const dataSize = u32le(bytes, offset + 8);
    if (recordSize < 12 || recordSize % 4 !== 0 || offset + recordSize > end) break;
    recordCount += 1;
    if (dataSize > recordSize - 12) {
      failure("IMAGE_HEADER_INVALID", "EMF+ image record data exceeds its record boundary");
    }
    if (recordType === EMF_PLUS_OBJECT) {
      const isContinuation = (flags & 0x8000) !== 0;
      const id = flags & 0xff;
      const objectType = (flags >> 8) & 0x7f;
      const dataOffset = offset + 12;
      if (isContinuation) {
        if (continuation === undefined) {
          if (dataSize >= 4) {
            const totalSize = u32le(bytes, dataOffset);
            if (totalSize === 0 || totalSize > bytes.length - dataOffset - 4) {
              failure("IMAGE_HEADER_INVALID", "EMF+ continued object size is invalid");
            }
            if (totalSize > MAX_EMF_PLUS_CONTINUATION_BYTES || totalSize > budget.maxBytes - budget.bytes) {
              failure("IMAGE_SIZE_LIMIT", "Embedded metafile object exceeds its continuation byte budget");
            }
            continuation = {
              id,
              objectType,
              data: objectType === EMF_PLUS_OBJECT_TYPE_IMAGE ? new Uint8Array(totalSize) : undefined,
              totalSize,
              segments: [],
              offset: 0,
            };
            appendContinuation(continuation, bytes, dataOffset + 4, dataSize - 4);
          }
        } else {
          appendContinuation(continuation, bytes, dataOffset, dataSize);
        }
      } else if (continuation !== undefined && id === continuation.id) {
        appendContinuation(continuation, bytes, dataOffset, dataSize);
        if (continuation.objectType === EMF_PLUS_OBJECT_TYPE_IMAGE && continuation.data !== undefined
          && inspectImageObject(continuation.data, budget)) {
          const completed = continuation;
          const imageTypeOffsets = [4, 5, 6, 7]
            .map((targetOffset) => sourceOffsetFor(completed, targetOffset))
            .filter((sourceOffset): sourceOffset is number => sourceOffset !== undefined);
          disableNested(imageTypeOffsets);
        }
        continuation = undefined;
      } else if (objectType === EMF_PLUS_OBJECT_TYPE_IMAGE) {
        const data = bytes.subarray(dataOffset, dataOffset + dataSize);
        if (inspectImageObject(data, budget)) {
          disableNested([dataOffset + 4, dataOffset + 5, dataOffset + 6, dataOffset + 7]);
        }
      }
    }
    offset += recordSize;
  }
}

function inspectGdiBitmapRecord(
  bytes: Uint8Array,
  recordOffset: number,
  recordSize: number,
  bmiFieldOffset: number,
  budget: MetafileBudget,
): void {
  if (recordSize < bmiFieldOffset + 16) {
    failure("IMAGE_HEADER_INVALID", "Embedded EMF bitmap record is truncated");
  }
  const bmiRelative = u32le(bytes, recordOffset + bmiFieldOffset);
  const bmiLength = u32le(bytes, recordOffset + bmiFieldOffset + 4);
  const bitsRelative = u32le(bytes, recordOffset + bmiFieldOffset + 8);
  const bitsLength = u32le(bytes, recordOffset + bmiFieldOffset + 12);
  if (bmiRelative === 0 || bmiLength === 0 || bitsRelative === 0 || bitsLength === 0) return;
  if (bmiRelative > recordSize || bmiLength > recordSize - bmiRelative
    || bitsRelative > recordSize || bitsLength > recordSize - bitsRelative) {
    failure("IMAGE_HEADER_INVALID", "Embedded EMF bitmap points outside its record");
  }
  const header = bytes.subarray(recordOffset + bmiRelative, recordOffset + bmiRelative + bmiLength);
  const identified = identifyImage(header, "image/bmp", budget.maxPixels);
  chargeBudget(budget, bmiLength + bitsLength, identified.pixels);
}

/**
 * Normalizes explicit GDI windows before conversion. EMF raster allocations
 * are preflighted and its unbounded recursive metafile path stays disabled.
 */
export function secureMetafileForConversion(
  bytes: Uint8Array,
  format: "emf" | "wmf",
  maxBytes: number,
  maxPixels: number,
  targetWidth?: number,
  targetHeight?: number,
  maxViewportScale = 1,
): Uint8Array {
  const budget: MetafileBudget = { maxBytes, maxPixels, bytes: 0, pixels: 0 };
  if (format === "wmf") {
    const normalized = addWmfPlaceableBounds(bytes, maxBytes);
    inspectWmfBitmapRecords(normalized, budget);
    return normalized;
  }
  let sanitized: Uint8Array | undefined;
  const contentBounds = bytes.length >= 88 && u32le(bytes, 0) === 1 && u32le(bytes, 4) >= 88
    ? {
        left: i32le(bytes, 8),
        top: i32le(bytes, 12),
        right: i32le(bytes, 16),
        bottom: i32le(bytes, 20),
      }
    : undefined;
  const frameWidth = contentBounds ? i32le(bytes, 32) - i32le(bytes, 24) : 0;
  const frameHeight = contentBounds ? i32le(bytes, 36) - i32le(bytes, 28) : 0;
  const deviceWidth = contentBounds ? i32le(bytes, 72) : 0;
  const deviceHeight = contentBounds ? i32le(bytes, 76) : 0;
  const millimetersWidth = contentBounds ? i32le(bytes, 80) : 0;
  const millimetersHeight = contentBounds ? i32le(bytes, 84) : 0;
  const framePixelWidth = deviceWidth > 0 && millimetersWidth > 0
    ? Math.round(frameWidth * deviceWidth / (millimetersWidth * 100))
    : 0;
  const framePixelHeight = deviceHeight > 0 && millimetersHeight > 0
    ? Math.round(frameHeight * deviceHeight / (millimetersHeight * 100))
    : 0;
  const usesFrameBounds = contentBounds !== undefined
    && framePixelWidth > 0 && framePixelHeight > 0
    && contentBounds.left >= 0 && contentBounds.top >= 0
    && contentBounds.right <= framePixelWidth && contentBounds.bottom <= framePixelHeight;
  const headerBounds = usesFrameBounds
    ? { left: 0, top: 0, right: framePixelWidth, bottom: framePixelHeight }
    : contentBounds;
  const boundsWidth = headerBounds ? headerBounds.right - headerBounds.left : 0;
  const boundsHeight = headerBounds ? headerBounds.bottom - headerBounds.top : 0;
  const viewportScale = Number.isFinite(maxViewportScale) && maxViewportScale > 0
    ? maxViewportScale
    : 1;
  const canvasScale = headerBounds && boundsWidth > 0 && boundsHeight > 0
    && targetWidth && targetWidth > 0 && targetHeight && targetHeight > 0
    ? Math.min(viewportScale, targetWidth / boundsWidth, targetHeight / boundsHeight)
    : undefined;
  const scaleX = canvasScale === undefined
    ? undefined
    : Math.max(1, Math.round(boundsWidth * canvasScale)) / boundsWidth;
  const scaleY = canvasScale === undefined
    ? undefined
    : Math.max(1, Math.round(boundsHeight * canvasScale)) / boundsHeight;
  const writeI32 = (offset: number, value: number) => {
    const rounded = Math.round(value);
    if (!Number.isSafeInteger(rounded) || rounded < -0x8000_0000 || rounded > 0x7fff_ffff) {
      failure("IMAGE_HEADER_INVALID", "Scaled EMF viewport coordinate exceeds the signed 32-bit range");
    }
    sanitized ??= bytes.slice();
    new DataView(sanitized.buffer, sanitized.byteOffset + offset, 4)
      .setInt32(0, rounded, true);
  };
  const scaledViewportExtent = (
    value: number,
    boundsExtent: number,
    scale: number,
    contentEndsAtExclusiveEdge: boolean,
  ) => {
    if (Math.abs(value) !== boundsExtent + 1
      && !(contentEndsAtExclusiveEdge && Math.abs(value) === boundsExtent)) return value * scale;
    const scaled = Math.sign(value) * Math.max(1, Math.round(boundsExtent * scale));
    return scaled > 0 ? scaled - 1 : scaled + 1;
  };
  const disableNested = (imageTypeOffsets: readonly number[]) => {
    if (imageTypeOffsets.length === 0) return;
    sanitized ??= bytes.slice();
    for (const imageTypeOffset of imageTypeOffsets) sanitized[imageTypeOffset] = 0;
  };
  let windowOrigin: readonly [number, number] | undefined;
  let windowExtent: readonly [number, number] | undefined;
  let explicitViewport = false;

  let offset = 0;
  let recordCount = 0;
  while (offset + 8 <= bytes.length && recordCount < MAX_METAFILE_RECORDS) {
    const recordType = u32le(bytes, offset);
    const recordSize = u32le(bytes, offset + 4);
    if (recordSize < 8 || recordSize % 4 !== 0 || offset + recordSize > bytes.length) break;
    recordCount += 1;
    if (recordType === EMR_COMMENT && recordSize >= 16) {
      const commentLength = u32le(bytes, offset + 8);
      if (commentLength > recordSize - 12) {
        failure("IMAGE_HEADER_INVALID", "EMF comment data exceeds its record boundary");
      }
      if (commentLength > 4 && u32le(bytes, offset + 12) === EMF_PLUS_SIGNATURE) {
        inspectEmfPlusRecords(bytes, offset + 16, commentLength - 4, budget, disableNested);
      }
    } else if (recordType === EMR_BITBLT) {
      inspectGdiBitmapRecord(bytes, offset, recordSize, 84, budget);
    } else if (recordType === EMR_STRETCHDIBITS) {
      inspectGdiBitmapRecord(bytes, offset, recordSize, 48, budget);
    } else if (recordType === EMR_SETWINDOWORGEX && recordSize >= 16) {
      windowOrigin = [i32le(bytes, offset + 8), i32le(bytes, offset + 12)];
    } else if (recordType === EMR_SETWINDOWEXTEX && recordSize >= 16) {
      windowExtent = [i32le(bytes, offset + 8), i32le(bytes, offset + 12)];
    } else if (recordType === EMR_SETVIEWPORTORGEX && recordSize >= 16
      && headerBounds && scaleX !== undefined && scaleY !== undefined) {
      explicitViewport = true;
      writeI32(offset + 8, (i32le(bytes, offset + 8) - headerBounds.left) * scaleX);
      writeI32(offset + 12, (i32le(bytes, offset + 12) - headerBounds.top) * scaleY);
    } else if (recordType === EMR_SETVIEWPORTEXTEX && recordSize >= 16
      && headerBounds && scaleX !== undefined && scaleY !== undefined) {
      explicitViewport = true;
      writeI32(offset + 8, scaledViewportExtent(
        i32le(bytes, offset + 8),
        boundsWidth,
        scaleX,
        usesFrameBounds && contentBounds?.right === headerBounds.right - 1,
      ));
      writeI32(offset + 12, scaledViewportExtent(
        i32le(bytes, offset + 12),
        boundsHeight,
        scaleY,
        usesFrameBounds && contentBounds?.bottom === headerBounds.bottom - 1,
      ));
    }
    offset += recordSize;
  }
  if (!explicitViewport && headerBounds && windowOrigin && windowExtent
    && windowExtent[0] > 0 && windowExtent[1] > 0) {
    const right = windowOrigin[0] + windowExtent[0];
    const bottom = windowOrigin[1] + windowExtent[1];
    const expandsBounds = windowOrigin[0] < headerBounds.left || windowOrigin[1] < headerBounds.top
      || right > headerBounds.right || bottom > headerBounds.bottom;
    if (expandsBounds && right <= 0x7fff_ffff && bottom <= 0x7fff_ffff) {
      writeI32(8, windowOrigin[0]);
      writeI32(12, windowOrigin[1]);
      writeI32(16, right);
      writeI32(20, bottom);
    }
  }
  return sanitized ?? bytes;
}
