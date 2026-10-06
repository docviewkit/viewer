import { OfficeEngineError } from "./types.js";

export type OfficeImageFormat =
  | "png"
  | "jpeg"
  | "gif"
  | "webp"
  | "bmp"
  | "dib"
  | "svg"
  | "tiff"
  | "jxr"
  | "ico"
  | "pcx"
  | "jp2"
  | "j2k"
  | "emf"
  | "wmf"
  | "emz"
  | "wmz"
  | "pdf-raster";

export type OfficeImageMediaType =
  | "image/png"
  | "image/jpeg"
  | "image/gif"
  | "image/webp"
  | "image/bmp"
  | "image/svg+xml"
  | "image/tiff"
  | "image/vnd.ms-photo"
  | "image/x-icon"
  | "image/x-pcx"
  | "image/jp2"
  | "image/j2k"
  | "image/x-emf"
  | "image/x-wmf"
  | "image/x-emz"
  | "image/x-wmz"
  | "image/x-officeviewer-pdf-raster";

export interface ImageInfo {
  readonly mediaType: OfficeImageMediaType;
  readonly width: number;
  readonly height: number;
  readonly pixels: number;
}

export interface IdentifiedImage extends ImageInfo {
  readonly format: OfficeImageFormat;
  readonly dimensionsKnown: boolean;
}

function ascii(bytes: Uint8Array, offset: number, length: number): string {
  return String.fromCharCode(...bytes.subarray(offset, offset + length));
}

function u16be(bytes: Uint8Array, offset: number): number {
  return (bytes[offset]! << 8) | bytes[offset + 1]!;
}

function u16le(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

function i16le(bytes: Uint8Array, offset: number): number {
  const value = u16le(bytes, offset);
  return value & 0x8000 ? value - 0x1_0000 : value;
}

function u24le(bytes: Uint8Array, offset: number): number {
  return bytes[offset]! | (bytes[offset + 1]! << 8) | (bytes[offset + 2]! << 16);
}

function u32be(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, false);
}

function u32le(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, true);
}

function i32le(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getInt32(0, true);
}

function boundedInteger(value: bigint, label: string): number {
  if (value < 0n || value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", `${label} exceeds the supported integer range`);
  }
  return Number(value);
}

function makeInfo(
  format: OfficeImageFormat,
  mediaType: OfficeImageMediaType,
  width: number,
  height: number,
  dimensionsKnown = true,
): IdentifiedImage {
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width <= 0 || height <= 0) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "Image dimensions are invalid");
  }
  const pixels = width * height;
  if (!Number.isSafeInteger(pixels)) {
    throw new OfficeEngineError("IMAGE_DIMENSION_LIMIT", "Image dimensions overflow the pixel counter");
  }
  return { format, mediaType, width, height, pixels, dimensionsKnown };
}

function png(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 24) return undefined;
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (!signature.every((value, index) => bytes[index] === value) || ascii(bytes, 12, 4) !== "IHDR") return undefined;
  return makeInfo("png", "image/png", u32be(bytes, 16), u32be(bytes, 20));
}

function gif(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 10) return undefined;
  const header = ascii(bytes, 0, 6);
  if (header !== "GIF87a" && header !== "GIF89a") return undefined;
  return makeInfo("gif", "image/gif", u16le(bytes, 6), u16le(bytes, 8));
}

function jpeg(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  let offset = 2;
  while (offset + 3 < bytes.length) {
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1;
    if (offset >= bytes.length) break;
    const marker = bytes[offset++]!;
    if (marker === 0xd9 || marker === 0xda) break;
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 2 > bytes.length) break;
    const length = u16be(bytes, offset);
    if (length < 2 || offset + length > bytes.length) break;
    const isStartOfFrame = (marker >= 0xc0 && marker <= 0xc3)
      || (marker >= 0xc5 && marker <= 0xc7)
      || (marker >= 0xc9 && marker <= 0xcb)
      || (marker >= 0xcd && marker <= 0xcf);
    if (isStartOfFrame) {
      if (length < 7) break;
      return makeInfo("jpeg", "image/jpeg", u16be(bytes, offset + 5), u16be(bytes, offset + 3));
    }
    offset += length;
  }
  throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG has no valid start-of-frame header");
}

function webp(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 16 || ascii(bytes, 0, 4) !== "RIFF" || ascii(bytes, 8, 4) !== "WEBP") return undefined;
  const kind = ascii(bytes, 12, 4);
  if (kind === "VP8X") {
    if (bytes.length < 30) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "WebP extended header is truncated");
    return makeInfo("webp", "image/webp", u24le(bytes, 24) + 1, u24le(bytes, 27) + 1);
  }
  if (kind === "VP8L") {
    if (bytes.length < 25 || bytes[20] !== 0x2f) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "WebP lossless signature is invalid");
    }
    const b1 = bytes[21]!;
    const b2 = bytes[22]!;
    const b3 = bytes[23]!;
    const b4 = bytes[24]!;
    return makeInfo(
      "webp",
      "image/webp",
      1 + b1 + ((b2 & 0x3f) << 8),
      1 + ((b2 & 0xc0) >> 6) + (b3 << 2) + ((b4 & 0x0f) << 10),
    );
  }
  if (kind === "VP8 ") {
    if (bytes.length < 30 || bytes[23] !== 0x9d || bytes[24] !== 0x01 || bytes[25] !== 0x2a) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "WebP lossy frame header is invalid");
    }
    return makeInfo("webp", "image/webp", u16le(bytes, 26) & 0x3fff, u16le(bytes, 28) & 0x3fff);
  }
  throw new OfficeEngineError("IMAGE_HEADER_INVALID", `Unsupported WebP chunk ${kind}`);
}

function bitmap(bytes: Uint8Array): IdentifiedImage | undefined {
  const fileHeader = bytes.length >= 14 && bytes[0] === 0x42 && bytes[1] === 0x4d;
  const offset = fileHeader ? 14 : 0;
  if (bytes.length < offset + 12) return undefined;
  const dibSize = u32le(bytes, offset);
  if (![12, 16, 40, 52, 56, 64, 108, 124].includes(dibSize)) return undefined;
  let width: number;
  let height: number;
  if (dibSize === 12) {
    width = u16le(bytes, offset + 4);
    height = u16le(bytes, offset + 6);
  } else {
    if (bytes.length < offset + 16) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "BMP DIB header is truncated");
    }
    width = Math.abs(i32le(bytes, offset + 4));
    height = Math.abs(i32le(bytes, offset + 8));
    if (dibSize === 16) {
      const planes = u16le(bytes, offset + 12);
      const bits = u16le(bytes, offset + 14);
      if (planes !== 1 || ![1, 4, 8, 16, 24, 32].includes(bits)) {
        throw new OfficeEngineError("IMAGE_HEADER_INVALID", "OS/2 DIB plane or bit depth is invalid");
      }
    }
  }
  return makeInfo(fileHeader ? "bmp" : "dib", "image/bmp", width, height);
}

function svgRoot(source: string): string | undefined {
  let offset = source.charCodeAt(0) === 0xfeff ? 1 : 0;
  const skipWhitespace = () => {
    while (offset < source.length && /\s/u.test(source[offset]!)) offset += 1;
  };
  skipWhitespace();
  while (offset < source.length) {
    if (source.startsWith("<?", offset)) {
      const end = source.indexOf("?>", offset + 2);
      if (end < 0) return undefined;
      offset = end + 2;
    } else if (source.startsWith("<!--", offset)) {
      const end = source.indexOf("-->", offset + 4);
      if (end < 0) return undefined;
      offset = end + 3;
    } else if (source.startsWith("<!DOCTYPE", offset)) {
      let quote: string | undefined;
      let subsetDepth = 0;
      let end = -1;
      for (let index = offset + 9; index < source.length; index += 1) {
        const character = source[index]!;
        if (quote !== undefined) {
          if (character === quote) quote = undefined;
        } else if (character === "\"" || character === "'") {
          quote = character;
        } else if (character === "[") {
          subsetDepth += 1;
        } else if (character === "]") {
          subsetDepth = Math.max(0, subsetDepth - 1);
        } else if (character === ">" && subsetDepth === 0) {
          end = index;
          break;
        }
      }
      if (end < 0) return undefined;
      offset = end + 1;
    } else {
      break;
    }
    skipWhitespace();
  }
  const rootStart = source.slice(offset).match(/^<(?:[A-Za-z_][\w.-]*:)?svg(?=[\s/>])/u);
  if (rootStart === null) return undefined;
  let quote: string | undefined;
  for (let index = offset + rootStart[0].length; index < source.length; index += 1) {
    const character = source[index]!;
    if (quote !== undefined) {
      if (character === quote) quote = undefined;
    } else if (character === "\"" || character === "'") {
      quote = character;
    } else if (character === ">") {
      return source.slice(offset, index + 1);
    }
  }
  return undefined;
}

function svgAttribute(root: string, name: string): string | undefined {
  const expression = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')`, "u");
  const match = expression.exec(root);
  return match?.[1] ?? match?.[2];
}

function svgLength(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(px|in|cm|mm|q|pt|pc)?$/iu.exec(value.trim());
  if (match === null) return undefined;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return undefined;
  const scale = {
    px: 1,
    in: 96,
    cm: 96 / 2.54,
    mm: 96 / 25.4,
    q: 96 / 101.6,
    pt: 96 / 72,
    pc: 16,
  }[match[2]?.toLowerCase() ?? "px"]!;
  return amount * scale;
}

function svg(bytes: Uint8Array): IdentifiedImage | undefined {
  let source: string;
  try {
    source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return undefined;
  }
  const root = svgRoot(source);
  if (root === undefined) return undefined;
  const viewBox = (svgAttribute(root, "viewBox") ?? "")
    .trim()
    .split(/[\s,]+/u)
    .map(Number);
  const viewWidth = viewBox.length === 4 && Number.isFinite(viewBox[2]) && viewBox[2]! > 0 ? viewBox[2] : undefined;
  const viewHeight = viewBox.length === 4 && Number.isFinite(viewBox[3]) && viewBox[3]! > 0 ? viewBox[3] : undefined;
  let width = svgLength(svgAttribute(root, "width"));
  let height = svgLength(svgAttribute(root, "height"));
  if (width === undefined && height === undefined && viewWidth !== undefined && viewHeight !== undefined) {
    width = viewWidth;
    height = viewHeight;
  } else if (width === undefined && height !== undefined && viewWidth !== undefined && viewHeight !== undefined) {
    width = height * viewWidth / viewHeight;
  } else if (height === undefined && width !== undefined && viewWidth !== undefined && viewHeight !== undefined) {
    height = width * viewHeight / viewWidth;
  }
  width ??= 300;
  height ??= 150;
  return makeInfo("svg", "image/svg+xml", Math.ceil(width), Math.ceil(height));
}

interface TiffReader {
  readonly littleEndian: boolean;
  u16(offset: number): number;
  u32(offset: number): number;
}

function tiffReader(bytes: Uint8Array): TiffReader | undefined {
  if (bytes.length < 8) return undefined;
  const order = ascii(bytes, 0, 2);
  if (order !== "II" && order !== "MM") return undefined;
  const littleEndian = order === "II";
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number) => {
    if (offset < 0 || offset + 2 > bytes.length) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF header is truncated");
    return view.getUint16(offset, littleEndian);
  };
  const u32 = (offset: number) => {
    if (offset < 0 || offset + 4 > bytes.length) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF header is truncated");
    return view.getUint32(offset, littleEndian);
  };
  return u16(2) === 42 ? { littleEndian, u16, u32 } : undefined;
}

function tiffValue(
  bytes: Uint8Array,
  reader: TiffReader,
  entryOffset: number,
  type: number,
  count: number,
): number | undefined {
  const typeSize = type === 3 ? 2 : type === 4 || type === 9 ? 4 : 0;
  if (typeSize === 0 || count < 1) return undefined;
  const valueField = entryOffset + 8;
  const valueOffset = typeSize * count <= 4 ? valueField : reader.u32(valueField);
  if (valueOffset < 0 || valueOffset + typeSize > bytes.length) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF tag value points outside the image");
  }
  if (type === 3) return reader.u16(valueOffset);
  if (type === 4) return reader.u32(valueOffset);
  if (type === 9) {
    const value = reader.u32(valueOffset);
    return value & 0x8000_0000 ? value - 0x1_0000_0000 : value;
  }
  return undefined;
}

function tiff(bytes: Uint8Array): IdentifiedImage | undefined {
  const reader = tiffReader(bytes);
  if (reader === undefined) return undefined;
  const ifdOffset = reader.u32(4);
  const count = reader.u16(ifdOffset);
  if (count > 65_535) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF has too many directory entries");
  const countBytes = 2;
  const entryBytes = 12;
  if (ifdOffset + countBytes + count * entryBytes > bytes.length) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF image directory is truncated");
  }
  let width: number | undefined;
  let height: number | undefined;
  let orientation = 1;
  for (let index = 0; index < count; index += 1) {
    const entry = ifdOffset + countBytes + index * entryBytes;
    const tag = reader.u16(entry);
    if (tag !== 256 && tag !== 257 && tag !== 274) continue;
    const type = reader.u16(entry + 2);
    const valueCount = reader.u32(entry + 4);
    const value = tiffValue(bytes, reader, entry, type, valueCount);
    if (value === undefined) continue;
    if (tag === 256) width = value;
    else if (tag === 257) height = value;
    else orientation = value;
  }
  if (width === undefined || height === undefined) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "TIFF has no first-frame dimensions");
  }
  if (orientation >= 5 && orientation <= 8) [width, height] = [height, width];
  return makeInfo("tiff", "image/tiff", width, height);
}

function jpegXr(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 8 || bytes[0] !== 0x49 || bytes[1] !== 0x49 || bytes[2] !== 0xbc || bytes[3] !== 0x01) {
    return undefined;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number) => view.getUint16(offset, true);
  const u32 = (offset: number) => view.getUint32(offset, true);
  const ifdOffset = u32(4);
  const count = u16(ifdOffset);
  if (ifdOffset + 2 + count * 12 > bytes.length) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG XR image directory is truncated");
  }
  let width: number | undefined;
  let height: number | undefined;
  for (let index = 0; index < count; index += 1) {
    const entry = ifdOffset + 2 + index * 12;
    const tag = u16(entry);
    if (tag !== 0xbc80 && tag !== 0xbc81 && tag !== 256 && tag !== 257) continue;
    const type = u16(entry + 2);
    const valueCount = u32(entry + 4);
    if (valueCount !== 1 || (type !== 3 && type !== 4)) continue;
    const value = type === 3 ? u16(entry + 8) : u32(entry + 8);
    if (tag === 0xbc80 || tag === 256) width = value;
    else height = value;
  }
  if (width === undefined || height === undefined) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG XR has no image dimensions");
  }
  return makeInfo("jxr", "image/vnd.ms-photo", width, height);
}

function icon(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 6 || u16le(bytes, 0) !== 0 || u16le(bytes, 2) !== 1) return undefined;
  const count = u16le(bytes, 4);
  if (count < 1 || count > 4_096 || 6 + count * 16 > bytes.length) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "ICO directory is invalid or truncated");
  }
  let bestWidth = 0;
  let bestHeight = 0;
  let bestDepth = -1;
  for (let index = 0; index < count; index += 1) {
    const offset = 6 + index * 16;
    const width = bytes[offset] === 0 ? 256 : bytes[offset]!;
    const height = bytes[offset + 1] === 0 ? 256 : bytes[offset + 1]!;
    const depth = u16le(bytes, offset + 6);
    const length = u32le(bytes, offset + 8);
    const dataOffset = u32le(bytes, offset + 12);
    if (length === 0 || dataOffset < 6 + count * 16 || dataOffset + length > bytes.length) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "ICO image entry points outside the file");
    }
    if (width * height > bestWidth * bestHeight || (width * height === bestWidth * bestHeight && depth > bestDepth)) {
      bestWidth = width;
      bestHeight = height;
      bestDepth = depth;
    }
  }
  return makeInfo("ico", "image/x-icon", bestWidth, bestHeight);
}

function pcx(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 128 || bytes[0] !== 0x0a) return undefined;
  if (![0, 2, 3, 4, 5].includes(bytes[1]!) || ![0, 1].includes(bytes[2]!) || ![1, 2, 4, 8].includes(bytes[3]!)) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "PCX header uses an unsupported version or pixel layout");
  }
  const left = u16le(bytes, 4);
  const top = u16le(bytes, 6);
  const right = u16le(bytes, 8);
  const bottom = u16le(bytes, 10);
  return makeInfo("pcx", "image/x-pcx", right - left + 1, bottom - top + 1);
}

export interface Jpeg2000ComponentInfo {
  readonly precision: number;
  readonly signed: boolean;
  readonly xSubsampling: number;
  readonly ySubsampling: number;
}

export interface Jpeg2000Structure {
  readonly boxed: boolean;
  readonly width: number;
  readonly height: number;
  readonly xOrigin: number;
  readonly yOrigin: number;
  readonly components: readonly Jpeg2000ComponentInfo[];
  readonly enumeratedColorSpace?: number;
}

function j2kStructure(
  bytes: Uint8Array,
  start = 0,
  end = bytes.length,
): Omit<Jpeg2000Structure, "boxed" | "enumeratedColorSpace"> | undefined {
  if (end - start < 4 || bytes[start] !== 0xff || bytes[start + 1] !== 0x4f) return undefined;
  const scanEnd = Math.min(end - 41, start + 1_048_576);
  for (let offset = start + 2; offset <= scanEnd; offset += 1) {
    if (bytes[offset] !== 0xff || bytes[offset + 1] !== 0x51) continue;
    const length = u16be(bytes, offset + 2);
    if (length < 41 || offset + 2 + length > end) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 SIZ marker is truncated");
    }
    const xSize = u32be(bytes, offset + 6);
    const ySize = u32be(bytes, offset + 10);
    const xOrigin = u32be(bytes, offset + 14);
    const yOrigin = u32be(bytes, offset + 18);
    const tileWidth = u32be(bytes, offset + 22);
    const tileHeight = u32be(bytes, offset + 26);
    const tileXOrigin = u32be(bytes, offset + 30);
    const tileYOrigin = u32be(bytes, offset + 34);
    if (xSize <= xOrigin || ySize <= yOrigin) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 canvas dimensions are invalid");
    }
    if (tileWidth === 0 || tileHeight === 0 || tileXOrigin > xOrigin || tileYOrigin > yOrigin
      || BigInt(tileXOrigin) + BigInt(tileWidth) <= BigInt(xOrigin)
      || BigInt(tileYOrigin) + BigInt(tileHeight) <= BigInt(yOrigin)) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 tile grid is invalid");
    }
    const componentCount = u16be(bytes, offset + 38);
    const expectedLength = 38 + componentCount * 3;
    if (componentCount < 1 || componentCount > 16_384 || length !== expectedLength) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 component table is invalid");
    }
    const components: Jpeg2000ComponentInfo[] = [];
    for (let index = 0; index < componentCount; index += 1) {
      const componentOffset = offset + 40 + index * 3;
      const precision = (bytes[componentOffset]! & 0x7f) + 1;
      const xSubsampling = bytes[componentOffset + 1]!;
      const ySubsampling = bytes[componentOffset + 2]!;
      if (precision > 38 || xSubsampling === 0 || ySubsampling === 0) {
        throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 component sampling is invalid");
      }
      components.push({
        precision,
        signed: (bytes[componentOffset]! & 0x80) !== 0,
        xSubsampling,
        ySubsampling,
      });
    }
    return { width: xSize - xOrigin, height: ySize - yOrigin, xOrigin, yOrigin, components };
  }
  throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 codestream has no bounded SIZ marker");
}

interface Jp2BoxState {
  header?: { readonly width: number; readonly height: number; readonly components: number };
  codestream?: Omit<Jpeg2000Structure, "boxed" | "enumeratedColorSpace">;
  enumeratedColorSpace?: number;
}

function inspectJp2Boxes(
  bytes: Uint8Array,
  start: number,
  end: number,
  depth: number,
  state: Jp2BoxState,
): void {
  if (depth > 4) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 box nesting is excessive");
  let offset = start;
  while (offset < end) {
    if (offset + 8 > end) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 box header is truncated");
    let length = u32be(bytes, offset);
    const type = ascii(bytes, offset + 4, 4);
    let header = 8;
    if (length === 1) {
      if (offset + 16 > end) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 extended box is truncated");
      length = boundedInteger(new DataView(bytes.buffer, bytes.byteOffset + offset + 8, 8).getBigUint64(0, false), "JPEG 2000 box length");
      header = 16;
    } else if (length === 0) {
      length = end - offset;
    }
    if (length < header || offset + length > end) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 box length is invalid");
    }
    const dataStart = offset + header;
    const boxEnd = offset + length;
    if (type === "ihdr") {
      if (boxEnd - dataStart < 14 || state.header !== undefined) {
        throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 image header is invalid");
      }
      state.header = {
        width: u32be(bytes, dataStart + 4),
        height: u32be(bytes, dataStart),
        components: u16be(bytes, dataStart + 8),
      };
    } else if (type === "colr" && state.enumeratedColorSpace === undefined) {
      if (boxEnd - dataStart >= 7 && bytes[dataStart] === 1) {
        state.enumeratedColorSpace = u32be(bytes, dataStart + 3);
      }
    } else if (type === "jp2h") {
      inspectJp2Boxes(bytes, dataStart, boxEnd, depth + 1, state);
    } else if (type === "jp2c") {
      if (state.codestream !== undefined) {
        throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 contains multiple codestream boxes");
      }
      const codestream = j2kStructure(bytes, dataStart, boxEnd);
      if (codestream === undefined) {
        throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JPEG 2000 codestream box has no SOC marker");
      }
      state.codestream = codestream;
    }
    offset = boxEnd;
  }
}

export function inspectJpeg2000Structure(bytes: Uint8Array): Jpeg2000Structure | undefined {
  if (bytes.length >= 12 && ascii(bytes, 4, 4) === "jP  " && u32be(bytes, 0) === 12 && u32be(bytes, 8) === 0x0d0a_870a) {
    const state: Jp2BoxState = {};
    inspectJp2Boxes(bytes, 0, bytes.length, 0, state);
    if (state.header === undefined || state.codestream === undefined) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JP2 file is missing its image header or codestream");
    }
    if (state.header.width !== state.codestream.width || state.header.height !== state.codestream.height
      || state.header.components !== state.codestream.components.length) {
      throw new OfficeEngineError("IMAGE_HEADER_INVALID", "JP2 image header disagrees with the codestream SIZ marker");
    }
    return state.enumeratedColorSpace === undefined
      ? { ...state.codestream, boxed: true }
      : { ...state.codestream, boxed: true, enumeratedColorSpace: state.enumeratedColorSpace };
  }
  const structure = j2kStructure(bytes);
  return structure === undefined ? undefined : { ...structure, boxed: false };
}

function jpeg2000(bytes: Uint8Array): IdentifiedImage | undefined {
  const structure = inspectJpeg2000Structure(bytes);
  if (structure === undefined) return undefined;
  return makeInfo(structure.boxed ? "jp2" : "j2k", structure.boxed ? "image/jp2" : "image/j2k", structure.width, structure.height);
}

function emf(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 88 || u32le(bytes, 0) !== 1 || ascii(bytes, 40, 4) !== " EMF") return undefined;
  const headerSize = u32le(bytes, 4);
  if (headerSize < 88 || headerSize % 4 !== 0 || headerSize > bytes.length) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "EMF header length is invalid");
  }
  const frameWidth = i32le(bytes, 32) - i32le(bytes, 24);
  const frameHeight = i32le(bytes, 36) - i32le(bytes, 28);
  if (frameWidth > 0 && frameHeight > 0) {
    const width = Math.max(1, Math.round(frameWidth * 96 / 2_540));
    const height = Math.max(1, Math.round(frameHeight * 96 / 2_540));
    return makeInfo("emf", "image/x-emf", width, height);
  }
  const boundsWidth = i32le(bytes, 16) - i32le(bytes, 8);
  const boundsHeight = i32le(bytes, 20) - i32le(bytes, 12);
  const deviceWidth = i32le(bytes, 72);
  const deviceHeight = i32le(bytes, 76);
  const millimetersWidth = i32le(bytes, 80);
  const millimetersHeight = i32le(bytes, 84);
  if (boundsWidth > 0 && boundsHeight > 0 && deviceWidth > 0 && deviceHeight > 0
    && millimetersWidth > 0 && millimetersHeight > 0) {
    const width = Math.max(1, Math.round(boundsWidth * millimetersWidth * 96 / (deviceWidth * 25.4)));
    const height = Math.max(1, Math.round(boundsHeight * millimetersHeight * 96 / (deviceHeight * 25.4)));
    return makeInfo("emf", "image/x-emf", width, height);
  }
  return makeInfo("emf", "image/x-emf", 1, 1, false);
}

function placeableWmfChecksum(bytes: Uint8Array): boolean {
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= u16le(bytes, offset);
  return checksum === u16le(bytes, 20);
}

function wmfHeader(bytes: Uint8Array, offset: number): boolean {
  return bytes.length >= offset + 18
    && (u16le(bytes, offset) === 1 || u16le(bytes, offset) === 2)
    && u16le(bytes, offset + 2) === 9
    && (u16le(bytes, offset + 4) === 0x0100 || u16le(bytes, offset + 4) === 0x0300);
}

function wmf(bytes: Uint8Array): IdentifiedImage | undefined {
  const placeable = bytes.length >= 40 && u32le(bytes, 0) === 0x9ac6_cdd7;
  const headerOffset = placeable ? 22 : 0;
  if (!wmfHeader(bytes, headerOffset)) return undefined;
  if (placeable) {
    if (!placeableWmfChecksum(bytes)) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "WMF placeable header checksum is invalid");
    const unitsPerInch = u16le(bytes, 14);
    if (unitsPerInch === 0) throw new OfficeEngineError("IMAGE_HEADER_INVALID", "WMF placeable units are invalid");
    const width = Math.ceil(Math.abs(i16le(bytes, 10) - i16le(bytes, 6)) * 96 / unitsPerInch);
    const height = Math.ceil(Math.abs(i16le(bytes, 12) - i16le(bytes, 8)) * 96 / unitsPerInch);
    return width > 0 && height > 0
      ? makeInfo("wmf", "image/x-wmf", width, height)
      : makeInfo("wmf", "image/x-wmf", 1, 1, false);
  }
  return makeInfo("wmf", "image/x-wmf", 1, 1, false);
}

function gzip(bytes: Uint8Array, declaredMediaType: OfficeImageMediaType): IdentifiedImage | undefined {
  if (bytes.length < 10 || bytes[0] !== 0x1f || bytes[1] !== 0x8b || bytes[2] !== 8 || (bytes[3]! & 0xe0) !== 0) return undefined;
  if (declaredMediaType === "image/x-emz") return makeInfo("emz", declaredMediaType, 1, 1, false);
  if (declaredMediaType === "image/x-wmz") return makeInfo("wmz", declaredMediaType, 1, 1, false);
  return undefined;
}

function pdfRaster(bytes: Uint8Array): IdentifiedImage | undefined {
  if (bytes.length < 21 || ascii(bytes, 0, 8) !== "OVPDFR01") return undefined;
  const components = bytes[16]!;
  const predictor = bytes[17]!;
  const paletteComponents = bytes[18]!;
  const paletteLength = paletteComponents * (bytes[19]! + 1);
  if (![1, 3, 4].includes(components)
    || (predictor !== 1 && predictor !== 2 && (predictor < 10 || predictor > 15))
    || (paletteComponents === 0
      ? bytes[19] !== 0
      : components !== 1
        || ![1, 3, 4].includes(paletteComponents)
        || bytes.length <= 28 + paletteLength)) {
    throw new OfficeEngineError("IMAGE_HEADER_INVALID", "Compressed PDF raster header is invalid");
  }
  return makeInfo(
    "pdf-raster",
    "image/x-officeviewer-pdf-raster",
    u32be(bytes, 8),
    u32be(bytes, 12),
  );
}

function canonicalMediaType(value: string): OfficeImageMediaType | undefined {
  switch (value.trim().toLowerCase().split(";", 1)[0]) {
    case "image/png":
    case "image/x-png": return "image/png";
    case "image/jpeg":
    case "image/jpg":
    case "image/pjpeg": return "image/jpeg";
    case "image/gif": return "image/gif";
    case "image/webp": return "image/webp";
    case "image/bmp":
    case "image/x-bmp":
    case "image/x-ms-bmp": return "image/bmp";
    case "image/svg+xml":
    case "image/svg": return "image/svg+xml";
    case "image/tiff":
    case "image/tif": return "image/tiff";
    case "image/vnd.ms-photo":
    case "image/jxr": return "image/vnd.ms-photo";
    case "image/x-icon":
    case "image/vnd.microsoft.icon":
    case "image/ico": return "image/x-icon";
    case "image/x-pcx":
    case "image/pcx":
    case "image/vnd.zbrush.pcx": return "image/x-pcx";
    case "image/jp2":
    case "image/jpx": return "image/jp2";
    case "image/j2k":
    case "image/jpc": return "image/j2k";
    case "image/x-emf":
    case "image/emf": return "image/x-emf";
    case "image/x-wmf":
    case "image/wmf": return "image/x-wmf";
    case "image/x-emz": return "image/x-emz";
    case "image/x-wmz": return "image/x-wmz";
    case "image/x-officeviewer-pdf-raster": return "image/x-officeviewer-pdf-raster";
    default: return undefined;
  }
}

export function reducedJpeg2000Dimensions(
  width: number,
  height: number,
  maxPixels: number,
): { readonly width: number; readonly height: number; readonly reducePower: number } {
  let reducePower = 0;
  let reducedWidth = width;
  let reducedHeight = height;
  while (reducedWidth * reducedHeight > maxPixels) {
    reducePower += 1;
    reducedWidth = Math.ceil(width / 2 ** reducePower);
    reducedHeight = Math.ceil(height / 2 ** reducePower);
  }
  return { width: reducedWidth, height: reducedHeight, reducePower };
}

export function identifyImage(
  bytes: Uint8Array,
  declaredMediaType: string,
  maxPixels: number,
  allowReducedJpeg2000 = false,
): IdentifiedImage {
  const declared = canonicalMediaType(declaredMediaType);
  if (declared === undefined) {
    throw new OfficeEngineError(
      "IMAGE_MEDIA_TYPE_UNSUPPORTED",
      `Declared image media type ${declaredMediaType || "(missing)"} is not supported`,
    );
  }
  const result = png(bytes)
    ?? jpeg(bytes)
    ?? gif(bytes)
    ?? webp(bytes)
    ?? bitmap(bytes)
    ?? svg(bytes)
    ?? jpegXr(bytes)
    ?? tiff(bytes)
    ?? icon(bytes)
    ?? pcx(bytes)
    ?? jpeg2000(bytes)
    ?? emf(bytes)
    ?? wmf(bytes)
    ?? gzip(bytes, declared)
    ?? pdfRaster(bytes);
  if (result === undefined) {
    throw new OfficeEngineError(
      "IMAGE_FORMAT_UNSUPPORTED",
      "Image bytes are not a supported Office or PDF image",
    );
  }
  if (declared !== result.mediaType) {
    throw new OfficeEngineError(
      "IMAGE_MEDIA_TYPE_MISMATCH",
      `Declared media type ${declaredMediaType || "(missing)"} does not match ${result.mediaType}`,
    );
  }
  if (result.dimensionsKnown
    && result.pixels > maxPixels
    && !(allowReducedJpeg2000 && (result.format === "jp2" || result.format === "j2k"))) {
    throw new OfficeEngineError(
      "IMAGE_DIMENSION_LIMIT",
      `Image is ${result.width}×${result.height} (${result.pixels} pixels); limit is ${maxPixels}`,
    );
  }
  return result;
}

export function inspectImage(bytes: Uint8Array, declaredMediaType: string, maxPixels: number): ImageInfo {
  const { mediaType, width, height, pixels } = identifyImage(bytes, declaredMediaType, maxPixels);
  return { mediaType, width, height, pixels };
}
