import { semanticSymbolFontText } from "./symbol-font-mappings.js";
import { initWasm, Resvg } from "@resvg/resvg-wasm";
import { convertEmfToDataUrl, convertWmfToDataUrl } from "emf-converter";
import OpenJPEG from "../third_party/openjpeg/openjpeg.js";
import pako from "pako";
import UTIF from "utif2";
import { decode as decodeJpegXr } from "@discourse/jxr";

import { identifyImage, inspectJpeg2000Structure, reducedJpeg2000Dimensions } from "./image.js";
import type { OfficeImageFormat, OfficeImageMediaType } from "./image.js";
import { approximateFontFamily } from "./font.js";
import { secureMetafileForConversion } from "./metafile-safety.js";
import { OfficeEngineError } from "./types.js";

const MAX_VECTOR_DIMENSION = 8_192;
const MAX_METAFILE_RECORDS = 200_000;
const MAX_TIFF_SEGMENTS = 1_000_000;
const MAX_TIFF_METADATA_VALUES = 2_000_000;
const MAX_TIFF_WORKING_BYTES = 256 * 1024 * 1024;
const TARGETED_VECTOR_FORMATS = new Set<OfficeImageFormat>(["svg", "emf", "wmf", "emz", "wmz"]);

export interface OfficeImageCodecInput {
  readonly format: OfficeImageFormat;
  readonly mediaType: OfficeImageMediaType;
  readonly bytes: ArrayBuffer | Uint8Array;
  readonly maxPixels: number;
  readonly maxBytes: number;
  readonly maxCompressionRatio: number;
  readonly targetWidth?: number;
  readonly targetHeight?: number;
  readonly fonts?: readonly OfficeImageCodecFont[];
}

export interface OfficeImageCodecFont {
  readonly family: string;
  readonly bytes: ArrayBuffer;
}

export interface RgbaCodecOutput {
  readonly kind: "rgba";
  readonly data: Uint8ClampedArray;
  readonly width: number;
  readonly height: number;
  readonly approximate: boolean;
}

export interface EncodedCodecOutput {
  readonly kind: "encoded";
  readonly data: Uint8Array;
  readonly mediaType: "image/png";
  readonly width: number;
  readonly height: number;
  readonly approximate: boolean;
}

export type OfficeImageCodecOutput = RgbaCodecOutput | EncodedCodecOutput;

function failure(code: string, message: string, cause?: unknown): never {
  throw new OfficeEngineError(code, message, cause === undefined ? {} : { cause });
}

function u16le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 2 > bytes.length) failure("IMAGE_HEADER_INVALID", "Image header is truncated");
  return bytes[offset]! | (bytes[offset + 1]! << 8);
}

function i16le(bytes: Uint8Array, offset: number): number {
  const value = u16le(bytes, offset);
  return value & 0x8000 ? value - 0x1_0000 : value;
}

function u32le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > bytes.length) failure("IMAGE_HEADER_INVALID", "Image header is truncated");
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, true);
}

function i32le(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > bytes.length) failure("IMAGE_HEADER_INVALID", "Image header is truncated");
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getInt32(0, true);
}

function assertDimensions(width: number, height: number, maxPixels: number): void {
  const pixels = width * height;
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width <= 0 || height <= 0) {
    failure("IMAGE_HEADER_INVALID", "Decoded image dimensions are invalid");
  }
  if (!Number.isSafeInteger(pixels) || pixels > maxPixels) {
    failure("IMAGE_DIMENSION_LIMIT", `Decoded image is ${width}×${height}; pixel limit is ${maxPixels}`);
  }
}

function rgbaOutput(
  data: Uint8Array | Uint8ClampedArray,
  width: number,
  height: number,
  maxPixels: number,
  approximate = false,
): RgbaCodecOutput {
  assertDimensions(width, height, maxPixels);
  if (data.byteLength !== width * height * 4) {
    failure("IMAGE_DECODE_FAILED", "Image decoder returned an invalid RGBA buffer");
  }
  const rgba = data instanceof Uint8ClampedArray
    ? data
    : new Uint8ClampedArray(data.buffer, data.byteOffset, data.byteLength);
  return { kind: "rgba", data: rgba, width, height, approximate };
}

function paletteColor(
  palette: Uint8Array,
  index: number,
  stride: 3 | 4,
  bgr = false,
): readonly [number, number, number, number] {
  const offset = index * stride;
  if (offset + stride > palette.length) failure("IMAGE_DECODE_FAILED", "Bitmap palette index is out of range");
  if (stride === 3 && !bgr) return [palette[offset]!, palette[offset + 1]!, palette[offset + 2]!, 255];
  return [palette[offset + 2]!, palette[offset + 1]!, palette[offset]!, 255];
}

function setPixel(
  output: Uint8ClampedArray,
  width: number,
  x: number,
  y: number,
  color: readonly [number, number, number, number],
): void {
  const offset = (y * width + x) * 4;
  output[offset] = color[0];
  output[offset + 1] = color[1];
  output[offset + 2] = color[2];
  output[offset + 3] = color[3];
}

function maskChannel(value: number, mask: number, defaultValue: number): number {
  if (mask === 0) return defaultValue;
  let shift = 0;
  let shifted = mask >>> 0;
  while ((shifted & 1) === 0) {
    shifted >>>= 1;
    shift += 1;
  }
  const sample = ((value & mask) >>> shift) >>> 0;
  return Math.round(sample * 255 / shifted);
}

interface DibOptions {
  readonly icon?: boolean;
}

function decodeRleBitmap(
  bytes: Uint8Array,
  offset: number,
  width: number,
  height: number,
  bits: 4 | 8,
  topDown: boolean,
  palette: Uint8Array,
  paletteStride: 3 | 4,
): Uint8ClampedArray {
  if (topDown) failure("IMAGE_DECODE_FAILED", "Top-down RLE bitmaps are invalid");
  const output = new Uint8ClampedArray(width * height * 4);
  const background = paletteColor(palette, 0, paletteStride, true);
  for (let pixel = 0; pixel < width * height; pixel += 1) {
    const outputOffset = pixel * 4;
    output[outputOffset] = background[0];
    output[outputOffset + 1] = background[1];
    output[outputOffset + 2] = background[2];
    output[outputOffset + 3] = background[3];
  }
  let x = 0;
  let y = height - 1;
  let position = offset;
  let ended = false;
  const emit = (index: number) => {
    if (x >= width || y < 0 || y >= height) failure("IMAGE_DECODE_FAILED", "RLE bitmap writes outside its canvas");
    setPixel(output, width, x, y, paletteColor(palette, index, paletteStride, true));
    x += 1;
  };
  while (!ended && position + 2 <= bytes.length) {
    const count = bytes[position++]!;
    const command = bytes[position++]!;
    if (count !== 0) {
      for (let index = 0; index < count; index += 1) {
        emit(bits === 8 ? command : index % 2 === 0 ? command >>> 4 : command & 0x0f);
      }
      continue;
    }
    if (command === 0) {
      x = 0;
      y -= 1;
    } else if (command === 1) {
      ended = true;
    } else if (command === 2) {
      if (position + 2 > bytes.length) failure("IMAGE_DECODE_FAILED", "RLE delta command is truncated");
      x += bytes[position++]!;
      y -= bytes[position++]!;
      if (x > width || y < 0 || y >= height) failure("IMAGE_DECODE_FAILED", "RLE delta leaves the bitmap canvas");
    } else if (bits === 8) {
      const encodedLength = command + (command & 1);
      if (position + encodedLength > bytes.length) failure("IMAGE_DECODE_FAILED", "RLE8 literal run is truncated");
      for (let index = 0; index < command; index += 1) emit(bytes[position + index]!);
      position += encodedLength;
    } else {
      const packedBytes = Math.ceil(command / 2);
      const encodedLength = packedBytes + (packedBytes & 1);
      if (position + encodedLength > bytes.length) failure("IMAGE_DECODE_FAILED", "RLE4 literal run is truncated");
      for (let index = 0; index < command; index += 1) {
        const packed = bytes[position + Math.floor(index / 2)]!;
        emit(index % 2 === 0 ? packed >>> 4 : packed & 0x0f);
      }
      position += encodedLength;
    }
  }
  if (!ended) failure("IMAGE_DECODE_FAILED", "RLE bitmap has no end marker");
  return output;
}

function decodeDib(
  bytes: Uint8Array,
  maxPixels: number,
  options: DibOptions = {},
): RgbaCodecOutput {
  const hasFileHeader = bytes.length >= 14 && bytes[0] === 0x42 && bytes[1] === 0x4d;
  const headerOffset = hasFileHeader ? 14 : 0;
  const headerSize = u32le(bytes, headerOffset);
  if (![12, 16, 40, 52, 56, 64, 108, 124].includes(headerSize)) {
    failure("IMAGE_DECODE_FAILED", `Unsupported DIB header size ${headerSize}`);
  }
  if (headerOffset + headerSize > bytes.length) failure("IMAGE_HEADER_INVALID", "DIB header is truncated");

  const core = headerSize === 12;
  const storedWidth = core ? u16le(bytes, headerOffset + 4) : i32le(bytes, headerOffset + 4);
  const storedHeight = core ? u16le(bytes, headerOffset + 6) : i32le(bytes, headerOffset + 8);
  if (storedWidth <= 0 || storedHeight === 0) failure("IMAGE_HEADER_INVALID", "DIB dimensions are invalid");
  const width = storedWidth;
  let height = Math.abs(storedHeight);
  if (options.icon) {
    if (height % 2 !== 0) failure("IMAGE_HEADER_INVALID", "ICO DIB height does not include a valid mask");
    height /= 2;
  }
  assertDimensions(width, height, maxPixels);
  const topDown = storedHeight < 0;
  const planes = u16le(bytes, headerOffset + (core ? 8 : 12));
  const bits = u16le(bytes, headerOffset + (core ? 10 : 14));
  if (planes !== 1 || ![1, 4, 8, 16, 24, 32].includes(bits)) {
    failure("IMAGE_DECODE_FAILED", `Unsupported DIB plane/bit depth ${planes}/${bits}`);
  }
  const compression = core || headerSize === 16 ? 0 : u32le(bytes, headerOffset + 16);
  if (![0, 1, 2, 3, 6].includes(compression)
      || (compression === 1 && bits !== 8)
      || (compression === 2 && bits !== 4)
      || ((compression === 3 || compression === 6) && bits !== 16 && bits !== 32)) {
    failure("IMAGE_DECODE_FAILED", `Unsupported DIB compression ${compression} for ${bits}-bit pixels`);
  }

  let tableOffset = headerOffset + headerSize;
  let redMask = 0;
  let greenMask = 0;
  let blueMask = 0;
  let alphaMask = 0;
  if (compression === 3 || compression === 6) {
    if (headerSize >= 52) {
      redMask = u32le(bytes, headerOffset + 40);
      greenMask = u32le(bytes, headerOffset + 44);
      blueMask = u32le(bytes, headerOffset + 48);
      if (headerSize >= 56) alphaMask = u32le(bytes, headerOffset + 52);
    } else {
      redMask = u32le(bytes, tableOffset);
      greenMask = u32le(bytes, tableOffset + 4);
      blueMask = u32le(bytes, tableOffset + 8);
      tableOffset += 12;
      if (compression === 6) {
        alphaMask = u32le(bytes, tableOffset);
        tableOffset += 4;
      }
    }
    if (redMask === 0 || greenMask === 0 || blueMask === 0 || (redMask & greenMask) !== 0
      || (redMask & blueMask) !== 0 || (greenMask & blueMask) !== 0) {
      failure("IMAGE_DECODE_FAILED", "DIB color masks are invalid");
    }
  } else if (bits === 16) {
    redMask = 0x7c00;
    greenMask = 0x03e0;
    blueMask = 0x001f;
  }

  const paletteEntries = bits <= 8
    ? (headerSize < 40 ? 1 << bits : u32le(bytes, headerOffset + 32) || 1 << bits)
    : 0;
  if (paletteEntries > 256) failure("IMAGE_DECODE_FAILED", "DIB palette is too large");
  const paletteStride: 3 | 4 = core ? 3 : 4;
  const paletteBytes = paletteEntries * paletteStride;
  if (tableOffset + paletteBytes > bytes.length) failure("IMAGE_HEADER_INVALID", "DIB palette is truncated");
  const palette = bytes.subarray(tableOffset, tableOffset + paletteBytes);
  const computedPixelOffset = tableOffset + paletteBytes;
  const pixelOffset = hasFileHeader ? u32le(bytes, 10) : computedPixelOffset;
  if (pixelOffset < computedPixelOffset || pixelOffset >= bytes.length) failure("IMAGE_HEADER_INVALID", "DIB pixel offset is invalid");

  if (compression === 1 || compression === 2) {
    const output = decodeRleBitmap(
      bytes,
      pixelOffset,
      width,
      height,
      bits as 4 | 8,
      topDown,
      palette,
      paletteStride,
    );
    return rgbaOutput(output, width, height, maxPixels);
  }

  const rowStride = Math.floor((bits * width + 31) / 32) * 4;
  const pixelBytes = rowStride * height;
  if (!Number.isSafeInteger(pixelBytes) || pixelOffset + pixelBytes > bytes.length) {
    failure("IMAGE_DECODE_FAILED", "DIB pixel rows are truncated");
  }
  const output = new Uint8ClampedArray(width * height * 4);
  let hasNonZeroAlpha = false;
  for (let y = 0; y < height; y += 1) {
    const sourceY = topDown ? y : height - 1 - y;
    const row = pixelOffset + sourceY * rowStride;
    for (let x = 0; x < width; x += 1) {
      let color: readonly [number, number, number, number];
      if (bits <= 8) {
        const bitOffset = x * bits;
        const packed = bytes[row + Math.floor(bitOffset / 8)]!;
        const shift = 8 - bits - (bitOffset % 8);
        const index = (packed >>> shift) & ((1 << bits) - 1);
        color = paletteColor(palette, index, paletteStride, true);
      } else if (bits === 16) {
        const value = u16le(bytes, row + x * 2);
        color = [maskChannel(value, redMask, 0), maskChannel(value, greenMask, 0), maskChannel(value, blueMask, 0), 255];
      } else if (bits === 24) {
        const offset = row + x * 3;
        color = [bytes[offset + 2]!, bytes[offset + 1]!, bytes[offset]!, 255];
      } else {
        const value = u32le(bytes, row + x * 4);
        const alpha = alphaMask !== 0
          ? maskChannel(value, alphaMask, 255)
          : options.icon ? bytes[row + x * 4 + 3]! : 255;
        hasNonZeroAlpha ||= alpha !== 0;
        color = compression === 3 || compression === 6
          ? [maskChannel(value, redMask, 0), maskChannel(value, greenMask, 0), maskChannel(value, blueMask, 0), alpha]
          : [bytes[row + x * 4 + 2]!, bytes[row + x * 4 + 1]!, bytes[row + x * 4]!, alpha];
      }
      setPixel(output, width, x, y, color);
    }
  }

  if (options.icon) {
    const maskOffset = pixelOffset + pixelBytes;
    const maskStride = Math.floor((width + 31) / 32) * 4;
    if (maskOffset + maskStride * height > bytes.length) failure("IMAGE_DECODE_FAILED", "ICO transparency mask is truncated");
    for (let y = 0; y < height; y += 1) {
      const sourceY = height - 1 - y;
      const row = maskOffset + sourceY * maskStride;
      for (let x = 0; x < width; x += 1) {
        const transparent = ((bytes[row + Math.floor(x / 8)]! >>> (7 - x % 8)) & 1) !== 0;
        const alphaOffset = (y * width + x) * 4 + 3;
        if (transparent) output[alphaOffset] = 0;
        else if (bits < 32 || !hasNonZeroAlpha) output[alphaOffset] = 255;
      }
    }
  }
  return rgbaOutput(output, width, height, maxPixels);
}

function decodeIcon(bytes: Uint8Array, maxPixels: number): OfficeImageCodecOutput {
  const count = u16le(bytes, 4);
  let bestOffset = -1;
  let bestLength = 0;
  let bestArea = -1;
  let bestDepth = -1;
  let bestWidth = 0;
  let bestHeight = 0;
  for (let index = 0; index < count; index += 1) {
    const entry = 6 + index * 16;
    const width = bytes[entry] === 0 ? 256 : bytes[entry]!;
    const height = bytes[entry + 1] === 0 ? 256 : bytes[entry + 1]!;
    const depth = u16le(bytes, entry + 6);
    const length = u32le(bytes, entry + 8);
    const offset = u32le(bytes, entry + 12);
    if (offset + length > bytes.length) failure("IMAGE_HEADER_INVALID", "ICO entry is truncated");
    if (width * height > bestArea || (width * height === bestArea && depth > bestDepth)) {
      bestOffset = offset;
      bestLength = length;
      bestArea = width * height;
      bestDepth = depth;
      bestWidth = width;
      bestHeight = height;
    }
  }
  if (bestOffset < 0) failure("IMAGE_DECODE_FAILED", "ICO has no image entries");
  const image = bytes.subarray(bestOffset, bestOffset + bestLength);
  if (image.length >= 24 && image[0] === 0x89 && image[1] === 0x50 && image[2] === 0x4e && image[3] === 0x47) {
    const info = identifyImage(image, "image/png", maxPixels);
    if (info.width !== bestWidth || info.height !== bestHeight) {
      failure("IMAGE_HEADER_INVALID", "ICO directory dimensions disagree with the embedded PNG");
    }
    return { kind: "encoded", data: image.slice(), mediaType: "image/png", width: info.width, height: info.height, approximate: false };
  }
  const dibHeaderSize = u32le(image, 0);
  if (![12, 16, 40, 52, 56, 64, 108, 124].includes(dibHeaderSize) || image.length < dibHeaderSize) {
    failure("IMAGE_HEADER_INVALID", "ICO embedded DIB header is invalid");
  }
  const dibWidth = dibHeaderSize === 12 ? u16le(image, 4) : Math.abs(i32le(image, 4));
  const storedHeight = dibHeaderSize === 12 ? u16le(image, 6) : Math.abs(i32le(image, 8));
  if (storedHeight === 0 || storedHeight % 2 !== 0) failure("IMAGE_HEADER_INVALID", "ICO embedded DIB mask height is invalid");
  const dibHeight = storedHeight / 2;
  assertDimensions(dibWidth, dibHeight, maxPixels);
  if (dibWidth !== bestWidth || dibHeight !== bestHeight) {
    failure("IMAGE_HEADER_INVALID", "ICO directory dimensions disagree with the embedded DIB");
  }
  return decodeDib(image, maxPixels, { icon: true });
}

function pcxPalette(bytes: Uint8Array, totalBits: number): readonly [Uint8Array, number] {
  if (totalBits <= 4) return [bytes.subarray(16, 64), bytes.length];
  if (bytes.length >= 897 && bytes[bytes.length - 769] === 0x0c) {
    return [bytes.subarray(bytes.length - 768), bytes.length - 769];
  }
  if (totalBits === 8) {
    const gray = new Uint8Array(256 * 3);
    for (let index = 0; index < 256; index += 1) gray.fill(index, index * 3, index * 3 + 3);
    return [gray, bytes.length];
  }
  return [new Uint8Array(), bytes.length];
}

function decodePcx(bytes: Uint8Array, maxPixels: number): RgbaCodecOutput {
  const left = u16le(bytes, 4);
  const top = u16le(bytes, 6);
  const width = u16le(bytes, 8) - left + 1;
  const height = u16le(bytes, 10) - top + 1;
  assertDimensions(width, height, maxPixels);
  const encoding = bytes[2]!;
  const bits = bytes[3]!;
  const planes = bytes[65]!;
  const bytesPerLine = u16le(bytes, 66);
  if (planes < 1 || planes > 4 || bits * planes > 32 || bytesPerLine < Math.ceil(width * bits / 8)) {
    failure("IMAGE_DECODE_FAILED", "PCX scanline layout is invalid");
  }
  const totalBits = bits * planes;
  const [palette, dataEnd] = pcxPalette(bytes, totalBits);
  const output = new Uint8ClampedArray(width * height * 4);
  const scanline = new Uint8Array(bytesPerLine * planes);
  let source = 128;
  let approximate = false;
  for (let y = 0; y < height; y += 1) {
    let target = 0;
    while (target < scanline.length) {
      if (source >= dataEnd) failure("IMAGE_DECODE_FAILED", "PCX scanline data is truncated");
      let value = bytes[source++]!;
      let count = 1;
      if (encoding === 1 && (value & 0xc0) === 0xc0) {
        count = value & 0x3f;
        if (count === 0 || source >= dataEnd) failure("IMAGE_DECODE_FAILED", "PCX RLE command is invalid");
        value = bytes[source++]!;
      }
      if (target + count > scanline.length) failure("IMAGE_DECODE_FAILED", "PCX RLE run crosses a scanline");
      scanline.fill(value, target, target + count);
      target += count;
    }
    for (let x = 0; x < width; x += 1) {
      let color: readonly [number, number, number, number];
      if (bits === 8 && planes >= 3) {
        if (planes === 4) {
          const cyan = scanline[x]! / 255;
          const magenta = scanline[bytesPerLine + x]! / 255;
          const yellow = scanline[bytesPerLine * 2 + x]! / 255;
          const black = scanline[bytesPerLine * 3 + x]! / 255;
          color = [
            Math.round(255 * (1 - cyan) * (1 - black)),
            Math.round(255 * (1 - magenta) * (1 - black)),
            Math.round(255 * (1 - yellow) * (1 - black)),
            255,
          ];
          approximate = true;
        } else {
          color = [scanline[x]!, scanline[bytesPerLine + x]!, scanline[bytesPerLine * 2 + x]!, 255];
        }
      } else {
        let paletteIndex = 0;
        for (let plane = 0; plane < planes; plane += 1) {
          const bitOffset = x * bits;
          const packed = scanline[plane * bytesPerLine + Math.floor(bitOffset / 8)]!;
          const shift = 8 - bits - (bitOffset % 8);
          paletteIndex |= ((packed >>> shift) & ((1 << bits) - 1)) << (plane * bits);
        }
        color = paletteColor(palette, paletteIndex, 3);
      }
      setPixel(output, width, x, y, color);
    }
  }
  return rgbaOutput(output, width, height, maxPixels, approximate);
}

function orientationValue(ifd: UTIF.IFD): number {
  const value = ifd.t274;
  return Array.isArray(value) && typeof value[0] === "number" ? value[0] : 1;
}

function orientRgba(
  source: Uint8Array,
  width: number,
  height: number,
  orientation: number,
): readonly [Uint8ClampedArray, number, number] {
  if (orientation < 2 || orientation > 8) {
    return [new Uint8ClampedArray(source.buffer, source.byteOffset, source.byteLength), width, height];
  }
  const swap = orientation >= 5;
  const outputWidth = swap ? height : width;
  const outputHeight = swap ? width : height;
  const output = new Uint8ClampedArray(source.byteLength);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let dx = x;
      let dy = y;
      switch (orientation) {
        case 2: dx = width - 1 - x; break;
        case 3: dx = width - 1 - x; dy = height - 1 - y; break;
        case 4: dy = height - 1 - y; break;
        case 5: dx = y; dy = x; break;
        case 6: dx = height - 1 - y; dy = x; break;
        case 7: dx = height - 1 - y; dy = width - 1 - x; break;
        case 8: dx = y; dy = width - 1 - x; break;
      }
      const sourceOffset = (y * width + x) * 4;
      output.set(source.subarray(sourceOffset, sourceOffset + 4), (dy * outputWidth + dx) * 4);
    }
  }
  return [output, outputWidth, outputHeight];
}

interface TiffEntry {
  readonly type: number;
  readonly count: number;
  readonly valueOffset: number;
}

interface TiffPreflight {
  readonly buffer: ArrayBuffer;
  readonly width: number;
  readonly height: number;
}

function preflightTiff(
  bytes: Uint8Array,
  maxPixels: number,
  maxCompressionRatio: number,
): TiffPreflight {
  if (bytes.length < 8) failure("IMAGE_HEADER_INVALID", "TIFF header is truncated");
  const order = String.fromCharCode(bytes[0]!, bytes[1]!);
  if (order !== "II" && order !== "MM") failure("IMAGE_HEADER_INVALID", "TIFF byte order is invalid");
  const littleEndian = order === "II";
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const readU16 = (offset: number): number => {
    if (offset < 0 || offset + 2 > bytes.length) failure("IMAGE_HEADER_INVALID", "TIFF field is truncated");
    return view.getUint16(offset, littleEndian);
  };
  const readU32 = (offset: number): number => {
    if (offset < 0 || offset + 4 > bytes.length) failure("IMAGE_HEADER_INVALID", "TIFF field is truncated");
    return view.getUint32(offset, littleEndian);
  };
  if (readU16(2) !== 42) failure("IMAGE_HEADER_INVALID", "TIFF version is unsupported");
  const ifdOffset = readU32(4);
  const entryCount = readU16(ifdOffset);
  const directoryEnd = ifdOffset + 2 + entryCount * 12;
  if (!Number.isSafeInteger(directoryEnd) || directoryEnd + 4 > bytes.length) {
    failure("IMAGE_HEADER_INVALID", "TIFF image directory is truncated");
  }
  const typeSizes = [0, 1, 1, 2, 4, 8, 1, 1, 2, 4, 8, 4, 8, 4] as const;
  const entries = new Map<number, TiffEntry>();
  const relevantTags = new Set([256, 257, 258, 259, 262, 273, 277, 278, 279, 284, 322, 323, 324, 325]);
  const recursiveIfdTags = new Set([330, 34_665, 34_853, 50_740, 61_440]);
  const recursiveIfdEntries: number[] = [];
  let metadataValues = 0;
  for (let index = 0; index < entryCount; index += 1) {
    const entryOffset = ifdOffset + 2 + index * 12;
    const tag = readU16(entryOffset);
    const type = readU16(entryOffset + 2);
    const count = readU32(entryOffset + 4);
    const typeSize = typeSizes[type] ?? 0;
    if (typeSize === 0) failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} has an invalid type`);
    if (count === 0) {
      if (relevantTags.has(tag)) failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} has no values`);
      continue;
    }
    metadataValues += count;
    if (!Number.isSafeInteger(metadataValues) || metadataValues > MAX_TIFF_METADATA_VALUES) {
      failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF metadata exceeds the codec working-set limit");
    }
    const valueBytes = typeSize * count;
    if (!Number.isSafeInteger(valueBytes)) failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} is too large`);
    const valueOffset = valueBytes <= 4 ? entryOffset + 8 : readU32(entryOffset + 8);
    if (valueOffset < 0 || valueOffset + valueBytes > bytes.length) {
      failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} points outside the image`);
    }
    if (relevantTags.has(tag)) {
      if (entries.has(tag)) failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} is duplicated`);
      entries.set(tag, { type, count, valueOffset });
    }
    if (recursiveIfdTags.has(tag)) recursiveIfdEntries.push(entryOffset);
  }
  const values = (tag: number, maximum: number): number[] | undefined => {
    const entry = entries.get(tag);
    if (entry === undefined) return undefined;
    if (![1, 3, 4].includes(entry.type) || entry.count > maximum) {
      failure("IMAGE_HEADER_INVALID", `TIFF tag ${tag} has an unsupported value layout`);
    }
    const output: number[] = [];
    for (let index = 0; index < entry.count; index += 1) {
      const offset = entry.valueOffset + index * (entry.type === 1 ? 1 : entry.type === 3 ? 2 : 4);
      output.push(entry.type === 1 ? bytes[offset]! : entry.type === 3 ? readU16(offset) : readU32(offset));
    }
    return output;
  };
  const scalar = (tag: number, defaultValue?: number): number => {
    const result = values(tag, 1)?.[0] ?? defaultValue;
    if (result === undefined) failure("IMAGE_HEADER_INVALID", `TIFF required tag ${tag} is missing`);
    return result;
  };
  const width = scalar(256);
  const height = scalar(257);
  assertDimensions(width, height, maxPixels);
  const samplesPerPixel = scalar(277, 1);
  if (samplesPerPixel < 1 || samplesPerPixel > 4) {
    failure("IMAGE_DECODE_FAILED", `Unsupported TIFF sample count ${samplesPerPixel}`);
  }
  const bits = values(258, 4) ?? [1];
  if ((bits.length !== 1 && bits.length !== samplesPerPixel)
    || bits.some((value) => value < 1 || value > 16 || value !== bits[0])) {
    failure("IMAGE_DECODE_FAILED", "Unsupported TIFF BitsPerSample layout");
  }
  const planarConfiguration = scalar(284, 1);
  if (planarConfiguration !== 1) failure("IMAGE_DECODE_FAILED", "Planar TIFF samples are unsupported");
  const compression = scalar(259, 1);
  if (![1, 2, 3, 4, 5, 6, 7, 8, 32_773, 32_946].includes(compression)) {
    failure("IMAGE_DECODE_FAILED", `Unsupported TIFF compression ${compression}`);
  }
  const bitsPerPixel = bits[0]! * samplesPerPixel;
  const rowBytes = Math.ceil(width * bitsPerPixel / 8);
  const rawBytes = rowBytes * height;
  const rgbaBytes = width * height * 4;
  if (!Number.isSafeInteger(rawBytes) || rawBytes + rgbaBytes * 2 > MAX_TIFF_WORKING_BYTES) {
    failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF decoded working set exceeds the codec memory limit");
  }

  const tiled = entries.has(322) || entries.has(323) || entries.has(324) || entries.has(325);
  let offsets: number[];
  let byteCounts: number[];
  let decodedPayloadBytes = rawBytes;
  if (tiled) {
    const tileWidth = scalar(322);
    const tileHeight = scalar(323);
    if (tileWidth < 1 || tileHeight < 1) failure("IMAGE_HEADER_INVALID", "TIFF tile dimensions are invalid");
    const tileColumns = Math.ceil(width / tileWidth);
    const tileRows = Math.ceil(height / tileHeight);
    const tileCount = tileColumns * tileRows;
    if (!Number.isSafeInteger(tileCount) || tileCount > MAX_TIFF_SEGMENTS) {
      failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF has too many tiles");
    }
    offsets = values(324, tileCount) ?? failure("IMAGE_HEADER_INVALID", "TIFF tile offsets are missing");
    byteCounts = values(325, tileCount) ?? failure("IMAGE_HEADER_INVALID", "TIFF tile byte counts are missing");
    if (offsets.length !== tileCount || byteCounts.length !== tileCount) {
      failure("IMAGE_HEADER_INVALID", "TIFF tile arrays do not match the tile grid");
    }
    const tileBytes = Math.ceil(tileWidth * tileHeight * bitsPerPixel / 8);
    decodedPayloadBytes = tileBytes * tileCount;
    if (!Number.isSafeInteger(decodedPayloadBytes)
      || rawBytes + rgbaBytes * 2 + tileBytes > MAX_TIFF_WORKING_BYTES) {
      failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF tile working set exceeds the codec memory limit");
    }
  } else {
    const rowsPerStrip = Math.min(height, scalar(278, height));
    if (rowsPerStrip < 1) failure("IMAGE_HEADER_INVALID", "TIFF RowsPerStrip is invalid");
    const stripCount = Math.ceil(height / rowsPerStrip);
    if (!Number.isSafeInteger(stripCount) || stripCount > MAX_TIFF_SEGMENTS) {
      failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF has too many strips");
    }
    offsets = values(273, stripCount) ?? failure("IMAGE_HEADER_INVALID", "TIFF strip offsets are missing");
    byteCounts = values(279, stripCount) ?? failure("IMAGE_HEADER_INVALID", "TIFF strip byte counts are missing");
    if (offsets.length !== stripCount || byteCounts.length !== stripCount) {
      failure("IMAGE_HEADER_INVALID", "TIFF strip arrays do not match RowsPerStrip");
    }
  }
  const ranges = offsets.map((offset, index) => {
    const byteCount = byteCounts[index]!;
    if (byteCount < 1 || offset + byteCount > bytes.length) {
      failure("IMAGE_HEADER_INVALID", "TIFF strip or tile points outside the image");
    }
    return { start: offset, end: offset + byteCount, byteCount };
  }).sort((left, right) => left.start - right.start);
  let compressedBytes = 0;
  let previousEnd = 0;
  for (const range of ranges) {
    if (range.start < previousEnd) failure("IMAGE_HEADER_INVALID", "TIFF strip or tile ranges overlap");
    previousEnd = range.end;
    compressedBytes += range.byteCount;
  }
  if (!Number.isSafeInteger(compressedBytes)) failure("IMAGE_HEADER_INVALID", "TIFF compressed byte count overflows");
  if (compression === 1) {
    if (compressedBytes < rawBytes) failure("IMAGE_HEADER_INVALID", "Uncompressed TIFF pixel bytes are truncated");
  } else if (BigInt(decodedPayloadBytes) > BigInt(compressedBytes) * BigInt(maxCompressionRatio)) {
    failure("IMAGE_DECOMPRESSION_LIMIT", "TIFF compression ratio exceeds the configured limit");
  }

  const bounded = bytes.slice();
  const boundedView = new DataView(bounded.buffer);
  boundedView.setUint32(directoryEnd, 0, littleEndian);
  for (const entryOffset of recursiveIfdEntries) boundedView.setUint32(entryOffset + 4, 0, littleEndian);
  return { buffer: bounded.buffer, width, height };
}

function decodeTiff(
  bytes: Uint8Array,
  maxPixels: number,
  maxCompressionRatio: number,
): RgbaCodecOutput {
  const preflight = preflightTiff(bytes, maxPixels, maxCompressionRatio);
  let directories: UTIF.IFD[];
  try {
    const decodeDirectories = UTIF.decode as unknown as (
      buffer: ArrayBuffer,
      options: { readonly parseMN: boolean; readonly debug: boolean },
    ) => UTIF.IFD[];
    directories = decodeDirectories(preflight.buffer, { parseMN: false, debug: false });
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", "TIFF directory decode failed", cause);
  }
  const directory = directories[0];
  if (directory === undefined) failure("IMAGE_DECODE_FAILED", "TIFF contains no image directory");
  try {
    UTIF.decodeImage(preflight.buffer, directory);
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", "TIFF pixel decode failed", cause);
  }
  if (directory.width !== preflight.width || directory.height !== preflight.height) {
    failure("IMAGE_DECODE_FAILED", "TIFF decoder dimensions disagree with the preflight header");
  }
  assertDimensions(directory.width, directory.height, maxPixels);
  let pixels: Uint8Array;
  try {
    pixels = UTIF.toRGBA8(directory);
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", "TIFF color conversion failed", cause);
  }
  const [oriented, width, height] = orientRgba(pixels, directory.width, directory.height, orientationValue(directory));
  return rgbaOutput(oriented, width, height, maxPixels);
}

let openJpegModule: ReturnType<typeof OpenJPEG> | undefined;

async function loadOpenJpeg(): ReturnType<typeof OpenJPEG> {
  const url = new URL("./openjpeg.wasm", import.meta.url);
  let wasmBinary: Uint8Array;
  if (url.protocol === "file:") {
    const nodeModule = "node:fs/promises";
    const { readFile } = await import(nodeModule) as {
      readFile(path: URL): Promise<Uint8Array>;
    };
    wasmBinary = await readFile(url);
  } else {
    const response = await fetch(url);
    if (!response.ok) failure("IMAGE_DECODE_FAILED", `OpenJPEG WASM request failed with ${response.status}`);
    wasmBinary = new Uint8Array(await response.arrayBuffer());
  }
  return OpenJPEG({
    wasmBinary,
    warn: () => undefined,
  });
}

async function decodeJpeg2000(bytes: Uint8Array, maxPixels: number): Promise<RgbaCodecOutput> {
  const reductionOffset = bytes.length - 9;
  const minimumReducePower = reductionOffset >= 0
    && String.fromCharCode(...bytes.subarray(reductionOffset, reductionOffset + 8)) === "OVPJPXRD"
    ? bytes[reductionOffset + 8]!
    : 0;
  if (minimumReducePower > 0) {
    const boxedMetadata = reductionOffset >= 8
      && new DataView(bytes.buffer, bytes.byteOffset + reductionOffset - 8, 4).getUint32(0, false) === 17
      && String.fromCharCode(...bytes.subarray(reductionOffset - 4, reductionOffset)) === "free";
    bytes = bytes.subarray(0, reductionOffset - (boxedMetadata ? 8 : 0));
  }
  const opaqueOffset = bytes.length - 9;
  const forceOpaque = opaqueOffset >= 0
    && String.fromCharCode(...bytes.subarray(opaqueOffset, opaqueOffset + 8)) === "OVPJPXOP";
  const pdfColorComponents = forceOpaque ? bytes[opaqueOffset + 8]! : 0;
  if (forceOpaque) {
    const boxedMetadata = opaqueOffset >= 8
      && new DataView(bytes.buffer, bytes.byteOffset + opaqueOffset - 8, 4).getUint32(0, false) === 17
      && String.fromCharCode(...bytes.subarray(opaqueOffset - 4, opaqueOffset)) === "free";
    bytes = bytes.subarray(0, opaqueOffset - (boxedMetadata ? 8 : 0));
  }
  const matteOffset = bytes.length - 11;
  const matte = matteOffset >= 0
    && String.fromCharCode(...bytes.subarray(matteOffset, matteOffset + 8)) === "OVPMATTE"
    ? bytes.subarray(matteOffset + 8)
    : undefined;
  if (matte !== undefined) {
    const boxedMetadata = matteOffset >= 8
      && new DataView(bytes.buffer, bytes.byteOffset + matteOffset - 8, 4).getUint32(0, false) === 19
      && String.fromCharCode(...bytes.subarray(matteOffset - 4, matteOffset)) === "free";
    bytes = bytes.subarray(0, matteOffset - (boxedMetadata ? 8 : 0));
  }
  const structure = inspectJpeg2000Structure(bytes);
  if (structure === undefined) failure("IMAGE_HEADER_INVALID", "JPEG 2000 structure is invalid");
  const primaryComponent = structure.components[0]!;
  if (primaryComponent.xSubsampling !== 1 || primaryComponent.ySubsampling !== 1) {
    failure("IMAGE_DECODE_FAILED", "JPEG 2000 primary-component subsampling is unsupported");
  }
  const bounded = reducedJpeg2000Dimensions(structure.width, structure.height, maxPixels);
  const reducePower = Math.max(minimumReducePower, bounded.reducePower);
  const width = Math.ceil(structure.width / 2 ** reducePower);
  const height = Math.ceil(structure.height / 2 ** reducePower);
  assertDimensions(width, height, maxPixels);
  if (![1, 2, 3, 4].includes(structure.components.length)) {
    failure("IMAGE_DECODE_FAILED", `Unsupported JPEG 2000 component count ${structure.components.length}`);
  }
  openJpegModule ??= loadOpenJpeg();
  const module = await openJpegModule;
  let pointer = 0;
  let decoded: Uint8ClampedArray;
  try {
    pointer = module._malloc(bytes.length);
    module.writeArrayToMemory(bytes, pointer);
    const status = module._jp2_decode(
      pointer,
      bytes.length,
      pdfColorComponents > 0 ? pdfColorComponents : structure.components.length,
      false,
      matte !== undefined,
      reducePower,
    );
    if (status !== 0 || module.imageData === null) {
      const message = module.errorMessages ?? "OpenJPEG rejected the image";
      delete module.errorMessages;
      failure("IMAGE_DECODE_FAILED", message);
    }
    decoded = module.imageData;
    module.imageData = null;
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", "JPEG 2000 decode failed", cause);
  } finally {
    if (pointer !== 0) module._free(pointer);
  }
  const pixels = width * height;
  const channels = decoded.length / pixels;
  if (!Number.isInteger(channels) || ![1, 2, 3, 4].includes(channels)) {
    failure("IMAGE_DECODE_FAILED", "OpenJPEG returned an invalid pixel buffer");
  }
  const output = new Uint8ClampedArray(pixels * 4);
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    const source = pixel * channels;
    const target = pixel * 4;
    if (channels <= 2) {
      output[target] = decoded[source]!;
      output[target + 1] = decoded[source]!;
      output[target + 2] = decoded[source]!;
      output[target + 3] = !forceOpaque && channels === 2 ? decoded[source + 1]! : 255;
    } else {
      output[target] = decoded[source]!;
      output[target + 1] = decoded[source + 1]!;
      output[target + 2] = decoded[source + 2]!;
      output[target + 3] = !forceOpaque && channels === 4 ? decoded[source + 3]! : 255;
    }
  }
  if (matte !== undefined) {
    for (let index = 0; index < output.length; index += 4) {
      const alpha = output[index + 3]!;
      if (alpha === 0) {
        output[index] = 255;
        output[index + 1] = 255;
        output[index + 2] = 255;
        continue;
      }
      const scale = 255 / alpha;
      output[index] = (output[index]! - matte[0]!) * scale + matte[0]!;
      output[index + 1] = (output[index + 1]! - matte[1]!) * scale + matte[1]!;
      output[index + 2] = (output[index + 2]! - matte[2]!) * scale + matte[2]!;
    }
  }
  const approximate = channels === 2
    || channels === 4
    || (channels === 3 && (!structure.boxed || structure.enumeratedColorSpace !== 16))
    || (channels === 1
      && structure.enumeratedColorSpace !== undefined
      && structure.enumeratedColorSpace !== 17)
    || structure.components.some((component) => component.precision !== 8
      || component.signed
      || component.xSubsampling !== 1
      || component.ySubsampling !== 1);
  return rgbaOutput(output, width, height, maxPixels, approximate || reducePower > 0);
}

let resvgInitialization: Promise<void> | undefined;

function initializeResvg(): Promise<void> {
  resvgInitialization ??= initWasm(new URL("./resvg.wasm", import.meta.url));
  return resvgInitialization;
}

const MAX_SVG_RESOURCE_DEPTH = 8;
const SVG_DATA_IMAGE_FORMATS = new Set<OfficeImageFormat>(["png", "jpeg", "gif", "webp", "bmp", "dib", "tiff", "svg"]);
const SVG_DATA_IMAGE_TYPES = new Set([
  "image/png",
  "image/x-png",
  "image/jpeg",
  "image/jpg",
  "image/jpe",
  "image/jfif",
  "image/pjpeg",
  "image/gif",
  "image/webp",
  "image/bmp",
  "image/x-bmp",
  "image/x-ms-bmp",
  "image/tiff",
  "image/tif",
  "image/svg+xml",
  "image/svg",
]);

interface SvgResourceBudget {
  bytes: number;
  pixels: number;
}

function decodeXmlReferences(value: string): string {
  const decoded = value.replace(/&(#(?:x[0-9a-f]+|\d+)|amp|quot|apos|lt|gt);/giu, (reference, entity: string) => {
    const lower = entity.toLowerCase();
    const named = { amp: "&", quot: "\"", apos: "'", lt: "<", gt: ">" }[lower as "amp" | "quot" | "apos" | "lt" | "gt"];
    if (named !== undefined) return named;
    const codePoint = lower.startsWith("#x")
      ? Number.parseInt(lower.slice(2), 16)
      : Number.parseInt(lower.slice(1), 10);
    if (!Number.isSafeInteger(codePoint) || codePoint <= 0 || codePoint > 0x10_ffff
      || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG resource URI contains an invalid character reference");
    }
    return String.fromCodePoint(codePoint);
  });
  if (/&(?:#|[a-z_][\w.-]*;)/iu.test(decoded)) {
    failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG resource URI contains an unsupported entity reference");
  }
  return decoded;
}

function cssResourceUrls(source: string): string[] {
  const urls: string[] = [];
  const expression = /\burl\s*\(/giu;
  for (let match = expression.exec(source); match !== null; match = expression.exec(source)) {
    let offset = expression.lastIndex;
    while (/\s/u.test(source[offset] ?? "")) offset += 1;
    const quote = source[offset] === "\"" || source[offset] === "'" ? source[offset++]! : undefined;
    const start = offset;
    if (quote !== undefined) {
      while (offset < source.length && source[offset] !== quote) {
        if (source[offset] === "\\") {
          failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "Escaped SVG CSS resource URIs are not allowed");
        }
        offset += 1;
      }
      if (offset >= source.length) failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG CSS resource URI is unterminated");
      urls.push(source.slice(start, offset));
      offset += 1;
      while (/\s/u.test(source[offset] ?? "")) offset += 1;
      if (source[offset] !== ")") failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG CSS resource URI is malformed");
    } else {
      while (offset < source.length && source[offset] !== ")") {
        if (source[offset] === "\\" || source[offset] === "\"" || source[offset] === "'") {
          failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "Escaped SVG CSS resource URIs are not allowed");
        }
        offset += 1;
      }
      if (offset >= source.length) failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG CSS resource URI is unterminated");
      urls.push(source.slice(start, offset).trim());
    }
    expression.lastIndex = offset + 1;
  }
  return urls;
}

function decodeSvgDataUri(value: string, maxBytes: number, budget: SvgResourceBudget): {
  readonly bytes: Uint8Array;
  readonly mediaType: string;
} {
  const comma = value.indexOf(",");
  if (comma < 5 || value.slice(0, 5).toLowerCase() !== "data:") {
    failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG image resource is not a valid data URI");
  }
  const metadata = value.slice(5, comma).split(";").map((part) => part.trim());
  const mediaType = metadata.shift()?.toLowerCase() ?? "";
  if (!SVG_DATA_IMAGE_TYPES.has(mediaType)) {
    failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", `SVG data resource type ${mediaType || "(missing)"} is unsupported`);
  }
  let base64 = false;
  for (const parameter of metadata) {
    const lower = parameter.toLowerCase();
    if (lower === "base64" && !base64) {
      base64 = true;
    } else if (lower === "utf8" || lower === "charset=utf-8" || lower === "charset=us-ascii") {
      if (mediaType !== "image/svg+xml" && mediaType !== "image/svg") {
        failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "Raster SVG data resources cannot declare a text charset");
      }
    } else {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", `SVG data resource parameter ${parameter || "(empty)"} is unsupported`);
    }
  }
  const payload = value.slice(comma + 1);
  let bytes: Uint8Array;
  if (base64) {
    const encoded = payload.replace(/[\t\n\f\r ]+/gu, "");
    const padding = encoded.endsWith("==") ? 2 : encoded.endsWith("=") ? 1 : 0;
    const estimatedBytes = Math.max(0, Math.floor(encoded.length * 3 / 4) - padding);
    if (estimatedBytes > maxBytes || budget.bytes + estimatedBytes > maxBytes) {
      failure("IMAGE_SIZE_LIMIT", "Embedded SVG image bytes exceed the configured limit");
    }
    let binary: string;
    try {
      binary = atob(encoded);
    } catch (cause) {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG image data URI has invalid base64", cause);
    }
    bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  } else {
    let decoded: string;
    try {
      decoded = decodeURIComponent(payload);
    } catch (cause) {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG image data URI has invalid percent encoding", cause);
    }
    bytes = new TextEncoder().encode(decoded);
  }
  if (bytes.length === 0 || bytes.length > maxBytes || budget.bytes + bytes.length > maxBytes) {
    failure("IMAGE_SIZE_LIMIT", "Embedded SVG image bytes exceed the configured limit");
  }
  budget.bytes += bytes.length;
  return { bytes, mediaType };
}

function inspectSvgDataImage(
  value: string,
  maxPixels: number,
  maxBytes: number,
  budget: SvgResourceBudget,
  depth: number,
): boolean {
  const resource = decodeSvgDataUri(value, maxBytes, budget);
  const identified = identifyImage(resource.bytes, resource.mediaType, maxPixels);
  if (!SVG_DATA_IMAGE_FORMATS.has(identified.format)) {
    failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", `Embedded SVG ${identified.format.toUpperCase()} images are unsupported`);
  }
  budget.pixels += identified.pixels;
  if (!Number.isSafeInteger(budget.pixels) || budget.pixels > maxPixels) {
    failure("IMAGE_DIMENSION_LIMIT", `Embedded SVG images exceed the cumulative ${maxPixels}-pixel limit`);
  }
  if (identified.format !== "svg") return false;
  if (depth >= MAX_SVG_RESOURCE_DEPTH) {
    failure("IMAGE_SVG_RECURSION_LIMIT", `Embedded SVG depth exceeds ${MAX_SVG_RESOURCE_DEPTH}`);
  }
  let nestedSource: string;
  try {
    nestedSource = new TextDecoder("utf-8", { fatal: true }).decode(resource.bytes);
  } catch (cause) {
    failure("IMAGE_HEADER_INVALID", "Embedded SVG text is not valid UTF-8", cause);
  }
  return inspectSvgResources(nestedSource, maxPixels, maxBytes, budget, depth + 1);
}

function inspectSvgResources(
  source: string,
  maxPixels: number,
  maxBytes: number,
  budget: SvgResourceBudget,
  depth: number,
): boolean {
  if (/<!DOCTYPE\b|@import\b|@font-face\b|<(?:[\w.-]+:)?font-face-uri(?:\s|>)/iu.test(source)) {
    failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG external and embedded font resources are not allowed");
  }
  for (const match of source.matchAll(/<(?:[\w.-]+:)?style(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?style\s*>|\bstyle\s*=\s*(["'])(.*?)\2/giu)) {
    if ((match[1] ?? match[3] ?? "").includes("\\")) {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "Escaped SVG CSS resources are not allowed");
    }
  }
  let hasText = /<(?:[\w.-]+:)?text(?:\s|>)/iu.test(source);
  for (const match of source.matchAll(/\b(?:href|xlink:href)\s*=\s*(["'])(.*?)\1/gisu)) {
    const tagStart = source.lastIndexOf("<", match.index);
    const tag = tagStart >= 0
      ? source.slice(tagStart, match.index).match(/^<\s*(?:[\w.-]+:)?([\w.-]+)/u)?.[1]?.toLowerCase()
      : undefined;
    const value = decodeXmlReferences(match[2]!).trim();
    if (value === "" || value.startsWith("#")) continue;
    if (value.slice(0, 5).toLowerCase() === "data:" && (tag === "image" || tag === "feimage")) {
      hasText ||= inspectSvgDataImage(value, maxPixels, maxBytes, budget, depth);
    } else {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG external or unsupported linked resources are not allowed");
    }
  }
  for (const rawValue of cssResourceUrls(source)) {
    let value = decodeXmlReferences(rawValue).trim();
    if ((value.startsWith("\"") && value.endsWith("\"")) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1).trim();
    }
    if (value === "" || value.startsWith("#")) continue;
    if (value.slice(0, 5).toLowerCase() === "data:") {
      hasText ||= inspectSvgDataImage(value, maxPixels, maxBytes, budget, depth);
    } else {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG external CSS resources are not allowed");
    }
  }
  return hasText;
}

function pngCrc32(bytes: Uint8Array): number {
  let crc = 0xffff_ffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb8_8320 : 0);
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = new TextEncoder().encode(type);
  const chunk = new Uint8Array(12 + data.length);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, data.length, false);
  chunk.set(typeBytes, 4);
  chunk.set(data, 8);
  view.setUint32(8 + data.length, pngCrc32(chunk.subarray(4, 8 + data.length)), false);
  return chunk;
}

function encodeRgbaPng(output: RgbaCodecOutput): Uint8Array {
  const rows = new Uint8Array((output.width * 4 + 1) * output.height);
  for (let row = 0; row < output.height; row += 1) {
    rows.set(output.data.subarray(row * output.width * 4, (row + 1) * output.width * 4), row * (output.width * 4 + 1) + 1);
  }
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, output.width, false);
  view.setUint32(4, output.height, false);
  header.set([8, 6, 0, 0, 0], 8);
  const chunks = [
    new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk("IHDR", header),
    pngChunk("IDAT", (pako as unknown as { deflate(bytes: Uint8Array): Uint8Array }).deflate(rows)),
    pngChunk("IEND", new Uint8Array()),
  ];
  const png = new Uint8Array(chunks.reduce((length, chunk) => length + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    png.set(chunk, offset);
    offset += chunk.length;
  }
  return png;
}

function base64Bytes(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function normalizeSvgRasterImages(
  source: string,
  maxPixels: number,
  maxBytes: number,
  maxCompressionRatio: number,
): string {
  const expression = /(<(?:[\w.-]+:)?(?:image|feimage)\b[^>]*?\b(?:href|xlink:href)\s*=\s*)(["'])(.*?)\2/gisu;
  return source.replace(expression, (attribute, prefix: string, quote: string, rawValue: string) => {
    const value = decodeXmlReferences(rawValue).trim();
    if (value.slice(0, 5).toLowerCase() !== "data:") return attribute;
    const resource = decodeSvgDataUri(value, maxBytes, { bytes: 0, pixels: 0 });
    const identified = identifyImage(resource.bytes, resource.mediaType, maxPixels);
    const decoded = identified.format === "bmp" || identified.format === "dib"
      ? decodeDib(resource.bytes, maxPixels)
      : identified.format === "tiff"
        ? decodeTiff(resource.bytes, maxPixels, maxCompressionRatio)
        : undefined;
    if (decoded === undefined) return attribute;
    return `${prefix}${quote}data:image/png;base64,${base64Bytes(encodeRgbaPng(decoded))}${quote}`;
  });
}

function svgFontOptions(fonts: readonly OfficeImageCodecFont[]) {
  const defaultFont = fonts.find(({ family }) => approximateFontFamily(family) === "sans-serif") ?? fonts[0]!;
  const serifFont = fonts.find(({ family }) => approximateFontFamily(family) === "serif") ?? defaultFont;
  const monospaceFont = fonts.find(({ family }) => approximateFontFamily(family) === "monospace") ?? defaultFont;
  return {
    fontBuffers: fonts.map(({ bytes }) => new Uint8Array(bytes)),
    defaultFontFamily: defaultFont.family,
    sansSerifFamily: defaultFont.family,
    serifFamily: serifFont.family,
    monospaceFamily: monospaceFont.family,
  };
}

async function decodeSvg(
  bytes: Uint8Array,
  maxPixels: number,
  maxBytes: number,
  maxCompressionRatio: number,
  fonts: readonly OfficeImageCodecFont[],
  sourceWidth: number,
  sourceHeight: number,
  targetWidth?: number,
  targetHeight?: number,
): Promise<EncodedCodecOutput> {
  let renderer: InstanceType<typeof Resvg> | undefined;
  try {
    const source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const hasText = inspectSvgResources(source, maxPixels, maxBytes, { bytes: 0, pixels: 0 }, 0);
    if (hasText && fonts.length === 0) {
      failure(
        "IMAGE_SVG_FONT_UNAVAILABLE",
        "SVG text requires a document-embedded or host-provided font binary inside the isolated codec Worker",
      );
    }
    await initializeResvg();
    const rasterScale = targetWidth === undefined || targetHeight === undefined
      ? 1
      : Math.min(
        Math.max(targetWidth / sourceWidth, targetHeight / sourceHeight),
        MAX_VECTOR_DIMENSION / sourceWidth,
        MAX_VECTOR_DIMENSION / sourceHeight,
        Math.sqrt(maxPixels / (sourceWidth * sourceHeight)),
      );
    renderer = new Resvg(new TextEncoder().encode(normalizeSvgRasterImages(
      source,
      maxPixels,
      maxBytes,
      maxCompressionRatio,
    )), {
      font: hasText ? svgFontOptions(fonts) : { loadSystemFonts: false },
      fitTo: rasterScale === 1
        ? { mode: "original" }
        : { mode: "zoom", value: rasterScale },
    });
    const unresolved = renderer.imagesToResolve();
    if (unresolved.length !== 0) {
      failure("IMAGE_EXTERNAL_RESOURCE_BLOCKED", "SVG external image resources are not allowed");
    }
    assertDimensions(Math.ceil(renderer.width), Math.ceil(renderer.height), maxPixels);
    const rendered = renderer.render();
    try {
      assertDimensions(rendered.width, rendered.height, maxPixels);
      return {
        kind: "encoded",
        data: rendered.asPng(),
        mediaType: "image/png",
        width: rendered.width,
        height: rendered.height,
        approximate: hasText,
      };
    } finally {
      rendered.free();
    }
  } catch (cause) {
    if (cause instanceof OfficeEngineError) throw cause;
    return failure("IMAGE_DECODE_FAILED", "SVG rasterization failed", cause);
  } finally {
    renderer?.free();
  }
}

function boundedInflate(
  bytes: Uint8Array,
  maxBytes: number,
  maxCompressionRatio: number,
  label: string,
  decodedSizeLimit?: number,
): Uint8Array {
  const ratioLimit = bytes.length * maxCompressionRatio;
  const outputLimit = decodedSizeLimit === undefined
    ? Math.min(maxBytes, Math.max(1, ratioLimit))
    : Math.min(maxBytes, decodedSizeLimit);
  const chunks: Uint8Array[] = [];
  let length = 0;
  const inflater = new pako.Inflate({ chunkSize: 64 * 1024 });
  inflater.onData = (chunk) => {
    length += chunk.length;
    if (!Number.isSafeInteger(length) || length > outputLimit) {
      failure("IMAGE_DECOMPRESSION_LIMIT", `Compressed ${label} expands beyond ${outputLimit} bytes`);
    }
    chunks.push(chunk.slice());
  };
  try {
    inflater.push(bytes, true);
  } catch (cause) {
    if (cause instanceof OfficeEngineError) throw cause;
    failure("IMAGE_DECODE_FAILED", `Compressed ${label} decompression failed`, cause);
  }
  if (inflater.err !== 0) failure("IMAGE_DECODE_FAILED", inflater.msg || "Compressed metafile is invalid");
  const output = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }
  return output;
}

function boundedGunzip(bytes: Uint8Array, maxBytes: number, maxCompressionRatio: number): Uint8Array {
  return boundedInflate(bytes, maxBytes, maxCompressionRatio, "metafile");
}

function pdfPaeth(left: number, up: number, upperLeft: number): number {
  const prediction = left + up - upperLeft;
  const leftDistance = Math.abs(prediction - left);
  const upDistance = Math.abs(prediction - up);
  const upperLeftDistance = Math.abs(prediction - upperLeft);
  return leftDistance <= upDistance && leftDistance <= upperLeftDistance
    ? left : upDistance <= upperLeftDistance ? up : upperLeft;
}

function decodePdfRaster(
  bytes: Uint8Array,
  maxPixels: number,
  maxBytes: number,
  maxCompressionRatio: number,
): RgbaCodecOutput {
  const info = identifyImage(bytes, "image/x-officeviewer-pdf-raster", maxPixels);
  const components = bytes[16]!;
  const predictor = bytes[17]!;
  const paletteComponents = bytes[18]!;
  const paletteLength = paletteComponents * (bytes[19]! + 1);
  const compressedStart = paletteComponents === 0 ? 20 : 28;
  const compressedEnd = bytes.length - paletteLength;
  const compressed = bytes.subarray(compressedStart, compressedEnd);
  const palette = bytes.subarray(compressedEnd);
  if (compressed.length === 0) failure("IMAGE_HEADER_INVALID", "Compressed PDF raster has no payload");
  const rowBytes = info.width * components;
  const expected = rowBytes * info.height;
  const decodedSizeLimit = predictor >= 10 ? expected + info.height : expected;
  const samples = boundedInflate(
    compressed,
    maxBytes,
    maxCompressionRatio,
    "PDF raster",
    decodedSizeLimit,
  );
  const hasFilterByte = predictor >= 10 && samples.length === expected + info.height;
  if (samples.length !== expected && !hasFilterByte) {
    failure("IMAGE_DECODE_FAILED", "Compressed PDF raster sample length is invalid");
  }
  const decoded = new Uint8Array(expected);
  for (let row = 0; row < info.height; row += 1) {
    const sourceStart = hasFilterByte ? row * (rowBytes + 1) + 1 : row * rowBytes;
    const filter = predictor <= 1 ? 0
      : predictor === 2 ? 1
        : hasFilterByte ? samples[sourceStart - 1]! : predictor - 10;
    if (filter > 4) failure("IMAGE_DECODE_FAILED", "Compressed PDF raster predictor is invalid");
    const targetStart = row * rowBytes;
    for (let column = 0; column < rowBytes; column += 1) {
      const left = column >= components ? decoded[targetStart + column - components]! : 0;
      const up = row > 0 ? decoded[targetStart + column - rowBytes]! : 0;
      const upperLeft = row > 0 && column >= components
        ? decoded[targetStart + column - rowBytes - components]! : 0;
      const prediction = filter === 0 ? 0
        : filter === 1 ? left
          : filter === 2 ? up
            : filter === 3 ? Math.floor((left + up) / 2)
              : pdfPaeth(left, up, upperLeft);
      decoded[targetStart + column] = (samples[sourceStart + column]! + prediction) & 0xff;
    }
  }
  const rgba = new Uint8ClampedArray(info.width * info.height * 4);
  const paletteView = paletteComponents === 0
    ? undefined
    : new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const paletteDecodeLow = paletteView?.getFloat32(20) ?? 0;
  const paletteDecodeHigh = paletteView?.getFloat32(24) ?? 0;
  for (let pixel = 0; pixel < info.width * info.height; pixel += 1) {
    const source = pixel * components;
    const target = pixel * 4;
    if (paletteComponents > 0) {
      const paletteIndex = Math.min(bytes[19]!, Math.max(0, Math.round(
        paletteDecodeLow + decoded[source]! / 255 * (paletteDecodeHigh - paletteDecodeLow),
      )));
      const paletteSource = paletteIndex * paletteComponents;
      if (paletteComponents === 1) {
        rgba[target] = palette[paletteSource]!;
        rgba[target + 1] = palette[paletteSource]!;
        rgba[target + 2] = palette[paletteSource]!;
      } else if (paletteComponents === 3) {
        rgba[target] = palette[paletteSource]!;
        rgba[target + 1] = palette[paletteSource + 1]!;
        rgba[target + 2] = palette[paletteSource + 2]!;
      } else {
        const key = palette[paletteSource + 3]!;
        rgba[target] = 255 - Math.min(255, palette[paletteSource]! + key);
        rgba[target + 1] = 255 - Math.min(255, palette[paletteSource + 1]! + key);
        rgba[target + 2] = 255 - Math.min(255, palette[paletteSource + 2]! + key);
      }
    } else if (components === 1) {
      rgba[target] = decoded[source]!;
      rgba[target + 1] = decoded[source]!;
      rgba[target + 2] = decoded[source]!;
    } else if (components === 3) {
      rgba[target] = decoded[source]!;
      rgba[target + 1] = decoded[source + 1]!;
      rgba[target + 2] = decoded[source + 2]!;
    } else {
      const key = decoded[source + 3]!;
      rgba[target] = 255 - Math.min(255, decoded[source]! + key);
      rgba[target + 1] = 255 - Math.min(255, decoded[source + 1]! + key);
      rgba[target + 2] = 255 - Math.min(255, decoded[source + 2]! + key);
    }
    rgba[target + 3] = 255;
  }
  return rgbaOutput(rgba, info.width, info.height, maxPixels);
}

function pngDataUrlBytes(value: string): Uint8Array {
  const prefix = "data:image/png;base64,";
  if (!value.startsWith(prefix)) failure("IMAGE_DECODE_FAILED", "Metafile converter returned a non-PNG result");
  let binary: string;
  try {
    binary = atob(value.slice(prefix.length));
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", "Metafile PNG output is invalid", cause);
  }
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function decodeFullFrameWmfStretchDib(
  bytes: Uint8Array,
  maxPixels: number,
): RgbaCodecOutput | undefined {
  const headerOffset = bytes.length >= 22 && u32le(bytes, 0) === 0x9ac6_cdd7 ? 22 : 0;
  if (bytes.length < headerOffset + 18) return undefined;
  const headerWords = u16le(bytes, headerOffset + 2);
  if (headerWords < 9) return undefined;
  let offset = headerOffset + headerWords * 2;
  let composite: RgbaCodecOutput | undefined;
  for (let records = 0; offset + 6 <= bytes.length && records < MAX_METAFILE_RECORDS; records += 1) {
    const size = u32le(bytes, offset) * 2;
    if (size < 6 || offset + size > bytes.length) return undefined;
    const kind = u16le(bytes, offset + 4);
    if (kind === 0x0f43 && size > 28) {
      const sourceHeight = i16le(bytes, offset + 12);
      const sourceWidth = i16le(bytes, offset + 14);
      const fullFrame = sourceWidth > 0 && sourceHeight > 0
        && i16le(bytes, offset + 16) === 0 && i16le(bytes, offset + 18) === 0
        && i16le(bytes, offset + 20) === sourceHeight
        && i16le(bytes, offset + 22) === sourceWidth
        && i16le(bytes, offset + 24) === 0 && i16le(bytes, offset + 26) === 0;
      if (fullFrame) {
        try {
          const preview = decodeDib(bytes.subarray(offset + 28, offset + size), maxPixels);
          if (preview.width === sourceWidth && preview.height === sourceHeight) {
            const rasterOperation = u32le(bytes, offset + 6);
            if (rasterOperation === 0x00cc_0020) { // SRCCOPY
              composite = { ...preview, approximate: true };
            } else if (rasterOperation === 0x00ee_0086 || rasterOperation === 0x0088_00c6 || rasterOperation === 0x0066_0046) {
              if (composite !== undefined
                && (composite.width !== preview.width || composite.height !== preview.height)) return undefined;
              const data = composite?.data.slice() ?? new Uint8ClampedArray(preview.data.length).fill(255);
              for (let index = 0; index < data.length; index += 4) {
                for (let channel = 0; channel < 3; channel += 1) {
                  data[index + channel] = rasterOperation === 0x00ee_0086
                    ? data[index + channel]! | preview.data[index + channel]!
                    : rasterOperation === 0x0088_00c6
                      ? data[index + channel]! & preview.data[index + channel]!
                      : data[index + channel]! ^ preview.data[index + channel]!;
                }
              }
              composite = { ...preview, data, approximate: true };
            } else {
              return undefined;
            }
          }
        } catch (error) {
          if (!(error instanceof OfficeEngineError)
            || !["IMAGE_HEADER_INVALID", "IMAGE_DECODE_FAILED"].includes(error.code)) throw error;
        }
      }
    }
    if (kind === 0) break;
    offset += size;
  }
  return composite;
}

async function decodeMetafile(
  bytes: Uint8Array,
  format: "emf" | "wmf",
  maxPixels: number,
  maxBytes: number,
  targetWidth?: number,
  targetHeight?: number,
): Promise<OfficeImageCodecOutput> {
  const info = identifyImage(bytes, format === "emf" ? "image/x-emf" : "image/x-wmf", maxPixels);
  const dimensionCap = Math.min(MAX_VECTOR_DIMENSION, Math.max(1, Math.floor(Math.sqrt(maxPixels))));
  let maxWidth = dimensionCap;
  let maxHeight = dimensionCap;
  if (targetWidth !== undefined && targetHeight !== undefined) {
    const scale = Math.min(
      1,
      MAX_VECTOR_DIMENSION / targetWidth,
      MAX_VECTOR_DIMENSION / targetHeight,
      Math.sqrt(maxPixels / (targetWidth * targetHeight)),
    );
    maxWidth = Math.max(1, Math.floor(targetWidth * scale));
    maxHeight = Math.max(1, Math.floor(targetHeight * scale));
  } else if (info.dimensionsKnown) {
    const scale = Math.min(
      1,
      MAX_VECTOR_DIMENSION / info.width,
      MAX_VECTOR_DIMENSION / info.height,
      Math.sqrt(maxPixels / info.pixels),
    );
    maxWidth = Math.max(1, Math.floor(info.width * scale));
    maxHeight = Math.max(1, Math.floor(info.height * scale));
  }
  let dpiScale = targetWidth !== undefined
    && targetHeight !== undefined
    && info.dimensionsKnown
    ? Math.max(1, targetWidth / info.width, targetHeight / info.height)
    : 1;
  const safeBytes = secureMetafileForConversion(
    bytes,
    format,
    maxBytes,
    maxPixels,
    maxWidth,
    maxHeight,
    dpiScale,
  );
  if (format === "wmf" && !info.dimensionsKnown
    && targetWidth !== undefined && targetHeight !== undefined) {
    const placeable = safeBytes.length >= 22 && u32le(safeBytes, 0) === 0x9ac6_cdd7;
    const width = placeable ? Math.abs(i16le(safeBytes, 10) - i16le(safeBytes, 6)) : 800;
    const height = placeable ? Math.abs(i16le(safeBytes, 12) - i16le(safeBytes, 8)) : 600;
    if (width > 0 && height > 0) dpiScale = Math.max(1, targetWidth / width, targetHeight / height);
  }
  if (format === "wmf") {
    const preview = decodeFullFrameWmfStretchDib(safeBytes, maxPixels);
    if (preview !== undefined) return preview;
  }
  let dataUrl: string | null;
  try {
    const buffer = safeBytes.slice().buffer;
    const options = {
      dpiScale,
      maxCanvasDimension: MAX_VECTOR_DIMENSION,
      maxRecords: MAX_METAFILE_RECORDS,
      symbolFontText: semanticSymbolFontText,
    };
    dataUrl = format === "emf"
      ? await convertEmfToDataUrl(buffer, maxWidth, maxHeight, options)
      : await convertWmfToDataUrl(buffer, maxWidth, maxHeight, options);
  } catch (cause) {
    failure("IMAGE_DECODE_FAILED", `${format.toUpperCase()} conversion failed`, cause);
  }
  if (dataUrl === null) failure("IMAGE_DECODE_FAILED", `${format.toUpperCase()} converter rejected the metafile`);
  const pngBytes = pngDataUrlBytes(dataUrl);
  const png = identifyImage(pngBytes, "image/png", maxPixels);
  return {
    kind: "encoded",
    data: pngBytes,
    mediaType: "image/png",
    width: png.width,
    height: png.height,
    approximate: true,
  };
}

export async function decodeOfficeImagePayload(input: OfficeImageCodecInput): Promise<OfficeImageCodecOutput> {
  if (!Number.isSafeInteger(input.maxBytes) || input.maxBytes <= 0
    || !Number.isSafeInteger(input.maxPixels) || input.maxPixels <= 0
    || !Number.isSafeInteger(input.maxCompressionRatio) || input.maxCompressionRatio <= 0) {
    failure("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec limits are invalid");
  }
  const hasTarget = input.targetWidth !== undefined || input.targetHeight !== undefined;
  const validTargetDimensions = typeof input.targetWidth === "number"
    && typeof input.targetHeight === "number"
    && Number.isSafeInteger(input.targetWidth)
    && Number.isSafeInteger(input.targetHeight)
    && input.targetWidth > 0
    && input.targetHeight > 0;
  const targetPixels = validTargetDimensions
    ? input.targetWidth * input.targetHeight
    : 0;
  if (hasTarget
    && (!TARGETED_VECTOR_FORMATS.has(input.format)
      || !validTargetDimensions
      || !Number.isSafeInteger(targetPixels)
      || targetPixels > input.maxPixels)) {
    failure("IMAGE_CODEC_PROTOCOL_INVALID", "Image codec target dimensions are invalid");
  }
  const bytes = input.bytes instanceof Uint8Array ? input.bytes : new Uint8Array(input.bytes);
  if (bytes.length === 0 || bytes.length > input.maxBytes) {
    failure("IMAGE_SIZE_LIMIT", `Encoded image is ${bytes.length} bytes; limit is ${input.maxBytes}`);
  }
  const identified = identifyImage(
    bytes,
    input.mediaType,
    input.maxPixels,
    input.format === "jp2" || input.format === "j2k",
  );
  if (identified.format !== input.format) failure("IMAGE_CODEC_PROTOCOL_INVALID", "Image format changed across the codec boundary");

  switch (input.format) {
    case "bmp":
    case "dib": return decodeDib(bytes, input.maxPixels);
    case "ico": return decodeIcon(bytes, input.maxPixels);
    case "pcx": return decodePcx(bytes, input.maxPixels);
    case "tiff": return decodeTiff(bytes, input.maxPixels, input.maxCompressionRatio);
    case "jxr": {
      const decoded = await decodeJpegXr(bytes.slice().buffer);
      assertDimensions(decoded.width, decoded.height, input.maxPixels);
      return {
        kind: "rgba",
        data: new Uint8ClampedArray(decoded.data),
        width: decoded.width,
        height: decoded.height,
        approximate: false,
      };
    }
    case "jp2":
    case "j2k": return await decodeJpeg2000(bytes, input.maxPixels);
    case "svg": return decodeSvg(
      bytes,
      input.maxPixels,
      input.maxBytes,
      input.maxCompressionRatio,
      input.fonts ?? [],
      identified.width,
      identified.height,
      input.targetWidth,
      input.targetHeight,
    );
    case "emf": return decodeMetafile(
      bytes,
      "emf",
      input.maxPixels,
      input.maxBytes,
      input.targetWidth,
      input.targetHeight,
    );
    case "wmf": return decodeMetafile(
      bytes,
      "wmf",
      input.maxPixels,
      input.maxBytes,
      input.targetWidth,
      input.targetHeight,
    );
    case "emz": {
      const inflated = boundedGunzip(bytes, input.maxBytes, input.maxCompressionRatio);
      return decodeMetafile(
        inflated,
        "emf",
        input.maxPixels,
        input.maxBytes,
        input.targetWidth,
        input.targetHeight,
      );
    }
    case "wmz": {
      const inflated = boundedGunzip(bytes, input.maxBytes, input.maxCompressionRatio);
      return decodeMetafile(
        inflated,
        "wmf",
        input.maxPixels,
        input.maxBytes,
        input.targetWidth,
        input.targetHeight,
      );
    }
    case "pdf-raster": return decodePdfRaster(
      bytes,
      input.maxPixels,
      input.maxBytes,
      input.maxCompressionRatio,
    );
    case "png": {
      return { kind: "encoded", data: bytes.slice(), mediaType: "image/png", width: identified.width, height: identified.height, approximate: false };
    }
    default: failure("IMAGE_CODEC_PROTOCOL_INVALID", `${input.format.toUpperCase()} should use the browser-native decoder`);
  }
}
