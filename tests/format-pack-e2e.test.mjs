import assert from "node:assert/strict";
import { createCipheriv, createHash } from "node:crypto";
import { readFile, readdir } from "node:fs/promises";
import test from "node:test";
import { deflateSync } from "node:zlib";

import { extendedFormatPack } from "../dist/extended-formats.js";
import { createOfficeEngine } from "../dist/engine.js";
import { createZip } from "./zip-fixture.mjs";

const encoder = new TextEncoder();
const FREE_SECTOR = 0xffff_ffff;
const END_OF_CHAIN = 0xffff_fffe;
const FAT_SECTOR = 0xffff_fffd;
const SECTOR_BYTES = 512;
const LEGACY_STREAM_BYTES = 4096;
const xpsFont = await readFile(new URL(
  "./fixtures/LiberationSans-Regular.ttf",
  import.meta.url,
));
const suppliedOfdrwSeal = await readFile(new URL("./fixtures/ofdrw-h.ofd", import.meta.url));
const PDF_PASSWORD_PADDING = Uint8Array.of(
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41,
  0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80,
  0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
);

function concat(parts) {
  const result = new Uint8Array(parts.reduce((length, part) => length + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function md5(bytes) {
  return new Uint8Array(createHash("md5").update(bytes).digest());
}

function rc4(key, input) {
  const state = Uint8Array.from({ length: 256 }, (_, index) => index);
  let j = 0;
  for (let index = 0; index < 256; index += 1) {
    j = (j + state[index] + key[index % key.length]) & 255;
    [state[index], state[j]] = [state[j], state[index]];
  }
  const output = new Uint8Array(input.length);
  let i = 0;
  j = 0;
  for (let index = 0; index < input.length; index += 1) {
    i = (i + 1) & 255;
    j = (j + state[i]) & 255;
    [state[i], state[j]] = [state[j], state[i]];
    output[index] = input[index] ^ state[(state[i] + state[j]) & 255];
  }
  return output;
}

function paddedPdfPassword(value) {
  const bytes = encoder.encode(value);
  const result = new Uint8Array(32);
  const copied = bytes.subarray(0, 32);
  result.set(copied);
  result.set(PDF_PASSWORD_PADDING.subarray(0, 32 - copied.length), copied.length);
  return result;
}

function encryptedPdfFixture(password = "secret") {
  const ownerKey = md5(paddedPdfPassword("owner")).subarray(0, 5);
  const owner = rc4(ownerKey, paddedPdfPassword(password));
  const permissions = -4;
  const fileId = encoder.encode("0123456789abcdef");
  const permissionBytes = new Uint8Array(4);
  new DataView(permissionBytes.buffer).setInt32(0, permissions, true);
  const fileKey = md5(concat([
    paddedPdfPassword(password),
    owner,
    permissionBytes,
    fileId,
  ])).subarray(0, 5);
  const user = rc4(fileKey, PDF_PASSWORD_PADDING);
  const objectBytes = Uint8Array.of(4, 0, 0, 0, 0);
  const objectKey = md5(concat([fileKey, objectBytes])).subarray(0, 10);
  const content = encoder.encode("BT /F1 12 Tf 1 0 0 1 72 720 Tm (Protected PDF) Tj ET\n");
  const encryptedContent = rc4(objectKey, content);
  const hex = (bytes) => Buffer.from(bytes).toString("hex");
  return concat([
    encoder.encode(
      `%PDF-1.4\n`
      + `1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n`
      + `2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >> endobj\n`
      + `3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj\n`
      + `4 0 obj << /Length ${encryptedContent.length} >> stream\n`,
    ),
    encryptedContent,
    encoder.encode(
      `\nendstream endobj\n`
      + `5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n`
      + `6 0 obj << /Filter /Standard /V 1 /R 2 /O <${hex(owner)}> /U <${hex(user)}> /P ${permissions} >> endobj\n`
      + `trailer << /Root 1 0 R /Encrypt 6 0 R /ID [<${hex(fileId)}> <${hex(fileId)}>] >>\n%%EOF`,
    ),
  ]);
}

function aesCbcEncrypt(key, iv, input, padding = true) {
  const cipher = createCipheriv(`aes-${key.length * 8}-cbc`, key, iv);
  cipher.setAutoPadding(padding);
  return new Uint8Array(Buffer.concat([cipher.update(input), cipher.final()]));
}

function encryptedPdfAes128Fixture(password = "secret") {
  const fileId = encoder.encode("aes128-file-id!!");
  const permissions = -4;
  const permissionBytes = new Uint8Array(4);
  new DataView(permissionBytes.buffer).setInt32(0, permissions, true);
  let ownerDigest = md5(paddedPdfPassword("owner"));
  for (let round = 0; round < 50; round += 1) ownerDigest = md5(ownerDigest);
  const ownerKey = ownerDigest.subarray(0, 16);
  let owner = rc4(ownerKey, paddedPdfPassword(password));
  for (let round = 1; round <= 19; round += 1) {
    owner = rc4(ownerKey.map((byte) => byte ^ round), owner);
  }
  let digest = md5(concat([paddedPdfPassword(password), owner, permissionBytes, fileId]));
  for (let round = 0; round < 50; round += 1) digest = md5(digest.subarray(0, 16));
  const fileKey = digest.subarray(0, 16);
  let user = rc4(fileKey, md5(concat([PDF_PASSWORD_PADDING, fileId])));
  for (let round = 1; round <= 19; round += 1) {
    user = rc4(fileKey.map((byte) => byte ^ round), user);
  }
  user = concat([user, new Uint8Array(16)]);
  const objectKey = md5(concat([fileKey, Uint8Array.of(4, 0, 0, 0, 0), encoder.encode("sAlT")]));
  const content = encoder.encode("BT /F1 12 Tf 1 0 0 1 72 720 Tm (AES 128 PDF) Tj ET\n");
  const iv = Uint8Array.from({ length: 16 }, (_, index) => index);
  const encryptedContent = concat([iv, aesCbcEncrypt(objectKey, iv, content)]);
  const hex = (bytes) => Buffer.from(bytes).toString("hex");
  return concat([
    encoder.encode(
      `%PDF-1.6\n`
      + `1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n`
      + `2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >> endobj\n`
      + `3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj\n`
      + `4 0 obj << /Length ${encryptedContent.length} >> stream\n`,
    ),
    encryptedContent,
    encoder.encode(
      `\nendstream endobj\n`
      + `5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n`
      + `6 0 obj << /Filter /Standard /V 4 /R 4 /Length 128 /O <${hex(owner)}> /U <${hex(user)}> /P ${permissions} `
      + `/CF << /StdCF << /CFM /AESV2 /Length 16 >> >> /StmF /StdCF /StrF /StdCF >> endobj\n`
      + `trailer << /Root 1 0 R /Encrypt 6 0 R /ID [<${hex(fileId)}> <${hex(fileId)}>] >>\n%%EOF`,
    ),
  ]);
}

function encryptedPdfAes256R5Fixture(password = "secret") {
  const passwordBytes = encoder.encode(password);
  const fileId = encoder.encode("aes256-file-id!!");
  const fileKey = Uint8Array.from({ length: 32 }, (_, index) => 0x40 + index);
  const userValidationSalt = encoder.encode("user-val");
  const userKeySalt = encoder.encode("user-key");
  const user = concat([
    new Uint8Array(createHash("sha256").update(concat([passwordBytes, userValidationSalt])).digest()),
    userValidationSalt,
    userKeySalt,
  ]);
  const zeroIv = new Uint8Array(16);
  const userIntermediate = new Uint8Array(
    createHash("sha256").update(concat([passwordBytes, userKeySalt])).digest(),
  );
  const userEncrypted = aesCbcEncrypt(userIntermediate, zeroIv, fileKey, false);
  const ownerPassword = encoder.encode("owner");
  const ownerValidationSalt = encoder.encode("ownr-val");
  const ownerKeySalt = encoder.encode("ownr-key");
  const owner = concat([
    new Uint8Array(
      createHash("sha256").update(concat([ownerPassword, ownerValidationSalt, user])).digest(),
    ),
    ownerValidationSalt,
    ownerKeySalt,
  ]);
  const ownerIntermediate = new Uint8Array(
    createHash("sha256").update(concat([ownerPassword, ownerKeySalt, user])).digest(),
  );
  const ownerEncrypted = aesCbcEncrypt(ownerIntermediate, zeroIv, fileKey, false);
  const content = encoder.encode("BT /F1 12 Tf 1 0 0 1 72 720 Tm (AES 256 PDF) Tj ET\n");
  const iv = Uint8Array.from({ length: 16 }, (_, index) => 0x20 + index);
  const encryptedContent = concat([iv, aesCbcEncrypt(fileKey, iv, content)]);
  const hex = (bytes) => Buffer.from(bytes).toString("hex");
  return concat([
    encoder.encode(
      `%PDF-1.7\n`
      + `1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n`
      + `2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >> endobj\n`
      + `3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj\n`
      + `4 0 obj << /Length ${encryptedContent.length} >> stream\n`,
    ),
    encryptedContent,
    encoder.encode(
      `\nendstream endobj\n`
      + `5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n`
      + `6 0 obj << /Filter /Standard /V 5 /R 5 /Length 256 /O <${hex(owner)}> /U <${hex(user)}> `
      + `/OE <${hex(ownerEncrypted)}> /UE <${hex(userEncrypted)}> /P -4 `
      + `/CF << /StdCF << /CFM /AESV3 /Length 32 >> >> /StmF /StdCF /StrF /StdCF >> endobj\n`
      + `trailer << /Root 1 0 R /Encrypt 6 0 R /ID [<${hex(fileId)}> <${hex(fileId)}>] >>\n%%EOF`,
    ),
  ]);
}

function view(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function writeU16(bytes, offset, value, littleEndian = true) {
  view(bytes).setUint16(offset, value, littleEndian);
}

function writeU32(bytes, offset, value, littleEndian = true) {
  view(bytes).setUint32(offset, value >>> 0, littleEndian);
}

function writeI32(bytes, offset, value) {
  view(bytes).setInt32(offset, value, true);
}

function installRecordingCanvas() {
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "OffscreenCanvas");
  const pathDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Path2D");
  const imageBitmapDescriptor = Object.getOwnPropertyDescriptor(globalThis, "createImageBitmap");
  const calls = { clips: 0, fills: [], images: 0, text: [], strokes: [] };
  let font = "";
  let fillStyle = "";
  let strokeStyle = "";
  let lineWidth = 1;
  const context = {
    save() {},
    restore() {},
    scale() {},
    translate() {},
    transform() {},
    setTransform() {},
    rotate() {},
    fillRect() {},
    beginPath() {},
    closePath() {},
    rect() {},
    ellipse() {},
    moveTo() {},
    lineTo() {},
    quadraticCurveTo() {},
    bezierCurveTo() {},
    fill() { calls.fills.push(fillStyle); },
    clip() { calls.clips += 1; },
    drawImage() { calls.images += 1; },
    createImageData(width, height) {
      return { data: new Uint8ClampedArray(width * height * 4), width, height };
    },
    putImageData() {},
    getTransform() {
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    },
    createPattern() {
      return { setTransform() {} };
    },
    setLineDash() {},
    measureText(value) {
      return { width: value.length * 8 };
    },
    fillText(text, x, y) {
      calls.text.push({ text, x, y, font, fillStyle });
    },
    stroke() {
      calls.strokes.push({ strokeStyle, lineWidth });
    },
    set font(value) {
      font = value;
    },
    set fillStyle(value) {
      fillStyle = value;
    },
    set strokeStyle(value) {
      strokeStyle = value;
    },
    set lineWidth(value) {
      lineWidth = value;
    },
    globalAlpha: 1,
  };
  class RecordingCanvas {
    constructor(width = 1, height = 1) {
      this.width = width;
      this.height = height;
    }

    getContext() {
      context.canvas = this;
      return context;
    }

    transferToImageBitmap() {
      return { close() {} };
    }
  }
  Object.defineProperty(globalThis, "OffscreenCanvas", {
    configurable: true,
    value: RecordingCanvas,
  });
  Object.defineProperty(globalThis, "createImageBitmap", {
    configurable: true,
    value: async () => ({ width: 120, height: 120, close() {} }),
  });
  class RecordingPath2D {
    moveTo() {}
    lineTo() {}
    quadraticCurveTo() {}
    bezierCurveTo() {}
    closePath() {}
    rect() {}
    ellipse() {}
  }
  Object.defineProperty(globalThis, "Path2D", {
    configurable: true,
    value: RecordingPath2D,
  });
  return {
    calls,
    restore() {
      if (descriptor) Object.defineProperty(globalThis, "OffscreenCanvas", descriptor);
      else delete globalThis.OffscreenCanvas;
      if (pathDescriptor) Object.defineProperty(globalThis, "Path2D", pathDescriptor);
      else delete globalThis.Path2D;
      if (imageBitmapDescriptor) Object.defineProperty(globalThis, "createImageBitmap", imageBitmapDescriptor);
      else delete globalThis.createImageBitmap;
    },
  };
}

function crc32(bytes) {
  let crc = 0xffff_ffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb8_8320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffff_ffff) >>> 0;
}

function pngChunk(type, payload) {
  const typeBytes = encoder.encode(type);
  const length = new Uint8Array(4);
  const checksum = new Uint8Array(4);
  writeU32(length, 0, payload.length, false);
  writeU32(checksum, 0, crc32(concat([typeBytes, payload])), false);
  return concat([length, typeBytes, payload, checksum]);
}

function png(width, height) {
  const header = new Uint8Array(13);
  writeU32(header, 0, width, false);
  writeU32(header, 4, height, false);
  header.set([8, 6, 0, 0, 0], 8); // 8-bit RGBA, standard compression/filter, no interlace.

  const stride = 1 + width * 4;
  const pixels = new Uint8Array(stride * height);
  for (let row = 0; row < height; row += 1) {
    const rowOffset = row * stride;
    pixels[rowOffset] = 0; // PNG filter method: None.
    for (let column = 0; column < width; column += 1) {
      pixels.set([0x24, 0x68, 0xac, 0xff], rowOffset + 1 + column * 4);
    }
  }

  return concat([
    Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a),
    pngChunk("IHDR", header),
    pngChunk("IDAT", new Uint8Array(deflateSync(pixels))),
    pngChunk("IEND", new Uint8Array()),
  ]);
}

function varint(value) {
  const bytes = [];
  do {
    let byte = value & 0x7f;
    value = Math.floor(value / 128);
    if (value !== 0) byte |= 0x80;
    bytes.push(byte);
  } while (value !== 0);
  return Uint8Array.from(bytes);
}

function iwaDocument(messageType) {
  const payload = Uint8Array.of(0x08, 0x01);
  const messageInfo = concat([
    Uint8Array.of(0x08),
    varint(messageType),
    Uint8Array.of(0x18),
    varint(payload.length),
  ]);
  const archiveInfo = concat([
    Uint8Array.of(0x08, 0x01, 0x12),
    varint(messageInfo.length),
    messageInfo,
  ]);
  const uncompressed = concat([varint(archiveInfo.length), archiveInfo, payload]);
  assert.ok(uncompressed.length <= 60, "fixture must fit one Snappy literal");
  const snappy = concat([
    varint(uncompressed.length),
    Uint8Array.of((uncompressed.length - 1) << 2),
    uncompressed,
  ]);
  return concat([
    Uint8Array.of(
      0,
      snappy.length & 0xff,
      (snappy.length >>> 8) & 0xff,
      (snappy.length >>> 16) & 0xff,
    ),
    snappy,
  ]);
}

function createPagesFixture({ redundantLocalZip64 = false } = {}) {
  return createZip({
    "Index/Document.iwa": iwaDocument(10_000),
    "Metadata/Properties.plist": encoder.encode(
      '<?xml version="1.0" encoding="UTF-8"?><plist version="1.0"><dict/></plist>',
    ),
    "preview.png": png(3, 2),
  }, { compress: false, redundantLocalZip64 });
}

function createPagesDirectoryPackageFixture() {
  return createZip({
    "Report.pages/Index.zip": Uint8Array.of(0),
    "Report.pages/Metadata/Properties.plist": encoder.encode("bplist00directory-package"),
  }, { compress: false });
}

function createOdtFixture() {
  return createZip({
    mimetype: encoder.encode("application/vnd.oasis.opendocument.text"),
    "META-INF/manifest.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
        <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
        <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
        <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
      </manifest:manifest>`),
    "styles.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">
        <office:automatic-styles><style:page-layout style:name="page"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin="1in"/></style:page-layout></office:automatic-styles>
        <office:master-styles><style:master-page style:name="Standard" style:page-layout-name="page"/></office:master-styles>
      </office:document-styles>`),
    "content.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
        <office:body><office:text><text:p>Lazy ODF</text:p></office:text></office:body>
      </office:document-content>`),
  }, { compress: false });
}

function createXpsFixture() {
  return createZip({
    "[Content_Types].xml": encoder.encode(`<?xml version="1.0"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="fdseq" ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>
        <Default Extension="fdoc" ContentType="application/vnd.ms-package.xps-fixeddocument+xml"/>
        <Default Extension="fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/>
        <Default Extension="ttf" ContentType="application/vnd.ms-opentype"/>
      </Types>`),
    "_rels/.rels": encoder.encode(`<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="r1" Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" Target="FixedDocSeq.fdseq"/>
      </Relationships>`),
    "FixedDocSeq.fdseq": encoder.encode(`<?xml version="1.0"?>
      <FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06">
        <DocumentReference Source="Documents/1/FixedDoc.fdoc"/>
      </FixedDocumentSequence>`),
    "Documents/1/FixedDoc.fdoc": encoder.encode(`<?xml version="1.0"?>
      <FixedDocument xmlns="http://schemas.microsoft.com/xps/2005/06">
        <PageContent Source="Pages/1.fpage"/>
        <PageContent Source="Pages/2.fpage"/>
      </FixedDocument>`),
    "Documents/1/Pages/_rels/1.fpage.rels": encoder.encode(`<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="font" Type="http://schemas.openxps.org/oxps/v1.0/required-resource" Target="../Resources/Fonts/LiberationSans-Regular.ttf"/>
      </Relationships>`),
    "Documents/1/Pages/_rels/2.fpage.rels": encoder.encode(`<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="font" Type="http://schemas.openxps.org/oxps/v1.0/required-resource" Target="../Resources/Fonts/LiberationSans-Regular.ttf"/>
      </Relationships>`),
    "Documents/1/Resources/Fonts/LiberationSans-Regular.ttf": xpsFont,
    "Documents/1/Pages/1.fpage": encoder.encode(`<?xml version="1.0"?>
      <FixedPage xmlns="http://schemas.microsoft.com/xps/2005/06" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="320" Height="240">
        <FixedPage.Resources>
          <ResourceDictionary>
            <MatrixTransform x:Key="motifShift" Matrix="1,0,0,1,3,2"/>
            <VisualBrush x:Key="motifBrush" Visual="{StaticResource motif}" Viewbox="0,0,1,1" ViewboxUnits="Absolute" Viewport="0,0,1,1" ViewportUnits="Absolute" Stretch="Uniform" Transform="300,0,0,100,0,0"/>
            <Canvas x:Key="motif" RenderTransform="1,0,0,1,1,1">
              <Path Data="M 0,0 L 10,0 L 10,10 L 0,10 Z" Fill="#FF1166CC"/>
              <Path Data="M 2,2 L 8,2 L 2,8 Z" Fill="#FFFFCC00"/>
            </Canvas>
          </ResourceDictionary>
        </FixedPage.Resources>
        <Canvas RenderTransform="1,0,0,1,5,6">
          <Path Data="F1 M 10,10 L 100,10 L 100,60 L 10,60 Z" Fill="#FFFF0000"/>
          <Path Data="M 10,65 L 100,65 L 100,90 L 10,90 Z">
            <Path.Fill><RadialGradientBrush RadiusX="0.45" RadiusY="0.8" GradientOrigin="0.3,0.4" Center="0.5,0.5" SpreadMethod="Reflect" ColorInterpolationMode="ScRgbLinearInterpolation" RelativeTransform="1,0.15,0,1,0,0"><GradientStop Offset="0" Color="#FFFFFFFF"/><GradientStop Offset="1" Color="#FF0066CC"/></RadialGradientBrush></Path.Fill>
          </Path>
          <Path Data="F1 M 120,10 L 300,10 L 300,90 L 120,90 Z">
            <Path.Fill>
              <VisualBrush Visual="{StaticResource motif}" Viewbox="0,0,10,10" ViewboxUnits="Absolute" Viewport="120,10,24,24" ViewportUnits="Absolute" TileMode="FlipXY" Stretch="Fill" Transform="{StaticResource motifShift}"/>
            </Path.Fill>
          </Path>
          <Path Data="F1 M 120,100 L 300,100 L 300,150 L 120,150 Z" Fill="{StaticResource motifBrush}"/>
          <Glyphs FontUri="../Resources/Fonts/LiberationSans-Regular.ttf" FontRenderingEmSize="20" OriginX="20" OriginY="100" UnicodeString="Lazy XPS" Fill="#FF000000" Indices=",55;,55;,55;,55;,30;,55;,55;,55"/>
          <Glyphs FontUri="../Resources/Fonts/LiberationSans-Regular.ttf" FontRenderingEmSize="18" OriginX="280" OriginY="130" BidiLevel="1" StyleSimulations="BoldItalicSimulation" Fill="#FF004488" Indices="36,100;41,100;52,100"/>
          <Glyphs FontUri="../Resources/Fonts/LiberationSans-Regular.ttf" FontRenderingEmSize="18" OriginX="20" OriginY="150" IsSideways="true" Fill="#FF448800" Indices="36,100;41,100;52,100"/>
          <Path Data="M 120,160 L 300,160 L 300,185 L 120,185 Z" Fill="#FF7A2CC8">
            <Path.OpacityMask><LinearGradientBrush MappingMode="Absolute" StartPoint="120,160" EndPoint="300,160"><GradientStop Offset="0" Color="#00000000"/><GradientStop Offset="1" Color="#FFFFFFFF"/></LinearGradientBrush></Path.OpacityMask>
          </Path>
          <Canvas>
            <Canvas.OpacityMask><LinearGradientBrush MappingMode="Absolute" StartPoint="120,195" EndPoint="300,195"><GradientStop Offset="0" Color="#00000000"/><GradientStop Offset="1" Color="#FFFFFFFF"/></LinearGradientBrush></Canvas.OpacityMask>
            <Path Data="M 120,195 L 230,195 L 230,225 L 120,225 Z" Fill="#FF0088CC"/>
            <Path Data="M 190,195 L 300,195 L 300,225 L 190,225 Z" Fill="#FFFF8800"/>
          </Canvas>
        </Canvas>
      </FixedPage>`),
    "Documents/1/Pages/2.fpage": encoder.encode(`<?xml version="1.0"?>
      <FixedPage xmlns="http://schemas.microsoft.com/xps/2005/06" Width="320" Height="240">
        <Glyphs FontUri="../Resources/Fonts/LiberationSans-Regular.ttf" FontRenderingEmSize="20" OriginX="20" OriginY="100" UnicodeString="Second XPS" Fill="#FF000000" Indices=",55;,55;,55;,55;,55;,30;,55;,55;,55;,55"/>
      </FixedPage>`),
  });
}

// XML content from OFDRW's Apache-2.0 `ofdrw-reader` helloworld.ofd fixture.
function derOctetString(bytes) {
  const length = bytes.length < 128
    ? Uint8Array.of(bytes.length)
    : Uint8Array.of(0x82, bytes.length >>> 8, bytes.length & 0xff);
  return concat([Uint8Array.of(0x04), length, bytes]);
}

function createSealOfdFixture() {
  return createZip({
    "OFD.xml": encoder.encode(`<?xml version="1.0"?><ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016" Version="1.1"><ofd:DocBody><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot></ofd:DocBody></ofd:OFD>`),
    "Doc_0/Document.xml": encoder.encode(`<?xml version="1.0"?><ofd:Document xmlns:ofd="http://www.ofdspec.org/2016"><ofd:CommonData><ofd:PageArea><ofd:PhysicalBox>0 0 20 20</ofd:PhysicalBox></ofd:PageArea></ofd:CommonData><ofd:Pages><ofd:Page ID="1" BaseLoc="Pages/Page_0/Content.xml"/></ofd:Pages></ofd:Document>`),
    "Doc_0/Pages/Page_0/Content.xml": encoder.encode(`<?xml version="1.0"?><ofd:Page xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Content><ofd:Layer ID="1"><ofd:PathObject ID="2" Boundary="0 0 20 20" Fill="true" Stroke="false"><ofd:FillColor Value="196 0 0"/><ofd:AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</ofd:AbbreviatedData></ofd:PathObject></ofd:Layer></ofd:Content></ofd:Page>`),
  });
}

function createOfdFixture({ tamperSignatureReference = false } = {}) {
  const seal = createSealOfdFixture();
  const sealContainer = derOctetString(seal);
  return createZip({
    "OFD.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016" Version="1.1" DocType="OFD"><ofd:DocBody><ofd:DocInfo><ofd:DocID>220c5913ebfe4f6e8070dabd3647f157</ofd:DocID><ofd:CreationDate>2020-09-21</ofd:CreationDate><ofd:Creator>OFD R&amp;W</ofd:Creator><ofd:CreatorVersion>1.5.5</ofd:CreatorVersion></ofd:DocInfo><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot><ofd:Signatures>Doc_0/Signs/Signatures.xml</ofd:Signatures></ofd:DocBody></ofd:OFD>`),
    "Doc_0/Document.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Document xmlns:ofd="http://www.ofdspec.org/2016"><ofd:CommonData><ofd:PublicRes>PublicRes.xml</ofd:PublicRes><ofd:DocumentRes>DocumentRes.xml</ofd:DocumentRes><ofd:TemplatePage ID="8" BaseLoc="Tpls/Tpl_0/Content.xml"/><ofd:MaxUnitID>13</ofd:MaxUnitID></ofd:CommonData><ofd:Pages><ofd:Page ID="1" BaseLoc="Pages/Page_0/Content.xml"/><ofd:Page ID="12" BaseLoc="Pages/Page_1/Content.xml"/></ofd:Pages><ofd:Outlines><ofd:OutlineElem Title="首页"><ofd:Actions><ofd:Action Event="CLICK"><ofd:Goto><ofd:Dest PageID="1" Type="XYZ"/></ofd:Goto></ofd:Action></ofd:Actions></ofd:OutlineElem></ofd:Outlines><ofd:Annotations>Annots/Annotations.xml</ofd:Annotations></ofd:Document>`),
    "Doc_0/PublicRes.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Res xmlns:ofd="http://www.ofdspec.org/2016" BaseLoc="Res"><ofd:Fonts><ofd:Font FontName="宋体" FamilyName="宋体" ID="3"/></ofd:Fonts></ofd:Res>`),
    "Doc_0/DocumentRes.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Res xmlns:ofd="http://www.ofdspec.org/2016" BaseLoc="Res"><ofd:DrawParams><ofd:DrawParam ID="9" LineWidth="1"><ofd:StrokeColor Value="#00 #00 #00"/></ofd:DrawParam><ofd:DrawParam ID="10" Relative="9" LineWidth="2"><ofd:StrokeColor Value="#ee #20 #25"/></ofd:DrawParam></ofd:DrawParams><ofd:MultiMedias><ofd:MultiMedia ID="5" Type="Image" Format="PNG"><ofd:MediaFile>blue.png</ofd:MediaFile></ofd:MultiMedia></ofd:MultiMedias><ofd:CompositeGraphicUnits><ofd:CompositeGraphicUnit ID="20" Width="10" Height="10"><ofd:Content><ofd:PathObject ID="21" Boundary="0 0 10 10" Fill="true" Stroke="false"><ofd:FillColor Value="0 80 220"/><ofd:AbbreviatedData>M 0 0 L 10 0 L 10 10 L 0 10 C</ofd:AbbreviatedData></ofd:PathObject></ofd:Content></ofd:CompositeGraphicUnit></ofd:CompositeGraphicUnits></ofd:Res>`),
    "Doc_0/Res/blue.png": png(2, 2),
    "Doc_0/Signs/Signatures.xml": encoder.encode(`<?xml version="1.0"?><ofd:Signatures xmlns:ofd="http://www.ofdspec.org/2016" MaxSignId="1"><ofd:Signature ID="1" Type="Seal" BaseLoc="Sign_0/Signature.xml"/></ofd:Signatures>`),
    "Doc_0/Signs/Sign_0/Signature.xml": encoder.encode(`<?xml version="1.0"?><ofd:Signature xmlns:ofd="http://www.ofdspec.org/2016"><ofd:SignedInfo><ofd:Provider ProviderName="fixture"/><ofd:SignatureMethod>1.2.156.10197.1.501</ofd:SignatureMethod><ofd:References CheckMethod="SM3"><ofd:Reference FileRef="payload.bin"><ofd:CheckValue>Zsfw9GLu7dnR8tRr3BDk4kFnxIdc8veiKX2gK49LqOA=</ofd:CheckValue></ofd:Reference></ofd:References><ofd:Seal><ofd:BaseLoc>Seal.esl</ofd:BaseLoc></ofd:Seal><ofd:StampAnnot ID="1" PageRef="1" Boundary="160 40 20 20"/></ofd:SignedInfo><ofd:SignedValue>SignedValue.dat</ofd:SignedValue></ofd:Signature>`),
    "Doc_0/Signs/Sign_0/payload.bin": encoder.encode(tamperSignatureReference ? "abd" : "abc"),
    "Doc_0/Signs/Sign_0/SignedValue.dat": Uint8Array.of(0x30, 0x00),
    "Doc_0/Signs/Sign_0/Seal.esl": sealContainer,
    "Doc_0/Tpls/Tpl_0/Content.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Page xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Content><ofd:Layer ID="8"><ofd:PathObject ID="7" Boundary="0 0 210 8" Fill="true" Stroke="false"><ofd:FillColor Value="230 240 250"/><ofd:AbbreviatedData>M 0 0 L 210 0 L 210 8 L 0 8 C</ofd:AbbreviatedData></ofd:PathObject></ofd:Layer></ofd:Content></ofd:Page>`),
    "Doc_0/Annots/Annotations.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Annotations xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Page PageID="1"><ofd:FileLoc>Page_0/Annot_0.xml</ofd:FileLoc></ofd:Page></ofd:Annotations>`),
    "Doc_0/Annots/Page_0/Annot_0.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:PageAnnot xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Annot ID="10" Type="Watermark"><ofd:Appearance Boundary="0 0 210 297"><ofd:TextObject ID="11" Boundary="150 275 30 5" Font="3" Size="3"><ofd:TextCode X="0" Y="3">批注</ofd:TextCode></ofd:TextObject></ofd:Appearance></ofd:Annot></ofd:PageAnnot>`),
    "Doc_0/Pages/Page_0/Content.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Page xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Area><ofd:PhysicalBox>0 0 210 297</ofd:PhysicalBox></ofd:Area><ofd:Template TemplateID="8" ZOrder="Background"/><ofd:Content><ofd:Layer ID="2"><ofd:TextObject ID="4" Boundary="31.7 25.4 40.5 5" Font="3" Size="3.0" Weight="700" Italic="true"><ofd:TextCode X="0" Y="3" DeltaX="3 3 3 3 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5 1.5">你好呀，OFD Reader&amp;Writer！</ofd:TextCode></ofd:TextObject><ofd:ImageObject ID="6" Boundary="31.7 40 20 10" ResourceID="5"/><ofd:PathObject ID="7" Boundary="60 40 20 10" Fill="true" Stroke="false"><ofd:FillColor Value="255 0 0"/><ofd:AbbreviatedData>M 0 0 L 20 0 L 20 10 L 0 10 C</ofd:AbbreviatedData></ofd:PathObject><ofd:PathObject ID="9" Boundary="60 55 20 2" DrawParam="10" Fill="false" Stroke="true"><ofd:Clips TransFlag="false"><ofd:Clip><ofd:Area><ofd:Path Boundary="0 0 20 2"><ofd:AbbreviatedData>M 0 0 L 20 0 L 20 2 L 0 2 C</ofd:AbbreviatedData></ofd:Path></ofd:Area></ofd:Clip></ofd:Clips><ofd:AbbreviatedData>S 0 1 L 20 1</ofd:AbbreviatedData></ofd:PathObject><ofd:CompositeObject ID="22" Boundary="85 40 10 10" CTM="1 0 0 1 85 40" ResourceID="20"/><ofd:PathObject ID="23" Boundary="105 40 40 20" Fill="true" Stroke="false"><ofd:FillColor><ofd:LaGouraudShd VerticesPerRow="2"><ofd:Point X="0" Y="0"><ofd:Color Value="255 0 255"/></ofd:Point><ofd:Point X="40" Y="0"><ofd:Color Value="0 255 255"/></ofd:Point><ofd:Point X="0" Y="20"><ofd:Color Value="255 255 0"/></ofd:Point><ofd:Point X="40" Y="20"><ofd:Color Value="0 255 0"/></ofd:Point></ofd:LaGouraudShd></ofd:FillColor><ofd:AbbreviatedData>M 0 0 L 40 0 L 40 20 L 0 20 C</ofd:AbbreviatedData></ofd:PathObject></ofd:Layer></ofd:Content></ofd:Page>`),
    "Doc_0/Pages/Page_1/Content.xml": encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>
      <ofd:Page xmlns:ofd="http://www.ofdspec.org/2016"><ofd:Content><ofd:Layer ID="12"><ofd:TextObject ID="13" Boundary="20 20 30 5" Font="3" Size="3"><ofd:TextCode X="0" Y="3">第二页</ofd:TextCode></ofd:TextObject></ofd:Layer></ofd:Content></ofd:Page>`),
  });
}

function createPdfFixture() {
  return encoder.encode(`%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj
4 0 obj << /Length 48 >> stream
BT /F1 12 Tf 1 0 0 1 72 720 Tm (Lazy PDF) Tj ET
endstream endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
trailer << /Root 1 0 R >>
%%EOF`);
}

function createTwoPagePdfFixture() {
  return encoder.encode(`%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 /MediaBox [0 0 612 792] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [
<< /Type /Annot /Subtype /Link /Rect [70 710 140 730] /A << /S /GoTo /D [6 0 R /Fit] >> >>
] >> endobj
4 0 obj << >> stream
BT /F1 12 Tf 1 0 0 1 72 720 Tm (First page) Tj ET
endstream endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
6 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >> endobj
7 0 obj << >> stream
BT /F1 12 Tf 1 0 0 1 72 720 Tm (Second page) Tj ET
endstream endobj
trailer << /Root 1 0 R >>
%%EOF`);
}

function pptRecord(options, recordType, payload) {
  const header = new Uint8Array(8);
  writeU16(header, 0, options);
  writeU16(header, 2, recordType);
  writeU32(header, 4, payload.length);
  return concat([header, payload]);
}

function directoryEntry(name, objectType, right, child, startSector, size) {
  const entry = new Uint8Array(128);
  for (let index = 0; index < name.length; index += 1) {
    writeU16(entry, index * 2, name.charCodeAt(index));
  }
  writeU16(entry, 64, (name.length + 1) * 2);
  entry[66] = objectType;
  entry[67] = 1;
  writeU32(entry, 68, FREE_SECTOR);
  writeU32(entry, 72, right);
  writeU32(entry, 76, child);
  writeU32(entry, 116, startSector);
  writeU32(entry, 120, size);
  writeU32(entry, 124, 0);
  return entry;
}

function compoundFile(streams) {
  assert.ok(streams.length > 0 && streams.length <= 3);
  assert.equal(
    streams.every(({ stream }) => stream.length === LEGACY_STREAM_BYTES),
    true,
  );
  const sectorsPerStream = LEGACY_STREAM_BYTES / SECTOR_BYTES;
  const sectorCount = 2 + streams.length * sectorsPerStream;
  const bytes = new Uint8Array((sectorCount + 1) * SECTOR_BYTES);
  bytes.set(Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1));
  writeU16(bytes, 24, 0x003e);
  writeU16(bytes, 26, 3);
  writeU16(bytes, 28, 0xfffe);
  writeU16(bytes, 30, 9);
  writeU16(bytes, 32, 6);
  writeU32(bytes, 40, 0);
  writeU32(bytes, 44, 1);
  writeU32(bytes, 48, 1);
  writeU32(bytes, 56, LEGACY_STREAM_BYTES);
  writeU32(bytes, 60, END_OF_CHAIN);
  writeU32(bytes, 64, 0);
  writeU32(bytes, 68, END_OF_CHAIN);
  writeU32(bytes, 72, 0);
  for (let offset = 76; offset < SECTOR_BYTES; offset += 4) {
    writeU32(bytes, offset, FREE_SECTOR);
  }
  writeU32(bytes, 76, 0);

  const fatOffset = SECTOR_BYTES;
  for (let offset = fatOffset; offset < fatOffset + SECTOR_BYTES; offset += 4) {
    writeU32(bytes, offset, FREE_SECTOR);
  }
  writeU32(bytes, fatOffset, FAT_SECTOR);
  writeU32(bytes, fatOffset + 4, END_OF_CHAIN);
  for (let streamIndex = 0; streamIndex < streams.length; streamIndex += 1) {
    const firstSector = 2 + streamIndex * sectorsPerStream;
    for (let index = 0; index < sectorsPerStream; index += 1) {
      const sector = firstSector + index;
      writeU32(
        bytes,
        fatOffset + sector * 4,
        index + 1 === sectorsPerStream ? END_OF_CHAIN : sector + 1,
      );
    }
  }

  const directoryOffset = 2 * SECTOR_BYTES;
  bytes.set(
    directoryEntry("Root Entry", 5, FREE_SECTOR, 1, END_OF_CHAIN, 0),
    directoryOffset,
  );
  streams.forEach(({ name, stream }, index) => {
    const firstSector = 2 + index * sectorsPerStream;
    const right = index + 1 < streams.length ? index + 2 : FREE_SECTOR;
    bytes.set(
      directoryEntry(
        name,
        2,
        right,
        FREE_SECTOR,
        firstSector,
        LEGACY_STREAM_BYTES,
      ),
      directoryOffset + (index + 1) * 128,
    );
    bytes.set(stream, (firstSector + 1) * SECTOR_BYTES);
  });
  return bytes;
}

function createLegacyPptFixture() {
  const dimensions = new Uint8Array(40);
  writeI32(dimensions, 0, 5760);
  writeI32(dimensions, 4, 4320);
  const documentAtom = pptRecord(1, 1001, dimensions);

  const slidePersist = new Uint8Array(20);
  writeU32(slidePersist, 0, 2);
  writeU32(slidePersist, 12, 0x100);
  const slideList = pptRecord(0x000f, 4080, pptRecord(0, 1011, slidePersist));
  const document = pptRecord(0x000f, 1000, concat([documentAtom, slideList]));

  const text = "Title\rBody";
  const textBytes = new Uint8Array(text.length * 2);
  for (let index = 0; index < text.length; index += 1) {
    writeU16(textBytes, index * 2, text.charCodeAt(index));
  }
  const slide = pptRecord(0x000f, 1006, pptRecord(0, 4000, textBytes));
  const persistOffset = document.length + slide.length;
  const persistPayload = new Uint8Array(12);
  writeU32(persistPayload, 0, 1 | (2 << 20));
  writeU32(persistPayload, 4, 0);
  writeU32(persistPayload, 8, document.length);
  const persist = pptRecord(0, 6002, persistPayload);

  const editOffset = persistOffset + persist.length;
  const editPayload = new Uint8Array(28);
  editPayload[7] = 3;
  writeU32(editPayload, 12, persistOffset);
  writeU32(editPayload, 16, 1);
  writeU32(editPayload, 20, 3);
  const edit = pptRecord(0, 4085, editPayload);
  const powerpointDocument = new Uint8Array(LEGACY_STREAM_BYTES);
  powerpointDocument.set(concat([document, slide, persist, edit]));

  const currentUserPayload = new Uint8Array(24);
  writeU32(currentUserPayload, 0, 0x14);
  writeU32(currentUserPayload, 4, 0xe391c05f);
  writeU32(currentUserPayload, 8, editOffset);
  writeU16(currentUserPayload, 14, 0x03f4);
  currentUserPayload[16] = 3;
  writeU32(currentUserPayload, 20, 8);
  const currentUser = new Uint8Array(LEGACY_STREAM_BYTES);
  currentUser.set(pptRecord(0, 4086, currentUserPayload));

  return compoundFile([
    { name: "PowerPoint Document", stream: powerpointDocument },
    { name: "Current User", stream: currentUser },
  ]);
}

function createLegacyDocFixture() {
  const word = new Uint8Array(LEGACY_STREAM_BYTES);
  writeU16(word, 0, 0xa5ec);
  writeU16(word, 2, 0x00c1);
  writeU16(word, 10, 1 << 2);
  writeU32(word, 24, 512);
  writeU32(word, 28, 523);
  writeU16(word, 32, 0);
  writeU16(word, 34, 11);
  writeU32(word, 36 + 3 * 4, 11);
  writeU16(word, 80, 34);
  writeU32(word, 82 + 12 * 8, 64);
  writeU32(word, 82 + 12 * 8 + 4, 12);
  writeU32(word, 82 + 15 * 8, 80);
  writeU32(word, 82 + 33 * 8, 0);
  writeU32(word, 82 + 33 * 8 + 4, 21);
  word.set(encoder.encode("BoldItalic\r"), 512);

  const chpxPage = 1024;
  writeU32(word, chpxPage, 512);
  writeU32(word, chpxPage + 4, 516);
  writeU32(word, chpxPage + 8, 522);
  writeU32(word, chpxPage + 12, 523);
  word[chpxPage + 16] = 200;
  word[chpxPage + 17] = 210;
  word[chpxPage + 18] = 0;
  word[chpxPage + 511] = 3;
  const bold = Uint8Array.of(
    0x35, 0x08, 0x01,
    0x43, 0x4a, 40, 0,
    0x70, 0x68, 0xff, 0x00, 0x00, 0x00,
    0x4f, 0x4a, 0x00, 0x00,
  );
  word[chpxPage + 400] = bold.length;
  word.set(bold, chpxPage + 401);
  const italic = Uint8Array.of(
    0x36, 0x08, 0x01,
    0x3e, 0x2a, 0x01,
    0x43, 0x4a, 28, 0,
    0x4f, 0x4a, 0x01, 0x00,
  );
  word[chpxPage + 420] = italic.length;
  word.set(italic, chpxPage + 421);

  const table = new Uint8Array(LEGACY_STREAM_BYTES);
  table[0] = 0x02;
  writeU32(table, 1, 16);
  writeU32(table, 5, 0);
  writeU32(table, 9, 11);
  writeU16(table, 13, 0);
  writeU32(table, 15, 0x4000_0000 | (512 * 2));
  writeU16(table, 19, 0);
  writeU32(table, 64, 512);
  writeU32(table, 68, 523);
  writeU32(table, 72, 2);

  const fontTable = [2, 0, 0, 0];
  for (const name of ["Arial", "Courier New"]) {
    const record = new Uint8Array(39 + (name.length + 1) * 2);
    for (let index = 0; index < name.length; index += 1) {
      writeU16(record, 39 + index * 2, name.charCodeAt(index));
    }
    fontTable.push(record.length, ...record);
  }
  writeU32(word, 82 + 15 * 8 + 4, fontTable.length);
  table.set(fontTable, 80);

  return compoundFile([
    { name: "WordDocument", stream: word },
    { name: "0Table", stream: table },
  ]);
}

function biffRecord(recordId, payload) {
  const header = new Uint8Array(4);
  writeU16(header, 0, recordId);
  writeU16(header, 2, payload.length);
  return concat([header, payload]);
}

function createLegacyEtFixture() {
  const workbookBofPayload = new Uint8Array(8);
  writeU16(workbookBofPayload, 0, 0x0600);
  writeU16(workbookBofPayload, 2, 0x0005);
  const workbookBof = biffRecord(0x0809, workbookBofPayload);
  const workbookEof = biffRecord(0x000a, new Uint8Array());

  const worksheetBofPayload = new Uint8Array(8);
  writeU16(worksheetBofPayload, 0, 0x0600);
  writeU16(worksheetBofPayload, 2, 0x0010);
  const worksheetBof = biffRecord(0x0809, worksheetBofPayload);
  const numberPayload = new Uint8Array(14);
  writeU16(numberPayload, 0, 0);
  writeU16(numberPayload, 2, 0);
  writeU16(numberPayload, 4, 0);
  view(numberPayload).setFloat64(6, 42, true);
  const number = biffRecord(0x0203, numberPayload);
  const worksheet = concat([worksheetBof, number, workbookEof]);

  const sheetName = "Sheet1";
  const boundSheetPayload = new Uint8Array(8 + sheetName.length);
  const provisionalBoundSheet = biffRecord(0x0085, boundSheetPayload);
  writeU32(
    boundSheetPayload,
    0,
    workbookBof.length + provisionalBoundSheet.length + workbookEof.length,
  );
  boundSheetPayload[6] = sheetName.length;
  boundSheetPayload[7] = 0;
  boundSheetPayload.set(encoder.encode(sheetName), 8);

  const workbook = new Uint8Array(LEGACY_STREAM_BYTES);
  workbook.set(concat([
    workbookBof,
    biffRecord(0x0085, boundSheetPayload),
    workbookEof,
    worksheet,
  ]));
  const etExtensionData = new Uint8Array(LEGACY_STREAM_BYTES);
  etExtensionData.set(encoder.encode("ETExtData"));
  return compoundFile([
    { name: "Workbook", stream: workbook },
    { name: "ETExtData", stream: etExtensionData },
  ]);
}

async function readArtifacts() {
  const [base, odf, iworkPdf, legacy, xps, ofd] = await Promise.all([
    readFile(new URL("../dist/office-viewer-core.wasm", import.meta.url)),
    readFile(new URL("../dist/office-viewer-odf.wasm", import.meta.url)),
    readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url)),
    readFile(new URL("../dist/office-viewer-legacy-office.wasm", import.meta.url)),
    readFile(new URL("../dist/office-viewer-xps.wasm", import.meta.url)),
    readFile(new URL("../dist/office-viewer-ofd.wasm", import.meta.url)),
  ]);
  return { base, odf, iworkPdf, legacy, pdf: iworkPdf, xps, ofd };
}

async function assertUnsupported(engine, bytes, label) {
  await assert.rejects(
    engine.open(bytes),
    (cause) => cause?.code === "UNSUPPORTED_FORMAT",
    `${label} must remain unclaimed`,
  );
}

const fixtures = {
  csv: {
    bytes: encoder.encode("name,value\r\nA,1\r\n"),
    async validate(document) {
      assert.equal(document.info.format, "csv");
      assert.equal(document.info.kind, "spreadsheet");
    },
  },
  odt: {
    bytes: createOdtFixture(),
    async validate(document) {
      assert.equal(document.info.format, "odt");
      assert.equal(document.info.kind, "text");
      const objects = await document.listObjects({ textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["Lazy ODF"]);
    },
  },
  pages: {
    bytes: createPagesFixture(),
    async validate(document) {
      assert.equal(document.info.format, "pages");
      assert.equal(document.info.kind, "text");
      assert.deepEqual(
        {
          type: document.info.units[0]?.type,
          width: document.info.units[0]?.width,
          height: document.info.units[0]?.height,
        },
        { type: "page", width: 3, height: 2 },
      );
      const objects = await document.listObjects();
      assert.equal(objects.length, 1);
      assert.equal(objects[0]?.type, "image");
      assert.deepEqual(objects[0]?.source, {
        format: "pages",
        part: "preview.png",
        kind: "preview",
        component: "preview.png",
        mapping: "exact",
      });
      assert.equal(
        document.diagnostics().some(({ message }) => message.startsWith("IWORK_PREVIEW_ONLY:")),
        true,
      );
    },
  },
  ppt: {
    bytes: createLegacyPptFixture(),
    async validate(document) {
      assert.equal(document.info.format, "ppt");
      assert.equal(document.info.kind, "presentation");
      assert.deepEqual(
        {
          type: document.info.units[0]?.type,
          width: document.info.units[0]?.width,
          height: document.info.units[0]?.height,
        },
        { type: "slide", width: 960, height: 720 },
      );
      const objects = await document.listObjects({ textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["Title\nBody"]);
      assert.equal(objects[0]?.source.format, "ppt");
      assert.equal(objects[0]?.source.kind, "text");
      assert.equal(objects[0]?.source.stream, "PowerPoint Document");
    },
  },
  doc: {
    bytes: createLegacyDocFixture(),
    async validate(document) {
      assert.equal(document.info.format, "doc");
      assert.equal(document.info.kind, "text");
      const objects = await document.listObjects({ textOnly: true });
      assert.equal(objects.length, 1);
      assert.equal(objects[0]?.text, "BoldItalic");
      assert.deepEqual(
        objects[0]?.fontRuns?.map(({ start, end, authoredFamily }) => ({
          start,
          end,
          authoredFamily,
        })),
        [
          { start: 0, end: 4, authoredFamily: "Arial" },
          { start: 4, end: 10, authoredFamily: "Courier New" },
        ],
      );
      assert.equal(objects[0]?.source.format, "doc");
      assert.equal(objects[0]?.source.kind, "paragraph");
      assert.equal(objects[0]?.source.stream, "WordDocument");
    },
  },
  et: {
    bytes: createLegacyEtFixture(),
    async validate(document) {
      assert.equal(document.info.format, "xls");
      assert.equal(document.info.kind, "spreadsheet");
      const objects = await document.listObjects({ textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["42"]);
      assert.equal(objects[0]?.source.format, "xls");
      assert.equal(objects[0]?.source.stream, "Workbook");
    },
  },
  pdf: {
    bytes: createPdfFixture(),
    async validate(document) {
      assert.equal(document.info.format, "pdf");
      assert.equal(document.info.kind, "text");
      assert.deepEqual(
        {
          type: document.info.units[0]?.type,
          width: document.info.units[0]?.width,
          height: document.info.units[0]?.height,
        },
        { type: "page", width: 816, height: 1056 },
      );
      const objects = await document.listObjects({ textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["Lazy PDF"]);
      assert.equal(objects[0]?.source.format, "pdf");
      assert.equal(objects[0]?.source.kind, "text");
    },
  },
  xps: {
    bytes: createXpsFixture(),
    async validate(document) {
      assert.equal(document.info.format, "xps");
      assert.equal(document.info.kind, "text");
      assert.deepEqual(
        {
          type: document.info.units[0]?.type,
          width: document.info.units[0]?.width,
          height: document.info.units[0]?.height,
        },
        { type: "page", width: 320, height: 240 },
      );
      assert.equal(document.info.units.length, 2);
      const objects = await document.listObjects({ unitIndex: 0, textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["Lazy XPS"]);
      assert.equal(objects[0]?.source.format, "xps");
      assert.equal(objects[0]?.source.kind, "glyphs");
      const secondPage = await document.listObjects({ unitIndex: 1, textOnly: true });
      assert.deepEqual(secondPage.map(({ text }) => text), ["Second XPS"]);
    },
  },
  ofd: {
    bytes: createOfdFixture(),
    async validate(document) {
      assert.equal(document.info.format, "ofd");
      assert.equal(document.info.kind, "text");
      const page = document.info.units[0];
      assert.equal(document.info.units.length, 2);
      assert.equal(page?.type, "page");
      assert.deepEqual(document.info.outline, [{ title: "首页", unitIndex: 0, level: 0 }]);
      assert.ok(Math.abs(page.width - 793.7008) < 0.001);
      assert.ok(Math.abs(page.height - 1122.5197) < 0.001);
      const objects = await document.listObjects({ unitIndex: 0, textOnly: true });
      assert.deepEqual(objects.map(({ text }) => text), ["你好呀，OFD Reader&Writer！", "批注"]);
      const secondPage = await document.listObjects({ unitIndex: 1, textOnly: true });
      assert.deepEqual(secondPage.map(({ text }) => text), ["第二页"]);
      assert.match(secondPage[0]?.id ?? "", /Doc_0\/Pages\/Page_1\/Content\.xml/u);
      assert.equal(objects[0]?.source.format, "ofd");
      assert.equal(objects[0]?.source.kind, "text");
      assert.equal(objects[1]?.source.part, "Doc_0/Annots/Page_0/Annot_0.xml");
      const allObjects = await document.listObjects();
      const image = allObjects.find(({ source }) => source.kind === "image");
      assert.equal(image?.type, "image", JSON.stringify(await document.diagnostics()));
      assert.equal(image?.source.format, "ofd");
      assert.equal(
        document.diagnostics().some(({ message }) => message.startsWith("OFD_SIGNATURE_VALUE_UNVERIFIABLE:")),
        true,
      );
      const paths = allObjects.filter(({ source }) => source.kind === "path");
      assert.equal(paths.length, 5);
      assert.equal(paths[0]?.source.part, "Doc_0/Tpls/Tpl_0/Content.xml");
      const composite = paths.find(({ id }) => /ofd-composite-22/u.test(id));
      assert.equal(composite?.source.part, "Doc_0/DocumentRes.xml");
      assert.equal(allObjects.some(({ source }) => source.kind === "seal"), true);
    },
  },
};

test("dist contains no conversion-based format extension artifact", async () => {
  const files = await readdir(new URL("../dist/", import.meta.url));
  assert.deepEqual(files.filter((name) => name.startsWith("format-extension.")), []);
});

test("dist Wasm artifacts contain only their assigned format families", async () => {
  const artifacts = await readArtifacts();
  const engines = {};
  const expected = {
    base: { csv: true, odt: false, pages: false, ppt: false, doc: false, et: false, pdf: false, xps: false, ofd: false },
    odf: { csv: false, odt: true, pages: false, ppt: false, doc: false, et: false, pdf: false, xps: false, ofd: false },
    iworkPdf: { csv: false, odt: false, pages: true, ppt: false, doc: false, et: false, pdf: true, xps: false, ofd: false },
    legacy: { csv: false, odt: false, pages: false, ppt: true, doc: true, et: true, pdf: false, xps: false, ofd: false },
    xps: { csv: false, odt: false, pages: false, ppt: false, doc: false, et: false, pdf: false, xps: true, ofd: false },
    ofd: { csv: false, odt: false, pages: false, ppt: false, doc: false, et: false, pdf: false, xps: false, ofd: true },
  };

  try {
    for (const name of Object.keys(expected)) {
      engines[name] = await createOfficeEngine({ wasm: artifacts[name], execution: "inline" });
    }
    for (const [engineName, engine] of Object.entries(engines)) {
      for (const [fixtureName, fixture] of Object.entries(fixtures)) {
        const label = `${engineName} Wasm for ${fixtureName}`;
        if (!expected[engineName][fixtureName]) {
          await assertUnsupported(engine, fixture.bytes, label);
          continue;
        }
        const document = await engine.open(fixture.bytes);
        try {
          await fixture.validate(document);
        } finally {
          document.close();
        }
      }
    }
    await assert.rejects(
      engines.base.open(createPagesFixture({ redundantLocalZip64: true })),
      (cause) => cause?.code === "UNSUPPORTED_FORMAT",
    );
  } finally {
    for (const engine of Object.values(engines)) engine.close();
  }
});

test("supplied multipage OFD opens substantially faster than materializing every page", async () => {
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  const bytes = await readFile(new URL("./fixtures/ofdrw-intro.ofd", import.meta.url));
  const ratios = [];
  try {
    for (let sample = 0; sample < 4; sample += 1) {
      const start = performance.now();
      const document = await engine.open(bytes);
      const openMs = performance.now() - start;
      try {
        assert.equal((await document.listObjects({ unitIndex: 0 })).length, 35);
        const queryStart = performance.now();
        const objects = await document.listObjects();
        const queryMs = performance.now() - queryStart;
        assert.equal(objects.length, 1804);
        // Compare the same real document on the same machine; discard Wasm warmup.
        if (sample > 0) ratios.push(openMs / queryMs);
      } finally {
        document.close();
      }
    }
    ratios.sort((a, b) => a - b);
    assert.ok(ratios[1] < 0.6, `first-page open / full materialization: ${ratios[1]}`);
  } finally {
    engine.close();
  }
});

test("supplied OFD Pattern renders watermark text through the shared visual brush", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/ofdrw-pattern.ofd", import.meta.url)));
    frame = await document.render({ unitIndex: 0 });
    const text = recording.calls.text.map(({ text }) => text).join("");
    assert.match(text, /严禁复制/u);
    assert.equal(recording.calls.fills.some((paint) => typeof paint === "object"), true);
    assert.deepEqual(frame.diagnostics, []);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("OFD 1.1 drawing parameters survive the public protocol and renderer", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(fixtures.ofd.bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(
      recording.calls.strokes.some(({ strokeStyle, lineWidth }) => (
        strokeStyle === "rgba(238, 32, 37, 1)"
        && lineWidth > 0
      )),
      true,
      JSON.stringify(recording.calls.strokes),
    );
    assert.equal(recording.calls.clips > 0, true);
    assert.equal(recording.calls.images > 0, true);
    assert.equal(new Set(recording.calls.fills).size > 10, true, JSON.stringify(recording.calls.fills));
    assert.match(recording.calls.text.find(({ text }) => text === "你")?.font ?? "", /^italic 700 /u);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("supplied OFDRW text uses authored origins and baselines without text-box padding", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/ofdrw-testImageNotFound.ofd", import.meta.url)));
    const objects = await document.listObjects({ unitIndex: 1 });
    const text = objects.find(({ text }) => text === "根据国家法律、法规规定，经审查");
    assert.ok(text);
    const location = objects.find(({ text }) => text === "地里位置：");
    assert.ok(location);
    const hits = await document.hitTest({
      unitIndex: 1,
      x: location.bounds.x + location.bounds.width / 3,
      y: location.bounds.y + location.bounds.height / 2,
      limit: 16,
    });
    assert.equal(hits[0]?.object.text, "地里位置：");
    assert.ok(hits.every(({ object }) => object.text !== "合格，授予探矿权，特发此证。"));
    frame = await document.render({ unitIndex: 1 });
    const first = recording.calls.text.find(({ text }) => text === "根");
    assert.ok(first);
    // Recorded coordinates precede the object's CTM. X/Y are local millimetres.
    assert.ok(Math.abs(first.x - 29.92 * 96 / 25.4) < 0.001, `x=${first.x}`);
    assert.ok(Math.abs(first.y - (24.09 + 0.79) * 96 / 25.4) < 0.001, `y=${first.y}`);
    const second = recording.calls.text.find(({ text }) => text === "据");
    assert.ok(Math.abs(second.x - first.x - 96 / 25.4) < 0.001);
    for (const [value, x, y] of [
      ["T32132132132132132", 64.57, 55.28 + 6.14],
      ["test123", 62.72, 68.24 + 7.36],
      ["云南省临沧市镇康县国土资源局南伞分局租房", 62.72, 81.2 + 6.86],
      ["探矿权新立", 62.72, 94.16 + 7.36],
      ["云南省镇康县", 62.98, 107.65 + 7.36],
      ["1.51", 63.51, 134.63 + 7.36],
      ["2022/11/01", 62.98 + 5.65, 148.12 + 7.36],
    ]) {
      const index = recording.calls.text.findIndex((_, start) => (
        recording.calls.text.slice(start, start + value.length).map(({ text }) => text).join("") === value
      ));
      assert.ok(index >= 0, `missing field: ${value}`);
      const draw = recording.calls.text[index];
      assert.ok(Math.abs(draw.x - x * 96 / 25.4) < 0.001, `${value}: x=${draw.x}`);
      assert.ok(Math.abs(draw.y - y * 96 / 25.4) < 0.001, `${value}: baseline=${draw.y}`);
    }
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("real OFD essay exposes positioned characters for selection", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/ofdrw-nalaizhuyi.ofd", import.meta.url)));
    frame = await document.render({ unitIndex: 0, includeTextFragments: true });
    assert.ok(frame.textFragments.length > 100, "OFD must expose painted glyph positions, not paragraph bounds");
    const objects = await document.listObjects({ unitIndex: 0 });
    const paragraph = objects.find(({ text }) => text?.startsWith("中国"));
    assert.ok(paragraph);
    const fragments = frame.textFragments.filter(({ objectId }) => objectId === paragraph.id);
    assert.equal(fragments.map(({ text }) => text).join(""), paragraph.text);
    let offset = 0;
    for (const fragment of fragments) {
      assert.equal(fragment.start, offset);
      offset += fragment.text.length;
      assert.equal(fragment.end, offset);
      assert.ok(fragment.width > 0 && fragment.height > 0);
    }
    assert.ok(fragments[1].transform.e > fragments[0].transform.e);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("supplied OFDRW signout table paths remain visible inside their page boundaries", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(await readFile(new URL("./fixtures/ofdrw-signout.ofd", import.meta.url)));
    const paths = (await document.listObjects({ unitIndex: 0 }))
      .filter(({ source }) => source.kind === "path" && source.part.endsWith("Content.xml"));
    assert.equal(paths.length, 250);
    frame = await document.render({ unitIndex: 0 });
    // The source has 249 filled paths, including the thin table borders.
    assert.equal(recording.calls.fills.length, 249);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("OFD SM3 reference verification reports tampering through the public API", async () => {
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  try {
    document = await engine.open(createOfdFixture({ tamperSignatureReference: true }));
    assert.equal(
      document.diagnostics().some(({ message }) => message.startsWith("OFD_SIGNATURE_REFERENCE_INVALID:")),
      true,
    );
  } finally {
    document?.close();
    engine.close();
  }
});

test("supplied OFDRW raster seals survive the public protocol and Canvas renderer", async () => {
  const recording = installRecordingCanvas();
  const { ofd } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: ofd, execution: "inline" });
  let document;
  try {
    document = await engine.open(suppliedOfdrwSeal);
    const seals = (await document.listObjects()).filter(({ source }) => source.kind === "seal");
    assert.equal(seals.length, 6);
    assert.equal(seals.filter(({ unitIndex }) => unitIndex === 0).length, 2);
    const renderDiagnostics = [];
    for (let unitIndex = 0; unitIndex < 5; unitIndex += 1) {
      const frame = await document.render({ unitIndex });
      renderDiagnostics.push(...frame.diagnostics);
      frame.bitmap.close();
    }
    assert.equal(recording.calls.clips >= 5, true);
    assert.equal(recording.calls.images >= 6, true, JSON.stringify(renderDiagnostics));
  } finally {
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("binary DOC styles survive the public protocol and renderer", async () => {
  const recording = installRecordingCanvas();
  const { legacy } = await readArtifacts();
  const engine = await createOfficeEngine({
    wasm: legacy,
    execution: "inline",
    fontPolicy: "local-first",
  });
  let document;
  let frame;
  try {
    document = await engine.open(fixtures.doc.bytes);
    frame = await document.render({ unitIndex: 0 });

    const bold = recording.calls.text.find(({ text }) => text === "Bold");
    const italic = recording.calls.text.find(({ text }) => text === "Italic");
    assert.ok(bold);
    assert.ok(italic);
    assert.match(bold.font, /^700 /);
    assert.match(bold.font, /px "Arial"(?:,|$)/);
    assert.ok(
      Math.abs(Number.parseFloat(bold.font.match(/ ([\d.]+)px /)?.[1]) - 20 * 96 / 72) < 0.01,
    );
    assert.equal(bold.fillStyle, "rgba(255, 0, 0, 1)");
    assert.match(italic.font, /^italic 400 /);
    assert.match(italic.font, /px monospace(?:,|$)/);
    assert.ok(
      Math.abs(Number.parseFloat(italic.font.match(/ ([\d.]+)px /)?.[1]) - 14 * 96 / 72) < 0.01,
    );
    assert.equal(italic.fillStyle, "rgba(0, 0, 0, 1)");
    assert.equal(
      recording.calls.strokes.some(({ strokeStyle }) => strokeStyle === "rgba(0, 0, 0, 1)"),
      true,
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("supplied Pages TOC preserves both saved pages and page-number fields", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  try {
    const document = await engine.open(await readFile(new URL("./fixtures/word-original.pages", import.meta.url)));
    const first = (await document.listObjects({ unitIndex: 2 })).filter((object) => object.id.includes("-toc-"));
    const next = (await document.listObjects({ unitIndex: 3 })).filter((object) => object.id.includes("-toc-"));
    assert.equal(first.length, 43);
    assert.equal(next.length, 18);
    assert.equal(first[0].text, "词汇表\t5\n");
    assert.equal(next[0].text, "地表水\t22\n");
    assert.equal(next.at(-1).text, "Bibliography\t30");
    assert.ok([...first, ...next].every((object) => object.bounds.y + object.bounds.height < 1122.67));
  } finally {
    engine.close();
  }
});

test("supplied Pages cell images survive the public protocol", async () => {
  const wasm = await readFile(new URL("../dist/office-viewer-pdf.wasm", import.meta.url));
  const engine = await createOfficeEngine({ wasm, execution: "inline" });
  try {
    const document = await engine.open(await readFile(new URL("./fixtures/word-original.pages", import.meta.url)));
    const objects = await document.listObjects({ unitIndex: 0 });
    const images = objects.filter((object) => object.type === "image");
    assert.equal(images.length, 2);
    assert.ok(images.every((image) => image.source.format === "pages" && image.source.kind === "inline-image"));
    const table = objects.find((object) => object.type === "table");
    assert.ok(images.every((image) => image.bounds.y + image.bounds.height <= table.bounds.y + table.bounds.height));
  } finally {
    engine.close();
  }
});

test("PDF text survives the public protocol and renderer", async () => {
  const recording = installRecordingCanvas();
  const { pdf } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: pdf, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(fixtures.pdf.bytes);
    frame = await document.render({ unitIndex: 0 });
    assert.equal(recording.calls.text.map(({ text }) => text).join("").includes("Lazy PDF"), true);
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("password-protected PDF reports actionable failures and opens with the password", async () => {
  const recording = installRecordingCanvas();
  const { pdf } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: pdf, execution: "inline" });
  const bytes = encryptedPdfFixture();
  let document;
  let frame;
  try {
    await assert.rejects(engine.open(bytes), (error) => error?.code === "PDF_PASSWORD_REQUIRED");
    await assert.rejects(
      engine.open(bytes, { password: "wrong" }),
      (error) => error?.code === "PDF_PASSWORD_INCORRECT",
    );
    document = await engine.open(bytes, { password: "secret" });
    const textCallStart = recording.calls.text.length;
    frame = await document.render({ unitIndex: 0 });
    assert.equal(
      recording.calls.text.slice(textCallStart).map(({ text }) => text).join(""),
      "Protected PDF",
    );
    frame.bitmap.close();
    document.close();
    frame = undefined;
    document = await engine.open(encryptedPdfFixture(""));
    assert.equal(document.info.format, "pdf");
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("password-protected PDF decrypts AES-128 and AES-256 revision 5 streams", async () => {
  const recording = installRecordingCanvas();
  const { pdf } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: pdf, execution: "inline" });
  try {
    for (const [bytes, expected] of [
      [encryptedPdfAes128Fixture(), "AES 128 PDF"],
      [encryptedPdfAes256R5Fixture(), "AES 256 PDF"],
    ]) {
      const document = await engine.open(bytes, { password: "secret" });
      const textCallStart = recording.calls.text.length;
      const frame = await document.render({ unitIndex: 0 });
      frame.bitmap.close();
      document.close();
      assert.equal(
        recording.calls.text.slice(textCallStart).map(({ text }) => text).join(""),
        expected,
      );
    }
  } finally {
    engine.close();
    recording.restore();
  }
});

test("PDF pages materialize on demand and remain queryable", async () => {
  const recording = installRecordingCanvas();
  const { pdf } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: pdf, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(createTwoPagePdfFixture());
    assert.equal(document.info.units.length, 2);
    assert.deepEqual(
      (await document.listObjects({ unitIndex: 0, textOnly: true })).map(({ text }) => text),
      ["First page"],
    );
    assert.deepEqual(
      (await document.listObjects({ unitIndex: 0, types: ["shape"] }))
        .flatMap(({ actions }) => actions ?? []),
      [{
        trigger: "click",
        kind: "command",
        action: "navigate",
        target: "1",
      }],
    );
    frame = await document.render({ unitIndex: 1 });
    assert.deepEqual(
      (await document.listObjects({ unitIndex: 1, textOnly: true })).map(({ text }) => text),
      ["Second page"],
    );
    assert.deepEqual(
      (await document.listObjects({ textOnly: true })).map(({ text }) => text),
      ["First page", "Second page"],
    );
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});

test("createOfficeEngine lazily opens all real dist format packs", async () => {
  const { base } = await readArtifacts();
  const loads = [];
  let factoryCalls = 0;
  const engine = await createOfficeEngine({
    wasm: base,
    execution: "inline",
    formatPack: async () => {
      factoryCalls += 1;
      return {
        async load(candidate) {
          const source = await extendedFormatPack.load(candidate);
          assert.equal(source instanceof URL, true);
          loads.push([candidate, source.pathname.split("/").at(-1)]);
          // Node fetch does not load file: URLs; the browser-facing pack still resolves
          // the real published asset, which this test supplies as the same bytes.
          return readFile(source);
        },
      };
    },
  });

  try {
    assert.equal(factoryCalls, 0);
    const csv = await engine.open(fixtures.csv.bytes);
    csv.close();
    assert.equal(factoryCalls, 0);
    assert.deepEqual(loads, []);

    const odt = await engine.open(fixtures.odt.bytes);
    try {
      await fixtures.odt.validate(odt);
    } finally {
      odt.close();
    }
    assert.equal(factoryCalls, 1);
    assert.deepEqual(loads, [["odf", "office-viewer-odf.wasm"]]);

    const pages = await engine.open(fixtures.pages.bytes);
    try {
      await fixtures.pages.validate(pages);
    } finally {
      pages.close();
    }
    assert.equal(factoryCalls, 1);
    assert.deepEqual(loads, [
      ["odf", "office-viewer-odf.wasm"],
      ["iwork", "office-viewer-pdf.wasm"],
    ]);

    const applePages = await engine.open(createPagesFixture({ redundantLocalZip64: true }));
    try {
      await fixtures.pages.validate(applePages);
    } finally {
      applePages.close();
    }
    assert.equal(loads.length, 2);

    await assert.rejects(
      engine.open(createPagesDirectoryPackageFixture()),
      (cause) => cause?.code === "UNSUPPORTED_FEATURE"
        && cause.message.includes("IWORK_DIRECTORY_PACKAGE_UNSUPPORTED"),
    );
    assert.equal(loads.length, 2);

    const ppt = await engine.open(fixtures.ppt.bytes);
    try {
      await fixtures.ppt.validate(ppt);
    } finally {
      ppt.close();
    }
    assert.equal(factoryCalls, 1);
    assert.deepEqual(loads, [
      ["odf", "office-viewer-odf.wasm"],
      ["iwork", "office-viewer-pdf.wasm"],
      ["legacy-office", "office-viewer-legacy-office.wasm"],
    ]);

    const doc = await engine.open(fixtures.doc.bytes);
    try {
      await fixtures.doc.validate(doc);
    } finally {
      doc.close();
    }
    assert.equal(loads.length, 3);

    const wps = await engine.open(fixtures.doc.bytes, { fileName: "report.wps" });
    try {
      await fixtures.doc.validate(wps);
    } finally {
      wps.close();
    }
    const et = await engine.open(fixtures.et.bytes, { fileName: "budget.et" });
    try {
      await fixtures.et.validate(et);
    } finally {
      et.close();
    }
    assert.deepEqual(loads.slice(-1), [["wps", "office-viewer-legacy-office.wasm"]]);

    const pdf = await engine.open(fixtures.pdf.bytes);
    try {
      await fixtures.pdf.validate(pdf);
    } finally {
      pdf.close();
    }
    assert.deepEqual(loads, [
      ["odf", "office-viewer-odf.wasm"],
      ["iwork", "office-viewer-pdf.wasm"],
      ["legacy-office", "office-viewer-legacy-office.wasm"],
      ["wps", "office-viewer-legacy-office.wasm"],
      ["pdf", "office-viewer-pdf.wasm"],
    ]);

    const xps = await engine.open(fixtures.xps.bytes);
    try {
      await fixtures.xps.validate(xps);
    } finally {
      xps.close();
    }
    assert.deepEqual(loads.at(-1), ["xps", "office-viewer-xps.wasm"]);

    const ofd = await engine.open(fixtures.ofd.bytes);
    try {
      await fixtures.ofd.validate(ofd);
    } finally {
      ofd.close();
    }
    assert.deepEqual(loads.at(-1), ["ofd", "office-viewer-ofd.wasm"]);

    const pagesAgain = await engine.open(fixtures.pages.bytes);
    pagesAgain.close();
    assert.equal(factoryCalls, 1);
    assert.equal(loads.length, 7);
  } finally {
    engine.close();
  }
});

test("XPS brush transforms do not multiply the offscreen surface twice", async () => {
  const recording = installRecordingCanvas();
  const { xps } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: xps, execution: "inline" });
  let document;
  let frame;
  try {
    document = await engine.open(fixtures.xps.bytes);
    frame = await document.render({ unitIndex: 0 });
  } finally {
    frame?.bitmap.close();
    document?.close();
    engine.close();
    recording.restore();
  }
});


test("XPS interleaved OPC parts require consecutive pieces and exactly one final piece", async () => {
  const { xps } = await readArtifacts();
  const engine = await createOfficeEngine({ wasm: xps, execution: "inline" });
  const entries = {
    "[Content_Types].xml": encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    "_rels/.rels": encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" Target="seq.fdseq"/></Relationships>'),
    "seq.fdseq": encoder.encode('<FixedDocumentSequence><DocumentReference Source="doc.fdoc"/></FixedDocumentSequence>'),
    "doc.fdoc": encoder.encode('<FixedDocument><PageContent Source="page.fpage"/></FixedDocument>'),
    "page.fpage/[0].piece": encoder.encode('<FixedPage Width="100" Height="100">'),
    "page.fpage/[1].last.piece": encoder.encode('<Path Data="M0,0 L20,0 20,20 Z" Fill="#000000"/></FixedPage>'),
  };
  try {
    const document = await engine.open(createZip(entries));
    assert.equal(document.info.units.length, 1);
    document.close();
    const gap = { ...entries, "page.fpage/[2].last.piece": entries["page.fpage/[1].last.piece"] };
    delete gap["page.fpage/[1].last.piece"];
    await assert.rejects(engine.open(createZip(gap)), /piece sequence/);
    const afterFinal = { ...entries, "page.fpage/[2].piece": encoder.encode("extra") };
    await assert.rejects(engine.open(createZip(afterFinal)), /piece sequence/);
  } finally { engine.close(); }
});
