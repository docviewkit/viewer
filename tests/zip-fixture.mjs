import { deflateRawSync } from "node:zlib";

const encoder = new TextEncoder();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function u32(value) {
  return Uint8Array.of(
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  );
}

function concat(parts) {
  const length = parts.reduce((sum, part) => sum + part.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

export function createZip(entries, { compress = true, redundantLocalZip64 = false } = {}) {
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const [name, value] of Object.entries(entries)) {
    const nameBytes = encoder.encode(name);
    const bytes = typeof value === "string" ? encoder.encode(value) : value;
    const compressed = compress ? deflateRawSync(bytes) : bytes;
    const method = compress ? 8 : 0;
    const checksum = crc32(bytes);
    const localExtra = redundantLocalZip64
      ? concat([u16(0x0001), u16(16), u32(bytes.length), u32(0), u32(compressed.length), u32(0)])
      : new Uint8Array();
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0x0800),
      u16(method),
      u16(0),
      u16(0),
      u32(checksum),
      u32(compressed.length),
      u32(bytes.length),
      u16(nameBytes.length),
      u16(localExtra.length),
      nameBytes,
      localExtra,
      compressed,
    ]);
    localParts.push(local);

    centralParts.push(
      concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0x0800),
        u16(method),
        u16(0),
        u16(0),
        u32(checksum),
        u32(compressed.length),
        u32(bytes.length),
        u16(nameBytes.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(localOffset),
        nameBytes,
      ]),
    );
    localOffset += local.length;
  }

  const central = concat(centralParts);
  return concat([
    ...localParts,
    central,
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(centralParts.length),
    u16(centralParts.length),
    u32(central.length),
    u32(localOffset),
    u16(0),
  ]);
}
