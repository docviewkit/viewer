import { deflateRawSync, inflateRawSync } from "node:zlib";

const decoder = new TextDecoder();
const encoder = new TextEncoder();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function u32(value) {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff);
}

function concat(parts) {
  const output = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function view(bytes) {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

function findEndOfCentralDirectory(bytes) {
  const data = view(bytes);
  const start = Math.max(0, bytes.length - 65_557);
  for (let offset = bytes.length - 22; offset >= start; offset -= 1) {
    if (data.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error("ZIP end-of-central-directory record not found");
}

/** Read ordinary single-disk, non-ZIP64 OOXML/ODF entries using central-directory sizes. */
export function readZipEntries(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const data = view(bytes);
  const eocd = findEndOfCentralDirectory(bytes);
  if (data.getUint16(eocd + 4, true) !== 0 || data.getUint16(eocd + 6, true) !== 0) {
    throw new Error("Multi-disk ZIP packages are unsupported by the metamorphic harness");
  }
  const entryCount = data.getUint16(eocd + 10, true);
  let centralOffset = data.getUint32(eocd + 16, true);
  const entries = [];
  for (let index = 0; index < entryCount; index += 1) {
    if (data.getUint32(centralOffset, true) !== 0x02014b50) throw new Error("Invalid ZIP central-directory entry");
    const flags = data.getUint16(centralOffset + 8, true);
    const method = data.getUint16(centralOffset + 10, true);
    const checksum = data.getUint32(centralOffset + 16, true);
    const compressedSize = data.getUint32(centralOffset + 20, true);
    const uncompressedSize = data.getUint32(centralOffset + 24, true);
    const nameLength = data.getUint16(centralOffset + 28, true);
    const extraLength = data.getUint16(centralOffset + 30, true);
    const commentLength = data.getUint16(centralOffset + 32, true);
    const localOffset = data.getUint32(centralOffset + 42, true);
    if ((flags & 1) !== 0) throw new Error("Encrypted ZIP entries are unsupported by the metamorphic harness");
    if (method !== 0 && method !== 8) throw new Error(`ZIP compression method ${method} is unsupported by the metamorphic harness`);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error("ZIP64 packages are unsupported by the metamorphic harness");
    }
    const name = decoder.decode(bytes.subarray(centralOffset + 46, centralOffset + 46 + nameLength));
    if (data.getUint32(localOffset, true) !== 0x04034b50) throw new Error(`Invalid local ZIP entry for ${name}`);
    const localNameLength = data.getUint16(localOffset + 26, true);
    const localExtraLength = data.getUint16(localOffset + 28, true);
    const payloadOffset = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.subarray(payloadOffset, payloadOffset + compressedSize);
    const inflated = method === 0 ? compressed.slice() : new Uint8Array(inflateRawSync(compressed));
    if (inflated.length !== uncompressedSize || crc32(inflated) !== checksum) throw new Error(`ZIP integrity check failed for ${name}`);
    entries.push({ name, data: inflated, compression: method === 0 ? "store" : "deflate" });
    centralOffset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

export function readZipComment(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const data = view(bytes);
  const eocd = findEndOfCentralDirectory(bytes);
  const length = data.getUint16(eocd + 20, true);
  return decoder.decode(bytes.subarray(eocd + 22, eocd + 22 + length));
}

export function writeZipEntries(entries, { comment = "" } = {}) {
  if (entries.length > 0xffff) throw new Error("Too many ZIP entries for the non-ZIP64 metamorphic harness");
  const commentBytes = encoder.encode(comment);
  if (commentBytes.length > 0xffff) throw new Error("ZIP comment exceeds the non-ZIP64 limit");
  const names = new Set();
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;
  for (const entry of entries) {
    if (names.has(entry.name)) throw new Error(`Duplicate ZIP entry ${entry.name}`);
    names.add(entry.name);
    const name = encoder.encode(entry.name);
    const bytes = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
    const compressed = entry.compression === "store" ? bytes : new Uint8Array(deflateRawSync(bytes));
    if (name.length > 0xffff || bytes.length > 0xffffffff || compressed.length > 0xffffffff) {
      throw new Error(`ZIP entry ${entry.name} exceeds non-ZIP64 limits`);
    }
    const method = entry.compression === "store" ? 0 : 8;
    const checksum = crc32(bytes);
    const flags = 0x0800;
    const local = concat([
      u32(0x04034b50), u16(20), u16(flags), u16(method), u16(0), u16(0),
      u32(checksum), u32(compressed.length), u32(bytes.length), u16(name.length), u16(0), name, compressed,
    ]);
    localParts.push(local);
    centralParts.push(concat([
      u32(0x02014b50), u16(20), u16(20), u16(flags), u16(method), u16(0), u16(0),
      u32(checksum), u32(compressed.length), u32(bytes.length), u16(name.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(localOffset), name,
    ]));
    localOffset += local.length;
  }
  const central = concat(centralParts);
  return concat([
    ...localParts,
    central,
    u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length),
    u32(central.length), u32(localOffset), u16(commentBytes.length), commentBytes,
  ]);
}

function findTagEnd(xml, start) {
  let quote = "";
  for (let index = start + 1; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote !== "") {
      if (character === quote) quote = "";
    } else if (character === "\"" || character === "'") {
      quote = character;
    } else if (character === ">") {
      return index;
    }
  }
  return -1;
}

function mapTags(xml, transform) {
  let output = "";
  let cursor = 0;
  while (cursor < xml.length) {
    const start = xml.indexOf("<", cursor);
    if (start < 0) return output + xml.slice(cursor);
    output += xml.slice(cursor, start);
    if (xml.startsWith("<!--", start)) {
      const end = xml.indexOf("-->", start + 4);
      if (end < 0) return output + xml.slice(start);
      output += xml.slice(start, end + 3);
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", start)) {
      const end = xml.indexOf("]]>", start + 9);
      if (end < 0) return output + xml.slice(start);
      output += xml.slice(start, end + 3);
      cursor = end + 3;
      continue;
    }
    const end = findTagEnd(xml, start);
    if (end < 0) return output + xml.slice(start);
    const tag = xml.slice(start, end + 1);
    output += tag.startsWith("<?") || tag.startsWith("<!") ? tag : transform(tag);
    cursor = end + 1;
  }
  return output;
}

function reorderAttributes(xml) {
  return mapTags(xml, (tag) => {
    if (tag.startsWith("</")) return tag;
    const selfClosing = /\/\s*>$/u.test(tag);
    const body = tag.slice(1, selfClosing ? tag.lastIndexOf("/") : -1).trim();
    const nameMatch = /^([^\s/>]+)/u.exec(body);
    if (nameMatch === null) return tag;
    const name = nameMatch[1];
    const attributesText = body.slice(name.length);
    const attributes = [];
    const matcher = /\s+([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/guy;
    let consumed = 0;
    while (matcher.lastIndex < attributesText.length) {
      const match = matcher.exec(attributesText);
      if (match === null) break;
      attributes.push(`${match[1]}=${match[2]}`);
      consumed = matcher.lastIndex;
    }
    if (attributes.length < 2 || attributesText.slice(consumed).trim() !== "") return tag;
    return `<${name} ${attributes.reverse().join(" ")}${selfClosing ? "/" : ""}>`;
  });
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

function renameNamespacePrefix(xml) {
  const declarations = [...xml.matchAll(/xmlns:([A-Za-z_][\w.-]*)\s*=/gu)];
  const declaration = declarations.find((match) => {
    const prefix = match[1];
    if (prefix === "xml") return false;
    const inDoubleQuotedValue = new RegExp(`="[^"]*\\b${escapeRegExp(prefix)}:[^"]*"`, "u");
    const inSingleQuotedValue = new RegExp(`='[^']*\\b${escapeRegExp(prefix)}:[^']*'`, "u");
    return !inDoubleQuotedValue.test(xml) && !inSingleQuotedValue.test(xml);
  });
  if (declaration === undefined) return xml;
  const original = declaration[1];
  let replacement = "ovm";
  let suffix = 0;
  while (new RegExp(`xmlns:${escapeRegExp(replacement)}\\s*=`, "u").test(xml)) replacement = `ovm${++suffix}`;
  const elementToken = new RegExp(`^(<\\/?)(?:${escapeRegExp(original)}):`, "u");
  const attributeToken = new RegExp(`(\\s)${escapeRegExp(original)}:(?=[A-Za-z_][\\w.-]*\\s*=)`, "gu");
  const declarationToken = new RegExp(`xmlns:${escapeRegExp(original)}(?=\\s*=)`, "gu");
  return mapTags(xml, (tag) => tag
    .replaceAll(declarationToken, `xmlns:${replacement}`)
    .replace(elementToken, `$1${replacement}:`)
    .replaceAll(attributeToken, `$1${replacement}:`));
}

function replaceFirstChangedXml(entries, transform) {
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    if (!/\.(?:xml|rels)$/iu.test(entry.name)) continue;
    const xml = decoder.decode(entry.data);
    const transformed = transform(xml);
    if (transformed === xml) continue;
    return entries.map((candidate, candidateIndex) => candidateIndex === index
      ? { ...candidate, data: encoder.encode(transformed) }
      : candidate);
  }
  throw new Error("Package has no transformable XML entry");
}

function reorderedEntries(entries) {
  const mimetype = entries.find(({ name }) => name === "mimetype");
  const remaining = entries.filter(({ name }) => name !== "mimetype").reverse();
  return mimetype === undefined ? remaining : [mimetype, ...remaining];
}

/** Generate four rendering-invariant package mutations from one OOXML or ODF file. */
export function createMetamorphicPackages(bytes) {
  const entries = readZipEntries(bytes);
  const variants = [
    { kind: "xml-attribute-order", entries: replaceFirstChangedXml(entries, reorderAttributes), comment: "" },
    { kind: "namespace-prefix", entries: replaceFirstChangedXml(entries, renameNamespacePrefix), comment: "" },
    { kind: "zip-entry-order", entries: reorderedEntries(entries), comment: "" },
    { kind: "irrelevant-metadata", entries, comment: "OfficeViewer accuracy harness metadata mutation" },
  ];
  return variants.map(({ kind, entries: transformed, comment }) => ({ kind, bytes: writeZipEntries(transformed, { comment }) }));
}
