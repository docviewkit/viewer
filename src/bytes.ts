const objectTag = Object.prototype.toString;

/** A bucket key only; callers must compare bytes before reusing a resource. */
export function byteFingerprint(bytes: Uint8Array): string {
  let first = 0x811c9dc5;
  let second = 0x9e3779b9;
  for (let index = 0; index < bytes.length; index += 1) {
    const byte = bytes[index]!;
    first = Math.imul(first ^ byte, 0x01000193);
    second = Math.imul(second ^ byte, 0x85ebca6b);
  }
  return `${bytes.byteLength}:${first >>> 0}:${second >>> 0}`;
}

export function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
}

export function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return objectTag.call(value) === "[object ArrayBuffer]";
}

export function isUint8Array(value: unknown): value is Uint8Array {
  return ArrayBuffer.isView(value) && objectTag.call(value) === "[object Uint8Array]";
}

/** Returns an exact, current-realm ArrayBuffer copy of a supported byte source. */
export function copyByteSource(source: ArrayBuffer | Uint8Array): ArrayBuffer {
  const view = isArrayBuffer(source) ? new Uint8Array(source) : source;
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  return copy.buffer;
}
