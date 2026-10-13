// Compression of the streams of a PDF: the zlib format (RFC 1950) that the
// FlateDecode filter reads. The built-in CompressionStream does it where
// there is one; everywhere else the engine's own deflate, wrapped here.

import { deflateRaw } from '../deflate';

// `globalThis` does not exist in the old WebViews the fallback codec is meant for.
const scope: unknown = typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : {};

/** The Adler-32 checksum the zlib format ends with. */
export function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  for (let i = 0; i < data.length; ) {
    // As many bytes as fit before the sums must be reduced.
    const end = Math.min(i + 3800, data.length);
    for (; i < end; i++) {
      a += data[i];
      b += a;
    }
    a %= 65521;
    b %= 65521;
  }
  return ((b << 16) | a) >>> 0;
}

/** zlib data made by the engine's own deflate. */
export function zlibOwn(data: Uint8Array): Uint8Array {
  const raw = deflateRaw(data);
  const out = new Uint8Array(raw.length + 6);
  // 32K window, default compression; the two bytes are a multiple of 31.
  out[0] = 0x78;
  out[1] = 0x9c;
  out.set(raw, 2);
  const sum = adler32(data);
  out[raw.length + 2] = sum >>> 24;
  out[raw.length + 3] = (sum >>> 16) & 255;
  out[raw.length + 4] = (sum >>> 8) & 255;
  out[raw.length + 5] = sum & 255;
  return out;
}

function nativeStream(): CompressionStream | undefined {
  const ctor = (scope as Record<string, unknown>).CompressionStream as (new (format: string) => CompressionStream) | undefined;
  if (typeof ctor !== 'function' || typeof Blob === 'undefined' || typeof Blob.prototype.stream !== 'function') return undefined;
  try {
    return new ctor('deflate');
  } catch {
    return undefined;
  }
}

/** Data in the zlib format. `native: false` keeps to the engine's own code. */
export async function zlibDeflate(data: Uint8Array, native = true): Promise<Uint8Array> {
  const stream = native ? nativeStream() : undefined;
  if (!stream) return zlibOwn(data);
  const reader = new Blob([data as BlobPart]).stream().pipeThrough(stream).getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(total);
  let at = 0;
  for (let i = 0; i < chunks.length; i++) {
    out.set(chunks[i], at);
    at += chunks[i].length;
  }
  return out;
}
