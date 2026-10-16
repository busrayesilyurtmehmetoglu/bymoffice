// ZIP container reader/writer.
//
// Entries that are not modified are written back with their original compressed
// bytes, so saving a package never re-encodes content we did not touch.

import { deflateRaw, inflateRaw } from './deflate';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_EOCD64 = 0x06064b50;
const SIG_EOCD64_LOCATOR = 0x07064b50;

const EXTRA_ZIP64 = 0x0001;
const MAX32 = 0xffffffff;

const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

export interface ZipEntry {
  name: string;
  method: number;
  flags: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  dosTime: number;
  dosDate: number;
  externalAttrs: number;
  /** Name bytes as stored; kept so names that are not valid UTF-8 are written back unchanged. */
  rawName?: Uint8Array;
  versionMadeBy?: number;
  internalAttrs?: number;
  /** Extra fields of the central and the local header, without the ZIP64 field. */
  extra?: Uint8Array;
  localExtra?: Uint8Array;
  comment?: Uint8Array;
  /** Bytes exactly as stored in the archive (compressed when method is deflate). */
  raw: Uint8Array;
  /** Uncompressed replacement content, pending until the next write(). */
  data?: Uint8Array;
}

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const EMPTY = new Uint8Array(0);

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

export function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// Built-in streams are used where they exist; engines without them (iOS before
// 16.4, old Android WebViews) and those that lack the "deflate-raw" format fall
// back to our own codec.
// `globalThis` does not exist in the old WebViews the fallback codec is meant for.
const scope: unknown = typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : {};

function nativeStream(kind: 'CompressionStream' | 'DecompressionStream'): CompressionStream | DecompressionStream | undefined {
  const ctor = (scope as Record<string, unknown>)[kind] as (new (format: string) => CompressionStream) | undefined;
  if (typeof ctor !== 'function') return undefined;
  try {
    return new ctor('deflate-raw');
  } catch {
    return undefined;
  }
}

/** Runs data through a stream; output beyond `limit` bytes aborts instead of filling memory. */
async function pipe(data: Uint8Array, transform: CompressionStream | DecompressionStream, limit = Infinity): Promise<Uint8Array> {
  const reader = new Blob([data as BlobPart]).stream().pipeThrough(transform).getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.length;
    if (total > limit) {
      await reader.cancel().catch(() => {});
      throw new Error('inflated data is larger than the declared size');
    }
    chunks.push(value);
  }
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(total);
  let p = 0;
  for (const chunk of chunks) {
    out.set(chunk, p);
    p += chunk.length;
  }
  return out;
}

async function inflate(data: Uint8Array, size: number): Promise<Uint8Array> {
  const native = nativeStream('DecompressionStream');
  return native ? pipe(data, native, size) : inflateRaw(data, size);
}

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const native = nativeStream('CompressionStream');
  return native ? pipe(data, native) : deflateRaw(data);
}

// 64-bit fields are read as two halves so engines without BigInt can open ZIP64 archives.
function readUint64(view: DataView, offset: number): number {
  const high = view.getUint32(offset + 4, true);
  if (high > 0x1fffff) throw new Error('Corrupt ZIP archive: 64-bit value out of range');
  return high * 0x100000000 + view.getUint32(offset, true);
}

function writeUint64(view: DataView, offset: number, value: number): void {
  view.setUint32(offset, value >>> 0, true);
  view.setUint32(offset + 4, Math.floor(value / 0x100000000), true);
}

/** Returns a header's extra data without the ZIP64 field, whose values we store in the 32-bit fields. */
function stripZip64(extra: Uint8Array): Uint8Array {
  const view = new DataView(extra.buffer, extra.byteOffset, extra.byteLength);
  for (let p = 0; p + 4 <= extra.length; ) {
    const next = p + 4 + view.getUint16(p + 2, true);
    if (next > extra.length) break;
    if (view.getUint16(p, true) === EXTRA_ZIP64) {
      const out = new Uint8Array(extra.length - (next - p));
      out.set(extra.subarray(0, p));
      out.set(extra.subarray(next), p);
      return out;
    }
    p = next;
  }
  return extra;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

// What reading an archive may unpack. A small file must not be able to fill memory: deflate
// shrinks a run of one byte about a thousand times, and the parts of a package are read whole.
// The limits are well above what applications write (the body of a sheet or document is the
// largest part; past about 512 MB it does not fit in a string of a JavaScript engine anyway).

/** Most one part of a package unpacks to: the body of a sheet or document, shared strings, pictures and other files. */
export const MAX_UNPACKED_PART_BYTES = 512 * 1024 * 1024;
/** Most any other XML part unpacks to (styles, workbook, relationships, drawings, ...); as other parts that are read whole (charts, drawings). */
export const MAX_UNPACKED_XML_BYTES = 64 * 1024 * 1024;
/**
 * Most the parts read from one archive unpack to together, counting each part once, as a
 * multiple of the size of the archive: what applications write packs XML 5 to 50 times, a
 * decompression bomb about a thousand times.
 */
export const MAX_UNPACK_RATIO = 256;
/** What may be unpacked from any archive, however small, before MAX_UNPACK_RATIO applies. */
export const MIN_UNPACK_BUDGET_BYTES = 64 * 1024 * 1024;

/** A part of an archive unpacks to more than this engine reads (see MAX_UNPACKED_PART_BYTES and the limits beside it). */
export class PartTooLargeError extends Error {
  override name = 'PartTooLargeError';
  /** 'part': one part is larger than its limit; 'archive': the parts read together unpack to too much for the size of the file. */
  readonly kind: 'part' | 'archive';
  /** The entry that was being read; not in the message, which is shown to people. */
  readonly part: string;
  /** Its size when unpacked, as the archive declares it. */
  readonly size: number;
  readonly limit: number;

  constructor(kind: 'part' | 'archive', part: string, size: number, limit: number) {
    super(
      kind === 'part'
        ? `A part of the file is too large to open: it unpacks to ${size} bytes, more than the ${limit} this engine reads`
        : `The file unpacks to far more than its size (more than ${limit} bytes): it is not opened, as it may be made to fill memory`,
    );
    this.kind = kind;
    this.part = part;
    this.size = size;
    this.limit = limit;
  }
}

/**
 * The most an entry of a package unpacks to, by its name: MAX_UNPACKED_PART_BYTES for the
 * bodies of sheets and documents, shared strings, records of pivot caches and every file that
 * is not XML; MAX_UNPACKED_XML_BYTES for any other XML part. String work only: names come
 * from the file.
 */
export function unpackedLimitOf(name: string): number {
  const lower = name.toLowerCase();
  const slash = lower.lastIndexOf('/');
  const base = lower.slice(slash + 1);
  const dir = lower.slice(0, slash + 1);
  const xml = base.endsWith('.xml') || base.endsWith('.rels') || base.endsWith('.vml');
  if (!xml) return MAX_UNPACKED_PART_BYTES;
  const body =
    dir.endsWith('worksheets/') ||
    dir.endsWith('macrosheets/') ||
    dir.endsWith('dialogsheets/') ||
    base.startsWith('sharedstrings') ||
    base.startsWith('pivotcacherecords') ||
    (dir === 'word/' && base.indexOf('document') >= 0);
  return body ? MAX_UNPACKED_PART_BYTES : MAX_UNPACKED_XML_BYTES;
}

export class ZipArchive {
  private entries = new Map<string, ZipEntry>();
  /**
   * The entries as the archive had them, in its order, by name. An entry that is given the
   * content it had is written as the archive had it (compressed bytes, extra fields), and
   * one that was deleted and comes back takes its place again: a package that is changed and
   * changed back is the file that was opened. These objects are never changed.
   */
  private originals = new Map<string, { entry: ZipEntry; index: number }>();
  /** Archive comment, written back as it was read. */
  private comment: Uint8Array = EMPTY;
  /** Size of the archive that was read, for MAX_UNPACK_RATIO. */
  private archiveBytes = 0;
  /** What the entries inflated so far unpack to together, each counted once. */
  private unpacked = 0;
  private counted = new Set<string>();

  static read(bytes: Uint8Array): ZipArchive {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let eocd = -1;
    const lowest = Math.max(0, bytes.length - 22 - 0xffff);
    for (let i = bytes.length - 22; i >= lowest; i--) {
      if (view.getUint32(i, true) === SIG_EOCD) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('Not a ZIP archive: end of central directory not found');

    let count = view.getUint16(eocd + 10, true);
    let size = view.getUint32(eocd + 12, true);
    let p = view.getUint32(eocd + 16, true);
    if (count === 0xffff || size === MAX32 || p === MAX32) {
      // ZIP64: the real values live in a separate record found through a locator
      // that sits right before this one. Without a locator the values are literal.
      const locator = eocd - 20;
      if (locator >= 0 && view.getUint32(locator, true) === SIG_EOCD64_LOCATOR) {
        const record = readUint64(view, locator + 8);
        if (record + 56 > bytes.length || view.getUint32(record, true) !== SIG_EOCD64) {
          throw new Error('Corrupt ZIP archive: bad ZIP64 end of central directory');
        }
        count = readUint64(view, record + 32);
        size = readUint64(view, record + 40);
        p = readUint64(view, record + 48);
      }
    }
    const end = p + size;
    if (end > bytes.length || count * 46 > size) throw new Error('Corrupt ZIP archive: central directory out of bounds');

    const zip = new ZipArchive();
    zip.archiveBytes = bytes.length;
    const commentLength = Math.min(view.getUint16(eocd + 20, true), bytes.length - eocd - 22);
    zip.comment = bytes.subarray(eocd + 22, eocd + 22 + commentLength);
    // Writers without ZIP64 store the count modulo 65536, so headers that follow
    // the announced ones inside the directory are entries too.
    for (let n = 0; n < count || (p + 46 <= end && view.getUint32(p, true) === SIG_CENTRAL); n++) {
      if (p + 46 > end || view.getUint32(p, true) !== SIG_CENTRAL) {
        throw new Error('Corrupt ZIP archive: bad central directory entry');
      }
      const flags = view.getUint16(p + 8, true);
      const method = view.getUint16(p + 10, true);
      let compressedSize = view.getUint32(p + 20, true);
      let uncompressedSize = view.getUint32(p + 24, true);
      const nameLength = view.getUint16(p + 28, true);
      const extraLength = view.getUint16(p + 30, true);
      const entryCommentLength = view.getUint16(p + 32, true);
      let localOffset = view.getUint32(p + 42, true);
      const next = p + 46 + nameLength + extraLength + entryCommentLength;
      if (next > end) throw new Error('Corrupt ZIP archive: bad central directory entry');
      const rawName = bytes.subarray(p + 46, p + 46 + nameLength);
      const name = decoder.decode(rawName);
      // Keeping only one of two entries with the same name would lose the other on save.
      if (zip.entries.has(name)) throw new Error(`Corrupt ZIP archive: duplicate entry name ${name}`);

      if (flags & 1) throw new Error(`Encrypted ZIP entry is not supported: ${name}`);
      if (method !== METHOD_STORE && method !== METHOD_DEFLATE) {
        throw new Error(`Unsupported ZIP compression method ${method}: ${name}`);
      }
      if (compressedSize === MAX32 || uncompressedSize === MAX32 || localOffset === MAX32) {
        // The ZIP64 extra field holds, in this order, only the values that overflowed.
        let field = p + 46 + nameLength;
        const extraEnd = field + extraLength;
        let found = false;
        while (field + 4 <= extraEnd) {
          const fieldEnd = field + 4 + view.getUint16(field + 2, true);
          if (fieldEnd > extraEnd) break;
          if (view.getUint16(field, true) === EXTRA_ZIP64) {
            let q = field + 4;
            const take = () => {
              if (q + 8 > fieldEnd) throw new Error(`Corrupt ZIP archive: short ZIP64 extra field for ${name}`);
              q += 8;
              return readUint64(view, q - 8);
            };
            if (uncompressedSize === MAX32) uncompressedSize = take();
            if (compressedSize === MAX32) compressedSize = take();
            if (localOffset === MAX32) localOffset = take();
            found = true;
            break;
          }
          field = fieldEnd;
        }
        if (!found) throw new Error(`Corrupt ZIP archive: ZIP64 extra field is missing for ${name}`);
      }
      if (localOffset + 30 > bytes.length || view.getUint32(localOffset, true) !== SIG_LOCAL) {
        throw new Error(`Corrupt ZIP archive: bad local header for ${name}`);
      }
      const localExtraStart = localOffset + 30 + view.getUint16(localOffset + 26, true);
      const dataStart = localExtraStart + view.getUint16(localOffset + 28, true);
      if (dataStart + compressedSize > bytes.length) throw new Error(`Corrupt ZIP archive: truncated data for ${name}`);

      zip.entries.set(name, {
        name,
        method,
        flags,
        crc32: view.getUint32(p + 16, true),
        compressedSize,
        uncompressedSize,
        dosTime: view.getUint16(p + 12, true),
        dosDate: view.getUint16(p + 14, true),
        externalAttrs: view.getUint32(p + 38, true),
        rawName,
        versionMadeBy: view.getUint16(p + 4, true),
        internalAttrs: view.getUint16(p + 36, true),
        extra: stripZip64(bytes.subarray(p + 46 + nameLength, p + 46 + nameLength + extraLength)),
        localExtra: stripZip64(bytes.subarray(localExtraStart, dataStart)),
        comment: bytes.subarray(next - entryCommentLength, next),
        raw: bytes.subarray(dataStart, dataStart + compressedSize),
      });
      zip.originals.set(name, { entry: zip.entries.get(name) as ZipEntry, index: zip.originals.size });
      p = next;
    }
    return zip;
  }

  names(): string[] {
    return [...this.entries.keys()];
  }

  has(name: string): boolean {
    return this.entries.has(name);
  }

  entry(name: string): ZipEntry | undefined {
    return this.entries.get(name);
  }

  /**
   * Returns the uncompressed content of an entry. Rejects with PartTooLargeError, before
   * anything is inflated, for an entry larger than unpackedLimitOf() its name, and for one
   * that would take what the entries read from the archive unpack to together past
   * MAX_UNPACK_RATIO times its size (MIN_UNPACK_BUDGET_BYTES at least).
   */
  async read(name: string): Promise<Uint8Array> {
    const entry = this.entries.get(name);
    if (!entry) throw new Error(`ZIP entry not found: ${name}`);
    if (entry.data) return entry.data;
    // The limits are for what the file brings; what was written here since (and compressed by a write) is not held to them.
    const original = this.originals.get(name);
    const fromFile = original !== undefined && original.entry === entry;
    const limit = unpackedLimitOf(name);
    if (fromFile && entry.uncompressedSize > limit) throw new PartTooLargeError('part', name, entry.uncompressedSize, limit);
    let data = entry.raw;
    if (entry.method !== METHOD_STORE) {
      // Stored entries are in the file already; only what inflating makes counts.
      if (fromFile && !this.counted.has(name)) {
        const budget = Math.max(MIN_UNPACK_BUDGET_BYTES, MAX_UNPACK_RATIO * this.archiveBytes);
        if (this.unpacked + entry.uncompressedSize > budget) throw new PartTooLargeError('archive', name, entry.uncompressedSize, budget);
        this.unpacked += entry.uncompressedSize;
        this.counted.add(name);
      }
      try {
        data = await inflate(entry.raw, entry.uncompressedSize);
      } catch (error) {
        throw new Error(`Corrupt ZIP entry (${error instanceof Error ? error.message : error}): ${name}`);
      }
    }
    if (data.length !== entry.uncompressedSize || crc32(data) !== entry.crc32) {
      throw new Error(`Corrupt ZIP entry (size or CRC mismatch): ${name}`);
    }
    return data;
  }

  /**
   * Adds an entry or replaces its content; an existing entry keeps its position, and so does
   * one the archive had and that was deleted since.
   */
  set(name: string, data: Uint8Array): void {
    const original = this.originals.get(name);
    // Headers of the entry as it is now; of one that comes back, as the archive had them.
    const existing = this.entries.get(name) || (original ? original.entry : undefined);
    this.entries.set(name, {
      name,
      method: METHOD_DEFLATE,
      // The UTF-8 bit says how the stored name bytes are to be read, so it stays with them.
      flags: existing ? existing.flags & 0x0800 : 0,
      crc32: 0,
      compressedSize: 0,
      uncompressedSize: 0,
      // 1980-01-01 00:00 keeps output deterministic for new entries.
      dosTime: existing?.dosTime ?? 0,
      dosDate: existing?.dosDate ?? 0x21,
      externalAttrs: existing?.externalAttrs ?? 0,
      // Extra fields may describe the old content, so they are not carried over.
      rawName: existing?.rawName,
      versionMadeBy: existing?.versionMadeBy,
      internalAttrs: existing?.internalAttrs,
      comment: existing?.comment,
      raw: data,
      data,
    });
  }

  delete(name: string): boolean {
    return this.entries.delete(name);
  }

  /**
   * Whether the content waiting to be written for an entry is the content the archive had
   * for it: the same length and checksum first, which settles nearly every case without
   * inflating anything, then the bytes themselves.
   */
  private async isOriginal(entry: ZipEntry, original: ZipEntry): Promise<boolean> {
    const data = entry.data as Uint8Array;
    if (data.length !== original.uncompressedSize || crc32(data) !== original.crc32) return false;
    if (original.method === METHOD_STORE) return sameBytes(data, original.raw);
    try {
      return sameBytes(data, await inflate(original.raw, original.uncompressedSize));
    } catch (error) {
      // An entry of the archive that cannot be read is not what anybody wrote here.
      return false;
    }
  }

  /** The write that is under way: one at a time, each of the archive as it is when its turn comes. */
  private writing: Promise<unknown> = Promise.resolve();

  /**
   * The archive as a file. Writes of one archive follow each other (a second one waits for the
   * first), and each writes the entries as they are when it starts: what is set or deleted while
   * it runs is in the next write, never half in this one.
   */
  write(): Promise<Uint8Array> {
    const run = this.writing.then(
      () => this.writeNow(),
      () => this.writeNow(),
    );
    this.writing = run.catch(() => undefined);
    return run;
  }

  private async writeNow(): Promise<Uint8Array> {
    // Entries the archive had come first, in its order; whatever was deleted and came back is among them again.
    let ordered = true;
    let last = -1;
    let added = false;
    for (const name of this.entries.keys()) {
      const original = this.originals.get(name);
      if (!original) added = true;
      else if (added || original.index < last) ordered = false;
      else last = original.index;
    }
    if (!ordered) {
      const names = [...this.entries.keys()];
      const place = (name: string) => {
        const original = this.originals.get(name);
        return original ? original.index : Infinity;
      };
      // A stable sort: entries the archive did not have keep the order they were added in.
      const sorted = names.map((name, at) => ({ name, at })).sort((a, b) => place(a.name) - place(b.name) || a.at - b.at);
      const next = new Map<string, ZipEntry>();
      for (const { name } of sorted) next.set(name, this.entries.get(name) as ZipEntry);
      this.entries = next;
    }
    // The entries of this write: fixed here, whatever becomes of the archive while they are compressed.
    const mine = [...this.entries];
    const written: ZipEntry[] = [];
    for (const [name, entry] of mine) {
      if (!entry.data) {
        written.push(entry);
        continue;
      }
      const original = this.originals.get(name);
      if (original && (await this.isOriginal(entry, original.entry))) {
        if (this.entries.get(name) === entry) this.entries.set(name, original.entry);
        written.push(original.entry);
        continue;
      }
      written.push(entry);
      const compressed = entry.data.length > 0 ? await deflate(entry.data) : entry.data;
      const useDeflate = compressed.length < entry.data.length;
      entry.method = useDeflate ? METHOD_DEFLATE : METHOD_STORE;
      entry.raw = useDeflate ? compressed : entry.data;
      entry.crc32 = crc32(entry.data);
      entry.compressedSize = entry.raw.length;
      entry.uncompressedSize = entry.data.length;
      entry.flags &= 0x0800;
      delete entry.data;
    }
    const chunks: Uint8Array[] = [];
    const central: Uint8Array[] = [];
    let offset = 0;
    for (const entry of written) {
      const name = entry.rawName ?? encoder.encode(entry.name);
      // Keep the deflate-level bits; drop the data-descriptor bit since sizes go in the header.
      let flags = entry.flags & 0x0006;
      if (entry.rawName ? entry.flags & 0x0800 : name.length !== entry.name.length) flags |= 0x0800;
      const extra = entry.extra ?? EMPTY;
      const localExtra = entry.localExtra ?? EMPTY;
      const comment = entry.comment ?? EMPTY;

      const local = new Uint8Array(30 + name.length + localExtra.length);
      const lv = new DataView(local.buffer);
      lv.setUint32(0, SIG_LOCAL, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, flags, true);
      lv.setUint16(8, entry.method, true);
      lv.setUint16(10, entry.dosTime, true);
      lv.setUint16(12, entry.dosDate, true);
      lv.setUint32(14, entry.crc32, true);
      lv.setUint32(18, entry.compressedSize, true);
      lv.setUint32(22, entry.uncompressedSize, true);
      lv.setUint16(26, name.length, true);
      lv.setUint16(28, localExtra.length, true);
      local.set(name, 30);
      local.set(localExtra, 30 + name.length);

      const header = new Uint8Array(46 + name.length + extra.length + comment.length);
      const cv = new DataView(header.buffer);
      cv.setUint32(0, SIG_CENTRAL, true);
      cv.setUint16(4, entry.versionMadeBy ?? 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, flags, true);
      cv.setUint16(10, entry.method, true);
      cv.setUint16(12, entry.dosTime, true);
      cv.setUint16(14, entry.dosDate, true);
      cv.setUint32(16, entry.crc32, true);
      cv.setUint32(20, entry.compressedSize, true);
      cv.setUint32(24, entry.uncompressedSize, true);
      cv.setUint16(28, name.length, true);
      cv.setUint16(30, extra.length, true);
      cv.setUint16(32, comment.length, true);
      cv.setUint16(36, entry.internalAttrs ?? 0, true);
      cv.setUint32(38, entry.externalAttrs, true);
      cv.setUint32(42, offset, true);
      header.set(name, 46);
      header.set(extra, 46 + name.length);
      header.set(comment, 46 + name.length + extra.length);

      chunks.push(local, entry.raw);
      central.push(header);
      offset += local.length + entry.raw.length;
      if (offset > 0xfffffffe) throw new Error('Archive too large: writing archives of 4 GB or more is not supported yet');
    }

    // Of this write, like everything else in it: what the archive holds by now may be more or less.
    const count = written.length;
    const centralSize = central.reduce((sum, c) => sum + c.length, 0);
    if (offset + centralSize > 0xfffffffe) {
      throw new Error('Archive too large: writing archives of 4 GB or more is not supported yet');
    }
    const tail: Uint8Array[] = [];
    if (count > 0xfffe) {
      // More entries than the 16-bit counters hold: add the ZIP64 record and its locator.
      const record = new Uint8Array(56);
      const rv = new DataView(record.buffer);
      rv.setUint32(0, SIG_EOCD64, true);
      writeUint64(rv, 4, 44);
      rv.setUint16(12, 45, true);
      rv.setUint16(14, 45, true);
      writeUint64(rv, 24, count);
      writeUint64(rv, 32, count);
      writeUint64(rv, 40, centralSize);
      writeUint64(rv, 48, offset);

      const locator = new Uint8Array(20);
      const tv = new DataView(locator.buffer);
      tv.setUint32(0, SIG_EOCD64_LOCATOR, true);
      writeUint64(tv, 8, offset + centralSize);
      tv.setUint32(16, 1, true);
      tail.push(record, locator);
    }
    const eocd = new Uint8Array(22 + this.comment.length);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, SIG_EOCD, true);
    ev.setUint16(8, Math.min(count, 0xffff), true);
    ev.setUint16(10, Math.min(count, 0xffff), true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, this.comment.length, true);
    eocd.set(this.comment, 22);
    tail.push(eocd);

    const out = new Uint8Array(offset + centralSize + tail.reduce((sum, c) => sum + c.length, 0));
    let p = 0;
    for (const chunk of [...chunks, ...central, ...tail]) {
      out.set(chunk, p);
      p += chunk.length;
    }
    return out;
  }
}
