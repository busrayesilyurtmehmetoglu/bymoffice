// Raw DEFLATE (RFC 1951) decoder and encoder.
//
// Synchronous, dependency-free fallback for environments that lack the built-in
// DecompressionStream/CompressionStream (iOS before 16.4, old Android WebViews).
// Only standard typed arrays are used, so the same code runs in browsers and Node.

const MAX_BITS = 15;
const MAX_CL_BITS = 7;
const NUM_LITLEN = 286;
const NUM_DIST = 30;
const END_OF_BLOCK = 256;
const WINDOW_SIZE = 32768;
const MIN_MATCH = 3;
const MAX_MATCH = 258;

const LEN_BASE = new Uint16Array([
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258,
]);
const LEN_EXTRA = new Uint8Array([
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
]);
const DIST_BASE = new Uint16Array([
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
  8193, 12289, 16385, 24577,
]);
const DIST_EXTRA = new Uint8Array([
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
]);
/** Order in which code length code lengths are stored in a dynamic block header. */
const CL_ORDER = new Uint8Array([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]);

function fixedLitLenLengths(): Uint8Array {
  const lengths = new Uint8Array(288);
  lengths.fill(8, 0, 144);
  lengths.fill(9, 144, 256);
  lengths.fill(7, 256, 280);
  lengths.fill(8, 280, 288);
  return lengths;
}

function corrupt(detail: string): Error {
  return new Error(`Invalid DEFLATE data: ${detail}`);
}

// ---------------------------------------------------------------------------
// Decoder
// ---------------------------------------------------------------------------

const KIND_CODE_LENGTHS = 0;
const KIND_LITLEN = 1;
const KIND_DIST = 2;

const LIT_FAST_BITS = 10;
const DIST_FAST_BITS = 8;

/**
 * Canonical Huffman decoding table. Codes of up to `fastBits` bits resolve with
 * one lookup; `fast` entries are `(symbol << 4) | length`, 0 meaning "not here".
 * Longer (rare) codes fall back to a canonical walk over `count`/`symbols`.
 */
interface DecodeTable {
  fast: Uint16Array;
  fastBits: number;
  /** Number of codes of each length. */
  count: Uint16Array;
  /** Symbols ordered by code length, then by value. */
  symbols: Uint16Array;
}

function newDecodeTable(maxSymbols: number, fastBits: number): DecodeTable {
  return {
    fast: new Uint16Array(1 << fastBits),
    fastBits,
    count: new Uint16Array(MAX_BITS + 1),
    symbols: new Uint16Array(maxSymbols),
  };
}

/**
 * Fills `table` from code lengths. Follows zlib's rules: over-subscribed sets are
 * rejected; incomplete sets are rejected except for a literal/length or distance
 * code that consists of a single one-bit code, or a set with no codes at all.
 */
function buildDecodeTable(table: DecodeTable, lengths: Uint8Array, n: number, kind: number): void {
  const { fast, fastBits, count, symbols } = table;
  const what = kind === KIND_CODE_LENGTHS ? 'code length' : kind === KIND_LITLEN ? 'literal/length' : 'distance';
  count.fill(0);
  fast.fill(0);
  for (let i = 0; i < n; i++) count[lengths[i]]++;
  count[0] = 0;

  let left = 1;
  let max = 0;
  for (let len = 1; len <= MAX_BITS; len++) {
    left = (left << 1) - count[len];
    if (left < 0) throw corrupt(`over-subscribed ${what} code`);
    if (count[len]) max = len;
  }
  if (max === 0) {
    // No codes: any attempt to decode a symbol fails. Legal for distances (a block of literals only).
    if (kind === KIND_CODE_LENGTHS) throw corrupt('empty code length code');
    return;
  }
  if (left > 0 && (kind === KIND_CODE_LENGTHS || max !== 1)) throw corrupt(`incomplete ${what} code`);

  const next = new Uint16Array(MAX_BITS + 2);
  const code = new Uint16Array(MAX_BITS + 2);
  for (let len = 1; len <= MAX_BITS; len++) {
    next[len + 1] = next[len] + count[len];
    code[len + 1] = (code[len] + count[len]) << 1;
  }
  const size = 1 << fastBits;
  for (let sym = 0; sym < n; sym++) {
    const len = lengths[sym];
    if (len === 0) continue;
    symbols[next[len]++] = sym;
    const c = code[len]++;
    if (len > fastBits) continue;
    // Codes are packed most significant bit first; the bit buffer is LSB first.
    let reversed = 0;
    for (let i = 0; i < len; i++) reversed |= ((c >>> i) & 1) << (len - 1 - i);
    const entry = (sym << 4) | len;
    for (let k = reversed; k < size; k += 1 << len) fast[k] = entry;
  }
}

/** Decodes a code that missed the fast table from 15 buffered bits. Returns `(symbol << 4) | length`, or -1. */
function slowDecode(table: DecodeTable, bits: number): number {
  const { count, symbols } = table;
  let code = 0;
  let first = 0;
  let index = 0;
  for (let len = 1; len <= MAX_BITS; len++) {
    code |= bits & 1;
    bits >>>= 1;
    const n = count[len];
    if (code - n < first) return (symbols[index + (code - first)] << 4) | len;
    index += n;
    first = (first + n) << 1;
    code <<= 1;
  }
  return -1;
}

/**
 * LSB-first bit reader. Past the end of the input it feeds zero bits and keeps
 * counting, so consuming a bit that was never there is detected right after.
 */
class BitReader {
  pos = 0;
  buf = 0;
  cnt = 0;

  constructor(readonly data: Uint8Array) {}

  /** Reads `n` (at most 16) bits. */
  bits(n: number): number {
    const data = this.data;
    while (this.cnt < n) {
      this.buf |= (this.pos < data.length ? data[this.pos] : 0) << this.cnt;
      this.pos++;
      this.cnt += 8;
    }
    const value = this.buf & ((1 << n) - 1);
    this.buf >>>= n;
    this.cnt -= n;
    this.checkEnd();
    return value;
  }

  symbol(table: DecodeTable, what: string): number {
    const data = this.data;
    while (this.cnt < MAX_BITS) {
      this.buf |= (this.pos < data.length ? data[this.pos] : 0) << this.cnt;
      this.pos++;
      this.cnt += 8;
    }
    let entry = table.fast[this.buf & ((1 << table.fastBits) - 1)];
    if (entry === 0) {
      entry = slowDecode(table, this.buf);
      if (entry < 0) {
        this.checkEnd();
        throw corrupt(`invalid ${what} code`);
      }
    }
    this.buf >>>= entry & 15;
    this.cnt -= entry & 15;
    this.checkEnd();
    return entry >>> 4;
  }

  /** Drops the rest of the current byte and returns whole buffered bytes to the input. */
  alignToByte(): void {
    this.cnt -= this.cnt & 7;
    this.pos -= this.cnt >>> 3;
    this.buf = 0;
    this.cnt = 0;
    if (this.pos > this.data.length) throw truncated();
  }

  private checkEnd(): void {
    if (this.pos > this.data.length && this.cnt < (this.pos - this.data.length) << 3) throw truncated();
  }
}

function truncated(): Error {
  return corrupt('unexpected end of data');
}

/** Output buffer that grows on demand and never beyond `limit`. */
interface Output {
  buf: Uint8Array;
  pos: number;
  limit: number;
}

function grow(buf: Uint8Array, used: number, extra: number, limit: number): Uint8Array {
  const need = used + extra;
  if (need > limit) throw new Error(`Inflated data exceeds the limit of ${limit} bytes`);
  const bigger = new Uint8Array(Math.min(limit, Math.max(buf.length * 2, need, 1024)));
  bigger.set(buf.subarray(0, used));
  return bigger;
}

let fixedLitLen: DecodeTable | undefined;
let fixedDist: DecodeTable | undefined;

function readStoredBlock(reader: BitReader, out: Output): void {
  reader.alignToByte();
  const data = reader.data;
  let pos = reader.pos;
  if (pos + 4 > data.length) throw truncated();
  const len = data[pos] | (data[pos + 1] << 8);
  const nlen = data[pos + 2] | (data[pos + 3] << 8);
  if ((len ^ nlen) !== 0xffff) throw corrupt('stored block length does not match its complement');
  pos += 4;
  if (pos + len > data.length) throw truncated();
  if (out.pos + len > out.buf.length) out.buf = grow(out.buf, out.pos, len, out.limit);
  out.buf.set(data.subarray(pos, pos + len), out.pos);
  out.pos += len;
  reader.pos = pos + len;
}

function readDynamicTables(
  reader: BitReader,
  codeLengths: DecodeTable,
  litLen: DecodeTable,
  dist: DecodeTable,
  lengths: Uint8Array,
): void {
  const hlit = reader.bits(5) + 257;
  const hdist = reader.bits(5) + 1;
  const hclen = reader.bits(4) + 4;
  if (hlit > NUM_LITLEN || hdist > NUM_DIST) throw corrupt('too many length or distance symbols');

  lengths.fill(0, 0, 19);
  for (let i = 0; i < hclen; i++) lengths[CL_ORDER[i]] = reader.bits(3);
  buildDecodeTable(codeLengths, lengths, 19, KIND_CODE_LENGTHS);

  const total = hlit + hdist;
  let i = 0;
  while (i < total) {
    const sym = reader.symbol(codeLengths, 'code length');
    if (sym < 16) {
      lengths[i++] = sym;
      continue;
    }
    let value = 0;
    let repeat: number;
    if (sym === 16) {
      if (i === 0) throw corrupt('code length repeat with no previous length');
      value = lengths[i - 1];
      repeat = 3 + reader.bits(2);
    } else if (sym === 17) {
      repeat = 3 + reader.bits(3);
    } else {
      repeat = 11 + reader.bits(7);
    }
    if (i + repeat > total) throw corrupt('code length repeat runs past the end of the table');
    lengths.fill(value, i, i + repeat);
    i += repeat;
  }
  if (lengths[END_OF_BLOCK] === 0) throw corrupt('missing end-of-block code');

  buildDecodeTable(litLen, lengths, hlit, KIND_LITLEN);
  buildDecodeTable(dist, lengths.subarray(hlit, total), hdist, KIND_DIST);
}

/** Decodes the symbols of one Huffman-coded block. The hot loop keeps all state in locals. */
function readCompressedBlock(reader: BitReader, litLen: DecodeTable, distTable: DecodeTable, out: Output): void {
  const data = reader.data;
  const end = data.length;
  let pos = reader.pos;
  let buf = reader.buf;
  let cnt = reader.cnt;
  let dst = out.buf;
  let op = out.pos;
  const limit = out.limit;
  const litFast = litLen.fast;
  const litMask = (1 << litLen.fastBits) - 1;
  const distFast = distTable.fast;
  const distMask = (1 << distTable.fastBits) - 1;

  for (;;) {
    while (cnt < MAX_BITS) {
      buf |= (pos < end ? data[pos] : 0) << cnt;
      pos++;
      cnt += 8;
    }
    let entry = litFast[buf & litMask];
    if (entry === 0) entry = slowDecode(litLen, buf);
    if (pos > end && (entry < 0 || cnt - (entry & 15) < (pos - end) << 3)) throw truncated();
    if (entry < 0) throw corrupt('invalid literal/length code');
    buf >>>= entry & 15;
    cnt -= entry & 15;
    let sym = entry >>> 4;

    if (sym < 256) {
      if (op === dst.length) dst = grow(dst, op, 1, limit);
      dst[op++] = sym;
      continue;
    }
    if (sym === END_OF_BLOCK) break;
    if (sym >= 257 + LEN_BASE.length) throw corrupt('invalid length symbol');
    sym -= 257;
    let length = LEN_BASE[sym];
    let extra = LEN_EXTRA[sym];
    if (extra) {
      while (cnt < extra) {
        buf |= (pos < end ? data[pos] : 0) << cnt;
        pos++;
        cnt += 8;
      }
      length += buf & ((1 << extra) - 1);
      buf >>>= extra;
      cnt -= extra;
    }

    while (cnt < MAX_BITS) {
      buf |= (pos < end ? data[pos] : 0) << cnt;
      pos++;
      cnt += 8;
    }
    entry = distFast[buf & distMask];
    if (entry === 0) entry = slowDecode(distTable, buf);
    if (pos > end && (entry < 0 || cnt - (entry & 15) < (pos - end) << 3)) throw truncated();
    if (entry < 0) throw corrupt('invalid distance code');
    buf >>>= entry & 15;
    cnt -= entry & 15;
    sym = entry >>> 4;
    if (sym >= NUM_DIST) throw corrupt('invalid distance symbol');
    let distance = DIST_BASE[sym];
    extra = DIST_EXTRA[sym];
    if (extra) {
      while (cnt < extra) {
        buf |= (pos < end ? data[pos] : 0) << cnt;
        pos++;
        cnt += 8;
      }
      distance += buf & ((1 << extra) - 1);
      buf >>>= extra;
      cnt -= extra;
    }
    if (pos > end && cnt < (pos - end) << 3) throw truncated();

    if (distance > op) throw corrupt('distance points before the start of the output');
    if (op + length > dst.length) dst = grow(dst, op, length, limit);
    let from = op - distance;
    if (distance >= length && length > 16) {
      dst.copyWithin(op, from, from + length);
      op += length;
    } else {
      // Overlapping copies must go byte by byte: the source catches up with the output.
      const stop = op + length;
      while (op < stop) dst[op++] = dst[from++];
    }
  }

  reader.pos = pos;
  reader.buf = buf;
  reader.cnt = cnt;
  out.buf = dst;
  out.pos = op;
}

/** Decodes a raw DEFLATE stream (no zlib/gzip wrapper). `maxOutput` caps the output size; exceeding it throws. */
export function inflateRaw(data: Uint8Array, maxOutput?: number): Uint8Array {
  let limit = Infinity;
  if (maxOutput !== undefined) {
    if (!(maxOutput >= 0)) throw new RangeError('maxOutput must be a non-negative number');
    limit = Math.floor(maxOutput);
  }
  // The declared limit is not trusted for allocation: start proportional to the input and grow.
  const out: Output = { buf: new Uint8Array(Math.min(limit, Math.max(1024, data.length * 4))), pos: 0, limit };
  const reader = new BitReader(data);
  let codeLengths: DecodeTable | undefined;
  let litLen: DecodeTable | undefined;
  let dist: DecodeTable | undefined;
  let lengths: Uint8Array | undefined;

  // Every block consumes at least three bits, so the loop is bounded by the input size.
  for (;;) {
    const final = reader.bits(1);
    const type = reader.bits(2);
    if (type === 0) {
      readStoredBlock(reader, out);
    } else if (type === 1) {
      if (!fixedLitLen || !fixedDist) {
        fixedLitLen = newDecodeTable(288, LIT_FAST_BITS);
        buildDecodeTable(fixedLitLen, fixedLitLenLengths(), 288, KIND_LITLEN);
        fixedDist = newDecodeTable(32, DIST_FAST_BITS);
        buildDecodeTable(fixedDist, new Uint8Array(32).fill(5), 32, KIND_DIST);
      }
      readCompressedBlock(reader, fixedLitLen, fixedDist, out);
    } else if (type === 2) {
      if (!codeLengths || !litLen || !dist || !lengths) {
        codeLengths = newDecodeTable(19, MAX_CL_BITS);
        litLen = newDecodeTable(NUM_LITLEN, LIT_FAST_BITS);
        dist = newDecodeTable(NUM_DIST, DIST_FAST_BITS);
        lengths = new Uint8Array(NUM_LITLEN + NUM_DIST);
      }
      readDynamicTables(reader, codeLengths, litLen, dist, lengths);
      readCompressedBlock(reader, litLen, dist, out);
    } else {
      throw corrupt('invalid block type');
    }
    if (final) break;
  }
  // Bytes after the final block are ignored.
  return out.pos === out.buf.length ? out.buf : out.buf.slice(0, out.pos);
}

// ---------------------------------------------------------------------------
// Encoder
// ---------------------------------------------------------------------------

const HASH_BITS = 15;
const HASH_SIZE = 1 << HASH_BITS;
const HASH_MASK = HASH_SIZE - 1;
const WINDOW_MASK = WINDOW_SIZE - 1;
/** Symbols per block; a block is closed when this many literals and matches are buffered. */
const BLOCK_TOKENS = 32768;
const MAX_STORED = 65535;

// Match search effort, equivalent to zlib level 6.
const MAX_CHAIN = 128;
const GOOD_LENGTH = 8;
const MAX_LAZY = 16;
const NICE_LENGTH = 128;
/** A minimum-length match this far back costs more than three literals. */
const TOO_FAR = 4096;

/** Match length (3..258) to length symbol minus 257. */
const LEN_SYMBOL = (() => {
  const table = new Uint8Array(MAX_MATCH + 1);
  for (let sym = 0; sym < LEN_BASE.length; sym++) {
    const last = Math.min(MAX_MATCH, LEN_BASE[sym] + (1 << LEN_EXTRA[sym]) - 1);
    for (let len = LEN_BASE[sym]; len <= last; len++) table[len] = sym;
  }
  return table;
})();

/** Distance to distance symbol: index `d - 1` up to 256, `256 + ((d - 1) >> 7)` beyond. */
const DIST_SYMBOL = (() => {
  const table = new Uint8Array(512);
  for (let sym = 0; sym < NUM_DIST; sym++) {
    const last = DIST_BASE[sym] + (1 << DIST_EXTRA[sym]) - 1;
    for (let d = DIST_BASE[sym]; d <= last; d++) table[d <= 256 ? d - 1 : 256 + ((d - 1) >> 7)] = sym;
  }
  return table;
})();

/**
 * Computes length-limited Huffman code lengths for the given symbol frequencies.
 * The result is always a complete prefix code; when fewer than two symbols are
 * used, two one-bit codes are assigned so that every decoder accepts the table.
 *
 * @internal Exported for tests only; not part of the public API.
 */
export function buildCodeLengths(freq: Uint32Array, limit: number): Uint8Array {
  const n = freq.length;
  const lengths = new Uint8Array(n);
  const used: number[] = [];
  for (let i = 0; i < n; i++) if (freq[i]) used.push(i);
  if (used.length < 2) {
    const first = used.length ? used[0] : 0;
    lengths[first] = 1;
    lengths[first === 0 ? 1 : 0] = 1;
    return lengths;
  }
  // Ascending frequency; ties by symbol so the result is deterministic.
  used.sort((a, b) => freq[a] - freq[b] || a - b);

  // Huffman tree with two queues: leaves are 0..m-1, internal nodes are created in ascending weight order.
  const m = used.length;
  const weight = new Float64Array(2 * m - 1);
  const parent = new Int32Array(2 * m - 1);
  for (let i = 0; i < m; i++) weight[i] = freq[used[i]];
  let leaf = 0;
  let node = m;
  for (let next = m; next < 2 * m - 1; next++) {
    const a = leaf < m && (node >= next || weight[leaf] <= weight[node]) ? leaf++ : node++;
    const b = leaf < m && (node >= next || weight[leaf] <= weight[node]) ? leaf++ : node++;
    weight[next] = weight[a] + weight[b];
    parent[a] = next;
    parent[b] = next;
  }
  const depth = new Int32Array(2 * m - 1);
  const count = new Int32Array(limit + 1);
  for (let k = 2 * m - 3; k >= 0; k--) {
    depth[k] = depth[parent[k]] + 1;
    if (k < m) count[Math.min(depth[k], limit)]++;
  }

  // Clamping deep leaves to `limit` over-subscribes the code. Each step below removes one
  // unit of the excess: a leaf above the limit level moves one level down and takes one
  // of the clamped leaves as its sibling.
  let total = 0;
  for (let len = 1; len <= limit; len++) total += count[len] * 2 ** (limit - len);
  const full = 2 ** limit;
  while (total > full) {
    count[limit]--;
    for (let len = limit - 1; len > 0; len--) {
      if (count[len]) {
        count[len]--;
        count[len + 1] += 2;
        break;
      }
    }
    total--;
  }

  // Longest codes go to the least frequent symbols.
  let i = 0;
  for (let len = limit; len > 0; len--) {
    for (let k = count[len]; k > 0; k--) lengths[used[i++]] = len;
  }
  return lengths;
}

/** Canonical codes for the given lengths, bit-reversed so they can be written LSB first. */
function reversedCodes(lengths: Uint8Array): Uint16Array {
  const count = new Uint16Array(MAX_BITS + 2);
  for (let i = 0; i < lengths.length; i++) count[lengths[i]]++;
  count[0] = 0;
  const next = new Uint16Array(MAX_BITS + 2);
  for (let len = 1; len <= MAX_BITS; len++) next[len + 1] = (next[len] + count[len]) << 1;
  const codes = new Uint16Array(lengths.length);
  for (let sym = 0; sym < lengths.length; sym++) {
    const len = lengths[sym];
    if (len === 0) continue;
    const c = next[len]++;
    let reversed = 0;
    for (let i = 0; i < len; i++) reversed |= ((c >>> i) & 1) << (len - 1 - i);
    codes[sym] = reversed;
  }
  return codes;
}

const FIXED_LIT_LENGTHS = fixedLitLenLengths();
const FIXED_LIT_CODES = reversedCodes(FIXED_LIT_LENGTHS);
const FIXED_DIST_LENGTHS = new Uint8Array(NUM_DIST).fill(5);
const FIXED_DIST_CODES = reversedCodes(FIXED_DIST_LENGTHS);

class BitWriter {
  out: Uint8Array;
  pos = 0;
  private buf = 0;
  private cnt = 0;

  constructor(capacity: number) {
    this.out = new Uint8Array(capacity);
  }

  private reserve(extra: number): void {
    if (this.pos + extra <= this.out.length) return;
    const bigger = new Uint8Array(Math.max(this.out.length * 2, this.pos + extra));
    bigger.set(this.out.subarray(0, this.pos));
    this.out = bigger;
  }

  /** Writes the low `n` (at most 16) bits of `value`, least significant bit first. */
  bits(value: number, n: number): void {
    this.buf |= value << this.cnt;
    this.cnt += n;
    if (this.pos + 3 > this.out.length) this.reserve(3);
    while (this.cnt >= 8) {
      this.out[this.pos++] = this.buf & 0xff;
      this.buf >>>= 8;
      this.cnt -= 8;
    }
  }

  /** Pads with zero bits up to the next byte boundary. */
  align(): void {
    if (this.cnt === 0) return;
    this.reserve(1);
    this.out[this.pos++] = this.buf & 0xff;
    this.buf = 0;
    this.cnt = 0;
  }

  /** Appends whole bytes; the writer must be byte-aligned. */
  bytes(src: Uint8Array): void {
    this.reserve(src.length);
    this.out.set(src, this.pos);
    this.pos += src.length;
  }

  finish(): Uint8Array {
    this.align();
    return this.out.slice(0, this.pos);
  }
}

/** Buffers the literals and matches of one block and writes it in the cheapest of the three block types. */
class BlockEncoder {
  /** Literal byte, or match length when the distance is non-zero. */
  private tokenValue = new Uint16Array(BLOCK_TOKENS);
  /** Match distance minus one, or 0xffff for a literal. */
  private tokenDist = new Uint16Array(BLOCK_TOKENS);
  private tokens = 0;
  private litFreq = new Uint32Array(NUM_LITLEN);
  private distFreq = new Uint32Array(NUM_DIST);
  /** Input range covered by the buffered tokens. */
  private start = 0;
  private length = 0;

  constructor(
    private readonly data: Uint8Array,
    private readonly writer: BitWriter,
  ) {}

  literal(byte: number): void {
    this.tokenValue[this.tokens] = byte;
    this.tokenDist[this.tokens++] = 0xffff;
    this.litFreq[byte]++;
    this.length++;
    if (this.tokens === BLOCK_TOKENS) this.flush(false);
  }

  match(length: number, distance: number): void {
    const d = distance - 1;
    this.tokenValue[this.tokens] = length;
    this.tokenDist[this.tokens++] = d;
    this.litFreq[257 + LEN_SYMBOL[length]]++;
    this.distFreq[DIST_SYMBOL[d < 256 ? d : 256 + (d >> 7)]]++;
    this.length += length;
    if (this.tokens === BLOCK_TOKENS) this.flush(false);
  }

  flush(final: boolean): void {
    const { writer, litFreq, distFreq } = this;
    litFreq[END_OF_BLOCK]++;
    const litLengths = buildCodeLengths(litFreq, MAX_BITS);
    const distLengths = buildCodeLengths(distFreq, MAX_BITS);

    let hlit = NUM_LITLEN;
    while (hlit > 257 && litLengths[hlit - 1] === 0) hlit--;
    let hdist = NUM_DIST;
    while (hdist > 1 && distLengths[hdist - 1] === 0) hdist--;

    // Run-length encode both length tables as one sequence (runs may cross the boundary).
    const all = new Uint8Array(hlit + hdist);
    all.set(litLengths.subarray(0, hlit));
    all.set(distLengths.subarray(0, hdist), hlit);
    const clSymbols = new Uint8Array(all.length);
    const clExtra = new Uint8Array(all.length);
    const clFreq = new Uint32Array(19);
    let clCount = 0;
    const emit = (sym: number, extra: number) => {
      clSymbols[clCount] = sym;
      clExtra[clCount++] = extra;
      clFreq[sym]++;
    };
    for (let i = 0; i < all.length; ) {
      const value = all[i];
      let run = 1;
      while (i + run < all.length && all[i + run] === value) run++;
      i += run;
      if (value !== 0) {
        emit(value, 0);
        run--;
      }
      while (run > 0) {
        let taken = 1;
        if (value === 0 && run >= 11) {
          taken = Math.min(run, 138);
          emit(18, taken - 11);
        } else if (value === 0 && run >= 3) {
          taken = run;
          emit(17, taken - 3);
        } else if (value !== 0 && run >= 3) {
          taken = Math.min(run, 6);
          emit(16, taken - 3);
        } else {
          emit(value, 0);
        }
        run -= taken;
      }
    }
    const clLengths = buildCodeLengths(clFreq, MAX_CL_BITS);
    let hclen = 19;
    while (hclen > 4 && clLengths[CL_ORDER[hclen - 1]] === 0) hclen--;

    let dynamicBits = 3 + 5 + 5 + 4 + 3 * hclen + 2 * clFreq[16] + 3 * clFreq[17] + 7 * clFreq[18];
    for (let sym = 0; sym < 19; sym++) dynamicBits += clFreq[sym] * clLengths[sym];
    dynamicBits += this.bodyBits(litLengths, distLengths);
    const fixedBits = 3 + this.bodyBits(FIXED_LIT_LENGTHS, FIXED_DIST_LENGTHS);
    const storedBlocks = Math.max(1, Math.ceil(this.length / MAX_STORED));
    const storedBits = (this.length + 5 * storedBlocks) * 8;

    if (storedBits <= Math.min(fixedBits, dynamicBits)) {
      let offset = this.start;
      let remaining = this.length;
      do {
        const chunk = Math.min(remaining, MAX_STORED);
        remaining -= chunk;
        writer.bits(final && remaining === 0 ? 1 : 0, 3);
        writer.align();
        writer.bits(chunk, 16);
        writer.bits(~chunk & 0xffff, 16);
        writer.bytes(this.data.subarray(offset, offset + chunk));
        offset += chunk;
      } while (remaining > 0);
    } else if (fixedBits <= dynamicBits) {
      writer.bits((final ? 1 : 0) | (1 << 1), 3);
      this.writeBody(FIXED_LIT_LENGTHS, FIXED_LIT_CODES, FIXED_DIST_LENGTHS, FIXED_DIST_CODES);
    } else {
      writer.bits((final ? 1 : 0) | (2 << 1), 3);
      writer.bits(hlit - 257, 5);
      writer.bits(hdist - 1, 5);
      writer.bits(hclen - 4, 4);
      for (let i = 0; i < hclen; i++) writer.bits(clLengths[CL_ORDER[i]], 3);
      const clCodes = reversedCodes(clLengths);
      for (let i = 0; i < clCount; i++) {
        const sym = clSymbols[i];
        writer.bits(clCodes[sym], clLengths[sym]);
        if (sym === 16) writer.bits(clExtra[i], 2);
        else if (sym === 17) writer.bits(clExtra[i], 3);
        else if (sym === 18) writer.bits(clExtra[i], 7);
      }
      this.writeBody(litLengths, reversedCodes(litLengths), distLengths, reversedCodes(distLengths));
    }

    this.start += this.length;
    this.length = 0;
    this.tokens = 0;
    litFreq.fill(0);
    distFreq.fill(0);
  }

  /** Size in bits of the buffered symbols (including end-of-block) under the given codes. */
  private bodyBits(litLengths: Uint8Array, distLengths: Uint8Array): number {
    const { litFreq, distFreq } = this;
    let bits = 0;
    for (let sym = 0; sym < NUM_LITLEN; sym++) bits += litFreq[sym] * litLengths[sym];
    for (let sym = 0; sym < LEN_EXTRA.length; sym++) bits += litFreq[257 + sym] * LEN_EXTRA[sym];
    for (let sym = 0; sym < NUM_DIST; sym++) bits += distFreq[sym] * (distLengths[sym] + DIST_EXTRA[sym]);
    return bits;
  }

  private writeBody(litLengths: Uint8Array, litCodes: Uint16Array, distLengths: Uint8Array, distCodes: Uint16Array): void {
    const { writer, tokenValue, tokenDist } = this;
    for (let i = 0; i < this.tokens; i++) {
      const value = tokenValue[i];
      const d = tokenDist[i];
      if (d === 0xffff) {
        writer.bits(litCodes[value], litLengths[value]);
        continue;
      }
      const lenSym = LEN_SYMBOL[value];
      writer.bits(litCodes[257 + lenSym], litLengths[257 + lenSym]);
      if (LEN_EXTRA[lenSym]) writer.bits(value - LEN_BASE[lenSym], LEN_EXTRA[lenSym]);
      const distSym = DIST_SYMBOL[d < 256 ? d : 256 + (d >> 7)];
      writer.bits(distCodes[distSym], distLengths[distSym]);
      if (DIST_EXTRA[distSym]) writer.bits(d + 1 - DIST_BASE[distSym], DIST_EXTRA[distSym]);
    }
    writer.bits(litCodes[END_OF_BLOCK], litLengths[END_OF_BLOCK]);
  }
}

/** Encodes data as a raw DEFLATE stream. */
export function deflateRaw(data: Uint8Array): Uint8Array {
  const n = data.length;
  const writer = new BitWriter((n >>> 1) + 64);
  const encoder = new BlockEncoder(data, writer);
  // Hash chains over three-byte sequences: `head` holds the latest position per hash,
  // `prev` links each position (modulo the window) to the previous one with the same hash.
  const head = new Int32Array(HASH_SIZE).fill(-1);
  const prev = new Int32Array(WINDOW_SIZE);

  // Lazy matching: the match found at a position is only emitted if the next position
  // does not yield a longer one; otherwise a single literal is emitted instead.
  let pos = 0;
  let matchLen = MIN_MATCH - 1;
  let matchDist = 0;
  let pending = false;
  while (pos < n) {
    let candidate = -1;
    if (pos + 2 < n) {
      const h = ((data[pos] << 10) ^ (data[pos + 1] << 5) ^ data[pos + 2]) & HASH_MASK;
      candidate = head[h];
      prev[pos & WINDOW_MASK] = candidate;
      head[h] = pos;
    }
    const prevLen = matchLen;
    const prevDist = matchDist;
    matchLen = MIN_MATCH - 1;

    const maxLen = Math.min(MAX_MATCH, n - pos);
    if (candidate >= 0 && prevLen < MAX_LAZY && prevLen < maxLen && pos - candidate <= WINDOW_SIZE) {
      // The chain length cap keeps degenerate inputs (long runs, many repeats) linear.
      let chain = prevLen >= GOOD_LENGTH ? MAX_CHAIN >> 2 : MAX_CHAIN;
      const nice = Math.min(NICE_LENGTH, maxLen);
      let best = prevLen;
      let bestDist = 0;
      let cur = candidate;
      for (;;) {
        if (data[cur + best] === data[pos + best] && data[cur] === data[pos]) {
          let len = 1;
          while (len < maxLen && data[cur + len] === data[pos + len]) len++;
          if (len > best) {
            best = len;
            bestDist = pos - cur;
            if (len >= nice) break;
          }
        }
        if (--chain === 0) break;
        // Slots of positions that left the window are reused, so a link may point forward; stop there.
        const next = prev[cur & WINDOW_MASK];
        if (next < 0 || next >= cur || pos - next > WINDOW_SIZE) break;
        cur = next;
      }
      if (bestDist !== 0 && !(best === MIN_MATCH && bestDist > TOO_FAR)) {
        matchLen = best;
        matchDist = bestDist;
      }
    }

    if (prevLen >= MIN_MATCH && matchLen <= prevLen) {
      // The match that started at pos - 1 wins; index the positions it covers.
      encoder.match(prevLen, prevDist);
      const next = pos + prevLen - 1;
      for (let p = pos + 1; p < next; p++) {
        if (p + 2 >= n) break;
        const h = ((data[p] << 10) ^ (data[p + 1] << 5) ^ data[p + 2]) & HASH_MASK;
        prev[p & WINDOW_MASK] = head[h];
        head[h] = p;
      }
      pos = next;
      pending = false;
      matchLen = MIN_MATCH - 1;
    } else if (pending) {
      encoder.literal(data[pos - 1]);
      pos++;
    } else {
      pending = true;
      pos++;
    }
  }
  if (pending) encoder.literal(data[n - 1]);
  encoder.flush(true);
  return writer.finish();
}
