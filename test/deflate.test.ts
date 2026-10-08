import assert from 'node:assert/strict';
import { test } from 'node:test';
import zlib from 'node:zlib';
import { buildCodeLengths, deflateRaw, inflateRaw } from '../src/deflate';

// --- helpers ---------------------------------------------------------------

/** Deterministic PRNG (mulberry32). */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomBytes(length: number, seed: number, alphabet = 256): Uint8Array {
  const rand = prng(seed);
  const out = new Uint8Array(length);
  for (let i = 0; i < length; i++) out[i] = Math.floor(rand() * alphabet);
  return out;
}

const utf8 = (text: string) => new TextEncoder().encode(text);

function xmlLike(targetBytes: number, seed: number): Uint8Array {
  const rand = prng(seed);
  const names = ['Ağrı', 'İstanbul', 'Şırnak', 'Çanakkale', 'Öğütçü', 'Üsküdar', 'Iğdır', 'gemi', 'yük', 'liman'];
  const parts: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sheetData>'];
  let size = 0;
  for (let row = 1; size < targetBytes; row++) {
    let line = `<row r="${row}" spans="1:6">`;
    for (let col = 0; col < 6; col++) {
      const ref = String.fromCharCode(65 + col) + row;
      if (rand() < 0.5) line += `<c r="${ref}" s="${Math.floor(rand() * 12)}"><v>${(rand() * 100000).toFixed(2)}</v></c>`;
      else line += `<c r="${ref}" t="inlineStr"><is><t>${names[Math.floor(rand() * names.length)]}</t></is></c>`;
    }
    line += '</row>';
    parts.push(line);
    size += line.length;
  }
  parts.push('</sheetData>');
  return utf8(parts.join('\n'));
}

function sameBytes(actual: Uint8Array, expected: Uint8Array, label = ''): void {
  // Buffer.equals is much faster than deepEqual on megabyte inputs.
  assert.equal(actual.length, expected.length, `${label} length`);
  assert.ok(Buffer.from(actual.buffer, actual.byteOffset, actual.length).equals(expected), `${label} content differs`);
}

/** Round trip through our encoder and both decoders. Returns the compressed bytes. */
function roundTrip(data: Uint8Array, label = ''): Uint8Array {
  const packed = deflateRaw(data);
  sameBytes(inflateRaw(packed), data, `${label} own inflate`);
  sameBytes(zlib.inflateRawSync(packed), data, `${label} zlib inflate`);
  return packed;
}

/** LSB-first bit writer for hand-built streams. */
class Bits {
  private bytes: number[] = [];
  private buf = 0;
  private cnt = 0;

  /** Writes `n` bits of `value`, least significant first (header fields, extra bits). */
  put(value: number, n: number): this {
    for (let i = 0; i < n; i++) {
      this.buf |= ((value >>> i) & 1) << this.cnt;
      if (++this.cnt === 8) this.flush();
    }
    return this;
  }

  /** Writes a Huffman code, most significant bit first. */
  code(value: number, n: number): this {
    for (let i = n - 1; i >= 0; i--) this.put((value >>> i) & 1, 1);
    return this;
  }

  private flush(): void {
    this.bytes.push(this.buf);
    this.buf = 0;
    this.cnt = 0;
  }

  done(): Uint8Array {
    if (this.cnt) this.flush();
    return new Uint8Array(this.bytes);
  }
}

/** Canonical Huffman codes for the given lengths. */
function canonical(lengths: number[]): number[] {
  const count = new Array<number>(17).fill(0);
  for (const len of lengths) if (len) count[len]++;
  const next = new Array<number>(17).fill(0);
  for (let len = 1; len <= 15; len++) next[len + 1] = (next[len] + count[len]) << 1;
  return lengths.map((len) => (len ? next[len]++ : 0));
}

/** Fixed-block literal/length symbol. */
function fixedSym(b: Bits, sym: number): Bits {
  if (sym < 144) return b.code(0x30 + sym, 8);
  if (sym < 256) return b.code(0x190 + sym - 144, 9);
  if (sym < 280) return b.code(sym - 256, 7);
  return b.code(0xc0 + sym - 280, 8);
}

const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
/** Code length code where symbols 0..15 all take four bits (complete). */
const CL_PLAIN = [4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, 0, 0];

interface DynamicCodes {
  lit: number[];
  litCodes: number[];
  dist: number[];
  distCodes: number[];
}

/**
 * Writes a final dynamic block header. `sequence` is the list of code length
 * symbols as `[symbol, extraValue]`; by default each length is written literally.
 */
function dynamicHeader(
  b: Bits,
  lit: number[],
  dist: number[],
  opts: { cl?: number[]; sequence?: [number, number][]; hlit?: number; hdist?: number } = {},
): DynamicCodes {
  const cl = opts.cl ?? CL_PLAIN;
  const clCodes = canonical(cl);
  b.put(1, 1).put(2, 2);
  b.put(opts.hlit ?? lit.length - 257, 5);
  b.put(opts.hdist ?? dist.length - 1, 5);
  b.put(19 - 4, 4);
  for (const sym of CL_ORDER) b.put(cl[sym], 3);
  const sequence = opts.sequence ?? [...lit, ...dist].map((len): [number, number] => [len, 0]);
  for (const [sym, extra] of sequence) {
    b.code(clCodes[sym], cl[sym]);
    if (sym === 16) b.put(extra, 2);
    else if (sym === 17) b.put(extra, 3);
    else if (sym === 18) b.put(extra, 7);
  }
  return { lit, litCodes: canonical(lit), dist, distCodes: canonical(dist) };
}

/** Literal/length lengths: the given symbols get the given lengths, 257 entries minimum. */
function litLengths(entries: Record<number, number>, size = 257): number[] {
  const lengths = new Array<number>(size).fill(0);
  for (const [sym, len] of Object.entries(entries)) lengths[Number(sym)] = len;
  return lengths;
}

/** Asserts that both our decoder and zlib reject the stream. */
function rejects(stream: Uint8Array, pattern: RegExp): void {
  assert.throws(() => inflateRaw(stream), pattern);
  assert.throws(() => zlib.inflateRawSync(stream), 'zlib accepts a stream we expected to be invalid');
}

// --- round trips -----------------------------------------------------------

test('deflate: empty input produces a valid stream', () => {
  const packed = roundTrip(new Uint8Array(0));
  assert.ok(packed.length > 0 && packed.length <= 5);
});

test('deflate: single byte', () => {
  for (const value of [0, 1, 0x7f, 0xff]) roundTrip(new Uint8Array([value]), `byte ${value}`);
});

test('deflate: one megabyte of the same byte', () => {
  const data = new Uint8Array(1 << 20).fill(0x41);
  const packed = roundTrip(data);
  assert.ok(packed.length < 2500, `expected a tiny stream, got ${packed.length} bytes`);
});

test('deflate: incompressible random data falls back to stored blocks', () => {
  const data = randomBytes(300_000, 1);
  const packed = roundTrip(data);
  assert.ok(packed.length <= data.length + 5 * Math.ceil(data.length / 32768) + 5, `stored overhead too high: ${packed.length}`);
});

test('deflate: repetitive text', () => {
  const data = utf8('The quick brown fox jumps over the lazy dog. '.repeat(5000));
  const packed = roundTrip(data);
  assert.ok(packed.length < data.length / 50);
});

test('deflate: UTF-8 XML with Turkish characters', () => {
  const data = xmlLike(50_000, 2);
  const packed = roundTrip(data);
  assert.ok(packed.length < data.length / 3);
  assert.ok(new TextDecoder().decode(inflateRaw(packed)).includes('Şırnak'));
});

test('deflate: sizes around the 65535-byte stored block limit', () => {
  for (const size of [65534, 65535, 65536, 65537, 131070, 131071]) {
    roundTrip(randomBytes(size, size), `random ${size}`);
    roundTrip(randomBytes(size, size, 4), `compressible ${size}`);
  }
});

test('deflate: matches of the maximum length 258', () => {
  const unit = randomBytes(258, 3);
  for (const extra of [0, 1, 2, 3, 257, 258, 259]) {
    const data = new Uint8Array(258 * 3 + extra);
    for (let i = 0; i < data.length; i++) data[i] = unit[i % 258];
    roundTrip(data, `extra ${extra}`);
  }
  for (const size of [257, 258, 259, 260, 516, 517, 1000]) roundTrip(new Uint8Array(size).fill(9), `run ${size}`);
});

test('deflate: matches at the 32768 distance limit', () => {
  for (const distance of [32766, 32767, 32768, 32769, 32770]) {
    const unit = randomBytes(64, distance);
    const filler = randomBytes(distance - 64, distance + 1);
    const data = new Uint8Array(distance + 64 + 10);
    data.set(unit, 0);
    data.set(filler, 64);
    data.set(unit, distance);
    const packed = roundTrip(data, `distance ${distance}`);
    // Everything but the repeat is incompressible, so the output is smaller than the input
    // exactly when the repeat was used: reachable up to and including 32768, not beyond.
    assert.equal(packed.length < data.length, distance <= 32768, `distance ${distance}: ${packed.length} bytes`);
  }
});

test('deflate: encoder is deterministic', () => {
  const data = xmlLike(30_000, 4);
  assert.deepEqual(deflateRaw(data), deflateRaw(data.slice()));
});

test('deflate: works on a subarray view with a byte offset', () => {
  const backing = randomBytes(5000, 5, 8);
  const view = backing.subarray(100, 4100);
  const packed = roundTrip(view);
  const padded = new Uint8Array(packed.length + 7);
  padded.set(packed, 7);
  sameBytes(inflateRaw(padded.subarray(7)), view);
});

// --- cross validation ------------------------------------------------------

test('inflate: decodes zlib output at every level and strategy', () => {
  const inputs: [string, Uint8Array][] = [
    ['empty', new Uint8Array(0)],
    ['xml', xmlLike(120_000, 6)],
    ['random', randomBytes(100_000, 7)],
    ['small alphabet', randomBytes(100_000, 8, 3)],
    ['run', new Uint8Array(200_000).fill(0xee)],
  ];
  for (const [name, data] of inputs) {
    for (const level of [0, 1, 6, 9]) {
      sameBytes(inflateRaw(zlib.deflateRawSync(data, { level })), data, `${name} level ${level}`);
    }
    for (const strategy of [zlib.constants.Z_FIXED, zlib.constants.Z_HUFFMAN_ONLY, zlib.constants.Z_RLE]) {
      sameBytes(inflateRaw(zlib.deflateRawSync(data, { strategy })), data, `${name} strategy ${strategy}`);
    }
    // Tiny blocks exercise many dynamic headers in a row.
    sameBytes(inflateRaw(zlib.deflateRawSync(data, { memLevel: 1 })), data, `${name} memLevel 1`);
  }
});

test('inflate: bytes after the final block are ignored', () => {
  const data = utf8('trailing garbage is fine');
  const packed = zlib.deflateRawSync(data);
  const padded = new Uint8Array(packed.length + 16).fill(0xff);
  padded.set(packed);
  sameBytes(inflateRaw(padded), data);
});

test('deflate: compression ratio on XML is close to zlib level 6', () => {
  const data = xmlLike(200_000, 9);
  const ours = deflateRaw(data).length;
  const theirs = zlib.deflateRawSync(data, { level: 6 }).length;
  assert.ok(ours <= theirs * 1.25, `ours ${ours} vs zlib ${theirs}`);
});

test('fuzz: random inputs round trip in both directions', () => {
  const rand = prng(0xbeef);
  for (let i = 0; i < 400; i++) {
    const length = Math.floor(rand() ** 3 * 20_000);
    const alphabet = 1 + Math.floor(rand() ** 2 * 256);
    const data = randomBytes(length, 1000 + i, alphabet);
    // Splice in copies of earlier content so long matches occur too.
    if (length > 100 && rand() < 0.5) {
      for (let k = 0; k < 5; k++) {
        const from = Math.floor(rand() * (length - 50));
        const to = Math.floor(rand() * (length - 50));
        data.copyWithin(to, from, from + Math.floor(rand() * 50));
      }
    }
    roundTrip(data, `case ${i} (length ${length}, alphabet ${alphabet})`);
    sameBytes(inflateRaw(zlib.deflateRawSync(data, { level: 1 + (i % 9) })), data, `case ${i} zlib`);
  }
});

// --- code length limiting --------------------------------------------------

function kraft(lengths: Uint8Array): number {
  let sum = 0;
  for (const len of lengths) if (len) sum += 2 ** -len;
  return sum;
}

test('deflate: Huffman code lengths stay within the limit for skewed frequencies', () => {
  for (const [symbols, limit, size] of [
    [40, 15, 286],
    [19, 7, 19],
    [30, 15, 30],
    [286, 15, 286],
    [128, 7, 128],
  ]) {
    const freq = new Uint32Array(size);
    let a = 1;
    let b = 1;
    for (let i = 0; i < symbols; i++) {
      // Fibonacci frequencies give the deepest possible Huffman tree; cap them to stay in 32 bits.
      freq[(i * 7) % size] = Math.min(a, 0xffffffff);
      [a, b] = [b, Math.min(a + b, 0xffffffff)];
    }
    const lengths = buildCodeLengths(freq, limit);
    const used = [...freq.keys()].filter((i) => freq[i] > 0);
    assert.ok(used.every((i) => lengths[i] >= 1 && lengths[i] <= limit), `${symbols}/${limit}: length out of range`);
    assert.ok([...freq.keys()].every((i) => freq[i] > 0 || lengths[i] === 0), 'unused symbol got a code');
    assert.equal(kraft(lengths), 1, `${symbols}/${limit}: code is not complete`);
    assert.equal(Math.max(...lengths), limit, `${symbols}/${limit}: the limiter was expected to kick in`);
    for (const i of used) for (const j of used) {
      if (freq[i] > freq[j]) assert.ok(lengths[i] <= lengths[j], 'a more frequent symbol has a longer code');
    }
  }
});

test('deflate: Huffman code lengths for degenerate and ordinary frequencies', () => {
  assert.deepEqual([...buildCodeLengths(new Uint32Array(5), 15)], [1, 1, 0, 0, 0]);
  assert.deepEqual([...buildCodeLengths(new Uint32Array([0, 0, 9, 0]), 15)], [1, 0, 1, 0]);
  assert.deepEqual([...buildCodeLengths(new Uint32Array([4, 0, 0]), 15)], [1, 1, 0]);
  assert.deepEqual([...buildCodeLengths(new Uint32Array([5, 5]), 15)], [1, 1]);
  assert.deepEqual([...buildCodeLengths(new Uint32Array([1, 1, 2, 4]), 15)], [3, 3, 2, 1]);
  assert.deepEqual([...buildCodeLengths(new Uint32Array([1, 1, 2, 4]), 2)], [2, 2, 2, 2]);
});

test('deflate: Fibonacci-distributed input round trips', () => {
  // Byte i occurs fib(i) times, shuffled, so that per-block trees are much deeper than 15 levels.
  const counts: number[] = [];
  let a = 1;
  let b = 1;
  for (let i = 0; i < 29; i++) {
    counts.push(a);
    [a, b] = [b, a + b];
  }
  const data = new Uint8Array(counts.reduce((x, y) => x + y, 0));
  let at = 0;
  counts.forEach((count, sym) => {
    data.fill(sym, at, at + count);
    at += count;
  });
  const rand = prng(11);
  for (let i = data.length - 1; i > 0; i--) {
    const j = Math.floor(rand() * (i + 1));
    [data[i], data[j]] = [data[j], data[i]];
  }
  roundTrip(data);
  // Unshuffled: long runs with Fibonacci lengths skew the length/distance and code length alphabets.
  roundTrip(data.slice().sort());
});

// --- malformed input -------------------------------------------------------

test('inflate: rejects truncated streams', () => {
  rejects(new Uint8Array(0), /unexpected end/);
  rejects(new Bits().put(1, 1).put(1, 2).done().subarray(0, 0), /unexpected end/);
  // Fixed block with one literal and no end-of-block symbol.
  rejects(fixedSym(new Bits().put(1, 1).put(1, 2), 65).done(), /unexpected end/);
  // Stored block whose header or payload is cut short.
  rejects(new Uint8Array([0x01]), /unexpected end/);
  rejects(new Uint8Array([0x01, 0x03, 0x00]), /unexpected end/);
  rejects(new Uint8Array([0x01, 0x03, 0x00, 0xfc, 0xff, 1, 2]), /unexpected end/);
  // A non-final block followed by nothing.
  rejects(new Uint8Array([0x00, 0x01, 0x00, 0xfe, 0xff, 7]), /unexpected end/);
  // Dynamic header cut in the middle.
  const b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 256: 1 }), [0]);
  rejects(b.done().subarray(0, 20), /unexpected end/);
});

test('inflate: rejects block type 3', () => {
  rejects(new Bits().put(1, 1).put(3, 2).done(), /invalid block type/);
});

test('inflate: rejects a stored block whose LEN and NLEN disagree', () => {
  rejects(new Uint8Array([0x01, 0x05, 0x00, 0x00, 0x00, 1, 2, 3, 4, 5]), /stored block length/);
  sameBytes(inflateRaw(new Uint8Array([0x01, 0x02, 0x00, 0xfd, 0xff, 8, 9])), new Uint8Array([8, 9]));
});

test('inflate: rejects over-subscribed code lengths', () => {
  // Code length code: three one-bit codes.
  let b = new Bits();
  const cl = new Array<number>(19).fill(0);
  cl[0] = cl[1] = cl[2] = 1;
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl, sequence: [] });
  rejects(b.done(), /over-subscribed code length code/);

  // Literal/length code: three one-bit codes.
  b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 66: 1, 256: 1 }), [0]);
  rejects(b.put(0, 16).done(), /over-subscribed literal\/length code/);

  // Distance code: three one-bit codes.
  b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 256: 1 }), [1, 1, 1]);
  rejects(b.put(0, 16).done(), /over-subscribed distance code/);
});

test('inflate: rejects incomplete code lengths', () => {
  // Code length code with a single one-bit code.
  let b = new Bits();
  const cl = new Array<number>(19).fill(0);
  cl[1] = 1;
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl, sequence: [] });
  rejects(b.put(0, 32).done(), /incomplete code length code/);

  // Code length code with no codes at all.
  b = new Bits();
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl: new Array<number>(19).fill(0), sequence: [] });
  rejects(b.put(0, 400).done(), /code length code|end-of-block/);

  // Literal/length code: lengths 1 and 2 leave a quarter of the code space unused.
  b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 256: 2 }), [0]);
  rejects(b.put(0, 16).done(), /incomplete literal\/length code/);

  // Distance code: a single two-bit code.
  b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 256: 1 }), [2]);
  rejects(b.put(0, 16).done(), /incomplete distance code/);
});

test('inflate: accepts the incomplete codes zlib accepts', () => {
  // A single one-bit distance code, used by a match.
  let b = new Bits();
  let codes = dynamicHeader(b, litLengths({ 65: 2, 256: 2, 257: 1 }, 258), [1]);
  b.code(codes.litCodes[65], 2); // 'A'
  b.code(codes.litCodes[257], 1).code(0, 1); // length 3, distance 1
  b.code(codes.litCodes[256], 2);
  let stream = b.done();
  sameBytes(inflateRaw(stream), utf8('AAAA'));
  sameBytes(zlib.inflateRawSync(stream), utf8('AAAA'));

  // The unused half of that distance code is an error when it shows up.
  b = new Bits();
  codes = dynamicHeader(b, litLengths({ 65: 2, 256: 2, 257: 1 }, 258), [1]);
  b.code(codes.litCodes[65], 2).code(codes.litCodes[257], 1).code(1, 1).code(codes.litCodes[256], 2);
  rejects(b.put(0, 16).done(), /invalid distance code/);

  // No distance codes at all: fine as long as the block has only literals.
  b = new Bits();
  codes = dynamicHeader(b, litLengths({ 65: 1, 256: 1 }), [0]);
  b.code(codes.litCodes[65], 1).code(codes.litCodes[65], 1).code(codes.litCodes[256], 1);
  stream = b.done();
  sameBytes(inflateRaw(stream), utf8('AA'));
  sameBytes(zlib.inflateRawSync(stream), utf8('AA'));

  // ...and an error as soon as a length symbol needs a distance.
  b = new Bits();
  codes = dynamicHeader(b, litLengths({ 65: 2, 256: 2, 257: 1 }, 258), [0]);
  b.code(codes.litCodes[65], 2).code(codes.litCodes[257], 1);
  rejects(b.put(0, 32).done(), /invalid distance code/);

  // A single one-bit literal/length code (end-of-block only): an empty block.
  b = new Bits();
  codes = dynamicHeader(b, litLengths({ 256: 1 }), [0]);
  b.code(codes.litCodes[256], 1);
  stream = b.done();
  assert.equal(inflateRaw(stream).length, 0);
  assert.equal(zlib.inflateRawSync(stream).length, 0);
});

test('inflate: rejects a block without an end-of-block code', () => {
  const b = new Bits();
  dynamicHeader(b, litLengths({ 65: 1, 66: 1 }), [0]);
  rejects(b.put(0, 16).done(), /missing end-of-block/);
});

test('inflate: rejects a distance that points before the output start', () => {
  // Fixed block: length 3 at distance 1 with nothing written yet.
  let b = new Bits().put(1, 1).put(1, 2);
  fixedSym(b, 257).code(0, 5);
  rejects(fixedSym(b, 256).done(), /distance points before/);

  // One literal written, distance 2.
  b = new Bits().put(1, 1).put(1, 2);
  fixedSym(fixedSym(b, 65), 257).code(1, 5);
  rejects(fixedSym(b, 256).done(), /distance points before/);

  // Same with distance 1: valid.
  b = new Bits().put(1, 1).put(1, 2);
  fixedSym(fixedSym(b, 65), 257).code(0, 5);
  sameBytes(inflateRaw(fixedSym(b, 256).done()), utf8('AAAA'));
});

test('inflate: rejects length symbols 286 and 287', () => {
  for (const sym of [286, 287]) {
    const b = new Bits().put(1, 1).put(1, 2);
    fixedSym(fixedSym(b, 65), sym).put(0, 32);
    rejects(b.done(), /invalid length symbol/);
  }
});

test('inflate: rejects distance symbols 30 and 31', () => {
  for (const sym of [30, 31]) {
    const b = new Bits().put(1, 1).put(1, 2);
    fixedSym(fixedSym(b, 65), 257).code(sym, 5).put(0, 32);
    rejects(b.done(), /invalid distance symbol/);
  }
});

test('inflate: rejects a code length repeat with no previous length', () => {
  const cl = [4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, 0, 4, 0, 4];
  const b = new Bits();
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl, sequence: [[16, 0], [1, 0]] });
  rejects(b.put(0, 32).done(), /repeat with no previous length/);
});

test('inflate: rejects code length repeats that run past HLIT + HDIST', () => {
  const cl = [4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, 0, 4, 0, 4];
  // 257 + 1 lengths expected; 138 + 138 zeros overflow.
  let b = new Bits();
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl, sequence: [[18, 127], [18, 127]] });
  rejects(b.put(0, 32).done(), /repeat runs past/);

  // Same with symbol 16: 256 literal lengths, then a length and a repeat of six.
  b = new Bits();
  const sequence: [number, number][] = [[18, 127], [18, 107], [1, 0], [16, 3]];
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { cl, sequence });
  rejects(b.put(0, 32).done(), /repeat runs past/);
});

test('inflate: rejects HLIT above 286 and HDIST above 30', () => {
  let b = new Bits();
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { hlit: 30 });
  rejects(b.put(0, 64).done(), /too many length or distance symbols/);

  b = new Bits();
  dynamicHeader(b, litLengths({ 256: 1 }), [0], { hdist: 30 });
  rejects(b.put(0, 64).done(), /too many length or distance symbols/);
});

test('inflate: enforces maxOutput without trusting it for allocation', () => {
  const bomb = zlib.deflateRawSync(new Uint8Array(64 << 20), { level: 9 });
  assert.ok(bomb.length < 100_000);
  const started = performance.now();
  assert.throws(() => inflateRaw(bomb, 1 << 20), /exceeds the limit of 1048576 bytes/);
  assert.ok(performance.now() - started < 2000, 'the bomb must be stopped early');

  const data = utf8('hello hello hello hello');
  const packed = zlib.deflateRawSync(data);
  sameBytes(inflateRaw(packed, data.length), data);
  assert.throws(() => inflateRaw(packed, data.length - 1), /exceeds the limit/);
  assert.throws(() => inflateRaw(packed, 0), /exceeds the limit/);
  // A huge declared limit must not be allocated up front.
  sameBytes(inflateRaw(packed, 2 ** 40), data);
  sameBytes(inflateRaw(packed, Infinity), data);
  // Stored blocks are limited too.
  const stored = zlib.deflateRawSync(randomBytes(1000, 12), { level: 0 });
  assert.throws(() => inflateRaw(stored, 999), /exceeds the limit/);
  assert.equal(inflateRaw(stored, 1000).length, 1000);
  assert.equal(inflateRaw(deflateRaw(new Uint8Array(0)), 0).length, 0);
  assert.throws(() => inflateRaw(packed, -1), RangeError);
  assert.throws(() => inflateRaw(packed, NaN), RangeError);
});

test('inflate: every prefix of a valid stream fails cleanly', () => {
  const data = xmlLike(6000, 13);
  const streams = [
    deflateRaw(data),
    zlib.deflateRawSync(data, { level: 9 }),
    zlib.deflateRawSync(data, { strategy: zlib.constants.Z_FIXED }),
    zlib.deflateRawSync(data.subarray(0, 700), { level: 0 }),
  ];
  for (const stream of streams) {
    sameBytes(inflateRaw(stream), data.subarray(0, inflateRaw(stream).length));
    for (let cut = 0; cut < stream.length; cut++) {
      assert.throws(() => inflateRaw(stream.subarray(0, cut), 1 << 20), /Invalid DEFLATE data/, `prefix ${cut}`);
    }
  }
});

test('inflate: random corruption either fails cleanly or terminates', () => {
  const rand = prng(0xc0ffee);
  const data = xmlLike(20_000, 14);
  const streams = [deflateRaw(data), zlib.deflateRawSync(data, { level: 6 }), zlib.deflateRawSync(data, { memLevel: 1 })];
  let failed = 0;
  let survived = 0;
  for (let i = 0; i < 600; i++) {
    const stream = streams[i % streams.length].slice();
    const flips = 1 + Math.floor(rand() * 4);
    for (let k = 0; k < flips; k++) stream[Math.floor(rand() * stream.length)] ^= 1 << Math.floor(rand() * 8);
    try {
      const out = inflateRaw(stream, 1 << 22);
      assert.ok(out.length <= 1 << 22);
      survived++;
    } catch (error) {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Invalid DEFLATE data|exceeds the limit/);
      failed++;
    }
  }
  assert.equal(failed + survived, 600);
  assert.ok(failed > 0);

  // Pure noise as input.
  for (let i = 0; i < 300; i++) {
    const noise = randomBytes(1 + Math.floor(rand() * 2000), 5000 + i);
    try {
      assert.ok(inflateRaw(noise, 1 << 22).length <= 1 << 22);
    } catch (error) {
      assert.match((error as Error).message, /Invalid DEFLATE data|exceeds the limit/);
    }
  }
});

// --- speed -----------------------------------------------------------------

test('speed: inflate and deflate stay within generous bounds', () => {
  const xml = xmlLike(10 << 20, 15);
  const packedByZlib = zlib.deflateRawSync(xml);
  let started = performance.now();
  const unpacked = inflateRaw(packedByZlib);
  const inflateMs = performance.now() - started;
  sameBytes(unpacked, xml);
  assert.ok(inflateMs < 3000, `inflating 10 MB took ${inflateMs.toFixed(0)} ms`);

  const five = xml.subarray(0, 5 << 20);
  started = performance.now();
  const packed = deflateRaw(five);
  const deflateMs = performance.now() - started;
  sameBytes(zlib.inflateRawSync(packed), five);
  assert.ok(deflateMs < 15000, `deflating 5 MB took ${deflateMs.toFixed(0)} ms`);

  // Degenerate inputs must not go quadratic.
  const worst: Uint8Array[] = [new Uint8Array(5 << 20), new Uint8Array(5 << 20), randomBytes(5 << 20, 16, 2)];
  for (let i = 0; i < worst[1].length; i++) worst[1][i] = (i % 300) & 0xff; // long-period repeats
  for (const data of worst) {
    started = performance.now();
    const out = deflateRaw(data);
    const ms = performance.now() - started;
    sameBytes(inflateRaw(out), data);
    assert.ok(ms < 150000, `degenerate input took ${ms.toFixed(0)} ms`);
  }
});
