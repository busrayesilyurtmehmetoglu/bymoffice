// The digests password-protected Office files are made with, written out
// here because the engine has no dependencies and has to run where there is
// no WebCrypto (and WebCrypto has no MD5): MD5 (RFC 1321), SHA-1, SHA-256,
// SHA-384 and SHA-512 (FIPS 180-4), and HMAC (RFC 2104) over any of them.
//
// A hash object is reused: reset() starts it over, so that the hundred
// thousand rounds of a key derivation allocate nothing.

export type HashName = 'MD5' | 'SHA1' | 'SHA256' | 'SHA384' | 'SHA512';

export const HASH_NAMES: readonly HashName[] = ['MD5', 'SHA1', 'SHA256', 'SHA384', 'SHA512'];

export interface Hash {
  /** Bytes in a block of the compression function: 64, or 128 for SHA-384 and SHA-512. */
  readonly blockSize: number;
  /** Bytes in a digest. */
  readonly digestSize: number;
  update(data: Uint8Array): this;
  /** The digest, as a new array or written into `out`; the hash is then where reset() leaves it. */
  digest(out?: Uint8Array): Uint8Array;
  reset(): this;
  /** Zeroes the state: for a hash that took in a password or a key. */
  wipe(): void;
}

/** What the five digests share: a block buffer, a byte count and the padding. */
abstract class BlockHash implements Hash {
  protected readonly block: Uint8Array;
  private filled = 0;
  /** Bytes taken in so far, as two numbers: a file can be longer than 2^32 bytes, and 2^53 is never reached. */
  private bytes = 0;

  constructor(
    readonly blockSize: number,
    readonly digestSize: number,
    private readonly bigEndianLength: boolean,
  ) {
    this.block = new Uint8Array(blockSize);
  }

  protected abstract init(): void;
  protected abstract compress(data: Uint8Array, offset: number): void;
  protected abstract write(out: Uint8Array): void;
  protected abstract wipeState(): void;

  reset(): this {
    this.filled = 0;
    this.bytes = 0;
    this.init();
    return this;
  }

  update(data: Uint8Array): this {
    const size = this.blockSize;
    let at = 0;
    this.bytes += data.length;
    if (this.filled > 0) {
      const take = Math.min(size - this.filled, data.length);
      this.block.set(data.subarray(0, take), this.filled);
      this.filled += take;
      at = take;
      if (this.filled < size) return this;
      this.compress(this.block, 0);
      this.filled = 0;
    }
    for (; at + size <= data.length; at += size) this.compress(data, at);
    if (at < data.length) {
      this.block.set(data.subarray(at), 0);
      this.filled = data.length - at;
    }
    return this;
  }

  digest(out?: Uint8Array): Uint8Array {
    const size = this.blockSize;
    const block = this.block;
    // The length field: 8 bytes, or 16 for the 128-byte block (whose upper 8 are always zero here).
    const lengthSize = size === 128 ? 16 : 8;
    block[this.filled++] = 0x80;
    if (this.filled > size - lengthSize) {
      block.fill(0, this.filled);
      this.compress(block, 0);
      this.filled = 0;
    }
    block.fill(0, this.filled);
    const high = Math.floor(this.bytes / 0x20000000);
    const low = (this.bytes << 3) >>> 0;
    if (this.bigEndianLength) {
      for (let i = 0; i < 4; i++) {
        block[size - 8 + i] = (high >>> (24 - i * 8)) & 0xff;
        block[size - 4 + i] = (low >>> (24 - i * 8)) & 0xff;
      }
    } else {
      for (let i = 0; i < 4; i++) {
        block[size - 8 + i] = (low >>> (i * 8)) & 0xff;
        block[size - 4 + i] = (high >>> (i * 8)) & 0xff;
      }
    }
    this.compress(block, 0);
    const result = out || new Uint8Array(this.digestSize);
    this.write(result);
    block.fill(0);
    this.reset();
    return result;
  }

  wipe(): void {
    this.block.fill(0);
    this.filled = 0;
    this.bytes = 0;
    this.wipeState();
  }
}

const rotl = (x: number, n: number): number => (x << n) | (x >>> (32 - n));

// MD5: the shift of each step, and the constants floor(2^32 * abs(sin(i + 1))).
const MD5_SHIFT = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
const MD5_K = new Int32Array(64);
for (let i = 0; i < 64; i++) MD5_K[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296) | 0;

class Md5 extends BlockHash {
  private readonly h = new Int32Array(4);
  private readonly w = new Int32Array(16);

  constructor() {
    super(64, 16, false);
    this.init();
  }

  protected init(): void {
    const h = this.h;
    h[0] = 0x67452301;
    h[1] = 0xefcdab89 | 0;
    h[2] = 0x98badcfe | 0;
    h[3] = 0x10325476;
  }

  protected compress(data: Uint8Array, offset: number): void {
    const w = this.w;
    const h = this.h;
    for (let i = 0; i < 16; i++, offset += 4) w[i] = data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
    let a = h[0];
    let b = h[1];
    let c = h[2];
    let d = h[3];
    for (let i = 0; i < 64; i++) {
      let f: number;
      let g: number;
      const round = i >> 4;
      if (round === 0) {
        f = (b & c) | (~b & d);
        g = i;
      } else if (round === 1) {
        f = (d & b) | (~d & c);
        g = (5 * i + 1) & 15;
      } else if (round === 2) {
        f = b ^ c ^ d;
        g = (3 * i + 5) & 15;
      } else {
        f = c ^ (b | ~d);
        g = (7 * i) & 15;
      }
      const sum = (a + f + MD5_K[i] + w[g]) | 0;
      a = d;
      d = c;
      c = b;
      b = (b + rotl(sum, MD5_SHIFT[(round << 2) | (i & 3)])) | 0;
    }
    h[0] = (h[0] + a) | 0;
    h[1] = (h[1] + b) | 0;
    h[2] = (h[2] + c) | 0;
    h[3] = (h[3] + d) | 0;
  }

  protected write(out: Uint8Array): void {
    for (let i = 0; i < 4; i++) {
      const word = this.h[i];
      out[i * 4] = word & 0xff;
      out[i * 4 + 1] = (word >>> 8) & 0xff;
      out[i * 4 + 2] = (word >>> 16) & 0xff;
      out[i * 4 + 3] = word >>> 24;
    }
  }

  protected wipeState(): void {
    this.h.fill(0);
    this.w.fill(0);
  }
}

/** Writes 32-bit words big-endian, as many as fit into `out`. */
function writeWords(words: Int32Array, out: Uint8Array): void {
  for (let i = 0; i * 4 < out.length && i < words.length; i++) {
    const word = words[i];
    out[i * 4] = word >>> 24;
    out[i * 4 + 1] = (word >>> 16) & 0xff;
    out[i * 4 + 2] = (word >>> 8) & 0xff;
    out[i * 4 + 3] = word & 0xff;
  }
}

class Sha1 extends BlockHash {
  private readonly h = new Int32Array(5);
  private readonly w = new Int32Array(80);

  constructor() {
    super(64, 20, true);
    this.init();
  }

  protected init(): void {
    const h = this.h;
    h[0] = 0x67452301;
    h[1] = 0xefcdab89 | 0;
    h[2] = 0x98badcfe | 0;
    h[3] = 0x10325476;
    h[4] = 0xc3d2e1f0 | 0;
  }

  protected compress(data: Uint8Array, offset: number): void {
    const w = this.w;
    const h = this.h;
    for (let i = 0; i < 16; i++, offset += 4) w[i] = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
    for (let i = 16; i < 80; i++) w[i] = rotl(w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16], 1);
    let a = h[0];
    let b = h[1];
    let c = h[2];
    let d = h[3];
    let e = h[4];
    for (let i = 0; i < 80; i++) {
      let f: number;
      if (i < 20) f = ((b & c) | (~b & d)) + 0x5a827999;
      else if (i < 40) f = (b ^ c ^ d) + 0x6ed9eba1;
      else if (i < 60) f = ((b & c) | (b & d) | (c & d)) + 0x8f1bbcdc;
      else f = (b ^ c ^ d) + 0xca62c1d6;
      const t = (rotl(a, 5) + f + e + w[i]) | 0;
      e = d;
      d = c;
      c = rotl(b, 30);
      b = a;
      a = t;
    }
    h[0] = (h[0] + a) | 0;
    h[1] = (h[1] + b) | 0;
    h[2] = (h[2] + c) | 0;
    h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0;
  }

  protected write(out: Uint8Array): void {
    writeWords(this.h, out);
  }

  protected wipeState(): void {
    this.h.fill(0);
    this.w.fill(0);
  }
}

/** The first `count` primes: the constants of SHA-2 are fractions of their roots. */
function primes(count: number): number[] {
  const out: number[] = [];
  for (let n = 2; out.length < count; n++) {
    let prime = true;
    for (let i = 0; i < out.length && out[i] * out[i] <= n; i++) {
      if (n % out[i] === 0) {
        prime = false;
        break;
      }
    }
    if (prime) out.push(n);
  }
  return out;
}

// SHA-256: the first 32 bits of the fractional parts of the cube roots of the first 64 primes. A double holds them exactly.
const SHA256_K = new Int32Array(64);
{
  const p = primes(64);
  for (let i = 0; i < 64; i++) {
    const root = Math.cbrt(p[i]);
    SHA256_K[i] = Math.floor((root - Math.floor(root)) * 4294967296) | 0;
  }
}

class Sha256 extends BlockHash {
  private readonly h = new Int32Array(8);
  private readonly w = new Int32Array(64);

  constructor() {
    super(64, 32, true);
    this.init();
  }

  protected init(): void {
    const h = this.h;
    h[0] = 0x6a09e667;
    h[1] = 0xbb67ae85 | 0;
    h[2] = 0x3c6ef372;
    h[3] = 0xa54ff53a | 0;
    h[4] = 0x510e527f;
    h[5] = 0x9b05688c | 0;
    h[6] = 0x1f83d9ab;
    h[7] = 0x5be0cd19;
  }

  protected compress(data: Uint8Array, offset: number): void {
    const w = this.w;
    const h = this.h;
    for (let i = 0; i < 16; i++, offset += 4) w[i] = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
    for (let i = 16; i < 64; i++) {
      const x = w[i - 15];
      const y = w[i - 2];
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
    }
    let a = h[0];
    let b = h[1];
    let c = h[2];
    let d = h[3];
    let e = h[4];
    let f = h[5];
    let g = h[6];
    let hh = h[7];
    for (let i = 0; i < 64; i++) {
      const s1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const t1 = (hh + s1 + ((e & f) ^ (~e & g)) + SHA256_K[i] + w[i]) | 0;
      const s0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const t2 = (s0 + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      hh = g;
      g = f;
      f = e;
      e = (d + t1) | 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) | 0;
    }
    h[0] = (h[0] + a) | 0;
    h[1] = (h[1] + b) | 0;
    h[2] = (h[2] + c) | 0;
    h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0;
    h[5] = (h[5] + f) | 0;
    h[6] = (h[6] + g) | 0;
    h[7] = (h[7] + hh) | 0;
  }

  protected write(out: Uint8Array): void {
    writeWords(this.h, out);
  }

  protected wipeState(): void {
    this.h.fill(0);
    this.w.fill(0);
  }
}

// SHA-512: 64-bit constants, each as two 32-bit halves (high, low). A double cannot hold the
// fraction of a cube root to 64 bits, so these are written out (FIPS 180-4, 4.2.3).
// prettier-ignore
const SHA512_K = new Int32Array([
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019,
  0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118, 0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2,
  0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694, 0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3,
  0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65, 0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725,
  0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70, 0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df,
  0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b, 0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001,
  0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb,
  0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3, 0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec,
  0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b, 0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207,
  0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a,
  0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
]);

// prettier-ignore
const SHA512_INIT = new Int32Array([
  0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1,
  0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179,
]);
// prettier-ignore
const SHA384_INIT = new Int32Array([
  0xcbbb9d5d, 0xc1059ed8, 0x629a292a, 0x367cd507, 0x9159015a, 0x3070dd17, 0x152fecd8, 0xf70e5939,
  0x67332667, 0xffc00b31, 0x8eb44a87, 0x68581511, 0xdb0c2e0d, 0x64f98fa7, 0x47b5481d, 0xbefa4fa4,
]);

/** SHA-512 and SHA-384, with every 64-bit word as two 32-bit halves (ES2019 has no BigInt). */
class Sha512 extends BlockHash {
  /** Eight words: high half at 2i, low half at 2i + 1. */
  private readonly h = new Int32Array(16);
  private readonly w = new Int32Array(160);

  constructor(
    private readonly start: Int32Array,
    digestSize: number,
  ) {
    super(128, digestSize, true);
    this.init();
  }

  protected init(): void {
    this.h.set(this.start);
  }

  protected compress(data: Uint8Array, offset: number): void {
    const w = this.w;
    const h = this.h;
    for (let i = 0; i < 32; i++, offset += 4) w[i] = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
    for (let i = 32; i < 160; i += 2) {
      // s0 = rotr(x, 1) ^ rotr(x, 8) ^ (x >>> 7), x = w[t - 15]
      const xh = w[i - 30];
      const xl = w[i - 29];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      // s1 = rotr(y, 19) ^ rotr(y, 61) ^ (y >>> 6), y = w[t - 2]
      const yh = w[i - 4];
      const yl = w[i - 3];
      const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
      const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
      // The low halves are added as unsigned numbers; what goes over 2^32 is carried into the high half.
      const low = (w[i - 31] >>> 0) + (s0l >>> 0) + (w[i - 13] >>> 0) + (s1l >>> 0);
      w[i + 1] = low | 0;
      w[i] = (w[i - 32] + s0h + w[i - 14] + s1h + Math.floor(low / 4294967296)) | 0;
    }
    let ah = h[0];
    let al = h[1];
    let bh = h[2];
    let bl = h[3];
    let ch = h[4];
    let cl = h[5];
    let dh = h[6];
    let dl = h[7];
    let eh = h[8];
    let el = h[9];
    let fh = h[10];
    let fl = h[11];
    let gh = h[12];
    let gl = h[13];
    let hh = h[14];
    let hl = h[15];
    for (let i = 0; i < 160; i += 2) {
      // S1 = rotr(e, 14) ^ rotr(e, 18) ^ rotr(e, 41)
      const s1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const s1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh);
      const chl = (el & fl) ^ (~el & gl);
      const t1low = (hl >>> 0) + (s1l >>> 0) + (chl >>> 0) + (SHA512_K[i + 1] >>> 0) + (w[i + 1] >>> 0);
      const t1l = t1low | 0;
      const t1h = (hh + s1h + chh + SHA512_K[i] + w[i] + Math.floor(t1low / 4294967296)) | 0;
      // S0 = rotr(a, 28) ^ rotr(a, 34) ^ rotr(a, 39)
      const s0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const s0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const majh = (ah & bh) ^ (ah & ch) ^ (bh & ch);
      const majl = (al & bl) ^ (al & cl) ^ (bl & cl);
      const t2low = (s0l >>> 0) + (majl >>> 0);
      const t2l = t2low | 0;
      const t2h = (s0h + majh + (t2low > 0xffffffff ? 1 : 0)) | 0;
      hh = gh;
      hl = gl;
      gh = fh;
      gl = fl;
      fh = eh;
      fl = el;
      const elow = (dl >>> 0) + (t1l >>> 0);
      el = elow | 0;
      eh = (dh + t1h + (elow > 0xffffffff ? 1 : 0)) | 0;
      dh = ch;
      dl = cl;
      ch = bh;
      cl = bl;
      bh = ah;
      bl = al;
      const alow = (t1l >>> 0) + (t2l >>> 0);
      al = alow | 0;
      ah = (t1h + t2h + (alow > 0xffffffff ? 1 : 0)) | 0;
    }
    const add = (at: number, high: number, low: number): void => {
      const sum = (h[at + 1] >>> 0) + (low >>> 0);
      h[at + 1] = sum | 0;
      h[at] = (h[at] + high + (sum > 0xffffffff ? 1 : 0)) | 0;
    };
    add(0, ah, al);
    add(2, bh, bl);
    add(4, ch, cl);
    add(6, dh, dl);
    add(8, eh, el);
    add(10, fh, fl);
    add(12, gh, gl);
    add(14, hh, hl);
  }

  protected write(out: Uint8Array): void {
    writeWords(this.h, out);
  }

  protected wipeState(): void {
    this.h.fill(0);
    this.w.fill(0);
  }
}

export function createHash(name: HashName): Hash {
  switch (name) {
    case 'MD5':
      return new Md5();
    case 'SHA1':
      return new Sha1();
    case 'SHA256':
      return new Sha256();
    case 'SHA384':
      return new Sha512(SHA384_INIT, 48);
    case 'SHA512':
      return new Sha512(SHA512_INIT, 64);
    default:
      throw new Error('Unknown hash algorithm');
  }
}

/** Bytes in a digest of the algorithm. */
export function digestSizeOf(name: HashName): number {
  return name === 'MD5' ? 16 : name === 'SHA1' ? 20 : name === 'SHA256' ? 32 : name === 'SHA384' ? 48 : 64;
}

/** The digest of the parts laid end to end. */
export function digest(name: HashName, ...parts: Uint8Array[]): Uint8Array {
  const hash = createHash(name);
  for (let i = 0; i < parts.length; i++) hash.update(parts[i]);
  const out = hash.digest();
  hash.wipe();
  return out;
}

/** An HMAC (RFC 2104) that takes its message in pieces: a large message can be worked on in slices. */
export class Hmac {
  private readonly hash: Hash;
  private readonly pad: Uint8Array;

  constructor(name: HashName, key: Uint8Array) {
    const hash = (this.hash = createHash(name));
    const pad = (this.pad = new Uint8Array(hash.blockSize));
    // A key longer than a block is hashed first; one of exactly a block is used as it is.
    if (key.length > pad.length) hash.update(key).digest(pad.subarray(0, hash.digestSize));
    else pad.set(key);
    for (let i = 0; i < pad.length; i++) pad[i] ^= 0x36;
    hash.reset().update(pad);
  }

  update(data: Uint8Array): this {
    this.hash.update(data);
    return this;
  }

  /** The HMAC of what was taken in. The object cannot be used afterwards: its key is zeroed. */
  digest(): Uint8Array {
    const hash = this.hash;
    const pad = this.pad;
    const inner = hash.digest();
    // 0x36 ^ 0x5c: from the inner pad to the outer one.
    for (let i = 0; i < pad.length; i++) pad[i] ^= 0x6a;
    const out = hash.update(pad).update(inner).digest();
    inner.fill(0);
    this.wipe();
    return out;
  }

  wipe(): void {
    this.pad.fill(0);
    this.hash.wipe();
  }
}

/** HMAC of the parts laid end to end (RFC 2104). */
export function hmac(name: HashName, key: Uint8Array, ...parts: Uint8Array[]): Uint8Array {
  const mac = new Hmac(name, key);
  for (let i = 0; i < parts.length; i++) mac.update(parts[i]);
  return mac.digest();
}
