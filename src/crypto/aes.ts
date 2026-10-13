// AES (FIPS 197) with keys of 128, 192 and 256 bits, in the two modes
// Office files use: ECB and CBC, both without padding (the formats pad by
// themselves). Written with the usual four lookup tables; the tables are
// computed when the module is first used, not stored.
//
// The lookups depend on the key and the data, so the time a block takes is
// not constant: fine for opening and saving a file on the user's own
// machine, not for a server that answers strangers.

let SBOX: Uint8Array | undefined;
let INV_SBOX: Uint8Array;
let TE0: Int32Array;
let TE1: Int32Array;
let TE2: Int32Array;
let TE3: Int32Array;
let TD0: Int32Array;
let TD1: Int32Array;
let TD2: Int32Array;
let TD3: Int32Array;

function tables(): void {
  if (SBOX) return;
  const sbox = new Uint8Array(256);
  const inv = new Uint8Array(256);
  // Powers of 3 and their inverses run through the whole field: p * q is 1 at every step.
  let p = 1;
  let q = 1;
  do {
    p = (p ^ (p << 1) ^ (p & 0x80 ? 0x11b : 0)) & 0xff;
    q ^= q << 1;
    q ^= q << 2;
    q ^= q << 4;
    q &= 0xff;
    if (q & 0x80) q ^= 0x09;
    const x = q ^ ((q << 1) | (q >> 7)) ^ ((q << 2) | (q >> 6)) ^ ((q << 3) | (q >> 5)) ^ ((q << 4) | (q >> 4));
    sbox[p] = (x ^ 0x63) & 0xff;
  } while (p !== 1);
  sbox[0] = 0x63;
  for (let i = 0; i < 256; i++) inv[sbox[i]] = i;

  const mul = (a: number, b: number): number => {
    let out = 0;
    for (; b; b >>= 1) {
      if (b & 1) out ^= a;
      a = (a << 1) ^ (a & 0x80 ? 0x11b : 0);
    }
    return out & 0xff;
  };
  TE0 = new Int32Array(256);
  TE1 = new Int32Array(256);
  TE2 = new Int32Array(256);
  TE3 = new Int32Array(256);
  TD0 = new Int32Array(256);
  TD1 = new Int32Array(256);
  TD2 = new Int32Array(256);
  TD3 = new Int32Array(256);
  for (let i = 0; i < 256; i++) {
    const s = sbox[i];
    const e = (mul(s, 2) << 24) | (s << 16) | (s << 8) | mul(s, 3);
    TE0[i] = e;
    TE1[i] = (e >>> 8) | (e << 24);
    TE2[i] = (e >>> 16) | (e << 16);
    TE3[i] = (e >>> 24) | (e << 8);
    const v = inv[i];
    const d = (mul(v, 14) << 24) | (mul(v, 9) << 16) | (mul(v, 13) << 8) | mul(v, 11);
    TD0[i] = d;
    TD1[i] = (d >>> 8) | (d << 24);
    TD2[i] = (d >>> 16) | (d << 16);
    TD3[i] = (d >>> 24) | (d << 8);
  }
  INV_SBOX = inv;
  SBOX = sbox;
}

/** One key, expanded for both directions. */
export class Aes {
  private readonly rounds: number;
  private readonly enc: Int32Array;
  private readonly dec: Int32Array;

  constructor(key: Uint8Array) {
    if (key.length !== 16 && key.length !== 24 && key.length !== 32) throw new Error('AES key must be 16, 24 or 32 bytes');
    tables();
    const sbox = SBOX as Uint8Array;
    const nk = key.length / 4;
    const rounds = (this.rounds = nk + 6);
    const size = 4 * (rounds + 1);
    const enc = (this.enc = new Int32Array(size));
    for (let i = 0; i < nk; i++) enc[i] = (key[i * 4] << 24) | (key[i * 4 + 1] << 16) | (key[i * 4 + 2] << 8) | key[i * 4 + 3];
    let rcon = 1;
    for (let i = nk; i < size; i++) {
      let t = enc[i - 1];
      if (i % nk === 0) {
        t = (sbox[(t >>> 16) & 0xff] << 24) | (sbox[(t >>> 8) & 0xff] << 16) | (sbox[t & 0xff] << 8) | sbox[t >>> 24];
        t ^= rcon << 24;
        rcon = (rcon << 1) ^ (rcon & 0x80 ? 0x11b : 0);
      } else if (nk > 6 && i % nk === 4) {
        t = (sbox[t >>> 24] << 24) | (sbox[(t >>> 16) & 0xff] << 16) | (sbox[(t >>> 8) & 0xff] << 8) | sbox[t & 0xff];
      }
      enc[i] = enc[i - nk] ^ t;
    }
    // The equivalent inverse cipher: the round keys in reverse order, the middle ones through InvMixColumns.
    const dec = (this.dec = new Int32Array(size));
    for (let round = 0; round <= rounds; round++) {
      for (let c = 0; c < 4; c++) {
        const word = enc[(rounds - round) * 4 + c];
        dec[round * 4 + c] =
          round === 0 || round === rounds ? word : TD0[sbox[word >>> 24]] ^ TD1[sbox[(word >>> 16) & 0xff]] ^ TD2[sbox[(word >>> 8) & 0xff]] ^ TD3[sbox[word & 0xff]];
      }
    }
  }

  /** Zeroes the expanded key. The object cannot be used afterwards. */
  wipe(): void {
    this.enc.fill(0);
    this.dec.fill(0);
  }

  /** Encrypts the 16 bytes at `input[at]` into `output[to]`; the two may be the same place. */
  encryptBlock(input: Uint8Array, at: number, output: Uint8Array, to: number): void {
    const rk = this.enc;
    const sbox = SBOX as Uint8Array;
    let s0 = ((input[at] << 24) | (input[at + 1] << 16) | (input[at + 2] << 8) | input[at + 3]) ^ rk[0];
    let s1 = ((input[at + 4] << 24) | (input[at + 5] << 16) | (input[at + 6] << 8) | input[at + 7]) ^ rk[1];
    let s2 = ((input[at + 8] << 24) | (input[at + 9] << 16) | (input[at + 10] << 8) | input[at + 11]) ^ rk[2];
    let s3 = ((input[at + 12] << 24) | (input[at + 13] << 16) | (input[at + 14] << 8) | input[at + 15]) ^ rk[3];
    let k = 4;
    for (let round = 1; round < this.rounds; round++, k += 4) {
      const t0 = TE0[s0 >>> 24] ^ TE1[(s1 >>> 16) & 0xff] ^ TE2[(s2 >>> 8) & 0xff] ^ TE3[s3 & 0xff] ^ rk[k];
      const t1 = TE0[s1 >>> 24] ^ TE1[(s2 >>> 16) & 0xff] ^ TE2[(s3 >>> 8) & 0xff] ^ TE3[s0 & 0xff] ^ rk[k + 1];
      const t2 = TE0[s2 >>> 24] ^ TE1[(s3 >>> 16) & 0xff] ^ TE2[(s0 >>> 8) & 0xff] ^ TE3[s1 & 0xff] ^ rk[k + 2];
      const t3 = TE0[s3 >>> 24] ^ TE1[(s0 >>> 16) & 0xff] ^ TE2[(s1 >>> 8) & 0xff] ^ TE3[s2 & 0xff] ^ rk[k + 3];
      s0 = t0;
      s1 = t1;
      s2 = t2;
      s3 = t3;
    }
    const t0 = ((sbox[s0 >>> 24] << 24) | (sbox[(s1 >>> 16) & 0xff] << 16) | (sbox[(s2 >>> 8) & 0xff] << 8) | sbox[s3 & 0xff]) ^ rk[k];
    const t1 = ((sbox[s1 >>> 24] << 24) | (sbox[(s2 >>> 16) & 0xff] << 16) | (sbox[(s3 >>> 8) & 0xff] << 8) | sbox[s0 & 0xff]) ^ rk[k + 1];
    const t2 = ((sbox[s2 >>> 24] << 24) | (sbox[(s3 >>> 16) & 0xff] << 16) | (sbox[(s0 >>> 8) & 0xff] << 8) | sbox[s1 & 0xff]) ^ rk[k + 2];
    const t3 = ((sbox[s3 >>> 24] << 24) | (sbox[(s0 >>> 16) & 0xff] << 16) | (sbox[(s1 >>> 8) & 0xff] << 8) | sbox[s2 & 0xff]) ^ rk[k + 3];
    put(output, to, t0, t1, t2, t3);
  }

  /** Decrypts the 16 bytes at `input[at]` into `output[to]`; the two may be the same place. */
  decryptBlock(input: Uint8Array, at: number, output: Uint8Array, to: number): void {
    const rk = this.dec;
    const inv = INV_SBOX;
    let s0 = ((input[at] << 24) | (input[at + 1] << 16) | (input[at + 2] << 8) | input[at + 3]) ^ rk[0];
    let s1 = ((input[at + 4] << 24) | (input[at + 5] << 16) | (input[at + 6] << 8) | input[at + 7]) ^ rk[1];
    let s2 = ((input[at + 8] << 24) | (input[at + 9] << 16) | (input[at + 10] << 8) | input[at + 11]) ^ rk[2];
    let s3 = ((input[at + 12] << 24) | (input[at + 13] << 16) | (input[at + 14] << 8) | input[at + 15]) ^ rk[3];
    let k = 4;
    for (let round = 1; round < this.rounds; round++, k += 4) {
      const t0 = TD0[s0 >>> 24] ^ TD1[(s3 >>> 16) & 0xff] ^ TD2[(s2 >>> 8) & 0xff] ^ TD3[s1 & 0xff] ^ rk[k];
      const t1 = TD0[s1 >>> 24] ^ TD1[(s0 >>> 16) & 0xff] ^ TD2[(s3 >>> 8) & 0xff] ^ TD3[s2 & 0xff] ^ rk[k + 1];
      const t2 = TD0[s2 >>> 24] ^ TD1[(s1 >>> 16) & 0xff] ^ TD2[(s0 >>> 8) & 0xff] ^ TD3[s3 & 0xff] ^ rk[k + 2];
      const t3 = TD0[s3 >>> 24] ^ TD1[(s2 >>> 16) & 0xff] ^ TD2[(s1 >>> 8) & 0xff] ^ TD3[s0 & 0xff] ^ rk[k + 3];
      s0 = t0;
      s1 = t1;
      s2 = t2;
      s3 = t3;
    }
    const t0 = ((inv[s0 >>> 24] << 24) | (inv[(s3 >>> 16) & 0xff] << 16) | (inv[(s2 >>> 8) & 0xff] << 8) | inv[s1 & 0xff]) ^ rk[k];
    const t1 = ((inv[s1 >>> 24] << 24) | (inv[(s0 >>> 16) & 0xff] << 16) | (inv[(s3 >>> 8) & 0xff] << 8) | inv[s2 & 0xff]) ^ rk[k + 1];
    const t2 = ((inv[s2 >>> 24] << 24) | (inv[(s1 >>> 16) & 0xff] << 16) | (inv[(s0 >>> 8) & 0xff] << 8) | inv[s3 & 0xff]) ^ rk[k + 2];
    const t3 = ((inv[s3 >>> 24] << 24) | (inv[(s2 >>> 16) & 0xff] << 16) | (inv[(s1 >>> 8) & 0xff] << 8) | inv[s0 & 0xff]) ^ rk[k + 3];
    put(output, to, t0, t1, t2, t3);
  }

  /** ECB over whole blocks of `data`, into `out` (which may be `data`). */
  ecb(encrypt: boolean, data: Uint8Array, out: Uint8Array): void {
    whole(data);
    for (let at = 0; at < data.length; at += 16) {
      if (encrypt) this.encryptBlock(data, at, out, at);
      else this.decryptBlock(data, at, out, at);
    }
  }

  /** CBC over whole blocks of `data`, into `out` (which may be `data`). `iv` is not changed. */
  cbc(encrypt: boolean, iv: Uint8Array, data: Uint8Array, out: Uint8Array): void {
    whole(data);
    if (iv.length !== 16) throw new Error('AES initialisation vector must be 16 bytes');
    const chain = iv.slice();
    const block = new Uint8Array(16);
    for (let at = 0; at < data.length; at += 16) {
      if (encrypt) {
        for (let i = 0; i < 16; i++) block[i] = data[at + i] ^ chain[i];
        this.encryptBlock(block, 0, out, at);
        chain.set(out.subarray(at, at + 16));
      } else {
        block.set(data.subarray(at, at + 16));
        this.decryptBlock(data, at, out, at);
        for (let i = 0; i < 16; i++) out[at + i] ^= chain[i];
        chain.set(block);
      }
    }
    block.fill(0);
    chain.fill(0);
  }
}

function put(out: Uint8Array, to: number, t0: number, t1: number, t2: number, t3: number): void {
  out[to] = t0 >>> 24;
  out[to + 1] = (t0 >>> 16) & 0xff;
  out[to + 2] = (t0 >>> 8) & 0xff;
  out[to + 3] = t0 & 0xff;
  out[to + 4] = t1 >>> 24;
  out[to + 5] = (t1 >>> 16) & 0xff;
  out[to + 6] = (t1 >>> 8) & 0xff;
  out[to + 7] = t1 & 0xff;
  out[to + 8] = t2 >>> 24;
  out[to + 9] = (t2 >>> 16) & 0xff;
  out[to + 10] = (t2 >>> 8) & 0xff;
  out[to + 11] = t2 & 0xff;
  out[to + 12] = t3 >>> 24;
  out[to + 13] = (t3 >>> 16) & 0xff;
  out[to + 14] = (t3 >>> 8) & 0xff;
  out[to + 15] = t3 & 0xff;
}

function whole(data: Uint8Array): void {
  if (data.length % 16 !== 0) throw new Error('AES data must be a multiple of 16 bytes');
}

function run(key: Uint8Array, work: (aes: Aes, out: Uint8Array) => void, size: number): Uint8Array {
  const aes = new Aes(key);
  const out = new Uint8Array(size);
  try {
    work(aes, out);
  } finally {
    aes.wipe();
  }
  return out;
}

export const aesEcbEncrypt = (key: Uint8Array, data: Uint8Array): Uint8Array => run(key, (aes, out) => aes.ecb(true, data, out), data.length);
export const aesEcbDecrypt = (key: Uint8Array, data: Uint8Array): Uint8Array => run(key, (aes, out) => aes.ecb(false, data, out), data.length);
export const aesCbcEncrypt = (key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array => run(key, (aes, out) => aes.cbc(true, iv, data, out), data.length);
export const aesCbcDecrypt = (key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array => run(key, (aes, out) => aes.cbc(false, iv, data, out), data.length);
