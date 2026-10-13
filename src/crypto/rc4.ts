// RC4: the stream cipher of password-protected xls and doc files. Broken as
// a cipher, and only ever used here to read files that were written with it.

export class Rc4 {
  private readonly s = new Uint8Array(256);
  private i = 0;
  private j = 0;

  constructor(key: Uint8Array) {
    if (key.length < 1 || key.length > 256) throw new Error('RC4 key must be 1 to 256 bytes');
    const s = this.s;
    for (let i = 0; i < 256; i++) s[i] = i;
    for (let i = 0, j = 0; i < 256; i++) {
      j = (j + s[i] + key[i % key.length]) & 0xff;
      const t = s[i];
      s[i] = s[j];
      s[j] = t;
    }
  }

  /** XORs the key stream into `data[start..end)`. */
  apply(data: Uint8Array, start = 0, end = data.length): void {
    const s = this.s;
    let i = this.i;
    let j = this.j;
    for (let at = start; at < end; at++) {
      i = (i + 1) & 0xff;
      const t = s[i];
      j = (j + t) & 0xff;
      s[i] = s[j];
      s[j] = t;
      data[at] ^= s[(s[i] + t) & 0xff];
    }
    this.i = i;
    this.j = j;
  }

  /** Lets `count` bytes of the key stream go by. */
  skip(count: number): void {
    const s = this.s;
    let i = this.i;
    let j = this.j;
    for (; count > 0; count--) {
      i = (i + 1) & 0xff;
      const t = s[i];
      j = (j + t) & 0xff;
      s[i] = s[j];
      s[j] = t;
    }
    this.i = i;
    this.j = j;
  }

  wipe(): void {
    this.s.fill(0);
    this.i = this.j = 0;
  }
}
