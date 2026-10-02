// Splits the bytes of an RTF file into tokens (RTF 1.9.1, "Conventions of an
// RTF Reader").
//
//   {  }            a group opens, closes
//   \word           a control word: letters, then an optional number (with a
//                   sign), then one space that belongs to the word
//   \x              a control symbol: a backslash and one other character
//   \'hh            a byte given as two hex digits
//   \binN           N bytes follow that are data, not RTF
//   anything else   text, up to the next backslash or brace
//
// The tokenizer works on the bytes: a file of megabytes of picture data is
// never turned into a string. A token is described by the fields of the
// tokenizer itself, so that reading allocates nothing but the names of the
// control words, and those only the first time each is seen.
//
// Nothing is trusted: a number is cut off where it would no longer fit into
// 32 bits, the length of \bin is cut to what the file still holds, and every
// call of next() moves forward by at least one byte.

export const T_EOF = 0;
export const T_OPEN = 1;
export const T_CLOSE = 2;
/** `word`, `param`, `hasParam`. */
export const T_WORD = 3;
/** `code`: the character after the backslash. */
export const T_SYMBOL = 4;
/** bytes[start..stop): text as written, line ends included (they are not part of the text). */
export const T_TEXT = 5;
/** `code`: the byte. */
export const T_HEX = 6;
/** bytes[start..stop): the data of a \bin. */
export const T_BIN = 7;

/** Letters a control word may have before it is cut; the specification allows 32. */
const MAX_WORD = 32;
const MAX_PARAM = 0x7fffffff;

const isLetter = (c: number): boolean => (c >= 97 && c <= 122) || (c >= 65 && c <= 90);
const isDigit = (c: number): boolean => c >= 48 && c <= 57;

/** Value of a hex digit, -1 for anything else. */
export function hexValue(c: number): number {
  if (c >= 48 && c <= 57) return c - 48;
  if (c >= 97 && c <= 102) return c - 87;
  if (c >= 65 && c <= 70) return c - 55;
  return -1;
}

export class RtfTokenizer {
  readonly bytes: Uint8Array;
  readonly end: number;
  /** Where the next token starts. */
  pos: number;

  word = '';
  param = 0;
  hasParam = false;
  code = 0;
  start = 0;
  stop = 0;

  /** Names of control words by a hash of their letters. */
  private readonly names = new Map<number, string>();

  constructor(bytes: Uint8Array, start = 0, end = bytes.length) {
    this.bytes = bytes;
    this.pos = start;
    this.end = Math.min(end, bytes.length);
  }

  /** The name bytes[from..to) as a string; one string for each name, however often it occurs. */
  private name(from: number, to: number): string {
    const b = this.bytes;
    let hash = to - from;
    for (let i = from; i < to; i++) hash = (Math.imul(hash, 31) + b[i]) | 0;
    const known = this.names.get(hash);
    if (known !== undefined && known.length === to - from) {
      let same = true;
      for (let i = from; i < to; i++) {
        if (known.charCodeAt(i - from) !== b[i]) {
          same = false;
          break;
        }
      }
      if (same) return known;
    }
    let text = '';
    for (let i = from; i < to; i++) text += String.fromCharCode(b[i]);
    this.names.set(hash, text);
    return text;
  }

  next(): number {
    const b = this.bytes;
    const end = this.end;
    let p = this.pos;
    if (p >= end) return T_EOF;
    const c = b[p];
    if (c === 0x7b) {
      this.pos = p + 1;
      return T_OPEN;
    }
    if (c === 0x7d) {
      this.pos = p + 1;
      return T_CLOSE;
    }
    if (c !== 0x5c) {
      this.start = p;
      p++;
      while (p < end) {
        const x = b[p];
        if (x === 0x5c || x === 0x7b || x === 0x7d) break;
        p++;
      }
      this.stop = p;
      this.pos = p;
      return T_TEXT;
    }

    p++;
    // A backslash that ends the file stands for nothing.
    if (p >= end) {
      this.pos = p;
      return T_EOF;
    }
    const d = b[p];
    if (!isLetter(d)) {
      if (d === 0x27) {
        p++;
        let value = 0;
        let digits = 0;
        while (digits < 2 && p < end) {
          const h = hexValue(b[p]);
          if (h < 0) break;
          value = value * 16 + h;
          digits++;
          p++;
        }
        this.pos = p;
        this.code = digits ? value : 0x27;
        // Without a digit after it the symbol stands for nothing; the reader passes over it.
        return digits ? T_HEX : T_SYMBOL;
      }
      this.code = d;
      this.pos = p + 1;
      return T_SYMBOL;
    }

    let q = p + 1;
    while (q < end && q - p < MAX_WORD && isLetter(b[q])) q++;
    this.word = this.name(p, q);
    p = q;
    let negative = false;
    this.hasParam = false;
    this.param = 0;
    if (p + 1 < end && b[p] === 0x2d && isDigit(b[p + 1])) {
      negative = true;
      p++;
    }
    if (p < end && isDigit(b[p])) {
      let value = 0;
      while (p < end && isDigit(b[p])) {
        // More digits than a number holds: the value stays at the largest one.
        value = value > 214748364 ? MAX_PARAM : value * 10 + (b[p] - 48);
        if (value > MAX_PARAM) value = MAX_PARAM;
        p++;
      }
      this.hasParam = true;
      this.param = negative ? -value : value;
    }
    if (p < end && b[p] === 0x20) p++;
    if (this.word === 'bin') {
      const length = this.hasParam && this.param > 0 ? Math.min(this.param, end - p) : 0;
      this.start = p;
      this.stop = p + length;
      this.pos = p + length;
      return T_BIN;
    }
    this.pos = p;
    return T_WORD;
  }

  /**
   * Passes over the rest of the group the tokenizer is in, groups inside it
   * included; the brace that closes it is the next token. At the end of a
   * file that is cut short the position is the end.
   */
  skipGroup(): void {
    const b = this.bytes;
    const end = this.end;
    let p = this.pos;
    let depth = 0;
    while (p < end) {
      const c = b[p];
      if (c === 0x7b) depth++;
      else if (c === 0x7d) {
        if (depth === 0) break;
        depth--;
      } else if (c === 0x5c) {
        // "\bin" hides bytes that may hold braces; every other backslash hides the character after it.
        if (p + 4 < end && b[p + 1] === 0x62 && b[p + 2] === 0x69 && b[p + 3] === 0x6e && isDigit(b[p + 4])) {
          let q = p + 4;
          let value = 0;
          while (q < end && isDigit(b[q])) {
            value = value > 214748364 ? MAX_PARAM : value * 10 + (b[q] - 48);
            q++;
          }
          if (q < end && b[q] === 0x20) q++;
          p = q + Math.min(value, end - q);
          continue;
        }
        p++;
      }
      p++;
    }
    this.pos = Math.min(p, end);
  }
}
