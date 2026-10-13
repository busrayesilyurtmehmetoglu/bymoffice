// The core of the PDF writer: values, indirect objects, streams, the cross
// reference table and the trailer (PDF 1.4, the classic `xref` table).
//
// Nothing is ever put into the file as text that was not made here: a value
// is a tree of numbers, names, strings and references, and every one of them
// is written by its own rule (names with #xx, strings with escapes or as
// hexadecimal), so that text from a document cannot become PDF syntax.
//
// The file is not held as one string or one array: what is written is a list
// of parts, and the writer only remembers where each object starts.

export type PdfErrorCode = 'invalid' | 'too-large' | 'empty' | 'font';

/** Why a PDF could not be written: a value that cannot be, a file beyond the limit, a document without pages, a font that could not be put into the file when it was ended. */
export class PdfError extends Error {
  readonly code: PdfErrorCode;
  constructor(code: PdfErrorCode, message: string) {
    super(message);
    this.name = 'PdfError';
    this.code = code;
  }
}

/** A name object: "/Type". */
export class PdfName {
  constructor(readonly name: string) {}
}

/** A reference to an indirect object: "12 0 R". */
export class PdfRef {
  constructor(readonly id: number) {}
}

/** A string object, as bytes: written as "(...)" with escapes, or as "<...>". */
export class PdfString {
  constructor(
    readonly bytes: Uint8Array,
    readonly hex = false,
  ) {}
}

export interface PdfDict {
  [key: string]: PdfValue | undefined;
}
export type PdfValue = number | boolean | null | PdfName | PdfRef | PdfString | PdfValue[] | PdfDict;

export const name = (text: string): PdfName => new PdfName(text);

/** The largest offset the cross reference table of this writer states: files stay below 4 GB. */
export const MAX_PDF_BYTES = 0xffffffff;
/** How deep a value may be nested. */
const MAX_DEPTH = 32;

/**
 * A number as PDF writes it: digits, at most `digits` of them after the
 * point, never an exponent and never anything that depends on a locale.
 * Throws for a number that is not finite or too large to write out.
 */
export function pdfNumber(value: number, digits = 4): string {
  if (typeof value !== 'number' || !isFinite(value)) throw new PdfError('invalid', 'A number in a PDF must be finite');
  if (Math.abs(value) >= 1e15) throw new PdfError('invalid', 'A number in a PDF is too large');
  if (Math.floor(value) === value) return value === 0 ? '0' : String(value);
  let text = value.toFixed(Math.min(Math.max(digits, 0), 8));
  if (text.indexOf('.') >= 0) text = text.replace(/0+$/, '').replace(/\.$/, '');
  return text === '-0' || text === '' ? '0' : text;
}

const HEX = '0123456789abcdef';

/** Bytes as hexadecimal digits. */
export function hexOf(bytes: Uint8Array): string {
  const parts: string[] = [];
  let piece = '';
  for (let i = 0; i < bytes.length; i++) {
    piece += HEX.charAt(bytes[i] >> 4) + HEX.charAt(bytes[i] & 15);
    if (piece.length >= 8192) {
      parts.push(piece);
      piece = '';
    }
  }
  parts.push(piece);
  return parts.join('');
}

/** A name with everything but plain letters, digits and a few signs written as #xx (of its UTF-8 bytes). */
function nameText(value: string): string {
  let out = '/';
  for (let i = 0; i < value.length; i++) {
    let code = value.charCodeAt(i);
    const char = value.charAt(i);
    if ((code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || char === '-' || char === '_' || char === '.' || char === '+') {
      out += char;
      continue;
    }
    if (code === 0) throw new PdfError('invalid', 'A name in a PDF cannot hold a NUL');
    // UTF-8, with a pair of surrogates as one character and a lone one as U+FFFD.
    if (code >= 0xd800 && code <= 0xdfff) {
      const next = i + 1 < value.length ? value.charCodeAt(i + 1) : 0;
      if (code <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
        code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
        i++;
      } else code = 0xfffd;
    }
    const bytes = code < 0x80 ? [code] : code < 0x800 ? [0xc0 | (code >> 6), 0x80 | (code & 63)] : code < 0x10000 ? [0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63)] : [0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63)];
    for (let b = 0; b < bytes.length; b++) out += '#' + HEX.charAt(bytes[b] >> 4) + HEX.charAt(bytes[b] & 15);
  }
  return out;
}

/** A string in parentheses: the three characters that mean something are escaped, everything outside printable ASCII is octal. */
function literalText(bytes: Uint8Array): string {
  let out = '(';
  for (let i = 0; i < bytes.length; i++) {
    const code = bytes[i];
    if (code === 0x28 || code === 0x29 || code === 0x5c) out += '\\' + String.fromCharCode(code);
    else if (code >= 0x20 && code < 0x7f) out += String.fromCharCode(code);
    else out += '\\' + ((code >> 6) & 7) + ((code >> 3) & 7) + (code & 7);
  }
  return out + ')';
}

/**
 * A text string of the document information: plain ASCII as it is, anything
 * else as UTF-16BE after a byte order mark. A surrogate without its partner
 * becomes U+FFFD, and characters that cannot be in a title (controls) are left out.
 */
export function pdfText(text: string): PdfString {
  let ascii = true;
  const units: number[] = [];
  for (let i = 0; i < text.length; i++) {
    let code = text.charCodeAt(i);
    if (code < 0x20 || (code >= 0x7f && code < 0xa0)) {
      // A tab or a line end separates words; other controls are dropped.
      if (code !== 9 && code !== 10 && code !== 13) continue;
      code = 0x20;
    }
    if (code >= 0xd800 && code <= 0xdfff) {
      const next = i + 1 < text.length ? text.charCodeAt(i + 1) : 0;
      if (code <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
        units.push(code, next);
        i++;
        ascii = false;
        continue;
      }
      code = 0xfffd;
    }
    if (code >= 0x7f) ascii = false;
    units.push(code);
  }
  if (ascii) {
    const bytes = new Uint8Array(units.length);
    for (let i = 0; i < units.length; i++) bytes[i] = units[i];
    return new PdfString(bytes);
  }
  const bytes = new Uint8Array(2 + units.length * 2);
  bytes[0] = 0xfe;
  bytes[1] = 0xff;
  for (let i = 0; i < units.length; i++) {
    bytes[2 + i * 2] = units[i] >> 8;
    bytes[3 + i * 2] = units[i] & 255;
  }
  return new PdfString(bytes);
}

/** A date of the document information, in universal time: "D:20261013120000Z". Undefined for a date that is none. */
export function pdfDate(date: Date): PdfString | undefined {
  if (!(date instanceof Date) || isNaN(date.getTime())) return undefined;
  const year = date.getUTCFullYear();
  if (year < 0 || year > 9999) return undefined;
  const two = (value: number) => (value < 10 ? '0' : '') + value;
  const text = 'D:' + ('000' + year).slice(-4) + two(date.getUTCMonth() + 1) + two(date.getUTCDate()) + two(date.getUTCHours()) + two(date.getUTCMinutes()) + two(date.getUTCSeconds()) + 'Z';
  return new PdfString(latin1(text));
}

/** The bytes of text that only holds characters below U+0100. */
export function latin1(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code > 255) throw new PdfError('invalid', 'Text that is written as bytes holds a character beyond U+00FF');
    bytes[i] = code;
  }
  return bytes;
}

/** A value as the text of a PDF object. */
export function serialize(value: PdfValue, depth = 0): string {
  if (depth > MAX_DEPTH) throw new PdfError('invalid', 'A value in a PDF is nested too deeply');
  if (value === null) return 'null';
  if (typeof value === 'number') return pdfNumber(value);
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  if (value instanceof PdfName) return nameText(value.name);
  if (value instanceof PdfRef) {
    if (!(value.id >= 1) || Math.floor(value.id) !== value.id) throw new PdfError('invalid', 'A reference to an object that cannot be');
    return value.id + ' 0 R';
  }
  if (value instanceof PdfString) return value.hex ? '<' + hexOf(value.bytes) + '>' : literalText(value.bytes);
  if (Array.isArray(value)) {
    const parts: string[] = [];
    for (let i = 0; i < value.length; i++) parts.push(serialize(value[i], depth + 1));
    return '[' + parts.join(' ') + ']';
  }
  if (typeof value === 'object') {
    const parts: string[] = [];
    for (const key of Object.keys(value)) {
      const entry = value[key];
      if (entry === undefined) continue;
      parts.push(nameText(key) + ' ' + serialize(entry, depth + 1));
    }
    return '<<' + parts.join(' ') + '>>';
  }
  throw new PdfError('invalid', 'A value that is no PDF object');
}

export interface PdfWriterOptions {
  /** The file may not grow beyond this many bytes; MAX_PDF_BYTES at most, which is also the default. */
  maxBytes?: number;
}

/**
 * Writes the objects of a file one after the other. An object number is
 * taken with reserve() (or by add() itself), so that objects can name each
 * other before they are written; finish() writes the table of where they
 * are. The parts of the file are handed out in order by takeParts(): a
 * caller that passes them on as they come never holds the file twice.
 */
export class PdfWriter {
  private parts: Uint8Array[] = [];
  /** Offset of every object, by its number; undefined for a number that was reserved and not written yet. */
  private readonly offsets: (number | undefined)[] = [undefined];
  private offset = 0;
  private finished = false;
  private readonly maxBytes: number;

  constructor(options: PdfWriterOptions = {}) {
    const max = options.maxBytes;
    this.maxBytes = typeof max === 'number' && max > 0 && max < MAX_PDF_BYTES ? Math.floor(max) : MAX_PDF_BYTES;
    // The second line holds bytes above 127: it tells programs that carry files around that this one is binary.
    this.push(latin1('%PDF-1.4\n%\xe2\xe3\xcf\xd3\n'));
  }

  /** Bytes written so far. */
  get size(): number {
    return this.offset;
  }

  /** Number of objects, the free one at 0 included: the /Size of the trailer. */
  get objectCount(): number {
    return this.offsets.length;
  }

  private push(bytes: Uint8Array): void {
    if (this.finished) throw new PdfError('invalid', 'The PDF is finished: nothing can be added to it');
    if (this.offset + bytes.length > this.maxBytes) throw new PdfError('too-large', `The PDF would be larger than ${this.maxBytes} bytes`);
    if (bytes.length === 0) return;
    this.parts.push(bytes);
    this.offset += bytes.length;
  }

  /** Takes a number for an object that is written later. */
  reserve(): PdfRef {
    this.offsets.push(undefined);
    return new PdfRef(this.offsets.length - 1);
  }

  private begin(ref: PdfRef | undefined): PdfRef {
    const at = ref || this.reserve();
    if (!(at.id >= 1) || at.id >= this.offsets.length || Math.floor(at.id) !== at.id) throw new PdfError('invalid', `Object ${at.id} was never reserved`);
    if (this.offsets[at.id] !== undefined) throw new PdfError('invalid', `Object ${at.id} is written twice`);
    return at;
  }

  /** Writes an object; under the number of `ref` when one was reserved for it. */
  add(value: PdfValue, ref?: PdfRef): PdfRef {
    const at = this.begin(ref);
    const bytes = latin1(at.id + ' 0 obj\n' + serialize(value) + '\nendobj\n');
    const offset = this.offset;
    this.push(bytes);
    this.offsets[at.id] = offset;
    return at;
  }

  /**
   * Writes a stream: `data` as it is to be in the file (already compressed
   * when the dictionary names a filter). /Length is set here.
   */
  addStream(dict: PdfDict, data: Uint8Array, ref?: PdfRef): PdfRef {
    const at = this.begin(ref);
    const head = latin1(at.id + ' 0 obj\n' + serialize({ ...dict, Length: data.length }) + '\nstream\n');
    const tail = latin1('\nendstream\nendobj\n');
    if (this.offset + head.length + data.length + tail.length > this.maxBytes) throw new PdfError('too-large', `The PDF would be larger than ${this.maxBytes} bytes`);
    const offset = this.offset;
    this.push(head);
    this.push(data);
    this.push(tail);
    this.offsets[at.id] = offset;
    return at;
  }

  /** The parts written since the last call, in order; the writer lets go of them. */
  takeParts(): Uint8Array[] {
    const parts = this.parts;
    this.parts = [];
    return parts;
  }

  /**
   * Writes the cross reference table and the trailer. `id` is the pair of
   * file identifiers; the same bytes are written twice when one is given.
   */
  finish(trailer: { root: PdfRef; info?: PdfRef; id?: Uint8Array }): void {
    const count = this.offsets.length;
    const lines: string[] = ['xref\n0 ' + count + '\n0000000000 65535 f \n'];
    for (let i = 1; i < count; i++) {
      const offset = this.offsets[i];
      if (offset === undefined) throw new PdfError('invalid', `Object ${i} was reserved and never written`);
      // Every entry is exactly 20 bytes long.
      lines.push(('0000000000' + offset).slice(-10) + ' 00000 n \n');
    }
    const dict: PdfDict = { Size: count, Root: trailer.root };
    if (trailer.info) dict.Info = trailer.info;
    if (trailer.id) dict.ID = [new PdfString(trailer.id, true), new PdfString(trailer.id, true)];
    const start = this.offset;
    this.push(latin1(lines.join('') + 'trailer\n' + serialize(dict) + '\nstartxref\n' + start + '\n%%EOF\n'));
    this.finished = true;
  }
}

/** The parts of a file as one array. */
export function joinParts(parts: readonly Uint8Array[]): Uint8Array {
  let total = 0;
  for (let i = 0; i < parts.length; i++) total += parts[i].length;
  const out = new Uint8Array(total);
  let at = 0;
  for (let i = 0; i < parts.length; i++) {
    out.set(parts[i], at);
    at += parts[i].length;
  }
  return out;
}
