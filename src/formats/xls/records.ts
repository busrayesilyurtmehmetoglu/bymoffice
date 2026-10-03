// BIFF record stream: record ids and a reader that follows CONTINUE records.
//
// A record is a 16-bit id, a 16-bit length and that many bytes. Content that
// does not fit goes on in CONTINUE records, which the reader presents as one
// run of bytes. Every read is checked against the end of the data.

export const REC_FORMULA = 0x0006;
export const REC_EOF = 0x000a;
export const REC_HEADER = 0x0014;
export const REC_FOOTER = 0x0015;
export const REC_EXTERNSHEET = 0x0017;
export const REC_NAME = 0x0018;
export const REC_VERTICALPAGEBREAKS = 0x001a;
export const REC_HORIZONTALPAGEBREAKS = 0x001b;
export const REC_NOTE = 0x001c;
export const REC_SELECTION = 0x001d;
export const REC_DATEMODE = 0x0022;
export const REC_EXTERNNAME = 0x0023;
export const REC_LEFTMARGIN = 0x0026;
export const REC_RIGHTMARGIN = 0x0027;
export const REC_TOPMARGIN = 0x0028;
export const REC_BOTTOMMARGIN = 0x0029;
export const REC_PRINTHEADERS = 0x002a;
export const REC_PRINTGRIDLINES = 0x002b;
export const REC_FILEPASS = 0x002f;
export const REC_FONT = 0x0031;
export const REC_CONTINUE = 0x003c;
export const REC_WINDOW1 = 0x003d;
export const REC_PANE = 0x0041;
export const REC_CODEPAGE = 0x0042;
export const REC_DEFCOLWIDTH = 0x0055;
export const REC_COLINFO = 0x007d;
export const REC_WSBOOL = 0x0081;
export const REC_GRIDSET = 0x0082;
export const REC_HCENTER = 0x0083;
export const REC_VCENTER = 0x0084;
export const REC_BOUNDSHEET = 0x0085;
export const REC_PALETTE = 0x0092;
export const REC_STANDARDWIDTH = 0x0099;
export const REC_SCL = 0x00a0;
export const REC_SETUP = 0x00a1;
export const REC_MULRK = 0x00bd;
export const REC_MULBLANK = 0x00be;
export const REC_RSTRING = 0x00d6;
export const REC_XF = 0x00e0;
export const REC_MERGECELLS = 0x00e5;
export const REC_SST = 0x00fc;
export const REC_LABELSST = 0x00fd;
export const REC_USERSVIEWBEGIN = 0x01aa;
export const REC_USERSVIEWEND = 0x01ab;
export const REC_SUPBOOK = 0x01ae;
export const REC_HLINK = 0x01b8;
export const REC_DIMENSION = 0x0200;
export const REC_BLANK = 0x0201;
export const REC_NUMBER = 0x0203;
export const REC_LABEL = 0x0204;
export const REC_BOOLERR = 0x0205;
export const REC_STRING = 0x0207;
export const REC_ROW = 0x0208;
export const REC_ARRAY = 0x0221;
export const REC_DEFAULTROWHEIGHT = 0x0225;
export const REC_TABLE = 0x0236;
export const REC_WINDOW2 = 0x023e;
export const REC_RK = 0x027e;
export const REC_STYLE = 0x0293;
export const REC_FORMAT = 0x041e;
export const REC_SHRFMLA = 0x04bc;
export const REC_HLINKTOOLTIP = 0x0800;
export const REC_BOF = 0x0809;
export const REC_XFEXT = 0x087d;
export const REC_CONTINUEFRT12 = 0x087f;
export const REC_THEME = 0x0896;
export const REC_HEADERFOOTER = 0x089c;

/** The file is an xls file, but of a kind this reader does not open (encrypted, older than BIFF5). */
export class XlsUnsupportedError extends Error {
  override name = 'UnsupportedFileError';
}

export const corrupt = (message: string): Error => new Error(`Invalid xls: ${message}`);

/** File content quoted in an error message, cut so a hostile value cannot produce a huge message. */
export const quote = (text: string): string => (text.length > 40 ? `"${text.slice(0, 40)}…"` : `"${text}"`);

/** The bytes of a Workbook (or Book) stream. Offsets are relative to the stream. */
export interface RecordStream {
  bytes: Uint8Array;
  view: DataView;
}

export function recordStream(bytes: Uint8Array): RecordStream {
  return { bytes, view: new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength) };
}

const scratch = new DataView(new ArrayBuffer(8));

/** Reads little-endian values from a record and the records that continue it. */
export class RecordReader {
  private segment = 0;
  private pos: number;
  private end: number;
  /** Bytes in the records after the current one: added up once, so that asking what is left costs nothing. */
  private after = 0;

  /** `segments` holds the start and the end offset of each run of bytes, in reading order. */
  constructor(
    private readonly stream: RecordStream,
    private readonly segments: number[],
  ) {
    this.pos = segments[0];
    this.end = segments[1];
    for (let i = 2; i < segments.length; i += 2) this.after += segments[i + 1] - segments[i];
  }

  /** Bytes left in the current record, not counting the records that continue it. */
  get left(): number {
    return this.end - this.pos;
  }

  /** Nothing is left, in this record or in the ones that continue it. */
  get done(): boolean {
    return this.pos >= this.end && this.after === 0;
  }

  /** Bytes left in all records together. */
  get remaining(): number {
    return this.end - this.pos + this.after;
  }

  /** Moves to the next record that has content; false when there is none. */
  private nextSegment(): boolean {
    while ((this.segment + 1) * 2 < this.segments.length) {
      this.segment++;
      this.pos = this.segments[this.segment * 2];
      this.end = this.segments[this.segment * 2 + 1];
      this.after -= this.end - this.pos;
      if (this.end > this.pos) return true;
    }
    return false;
  }

  u8(): number {
    if (this.pos >= this.end && !this.nextSegment()) throw corrupt('a record ends in the middle of a field');
    return this.stream.bytes[this.pos++];
  }

  u16(): number {
    if (this.end - this.pos >= 2) {
      const value = this.stream.view.getUint16(this.pos, true);
      this.pos += 2;
      return value;
    }
    const low = this.u8();
    return low | (this.u8() << 8);
  }

  i16(): number {
    const value = this.u16();
    return value >= 0x8000 ? value - 0x10000 : value;
  }

  u32(): number {
    if (this.end - this.pos >= 4) {
      const value = this.stream.view.getUint32(this.pos, true);
      this.pos += 4;
      return value;
    }
    const low = this.u16();
    return (low | (this.u16() << 16)) >>> 0;
  }

  f64(): number {
    if (this.end - this.pos >= 8) {
      const value = this.stream.view.getFloat64(this.pos, true);
      this.pos += 8;
      return value;
    }
    for (let i = 0; i < 8; i++) scratch.setUint8(i, this.u8());
    return scratch.getFloat64(0, true);
  }

  skip(count: number): void {
    while (count > 0) {
      if (this.pos >= this.end && !this.nextSegment()) throw corrupt('a record ends in the middle of a field');
      const take = Math.min(count, this.end - this.pos);
      this.pos += take;
      count -= take;
    }
  }

  /** Skips up to `count` bytes; stops quietly where the data ends. */
  skipUpTo(count: number): void {
    this.skip(Math.min(count, this.remaining));
  }

  /** A copy of the next `count` bytes. */
  take(count: number): Uint8Array {
    if (count > this.remaining) throw corrupt('a record ends in the middle of a field');
    const out = new Uint8Array(count);
    let at = 0;
    while (at < count) {
      if (this.pos >= this.end) this.nextSegment();
      const size = Math.min(count - at, this.end - this.pos);
      out.set(this.stream.bytes.subarray(this.pos, this.pos + size), at);
      this.pos += size;
      at += size;
    }
    return out;
  }

  /** A copy of everything that is left. */
  rest(): Uint8Array {
    return this.take(this.remaining);
  }

  /**
   * `count` characters of a BIFF8 string, one byte each or two (`wide`). Where the text goes on
   * in a CONTINUE record, that record starts with a byte that states the width again, and the
   * width may change there.
   */
  chars(count: number, wide: boolean): string {
    const bytes = this.stream.bytes;
    let out = '';
    let codes: number[] = [];
    while (count > 0) {
      if (this.pos >= this.end) {
        if (!this.nextSegment()) throw corrupt('a string runs past the end of its record');
        wide = (bytes[this.pos++] & 1) !== 0;
        continue;
      }
      const available = this.end - this.pos;
      if (wide && available < 2) throw corrupt('a character is split between two records');
      const size = Math.min(count, wide ? available >> 1 : available, 8192);
      let p = this.pos;
      if (wide) {
        for (let i = 0; i < size; i++, p += 2) codes.push(bytes[p] | (bytes[p + 1] << 8));
      } else {
        for (let i = 0; i < size; i++) codes.push(bytes[p++]);
      }
      this.pos = p;
      count -= size;
      if (codes.length >= 8192) {
        out += String.fromCharCode.apply(null, codes);
        codes = [];
      }
    }
    return codes.length > 0 ? out + String.fromCharCode.apply(null, codes) : out;
  }
}

/**
 * Length of the record at `offset`, after checking that the record lies inside the stream.
 * The content starts at offset + 4.
 */
export function recordLength(stream: RecordStream, offset: number): number {
  const size = stream.bytes.length;
  if (offset + 4 > size) throw corrupt('the record stream ends in the middle of a record header');
  const length = stream.view.getUint16(offset + 2, true);
  if (offset + 4 + length > size) throw corrupt(`the record at offset ${offset} runs past the end of the stream`);
  return length;
}

/**
 * A reader over the record at `offset` and the CONTINUE records after it. Records of the
 * future record type family continue in CONTINUEFRT12 records, which repeat a 12 byte header.
 */
export function readerAt(stream: RecordStream, offset: number): RecordReader {
  const view = stream.view;
  const size = stream.bytes.length;
  const length = recordLength(stream, offset);
  const segments = [offset + 4, offset + 4 + length];
  let p = offset + 4 + length;
  while (p + 4 <= size) {
    const id = view.getUint16(p, true);
    if (id !== REC_CONTINUE && id !== REC_CONTINUEFRT12) break;
    const next = recordLength(stream, p);
    if (id === REC_CONTINUE) segments.push(p + 4, p + 4 + next);
    else if (next > 12) segments.push(p + 16, p + 4 + next);
    p += 4 + next;
  }
  return new RecordReader(stream, segments);
}
