// Pictures, shapes and comments of a BIFF8 workbook.
//
// The pictures of a workbook are kept once, in the OfficeArt drawing group of
// the globals (MSODRAWINGGROUP records and the CONTINUE records after them: a
// store of "BLIPs"). Each sheet has one OfficeArt drawing, cut into
// MSODRAWING records between which the OBJ and TXO records of its shapes
// stand: the n-th client data of the drawing belongs to the n-th OBJ record,
// the n-th client text box to the n-th TXO record. A comment is a NOTE record
// (cell, author, whether its box is always shown) that names the object whose
// text box holds its text.
//
// Everything is read into the same read-only view as the drawings and
// comments of an xlsx file. Nothing in here is trusted: record lengths are
// checked against the record around them, nesting and counts are bounded,
// and pictures are unpacked only when they are first shown.

import { readBlip, type MediaBudget } from '../doc/drawings';
import { i16, i32, u16, u32, u8 } from '../doc/bytes';
import { resolveColor } from '../../model/sheet/color';
import type {
  DrawingAnchor,
  DrawingGroup,
  DrawingObject,
  DrawingObjectBase,
  DrawingPicture,
  DrawingShape,
  DrawingText,
  DrawingTextRun,
  Font,
  RichRun,
  SheetComment,
  SheetDrawing,
  StyleSheet,
  WorkbookMedia,
} from '../../model/sheet/types';
import type { XlsChartRead } from './chart';
import { readerAt, type RecordStream } from './records';
import { readString16, toSharedString, type RawString } from './strings';

export const REC_OBJ = 0x005d;
export const REC_MSODRAWINGGROUP = 0x00eb;
export const REC_MSODRAWING = 0x00ec;
export const REC_TXO = 0x01b6;

/** Shapes read from one sheet, group members included; a sheet with more shows the first ones. */
export const XLS_MAX_SHAPES = 20000;
/** Pictures of the store of one workbook. */
const MAX_BLIPS = 65536;
/** Deepest nesting of OfficeArt containers that is followed. */
const MAX_ART_DEPTH = 16;
/** Bytes all pictures of one workbook may take that are not just a view of the file (bitmaps given a file header). */
const MEDIA_BUDGET = 256 * 1024 * 1024;
const MAX_TEXT = 32767;
const MAX_COMMENTS = 100000;
const EMU_PER_PIXEL = 9525;
const KEY_PREFIX = 'xls:blip:';

const RT_BSTORE = 0xf001;
const RT_DG_CONTAINER = 0xf002;
const RT_SPGR_CONTAINER = 0xf003;
const RT_SP_CONTAINER = 0xf004;
const RT_DGG_CONTAINER = 0xf000;
const RT_BSE = 0xf007;
const RT_SPGR = 0xf009;
const RT_SP = 0xf00a;
const RT_OPT = 0xf00b;
const RT_CLIENT_TEXTBOX = 0xf00d;
const RT_CHILD_ANCHOR = 0xf00f;
const RT_CLIENT_ANCHOR = 0xf010;
const RT_CLIENT_DATA = 0xf011;
const RT_OPT2 = 0xf121;
const RT_OPT3 = 0xf122;
const RT_BLIP_FIRST = 0xf018;
const RT_BLIP_LAST = 0xf117;
/** BLIP records that hold a picture a browser may decode; metafiles are never unpacked. */
const RASTER_BLIPS: { [type: number]: true } = { 0xf01d: true, 0xf01e: true, 0xf01f: true, 0xf02a: true };

// Properties of a shape (MS-ODRAW 2.3), by property id.
const P_ROTATION = 0x0004;
const P_TEXT_LEFT = 0x0081;
const P_CROP_TOP = 0x0100;
const P_CROP_BOTTOM = 0x0101;
const P_CROP_LEFT = 0x0102;
const P_CROP_RIGHT = 0x0103;
const P_PICTURE = 0x0104;
const P_FILL_COLOR = 0x0181;
const P_FILL_FLAGS = 0x01bf;
const P_LINE_COLOR = 0x01c0;
const P_LINE_WIDTH = 0x01cb;
const P_LINE_DASHING = 0x01ce;
const P_LINE_FLAGS = 0x01ff;
const P_NAME = 0x0380;
const P_DESCRIPTION = 0x0381;
const P_GROUP_FLAGS = 0x03bf;

// Object types of an OBJ record (MS-XLS 2.5.143).
const OT_GROUP = 0x00;
const OT_CHART = 0x05;
const OT_PICTURE = 0x08;
const OT_NOTE = 0x19;
/** Buttons, check boxes, lists, the buttons of a filter: controls, not drawings. */
const CONTROLS: { [ot: number]: true } = { 0x07: true, 0x0b: true, 0x0c: true, 0x0d: true, 0x0e: true, 0x0f: true, 0x10: true, 0x11: true, 0x12: true, 0x13: true, 0x14: true };

/** Shape types (MSOSPT) that are drawn as something other than their box. */
const GEOMETRY: { [type: number]: string } = { 1: 'rect', 2: 'roundRect', 3: 'ellipse', 20: 'line', 32: 'straightConnector1', 33: 'bentConnector2', 34: 'bentConnector3', 35: 'bentConnector4', 36: 'bentConnector5', 37: 'curvedConnector2', 38: 'curvedConnector3', 39: 'curvedConnector4', 40: 'curvedConnector5', 202: 'rect' };

interface ArtRecord {
  version: number;
  instance: number;
  type: number;
  /** Start and end of the content. */
  start: number;
  end: number;
}

function recordAt(b: Uint8Array, at: number, end: number): ArtRecord | undefined {
  if (at + 8 > end) return undefined;
  const head = u16(b, at);
  const length = u32(b, at + 4);
  const start = at + 8;
  if (start + length > end) return undefined;
  return { version: head & 15, instance: head >> 4, type: u16(b, at + 2), start, end: start + length };
}

/** Runs of bytes of one OfficeArt stream, gathered from the records it is cut into. */
class ArtBytes {
  private parts: Uint8Array[] = [];
  private total = 0;
  private joined: Uint8Array | undefined;
  /** Offset in the record stream at which a CONTINUE record goes on with this data; -1 when none does. */
  expect = -1;

  add(part: Uint8Array, next: number): void {
    this.parts.push(part);
    this.total += part.length;
    this.joined = undefined;
    this.expect = next;
  }

  get length(): number {
    return this.total;
  }

  /** All of it as one array: the record itself when there is only one, a copy otherwise. */
  bytes(): Uint8Array {
    if (this.joined) return this.joined;
    if (this.parts.length === 1) this.joined = this.parts[0];
    else {
      const out = new Uint8Array(this.total);
      let at = 0;
      for (const part of this.parts) {
        out.set(part, at);
        at += part.length;
      }
      this.joined = out;
    }
    this.parts = [this.joined];
    return this.joined;
  }
}

/** The picture store of a workbook: where each BLIP is, read when a picture is first shown. */
export class XlsPictures {
  private readonly art = new ArtBytes();
  private blips: { at: number; end: number }[] | undefined;
  private readonly budget: MediaBudget = { left: MEDIA_BUDGET };
  private readonly used = new Set<number>();

  /** A MSODRAWINGGROUP record, or a CONTINUE record at `offset` that goes on with one; false when the record is neither. */
  add(stream: RecordStream, id: number, offset: number, length: number): boolean {
    if (id !== REC_MSODRAWINGGROUP && offset !== this.art.expect) return false;
    this.art.add(stream.bytes.subarray(offset + 4, offset + 4 + length), offset + 4 + length);
    this.blips = undefined;
    return true;
  }

  /** Any other record ends the run of records that belong to the store. */
  stop(): void {
    this.art.expect = -1;
  }

  private index(): { at: number; end: number }[] {
    if (this.blips) return this.blips;
    const blips: { at: number; end: number }[] = [];
    this.blips = blips;
    if (this.art.length < 16) return blips;
    const b = this.art.bytes();
    const head = recordAt(b, 0, b.length);
    if (!head || head.type !== RT_DGG_CONTAINER) return blips;
    for (let p = head.start; p < head.end; ) {
      const record = recordAt(b, p, head.end);
      if (!record) break;
      if (record.type === RT_BSTORE) {
        for (let q = record.start; q < record.end && blips.length < MAX_BLIPS; ) {
          const entry = recordAt(b, q, record.end);
          if (!entry) break;
          if (entry.type === RT_BSE && entry.end - entry.start >= 36) {
            // The picture follows its entry and the name the entry may have.
            const embedded = entry.start + 36 + u8(b, entry.start + 33);
            blips.push(embedded + 8 <= entry.end ? { at: embedded, end: entry.end } : { at: 0, end: 0 });
          } else if (entry.type >= RT_BLIP_FIRST && entry.type <= RT_BLIP_LAST) blips.push({ at: q, end: entry.end });
          else blips.push({ at: 0, end: 0 });
          q = entry.end;
        }
      }
      p = record.end;
    }
    return blips;
  }

  /** The key of the picture a shape names (its "pib", counted from 1); undefined when the store has no such picture. */
  key(pib: number): string | undefined {
    const blip = this.index()[pib - 1];
    if (!blip || blip.end <= blip.at) return undefined;
    this.used.add(pib);
    return KEY_PREFIX + pib;
  }

  private read(key: string): Uint8Array {
    const pib = key.indexOf(KEY_PREFIX) === 0 && /^\d{1,6}$/.test(key.slice(KEY_PREFIX.length)) ? Number(key.slice(KEY_PREFIX.length)) : 0;
    const blip = this.used.has(pib) ? this.index()[pib - 1] : undefined;
    if (!blip) throw new Error(`Not a picture of this workbook: ${key}`);
    const b = this.art.bytes();
    // Metafiles are not shown, so they are not unpacked either.
    if (!RASTER_BLIPS[u16(b, blip.at + 2)]) throw new Error('The picture is not of a kind that is shown');
    const media = readBlip(b, blip.at, blip.end, this.budget);
    if (!media) throw new Error('The picture cannot be read');
    return media.bytes;
  }

  /** Where the view asks for the bytes of the pictures; undefined when no drawing shows one. */
  media(): WorkbookMedia | undefined {
    if (this.used.size === 0) return undefined;
    return {
      readPart: (key: string) => {
        try {
          return Promise.resolve(this.read(key));
        } catch (error) {
          return Promise.reject(error);
        }
      },
    };
  }
}

interface RawObject {
  type: number;
  id: number;
  /** A chart: the offset of the BOF record of its substream, which follows the OBJ record. */
  chart?: number;
}

interface RawText {
  text: string;
  runs?: number[];
  align: number;
  anchor: number;
}

interface RawNote {
  row: number;
  col: number;
  visible: boolean;
  object: number;
  author: string;
}

interface RawShape {
  /** MSOSPT of the shape, and the flags of its FSP record. */
  kind: number;
  flags: number;
  values: Map<number, number>;
  texts: Map<number, string>;
  /** Client anchor: where the shape is on the sheet. */
  anchor?: DrawingAnchor;
  /** Child anchor: left, top, right, bottom in the coordinates of the group. */
  child?: number[];
  /** Coordinates of the members, for the shape that stands for a group. */
  space?: number[];
  object?: RawObject;
  text?: RawText;
  children?: RawShape[];
}

export interface XlsArtContext {
  styles: StyleSheet;
  fontOf(index: number): Font | undefined;
  pictures: XlsPictures;
  /** Last row and column of the format. */
  maxRows: number;
  maxCols: number;
  /** Reads the chart whose substream starts at the offset; undefined for a chart object no substream follows. Without it charts are placeholders. */
  chart?(offset: number | undefined): XlsChartRead;
}

/** Warnings about single charts of one sheet; what goes wrong beyond them is counted in one line. */
const MAX_CHART_WARNINGS = 20;

/** The drawing of one sheet while its records are read, and what it comes to. */
export class XlsSheetArt {
  private readonly art = new ArtBytes();
  private readonly objects: RawObject[] = [];
  private readonly texts: RawText[] = [];
  private readonly notes: RawNote[] = [];
  private skipped = 0;
  /** Why charts of the sheet are placeholders, and how many more are. */
  private readonly chartProblems: string[] = [];
  private chartsUnsaid = 0;

  constructor(private readonly stream: RecordStream) {}

  /** A MSODRAWING record at `offset`, or a CONTINUE record that goes on with one; false when the record is neither. */
  drawing(id: number, offset: number, length: number): boolean {
    if (id !== REC_MSODRAWING && offset !== this.art.expect) return false;
    this.art.add(this.stream.bytes.subarray(offset + 4, offset + 4 + length), offset + 4 + length);
    return true;
  }

  /** An OBJ record: the kind of object the shape before it is, and its id. */
  object(offset: number, length: number): void {
    this.art.expect = -1;
    const b = this.stream.bytes;
    const d = offset + 4;
    // The common object data comes first: ft = 0x15, cb = 0x12, then the type and the id.
    if (this.objects.length >= XLS_MAX_SHAPES) return;
    if (length >= 8 && u16(b, d) === 0x15) this.objects.push({ type: u16(b, d + 4), id: u16(b, d + 6) });
    else this.objects.push({ type: -1, id: -1 });
  }

  /** The BOF record of a chart substream: the chart of the object record before it. */
  chart(offset: number): void {
    this.art.expect = -1;
    const object = this.objects[this.objects.length - 1];
    if (object && object.type === OT_CHART && object.chart === undefined) object.chart = offset;
  }

  /** A TXO record: the text of the shape before it, in the CONTINUE records that follow. */
  text(offset: number, length: number): void {
    this.art.expect = -1;
    if (this.texts.length >= XLS_MAX_SHAPES) return;
    const raw: RawText = { text: '', align: 1, anchor: 1 };
    this.texts.push(raw);
    if (length < 18) return;
    try {
      const r = readerAt(this.stream, offset);
      const flags = r.u16();
      raw.align = (flags >> 1) & 7;
      raw.anchor = (flags >> 4) & 7;
      r.skip(8);
      const count = r.u16();
      const runBytes = r.u16();
      r.skip(4);
      if (count === 0 || r.done) return;
      // The first CONTINUE record: a byte that says how wide the characters are, then the characters.
      const wide = (r.u8() & 1) !== 0;
      raw.text = r.chars(Math.min(count, MAX_TEXT), wide);
      if (count > MAX_TEXT) return;
      // The next one: runs of eight bytes, the first character and the font of each; the last run only ends the text.
      const runs: number[] = [];
      for (let i = 0; i + 8 <= runBytes && r.remaining >= 8 && runs.length < 2048; i += 8) {
        const first = r.u16();
        const font = r.u16();
        r.skip(4);
        if (first < count) runs.push(first, font);
      }
      if (runs.length > 0) raw.runs = runs;
    } catch (error) {
      // Text that cannot be read is left as far as it was read.
    }
  }

  /** A NOTE record: the cell of a comment and the object that holds its text. */
  note(offset: number, length: number): void {
    this.art.expect = -1;
    if (length < 8 || this.notes.length >= MAX_COMMENTS) return;
    const b = this.stream.bytes;
    const d = offset + 4;
    const note: RawNote = { row: u16(b, d), col: u16(b, d + 2), visible: (u16(b, d + 4) & 2) !== 0, object: u16(b, d + 6), author: '' };
    if (length > 11) {
      try {
        const r = readerAt(this.stream, offset);
        r.skip(8);
        note.author = readString16(r).text;
      } catch (error) {
        // A comment without the name of who wrote it.
      }
    }
    this.notes.push(note);
  }

  /** Any other record ends the run of records that belong to the drawing. */
  stop(): void {
    this.art.expect = -1;
  }

  /** What was gathered, as the drawings and comments of the sheet. */
  finish(context: XlsArtContext): { drawings?: SheetDrawing[]; comments?: SheetComment[]; warnings?: string[] } {
    const out: { drawings?: SheetDrawing[]; comments?: SheetComment[]; warnings?: string[] } = {};
    if (this.art.length === 0 && this.notes.length === 0) return out;
    const drawings: SheetDrawing[] = [];
    /** The shapes that are the boxes of comments, by object id. */
    const boxes = new Map<number, RawShape>();
    try {
      const shapes = this.shapes();
      for (const shape of shapes) {
        if (shape.object && shape.object.type === OT_NOTE) {
          if (!boxes.has(shape.object.id)) boxes.set(shape.object.id, shape);
          continue;
        }
        const object = shape.anchor ? this.objectOf(shape, context, 0) : undefined;
        if (object && shape.anchor) drawings.push({ anchor: shape.anchor, object });
        else if (!(shape.object && CONTROLS[shape.object.type])) this.skipped++;
      }
    } catch (error) {
      this.skipped++;
    }
    if (drawings.length > 0) out.drawings = drawings;

    const comments: SheetComment[] = [];
    const seen = new Set<number>();
    for (const note of this.notes) {
      if (note.row >= context.maxRows || note.col >= context.maxCols) continue;
      const key = note.row * 16384 + note.col;
      if (seen.has(key)) continue;
      seen.add(key);
      const shape = boxes.get(note.object);
      const comment: SheetComment = { row: note.row, col: note.col, text: '' };
      if (shape && shape.text) {
        const rich = richText(shape.text, context.fontOf);
        comment.text = rich.text;
        if (rich.runs) comment.runs = rich.runs;
      }
      if (note.author !== '') comment.author = note.author.slice(0, 255);
      if (note.visible) comment.visible = true;
      if (shape && shape.anchor) comment.box = { kind: 'twoCell', from: shape.anchor.from, to: shape.anchor.to };
      comments.push(comment);
    }
    if (comments.length > 0) out.comments = comments;
    if (this.skipped > 0) out.warnings = [`Drawing: ${this.skipped} object(s) could not be read and are not shown`];
    if (this.chartProblems.length > 0) {
      out.warnings = (out.warnings || []).concat(this.chartProblems);
      if (this.chartsUnsaid > 0) out.warnings.push(`More charts that are not shown: ${this.chartsUnsaid}`);
    }
    return out;
  }

  /** The shapes of the drawing, top level ones; the n-th client data and text box take the n-th OBJ and TXO record. */
  private shapes(): RawShape[] {
    const b = this.art.bytes();
    let count = 0;
    let objectIndex = 0;
    let textIndex = 0;

    const readShape = (container: ArtRecord): RawShape => {
      const shape: RawShape = { kind: 0, flags: 0, values: new Map(), texts: new Map() };
      // Each round takes at least the eight bytes of a header.
      for (let p = container.start; p < container.end; ) {
        const record = recordAt(b, p, container.end);
        if (!record) break;
        const size = record.end - record.start;
        if (record.type === RT_SP && size >= 8) {
          shape.kind = record.instance;
          shape.flags = u32(b, record.start + 4);
        } else if (record.type === RT_SPGR && size >= 16) {
          shape.space = [i32(b, record.start), i32(b, record.start + 4), i32(b, record.start + 8), i32(b, record.start + 12)];
        } else if (record.type === RT_OPT || record.type === RT_OPT2 || record.type === RT_OPT3) readOptions(b, record, shape);
        else if (record.type === RT_CLIENT_ANCHOR && size >= 18) shape.anchor = clientAnchor(b, record.start);
        else if (record.type === RT_CHILD_ANCHOR && size >= 16) {
          shape.child = [i32(b, record.start), i32(b, record.start + 4), i32(b, record.start + 8), i32(b, record.start + 12)];
        } else if (record.type === RT_CLIENT_DATA) shape.object = this.objects[objectIndex++];
        else if (record.type === RT_CLIENT_TEXTBOX) shape.text = this.texts[textIndex++];
        p = record.end;
      }
      return shape;
    };

    /** The shapes of a group container: the first one is the group itself, the others are its members. */
    const group = (container: ArtRecord, depth: number): RawShape | undefined => {
      let first: RawShape | undefined;
      for (let p = container.start; p < container.end && count < XLS_MAX_SHAPES; ) {
        const record = recordAt(b, p, container.end);
        if (!record) break;
        let shape: RawShape | undefined;
        if (record.type === RT_SP_CONTAINER) {
          shape = readShape(record);
          count++;
        } else if (record.type === RT_SPGR_CONTAINER && depth < MAX_ART_DEPTH) shape = group(record, depth + 1);
        if (shape) {
          if (!first) {
            first = shape;
            first.children = [];
          } else (first.children as RawShape[]).push(shape);
        }
        p = record.end;
      }
      return first;
    };

    const top: RawShape[] = [];
    const drawing = (container: ArtRecord) => {
      for (let p = container.start; p < container.end; ) {
        const record = recordAt(b, p, container.end);
        if (!record) break;
        if (record.type === RT_SPGR_CONTAINER) {
          // The outermost group stands for the sheet; its members are the shapes of the sheet.
          const patriarch = group(record, 0);
          if (patriarch && patriarch.children) for (const shape of patriarch.children) top.push(shape);
        } else if (record.type === RT_SP_CONTAINER && count < XLS_MAX_SHAPES) {
          top.push(readShape(record));
          count++;
        }
        p = record.end;
      }
    };
    for (let p = 0; p < b.length; ) {
      const record = recordAt(b, p, b.length);
      if (!record) break;
      if (record.type === RT_DG_CONTAINER) drawing(record);
      // Shapes added to a drawing that was saved before: more group and shape containers after the drawing container.
      else if (record.type === RT_SPGR_CONTAINER) {
        const shape = group(record, 1);
        if (shape) top.push(shape);
      } else if (record.type === RT_SP_CONTAINER && count < XLS_MAX_SHAPES) {
        top.push(readShape(record));
        count++;
      }
      p = record.end;
    }
    return top;
  }

  private objectOf(shape: RawShape, context: XlsArtContext, depth: number): DrawingObject | undefined {
    const ot = shape.object ? shape.object.type : -1;
    if (CONTROLS[ot] || ot === OT_NOTE) return undefined;
    const base: DrawingObjectBase = {};
    const name = shape.texts.get(P_NAME);
    if (name !== undefined && name !== '') base.name = name.slice(0, 255);
    const description = shape.texts.get(P_DESCRIPTION);
    if (description !== undefined && description !== '') base.description = description.slice(0, 2000);
    const groupFlags = shape.values.get(P_GROUP_FLAGS);
    if (groupFlags !== undefined && (groupFlags & 0x20000) !== 0 && (groupFlags & 0x2) !== 0) base.hidden = true;
    const rotation = shape.values.get(P_ROTATION);
    if (rotation !== undefined) {
      // 16.16 fixed point degrees, signed.
      const degrees = (((rotation | 0) / 65536) % 360 + 360) % 360;
      if (degrees !== 0) base.rotation = degrees;
    }
    if (shape.flags & 0x40) base.flipH = true;
    if (shape.flags & 0x80) base.flipV = true;

    if (shape.children) {
      const group: DrawingGroup = { ...base, type: 'group', children: [] };
      const space = shape.space;
      if (!space || depth >= MAX_ART_DEPTH) return group;
      const width = space[2] - space[0];
      const height = space[3] - space[1];
      if (!(width > 0 && height > 0)) return group;
      for (const member of shape.children) {
        const child = member.child;
        if (!child) continue;
        const object = this.objectOf(member, context, depth + 1);
        if (!object) continue;
        const box = { x: (child[0] - space[0]) / width, y: (child[1] - space[1]) / height, width: (child[2] - child[0]) / width, height: (child[3] - child[1]) / height };
        if (!(box.width >= 0 && box.height >= 0 && Math.abs(box.x) < 1e6 && Math.abs(box.y) < 1e6 && box.width < 1e6 && box.height < 1e6)) continue;
        object.box = box;
        group.children.push(object);
      }
      return group;
    }
    if (ot === OT_CHART) {
      // A chart that cannot be read for certain stays a placeholder, and the sheet says why.
      const read = context.chart ? context.chart(shape.object ? shape.object.chart : undefined) : undefined;
      if (read && read.chart) return { ...base, type: 'chart', chart: read.chart };
      if (read) {
        if (this.chartProblems.length < MAX_CHART_WARNINGS) this.chartProblems.push(`Chart${base.name !== undefined ? ` "${base.name.slice(0, 40)}"` : ''} is not shown: ${read.problem || 'its records cannot be read'}`);
        else this.chartsUnsaid++;
      }
      return { ...base, type: 'placeholder', what: 'chart' };
    }
    if (ot === OT_GROUP) return undefined;

    const pib = shape.values.get(P_PICTURE);
    if (pib !== undefined || ot === OT_PICTURE || shape.kind === 75) {
      const picture: DrawingPicture = { ...base, type: 'picture' };
      const key = pib !== undefined && pib > 0 ? context.pictures.key(pib) : undefined;
      if (key !== undefined) picture.media = key;
      const side = (id: number): number => {
        const value = shape.values.get(id);
        // 16.16 fixed point share of the picture, signed; a negative one adds room, which is not shown.
        return value === undefined ? 0 : Math.min(Math.max((value | 0) / 65536, 0), 1);
      };
      const crop = { left: side(P_CROP_LEFT), top: side(P_CROP_TOP), right: side(P_CROP_RIGHT), bottom: side(P_CROP_BOTTOM) };
      if ((crop.left > 0 || crop.top > 0 || crop.right > 0 || crop.bottom > 0) && crop.left + crop.right < 1 && crop.top + crop.bottom < 1) picture.crop = crop;
      return picture;
    }

    const geometry = Object.prototype.hasOwnProperty.call(GEOMETRY, shape.kind) ? GEOMETRY[shape.kind] : 'rect';
    const out: DrawingShape = { ...base, type: 'shape', geometry };
    const line = shape.kind === 20 || (shape.kind >= 32 && shape.kind <= 40);
    if (line) out.connector = true;
    // A shape is filled white and outlined black unless it says otherwise.
    const fillFlags = shape.values.get(P_FILL_FLAGS);
    const filled = fillFlags !== undefined && (fillFlags & 0x100000) !== 0 ? (fillFlags & 0x10) !== 0 : true;
    if (filled && !line) {
      const fill = artColor(shape.values.get(P_FILL_COLOR), context.styles) || '#FFFFFF';
      out.fill = fill;
    }
    const lineFlags = shape.values.get(P_LINE_FLAGS);
    const outlined = lineFlags !== undefined && (lineFlags & 0x80000) !== 0 ? (lineFlags & 0x8) !== 0 : true;
    if (outlined) {
      const width = shape.values.get(P_LINE_WIDTH);
      const pixels = width === undefined ? 1 : Math.min(Math.max(width / EMU_PER_PIXEL, 0.75), 32);
      out.line = { color: artColor(shape.values.get(P_LINE_COLOR), context.styles) || '#000000', width: pixels };
      const dashing = shape.values.get(P_LINE_DASHING);
      if (dashing !== undefined && dashing !== 0) out.line.dash = true;
    }
    if (shape.text && shape.text.text !== '') {
      const text = shapeText(shape.text, context);
      // The four insets follow each other from the left one on; EMU.
      const insets: [number, number, number, number] = [0, 0, 0, 0];
      const defaults = [91440, 45720, 91440, 45720];
      for (let i = 0; i < 4; i++) {
        const stated = shape.values.get(P_TEXT_LEFT + i);
        const emu = stated === undefined ? defaults[i] : stated | 0;
        insets[i] = Math.min(Math.max(emu, 0), EMU_PER_PIXEL * 2000) / EMU_PER_PIXEL;
      }
      text.insets = insets;
      out.text = text;
    }
    return out;
  }
}

/** Reads the property table of a shape into its maps. */
function readOptions(b: Uint8Array, record: ArtRecord, shape: RawShape): void {
  const count = Math.min(record.instance, Math.floor((record.end - record.start) / 6));
  // The values of "complex" properties follow the table, in the order of the table.
  let complex = record.start + count * 6;
  for (let i = 0; i < count; i++) {
    const p = record.start + i * 6;
    const head = u16(b, p);
    const id = head & 0x3fff;
    const value = u32(b, p + 2);
    if (head & 0x8000) {
      if (complex + value > record.end) return;
      if (id === P_NAME || id === P_DESCRIPTION) {
        let text = '';
        for (let k = 0; k + 1 < value && k < 4096; k += 2) {
          const code = u16(b, complex + k);
          if (code === 0) break;
          text += String.fromCharCode(code);
        }
        shape.texts.set(id, text);
      }
      complex += value;
    } else shape.values.set(id, value);
  }
}

/**
 * Where a shape is on its sheet (OfficeArtClientAnchorSheet): flags, then the
 * column, the offset in 1/1024 of its width, the row and the offset in 1/256
 * of its height, for the top left and the bottom right corner.
 */
function clientAnchor(b: Uint8Array, at: number): DrawingAnchor | undefined {
  const flags = u16(b, at);
  const share = (value: number, unit: number): number => Math.min(Math.max(value / unit, 0), 1);
  const from = { col: u16(b, at + 2), colOff: 0, colShare: share(i16(b, at + 4), 1024), row: u16(b, at + 6), rowOff: 0, rowShare: share(i16(b, at + 8), 256) };
  const to = { col: u16(b, at + 10), colOff: 0, colShare: share(i16(b, at + 12), 1024), row: u16(b, at + 14), rowOff: 0, rowShare: share(i16(b, at + 16), 256) };
  if (from.col >= 256 || to.col >= 256) return undefined;
  const anchor: DrawingAnchor = { kind: 'twoCell', from, to };
  // Bit 0: the shape does not move with cells; bit 1: it does not size with them.
  if (flags & 1) anchor.editAs = 'absolute';
  else if (flags & 2) anchor.editAs = 'oneCell';
  return anchor;
}

/**
 * A colour of a shape property (OfficeArtCOLORREF) as "#RRGGBB": red, green and
 * blue in the three low bytes, or an index into the palette of the workbook;
 * undefined for a colour of the system, which has no fixed value.
 */
function artColor(value: number | undefined, styles: StyleSheet): string | undefined {
  if (value === undefined) return undefined;
  const flags = value >>> 24;
  if (flags & 0x08) {
    const index = value & 0xff;
    // 64 and 65: the text and the window colour of the system.
    if (index === 64) return '#000000';
    if (index === 65) return '#FFFFFF';
    return resolveColor({ indexed: index }, styles);
  }
  if (flags & 0x10) return undefined;
  let out = '#';
  for (let i = 0; i < 3; i++) {
    const channel = (value >>> (i * 8)) & 0xff;
    out += (channel < 16 ? '0' : '') + channel.toString(16).toUpperCase();
  }
  return out;
}

function richText(raw: RawText, fontOf: (index: number) => Font | undefined): { text: string; runs?: RichRun[] } {
  const source: RawString = raw.runs ? { text: raw.text, runs: raw.runs } : { text: raw.text };
  const shared = toSharedString(source, fontOf);
  return shared.runs ? { text: shared.text, runs: shared.runs } : { text: shared.text };
}

const ALIGN: (DrawingText['paragraphs'][number]['align'] | undefined)[] = [undefined, undefined, 'center', 'right', 'justify'];
const ANCHOR: (DrawingText['anchor'] | undefined)[] = [undefined, undefined, 'middle', 'bottom'];

function shapeText(raw: RawText, context: XlsArtContext): DrawingText {
  const rich = richText(raw, context.fontOf);
  const runs: DrawingTextRun[] = [];
  const pieces: RichRun[] = rich.runs || [{ text: rich.text, font: raw.runs && raw.runs.length >= 2 ? context.fontOf(raw.runs[1]) : undefined }];
  for (const piece of pieces) {
    const run: DrawingTextRun = { text: piece.text };
    const font = piece.font;
    if (font) {
      if (font.name !== undefined) run.font = font.name;
      if (font.size !== undefined) run.size = font.size;
      if (font.bold) run.bold = true;
      if (font.italic) run.italic = true;
      if (font.underline !== 'none') run.underline = true;
      if (font.strike) run.strike = true;
      const color = resolveColor(font.color, context.styles);
      if (color) run.color = color;
    }
    runs.push(run);
  }
  const paragraph: DrawingText['paragraphs'][number] = { runs };
  const align = ALIGN[raw.align];
  if (align) paragraph.align = align;
  const text: DrawingText = { paragraphs: [paragraph] };
  const anchor = ANCHOR[raw.anchor];
  if (anchor) text.anchor = anchor;
  return text;
}
