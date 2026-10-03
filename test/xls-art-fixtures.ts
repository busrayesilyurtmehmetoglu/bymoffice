// OfficeArt records for the xls tests: the drawing of a sheet with its OBJ, TXO
// and NOTE records, and the picture store of the workbook. Everything is built
// here, record by record.

import { cat, narrow, rec, str16, u16, u32, type Bytes } from './xls-fixtures';

// --- OfficeArt ---------------------------------------------------------------------------

export const artHeader = (type: number, instance: number, version: number, length: number): number[] => [...u16((instance << 4) | version), ...u16(type), ...u32(length)];
export const art = (type: number, instance: number, version: number, content: Bytes = []): number[] => [...artHeader(type, instance, version, content.length), ...content];

/** A run of OfficeArt bytes and the BIFF records that stand after it in the sheet. */
export interface Segment {
  art: number[];
  after: Uint8Array[];
}

export const sizeOf = (segments: Segment[]): number => segments.reduce((sum, segment) => sum + segment.art.length, 0);

/** A container around segments: its header counts all the OfficeArt bytes, whatever records stand between them. */
export function wrap(type: number, segments: Segment[]): Segment[] {
  const head = artHeader(type, 0, 0xf, sizeOf(segments));
  return [{ art: [...head, ...segments[0].art], after: segments[0].after }, ...segments.slice(1)];
}

export type Prop = [id: number, value: number] | [id: number, text: string];

export function options(props: Prop[]): number[] {
  const table: number[] = [];
  const complex: number[] = [];
  for (const [id, value] of props) {
    if (typeof value === 'string') {
      const data: number[] = [];
      for (let i = 0; i < value.length; i++) data.push(value.charCodeAt(i) & 0xff, value.charCodeAt(i) >> 8);
      data.push(0, 0);
      table.push(...u16(id | 0x8000), ...u32(data.length));
      complex.push(...data);
    } else table.push(...u16(id), ...u32(value));
  }
  return art(0xf00b, props.length, 3, [...table, ...complex]);
}

/** Flags, then column, offset (1/1024), row, offset (1/256) of the top left and the bottom right corner. */
export const clientAnchor = (from: [number, number, number, number], to: [number, number, number, number], flags = 0): number[] =>
  art(0xf010, 0, 0, [...u16(flags), ...u16(from[0]), ...u16(from[1]), ...u16(from[2]), ...u16(from[3]), ...u16(to[0]), ...u16(to[1]), ...u16(to[2]), ...u16(to[3])]);

export const childAnchor = (left: number, top: number, right: number, bottom: number): number[] => art(0xf00f, 0, 0, [...u32(left), ...u32(top), ...u32(right), ...u32(bottom)]);

export const obj = (type: number, id: number): Uint8Array => rec(0x005d, u16(0x15), u16(0x12), u16(type), u16(id), u16(0x6011), new Array(12).fill(0), u16(0), u16(0));

/** TXO with its two CONTINUE records. `runs`: first character and font index of each run. */
export function txo(text: string, runs: [number, number][] = [[0, 0]], flags = 0x0212): Uint8Array[] {
  const formatting = [...runs, [text.length, 0]].flatMap(([first, fontIndex]) => [...u16(first), ...u16(fontIndex), ...u32(0)]);
  return [rec(0x01b6, u16(flags), u16(0), new Array(6).fill(0), u16(text.length), u16(formatting.length), u32(0)), rec(0x003c, [0], narrow(text)), rec(0x003c, formatting)];
}

export const note = (row: number, col: number, id: number, author: string, visible = false): Uint8Array => rec(0x001c, u16(row), u16(col), u16(visible ? 2 : 0), u16(id), str16(author), [0]);

export interface ShapeSpec {
  /** MSOSPT: 75 picture frame, 1 rectangle, 3 ellipse, 20 line, 202 text box. */
  kind: number;
  /** Object type of the OBJ record: 8 picture, 2 rectangle, 6 text, 0x19 note, 5 chart, 0x14 drop-down. */
  ot: number;
  id: number;
  props?: Prop[];
  anchor?: number[];
  flags?: number;
  text?: Uint8Array[];
  /** Records after the OBJ record (the substream of a chart). */
  extra?: Uint8Array[];
}

export function shape(spec: ShapeSpec): Segment[] {
  const body = [...art(0xf00a, spec.kind, 2, [...u32(1024 + spec.id), ...u32(spec.flags ?? 0xa00)]), ...options(spec.props ?? []), ...(spec.anchor ?? clientAnchor([1, 0, 1, 0], [3, 0, 5, 0])), ...art(0xf011, 0, 0)];
  const first: Segment = { art: body, after: [obj(spec.ot, spec.id), ...(spec.extra ?? [])] };
  const segments = spec.text ? [first, { art: art(0xf00d, 0, 0), after: spec.text }] : [first];
  return wrap(0xf004, segments);
}

/** A group: the shape that stands for it (with the coordinates of its members), then the members. */
export function groupOf(id: number, space: [number, number, number, number], anchor: number[], members: Segment[][]): Segment[] {
  const self: Segment = {
    art: [...art(0xf009, 0, 1, [...u32(space[0]), ...u32(space[1]), ...u32(space[2]), ...u32(space[3])]), ...art(0xf00a, 0, 2, [...u32(1024 + id), ...u32(0x201)]), ...anchor, ...art(0xf011, 0, 0)],
    after: [obj(0, id)],
  };
  return wrap(0xf003, [...wrap(0xf004, [self]), ...members.flat()]);
}

/** The records of a sheet's drawing: the drawing container cut into MSODRAWING records with the object records between them. */
export function drawingRecords(shapes: Segment[][]): Uint8Array {
  const patriarch: Segment = { art: art(0xf004, 0, 0xf, [...art(0xf009, 0, 1, new Array(16).fill(0)), ...art(0xf00a, 0, 2, [...u32(1024), ...u32(5)])]), after: [] };
  const group = wrap(0xf003, [patriarch, ...shapes.flat()]);
  // The patriarch has no object record: its bytes go into the same record as the first shape.
  const merged = group.length > 1 ? [{ art: [...group[0].art, ...group[1].art], after: group[1].after }, ...group.slice(2)] : group;
  const dg = wrap(0xf002, [{ art: art(0xf008, 1, 0, [...u32(shapes.length + 1), ...u32(1024 + shapes.length)]), after: [] }, ...merged]);
  const all = dg.length > 1 ? [{ art: [...dg[0].art, ...dg[1].art], after: dg[1].after }, ...dg.slice(2)] : dg;
  return cat(...all.map((segment) => cat(rec(0x00ec, segment.art), ...segment.after)));
}

/** A BLIP inside its store entry: a PNG (type 6) unless stated. */
export function bse(data: Uint8Array, blipType = 0xf01e, kind = 6): number[] {
  const blip = blipType === 0xf01a ? art(blipType, 0x3d4, 0, [...new Array(16).fill(1), ...new Array(34).fill(0), ...data]) : art(blipType, 0x6e0, 0, [...new Array(16).fill(1), 0xff, ...data]);
  return art(0xf007, kind, 2, [kind, kind, ...new Array(16).fill(1), ...u16(0xff), ...u32(blip.length), ...u32(1), ...u32(0), 0, 0, 0, 0, ...blip]);
}

/** The drawing group of the globals with its picture store, as one run of OfficeArt bytes. */
export function drawingGroup(entries: number[][]): number[] {
  const store = art(0xf001, entries.length, 0xf, entries.flat());
  return art(0xf000, 0, 0xf, [...art(0xf006, 0, 0, new Array(24).fill(0)), ...store]);
}
