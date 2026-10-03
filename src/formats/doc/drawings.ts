// Pictures and drawing objects of a Word binary file.
//
// An inline picture is a 0x01 character whose properties point to a PICF in
// the Data stream (MS-DOC 2.9.191): size, scaling and cropping, followed by
// an OfficeArt shape and the picture itself (a "BLIP": PNG, JPEG, DIB, EMF,
// WMF, ...).
//
// A drawing object (a floating picture, a text box, a shape) is a 0x08
// character; PlcSpaMom / PlcSpaHdr say where it sits (Spa, MS-DOC 2.9.253) and
// which shape of the OfficeArt drawing (DggInfo, MS-ODRAW) it is. The pictures
// of those shapes are in a store shared by the document.
//
// Nothing in here is trusted: record lengths are checked against the record
// around them, nesting and counts are bounded, and a picture that says it
// unpacks to more than MAX_BLIP bytes is not unpacked.

import { inflateRaw } from '../../deflate';
import type { Block, DocMedia, DrawingAnchor, DrawingItem, DrawingPosition, ImageCrop, ImageRef } from '../../model/doc/types';
import { officeArtLook } from '../../model/doc/shapes';
import type { XmlElement } from '../../xml';
import { ansiText, i16, i32, u16, u32, u8 } from './bytes';

/** Largest picture that is unpacked, bytes. */
const MAX_BLIP = 64 * 1024 * 1024;
/** Bytes all pictures of one document may take that are not just a view of the file: unpacked metafiles, bitmaps given a file header. */
const MEDIA_BUDGET = 256 * 1024 * 1024;

/** How much of MEDIA_BUDGET is left. */
export interface MediaBudget {
  left: number;
}
/** Deepest nesting of OfficeArt containers that is followed. */
const MAX_ART_DEPTH = 32;
/** Pictures and text boxes one drawing object shows: those of a group are gathered, and laid out, for each use of it. */
const MAX_SHAPE_PARTS = 256;
/** Most shapes of one document that are read. */
const MAX_SHAPES = 100000;
const TWIP = 635; // EMU

const RT_DGG_CONTAINER = 0xf000;
const RT_BSTORE = 0xf001;
const RT_DG_CONTAINER = 0xf002;
const RT_SPGR_CONTAINER = 0xf003;
const RT_SP_CONTAINER = 0xf004;
const RT_BSE = 0xf007;
const RT_SP = 0xf00a;
const RT_OPT = 0xf00b;
const RT_OPT2 = 0xf121;
const RT_OPT3 = 0xf122;
const RT_BLIP_FIRST = 0xf018;
const RT_BLIP_LAST = 0xf117;

interface BlipKind {
  type: string;
  extension: string;
  metafile: boolean;
}

const BLIPS: Record<number, BlipKind> = {
  0xf01a: { type: 'image/x-emf', extension: 'emf', metafile: true },
  0xf01b: { type: 'image/x-wmf', extension: 'wmf', metafile: true },
  0xf01c: { type: 'image/pict', extension: 'pict', metafile: true },
  0xf01d: { type: 'image/jpeg', extension: 'jpg', metafile: false },
  0xf01e: { type: 'image/png', extension: 'png', metafile: false },
  0xf01f: { type: 'image/bmp', extension: 'bmp', metafile: false },
  0xf029: { type: 'image/tiff', extension: 'tif', metafile: false },
  0xf02a: { type: 'image/jpeg', extension: 'jpg', metafile: false },
};

/** A device independent bitmap as a BMP file: the file header it lacks is put in front. */
function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  if (dib.length < 16) return undefined;
  const headerSize = u32(dib, 0);
  if (headerSize < 12 || headerSize > dib.length) return undefined;
  let colors = 0;
  if (headerSize >= 40) {
    const bits = u16(dib, 14);
    const compression = u32(dib, 16);
    colors = u32(dib, 32);
    if (colors === 0 && bits <= 8) colors = 1 << bits;
    if (compression === 3 && headerSize === 40) colors += 3;
    if (colors > 256 + 3) colors = 0;
  } else {
    const bits = u16(dib, 10);
    if (bits <= 8) colors = 1 << bits;
  }
  const pixels = 14 + headerSize + colors * (headerSize >= 40 ? 4 : 3);
  const out = new Uint8Array(14 + dib.length);
  out[0] = 0x42;
  out[1] = 0x4d;
  const size = out.length;
  out[2] = size & 0xff;
  out[3] = (size >>> 8) & 0xff;
  out[4] = (size >>> 16) & 0xff;
  out[5] = (size >>> 24) & 0xff;
  out[10] = pixels & 0xff;
  out[11] = (pixels >>> 8) & 0xff;
  out[12] = (pixels >>> 16) & 0xff;
  out[13] = (pixels >>> 24) & 0xff;
  out.set(dib, 14);
  return out;
}

/**
 * The picture of a BLIP record at `at` (its header included) in `b`, which
 * must end by `end`; undefined when the record is not a picture the reader
 * knows or is damaged.
 */
export function readBlip(b: Uint8Array, at: number, end: number, budget: MediaBudget): DocMedia | undefined {
  if (at + 8 > end) return undefined;
  const kind = BLIPS[u16(b, at + 2)];
  if (!kind) return undefined;
  const instance = u16(b, at) >> 4;
  const stop = at + 8 + u32(b, at + 4);
  if (stop > end || stop > b.length) return undefined;
  // One identifier of 16 bytes, two when the instance is odd.
  let p = at + 8 + (instance & 1 ? 32 : 16);
  if (!kind.metafile) {
    p += 1;
    if (p >= stop) return undefined;
    const bytes = b.subarray(p, stop);
    if (kind.extension === 'bmp') {
      if (bytes.length + 14 > budget.left) return undefined;
      const bmp = dibToBmp(bytes);
      if (bmp) budget.left -= bmp.length;
      return bmp ? { bytes: bmp, type: kind.type } : undefined;
    }
    return { bytes, type: kind.type };
  }
  // A metafile: its unpacked size, bounds and size, the packed size, then how it is packed.
  if (p + 34 > stop) return undefined;
  const size = u32(b, p);
  const packed = u32(b, p + 28);
  const compression = u8(b, p + 32);
  p += 34;
  const data = b.subarray(p, Math.min(stop, p + packed));
  if (compression === 0xfe) return { bytes: data, type: kind.type };
  if (compression !== 0 || size > MAX_BLIP || size > budget.left || data.length < 6) return undefined;
  // What the picture says it unpacks to is taken from the budget before the attempt: one that fails has cost as much.
  budget.left -= Math.max(size, 4096);
  try {
    // A zlib stream: two bytes of header, the deflate data, a checksum. Unpacking stops with an error at the size stated.
    return { bytes: inflateRaw(data.subarray(2), size), type: kind.type };
  } catch (error) {
    return undefined;
  }
}

/** Properties of a shape (MS-ODRAW 2.3) the reader looks at, by property id. */
const P_ROTATION = 0x0004;
const P_TEXT_ID = 0x0080;
const P_TEXT_LEFT = 0x0081;
const P_CROP_TOP = 0x0100;
const P_PICTURE = 0x0104;
const P_ADJUST = 0x0147;
const P_FILL_TYPE = 0x0180;
const P_FILL_COLOR = 0x0181;
const P_FILL_OPACITY = 0x0182;
/** The kind of shape that shows a picture (msosptPictureFrame). */
const SHAPE_PICTURE_FRAME = 75;
/** Boolean properties come in groups: the low bits are the values, the bits sixteen above them say which values are stated. */
const P_FILL_FLAGS = 0x01bf;
const P_LINE_COLOR = 0x01c0;
const P_LINE_WIDTH = 0x01cb;
const P_LINE_DASHING = 0x01ce;
const P_LINE_FLAGS = 0x01ff;
const P_NAME = 0x0380;
const P_DESCRIPTION = 0x0381;
const P_WRAP_LEFT = 0x0384;
const P_POS_H = 0x038f;
const P_POS_REL_H = 0x0390;
const P_POS_V = 0x0391;
const P_POS_REL_V = 0x0392;
const P_GROUP_FLAGS = 0x03bf;

export interface Shape {
  id: number;
  /** The kind of shape (MSOSPT): 1 a rectangle, 3 an ellipse, 20 a line, 75 a picture, 202 a text box, ... */
  type: number;
  group: boolean;
  flipH: boolean;
  flipV: boolean;
  /** Simple properties by id. */
  values: Map<number, number>;
  /** Text properties (names, descriptions) by id. */
  texts: Map<number, string>;
  /** Shapes of a group, groups in it included. */
  children: Shape[];
}

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

/** Reads the property table of a shape into its maps. */
function readOptions(b: Uint8Array, record: ArtRecord, shape: Shape): void {
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
        for (let k = 0; k + 1 < value && k < 2048; k += 2) {
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

function readShape(b: Uint8Array, container: ArtRecord): Shape {
  const shape: Shape = { id: 0, type: 0, group: false, flipH: false, flipV: false, values: new Map(), texts: new Map(), children: [] };
  let p = container.start;
  // Each round takes at least the eight bytes of a header.
  while (p < container.end) {
    const record = recordAt(b, p, container.end);
    if (!record) break;
    if (record.type === RT_SP && record.end - record.start >= 8) {
      shape.id = u32(b, record.start);
      shape.type = record.instance;
      const flags = u32(b, record.start + 4);
      shape.group = (flags & 0x01) !== 0;
      shape.flipH = (flags & 0x40) !== 0;
      shape.flipV = (flags & 0x80) !== 0;
    } else if (record.type === RT_OPT || record.type === RT_OPT2 || record.type === RT_OPT3) readOptions(b, record, shape);
    p = record.end;
  }
  return shape;
}

/** What a shape shows: its pictures and the texts of its text boxes, and those of the shapes grouped in it. */
interface ShapeContent {
  images: ImageRef[];
  boxes: Block[][];
}

export interface Drawings {
  /** Shapes that a 0x08 character may refer to, by shape id. */
  shapes: Map<number, Shape>;
  /** What each shape shows, gathered the first time a character refers to it. */
  content: Map<number, ShapeContent>;
  /** The picture store: where each picture is, by index (the "pib" of a shape is the index plus one). */
  blips: { stream: 'word' | 'table'; at: number; end: number }[];
}

/** Reads the OfficeArt content of the document (DggInfo) at [fc, fc + lcb) of the table stream. */
export function readDrawings(table: Uint8Array, word: Uint8Array, fc: number, lcb: number): Drawings {
  const out: Drawings = { shapes: new Map(), content: new Map(), blips: [] };
  const end = fc + lcb;
  if (lcb < 8 || fc < 0 || end > table.length) return out;
  let count = 0;

  /** Shapes of a group container, in order: the first is the group itself. */
  const group = (container: ArtRecord, depth: number): Shape | undefined => {
    let first: Shape | undefined;
    let p = container.start;
    while (p < container.end && count < MAX_SHAPES) {
      const record = recordAt(table, p, container.end);
      if (!record) break;
      let shape: Shape | undefined;
      if (record.type === RT_SP_CONTAINER) {
        shape = readShape(table, record);
        count++;
      } else if (record.type === RT_SPGR_CONTAINER && depth < MAX_ART_DEPTH) shape = group(record, depth + 1);
      if (shape) {
        if (!first) first = shape;
        else first.children.push(shape);
        if (shape.id !== 0 && !out.shapes.has(shape.id)) out.shapes.set(shape.id, shape);
      }
      p = record.end;
    }
    return first;
  };

  const store = (container: ArtRecord) => {
    let p = container.start;
    while (p < container.end && out.blips.length < MAX_SHAPES) {
      const record = recordAt(table, p, container.end);
      if (!record) break;
      if (record.type === RT_BSE && record.end - record.start >= 36) {
        const nameLength = u8(table, record.start + 33);
        const embedded = record.start + 36 + nameLength;
        const size = u32(table, record.start + 20);
        const delay = u32(table, record.start + 28);
        // The picture follows its entry here, or lies in the WordDocument stream at the offset the entry gives.
        if (embedded + 8 <= record.end) out.blips.push({ stream: 'table', at: embedded, end: record.end });
        else if (delay !== 0xffffffff && size >= 8 && delay + size <= word.length) out.blips.push({ stream: 'word', at: delay, end: delay + size });
        else out.blips.push({ stream: 'word', at: 0, end: 0 });
      } else if (record.type >= RT_BLIP_FIRST && record.type <= RT_BLIP_LAST) out.blips.push({ stream: 'table', at: p, end: record.end });
      else out.blips.push({ stream: 'word', at: 0, end: 0 });
      p = record.end;
    }
  };

  let p = fc;
  const head = recordAt(table, p, end);
  if (head && head.type === RT_DGG_CONTAINER) {
    let q = head.start;
    while (q < head.end) {
      const inner = recordAt(table, q, head.end);
      if (!inner) break;
      if (inner.type === RT_BSTORE) store(inner);
      q = inner.end;
    }
    p = head.end;
  }
  // The drawings (of the main text, of the headers): one byte that says which, then a drawing container.
  while (p + 9 <= end) {
    const record = recordAt(table, p + 1, end);
    if (!record || record.type !== RT_DG_CONTAINER) break;
    let q = record.start;
    while (q < record.end) {
      const inner = recordAt(table, q, record.end);
      if (!inner) break;
      if (inner.type === RT_SPGR_CONTAINER) group(inner, 0);
      else if (inner.type === RT_SP_CONTAINER) {
        const shape = readShape(table, inner);
        if (shape.id !== 0 && !out.shapes.has(shape.id)) out.shapes.set(shape.id, shape);
      }
      q = inner.end;
    }
    p = record.end;
  }
  return out;
}

/** Hands out pictures as media of the model: each picture once, under a part name of its own. */
export class MediaStore {
  readonly media = new Map<string, DocMedia>();
  readonly budget: MediaBudget = { left: MEDIA_BUDGET };
  /** Inline pictures by their place in the Data stream: a place that many characters refer to is read once. */
  readonly inline = new Map<number, DrawingItem | null>();
  private readonly known = new Map<string, ImageRef | null>();

  /** `key` names the place of the picture in the file; `read` is called the first time the key is seen. */
  get(key: string, read: () => DocMedia | undefined): ImageRef | undefined {
    const seen = this.known.get(key);
    if (seen !== undefined) return seen === null ? undefined : seen;
    const item = read();
    if (!item) {
      this.known.set(key, null);
      return undefined;
    }
    const number = this.media.size + 1;
    const slash = item.type.indexOf('/');
    const extension = item.type === 'image/jpeg' ? 'jpg' : item.type === 'image/x-emf' ? 'emf' : item.type === 'image/x-wmf' ? 'wmf' : item.type.slice(slash + 1);
    const target = `/doc/media/image${number}.${extension}`;
    this.media.set(target, item);
    const ref: ImageRef = { rId: 'doc:image' + number, target };
    this.known.set(key, ref);
    return ref;
  }
}

/** A stand-in for the element a docx reader keeps a drawing in: a doc file has none. */
function drawingXml(insets?: (number | undefined)[]): XmlElement {
  const xml: XmlElement = { type: 'element', name: 'doc:drawing', attrs: new Map(), children: [] };
  if (insets && insets.some((value) => value !== undefined)) {
    // The insets of a text box, where layout looks for them in a DrawingML shape.
    const body: XmlElement = { type: 'element', name: 'wps:bodyPr', attrs: new Map(), children: [] };
    const names = ['lIns', 'tIns', 'rIns', 'bIns'];
    insets.forEach((value, i) => {
      if (value !== undefined && value >= 0 && value <= 100000000) body.attrs.set(names[i], String(value));
    });
    xml.children.push(body);
  }
  return xml;
}

const MM_SHAPE = 0x64;
const MM_SHAPE_FILE = 0x66;
const MM_METAFILE = 8;

/**
 * The inline picture at `offset` of the Data stream; undefined when there is
 * no picture there (the place holds the data of a form field or an object,
 * or nothing the reader can make out).
 */
export function inlinePicture(data: Uint8Array, offset: number, store: MediaStore): DrawingItem | undefined {
  const known = store.inline.get(offset);
  if (known !== undefined) return known === null ? undefined : known;
  const item = readInlinePicture(data, offset, store);
  store.inline.set(offset, item || null);
  return item;
}

/** The largest width and height Word gives a picture: 22 inches. */
const MAX_PICTURE_TWIPS = 31680;

/** Segments of a JPEG file that are looked at for its size; a genuine file has the size within the first few. */
const MAX_JPEG_SEGMENTS = 4096;

const be16 = (b: Uint8Array, o: number): number => (b[o] << 8) | b[o + 1];
const be32 = (b: Uint8Array, o: number): number => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;

/**
 * Width and height of a picture as its own header states them: pixels for PNG, JPEG and
 * BMP, the frame for EMF and placeable WMF. Only the proportion of the two is used.
 * Undefined when the header is not there, is cut off, or states no size.
 */
export function pictureSize(media: DocMedia): { width: number; height: number } | undefined {
  const b = media.bytes;
  let width = 0;
  let height = 0;
  if (media.type === 'image/png') {
    // The signature, then the IHDR chunk: length, name, width, height.
    if (b.length < 24 || b[0] !== 0x89 || b[1] !== 0x50 || b[12] !== 0x49 || b[13] !== 0x48 || b[14] !== 0x44 || b[15] !== 0x52) return undefined;
    width = be32(b, 16);
    height = be32(b, 20);
  } else if (media.type === 'image/jpeg') {
    if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return undefined;
    let p = 2;
    // Each round takes a marker and its segment: at least two bytes.
    for (let i = 0; i < MAX_JPEG_SEGMENTS && p + 4 <= b.length; i++) {
      if (b[p] !== 0xff) return undefined;
      const marker = b[p + 1];
      // Fill bytes before a marker, and markers that stand alone.
      if (marker === 0xff) {
        p += 1;
        continue;
      }
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        p += 2;
        continue;
      }
      // The picture data starts, or the file ends: no size was stated.
      if (marker === 0xda || marker === 0xd9) return undefined;
      const length = be16(b, p + 2);
      if (length < 2) return undefined;
      // A frame header (SOF0..SOF15 without the three markers in that range that are something else).
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        if (length < 7 || p + 9 > b.length) return undefined;
        height = be16(b, p + 5);
        width = be16(b, p + 7);
        break;
      }
      p += 2 + length;
    }
  } else if (media.type === 'image/bmp') {
    // The file header, then the bitmap header: the newer one has 32-bit sizes, the older one 16-bit ones.
    if (b.length < 26 || b[0] !== 0x42 || b[1] !== 0x4d) return undefined;
    if (u32(b, 14) >= 40) {
      width = Math.abs(i32(b, 18));
      height = Math.abs(i32(b, 22));
    } else {
      width = u16(b, 18);
      height = u16(b, 20);
    }
  } else if (media.type === 'image/x-emf') {
    // The header record: its frame in hundredths of a millimetre, then the signature.
    if (b.length < 44 || u32(b, 0) !== 1 || u32(b, 40) !== 0x464d4520) return undefined;
    width = i32(b, 32) - i32(b, 24);
    height = i32(b, 36) - i32(b, 28);
  } else if (media.type === 'image/x-wmf') {
    // Only a "placeable" metafile says how large it is.
    if (b.length < 14 || u32(b, 0) !== 0x9ac6cdd7) return undefined;
    width = i16(b, 10) - i16(b, 6);
    height = i16(b, 12) - i16(b, 8);
  }
  return width > 0 && height > 0 ? { width, height } : undefined;
}

function readInlinePicture(data: Uint8Array, offset: number, store: MediaStore): DrawingItem | undefined {
  if (offset < 0 || offset + 68 > data.length) return undefined;
  const size = i32(data, offset);
  const header = u16(data, offset + 4);
  if (size < header || header < 68 || offset + size > data.length) return undefined;
  const end = offset + size;
  const mode = u16(data, offset + 6);
  const goalWidth = u16(data, offset + 28);
  const goalHeight = u16(data, offset + 30);
  const scaleX = u16(data, offset + 32) || 1000;
  const scaleY = u16(data, offset + 34) || 1000;
  const cropLeft = i16(data, offset + 36);
  const cropTop = i16(data, offset + 38);
  const cropRight = i16(data, offset + 40);
  const cropBottom = i16(data, offset + 42);

  const item: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', xml: drawingXml() };
  if ((cropLeft || cropTop || cropRight || cropBottom) && goalWidth > 0 && goalHeight > 0) {
    const crop: ImageCrop = {};
    if (cropLeft) crop.left = Math.round((cropLeft / goalWidth) * 100000);
    if (cropRight) crop.right = Math.round((cropRight / goalWidth) * 100000);
    if (cropTop) crop.top = Math.round((cropTop / goalHeight) * 100000);
    if (cropBottom) crop.bottom = Math.round((cropBottom / goalHeight) * 100000);
    item.crop = crop;
  }

  let p = offset + header;
  let image: ImageRef | undefined;
  if (mode === MM_SHAPE || mode === MM_SHAPE_FILE) {
    if (mode === MM_SHAPE_FILE) {
      // The name of the linked file comes first.
      const length = u8(data, p);
      const name = ansiText(data, p + 1, length);
      if (name) item.name = name;
      p += 1 + length;
    }
    // A shape, then the pictures it uses: entries of a picture store that carry their picture, or pictures as they are.
    let shape: Shape | undefined;
    const pictures: number[] = [];
    while (p < end) {
      const record = recordAt(data, p, end);
      if (!record) break;
      if (record.type === RT_SP_CONTAINER && !shape) shape = readShape(data, record);
      else if (record.type === RT_BSE && record.end - record.start >= 36) {
        const embedded = record.start + 36 + u8(data, record.start + 33);
        pictures.push(embedded + 8 <= record.end ? embedded : -1, record.end);
      } else if (record.type >= RT_BLIP_FIRST && record.type <= RT_BLIP_LAST) pictures.push(p, record.end);
      p = record.end;
    }
    const wanted = shape ? shape.values.get(P_PICTURE) : undefined;
    const index = wanted !== undefined && wanted >= 1 && wanted * 2 <= pictures.length ? wanted - 1 : 0;
    if (index * 2 < pictures.length && pictures[index * 2] >= 0) {
      const at = pictures[index * 2];
      const stop = pictures[index * 2 + 1];
      image = store.get('data:' + at, () => readBlip(data, at, stop, store.budget));
    }
    if (shape) {
      const description = shape.texts.get(P_DESCRIPTION);
      if (description) item.description = description;
      const name = shape.texts.get(P_NAME);
      if (name && item.name === undefined) item.name = name;
    }
  } else if (mode === MM_METAFILE && end > p) {
    // A Windows metafile as it is.
    const from = p;
    image = store.get('data:' + from, () => ({ bytes: data.subarray(from, end), type: 'image/x-wmf' }));
  }
  if (image) {
    item.image = image;
    item.images = item.crop ? [{ image, crop: item.crop }] : [{ image }];
  }

  // The size shown: the natural size less what is cropped away, scaled (in thousandths).
  let width = Math.max(0, ((goalWidth - cropLeft - cropRight) * scaleX) / 1000);
  let height = Math.max(0, ((goalHeight - cropTop - cropBottom) * scaleY) / 1000);
  // No picture is wider or taller than 22 inches in Word. Files of other producers carry a
  // scale on one axis that gives several metres; the picture then keeps its shape at the
  // size the other axis gives. The shape is that of the picture itself where its header says
  // how large it is: the natural size in the file is no measure of it, since it stops at 22
  // inches too and a picture larger than that has lost its proportion there.
  const wide = width > MAX_PICTURE_TWIPS && height <= MAX_PICTURE_TWIPS;
  const tall = height > MAX_PICTURE_TWIPS && width <= MAX_PICTURE_TWIPS;
  if (wide || tall) {
    let ratio = goalWidth > 0 && goalHeight > 0 ? goalWidth / goalHeight : 0;
    const media = image && image.target !== undefined ? store.media.get(image.target) : undefined;
    const size = media ? pictureSize(media) : undefined;
    if (size) {
      ratio = size.width / size.height;
      // What is cropped away, as the share of each side that is left.
      const shownX = goalWidth > 0 ? 1 - (cropLeft + cropRight) / goalWidth : 1;
      const shownY = goalHeight > 0 ? 1 - (cropTop + cropBottom) / goalHeight : 1;
      if (shownX > 0 && shownY > 0) ratio = (ratio * shownX) / shownY;
    }
    if (ratio > 0) {
      if (wide) width = height * ratio;
      else height = width / ratio;
    }
  }
  width = Math.min(width, MAX_PICTURE_TWIPS);
  height = Math.min(height, MAX_PICTURE_TWIPS);
  item.width = Math.round(width * TWIP);
  item.height = Math.round(height * TWIP);
  return item;
}

const POSITION_H = [undefined, 'left', 'center', 'right', 'inside', 'outside'];
const POSITION_V = [undefined, 'top', 'center', 'bottom', 'inside', 'outside'];
const RELATIVE_H = ['margin', 'page', 'column', 'character'];
const RELATIVE_V = ['margin', 'page', 'paragraph', 'line'];
/** Wrapping by the "wr" of a Spa: around both sides, above and below, around, none, tight, through. */
const WRAPS: NonNullable<DrawingAnchor['wrap']>[] = ['square', 'topAndBottom', 'square', 'none', 'tight', 'through'];
const WRAP_SIDES = ['bothSides', 'left', 'right', 'largest'];

const signed = (value: number | undefined): number | undefined => (value === undefined ? undefined : value | 0);

/** Largest length in twips whose EMU value the model holds. */
const MAX_TWIPS = 3000000;
const limited = (twips: number): number => (twips > MAX_TWIPS ? MAX_TWIPS : twips < -MAX_TWIPS ? -MAX_TWIPS : twips);

/**
 * The drawing object a 0x08 character stands for: `spa` is the place of its
 * Spa in the table stream. `text` gives the blocks of the text box of a
 * shape, by the id of the shape and the 1-based number the shape has for its
 * text (0 when it states none); undefined when the shape has no text.
 */
export function floatingDrawing(
  table: Uint8Array,
  word: Uint8Array,
  spa: number,
  drawings: Drawings,
  store: MediaStore,
  text: (shapeId: number, textId: number) => Block[] | undefined,
): DrawingItem {
  const id = u32(table, spa);
  // The rectangle, in twips; kept within what fits into EMU as the model holds them.
  const left = limited(i32(table, spa + 4));
  const top = limited(i32(table, spa + 8));
  const right = limited(i32(table, spa + 12));
  const bottom = limited(i32(table, spa + 16));
  const flags = u16(table, spa + 20);
  const shape = drawings.shapes.get(id);

  const horizontal: DrawingPosition = { relativeFrom: RELATIVE_H[(flags >> 1) & 3] || 'column', offset: left * TWIP };
  const vertical: DrawingPosition = { relativeFrom: RELATIVE_V[(flags >> 3) & 3] || 'paragraph', offset: top * TWIP };
  const anchor: DrawingAnchor = { horizontal, vertical, wrap: WRAPS[(flags >> 5) & 15] || 'none' };
  const side = WRAP_SIDES[(flags >> 9) & 15];
  if (side !== undefined && anchor.wrap !== 'none' && anchor.wrap !== 'topAndBottom') anchor.wrapText = side;
  if (flags & 0x4000) anchor.behindDoc = true;
  if (flags & 0x8000) anchor.locked = true;

  const insets: (number | undefined)[] = [];
  let kind: DrawingItem['kind'] = 'shape';
  let images: ImageRef[] = [];
  let boxes: NonNullable<DrawingItem['textBoxes']> = [];
  const item: DrawingItem = { type: 'drawing', placement: 'anchor', kind, xml: drawingXml() };
  if (shape) {
    const values = shape.values;
    // Word 2000 and later place a shape by these properties; the rectangle of the Spa is what Word 97 would show.
    const relH = values.get(P_POS_REL_H);
    if (relH !== undefined && RELATIVE_H[relH] !== undefined) horizontal.relativeFrom = RELATIVE_H[relH];
    const relV = values.get(P_POS_REL_V);
    if (relV !== undefined && RELATIVE_V[relV] !== undefined) vertical.relativeFrom = RELATIVE_V[relV];
    const posH = values.get(P_POS_H);
    if (posH !== undefined && POSITION_H[posH] !== undefined) {
      horizontal.align = POSITION_H[posH];
      delete horizontal.offset;
    }
    const posV = values.get(P_POS_V);
    if (posV !== undefined && POSITION_V[posV] !== undefined) {
      vertical.align = POSITION_V[posV];
      delete vertical.offset;
    }
    const distances = ['distL', 'distT', 'distR', 'distB'] as const;
    distances.forEach((name, i) => {
      const value = values.get(P_WRAP_LEFT + i);
      if (value !== undefined && value <= 100000000) anchor[name] = value;
    });
    const groupFlags = values.get(P_GROUP_FLAGS);
    // fBehindDocument, when the property says that it states it.
    if (groupFlags !== undefined && groupFlags & 0x00200000) {
      if (groupFlags & 0x20) anchor.behindDoc = true;
      else delete anchor.behindDoc;
    }
    // UNSURE: for a rotation between 45 and 135 or between 225 and 315 degrees other readers of the format take
    // the rectangle of the Spa to be that of the shape turned by a quarter, and swap its width and height. The
    // specification read for this reader does not say so, and no file at hand has a rotated shape: the rectangle
    // is taken as it is.
    const rotation = signed(values.get(P_ROTATION));
    if (rotation || shape.flipH || shape.flipV) {
      item.transform = {};
      // Degrees as a 16.16 fixed point number -> 60000ths of a degree.
      if (rotation) item.transform.rotation = Math.round((rotation / 65536) * 60000);
      if (shape.flipH) item.transform.flipH = true;
      if (shape.flipV) item.transform.flipV = true;
    }
    const name = shape.texts.get(P_NAME);
    if (name) item.name = name;
    const description = shape.texts.get(P_DESCRIPTION);
    if (description) item.description = description;
    for (let i = 0; i < 4; i++) insets.push(values.get(P_TEXT_LEFT + i));

    // The pictures and text boxes of the shape and, for a group, of everything in it. Gathered once for each
    // shape: many characters may refer to the same one.
    let content = drawings.content.get(shape.id);
    if (!content) {
      const found: ShapeContent = { images: [], boxes: [] };
      const visit = (current: Shape, depth: number) => {
        const picture = current.values.get(P_PICTURE);
        if (picture !== undefined && picture >= 1 && picture <= drawings.blips.length) {
          const blip = drawings.blips[picture - 1];
          if (blip.end > blip.at) {
            const source = blip.stream === 'word' ? word : table;
            const image = store.get(blip.stream + ':' + blip.at, () => readBlip(source, blip.at, blip.end, store.budget));
            if (image && found.images.length < MAX_SHAPE_PARTS) found.images.push(image);
          }
        }
        const textId = current.values.get(P_TEXT_ID);
        const blocks = text(current.id, textId === undefined ? 0 : textId >>> 16);
        if (blocks && found.boxes.length < MAX_SHAPE_PARTS && found.boxes.indexOf(blocks) < 0) found.boxes.push(blocks);
        if (depth < MAX_ART_DEPTH) for (const child of current.children) visit(child, depth + 1);
      };
      visit(shape, 0);
      drawings.content.set(shape.id, (content = found));
    }
    // Drawings of the same shape share these lists: they are not copied for each character that refers to it.
    images = content.images;
    boxes = content.boxes;
    if (shape.group) kind = 'group';
    else if (images.length && !boxes.length) kind = 'picture';
    else if (!images.length && (shape.type === SHAPE_PICTURE_FRAME || values.has(P_PICTURE))) {
      // A shape that is to show a picture the file does not have, or the reader cannot read: it is not a box
      // filled in white. Without text it stays a picture, without an image.
      if (!boxes.length) kind = 'picture';
    } else {
      /** A flag of a group of boolean properties; undefined when the group does not state it. */
      const flag = (id: number, bit: number): boolean | undefined => {
        const flags = values.get(id);
        return flags === undefined || !(flags & (bit << 16)) ? undefined : (flags & bit) !== 0;
      };
      item.shape = officeArtLook({
        type: shape.type,
        fFilled: flag(P_FILL_FLAGS, 0x10),
        fillType: values.get(P_FILL_TYPE),
        fillColor: values.get(P_FILL_COLOR),
        fillOpacity: signed(values.get(P_FILL_OPACITY)),
        fLine: flag(P_LINE_FLAGS, 0x08),
        lineColor: values.get(P_LINE_COLOR),
        lineWidth: values.get(P_LINE_WIDTH),
        lineDashing: values.get(P_LINE_DASHING),
        adjust: signed(values.get(P_ADJUST)),
      });
    }

    const cropSides = ['top', 'bottom', 'left', 'right'] as const;
    let crop: ImageCrop | undefined;
    cropSides.forEach((name, i) => {
      const value = signed(values.get(P_CROP_TOP + i));
      // A fraction of the picture as a 16.16 fixed point number -> thousandths of a percent.
      if (value) (crop || (crop = {}))[name] = Math.round((value / 65536) * 100000);
    });
    if (crop && kind === 'picture') item.crop = crop;
  } else {
    // The drawing data does not have the shape; the list of text boxes may still name it.
    const blocks = text(id, 0);
    if (blocks) boxes.push(blocks);
  }
  item.kind = kind;
  item.width = Math.max(0, right - left) * TWIP;
  item.height = Math.max(0, bottom - top) * TWIP;
  item.anchor = anchor;
  if (images.length) {
    item.image = images[0];
    item.images = images.map((image, i) => (i === 0 && item.crop ? { image, crop: item.crop } : { image }));
  }
  if (boxes.length) {
    item.textBoxes = boxes;
    item.xml = drawingXml(insets);
  }
  return item;
}
