// Pictures and shapes of an RTF file.
//
// A picture is a group "{\pict ...}": control words that say what kind it is
// and how large it is shown, then the bytes of the picture as hex digits (or
// as they are, after "\binN"). A shape is a group "{\shp{\*\shpinst ...}}":
// where it sits, a list of named properties ("{\sp{\sn name}{\sv value}}",
// the same names an OfficeArt shape has), the picture it shows as the value
// of the property "pib", and the text of its text box in "{\shptxt ...}".
//
// The bytes of a picture are decoded straight into the buffer the model
// keeps; megabytes of hex digits never become a string. Pictures with the
// same bytes are kept once.

import type { Block, DocMedia, DrawingAnchor, DrawingItem, DrawingPosition, ImageCrop, ImageRef } from '../../model/doc/types';
import { officeArtLook } from '../../model/doc/shapes';
import type { XmlElement } from '../../xml';
import type { MediaStore } from '../doc/drawings';

/** Largest picture that is kept, bytes. */
const MAX_PICTURE = 64 * 1024 * 1024;
/** The largest width and height Word gives a picture: 22 inches. */
const MAX_PICTURE_TWIPS = 31680;
const TWIP = 635; // EMU
/** Largest length in twips whose EMU value the model holds. */
const MAX_TWIPS = 3000000;
const limited = (twips: number): number => (twips > MAX_TWIPS ? MAX_TWIPS : twips < -MAX_TWIPS ? -MAX_TWIPS : twips);

/** A picture while its group is read. */
export interface PictState {
  /** Media type; undefined for a kind the reader does not keep. */
  type: string | undefined;
  /** Size of the picture itself: pixels for a bitmap, hundredths of a millimetre for a metafile. */
  width: number;
  height: number;
  /** Size it is meant to have, twips. */
  goalWidth: number;
  goalHeight: number;
  /** Percent. */
  scaleX: number;
  scaleY: number;
  /** Twips. */
  cropLeft: number;
  cropTop: number;
  cropRight: number;
  cropBottom: number;
  metafile: boolean;
  chunks: Uint8Array[];
  /** Small pieces on their way into `chunks`, and how much of the array they fill. */
  gathered: Uint8Array | undefined;
  gatheredLength: number;
  size: number;
  /** A hex digit whose partner is in the next piece of text; -1 for none. */
  half: number;
  /** The picture is larger than the reader keeps. */
  tooLarge: boolean;
}

export function newPict(): PictState {
  return {
    type: undefined,
    width: 0,
    height: 0,
    goalWidth: 0,
    goalHeight: 0,
    scaleX: 100,
    scaleY: 100,
    cropLeft: 0,
    cropTop: 0,
    cropRight: 0,
    cropBottom: 0,
    metafile: false,
    chunks: [],
    gathered: undefined,
    gatheredLength: 0,
    size: 0,
    half: -1,
    tooLarge: false,
  };
}

/** A control word of a picture group; false when it is not one. */
export function pictWord(pict: PictState, word: string, param: number): boolean {
  switch (word) {
    case 'pngblip':
      pict.type = 'image/png';
      return true;
    case 'jpegblip':
      pict.type = 'image/jpeg';
      return true;
    case 'emfblip':
      pict.type = 'image/x-emf';
      pict.metafile = true;
      return true;
    case 'wmetafile':
      pict.type = 'image/x-wmf';
      pict.metafile = true;
      return true;
    case 'macpict':
      pict.type = 'image/pict';
      pict.metafile = true;
      return true;
    case 'dibitmap':
      pict.type = 'image/bmp';
      return true;
    case 'wbitmap':
    case 'pmmetafile':
      // A device dependent bitmap and an OS/2 metafile: nothing can show them.
      pict.type = undefined;
      return true;
    case 'picw':
      pict.width = param;
      return true;
    case 'pich':
      pict.height = param;
      return true;
    case 'picwgoal':
      pict.goalWidth = param;
      return true;
    case 'pichgoal':
      pict.goalHeight = param;
      return true;
    case 'picscalex':
      pict.scaleX = param;
      return true;
    case 'picscaley':
      pict.scaleY = param;
      return true;
    case 'piccropl':
      pict.cropLeft = param;
      return true;
    case 'piccropt':
      pict.cropTop = param;
      return true;
    case 'piccropr':
      pict.cropRight = param;
      return true;
    case 'piccropb':
      pict.cropBottom = param;
      return true;
    case 'picscaled':
    case 'picbmp':
    case 'picbpp':
    case 'wbmbitspixel':
    case 'wbmplanes':
    case 'wbmwidthbytes':
    case 'bliptag':
    case 'blipupi':
    case 'defshp':
      return true;
  }
  return false;
}

/** Pieces shorter than this are gathered into one array. */
const SMALL_PIECE = 2048;
const GATHER_SIZE = 16384;

function flushGathered(pict: PictState): void {
  if (pict.gathered && pict.gatheredLength) pict.chunks.push(pict.gathered.subarray(0, pict.gatheredLength));
  pict.gathered = undefined;
  pict.gatheredLength = 0;
}

function addChunk(pict: PictState, chunk: Uint8Array): void {
  if (!chunk.length) return;
  if (pict.size + chunk.length > MAX_PICTURE) {
    pict.tooLarge = true;
    pict.chunks = [];
    pict.gathered = undefined;
    pict.gatheredLength = 0;
    return;
  }
  pict.size += chunk.length;
  if (chunk.length >= SMALL_PIECE) {
    flushGathered(pict);
    pict.chunks.push(chunk);
    return;
  }
  // A picture cut into thousands of small pieces (by groups between its digits) must not cost an array for each.
  if (!pict.gathered || pict.gatheredLength + chunk.length > pict.gathered.length) {
    flushGathered(pict);
    pict.gathered = new Uint8Array(GATHER_SIZE);
  }
  pict.gathered.set(chunk, pict.gatheredLength);
  pict.gatheredLength += chunk.length;
}

/** Adds the bytes the hex digits in bytes[start..stop) stand for; everything that is not a hex digit is passed over. */
export function pictHex(pict: PictState, bytes: Uint8Array, start: number, stop: number): void {
  if (pict.tooLarge) return;
  const out = new Uint8Array(((stop - start) >> 1) + 1);
  let n = 0;
  let half = pict.half;
  for (let i = start; i < stop; i++) {
    const c = bytes[i];
    // The digits, by range; line ends and spaces between them fall through.
    let v: number;
    if (c >= 48 && c <= 57) v = c - 48;
    else if (c >= 97 && c <= 102) v = c - 87;
    else if (c >= 65 && c <= 70) v = c - 55;
    else continue;
    if (half < 0) half = v;
    else {
      out[n++] = (half << 4) | v;
      half = -1;
    }
  }
  pict.half = half;
  addChunk(pict, n === out.length ? out : out.subarray(0, n));
}

/** Adds bytes that are given as they are ("\bin"): a view of the file, not a copy. */
export function pictBin(pict: PictState, bytes: Uint8Array, start: number, stop: number): void {
  if (pict.tooLarge) return;
  pict.half = -1;
  addChunk(pict, bytes.subarray(start, stop));
}

function pictBytes(pict: PictState): Uint8Array {
  flushGathered(pict);
  if (pict.chunks.length === 1) return pict.chunks[0];
  const out = new Uint8Array(pict.size);
  let p = 0;
  for (const chunk of pict.chunks) {
    out.set(chunk, p);
    p += chunk.length;
  }
  return out;
}

const u16le = (b: Uint8Array, o: number): number => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
const u32le = (b: Uint8Array, o: number): number => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
const u32be = (b: Uint8Array, o: number): number => (o + 4 <= b.length ? ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0 : 0);

/** A device independent bitmap as a BMP file: the file header it lacks is put in front. */
function dibToBmp(dib: Uint8Array): Uint8Array | undefined {
  if (dib.length < 16) return undefined;
  const headerSize = u32le(dib, 0);
  if (headerSize < 12 || headerSize > dib.length) return undefined;
  let colors = 0;
  if (headerSize >= 40) {
    const bits = u16le(dib, 14);
    const compression = u32le(dib, 16);
    colors = u32le(dib, 32);
    if (colors === 0 && bits <= 8) colors = 1 << bits;
    if (compression === 3 && headerSize === 40) colors += 3;
    if (colors > 256 + 3) colors = 0;
  } else {
    const bits = u16le(dib, 10);
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

/** Width and height in pixels of a PNG or JPEG file; undefined when its header does not say. */
export function bitmapSize(bytes: Uint8Array, type: string): { width: number; height: number } | undefined {
  if (type === 'image/png') {
    if (bytes.length < 24 || bytes[0] !== 0x89 || bytes[1] !== 0x50) return undefined;
    return { width: u32be(bytes, 16), height: u32be(bytes, 20) };
  }
  if (type !== 'image/jpeg' || bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  // The segments of the file, up to the one that states the size. Each round moves on by at least two bytes.
  let p = 2;
  while (p + 4 <= bytes.length) {
    if (bytes[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = bytes[p + 1];
    if (marker === 0xff) {
      p++;
      continue;
    }
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      p += 2;
      continue;
    }
    const length = (bytes[p + 2] << 8) | bytes[p + 3];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (p + 9 > bytes.length) return undefined;
      return { height: (bytes[p + 5] << 8) | bytes[p + 6], width: (bytes[p + 7] << 8) | bytes[p + 8] };
    }
    if (marker === 0xda || marker === 0xd9) return undefined;
    p += 2 + Math.max(2, length);
  }
  return undefined;
}

/** A stand-in for the element a docx reader keeps a drawing in: an RTF file has none. */
function drawingXml(insets?: (number | undefined)[]): XmlElement {
  const xml: XmlElement = { type: 'element', name: 'rtf:drawing', attrs: new Map(), children: [] };
  if (insets && insets.some((value) => value !== undefined)) {
    // The insets of a text box, where layout looks for them in a DrawingML shape.
    const body: XmlElement = { type: 'element', name: 'wps:bodyPr', attrs: new Map(), children: [] };
    const names = ['lIns', 'tIns', 'rIns', 'bIns'];
    insets.forEach((value, i) => {
      if (value !== undefined && value >= 0 && value <= 100000000) body.attrs.set(names[i], String(Math.round(value)));
    });
    xml.children.push(body);
  }
  return xml;
}

/** A hash of the bytes, to tell pictures apart: the same picture is often in a file many times. */
function hashBytes(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0x1b873593;
  for (let i = 0; i < bytes.length; i++) {
    a = Math.imul(a ^ bytes[i], 0x01000193);
    b = (Math.imul(b, 31) + bytes[i]) | 0;
  }
  return (a >>> 0).toString(36) + ':' + (b >>> 0).toString(36) + ':' + bytes.length;
}

/**
 * The picture as an inline drawing, with its bytes in the media of the model.
 * A picture without bytes the reader keeps is a drawing of its size without
 * an image; undefined when it has neither bytes nor a size.
 */
export function pictItem(pict: PictState, store: MediaStore): DrawingItem | undefined {
  let media: DocMedia | undefined;
  if (pict.type !== undefined && pict.size > 0 && !pict.tooLarge) {
    let bytes: Uint8Array | undefined = pictBytes(pict);
    if (pict.type === 'image/bmp') bytes = dibToBmp(bytes);
    if (bytes) media = { bytes, type: pict.type };
  }
  pict.chunks = [];

  // The size the picture itself has, in twips.
  let naturalWidth = pict.goalWidth > 0 ? pict.goalWidth : 0;
  let naturalHeight = pict.goalHeight > 0 ? pict.goalHeight : 0;
  if (!naturalWidth || !naturalHeight) {
    let width = pict.width;
    let height = pict.height;
    if (pict.metafile) {
      // Hundredths of a millimetre.
      width = (width * 1440) / 2540;
      height = (height * 1440) / 2540;
    } else {
      if ((!(width > 0) || !(height > 0)) && media) {
        const size = bitmapSize(media.bytes, media.type);
        if (size) {
          width = size.width;
          height = size.height;
        }
      }
      // Pixels, taken at 96 to the inch.
      width *= 15;
      height *= 15;
    }
    if (!naturalWidth) naturalWidth = width > 0 ? width : 0;
    if (!naturalHeight) naturalHeight = height > 0 ? height : 0;
  }
  if (!media && !(naturalWidth > 0 && naturalHeight > 0)) return undefined;

  const scaleX = pict.scaleX > 0 ? pict.scaleX : 100;
  const scaleY = pict.scaleY > 0 ? pict.scaleY : 100;
  let width = Math.max(0, ((naturalWidth - pict.cropLeft - pict.cropRight) * scaleX) / 100);
  let height = Math.max(0, ((naturalHeight - pict.cropTop - pict.cropBottom) * scaleY) / 100);
  // No picture is wider or taller than 22 inches in Word. A scale on one axis that gives more is taken to be
  // damaged: the picture keeps its shape at the size the other axis gives.
  if (width > MAX_PICTURE_TWIPS && height <= MAX_PICTURE_TWIPS && naturalHeight > 0) width = (height * naturalWidth) / naturalHeight;
  else if (height > MAX_PICTURE_TWIPS && width <= MAX_PICTURE_TWIPS && naturalWidth > 0) height = (width * naturalHeight) / naturalWidth;
  width = Math.min(width, MAX_PICTURE_TWIPS);
  height = Math.min(height, MAX_PICTURE_TWIPS);

  const item: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', xml: drawingXml() };
  item.width = Math.round(width * TWIP);
  item.height = Math.round(height * TWIP);
  if ((pict.cropLeft || pict.cropTop || pict.cropRight || pict.cropBottom) && naturalWidth > 0 && naturalHeight > 0) {
    const crop: ImageCrop = {};
    const share = (value: number, whole: number) => Math.max(-1000000, Math.min(1000000, Math.round((value / whole) * 100000)));
    if (pict.cropLeft) crop.left = share(pict.cropLeft, naturalWidth);
    if (pict.cropRight) crop.right = share(pict.cropRight, naturalWidth);
    if (pict.cropTop) crop.top = share(pict.cropTop, naturalHeight);
    if (pict.cropBottom) crop.bottom = share(pict.cropBottom, naturalHeight);
    item.crop = crop;
  }
  if (media) {
    const found = media;
    const image = store.get('rtf:' + found.type + ':' + hashBytes(found.bytes), () => found);
    if (image) {
      item.image = image;
      item.images = item.crop ? [{ image, crop: item.crop }] : [{ image }];
    }
  }
  return item;
}

// --------------------------------------------------------------------- shapes

/** A shape while its group is read. */
export interface ShapeState {
  left: number;
  top: number;
  right: number;
  bottom: number;
  /** Width and height of a drawing object of Word 95, which states them instead of its right and bottom edge. */
  sizeX: number | undefined;
  sizeY: number | undefined;
  /** What the position is measured from, as the control words say; undefined when the properties of the shape say it. */
  relativeH: string | undefined;
  relativeV: string | undefined;
  /** "\shpwr": 1 above and below, 2 around, 3 none, 4 tight, 5 through. */
  wrap: number;
  wrapSide: number;
  behind: boolean;
  group: boolean;
  /** Properties by name ("{\sp{\sn name}{\sv value}}"). */
  props: Map<string, string>;
  /** Name of the property whose value is being read. */
  name: string;
  images: ImageRef[];
  crops: (ImageCrop | undefined)[];
  boxes: Block[][];
}

export function newShape(group: boolean): ShapeState {
  return {
    left: 0,
    top: 0,
    right: 0,
    bottom: 0,
    sizeX: undefined,
    sizeY: undefined,
    relativeH: 'column',
    relativeV: 'paragraph',
    wrap: 3,
    wrapSide: 0,
    behind: false,
    group,
    props: new Map(),
    name: '',
    images: [],
    crops: [],
    boxes: [],
  };
}

/** A control word of a shape; false when it is not one. */
export function shapeWord(shape: ShapeState, word: string, param: number): boolean {
  switch (word) {
    case 'shpleft':
    case 'dpx':
      shape.left = limited(param);
      return true;
    case 'shptop':
    case 'dpy':
      shape.top = limited(param);
      return true;
    case 'shpright':
      shape.right = limited(param);
      return true;
    case 'shpbottom':
      shape.bottom = limited(param);
      return true;
    case 'dpxsize':
      shape.sizeX = Math.max(0, limited(param));
      return true;
    case 'dpysize':
      shape.sizeY = Math.max(0, limited(param));
      return true;
    case 'shpbxpage':
    case 'dobxpage':
      shape.relativeH = 'page';
      return true;
    case 'shpbxmargin':
    case 'dobxmargin':
      shape.relativeH = 'margin';
      return true;
    case 'shpbxcolumn':
    case 'dobxcolumn':
      shape.relativeH = 'column';
      return true;
    case 'shpbypage':
    case 'dobypage':
      shape.relativeV = 'page';
      return true;
    case 'shpbymargin':
    case 'dobymargin':
      shape.relativeV = 'margin';
      return true;
    case 'shpbypara':
    case 'dobypara':
      shape.relativeV = 'paragraph';
      return true;
    case 'shpwr':
      shape.wrap = param;
      return true;
    case 'shpwrk':
      shape.wrapSide = param;
      return true;
    case 'shpfblwtxt':
      shape.behind = param === 1;
      return true;
    case 'shpbxignore':
    case 'shpbyignore':
    case 'shpfhdr':
    case 'shpz':
    case 'shplid':
    case 'shplockanchor':
    case 'dodhgt':
    case 'dolock':
      return true;
  }
  return false;
}

const POSITION_H = [undefined, 'left', 'center', 'right', 'inside', 'outside'];
const POSITION_V = [undefined, 'top', 'center', 'bottom', 'inside', 'outside'];
const RELATIVE_H = ['margin', 'page', 'column', 'character'];
const RELATIVE_V = ['margin', 'page', 'paragraph', 'line'];
const WRAPS: (NonNullable<DrawingAnchor['wrap']> | undefined)[] = [undefined, 'topAndBottom', 'square', 'none', 'tight', 'through'];
const WRAP_SIDES = ['bothSides', 'left', 'right', 'largest'];

/** The whole number a property value states; undefined when it is not one. */
function numberOf(text: string | undefined): number | undefined {
  if (text === undefined) return undefined;
  let p = 0;
  while (p < text.length && text.charCodeAt(p) === 0x20) p++;
  let negative = false;
  if (text.charCodeAt(p) === 0x2d) {
    negative = true;
    p++;
  }
  let value = 0;
  let digits = 0;
  for (; p < text.length && digits < 12; p++, digits++) {
    const c = text.charCodeAt(p);
    if (c < 48 || c > 57) break;
    value = value * 10 + (c - 48);
  }
  if (!digits) return undefined;
  return negative ? -value : value;
}

/**
 * The anchored drawing a shape stands for. `legacy` is a drawing object of
 * Word 95, whose lines and fills are control words the reader does not read:
 * it gets no look.
 */
export function shapeItem(shape: ShapeState, legacy = false): DrawingItem {
  const props = shape.props;
  const horizontal: DrawingPosition = { relativeFrom: shape.relativeH || 'column', offset: shape.left * TWIP };
  const vertical: DrawingPosition = { relativeFrom: shape.relativeV || 'paragraph', offset: shape.top * TWIP };
  const anchor: DrawingAnchor = { horizontal, vertical, wrap: WRAPS[shape.wrap] || 'none' };
  const side = WRAP_SIDES[shape.wrapSide];
  if (side !== undefined && anchor.wrap !== 'none' && anchor.wrap !== 'topAndBottom') anchor.wrapText = side;
  if (shape.behind) anchor.behindDoc = true;

  // Word 2000 and later place a shape by these properties; the control words are what Word 97 would show.
  const relH = numberOf(props.get('posrelh'));
  if (relH !== undefined && RELATIVE_H[relH] !== undefined) horizontal.relativeFrom = RELATIVE_H[relH];
  const relV = numberOf(props.get('posrelv'));
  if (relV !== undefined && RELATIVE_V[relV] !== undefined) vertical.relativeFrom = RELATIVE_V[relV];
  const posH = numberOf(props.get('posh'));
  if (posH !== undefined && POSITION_H[posH] !== undefined) {
    horizontal.align = POSITION_H[posH];
    delete horizontal.offset;
  }
  const posV = numberOf(props.get('posv'));
  if (posV !== undefined && POSITION_V[posV] !== undefined) {
    vertical.align = POSITION_V[posV];
    delete vertical.offset;
  }
  const distances = [
    ['dxWrapDistLeft', 'distL'],
    ['dyWrapDistTop', 'distT'],
    ['dxWrapDistRight', 'distR'],
    ['dyWrapDistBottom', 'distB'],
  ] as const;
  for (const [name, field] of distances) {
    const value = numberOf(props.get(name));
    if (value !== undefined && value >= 0 && value <= 100000000) anchor[field] = value;
  }
  const behind = numberOf(props.get('fBehindDocument'));
  if (behind === 1) anchor.behindDoc = true;
  else if (behind === 0) delete anchor.behindDoc;

  const insets = ['dxTextLeft', 'dyTextTop', 'dxTextRight', 'dyTextBottom'].map((name) => numberOf(props.get(name)));
  const item: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'shape', xml: drawingXml() };
  const rotation = numberOf(props.get('rotation'));
  const flipH = numberOf(props.get('fFlipH')) === 1;
  const flipV = numberOf(props.get('fFlipV')) === 1;
  if (rotation || flipH || flipV) {
    item.transform = {};
    // Degrees as a 16.16 fixed point number -> 60000ths of a degree.
    if (rotation) item.transform.rotation = Math.round((rotation / 65536) * 60000);
    if (flipH) item.transform.flipH = true;
    if (flipV) item.transform.flipV = true;
  }
  const name = props.get('wzName');
  if (name) item.name = name;
  const description = props.get('wzDescription');
  if (description) item.description = description;

  if (shape.group) item.kind = 'group';
  else if (shape.images.length && !shape.boxes.length) item.kind = 'picture';
  else if (!shape.images.length && (numberOf(props.get('shapeType')) === 75 || props.has('pib'))) {
    // A shape that is to show a picture the file does not have, or the reader does not keep: it is not a box
    // filled in white. Without text it stays a picture, without an image.
    if (!shape.boxes.length) item.kind = 'picture';
  } else if (!legacy) {
    const flag = (name: string): boolean | undefined => {
      const value = numberOf(props.get(name));
      return value === undefined ? undefined : value !== 0;
    };
    item.shape = officeArtLook({
      type: numberOf(props.get('shapeType')),
      fFilled: flag('fFilled'),
      fillType: numberOf(props.get('fillType')),
      fillColor: numberOf(props.get('fillColor')),
      fillOpacity: numberOf(props.get('fillOpacity')),
      fLine: flag('fLine'),
      lineColor: numberOf(props.get('lineColor')),
      lineWidth: numberOf(props.get('lineWidth')),
      lineDashing: numberOf(props.get('lineDashing')),
      adjust: numberOf(props.get('adjustValue')),
    });
  }
  item.width = (shape.sizeX !== undefined ? shape.sizeX : Math.max(0, shape.right - shape.left)) * TWIP;
  item.height = (shape.sizeY !== undefined ? shape.sizeY : Math.max(0, shape.bottom - shape.top)) * TWIP;
  item.anchor = anchor;
  if (shape.images.length) {
    item.image = shape.images[0];
    const crop = shape.crops[0];
    if (crop && item.kind === 'picture') item.crop = crop;
    item.images = shape.images.map((image, i) => (i === 0 && item.crop ? { image, crop: item.crop } : { image }));
  }
  if (shape.boxes.length) {
    item.textBoxes = shape.boxes;
    item.xml = drawingXml(insets);
  }
  return item;
}
