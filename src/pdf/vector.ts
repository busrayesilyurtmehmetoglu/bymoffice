// What a view needs to write a page of a PDF as text and drawings instead of
// as a picture: the font files of the page (the fonts the application
// registered and the font of the package - a font of the device has no file
// here), a measurer that asks the real canvas, a reader of the pictures the
// painters draw, and drawPage(), which lets a painter paint on a PDF canvas
// and says whether the page could be written that way.
//
// Nothing touches the DOM when the module is imported.

import { deviceHasFamily, embeddedFontFiles } from '../render/embedded-fonts';
import { registeredFontFiles, type FontFileSource } from '../render/fonts';
import { cleanReason, createPdfCanvas, type PdfCanvasFonts, type PdfInk, type PdfCanvasImages, type PdfCanvasMeasure, type PdfPicture, type PdfVectorPage } from './canvas';
import { PdfFontSet, type PdfFontSource } from './font-set';
import { PHOTO_SHARE, changeShare, jpegInfo, type PdfImage } from './image';
import { MAX_FONT_BYTES } from './sfnt';

/** What drawPage() draws with. */
export interface PdfVectorTools {
  fonts: PdfCanvasFonts;
  measure: PdfCanvasMeasure;
  images?: PdfCanvasImages;
}

/** How long the file of a font may take to arrive before the font counts as one without a file. */
const FONT_FETCH_MS = 15000;
const MAX_FETCHED = 32;
/** Font files that were fetched, by their address: the same bytes again for the next file, so that a font is read once. */
const fetched = new Map<string, Promise<Uint8Array | undefined>>();

type Fetch = (url: string, init?: { signal?: unknown }) => Promise<{ ok: boolean; arrayBuffer(): Promise<ArrayBuffer> }>;

function fetchFont(url: string, fetcher: Fetch | undefined): Promise<Uint8Array | undefined> {
  const known = fetched.get(url);
  if (known) return known;
  const scope = (typeof globalThis !== 'undefined' ? globalThis : {}) as { fetch?: Fetch; AbortController?: new () => { signal: unknown; abort(): void } };
  const get = fetcher || (typeof scope.fetch === 'function' ? scope.fetch.bind(scope) : undefined);
  if (!get) return Promise.resolve(undefined);
  const control = typeof scope.AbortController === 'function' ? new scope.AbortController() : undefined;
  const timer = control ? setTimeout(() => control.abort(), FONT_FETCH_MS) : undefined;
  const job = Promise.resolve()
    .then(() => get(url, control ? { signal: control.signal } : undefined))
    .then((response) => (response && response.ok ? response.arrayBuffer() : undefined))
    .then((buffer) => (buffer && buffer.byteLength > 0 && buffer.byteLength <= MAX_FONT_BYTES ? new Uint8Array(buffer) : undefined))
    .catch(() => undefined)
    .then((bytes) => {
      if (timer !== undefined) clearTimeout(timer);
      // A font that did not arrive is asked for again the next time.
      if (!bytes) fetched.delete(url);
      return bytes;
    });
  if (fetched.size >= MAX_FETCHED) fetched.clear();
  fetched.set(url, job);
  return job;
}

/**
 * The fonts a PDF can be written with now: every face the application
 * registered and the faces of the font of the package that are on the page,
 * each with its font file - the bytes it was given as, or what its address
 * gives (asked for here; a file that does not arrive leaves its face out, and
 * text in it stays a picture). `sources` and `fetcher` stand in for the fonts
 * of the page and for fetch().
 */
export async function loadPdfFonts(sources?: readonly FontFileSource[], fetcher?: Fetch, deviceHas: (family: string) => boolean = deviceHasFamily): Promise<PdfFontSet> {
  const all = sources || registeredFontFiles().concat(embeddedFontFiles());
  const read = await Promise.all(
    all.map((source): Promise<PdfFontSource | undefined> => {
      const made = (bytes: Uint8Array | undefined): PdfFontSource | undefined =>
        // A face without a file still names its family: text in it is not passed on to the next family of a list.
        ({ family: source.family, weight: source.weight, style: source.style, bytes });
      if (source.bytes instanceof Uint8Array) return Promise.resolve(made(source.bytes));
      if (typeof source.url === 'string' && source.url !== '') return fetchFont(source.url, fetcher).then(made);
      return Promise.resolve(made(undefined));
    }),
  );
  const usable: PdfFontSource[] = [];
  for (let i = 0; i < read.length; i++) if (read[i]) usable.push(read[i] as PdfFontSource);
  return new PdfFontSet(usable, deviceHas);
}

interface MeasureContext {
  font: string;
  textBaseline: string;
  fontKerning?: string;
  measureText(text: string): { width: number; actualBoundingBoxAscent?: number };
}

interface PaintContext extends MeasureContext {
  fillStyle: unknown;
  clearRect(x: number, y: number, width: number, height: number): void;
  fillText(text: string, x: number, y: number): void;
  getImageData(x: number, y: number, width: number, height: number): { data: ArrayLike<number> };
}

/** The size, in px, a text is painted at to see how it is painted. */
const INK_SIZE = 64;
/** Characters of a text that is painted for that, and the widest canvas it is painted on. */
const MAX_INK_TEXT = 32;
const MAX_INK_WIDTH = 4096;

/**
 * A measurer that asks a canvas of the document: the widths the painters
 * measured their text with, and where the baselines of a font lie. Undefined
 * where there is no canvas.
 */
export function canvasMeasure(doc: Document): PdfCanvasMeasure | undefined {
  let context: MeasureContext | null = null;
  try {
    context = doc.createElement('canvas').getContext('2d') as unknown as MeasureContext | null;
  } catch {
    context = null;
  }
  if (!context) return undefined;
  const ctx = context;
  let current = '';
  const use = (font: string): void => {
    if (current !== font) {
      ctx.font = font;
      current = font;
    }
  };
  const widths = new Map<string, number>();
  const lines = new Map<string, number>();
  const inks = new Map<string, PdfInk | null>();
  let scratch: { canvas: { width: number; height: number }; context: PaintContext } | null | undefined;
  /** The width of a text without kerning, and the pixels the canvas paints for it at INK_SIZE px, as a digest. */
  const inkOf = (font: string, text: string): PdfInk | null => {
    // Only a canvas that takes the setting has turned kerning off.
    if (!('fontKerning' in ctx) || text.length > MAX_INK_TEXT) return null;
    use(font);
    ctx.textBaseline = 'alphabetic';
    ctx.fontKerning = 'none';
    if (ctx.fontKerning !== 'none') return null;
    const width = ctx.measureText(text).width;
    ctx.fontKerning = 'auto';
    // The same font, larger: what differs by a hair at the size of the text differs by pixels there.
    const large = font.replace(/(\d+(?:\.\d+)?|\.\d+)px/, INK_SIZE + 'px');
    if (large === font && font.indexOf(INK_SIZE + 'px') < 0) return null;
    if (scratch === undefined) {
      const canvas = doc.createElement('canvas');
      const context = canvas.getContext('2d') as unknown as PaintContext | null;
      scratch = context && 'fontKerning' in context ? { canvas, context } : null;
    }
    if (!scratch) return null;
    const paint = scratch.context;
    // As wide for a text with joiners as for the text without them: the pictures of the two are compared.
    const pixelWidth = Math.min(MAX_INK_WIDTH, text.replace(/\u200c/g, '').length * INK_SIZE * 2 + INK_SIZE * 2);
    scratch.canvas.width = pixelWidth;
    scratch.canvas.height = INK_SIZE * 3;
    try {
      paint.font = large;
      paint.textBaseline = 'alphabetic';
      paint.fontKerning = 'none';
      if (paint.fontKerning !== 'none') return null;
      paint.clearRect(0, 0, pixelWidth, INK_SIZE * 3);
      paint.fillStyle = '#000000';
      paint.fillText(text, INK_SIZE, INK_SIZE * 2);
      const data = paint.getImageData(0, 0, pixelWidth, INK_SIZE * 3).data;
      // Two numbers over how opaque every pixel is: the same pixels, the same two numbers.
      let first = 0x811c9dc5;
      let second = 0;
      for (let at = 3; at < data.length; at += 4) {
        first = Math.imul(first ^ data[at], 0x01000193) >>> 0;
        second = (Math.imul(second, 31) + data[at] + (at >> 2)) >>> 0;
      }
      return { width, ink: first.toString(16) + '-' + second.toString(16) };
    } finally {
      paint.fontKerning = 'auto';
    }
  };
  return {
    plain: (font, text) => {
      const key = font + '\n' + text;
      const known = inks.get(key);
      if (known !== undefined) return known || undefined;
      let ink: PdfInk | null = null;
      try {
        ink = inkOf(font, text);
      } catch {
        // Pixels that cannot be read: the canvas cannot be asked.
        ink = null;
      } finally {
        if ('fontKerning' in ctx) ctx.fontKerning = 'auto';
      }
      if (inks.size >= 5000) inks.clear();
      inks.set(key, ink);
      return ink || undefined;
    },
    width: (font, text) => {
      const key = text.length <= 64 ? font + '\n' + text : '';
      const known = key === '' ? undefined : widths.get(key);
      if (known !== undefined) return known;
      use(font);
      ctx.textBaseline = 'alphabetic';
      const width = ctx.measureText(text).width;
      if (key !== '') {
        if (widths.size >= 20000) widths.clear();
        widths.set(key, width);
      }
      return width;
    },
    baseline: (font, baseline) => {
      const key = font + '\n' + baseline;
      let down = lines.get(key);
      if (down === undefined) {
        use(font);
        // How far the ink of a letter rises above each of the two lines: the difference is how far the lines are apart.
        ctx.textBaseline = 'alphabetic';
        const above = ctx.measureText('x').actualBoundingBoxAscent;
        ctx.textBaseline = baseline;
        const from = ctx.measureText('x').actualBoundingBoxAscent;
        ctx.textBaseline = 'alphabetic';
        down = typeof above === 'number' && typeof from === 'number' ? above - from : NaN;
        if (lines.size >= 2000) lines.clear();
        lines.set(key, down);
      }
      return down;
    },
  };
}

/** Pixels of one picture that are read; a larger one leaves its page a picture. */
export const MAX_PICTURE_PIXELS = 16000000;

/** The picture RGBA pixels are: RGB or grey samples, with a mask when a pixel lets through what is behind it. */
export function pictureOfPixels(rgba: ArrayLike<number>, width: number, height: number): PdfImage & { kind: 'samples' } {
  const count = width * height;
  let grey = true;
  let opaque = true;
  for (let i = 0, at = 0; i < count && (grey || opaque); i++, at += 4) {
    if (rgba[at] !== rgba[at + 1] || rgba[at + 1] !== rgba[at + 2]) grey = false;
    if (rgba[at + 3] !== 255) opaque = false;
  }
  const colors = grey ? 1 : 3;
  const data = new Uint8Array(count * colors);
  const alpha = opaque ? undefined : new Uint8Array(count);
  for (let i = 0, at = 0, to = 0; i < count; i++, at += 4) {
    data[to++] = rgba[at];
    if (!grey) {
      data[to++] = rgba[at + 1];
      data[to++] = rgba[at + 2];
    }
    if (alpha) alpha[i] = rgba[at + 3];
  }
  return { kind: 'samples', data, width, height, colors, alpha };
}

interface ScratchCanvas {
  width: number;
  height: number;
  getContext(kind: '2d'): {
    clearRect(x: number, y: number, width: number, height: number): void;
    drawImage(source: unknown, x: number, y: number): void;
    getImageData(x: number, y: number, width: number, height: number): { data: ArrayLike<number> };
  } | null;
  toDataURL?(type: string, quality?: number): string;
}

const JPEG_HEAD = 'data:image/jpeg;base64,';

/**
 * Reads the pictures painters draw (decoded pictures, other canvases) through
 * a canvas of the document: their pixels, with how opaque each is. A picture
 * that is a photograph and hides everything behind it is stored as a JPEG of
 * `quality` when `kind` allows it ('auto': by the share of its pixels that
 * differ from their neighbour, as for the picture of a page), otherwise as
 * its pixels. A picture is read once.
 */
export function canvasPictures(doc: Document, kind: 'auto' | 'jpeg' | 'lossless' = 'auto', quality = 0.85): PdfCanvasImages {
  const known = new Map<unknown, PdfPicture | null>();
  /** The sources read since the last page ended: what only the page that is being painted needs so far. */
  let fresh: unknown[] = [];
  let scratch: ScratchCanvas | undefined;
  const read = (source: unknown): PdfPicture | undefined => {
    if (source === null || typeof source !== 'object') return undefined;
    const sized = source as { naturalWidth?: unknown; naturalHeight?: unknown; width?: unknown; height?: unknown };
    const width = typeof sized.naturalWidth === 'number' ? sized.naturalWidth : sized.width;
    const height = typeof sized.naturalHeight === 'number' ? sized.naturalHeight : sized.height;
    if (typeof width !== 'number' || typeof height !== 'number' || !(width >= 1) || !(height >= 1) || Math.floor(width) !== width || Math.floor(height) !== height || width * height > MAX_PICTURE_PIXELS) return undefined;
    if (!scratch) scratch = doc.createElement('canvas') as unknown as ScratchCanvas;
    const canvas = scratch;
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext('2d');
    if (!context) return undefined;
    context.clearRect(0, 0, width, height);
    context.drawImage(source, 0, 0);
    const samples = pictureOfPixels(context.getImageData(0, 0, width, height).data, width, height);
    let image: PdfImage = samples;
    if (kind !== 'lossless' && !samples.alpha && (kind === 'jpeg' || changeShare(samples) > PHOTO_SHARE) && typeof canvas.toDataURL === 'function' && typeof atob === 'function') {
      const url = canvas.toDataURL('image/jpeg', quality);
      if (url.slice(0, JPEG_HEAD.length) === JPEG_HEAD) {
        const binary = atob(url.slice(JPEG_HEAD.length));
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const info = jpegInfo(bytes);
        if (info.width === width && info.height === height) image = { kind: 'jpeg', data: bytes, width, height, components: info.components };
      }
    }
    return { width, height, image };
  };
  return {
    read: (source) => {
      const before = known.get(source);
      if (before !== undefined) return before || undefined;
      let picture: PdfPicture | undefined;
      try {
        picture = read(source);
      } catch {
        // A canvas that may not be read (pixels of another origin), or one that is too large.
        picture = undefined;
      }
      known.set(source, picture || null);
      if (picture) fresh.push(source);
      // The scratch canvas holds nothing between two pictures.
      if (scratch) {
        scratch.width = 1;
        scratch.height = 1;
      }
      return picture;
    },
    pageEnded: (kept) => {
      // The pixels read for a page that is not written, and that no page before it drew, are let go of.
      if (!kept) for (let i = 0; i < fresh.length; i++) known.delete(fresh[i]);
      fresh = [];
    },
  };
}

/** How a page went into the file. */
export interface PdfPageInfo {
  /** The number of the page, from 1. */
  page: number;
  /** 'vector': text in embedded fonts, drawings and pictures of their own; 'image': one picture of the page with its text lying invisible over it. */
  mode: 'vector' | 'image';
  /** For 'image' where 'vector' was tried: why the page could not be written as text and drawings. */
  reason?: string;
}

/**
 * Lets a painter paint a page on a PDF canvas of the size its real canvas
 * has. Gives the page when all of it could be written, the reason when not -
 * also when the painter threw: a page is written whole or not at all.
 */
export function drawPage(tools: PdfVectorTools, size: { pixelWidth: number; pixelHeight: number; pageWidth: number; pageHeight: number }, paint: (context: CanvasRenderingContext2D) => void): { vector?: PdfVectorPage; reason?: string } {
  const made = createPdfCanvas({ width: size.pixelWidth, height: size.pixelHeight, pageWidth: size.pageWidth, pageHeight: size.pageHeight, fonts: tools.fonts, measure: tools.measure, images: tools.images });
  const ended = (kept: boolean): void => {
    if (tools.images && tools.images.pageEnded) tools.images.pageEnded(kept);
  };
  try {
    paint(made.context);
  } catch (error) {
    if (made.canvas.unsupported === undefined) {
      ended(false);
      return { reason: cleanReason(`the page could not be painted: ${error instanceof Error ? error.message : String(error)}`) };
    }
  }
  if (made.canvas.unsupported !== undefined) {
    ended(false);
    return { reason: made.canvas.unsupported };
  }
  ended(true);
  return { vector: made.canvas.finish() };
}
