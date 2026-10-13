// Making a PDF of what a view shows: the options and their limits, the
// picture of one page as a canvas gives it, the name and the download of the
// file. The views paint the pages with the code they print with and hand
// each one to a PdfDocument; what both need of each other is here. Nothing
// touches the DOM when the module is imported.
//
// A page goes into the file in one of two ways (PdfExportOptions.text): as
// text in embedded fonts, drawings and pictures of their own ('vector': sharp
// at every size, small), or as one picture of the page with its text lying
// invisible over it ('image'). Which one a page can be is decided page by
// page: see writePdfPage().

import { DocumentPrintError, printLimits, type PrintLimits } from '../ui/doc-print';
import type { PdfDocument, PdfPage } from './document';
import { PHOTO_SHARE, TEXT_SHARE, changeShare, imageOfPixels, jpegInfo, packSamples, type PdfImage } from './image';
import { canvasMeasure, canvasPictures, drawPage, loadPdfFonts, type PdfPageInfo, type PdfVectorTools } from './vector';
import { cleanReason } from './canvas';
import { PdfError } from './writer';

export type { PdfPageInfo };

/** Dots per inch the pages are painted at when nothing else is asked for. */
export const DEFAULT_PDF_DPI = 150;
export const MIN_PDF_DPI = 72;
export const MAX_PDF_DPI = 300;
/** Quality of the JPEG pictures, between 0 and 1. */
export const DEFAULT_PDF_QUALITY = 0.85;
/** Bytes of a whole file; a phone holds less. */
export const MAX_PDF_FILE_BYTES = 768 * 1024 * 1024;
export const MAX_PDF_FILE_BYTES_COARSE = 192 * 1024 * 1024;
/** Rows of pixels that are read at once to see whether a page is of one colour. */
const STRIP_ROWS = 256;

/** What says that the caller does not want the file any more: an AbortSignal, or anything with its `aborted`. */
export interface PdfAbortSignal {
  readonly aborted: boolean;
}

export interface PdfExportOptions {
  /**
   * Resolution the pages are painted at, in dots per inch. Default 150,
   * between 72 and 300. A page larger than a canvas can be is painted at less.
   */
  dpi?: number;
  /**
   * How the picture of a page is stored: 'lossless' (every pixel as it was
   * painted; small for text, large for photographs), 'jpeg' (small for
   * photographs; text gets larger, and blurred at its edges) or 'auto' (the
   * default): page by page whichever suits the page - see canvasPageImage().
   */
  image?: 'auto' | 'jpeg' | 'lossless';
  /** Quality of the JPEG pictures, above 0 and at most 1. Default 0.85. */
  quality?: number;
  /**
   * How a page is written. 'auto' (the default): as text in embedded fonts,
   * drawings and pictures of their own where every bit of the page can be
   * written that way, and as a picture of the page (with its text lying
   * invisible over it) where not - page by page. 'image': every page as a
   * picture. 'vector': as 'auto', but a page that cannot be written as text
   * and drawings ends the work with a PdfExportError 'unsupported' that
   * names the page and the reason.
   *
   * A page can be written as text when the font files of all its text are
   * at hand - the font the package brings along and the fonts registered
   * with fonts.registerFont() (TrueType or OpenType, with a licence that
   * allows embedding) - and its fonts have a glyph for every character. Text
   * in a font of the device leaves its page a picture.
   */
  text?: 'auto' | 'image' | 'vector';
  /** Called for every page once it is in the file, with how it was written. */
  onPage?(info: PdfPageInfo): void;
  /** Title and author in the information of the file; left out when missing. */
  title?: string;
  author?: string;
  /** The date the file states as that of its making; now when missing. */
  date?: Date;
  /** Aborting it ends the work: the promise is rejected with a PdfExportError 'cancelled'. */
  signal?: PdfAbortSignal;
  /** Called when a page is done, with the number of pages done and of all pages. */
  onProgress?(page: number, total: number): void;
}

/** What downloadPdf() of a view says of the file it handed to the browser. */
export interface PdfDownload {
  fileName: string;
  /** Bytes of the file. */
  size: number;
}

export type PdfExportErrorCode = 'unavailable' | 'empty' | 'too-long' | 'cancelled' | 'unsupported' | 'failed';

/**
 * Why no PDF was made: 'unavailable' where there is no canvas to paint the
 * pages on, 'empty' when there is no page, 'too-long' beyond the limits on
 * pages and bytes, 'cancelled' when the signal was aborted, the view was
 * destroyed or what is being typed could not be applied, 'unsupported' when
 * `text: 'vector'` was asked for and a page cannot be written as text and
 * drawings (`page` is its number from 1, `reason` says why), 'failed' for
 * everything else.
 */
export class PdfExportError extends Error {
  readonly code: PdfExportErrorCode;
  /** For 'unsupported': the number of the page, from 1. */
  readonly page?: number;
  /** For 'unsupported': why the page cannot be written as text and drawings. */
  readonly reason?: string;
  constructor(code: PdfExportErrorCode, message: string, page?: number, reason?: string) {
    super(message);
    this.name = 'PdfExportError';
    this.code = code;
    if (page !== undefined) this.page = page;
    if (reason !== undefined) this.reason = reason;
  }
}

export interface PdfSettings {
  dpi: number;
  image: 'auto' | 'jpeg' | 'lossless';
  quality: number;
  text: 'auto' | 'image' | 'vector';
}

/**
 * The options as they are used. Throws a RangeError for a `dpi` or a
 * `quality` that is no number in its range and for an unknown `image` or
 * `text`, a TypeError for an `onProgress` or an `onPage` that is no function
 * or a `signal` that is no object.
 */
export function pdfSettings(options: PdfExportOptions): PdfSettings {
  if (options === null || typeof options !== 'object') throw new TypeError('The options of a PDF must be an object');
  let dpi = DEFAULT_PDF_DPI;
  if (options.dpi !== undefined) {
    if (typeof options.dpi !== 'number' || !(options.dpi > 0) || options.dpi === Infinity) throw new RangeError('The resolution of a PDF must be a positive number of dots per inch');
    dpi = Math.min(MAX_PDF_DPI, Math.max(MIN_PDF_DPI, options.dpi));
  }
  let quality = DEFAULT_PDF_QUALITY;
  if (options.quality !== undefined) {
    if (typeof options.quality !== 'number' || !(options.quality > 0) || !(options.quality <= 1)) throw new RangeError('The quality of the pictures of a PDF must be a number above 0 and at most 1');
    quality = options.quality;
  }
  if (options.image !== undefined && options.image !== 'auto' && options.image !== 'jpeg' && options.image !== 'lossless') throw new RangeError('"image" must be "auto", "jpeg" or "lossless"');
  if (options.text !== undefined && options.text !== 'auto' && options.text !== 'image' && options.text !== 'vector') throw new RangeError('"text" must be "auto", "image" or "vector"');
  if (options.onProgress !== undefined && typeof options.onProgress !== 'function') throw new TypeError('"onProgress" must be a function');
  if (options.onPage !== undefined && typeof options.onPage !== 'function') throw new TypeError('"onPage" must be a function');
  if (options.signal !== undefined && (options.signal === null || typeof options.signal !== 'object')) throw new TypeError('"signal" must be an AbortSignal');
  return { dpi, image: options.image || 'auto', quality, text: options.text || 'auto' };
}

/**
 * What pages are drawn with when they are written as text and drawings: the
 * font files of the page (asked for here where only their address is known),
 * a measurer and a reader of pictures on canvases of the document. Undefined
 * for `text: 'image'`, and where there is no canvas to measure with.
 */
export async function pdfVectorTools(doc: Document, settings: PdfSettings): Promise<PdfVectorTools | undefined> {
  if (settings.text === 'image') return undefined;
  const measure = canvasMeasure(doc);
  if (!measure) return undefined;
  const fonts = await loadPdfFonts();
  return { fonts, measure, images: canvasPictures(doc, settings.image, settings.quality) };
}

/**
 * Makes a file with `make`; when a font could not be put into it as the file
 * was ended (PdfError 'font': not expected, since every glyph is looked at
 * before it is written on a page), the file is made once more with every
 * page as a picture - or, where only text and drawings were asked for, the
 * work ends with a PdfExportError 'unsupported'.
 */
export async function pdfWithFallback<T>(settings: PdfSettings, make: (settings: PdfSettings) => Promise<T>): Promise<T> {
  try {
    return await make(settings);
  } catch (error) {
    if (!(error instanceof PdfError) || error.code !== 'font' || settings.text === 'image') throw error;
    if (settings.text === 'vector') throw new PdfExportError('unsupported', `The file cannot be written as text and drawings: ${cleanReason(error.message)}`, undefined, cleanReason(error.message));
    return make({ ...settings, text: 'image' });
  }
}

export interface PdfPageJob {
  /** The number of the page, from 1. */
  number: number;
  /** Size of the page in points. */
  width: number;
  height: number;
  /** Size of the canvas the page is painted on, in its pixels. */
  pixelWidth: number;
  pixelHeight: number;
  /** Paints the page on a canvas of that size. */
  paint(context: CanvasRenderingContext2D): void;
  /** The page as a picture, with its text: asked for only when the page is not written as text and drawings. */
  picture(): Promise<Pick<PdfPage, 'image' | 'fill' | 'text'>>;
  /** Called before the page is written into the file; false ends the work without a page (what it throws ends it too). */
  proceed?(): boolean;
}

/**
 * Writes one page into a file, as `settings.text` says: as text and drawings
 * when `tools` can draw all of it, as its picture otherwise - or not at all,
 * with a PdfExportError 'unsupported', when only text and drawings were asked
 * for. Tells `onPage` how the page was written. False when `proceed` said so.
 */
export async function writePdfPage(pdf: PdfDocument, job: PdfPageJob, settings: PdfSettings, tools: PdfVectorTools | undefined, onPage?: (info: PdfPageInfo) => void): Promise<boolean> {
  const info: PdfPageInfo = { page: job.number, mode: 'image' };
  let written = false;
  if (settings.text !== 'image') {
    const drawn = tools ? drawPage(tools, { pixelWidth: job.pixelWidth, pixelHeight: job.pixelHeight, pageWidth: job.width, pageHeight: job.height }, job.paint) : { reason: 'there is no canvas to measure text with' };
    if (drawn.vector) {
      if (job.proceed && !job.proceed()) return false;
      await pdf.addPage({ width: job.width, height: job.height, vector: drawn.vector });
      info.mode = 'vector';
      written = true;
    } else {
      info.reason = drawn.reason;
      if (settings.text === 'vector') throw new PdfExportError('unsupported', `Page ${job.number} cannot be written as text and drawings: ${String(drawn.reason)}`, job.number, drawn.reason);
    }
  }
  if (!written) {
    const picture = await job.picture();
    if (job.proceed && !job.proceed()) return false;
    await pdf.addPage({ width: job.width, height: job.height, image: picture.image, fill: picture.fill, text: picture.text });
  }
  if (onPage) {
    try {
      onPage(info);
    } catch {
      // What the caller does with the news is its own; the file is made all the same.
    }
  }
  return true;
}

/**
 * The limits of a PDF for a device. One page is limited as it is in print
 * (the pixels of a canvas); all pages together are not, because a page is
 * let go of once it is in the file; the file is.
 */
export function pdfLimits(coarse: boolean): PrintLimits {
  return { maxPagePixels: printLimits(coarse).maxPagePixels, maxTotalPixels: Infinity, maxBytes: coarse ? MAX_PDF_FILE_BYTES_COARSE : MAX_PDF_FILE_BYTES };
}

/** Throws the PdfExportError of a signal that was aborted. */
export function checkCancelled(signal: PdfAbortSignal | undefined): void {
  if (signal && signal.aborted === true) throw new PdfExportError('cancelled', 'Making the PDF was cancelled');
}

/**
 * What exportPdf() rejects with for an error of the parts it is made of: the
 * same reason as a PdfExportError. A wrong option (RangeError, TypeError)
 * stays what it is.
 */
export function toPdfExportError(error: unknown): unknown {
  if (error instanceof PdfExportError || error instanceof RangeError || error instanceof TypeError) return error;
  if (error instanceof DocumentPrintError) return new PdfExportError(error.code, error.message);
  if (error instanceof PdfError) return new PdfExportError(error.code === 'too-large' ? 'too-long' : error.code === 'empty' ? 'empty' : 'failed', error.message);
  return new PdfExportError('failed', `The PDF could not be made: ${error instanceof Error ? error.message : String(error)}`);
}

function blobBytes(blob: Blob): Promise<Uint8Array> {
  if (typeof blob.arrayBuffer === 'function') return blob.arrayBuffer().then((buffer) => new Uint8Array(buffer));
  // Browsers from before Blob.arrayBuffer().
  return new Promise<Uint8Array>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
    reader.onerror = () => reject(new PdfExportError('failed', 'The picture of a page could not be read'));
    reader.readAsArrayBuffer(blob);
  });
}

const JPEG = 'image/jpeg';

/** The canvas as a JPEG file; undefined where the browser makes none (it then answers with a PNG). */
function canvasJpeg(canvas: HTMLCanvasElement, quality: number, frame?: (callback: () => void) => unknown): Promise<Uint8Array | undefined> {
  return new Promise<Uint8Array | undefined>((resolve, reject) => {
    const failed = () => reject(new PdfExportError('failed', 'A page could not be turned into a picture'));
    try {
      if (typeof canvas.toBlob === 'function') {
        // Some browsers encode in idle time, and a page on which nothing moves has none for seconds: frames are kept going until the picture is there.
        let waiting = true;
        const tick = () => {
          if (waiting && frame) frame(tick);
        };
        canvas.toBlob(
          (blob) => {
            waiting = false;
            if (!blob) return failed();
            if (blob.type !== JPEG) return resolve(undefined);
            blobBytes(blob).then(resolve, failed);
          },
          JPEG,
          quality,
        );
        tick();
        return;
      }
      const url = canvas.toDataURL(JPEG, quality);
      const head = 'data:' + JPEG + ';base64,';
      if (url.slice(0, head.length) !== head || typeof atob !== 'function') {
        setTimeout(() => resolve(undefined), 0);
        return;
      }
      const binary = atob(url.slice(head.length));
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      // Not in the same turn: the pages are made one per turn, so that the page stays responsive.
      setTimeout(() => resolve(bytes), 0);
    } catch {
      failed();
    }
  });
}

/**
 * What a painted canvas is on a page of the PDF: the colour alone when the
 * whole page has one (an empty page needs no picture), a JPEG, or the pixels
 * themselves - for 'lossless' and where the browser makes no JPEG.
 *
 * For 'auto' the page decides: by the share of its pixels that differ from
 * their neighbour (changeShare) it is text (below TEXT_SHARE) and is stored
 * as its pixels, or a photograph (above PHOTO_SHARE) and becomes a JPEG; in
 * between both are made and the smaller one is taken.
 */
export async function canvasPageImage(
  canvas: HTMLCanvasElement,
  context: CanvasRenderingContext2D,
  settings: PdfSettings,
  frame?: (callback: () => void) => unknown,
): Promise<{ image?: PdfImage; fill?: [number, number, number] }> {
  const width = canvas.width;
  const height = canvas.height;
  if (!(width >= 1) || !(height >= 1)) throw new PdfExportError('failed', 'A page has no pixels');
  const pixels = () => {
    try {
      return imageOfPixels(context.getImageData(0, 0, width, height).data, width, height);
    } catch (error) {
      throw new PdfExportError('failed', `The pixels of a page could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  };
  if (settings.image === 'lossless') return pixels();
  if (settings.image === 'auto') {
    const read = pixels();
    const samples = read.image;
    if (!samples || samples.kind !== 'samples') return read;
    const share = changeShare(samples);
    if (share < TEXT_SHARE) return read;
    const made = await canvasJpeg(canvas, settings.quality, frame);
    if (!made) return read;
    let size: { width: number; height: number; components: 1 | 3 | 4 };
    try {
      size = jpegInfo(made);
    } catch (error) {
      throw new PdfExportError('failed', error instanceof Error ? error.message : String(error));
    }
    const jpeg: PdfImage = { kind: 'jpeg', data: made, width: size.width, height: size.height, components: size.components };
    if (share > PHOTO_SHARE) return { image: jpeg };
    // Neither clearly the one nor the other: the smaller file decides.
    const packed = await packSamples(samples, true, true);
    return { image: made.length < packed.data.length ? jpeg : packed };
  }
  // Strip by strip, until one is not of the colour of the first pixel: a page with anything on it is known after a few.
  let uniform = true;
  let first: [number, number, number] | undefined;
  try {
    for (let y = 0; y < height && uniform; y += STRIP_ROWS) {
      const rows = Math.min(STRIP_ROWS, height - y);
      const data = context.getImageData(0, y, width, rows).data;
      if (!first) first = [data[0], data[1], data[2]];
      for (let at = 0, end = width * rows * 4; at < end; at += 4) {
        if (data[at] !== first[0] || data[at + 1] !== first[1] || data[at + 2] !== first[2]) {
          uniform = false;
          break;
        }
      }
    }
  } catch (error) {
    throw new PdfExportError('failed', `The pixels of a page could not be read: ${error instanceof Error ? error.message : String(error)}`);
  }
  if (uniform && first) return { fill: first };
  const jpeg = await canvasJpeg(canvas, settings.quality, frame);
  if (!jpeg) return pixels();
  let info: { width: number; height: number; components: 1 | 3 | 4 };
  try {
    info = jpegInfo(jpeg);
  } catch (error) {
    throw new PdfExportError('failed', error instanceof Error ? error.message : String(error));
  }
  return { image: { kind: 'jpeg', data: jpeg, width: info.width, height: info.height, components: info.components } };
}

/** Bytes of UTF-8 the name has at most before its extension: with ".pdf" it fits the 255 bytes most file systems allow, and paths have room left. */
const MAX_NAME_BYTES = 200;
/** Names Windows keeps for its devices, whatever extension follows them. */
const DEVICE_NAME = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\.|$)/i;

/**
 * The name of the PDF of a file: the name of the file with ".pdf" for its
 * extension, without anything of a path, without the characters file systems
 * refuse, without controls and without the marks that turn the direction of
 * text around (they can make a name look like another one); at most 200
 * bytes long as UTF-8, never cut inside a character; with "_" before a name
 * Windows keeps for a device ("CON", "LPT1"). `fallback` + ".pdf" when no
 * name is left.
 */
export function pdfFileName(name: string | undefined, fallback: string): string {
  let base = typeof name === 'string' ? name : '';
  const slash = Math.max(base.lastIndexOf('/'), base.lastIndexOf('\\'));
  if (slash >= 0) base = base.slice(slash + 1);
  const dot = base.lastIndexOf('.');
  if (dot > 0) base = base.slice(0, dot);
  base = base
    .replace(/[\u0080-\u009f\u202a-\u202e\u2066-\u2069]/g, '')
    .replace(/[\u0000-\u001f\u007f<>:"|?*]/g, '_')
    .trim();
  // Character by character, as long as the bytes last.
  let kept = '';
  let bytes = 0;
  for (let i = 0; i < base.length; i++) {
    let code = base.charCodeAt(i);
    let units = 1;
    if (code >= 0xd800 && code <= 0xdbff && i + 1 < base.length && (base.charCodeAt(i + 1) & 0xfc00) === 0xdc00) {
      code = 0x10000;
      units = 2;
    } else if (code >= 0xd800 && code <= 0xdfff) {
      // Half a pair is no character of a name.
      continue;
    }
    const size = code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4;
    if (bytes + size > MAX_NAME_BYTES) break;
    bytes += size;
    kept += base.substr(i, units);
    i += units - 1;
  }
  base = kept.trim();
  if (base === '' || /^\.+$/.test(base)) base = fallback;
  if (DEVICE_NAME.test(base)) base = '_' + base;
  return base + '.pdf';
}

/**
 * Hands a file to the user as a download: a link to the parts as one blob,
 * pressed and taken away again. Returns false where a page cannot do that.
 */
export function downloadParts(doc: Document, parts: readonly Uint8Array[], fileName: string): boolean {
  const win = doc.defaultView;
  const urls = (win ? (win as unknown as { URL?: typeof URL }).URL : undefined) || (typeof URL !== 'undefined' ? URL : undefined);
  if (!win || !urls || typeof urls.createObjectURL !== 'function' || typeof Blob === 'undefined') return false;
  const url = urls.createObjectURL(new Blob(parts as unknown as BlobPart[], { type: 'application/pdf' }));
  const link = doc.createElement('a');
  link.href = url;
  link.download = fileName;
  link.rel = 'noopener';
  link.style.cssText = 'display:none';
  (doc.body || doc.documentElement).appendChild(link);
  try {
    link.click();
  } finally {
    if (link.parentNode) link.parentNode.removeChild(link);
    // The download has started by then; the blob is given back.
    win.setTimeout(() => {
      try {
        urls.revokeObjectURL(url);
      } catch {
        // Nothing depends on it.
      }
    }, 60000);
  }
  return true;
}
