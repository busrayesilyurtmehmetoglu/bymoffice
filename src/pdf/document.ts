// A PDF of pages that are pictures with their text lying invisible over
// them. Pages are written as they are added: a page is compressed, goes into
// the file and is let go of, so that a long document never holds more than
// one of its pictures. What a later stage adds (text drawn with embedded
// fonts, links) are further resources and operators of a page; the objects,
// the page tree and the trailer stay as they are here.

import { adler32, zlibDeflate } from './flate';
import { PdfTextFont } from './font';
import { jpegInfo, packSamples, type PdfImage, type PdfPacked } from './image';
import { textLayer, type PdfTextRun } from './text';
import { PdfError, PdfWriter, joinParts, latin1, name, pdfDate, pdfNumber, pdfText, type PdfDict, type PdfRef } from './writer';

export interface PdfDocumentOptions {
  /** The title and the author in the information of the document; left out when missing or empty. */
  title?: string;
  author?: string;
  /** When the document was made; now when missing. */
  date?: Date;
  /** The file may not grow beyond this many bytes (4 GB at most). */
  maxBytes?: number;
  /** False keeps to the engine's own compression where the built-in one could be used. */
  nativeCompression?: boolean;
}

export interface PdfPage {
  /** Size of the page in points (1/72 inch). */
  width: number;
  height: number;
  /** The picture that fills the page. */
  image?: PdfImage;
  /** Without a picture: the colour of the paper (0 to 255 each); white is not painted. */
  fill?: [number, number, number];
  /** The text of the page: the runs, and the width of the page in the unit they are measured in. */
  text?: { runs: readonly PdfTextRun[]; width: number };
}

/** A page of a PDF may be 200 inches long at most. */
export const MAX_PAGE_POINTS = 14400;
/** And no smaller than this: what readers of PDF files take for a page at all. */
export const MIN_PAGE_POINTS = 3;
const MAX_INFO_CHARS = 1024;
const PRODUCER = 'BYM Office';

const whole = (value: unknown): value is number => typeof value === 'number' && value >= 1 && Math.floor(value) === value && value <= 0x7fffffff;

export class PdfDocument {
  private readonly writer: PdfWriter;
  private readonly options: PdfDocumentOptions;
  private readonly pagesRef: PdfRef;
  private fontRef: PdfRef | undefined;
  private readonly font = new PdfTextFont();
  private readonly kids: PdfRef[] = [];
  /** Changes with everything that is written: what the identifier of the file is made of. */
  private readonly digest = [0x811c9dc5, 0x01000193, 0x9e3779b9, 0x85ebca6b];
  private busy = false;
  private over = false;

  constructor(options: PdfDocumentOptions = {}) {
    this.options = options;
    this.writer = new PdfWriter({ maxBytes: options.maxBytes });
    this.pagesRef = this.writer.reserve();
  }

  /** Pages added so far. */
  get pageCount(): number {
    return this.kids.length;
  }

  /** Bytes of the file so far. */
  get size(): number {
    return this.writer.size;
  }

  /** The parts of the file written since the last call, in order. */
  takeParts(): Uint8Array[] {
    return this.writer.takeParts();
  }

  private mix(value: number): void {
    const digest = this.digest;
    for (let i = 0; i < 4; i++) {
      // FNV-1a over the four bytes of the value, from four different starts.
      let hash = digest[i];
      for (let b = 0; b < 4; b++) {
        hash ^= ((value >>> (b * 8)) & 255) ^ i;
        hash = Math.imul(hash, 0x01000193) >>> 0;
      }
      digest[i] = hash;
    }
  }

  private async stream(dict: PdfDict, data: Uint8Array): Promise<PdfRef> {
    const packed = await zlibDeflate(data, this.options.nativeCompression !== false);
    this.mix(adler32(packed));
    return this.writer.addStream({ ...dict, Filter: name('FlateDecode') }, packed);
  }

  private async image(image: PdfImage): Promise<PdfRef> {
    if (!image || !whole(image.width) || !whole(image.height) || !image.data) throw new PdfError('invalid', 'A picture needs a width and a height in whole pixels');
    const base: PdfDict = { Type: name('XObject'), Subtype: name('Image'), Width: image.width, Height: image.height, BitsPerComponent: 8 };
    if (image.kind === 'jpeg') {
      const info = jpegInfo(image.data);
      if (info.width !== image.width || info.height !== image.height || info.components !== image.components) throw new PdfError('invalid', 'The JPEG is not of the size and the components it is said to be');
      base.ColorSpace = name(info.components === 1 ? 'DeviceGray' : info.components === 3 ? 'DeviceRGB' : 'DeviceCMYK');
      // Adobe's programs write CMYK with 255 for "no ink"; a PDF reads 0 for that.
      if (info.components === 4 && info.adobe) base.Decode = [1, 0, 1, 0, 1, 0, 1, 0];
      base.Filter = name('DCTDecode');
      this.mix(adler32(image.data));
      return this.writer.addStream(base, image.data);
    }
    if ((image.kind !== 'samples' && image.kind !== 'packed') || (image.colors !== 1 && image.colors !== 3)) throw new PdfError('invalid', 'A picture is neither a JPEG nor samples');
    const pixels = image.width * image.height;
    base.ColorSpace = name(image.colors === 1 ? 'DeviceGray' : 'DeviceRGB');
    let packed: PdfPacked;
    if (image.kind === 'packed') packed = image;
    else {
      if (image.data.length !== pixels * image.colors) throw new PdfError('invalid', `A picture of ${image.width} x ${image.height} with ${image.colors} components needs ${pixels * image.colors} bytes`);
      if (image.alpha) {
        if (image.alpha.length !== pixels) throw new PdfError('invalid', 'The mask of a picture needs one byte per pixel');
        base.SMask = await this.stream({ Type: name('XObject'), Subtype: name('Image'), Width: image.width, Height: image.height, BitsPerComponent: 8, ColorSpace: name('DeviceGray') }, image.alpha);
      }
      packed = await packSamples(image, this.options.nativeCompression !== false);
    }
    base.Filter = name('FlateDecode');
    if (packed.predictor) base.DecodeParms = { Predictor: 15, Colors: packed.colors, BitsPerComponent: 8, Columns: packed.width };
    this.mix(adler32(packed.data));
    return this.writer.addStream(base, packed.data);
  }

  /** Writes a page. The next one may be added when the promise is resolved. */
  async addPage(page: PdfPage): Promise<void> {
    if (this.over) throw new PdfError('invalid', 'The PDF is finished: no page can be added to it');
    if (this.busy) throw new PdfError('invalid', 'A page is added while the one before it is still being written');
    const width = page ? page.width : NaN;
    const height = page ? page.height : NaN;
    // As they are written: a size that is rounded to nothing is no size.
    if (!(width > 0) || !(height > 0) || !(Number(pdfNumber(width, 2)) >= MIN_PAGE_POINTS) || !(Number(pdfNumber(height, 2)) >= MIN_PAGE_POINTS) || !(width <= MAX_PAGE_POINTS) || !(height <= MAX_PAGE_POINTS))
      throw new PdfError('invalid', `A page needs a width and a height between ${MIN_PAGE_POINTS} and ${MAX_PAGE_POINTS} points`);
    this.busy = true;
    try {
      const resources: PdfDict = {};
      const content: string[] = [];
      const w = pdfNumber(width, 2);
      const h = pdfNumber(height, 2);
      if (page.image) {
        resources.XObject = { Im0: await this.image(page.image) };
        // An image is one unit wide and high: the matrix makes it the page.
        content.push(`q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q`);
      } else if (page.fill) {
        const fill = page.fill;
        const part = (value: number): string => pdfNumber(Math.min(255, Math.max(0, typeof value === 'number' && isFinite(value) ? value : 255)) / 255, 3);
        if (!(fill[0] === 255 && fill[1] === 255 && fill[2] === 255)) content.push(`${part(fill[0])} ${part(fill[1])} ${part(fill[2])} rg 0 0 ${w} ${h} re f`);
      }
      const text = page.text;
      if (text && text.width > 0 && isFinite(text.width)) {
        const scale = width / text.width;
        const layer = textLayer(text.runs, { height: height / scale, scale, font: this.font });
        if (layer !== '') {
          if (!this.fontRef) this.fontRef = this.writer.reserve();
          resources.Font = { F1: this.fontRef };
          content.push(layer.replace(/\n$/, ''));
        }
      }
      const dict: PdfDict = { Type: name('Page'), Parent: this.pagesRef, MediaBox: [0, 0, Number(w), Number(h)], Resources: resources };
      if (content.length > 0) dict.Contents = await this.stream({}, latin1(content.join('\n')));
      this.kids.push(this.writer.add(dict));
      this.mix(this.writer.size);
    } catch (error) {
      // A page that was written in part cannot be taken back: the file is given up.
      this.over = true;
      throw error;
    } finally {
      this.busy = false;
    }
  }

  /** Writes what ends the file: the font, the page tree, the information and the table of objects. */
  async finish(): Promise<void> {
    if (this.over) throw new PdfError('invalid', 'The PDF is finished already');
    if (this.busy) throw new PdfError('invalid', 'The PDF is finished while a page is still being written');
    if (this.kids.length === 0) {
      this.over = true;
      throw new PdfError('empty', 'A PDF needs at least one page');
    }
    this.over = true;
    const writer = this.writer;
    const options = this.options;
    if (this.fontRef) {
      await this.font.write(writer, this.fontRef, options.nativeCompression !== false);
      // Which characters the codes of the text are is part of what the file says.
      this.mix(adler32(latin1(this.font.toUnicode())));
    }
    writer.add({ Type: name('Pages'), Kids: this.kids, Count: this.kids.length }, this.pagesRef);
    const root = writer.add({ Type: name('Catalog'), Pages: this.pagesRef });
    const info: PdfDict = {};
    const said = (value: unknown): string => (typeof value === 'string' ? value.slice(0, MAX_INFO_CHARS) : '');
    const title = pdfText(said(options.title));
    const author = pdfText(said(options.author));
    if (title.bytes.length > 0) info.Title = title;
    if (author.bytes.length > 0) info.Author = author;
    info.Creator = pdfText(PRODUCER);
    info.Producer = pdfText(PRODUCER);
    const date = pdfDate(options.date === undefined ? new Date() : options.date);
    if (date) {
      info.CreationDate = date;
      info.ModDate = date;
    }
    // The identifier: sixteen bytes that change with the content, the information and the size of the file.
    for (const key of Object.keys(info)) this.mix(adler32((info[key] as { bytes: Uint8Array }).bytes));
    this.mix(writer.size);
    const id = new Uint8Array(16);
    for (let i = 0; i < 4; i++) for (let b = 0; b < 4; b++) id[i * 4 + b] = (this.digest[i] >>> (24 - b * 8)) & 255;
    writer.finish({ root, info: writer.add(info), id });
  }
}

/** A PDF of the given pages, as one array. */
export async function buildPdf(pages: readonly PdfPage[], options: PdfDocumentOptions = {}): Promise<Uint8Array> {
  const document = new PdfDocument(options);
  const parts: Uint8Array[] = [];
  for (let i = 0; i < pages.length; i++) {
    await document.addPage(pages[i]);
    const taken = document.takeParts();
    for (let k = 0; k < taken.length; k++) parts.push(taken[k]);
  }
  await document.finish();
  const taken = document.takeParts();
  for (let k = 0; k < taken.length; k++) parts.push(taken[k]);
  return joinParts(parts);
}
