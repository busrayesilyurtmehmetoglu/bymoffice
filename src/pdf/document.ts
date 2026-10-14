// A PDF of pages that are pictures with their text lying invisible over
// them. Pages are written as they are added: a page is compressed, goes into
// the file and is let go of, so that a long document never holds more than
// one of its pictures.
//
// A page may instead be what a PDF canvas wrote of it (canvas.ts): drawings,
// text in embedded fonts and pictures of their own. Such a page brings the
// fonts and pictures it needs; the document writes each of them once, the
// fonts at the end, when the glyphs all pages use of them are known.

import type { PdfPicture, PdfVectorPage } from './canvas';
import { PdfEmbeddedFont } from './embed';
import { adler32, zlibDeflate } from './flate';
import { PdfTextFont } from './font';
import { jpegInfo, packSamples, type PdfImage, type PdfPacked } from './image';
import { MAX_FILE_LINKS, MAX_LINK_RECTS, MAX_OUTLINE_ENTRIES, MAX_PAGE_LINKS, emptyLinkReport, finitePlace, outlineTitle, outlineTree, pdfLinkUri, type PdfDestination, type PdfLink, type PdfLinkReport, type PdfOutlineEntry } from './links';
import { textLayer, type PdfTextRun } from './text';
import type { PdfFontFace } from './sfnt';
import { PdfError, PdfRef, PdfString, PdfWriter, joinParts, latin1, name, pdfDate, pdfNumber, pdfText, type PdfDict, type PdfValue } from './writer';

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
  /**
   * The page as a PDF canvas wrote it. It is the whole page: `image`, `fill`
   * and `text` are not looked at with it (its text is text already).
   */
  vector?: PdfVectorPage;
  /**
   * The links of the page, in points from its top left corner (see
   * links.ts): written as link annotations over the page, whichever way the
   * page itself is written.
   */
  links?: readonly PdfLink[];
}

/** A page of a PDF may be 200 inches long at most. */
export const MAX_PAGE_POINTS = 14400;
/** And no smaller than this: what readers of PDF files take for a page at all. */
export const MIN_PAGE_POINTS = 3;
const MAX_INFO_CHARS = 1024;
/**
 * Links into the file held until the end of the file, of one page and of all
 * pages: more than can be written (some may have no page to go to), but not
 * without end.
 */
const MAX_WAITING_PAGE_LINKS = MAX_PAGE_LINKS * 4;
const MAX_WAITING_LINKS = MAX_FILE_LINKS * 4;
const PRODUCER = 'BYM Office';

const whole = (value: unknown): value is number => typeof value === 'number' && value >= 1 && Math.floor(value) === value && value <= 0x7fffffff;

export class PdfDocument {
  private readonly writer: PdfWriter;
  private readonly options: PdfDocumentOptions;
  private readonly pagesRef: PdfRef;
  private fontRef: PdfRef | undefined;
  private readonly font = new PdfTextFont();
  /** The fonts of the drawn pages, by their font file. */
  private readonly faces = new Map<PdfFontFace, PdfEmbeddedFont>();
  /** The states of opacity the drawn pages use, by what they say. */
  private readonly states = new Map<string, PdfRef>();
  /** The pictures of the drawn pages that are in the file, with where. */
  private readonly pictures = new WeakMap<PdfPicture, PdfRef>();
  private readonly kids: PdfRef[] = [];
  /** Changes with everything that is written: what the identifier of the file is made of. */
  private readonly digest = [0x811c9dc5, 0x01000193, 0x9e3779b9, 0x85ebca6b];
  private busy = false;
  private over = false;
  /** The heights of the pages written, in points: where a place on a page is from its bottom. */
  private readonly heights: number[] = [];
  /**
   * The annotations of the pages with links into the file: such a page names
   * its /Annots array by reference, and the array (and the links into the
   * file in it) is written when the file is ended, when every page a link can
   * go to is in it. A link whose page is not is then left out of its array.
   */
  private readonly pending: { array: PdfRef; items: (PdfRef | { rects: number[][]; dest: PdfDestination })[]; written: number }[] = [];
  private outline: PdfOutlineEntry[] = [];
  private outlineDropped = 0;
  /** Links written (out of the file as they come, into it at the end): what the limits count. */
  private linkCount = 0;
  /** Links into the file waiting for the end of the file. */
  private waiting = 0;
  private readonly report: PdfLinkReport = emptyLinkReport();

  constructor(options: PdfDocumentOptions = {}) {
    this.options = options;
    this.writer = new PdfWriter({ maxBytes: options.maxBytes });
    this.pagesRef = this.writer.reserve();
  }

  /** Pages added so far. */
  get pageCount(): number {
    return this.kids.length;
  }

  /** What became of the links and the outline so far; final once the file is finished. */
  get links(): PdfLinkReport {
    return { ...this.report };
  }

  /**
   * The outline of the file: its entries in the order they are shown, each
   * with its level (see links.ts). Entries beyond MAX_OUTLINE_ENTRIES are left
   * out and counted; an entry without a title, or whose page is not in the
   * file when it is ended, is left out. Set before finish(); a second call
   * replaces the first.
   */
  setOutline(entries: readonly PdfOutlineEntry[]): void {
    if (this.over) throw new PdfError('invalid', 'The PDF is finished: no outline can be given to it');
    const kept: PdfOutlineEntry[] = [];
    let dropped = 0;
    const list = Array.isArray(entries) ? entries : [];
    for (let i = 0; i < list.length; i++) {
      const entry = list[i];
      if (!entry || typeof entry.title !== 'string' || !entry.dest) continue;
      if (kept.length >= MAX_OUTLINE_ENTRIES) {
        dropped = list.length - i;
        break;
      }
      kept.push(entry);
    }
    this.outlineDropped = dropped;
    this.outline = kept;
  }

  /** The rectangles of a link as they are written, [left bottom right top] in points, cut to the page; empty ones are left out. */
  private linkRects(link: PdfLink, width: number, height: number): number[][] {
    const out: number[][] = [];
    const rects = Array.isArray(link.rects) ? link.rects : [];
    for (let i = 0; i < rects.length && out.length < MAX_LINK_RECTS; i++) {
      const rect = rects[i];
      if (!rect || !finitePlace(rect.x) || !finitePlace(rect.y) || !finitePlace(rect.width) || !finitePlace(rect.height)) continue;
      const left = Math.max(0, Math.min(rect.x, rect.x + rect.width));
      const right = Math.min(width, Math.max(rect.x, rect.x + rect.width));
      const top = Math.max(0, Math.min(rect.y, rect.y + rect.height));
      const bottom = Math.min(height, Math.max(rect.y, rect.y + rect.height));
      if (!(right - left >= 0.01) || !(bottom - top >= 0.01)) continue;
      out.push([Number(pdfNumber(left, 2)), Number(pdfNumber(height - bottom, 2)), Number(pdfNumber(right, 2)), Number(pdfNumber(height - top, 2))]);
    }
    return out;
  }

  /**
   * The /Annots of a page: the annotations of its links that go out of the
   * file are written now; with a link into the file among them, the array is
   * a reference to one that is written when the file is ended. Undefined for
   * a page without links.
   */
  private annotations(links: readonly PdfLink[] | undefined, width: number, height: number): PdfRef[] | PdfRef | undefined {
    const items: (PdfRef | { rects: number[][]; dest: PdfDestination })[] = [];
    let inward = false;
    if (!Array.isArray(links) || links.length === 0) return undefined;
    const report = this.report;
    // Only what is written counts against the limits: a link into the file is counted when it is written, at the end.
    let onPage = 0;
    let waiting = 0;
    for (let i = 0; i < links.length; i++) {
      const link = links[i];
      if (!link) continue;
      const inwardLink = !!link.dest;
      if (inwardLink ? waiting >= MAX_WAITING_PAGE_LINKS || this.waiting >= MAX_WAITING_LINKS : onPage >= MAX_PAGE_LINKS || this.linkCount >= MAX_FILE_LINKS) {
        report.dropped++;
        continue;
      }
      let uri: string | undefined;
      let dest: PdfDestination | undefined;
      if (link.dest) {
        const page = link.dest.page;
        if (typeof page !== 'number' || !(page >= 0) || Math.floor(page) !== page || page > 0x7fffffff) {
          report.unresolved++;
          continue;
        }
        dest = { page };
        if (finitePlace(link.dest.x)) dest.x = link.dest.x;
        if (finitePlace(link.dest.y)) dest.y = link.dest.y;
      } else {
        uri = pdfLinkUri(link.uri);
        if (uri === undefined) {
          report.refused++;
          continue;
        }
      }
      const rects = this.linkRects(link, width, height);
      if (rects.length === 0) {
        // A link that lies nowhere on the page cannot be pressed.
        report.dropped++;
        continue;
      }
      if (dest) {
        // Counted when it is written, once its page is known to be in the file.
        items.push({ rects, dest });
        inward = true;
        waiting++;
        this.waiting++;
        continue;
      }
      onPage++;
      this.linkCount++;
      report.external++;
      for (let r = 0; r < rects.length; r++) items.push(this.writer.add(linkAnnotation(rects[r], { A: { S: name('URI'), URI: new PdfString(latin1(uri as string)) } })));
    }
    if (items.length === 0) return undefined;
    if (!inward) return items as PdfRef[];
    const array = this.writer.reserve();
    this.pending.push({ array, items, written: onPage });
    return array;
  }

  /** [page /XYZ left top null] of a place in the file; undefined for a page that is not in it. */
  private destination(dest: PdfDestination): PdfValue[] | undefined {
    const page = this.kids[dest.page];
    if (!page) return undefined;
    const height = this.heights[dest.page];
    const left = finitePlace(dest.x) ? Number(pdfNumber(Math.max(0, dest.x), 2)) : null;
    const top = finitePlace(dest.y) ? Number(pdfNumber(Math.min(height, Math.max(0, height - dest.y)), 2)) : null;
    return [page, name('XYZ'), left, top, null];
  }

  /** Writes the outline; its reference, undefined when no entry is left. */
  private writeOutline(): PdfRef | undefined {
    const usable: PdfOutlineEntry[] = [];
    for (let i = 0; i < this.outline.length; i++) {
      const entry = this.outline[i];
      const title = outlineTitle(entry.title);
      if (title === '' || !this.destination(entry.dest)) {
        this.report.outlineSkipped++;
        continue;
      }
      usable.push({ title, level: entry.level, dest: entry.dest });
    }
    this.outline = [];
    this.report.dropped += this.outlineDropped;
    if (usable.length === 0) return undefined;
    const writer = this.writer;
    const { nodes, top } = outlineTree(usable);
    const root = writer.reserve();
    const refs: PdfRef[] = [];
    for (let i = 0; i < nodes.length; i++) refs.push(writer.reserve());
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i];
      const siblings = node.parent < 0 ? top : nodes[node.parent].children;
      const at = node.index;
      const dict: PdfDict = { Title: pdfText(node.entry.title), Parent: node.parent < 0 ? root : refs[node.parent] };
      if (at > 0) dict.Prev = refs[siblings[at - 1]];
      if (at < siblings.length - 1) dict.Next = refs[siblings[at + 1]];
      if (node.children.length > 0) {
        dict.First = refs[node.children[0]];
        dict.Last = refs[node.children[node.children.length - 1]];
        // Open: every entry under it is shown.
        dict.Count = node.descendants;
      }
      dict.Dest = this.destination(node.entry.dest);
      writer.add(dict, refs[i]);
    }
    writer.add({ Type: name('Outlines'), First: refs[top[0]], Last: refs[top[top.length - 1]], Count: nodes.length }, root);
    this.report.outline = nodes.length;
    return root;
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

  /** The resources and the content of a page a PDF canvas wrote. */
  private async drawn(vector: PdfVectorPage, resources: PdfDict): Promise<string> {
    if (!vector || typeof vector.content !== 'string' || !Array.isArray(vector.fonts) || !Array.isArray(vector.pictures) || !Array.isArray(vector.alphas)) throw new PdfError('invalid', 'A drawn page is not what a PDF canvas gives');
    if (vector.fonts.length > 0) {
      const fonts: PdfDict = {};
      for (let i = 0; i < vector.fonts.length; i++) {
        const used = vector.fonts[i];
        let font = this.faces.get(used.face);
        if (!font) {
          font = new PdfEmbeddedFont(used.face, this.writer.reserve());
          this.faces.set(used.face, font);
        }
        font.use(used.glyphs);
        fonts['F' + i] = font.ref;
      }
      resources.Font = fonts;
    }
    if (vector.pictures.length > 0) {
      const pictures: PdfDict = {};
      for (let i = 0; i < vector.pictures.length; i++) {
        const picture: PdfPicture = vector.pictures[i];
        // A picture that is in this file already is drawn again, not written again.
        if (!picture.ref || this.pictures.get(picture) !== picture.ref) {
          if (!picture.image) throw new PdfError('invalid', 'A picture of a drawn page was written into another file and let go of');
          picture.ref = await this.image(picture.image);
          this.pictures.set(picture, picture.ref);
          // The file has it now: its pixels are not held for the pages to come.
          picture.image = undefined;
        }
        pictures['I' + i] = picture.ref;
      }
      resources.XObject = pictures;
    }
    if (vector.alphas.length > 0) {
      const states: PdfDict = {};
      for (let i = 0; i < vector.alphas.length; i++) {
        const alpha = vector.alphas[i];
        const value = Number(pdfNumber(Math.min(1, Math.max(0, alpha.value)), 3));
        const key = (alpha.stroke ? 'CA' : 'ca') + value;
        let ref = this.states.get(key);
        if (!ref) {
          ref = this.writer.add(alpha.stroke ? { Type: name('ExtGState'), CA: value } : { Type: name('ExtGState'), ca: value });
          this.states.set(key, ref);
        }
        states['G' + i] = ref;
      }
      resources.ExtGState = states;
    }
    return vector.content;
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
      if (page.vector) {
        const drawn = await this.drawn(page.vector, resources);
        if (drawn !== '') content.push(drawn);
      } else if (page.image) {
        resources.XObject = { Im0: await this.image(page.image) };
        // An image is one unit wide and high: the matrix makes it the page.
        content.push(`q ${w} 0 0 ${h} 0 0 cm /Im0 Do Q`);
      } else if (page.fill) {
        const fill = page.fill;
        const part = (value: number): string => pdfNumber(Math.min(255, Math.max(0, typeof value === 'number' && isFinite(value) ? value : 255)) / 255, 3);
        if (!(fill[0] === 255 && fill[1] === 255 && fill[2] === 255)) content.push(`${part(fill[0])} ${part(fill[1])} ${part(fill[2])} rg 0 0 ${w} ${h} re f`);
      }
      const text = page.vector ? undefined : page.text;
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
      const annotations = this.annotations(page.links, Number(w), Number(h));
      if (annotations) dict.Annots = annotations;
      this.heights.push(Number(h));
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
    const faces = Array.from(this.faces.values());
    const names = new Set<string>();
    for (let i = 0; i < faces.length; i++) {
      let map: string;
      try {
        map = await faces[i].write(writer, options.nativeCompression !== false, names);
      } catch (error) {
        // The glyphs were found sound when they were written on their pages: this is not expected. The file cannot
        // be ended; whoever makes it makes it anew without this font.
        if (error instanceof PdfError && error.code !== 'invalid') throw error;
        throw new PdfError('font', `The font ${faces[i].face.name} could not be put into the file: ${error instanceof Error ? error.message : String(error)}`);
      }
      this.mix(adler32(latin1(map)));
    }
    writer.add({ Type: name('Pages'), Kids: this.kids, Count: this.kids.length }, this.pagesRef);
    // The links into the file, now that every page they can go to is in it; one whose page is not is left out of its page.
    const pending = this.pending;
    for (let i = 0; i < pending.length; i++) {
      const refs: PdfRef[] = [];
      const items = pending[i].items;
      let onPage = pending[i].written;
      for (let k = 0; k < items.length; k++) {
        const item = items[k];
        if (item instanceof PdfRef) {
          refs.push(item);
          continue;
        }
        const dest = this.destination(item.dest);
        if (!dest) {
          this.report.unresolved++;
          continue;
        }
        if (onPage >= MAX_PAGE_LINKS || this.linkCount >= MAX_FILE_LINKS) {
          this.report.dropped++;
          continue;
        }
        onPage++;
        this.linkCount++;
        this.report.internal++;
        for (let r = 0; r < item.rects.length; r++) refs.push(writer.add(linkAnnotation(item.rects[r], { Dest: dest })));
      }
      writer.add(refs, pending[i].array);
    }
    pending.length = 0;
    const catalog: PdfDict = { Type: name('Catalog'), Pages: this.pagesRef };
    const outline = this.writeOutline();
    if (outline) {
      catalog.Outlines = outline;
      catalog.PageMode = name('UseOutlines');
    }
    const report = this.report;
    this.mix(report.external * 31 + report.internal * 7 + report.outline);
    const root = writer.add(catalog);
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

/** A link annotation: no border, printed, with what it does. */
function linkAnnotation(rect: number[], action: PdfDict): PdfDict {
  return { Type: name('Annot'), Subtype: name('Link'), Rect: rect, Border: [0, 0, 0], F: 4, ...action };
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
