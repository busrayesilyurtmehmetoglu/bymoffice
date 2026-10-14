// Charts of objects embedded in a doc or RTF file: an Excel chart or
// workbook (Excel.Chart.8, Excel.Sheet.8) or a chart of Microsoft Graph
// (MSGraph.Chart.8). Such an object is a compound file of its own - the
// storage ObjectPool/_<id> of a doc file, the native data of the OLE1 object
// in "\objdata" of an RTF file - and holds a BIFF workbook in its stream
// "Workbook" (or "Book"). That workbook is read with the reader of xls files,
// whose charts become the chart of the object; the series are drawn from the
// values the chart cached, and where it cached none from the cells of the
// workbook (see fillChartFromBook). A view for showing: nothing of it is ever
// written, and a document converted to docx carries the picture the file
// stores for the object (see buildDocx).
//
// Which chart an object shows: the one of the sheet the workbook has active,
// when that is a chart sheet; otherwise, for an object of a chart (Excel.Chart,
// MSGraph.Chart), the first chart sheet, then the first chart on a sheet. An
// Excel.Sheet object whose active sheet is a worksheet shows cells, which are
// not drawn here: it is no chart.
//
// UNSURE: no object written by Office was at hand. The streams of an object
// (\1CompObj, Workbook, \3PIC) are those of the specifications as far as they
// are remembered; that a chart of Microsoft Graph 97 keeps its data sheet and
// its chart as a BIFF8 workbook in the stream "Workbook" is an assumption. An
// object whose workbook cannot be read keeps the picture the file stores for
// it, and the document says why in its warnings.
//
// Nothing is trusted: the size of a workbook is checked before it is read,
// the objects of a document share the limits of the charts of a docx file
// (charts, cached values, workbooks and their bytes), and the reader of xls
// files bounds whatever a workbook holds.

import type { Chart, ChartData } from '../../model/sheet/chart';
import type { Workbook } from '../../model/sheet/types';
import { readXlsStream } from '../xls/workbook';
import {
  fillChartFromBook,
  MAX_DOC_CHARTS,
  MAX_DOC_CHART_BOOK_BYTES,
  MAX_DOC_CHART_BOOKS,
  MAX_DOC_CHART_BOOKS_BYTES,
  MAX_DOC_CHART_POINTS,
  MAX_DOC_CHART_WARNINGS,
  type DocChartLimits,
} from '../docx/charts';

/** The streams of a compound file an object is read from; `entries` gives the sizes of streams before they are read. */
export interface OleSource {
  has(path: string): boolean;
  stream(path: string): Uint8Array;
  readonly entries?: readonly { path: string; size: number }[];
}

/** What an embedded object is, by its class: only the kinds that may hold a chart. */
export type OleChartKind = 'excelChart' | 'excelSheet' | 'graph';

/** The kind of an object of the class (a ProgID: "Excel.Chart.8", "MSGraph.Chart.8"); undefined for a class that holds no chart. */
export function oleChartKind(className: string | undefined): OleChartKind | undefined {
  if (typeof className !== 'string') return undefined;
  const name = className.trim();
  if (/^Excel\.Chart(\.\d{1,3})?$/i.test(name)) return 'excelChart';
  if (/^Excel\.Sheet(\.\d{1,3})?$/i.test(name)) return 'excelSheet';
  if (/^MSGraph\.Chart(\.\d{1,3})?$/i.test(name)) return 'graph';
  return undefined;
}

const u32 = (b: Uint8Array, o: number): number => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

/** A string with its length (bytes, the terminating zero among them) in front; undefined where it does not fit. */
function lengthPrefixed(b: Uint8Array, at: number, max: number): { text: string; next: number } | undefined {
  if (at + 4 > b.length) return undefined;
  const length = u32(b, at);
  if (length > max || at + 4 + length > b.length) return undefined;
  let text = '';
  for (let i = at + 4; i < at + 4 + length && b[i] !== 0; i++) text += String.fromCharCode(b[i]);
  return { text, next: at + 4 + length };
}

/**
 * The ProgID the stream "\1CompObj" of an object states: a header of 28 bytes, the name of the
 * type shown to users, the clipboard format (a number behind 0xFFFFFFFF or 0xFFFFFFFE, or a
 * name), then the ProgID; each name a string with its length in front. Undefined when the
 * stream does not hold one.
 */
export function compObjProgId(b: Uint8Array): string | undefined {
  const user = lengthPrefixed(b, 28, 4096);
  if (!user || user.next + 4 > b.length) return undefined;
  let at = user.next;
  const marker = u32(b, at);
  if (marker === 0xffffffff || marker === 0xfffffffe) at += 8;
  else {
    const format = lengthPrefixed(b, at, 4096);
    if (!format) return undefined;
    at = format.next;
  }
  const prog = lengthPrefixed(b, at, 256);
  return prog && prog.text !== '' ? prog.text : undefined;
}

/**
 * An OLE1 object as RTF stores it in "\objdata": the version, the format (2: embedded), the
 * class, the topic and the item (strings with their length in front), then the size of the
 * native data and the data. For an object of OLE2 the native data is a compound file.
 */
export function ole1Native(b: Uint8Array): { className: string; data: Uint8Array } | undefined {
  if (b.length < 8 || u32(b, 4) !== 2) return undefined;
  const className = lengthPrefixed(b, 8, 256);
  if (!className) return undefined;
  const topic = lengthPrefixed(b, className.next, 4096);
  if (!topic) return undefined;
  const item = lengthPrefixed(b, topic.next, 4096);
  if (!item || item.next + 4 > b.length) return undefined;
  const size = u32(b, item.next);
  const start = item.next + 4;
  if (size > b.length - start) return undefined;
  return { className: className.text, data: b.subarray(start, start + size) };
}

/** The size a doc file states for an object in its stream "\3PIC" (a PICF), EMU; undefined when the stream does not state one. */
export function olePictureSize(b: Uint8Array): { width: number; height: number } | undefined {
  if (b.length < 44) return undefined;
  const u16 = (o: number) => b[o] | (b[o + 1] << 8);
  const i16 = (o: number) => (u16(o) << 16) >> 16;
  const scaleX = u16(32) || 1000;
  const scaleY = u16(34) || 1000;
  // The natural size less what is cropped away, scaled (in thousandths); twips, at most 22 inches as for a picture.
  const width = Math.min(31680, ((u16(28) - i16(36) - i16(40)) * scaleX) / 1000);
  const height = Math.min(31680, ((u16(30) - i16(38) - i16(42)) * scaleY) / 1000);
  return width > 0 && height > 0 ? { width: Math.round(width * 635), height: Math.round(height * 635) } : undefined;
}

/** True where a Windows metafile header starts: the type (1 in memory, 2 on disk), a header of nine words, the version 0x0100 or 0x0300. */
function wmfHeaderAt(b: Uint8Array, at: number): boolean {
  return at + 18 <= b.length && (b[at] === 1 || b[at] === 2) && b[at + 1] === 0 && b[at + 2] === 9 && b[at + 3] === 0 && (b[at + 5] === 1 || b[at + 5] === 3) && b[at + 4] === 0;
}

/** Records of a metafile looked at for its window. */
const MAX_WMF_SCAN = 4096;

/**
 * The metafile an object stores as its picture ("\3META"), as a placeable metafile that Word
 * and other readers show: a metafile as it is, or behind the header of a METAFILEPICT (8 or 16
 * bytes); one that is placeable already is kept. The bounding box of the header is the window
 * the metafile sets (the last SetWindowOrg and the last SetWindowExt among its first 4,096 records, in whatever order), scaled so that it
 * fills `width` x `height` twips; without a window, those twips. Undefined for bytes that are
 * no metafile.
 *
 * UNSURE: no "\3META" written by Word was at hand: whether it carries a METAFILEPICT in front,
 * whether a picture without a window is in twips, and whether the last window of the records is the one the picture is drawn in.
 */
export function placeableWmf(b: Uint8Array, width: number, height: number): Uint8Array | undefined {
  if (b.length >= 22 && b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) return wmfHeaderAt(b, 22) ? b : undefined;
  const start = wmfHeaderAt(b, 0) ? 0 : wmfHeaderAt(b, 8) ? 8 : wmfHeaderAt(b, 16) ? 16 : -1;
  if (start < 0) return undefined;
  const i16 = (o: number) => ((b[o] | (b[o + 1] << 8)) << 16) >> 16;
  let left = 0;
  let top = 0;
  let extX = 0;
  let extY = 0;
  // The records after the header of nine words: a size in words, a function, its parameters.
  let p = start + 18;
  for (let n = 0; n < MAX_WMF_SCAN && p + 6 <= b.length; n++) {
    const words = u32(b, p);
    const fn = b[p + 4] | (b[p + 5] << 8);
    if (words < 3 || fn === 0) break;
    // SetWindowOrg and SetWindowExt: y, then x.
    if (fn === 0x020b && p + 10 <= b.length) {
      top = i16(p + 6);
      left = i16(p + 8);
    } else if (fn === 0x020c && p + 10 <= b.length) {
      extY = i16(p + 6);
      extX = i16(p + 8);
    }
    if (words > (b.length - p) / 2) break;
    p += words * 2;
  }
  const twipsX = Math.max(1, Math.min(32767, Math.round(width)));
  const twipsY = Math.max(1, Math.min(32767, Math.round(height)));
  let inch = 1440;
  let right = left + twipsX;
  let bottom = top + twipsY;
  if (extX !== 0 && extY !== 0) {
    right = left + extX;
    bottom = top + extY;
    // Units of the window to an inch, so that the window fills the width.
    inch = Math.max(1, Math.min(65535, Math.round((Math.abs(extX) * 1440) / twipsX)));
  }
  const clamp16 = (v: number) => Math.max(-32768, Math.min(32767, v));
  const head = new Uint8Array(22);
  const view = new DataView(head.buffer);
  view.setUint32(0, 0x9ac6cdd7, true);
  view.setInt16(6, clamp16(Math.min(left, right)), true);
  view.setInt16(8, clamp16(Math.min(top, bottom)), true);
  view.setInt16(10, clamp16(Math.max(left, right)), true);
  view.setInt16(12, clamp16(Math.max(top, bottom)), true);
  view.setUint16(14, inch, true);
  let sum = 0;
  for (let i = 0; i < 10; i++) sum ^= view.getUint16(i * 2, true);
  view.setUint16(20, sum, true);
  const out = new Uint8Array(22 + b.length - start);
  out.set(head);
  out.set(b.subarray(start), 22);
  return out;
}

export interface OleChartRead {
  chart?: Chart;
  /** Why an object of a chart shows no chart; missing for an object that is not one (cells of a worksheet). */
  problem?: string;
  /** The size the workbook states for its chart sheet, EMU. */
  width?: number;
  height?: number;
}

/** Why an object is not read. */
class OleError extends Error {}

/** Bytes of a stream "\1CompObj" or "\3PIC" that are read: a genuine one holds a few dozen. A larger one is not read: the object is no chart. */
export const MAX_OLE_SMALL_STREAM = 4096;

/** True for data the chart cached no value for: a chart of a workbook caches none behind SIIndex, whatever number of values its series states. */
const noPoints = (data: ChartData): boolean => !data.cache || !Array.isArray(data.cache.index) || data.cache.index.length === 0;

/** The points a chart holds in its caches. */
function cachedPoints(chart: Chart): number {
  let points = 0;
  const of = (data: ChartData | undefined) => {
    if (data && data.cache && Array.isArray(data.cache.index)) points += data.cache.index.length;
  };
  for (const group of chart.groups) {
    for (const series of group.series) {
      of(series.values);
      of(series.categories);
      of(series.bubbleSizes);
    }
  }
  return points;
}

/** The chart an object of the kind shows of its workbook, with the size its sheet states. */
function chartOfBook(book: Workbook, kind: OleChartKind): { chart?: Chart; problem?: string; width?: number; height?: number } {
  const sheets = book.sheets;
  const ofSheet = (index: number) => {
    const sheet = sheets[index];
    const drawing = sheet && sheet.drawings ? sheet.drawings.find((one) => one && one.object && one.object.type === 'chart') : undefined;
    if (!drawing || drawing.object.type !== 'chart') return undefined;
    const anchor = drawing.anchor;
    return anchor.kind === 'absolute' ? { chart: drawing.object.chart, width: anchor.width, height: anchor.height } : { chart: drawing.object.chart };
  };
  /** Why the chart of a sheet that should show one was not read: the warning the reader gave it. */
  const why = (index: number): string | undefined => {
    const sheet = sheets[index];
    return sheet && sheet.warnings && sheet.warnings.length > 0 ? sheet.warnings[0] : undefined;
  };
  const active = book.activeSheet >= 0 && book.activeSheet < sheets.length ? book.activeSheet : 0;
  if (sheets[active] && sheets[active].kind === 'chartsheet') {
    const found = ofSheet(active);
    return found || { problem: why(active) || 'its chart sheet holds no chart that is read' };
  }
  // An Excel workbook that shows a worksheet shows cells.
  if (kind === 'excelSheet') return {};
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].kind !== 'chartsheet') continue;
    const found = ofSheet(i);
    if (found) return found;
  }
  for (let i = 0; i < sheets.length; i++) {
    const found = ofSheet(i);
    if (found) return found;
  }
  for (let i = 0; i < sheets.length; i++) if (sheets[i].kind === 'chartsheet' && why(i)) return { problem: why(i) };
  return { problem: 'its workbook holds no chart that is read' };
}

const limitOf = (value: number | undefined, fallback: number): number => (typeof value === 'number' && value >= 0 ? Math.min(value, fallback) : fallback);

/**
 * Reads the charts of the objects of one document, all within one set of limits (those of the
 * charts of a docx file). Never throws: what goes wrong is added to `warnings`.
 */
export class OleChartReader {
  /** The charts read, by the key each object was read under (DrawingItem.chartPart). */
  readonly charts = new Map<string, Chart>();
  /** What reading an object came to, by its key: an object that many characters show is read once. */
  private readonly done = new Map<string, OleChartRead>();
  private readonly maxCharts: number;
  private readonly maxBookBytes: number;
  private readonly budget: { points: number };
  private books: number;
  /** Bytes of the streams of objects that may still be read, all streams together (workbooks, pictures, the small ones). */
  private bytesLeft: number;
  /** The sizes of the streams of a source by path, made once for each source from its entries. */
  private readonly sizes = new WeakMap<object, Map<string, number>>();
  private read = 0;
  private told = 0;
  private unsaid = 0;

  constructor(
    private readonly warnings: string[],
    limits: DocChartLimits = {},
  ) {
    this.maxCharts = limitOf(limits.maxCharts, MAX_DOC_CHARTS);
    this.maxBookBytes = limitOf(limits.maxChartBookBytes, MAX_DOC_CHART_BOOK_BYTES);
    this.budget = { points: limitOf(limits.maxChartPoints, MAX_DOC_CHART_POINTS) };
    this.books = limitOf(limits.maxChartBooks, MAX_DOC_CHART_BOOKS);
    this.bytesLeft = limitOf(limits.maxChartBooksBytes, MAX_DOC_CHART_BOOKS_BYTES);
  }

  /** The size of a stream as the source states it before it is read; undefined when it states none. */
  sizeOf(source: OleSource, path: string): number | undefined {
    if (!source.entries) return undefined;
    let sizes = this.sizes.get(source);
    if (!sizes) {
      sizes = new Map();
      for (const entry of source.entries) sizes.set(entry.path, entry.size);
      this.sizes.set(source, sizes);
    }
    return sizes.get(path);
  }

  /**
   * Reads a stream of an object of at most `max` bytes, within what the streams of objects may
   * still take; the bytes are taken from that. A stream that is missing, larger, or beyond what
   * is left gives a reason instead (nothing is read of a stream whose stated size is beyond), and
   * one that cannot be read the error.
   */
  readStream(source: OleSource, path: string, max: number): Uint8Array | 'missing' | 'large' | 'budget' | Error {
    if (!source.has(path)) return 'missing';
    const stated = this.sizeOf(source, path);
    if (stated !== undefined && stated > max) return 'large';
    if (stated !== undefined && stated > this.bytesLeft) return 'budget';
    let bytes: Uint8Array;
    try {
      bytes = source.stream(path);
    } catch (error) {
      return error instanceof Error ? error : new Error(String(error));
    }
    if (bytes.length > max) return 'large';
    if (bytes.length > this.bytesLeft) return 'budget';
    this.bytesLeft -= bytes.length;
    return bytes;
  }

  /** The kind of the object whose streams start with `prefix` ("ObjectPool/_12/", "" for the root): by its \1CompObj, else by the class the file names for it. */
  kindOf(source: OleSource, prefix: string, className?: string): OleChartKind | undefined {
    const bytes = this.readStream(source, prefix + '\u0001CompObj', MAX_OLE_SMALL_STREAM);
    const kind = bytes instanceof Uint8Array ? oleChartKind(compObjProgId(bytes)) : undefined;
    return kind || oleChartKind(className);
  }

  /**
   * Reads the chart of the object whose streams start with `prefix`, under `key`. `label`
   * names the object in a warning. Gives the chart, or why an object of a chart shows none (the
   * warning is given here), or neither for an object that shows cells.
   */
  readChart(source: OleSource, prefix: string, key: string, kind: OleChartKind, label: string): OleChartRead {
    const known = this.done.get(key);
    if (known) return known;
    let out: OleChartRead;
    try {
      out = this.readBook(source, prefix, key, kind);
    } catch (error) {
      out = { problem: error instanceof OleError ? error.message : `its workbook cannot be read (${error instanceof Error ? error.message : String(error)})` };
    }
    this.done.set(key, out);
    if (out.problem !== undefined) this.warn(`The chart of an embedded object (${label}) is not shown: ${out.problem}`);
    return out;
  }

  private readBook(source: OleSource, prefix: string, key: string, kind: OleChartKind): OleChartRead {
    const name = source.has(prefix + 'Workbook') ? 'Workbook' : source.has(prefix + 'Book') ? 'Book' : undefined;
    if (name === undefined) {
      throw new OleError(source.has(prefix + 'Package') ? 'its workbook is a package (xlsx), which is not read in a document of this format' : 'it holds no workbook');
    }
    if (this.read >= this.maxCharts) throw new OleError(`the document holds more than ${this.maxCharts} charts`);
    if (this.books <= 0) throw new OleError('the document holds too many embedded workbooks');
    // A stream beyond the limits is not even read.
    const bytes = this.readStream(source, prefix + name, this.maxBookBytes);
    if (bytes === 'large') throw new OleError('its workbook is too large');
    if (bytes === 'budget') throw new OleError('the embedded objects of the document hold too many bytes');
    if (!(bytes instanceof Uint8Array)) throw new OleError(`its workbook cannot be read${bytes instanceof Error ? ` (${bytes.message})` : ''}`);
    this.books--;
    let book: Workbook;
    try {
      book = readXlsStream(bytes, name).workbook;
    } catch (error) {
      throw new OleError(`its workbook cannot be read (${error instanceof Error ? error.message : String(error)})`);
    }
    const found = chartOfBook(book, kind);
    if (!found.chart) return found.problem === undefined ? {} : { problem: found.problem };
    const chart = found.chart;
    const points = cachedPoints(chart);
    if (points > this.budget.points) throw new OleError('the charts of the document hold too many values');
    this.budget.points -= points;
    this.read++;
    // What the chart cached nothing for is taken from the cells of its workbook.
    fillChartFromBook(chart, book, this.budget, () => true, noPoints);
    this.charts.set(key, chart);
    const out: OleChartRead = { chart };
    if (found.width !== undefined && found.height !== undefined && found.width > 0 && found.height > 0) {
      out.width = found.width;
      out.height = found.height;
    }
    return out;
  }

  /** Adds a warning about the chart of an object, up to the limit of them; the rest are counted (see finish()). */
  warn(text: string): void {
    if (this.told < MAX_DOC_CHART_WARNINGS) {
      this.warnings.push(text);
      this.told++;
    } else this.unsaid++;
  }

  /** Says how many more objects had a chart or a picture that is not shown; once, after the last object was read. */
  finish(): void {
    if (this.unsaid > 0) this.warnings.push(`More embedded objects whose chart or picture is not shown: ${this.unsaid}`);
    this.unsaid = 0;
  }
}
