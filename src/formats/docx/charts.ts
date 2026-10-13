// The charts of a document: the chart parts its frames name
// (word/charts/chartN.xml), read with the reader of the charts of a workbook
// into the same model. A chart of a document is drawn from the values its
// part cached (<c:numCache>, <c:strCache>, literals). Only where a reference
// of a series has no cache at all - no cache element; one that is there and
// holds no point is a cache, and is what is drawn - the workbook Word embeds
// beside the part (word/embeddings/*.xlsx, named by <c:externalData>) is
// opened, once per workbook and within the limits below, and the cells the
// reference names are taken as the cache the part lacks. The notes of the
// chart say which of its parts were read there, and which stayed empty and
// why. That happens while the document is read:
// the reader waits for its parts anyway, the charts of a document are painted
// from their models alone, and a workbook that is read is let go of again.
//
// Nothing here is ever serialized: the chart part, its relationships and the
// embedded workbook stay in the package and are written back from there. A
// part that cannot be read, that holds nothing that is drawn, or that lies
// beyond the limits of the document is said in the warnings of the document;
// its frame shows a placeholder.

import { CfbFile } from '../../cfb';
import { MAX_WORKBOOK_CHARTS, MAX_WORKBOOK_CHART_POINTS, type Chart, type ChartData } from '../../model/sheet/chart';
import type { Workbook } from '../../model/sheet/types';
import { OpcPackage } from '../../opc';
import { parseXml, type XmlElement } from '../../xml';
import { readXls } from '../xls/workbook';
import { ChartBudgetError, parseChart } from '../xlsx/chart';
import { chartCellsOf } from '../xlsx/chart-write';
import { readXlsx } from '../xlsx/workbook';
import { MAX_CHART_PART_BYTES, MAX_WORKBOOK_CHART_BYTES } from '../xlsx/drawings';
import { quote } from './common';

/** Chart parts read for one document, and cached values read for all of them together: the limits of a workbook. */
export const MAX_DOC_CHARTS = MAX_WORKBOOK_CHARTS;
export const MAX_DOC_CHART_POINTS = MAX_WORKBOOK_CHART_POINTS;
/** Bytes of chart parts read for one document, all charts together: the limit of a workbook. */
export const MAX_DOC_CHART_BYTES = MAX_WORKBOOK_CHART_BYTES;
/** Bytes of one embedded workbook that is opened for a chart, all its parts uncompressed together; a larger one is not opened. */
export const MAX_DOC_CHART_BOOK_BYTES = 16 * 1024 * 1024;
/** Bytes of the embedded workbooks opened for the charts of one document, all together. */
export const MAX_DOC_CHART_BOOKS_BYTES = 64 * 1024 * 1024;
/** Embedded workbooks opened for the charts of one document. */
export const MAX_DOC_CHART_BOOKS = 64;
/** Warnings about single charts of one document; what goes wrong beyond them is counted in one line. */
export const MAX_DOC_CHART_WARNINGS = 20;

const EMPTY_NOTE = 'Shown empty, without cached values: ';
const BOOK_NOTE = 'Read from the embedded workbook of the chart, where the chart part caches nothing: ';

/** What of a chart may lack a cache, in the order the notes name it. */
type Wanted = 'values of series' | 'categories' | 'bubble sizes' | 'names of series' | 'titles';
const WANTED: Wanted[] = ['values of series', 'categories', 'bubble sizes', 'names of series', 'titles'];

export interface DocChartLimits {
  /** Chart parts that are read. Default MAX_DOC_CHARTS. */
  maxCharts?: number;
  /** Cached values (<c:pt>) that are read, all charts together. Default MAX_DOC_CHART_POINTS. */
  maxChartPoints?: number;
  /** A chart part larger than this is not parsed. Default 32 MB. */
  maxChartPartBytes?: number;
  /** Bytes of chart parts that are read, all charts together. Default MAX_DOC_CHART_BYTES. */
  maxChartBytes?: number;
  /** An embedded workbook larger than this (uncompressed) is not opened. Default MAX_DOC_CHART_BOOK_BYTES. */
  maxChartBookBytes?: number;
  /** Bytes of embedded workbooks that are opened, all together. Default MAX_DOC_CHART_BOOKS_BYTES. */
  maxChartBooksBytes?: number;
  /** Embedded workbooks that are opened; 0 opens none. Default MAX_DOC_CHART_BOOKS. */
  maxChartBooks?: number;
}

const limitOf = (value: number | undefined, fallback: number): number => (typeof value === 'number' && value >= 0 ? Math.min(value, fallback) : fallback);

/** A reference as the model holds it: without the "=" a part may write before it. */
const refKey = (reference: string): string => {
  const text = reference.trim();
  return text.charAt(0) === '=' ? text.slice(1) : text;
};

/**
 * The references of a chart part that stand without a cache element (<c:numCache>, <c:strCache>,
 * <c:multiLvlStrCache>) beside their formula. A reference whose cache element is there and holds
 * no point is not among them: the part says that there is nothing, and that is what is drawn.
 */
function uncachedRefs(root: XmlElement): Set<string> {
  const out = new Set<string>();
  const stack: XmlElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop() as XmlElement;
    const local = el.name.replace(/^.*:/, '');
    let formula: string | undefined;
    let cache = false;
    for (const node of el.children) {
      if (node.type !== 'element') continue;
      stack.push(node);
      const name = node.name.replace(/^.*:/, '');
      if (name === 'f') {
        formula = '';
        for (const text of node.children) if (text.type === 'text') formula += text.value;
      } else if (name === 'numCache' || name === 'strCache' || name === 'multiLvlStrCache') cache = true;
    }
    if ((local === 'numRef' || local === 'strRef' || local === 'multiLvlStrRef') && formula !== undefined && !cache) out.add(refKey(formula));
  }
  return out;
}

/** True for data that holds no cached point and no count of points. */
const empty = (data: ChartData | undefined): boolean => !data || !data.cache || !Array.isArray(data.cache.index) || (data.cache.index.length === 0 && !(data.count > 0));

/** One thing of a chart the part caches nothing for: the data of a series, or a text (a name, a title) read from a cell. */
interface Want {
  kind: Wanted;
  data?: ChartData;
  text?: { ref?: string; text?: string };
}

/** What of a chart has a reference without a cache element, in the order the series stand in. */
function wantsOf(chart: Chart, uncached: Set<string>): Want[] {
  const out: Want[] = [];
  const named = (reference: string | undefined): boolean => reference !== undefined && uncached.has(refKey(reference));
  for (const group of chart.groups) {
    for (const series of group.series) {
      const sources: [Wanted, ChartData | undefined][] = [['values of series', series.values], ['categories', series.categories], ['bubble sizes', series.bubbleSizes]];
      for (const [kind, data] of sources) {
        // UNSURE: categories of several levels are left to the cache: which column of the block is which level is not certain.
        if (data && !data.multiLevel && empty(data) && named(data.ref)) out.push({ kind, data });
      }
      if (series.name && series.name.text === undefined && named(series.name.ref)) out.push({ kind: 'names of series', text: series.name });
    }
  }
  for (const title of [chart.title].concat(chart.axes.map((axis) => axis.title))) {
    if (title && title.text === undefined && named(title.ref)) out.push({ kind: 'titles', text: title });
  }
  return out;
}

/** The kinds of the things listed, each once, in the order the notes name them. */
const kindsOf = (wants: Want[]): Wanted[] => WANTED.filter((kind) => wants.some((want) => want.kind === kind));

/** Why a workbook was not read: said in the note of the chart. */
class BookError extends Error {}

/** The embedded workbooks of one document that were asked for, and what opening more of them may still cost. */
interface Books {
  pkg: OpcPackage;
  /** By the name of the part in lower case: the workbook, or why there is none. Each part is opened once. */
  opened: Map<string, Promise<Workbook | BookError>>;
  left: number;
  bytesLeft: number;
  maxBytes: number;
}

/** The size the package states for a part, uncompressed; undefined when it cannot be told before the part is read. */
function statedSize(pkg: OpcPackage, part: string): number | undefined {
  const zip = pkg.zip;
  if (!zip || typeof zip.names !== 'function') return undefined;
  const lower = part.slice(1).toLowerCase();
  const name = zip.has(part.slice(1)) ? part.slice(1) : zip.names().find((candidate) => candidate.toLowerCase() === lower);
  const entry = name === undefined ? undefined : zip.entry(name);
  return entry ? entry.uncompressedSize : undefined;
}

/**
 * Opens an embedded workbook and reads its cells: a package (xlsx, and xlsm, of which only the
 * cells are read - nothing of a workbook is ever run) or a compound file with a BIFF workbook
 * (xls). Throws a BookError for whatever is not read: too large, a binary workbook (xlsb),
 * encrypted, no workbook.
 */
async function openBook(books: Books, part: string): Promise<Workbook> {
  const { pkg } = books;
  const stated = statedSize(pkg, part);
  // The compressed part is at most as large as what it holds: one beyond the limit is not even inflated.
  if (stated !== undefined && stated > books.maxBytes) throw new BookError('the embedded workbook of the chart is too large');
  if (books.left <= 0) throw new BookError('the document holds too many embedded workbooks');
  books.left--;
  const bytes = await pkg.readPart(part);
  if (bytes.length > books.maxBytes) throw new BookError('the embedded workbook of the chart is too large');
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    const inner = await OpcPackage.open(bytes);
    // What reading the workbook inflates: every entry of it, as large as the archive says it is (the archive holds each to that).
    let size = 0;
    for (const name of inner.zip.names()) {
      const entry = inner.zip.entry(name);
      size += entry ? entry.uncompressedSize : 0;
    }
    if (size > books.maxBytes) throw new BookError('the embedded workbook of the chart is too large');
    if (size > books.bytesLeft) throw new BookError('the embedded workbooks of the document hold too many bytes');
    books.bytesLeft -= size;
    const main = await inner.mainPart();
    const type = (main && inner.contentTypeOf(main)) || '';
    if (/sheet\.binary/i.test(type) || /\.bin$/i.test(main || '')) throw new BookError('the embedded workbook of the chart is a binary workbook (xlsb), which is not read');
    if (!/spreadsheetml|ms-excel/i.test(type)) throw new BookError('what stands for the embedded workbook of the chart is not a workbook');
    return readXlsx(inner);
  }
  if (CfbFile.isCfb(bytes)) {
    if (bytes.length > books.bytesLeft) throw new BookError('the embedded workbooks of the document hold too many bytes');
    books.bytesLeft -= bytes.length;
    const cfb = CfbFile.read(bytes);
    if (cfb.has('EncryptedPackage')) throw new BookError('the embedded workbook of the chart is encrypted');
    if (!cfb.has('Workbook') && !cfb.has('Book')) throw new BookError('what stands for the embedded workbook of the chart is not a workbook');
    return readXls(cfb);
  }
  throw new BookError('what stands for the embedded workbook of the chart is not a workbook');
}

/** The workbook the chart part names as its data (<c:externalData r:id>), read once for all charts that name it. */
async function bookOf(books: Books, part: string, root: XmlElement): Promise<Workbook | BookError> {
  let id: string | undefined;
  for (const node of root.children) {
    if (node.type !== 'element' || node.name.replace(/^.*:/, '') !== 'externalData') continue;
    node.attrs.forEach((value, name) => {
      if (name.replace(/^.*:/, '') === 'id') id = value;
    });
    break;
  }
  if (id === undefined || typeof books.pkg.relationships !== 'function') return new BookError('the chart names no embedded workbook');
  let target: string | undefined;
  try {
    const rel = (await books.pkg.relationships(part)).find((candidate) => candidate.id === id);
    if (rel && rel.external) return new BookError('the workbook of the chart is linked (no embedded workbook): it is not fetched');
    if (rel && rel.part && books.pkg.hasPart(rel.part)) target = rel.part;
  } catch (error) {
    // Relationships that cannot be read name nothing.
  }
  if (target === undefined) return new BookError('the chart names no embedded workbook');
  const key = target.toLowerCase();
  let opened = books.opened.get(key);
  if (!opened) {
    opened = openBook(books, target).catch((error) =>
      error instanceof BookError ? error : new BookError(`the embedded workbook of the chart could not be read (${error instanceof Error ? error.message : String(error)})`),
    );
    books.opened.set(key, opened);
  }
  return opened;
}

/**
 * Gives what a chart lacks a cache for the cells of the workbook: the data of series as their
 * cache, names and titles as their text. The values taken count against `budget` like cached
 * ones. Says what was read, what stayed empty, and why the last of those did.
 */
function fillFromBook(chart: Chart, book: Workbook, wants: Want[], budget: { points: number }): { read: Want[]; missing: Want[]; problem?: string } {
  const cells = chartCellsOf(book, false);
  const read: Want[] = [];
  const missing: Want[] = [];
  let problem: string | undefined;
  let numbers = false;
  for (const want of wants) {
    const { data, text } = want;
    if (data) {
      const found = cells(data.ref as string, data.kind);
      if (!found) {
        problem = 'its reference cannot be followed in the embedded workbook of the chart';
        missing.push(want);
        continue;
      }
      if (found.cache.index.length > budget.points) {
        problem = 'the charts of the document hold too many values; the embedded workbook of the chart is not read for it';
        missing.push(want);
        continue;
      }
      budget.points -= found.cache.index.length;
      data.cache = found.cache;
      data.count = found.count;
      if (data.kind === 'num' && data.formatCode === undefined && found.formatCode !== undefined) data.formatCode = found.formatCode;
      if (data.kind === 'num') numbers = true;
      read.push(want);
    } else if (text) {
      const found = cells(text.ref as string, 'str');
      const first = found && found.cache.index[0] === 0 ? found.cache.values[0] : undefined;
      if (typeof first === 'string') {
        text.text = first;
        read.push(want);
      } else if (!found) {
        problem = 'its reference cannot be followed in the embedded workbook of the chart';
        missing.push(want);
      }
      // A cell that is there and holds nothing is an empty name: nothing is missing.
    }
  }
  // The numbers of dates are those of the workbook they were read from.
  if (numbers && book.date1904 === true) chart.date1904 = true;
  return { read, missing, problem };
}

/**
 * Reads the chart parts (names of parts as the package spells them, each
 * once, in the order their frames stand in the document) into models, by
 * part name.
 * `themeColors` as parseChart() takes them. Undefined when no part gave a
 * chart. Never throws: what goes wrong is added to `warnings`.
 */
export async function readDocCharts(pkg: OpcPackage, parts: Iterable<string>, themeColors: string[], warnings: string[], limits: DocChartLimits = {}): Promise<Map<string, Chart> | undefined> {
  const maxCharts = limitOf(limits.maxCharts, MAX_DOC_CHARTS);
  const maxBytes = limitOf(limits.maxChartPartBytes, MAX_CHART_PART_BYTES);
  const budget = { points: limitOf(limits.maxChartPoints, MAX_DOC_CHART_POINTS) };
  let bytesLeft = limitOf(limits.maxChartBytes, MAX_DOC_CHART_BYTES);
  const books: Books = {
    pkg,
    opened: new Map(),
    left: limitOf(limits.maxChartBooks, MAX_DOC_CHART_BOOKS),
    bytesLeft: limitOf(limits.maxChartBooksBytes, MAX_DOC_CHART_BOOKS_BYTES),
    maxBytes: limitOf(limits.maxChartBookBytes, MAX_DOC_CHART_BOOK_BYTES),
  };
  const charts = new Map<string, Chart>();
  let read = 0;
  let beyond = 0;
  let said = 0;
  let unsaid = 0;
  const warn = (text: string) => {
    if (said < MAX_DOC_CHART_WARNINGS) {
      warnings.push(text);
      said++;
    } else unsaid++;
  };
  for (const part of parts) {
    if (read >= maxCharts) {
      beyond++;
      continue;
    }
    read++;
    try {
      const bytes = await pkg.readPart(part);
      if (bytes.length > maxBytes) throw new Error('the part is too large');
      if (bytes.length > bytesLeft) {
        warn(`Chart ${quote(part)} is not shown: the charts of the document hold too many bytes`);
        continue;
      }
      bytesLeft -= bytes.length;
      const root = parseXml(bytes).root;
      const chart = parseChart(root, themeColors, budget);
      if (!chart) {
        warn(`Chart ${quote(part)} holds nothing that is drawn (no series, or a chart type that is not drawn); a placeholder is shown`);
        continue;
      }
      const wants = wantsOf(chart, uncachedRefs(root));
      if (wants.length > 0) {
        // Only now, and only for this chart, the workbook beside the part is looked at.
        let result: { read: Want[]; missing: Want[]; problem?: string };
        try {
          const book = await bookOf(books, part, root);
          result = book instanceof BookError ? { read: [], missing: wants, problem: book.message } : fillFromBook(chart, book, wants, budget);
        } catch (error) {
          result = { read: [], missing: wants, problem: 'the embedded workbook of the chart could not be read' };
        }
        // What came from where: each part of the chart is named in the note that is true of it.
        if (result.read.length > 0) chart.notes.push(BOOK_NOTE + kindsOf(result.read).join(', '));
        if (result.missing.length > 0) chart.notes.push(`${EMPTY_NOTE}${kindsOf(result.missing).join(', ')} (${result.problem || 'the embedded workbook of the chart could not be read'})`);
      }
      // Values that name no cells and hold nothing have no workbook to be read from.
      if (chart.groups.some((group) => group.series.some((series) => empty(series.values) && (!series.values || series.values.ref === undefined)))) {
        chart.notes.push(`${EMPTY_NOTE}values of series (they name no cells)`);
      }
      charts.set(part, chart);
    } catch (error) {
      // A chart part that cannot be read (not XML, a DTD, too much of it) leaves its frame a placeholder.
      if (error instanceof ChartBudgetError) warn(`Chart ${quote(part)} is not shown: the charts of the document hold too many cached values`);
      else warn(`Chart ${quote(part)} could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  if (unsaid > 0) warnings.push(`More charts that are not shown: ${unsaid}`);
  if (beyond > 0) warnings.push(`Charts beyond the limit of a document (${maxCharts}) are not shown: ${beyond}`);
  return charts.size > 0 ? charts : undefined;
}
