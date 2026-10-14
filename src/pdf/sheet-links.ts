// The links and the outline of the PDF of a workbook, from the pages
// printing lays the sheets out on:
//
// - a hyperlink of a cell (Sheet.hyperlinks) lies over the cells of its range
//   wherever they are printed - on the page that holds them, and on every
//   page that repeats them as print titles - and over the whole of a merged
//   cell it starts;
// - a link to a place in the workbook (`location`: "Sheet2!B5",
//   "'My sheet'!A1:C3", "B5", or a defined name that stands for one of those)
//   goes to the first page that prints that cell, at the top left corner of
//   the cell. A sheet that is not in the file, or a cell that is printed on
//   no page, leaves the link out, counted;
// - the outline has an entry for every sheet that is printed, with its name,
//   at its first page - when more than one sheet is printed.
//
// What an address outside the file may be is decided when it is written
// (pdfLinkUri): javascript:, file: and relative addresses are left out there,
// counted.

import type { SheetMetrics } from '../layout/sheet-metrics';
import type { PrintBand, SheetPrintLayout, SheetPrintPage } from '../layout/sheet-pages';
import { parseRange, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { MAX_FILE_LINKS, MAX_PAGE_LINKS, type PdfLink, type PdfLinkRect, type PdfOutlineEntry } from './links';

/** What a page of the plan is: a page of a sheet, laid out (SheetJobPage of the print plan). */
export interface SheetLinkPage {
  sheet: number;
  layout: SheetPrintLayout;
  page: SheetPrintPage;
  /** A page that is a chart sheet has no cells. */
  chart?: unknown;
}

export interface SheetPdfLinkOptions {
  /** Points per px of the paper: 0.75 for paper laid out at 96 px per inch. */
  scale: number;
  /** The row and column sizes the pages were laid out with. */
  metrics(sheet: Sheet): SheetMetrics;
  links?: boolean;
  outline?: boolean;
}

export interface SheetPdfLinks {
  /** The links of every page, by its index: points from the top left corner of the page. */
  pages: PdfLink[][];
  outline: PdfOutlineEntry[];
  /** Links to a place in the workbook that is not printed, or that is no place. */
  unresolved: number;
  /** What was left out beyond the limits. */
  dropped: number;
}

/** Hyperlinks of one sheet that are looked at. */
const MAX_SHEET_LINKS = 20000;
/** Merged cells of one sheet that are looked at. */
const MAX_SHEET_MERGES = 200000;
/** Pairs of a page and a link that are looked at, all sheets together. */
const MAX_LINK_CHECKS = 20000000;
/** Characters of a location. */
const MAX_LOCATION = 1024;
/** Letters beyond ASCII, as a range of a class: written with escapes, so that the file reads the same in every encoding. */
const WIDE = '\\u00a0-\\ufffd';
/** What a defined name can be: a letter, "_" or "\" first, then letters, digits, "_", "." and "\". */
const NAME = new RegExp('^[A-Za-z_\\\\' + WIDE + '][\\w.\\\\' + WIDE + ']*$');

/** The sheet and the cells a location names; undefined for one that is not a place. */
export function parseLocation(location: string, workbook: Workbook, from: number, depth = 0): { sheet: number; range: CellRange } | undefined {
  if (typeof location !== 'string' || location.length === 0 || location.length > MAX_LOCATION) return undefined;
  let text = location.trim();
  if (text.charAt(0) === '#') text = text.slice(1);
  let sheetName: string | undefined;
  let ref = text;
  if (text.charAt(0) === "'") {
    let name = '';
    let i = 1;
    for (; i < text.length; i++) {
      const char = text.charAt(i);
      if (char === "'") {
        if (text.charAt(i + 1) === "'") {
          name += "'";
          i++;
          continue;
        }
        break;
      }
      name += char;
    }
    if (i >= text.length || text.charAt(i + 1) !== '!') return undefined;
    sheetName = name;
    ref = text.slice(i + 2);
  } else {
    const bang = text.lastIndexOf('!');
    if (bang >= 0) {
      sheetName = text.slice(0, bang);
      ref = text.slice(bang + 1);
    }
  }
  let sheet = from;
  if (sheetName !== undefined) {
    const wanted = sheetName.toLowerCase();
    sheet = -1;
    for (let i = 0; i < workbook.sheets.length; i++) {
      if (workbook.sheets[i].name.toLowerCase() === wanted) {
        sheet = i;
        break;
      }
    }
    if (sheet < 0) return undefined;
  }
  const range = parseRange(ref.replace(/\$/g, ''));
  if (range) return { sheet, range };
  // A defined name: of the sheet first, then of the workbook; one step only.
  if (sheetName !== undefined || depth > 0 || !NAME.test(ref)) return undefined;
  const names = Array.isArray(workbook.definedNames) ? workbook.definedNames : [];
  const lowerName = ref.toLowerCase();
  let found: string | undefined;
  for (let pass = 0; pass < 2 && found === undefined; pass++) {
    for (let i = 0; i < names.length; i++) {
      const entry = names[i];
      if (!entry || typeof entry.name !== 'string' || entry.name.toLowerCase() !== lowerName) continue;
      if (pass === 0 ? entry.localSheet === from : entry.localSheet === undefined) {
        found = entry.value;
        break;
      }
    }
  }
  return found === undefined ? undefined : parseLocation(found.replace(/^=/, ''), workbook, from, depth + 1);
}

/** The bands of rows or columns of a page as they are printed: where each starts, in px of the sheet from the top left corner of the printed block. */
function placedBands(title: PrintBand | undefined, own: PrintBand, axis: SheetMetrics['rows'], start: number): { first: number; last: number; pos: number }[] {
  const out: { first: number; last: number; pos: number }[] = [];
  let pos = start;
  const bands = title ? [title, own] : [own];
  for (let i = 0; i < bands.length && out.length < 2; i++) {
    const first = Math.max(Math.floor(bands[i].first), 0);
    const last = Math.min(Math.floor(bands[i].last), axis.count - 1);
    if (!(first <= last)) continue;
    const size = axis.offset(last + 1) - axis.offset(first);
    if (!(size > 0)) continue;
    out.push({ first, last, pos });
    pos += size;
  }
  return out;
}

/** The rectangles (paper px) of a range on a page: one per pair of bands it is in, cut to the printed block. */
function rangeRects(range: CellRange, page: SheetPrintPage, layout: SheetPrintLayout, metrics: SheetMetrics): { x: number; y: number; width: number; height: number }[] {
  const s = layout.scale > 0 && isFinite(layout.scale) ? layout.scale : 1;
  const cols = placedBands(page.titleCols, page.cols, metrics.cols, layout.headings && layout.headingWidth > 0 ? layout.headingWidth : 0);
  const rows = placedBands(page.titleRows, page.rows, metrics.rows, layout.headings && layout.headingHeight > 0 ? layout.headingHeight : 0);
  const out: { x: number; y: number; width: number; height: number }[] = [];
  for (let r = 0; r < rows.length; r++) {
    const top = Math.max(range.top, rows[r].first);
    const bottom = Math.min(range.bottom, rows[r].last);
    if (top > bottom) continue;
    const y1 = rows[r].pos + metrics.rows.offset(top) - metrics.rows.offset(rows[r].first);
    const y2 = rows[r].pos + metrics.rows.offset(bottom + 1) - metrics.rows.offset(rows[r].first);
    for (let c = 0; c < cols.length; c++) {
      const left = Math.max(range.left, cols[c].first);
      const right = Math.min(range.right, cols[c].last);
      if (left > right) continue;
      const x1 = cols[c].pos + metrics.cols.offset(left) - metrics.cols.offset(cols[c].first);
      const x2 = cols[c].pos + metrics.cols.offset(right + 1) - metrics.cols.offset(cols[c].first);
      // On the paper, cut where printing cuts the block: at its size.
      const px1 = Math.max(page.x, page.x + x1 * s);
      const px2 = Math.min(page.x + page.width, page.x + x2 * s);
      const py1 = Math.max(page.y, page.y + y1 * s);
      const py2 = Math.min(page.y + page.height, page.y + y2 * s);
      if (px2 - px1 > 0.01 && py2 - py1 > 0.01) out.push({ x: px1, y: py1, width: px2 - px1, height: py2 - py1 });
    }
  }
  return out;
}

/** The range of a link with the merged cells it starts made whole. */
function withMerges(range: CellRange, merges: Map<string, CellRange>): CellRange {
  let out = range;
  if (merges.size === 0) return out;
  // Only the corners matter for a link of one cell, which is what links nearly always are.
  const merged = merges.get(range.top + ',' + range.left);
  if (merged) out = { top: Math.min(out.top, merged.top), left: Math.min(out.left, merged.left), bottom: Math.max(out.bottom, merged.bottom), right: Math.max(out.right, merged.right) };
  return out;
}

/**
 * The links of every page of the print plan of a workbook and the outline of
 * its sheets, in points (see SheetPdfLinks).
 */
export function sheetPdfLinks(workbook: Workbook, plan: readonly SheetLinkPage[], options: SheetPdfLinkOptions): SheetPdfLinks {
  const scale = options.scale > 0 && isFinite(options.scale) ? options.scale : 0.75;
  const out: SheetPdfLinks = { pages: [], outline: [], unresolved: 0, dropped: 0 };
  for (let i = 0; i < plan.length; i++) out.pages.push([]);
  const firstPage = new Map<number, number>();
  for (let i = 0; i < plan.length; i++) if (plan[i] && !firstPage.has(plan[i].sheet)) firstPage.set(plan[i].sheet, i);
  // One sheet needs no outline: a panel with one entry beside its pages tells nothing.
  if (options.outline !== false && firstPage.size > 1) {
    firstPage.forEach((page, sheet) => {
      const entry = workbook.sheets[sheet];
      if (entry) out.outline.push({ title: entry.name, level: 0, dest: { page, x: 0, y: 0 } });
    });
    out.outline.sort((a, b) => a.dest.page - b.dest.page);
  }
  if (options.links === false) return out;

  const metricsCache = new Map<Sheet, SheetMetrics>();
  const metricsOf = (sheet: Sheet): SheetMetrics => {
    let metrics = metricsCache.get(sheet);
    if (!metrics) {
      metrics = options.metrics(sheet);
      metricsCache.set(sheet, metrics);
    }
    return metrics;
  };
  const mergesCache = new Map<Sheet, Map<string, CellRange>>();
  const mergesOf = (sheet: Sheet): Map<string, CellRange> => {
    let map = mergesCache.get(sheet);
    if (!map) {
      map = new Map();
      const merges = Array.isArray(sheet.merges) ? sheet.merges : [];
      for (let i = 0; i < merges.length && i < MAX_SHEET_MERGES; i++) if (merges[i]) map.set(merges[i].top + ',' + merges[i].left, merges[i]);
      mergesCache.set(sheet, map);
    }
    return map;
  };
  const indexes = new Map<Sheet, LinkIndex>();
  const indexOf = (sheet: Sheet, at: number): LinkIndex => {
    let index = indexes.get(sheet);
    if (!index) {
      index = linkIndex(sheet, mergesOf(sheet), plan.filter((one) => one && one.sheet === at && !one.chart).map((one) => one.page));
      indexes.set(sheet, index);
    }
    return index;
  };
  /** Where a cell is printed first: the page, and its top left corner in paper px. */
  const placeOf = (sheet: number, row: number, col: number): { page: number; x: number; y: number } | undefined => {
    const target = workbook.sheets[sheet];
    if (!target) return undefined;
    for (let i = firstPage.has(sheet) ? (firstPage.get(sheet) as number) : plan.length; i < plan.length; i++) {
      const entry = plan[i];
      if (entry.sheet !== sheet || entry.chart) continue;
      const page = entry.page;
      if (row < page.rows.first || row > page.rows.last || col < page.cols.first || col > page.cols.last) continue;
      // The cells of the page itself come last, after those of the print titles.
      const rects = rangeRects({ top: row, left: col, bottom: row, right: col }, page, entry.layout, metricsOf(target));
      const own = rects.length > 0 ? rects[rects.length - 1] : undefined;
      if (own) return { page: i, x: own.x, y: own.y };
      // A hidden cell takes no room on the page: it is shown at the place it would be.
      return { page: i, x: page.x, y: page.y };
    }
    return undefined;
  };
  const resolved = new Map<string, { page: number; x: number; y: number } | null>();

  // Each link is counted once, however many pages it was left out of (a print title repeats it). The limits of a
  // page and of the file are those of the PDF writer (links.ts), kept here so that no more links are made than it takes.
  const droppedLinks = new Set<string>();
  const unresolvedLinks = new Set<string>();
  const writtenLinks = new Set<string>();
  let written = 0;

  let checks = MAX_LINK_CHECKS;
  let i = 0;
  for (; i < plan.length && written < MAX_FILE_LINKS; i++) {
    const entry = plan[i];
    if (!entry || entry.chart) continue;
    const sheet = workbook.sheets[entry.sheet];
    if (!sheet || !Array.isArray(sheet.hyperlinks) || sheet.hyperlinks.length === 0) continue;
    const metrics = metricsOf(sheet);
    const index = indexOf(sheet, entry.sheet);
    const list = out.pages[i];
    const page = entry.page;
    const links = sheet.hyperlinks;
    // Counted once, with the first page of the sheet.
    if (firstPage.get(entry.sheet) === i && links.length > index.ranges.length) out.dropped += links.length - index.ranges.length;
    const candidates = index.linksIn(page);
    for (let c = 0; c < candidates.length; c++) {
      const k = candidates[c];
      const id = entry.sheet + ':' + k;
      if (list.length >= MAX_PAGE_LINKS || written >= MAX_FILE_LINKS || --checks < 0) {
        // The page (or the file) is full: what is left of the page is left out.
        for (let rest = c; rest < candidates.length; rest++) droppedLinks.add(entry.sheet + ':' + candidates[rest]);
        break;
      }
      const link = links[k];
      const range = index.ranges[k];
      if (!link || !range) continue;
      const rects = rangeRects(range, page, entry.layout, metrics);
      if (rects.length === 0) continue;
      const scaled: PdfLinkRect[] = rects.map((rect) => ({ x: rect.x * scale, y: rect.y * scale, width: rect.width * scale, height: rect.height * scale }));
      const target = typeof link.target === 'string' && link.target !== '' ? link.target : undefined;
      const location = typeof link.location === 'string' && link.location !== '' ? link.location : undefined;
      let made: PdfLink | undefined;
      if (target !== undefined) {
        // The place in the file the address goes to is part of the address, as Excel opens it.
        made = { rects: scaled, uri: location !== undefined && target.indexOf('#') < 0 ? target + '#' + location : target };
      } else if (location === undefined) {
        // A link that had an address the file no longer holds (a relationship without a target).
        if (link.rId !== undefined) unresolvedLinks.add(id);
      } else {
        const key = entry.sheet + '!' + location;
        let place = resolved.get(key);
        if (place === undefined) {
          const parsed = parseLocation(location, workbook, entry.sheet);
          place = (parsed && placeOf(parsed.sheet, parsed.range.top, parsed.range.left)) || null;
          resolved.set(key, place);
        }
        if (place === null) unresolvedLinks.add(id);
        else made = { rects: scaled, dest: { page: place.page, x: place.x * scale, y: place.y * scale } };
      }
      if (!made) continue;
      list.push(made);
      written++;
      writtenLinks.add(id);
    }
  }
  if (i < plan.length) {
    // The file is full: every link of the pages not looked at that was written nowhere is left out, once. The links of
    // a sheet that reach into its printed cells stand for those of its pages (a print title repeats links of other pages).
    const seen = new Set<number>();
    for (let p = i; p < plan.length; p++) {
      const entry = plan[p];
      if (!entry || entry.chart || seen.has(entry.sheet)) continue;
      seen.add(entry.sheet);
      const sheet = workbook.sheets[entry.sheet];
      if (!sheet || !Array.isArray(sheet.hyperlinks)) continue;
      const printed = indexOf(sheet, entry.sheet).printed;
      for (let k = 0; k < printed.length; k++) {
        const id = entry.sheet + ':' + printed[k];
        if (!writtenLinks.has(id) && !unresolvedLinks.has(id)) droppedLinks.add(id);
      }
    }
  }
  out.dropped += droppedLinks.size;
  out.unresolved += unresolvedLinks.size;
  return out;
}

/** Links whose range is taller than this are not sorted in: they are looked at on every page of their sheet. */
const TALL_LINK = 256;

/** The links of a sheet by the cells they cover: those of a page are found without looking at the others. */
interface LinkIndex {
  /** The range of every link (merged cells it starts made whole), by its index; at most MAX_SHEET_LINKS. */
  ranges: (CellRange | undefined)[];
  /** The indexes of the links that reach into the cells some page of the sheet prints. */
  printed: number[];
  /** The indexes of the links that cover a cell of the page (its own and its print titles), in the order of the sheet. */
  linksIn(page: SheetPrintPage): number[];
}

const meets = (range: CellRange, rows: PrintBand | undefined, cols: PrintBand | undefined): boolean =>
  rows !== undefined && cols !== undefined && range.top <= rows.last && range.bottom >= rows.first && range.left <= cols.last && range.right >= cols.first;

/** Whether the range is on the page: in one of the blocks the page prints (its cells, the print titles and their corner). */
function onPage(range: CellRange, page: SheetPrintPage): boolean {
  return meets(range, page.rows, page.cols) || meets(range, page.titleRows, page.cols) || meets(range, page.rows, page.titleCols) || meets(range, page.titleRows, page.titleCols);
}

function linkIndex(sheet: Sheet, merges: Map<string, CellRange>, pages: readonly SheetPrintPage[]): LinkIndex {
  const links = sheet.hyperlinks;
  const count = Math.min(links.length, MAX_SHEET_LINKS);
  // What the pages of the sheet print, all together: a link outside it is on no page and is not looked at again.
  let top = Infinity;
  let bottom = -Infinity;
  let left = Infinity;
  let right = -Infinity;
  for (let p = 0; p < pages.length; p++) {
    const page = pages[p];
    const rows = page.titleRows ? [page.rows, page.titleRows] : [page.rows];
    const cols = page.titleCols ? [page.cols, page.titleCols] : [page.cols];
    for (let r = 0; r < rows.length; r++) {
      top = Math.min(top, rows[r].first);
      bottom = Math.max(bottom, rows[r].last);
    }
    for (let c = 0; c < cols.length; c++) {
      left = Math.min(left, cols[c].first);
      right = Math.max(right, cols[c].last);
    }
  }
  const area = { first: top, last: bottom };
  const across = { first: left, last: right };
  const ranges: (CellRange | undefined)[] = [];
  const printed: number[] = [];
  const short: number[] = [];
  const tall: number[] = [];
  for (let k = 0; k < count; k++) {
    const link = links[k];
    const range = link && link.range ? withMerges(link.range, merges) : undefined;
    ranges.push(range);
    if (!range || !meets(range, area, across)) continue;
    printed.push(k);
    if (range.bottom - range.top > TALL_LINK) tall.push(k);
    else short.push(k);
  }
  // Sorted by the first row, once: a band of rows is a stretch of this list.
  short.sort((a, b) => (ranges[a] as CellRange).top - (ranges[b] as CellRange).top || a - b);
  const firstAtOrAfter = (row: number): number => {
    let low = 0;
    let high = short.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if ((ranges[short[middle]] as CellRange).top < row) low = middle + 1;
      else high = middle;
    }
    return low;
  };
  return {
    ranges,
    printed,
    linksIn(page) {
      const found = new Set<number>();
      const bands = page.titleRows ? [page.titleRows, page.rows] : [page.rows];
      for (let b = 0; b < bands.length; b++) {
        const band = bands[b];
        // A short link that reaches into the band starts at most TALL_LINK rows above it.
        for (let s = firstAtOrAfter(band.first - TALL_LINK); s < short.length; s++) {
          const range = ranges[short[s]] as CellRange;
          if (range.top > band.last) break;
          if (range.bottom >= band.first && onPage(range, page)) found.add(short[s]);
        }
      }
      for (let t = 0; t < tall.length; t++) if (onPage(ranges[tall[t]] as CellRange, page)) found.add(tall[t]);
      return Array.from(found).sort((a, b) => a - b);
    },
  };
}
