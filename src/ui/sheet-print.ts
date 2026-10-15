// Printing a workbook: which pages a print job has (planSheetPages) and how
// one of them is painted on a canvas (paintSheetPage). The cells are drawn by
// the grid's own code (renderGridPrint); the pages come from the sheet's page
// setup (layout/sheet-pages) and the job itself is run by the view with the
// parts documents are printed with (doc-print).
//
// Nothing here touches the DOM at import time, and the plan needs none at all.

import { formatValue } from '../format/number-format';
import type { ConditionalSource } from '../format/conditional';
import type { FormatLocale } from '../format/locale';
import { headerFooterLines, parseHeaderFooter, type HeaderFooterFormat, type HeaderFooterPiece, type HeaderFooterRun, type HeaderFooterValues } from '../layout/header-footer';
import { createSheetMetrics, type SheetMetrics } from '../layout/sheet-metrics';
import { MAX_SHEET_PAGES, PAPER_A4, PAPER_LETTER, SheetPagesError, marginsOf, paginateSheet, paperSize, printRanges, type SheetPrintLayout, type SheetPrintPage } from '../layout/sheet-pages';
import type { Chart } from '../model/sheet/chart';
import { resolveCellStyle } from '../model/sheet/color';
import type { Font, HeaderFooter, Sheet, Workbook } from '../model/sheet/types';
import { ChartRenderer } from '../render/chart';
import { renderGridPrint, type GridBand, type GridDrawingSource } from '../render/grid';
import { cssFont, fontPixels } from '../render/text';
import { DocumentPrintError, type DocumentPrintOptions } from './doc-print';

export interface SheetPrintOptions extends DocumentPrintOptions {
  /** 'active' (the default) prints the sheet that is shown; 'all' every visible worksheet and chart sheet, in the order of the tabs. */
  sheets?: 'active' | 'all';
  /** What "&F" in a header or footer stands for; WorkbookViewOptions.fileName when it is not given (empty without one). */
  fileName?: string;
}

/** One page of a print job. */
export interface SheetJobPage {
  /** Index of the sheet in the workbook. */
  sheet: number;
  layout: SheetPrintLayout;
  page: SheetPrintPage;
  /** Its position among the pages of its sheet, from 0. */
  index: number;
  /** The number "&P" shows. */
  number: number;
  /** The page of a chart sheet: the chart that fills what the margins leave of the paper. */
  chart?: Chart;
}

export interface SheetJobPlan {
  pages: SheetJobPage[];
  /** The sheets that are printed, in order; the ones with nothing to print are not among them. */
  sheets: number[];
}

export interface SheetPlanOptions {
  /** Paper size code of sheets that state none. */
  defaultPaper?: number;
  /** Pages of all sheets together; more are refused. Default MAX_SHEET_PAGES. */
  maxPages?: number;
  /** Row and column sizes of a sheet; made from the sheet when missing. */
  metrics?(sheet: Sheet, index: number): SheetMetrics;
  /** The boxes of the drawings of a sheet, in px from the corner of A1. */
  drawings?(sheet: Sheet, index: number, metrics: SheetMetrics): readonly { x: number; y: number; width: number; height: number }[] | undefined;
}

/**
 * The paper of a sheet that states none: US Letter where the language of the
 * view is that of the United States or Canada ("en", "en-US", "fr-CA"), A4
 * everywhere else. Excel takes the paper of the default printer, which a
 * page cannot ask for.
 */
export function defaultPaperFor(tag: string | undefined): number {
  return typeof tag === 'string' && (/^en$/i.test(tag) || /[-_](US|CA)$/i.test(tag)) ? PAPER_LETTER : PAPER_A4;
}

/** The chart a chart sheet shows: the first one of its drawing that is drawn. */
function chartOfSheet(sheet: Sheet): Chart | undefined {
  if (sheet.kind !== 'chartsheet' || !Array.isArray(sheet.drawings)) return undefined;
  for (const drawing of sheet.drawings) if (drawing && drawing.object && drawing.object.type === 'chart' && drawing.object.chart) return drawing.object.chart;
  return undefined;
}

/**
 * The one page of a chart sheet: the chart as large as the margins leave the paper.
 * UNSURE: a chart sheet lies on its side unless its page setup says "portrait" - Excel makes
 * a new chart sheet that way, and "default" leaves the choice to the application.
 */
function chartSheetLayout(sheet: Sheet, defaultPaper: number | undefined): SheetPrintLayout {
  const setup = sheet.pageSetup;
  const stated = setup ? setup.page : undefined;
  const paper = paperSize({ ...(stated || {}), orientation: stated && stated.orientation === 'portrait' ? 'portrait' : 'landscape' }, defaultPaper);
  const margins = marginsOf(setup ? setup.margins : undefined, paper);
  const page: SheetPrintPage = {
    area: 0,
    // No cells: bands that hold nothing.
    rows: { first: 0, last: -1 },
    cols: { first: 0, last: -1 },
    x: margins.left,
    y: margins.top,
    width: paper.width - margins.left - margins.right,
    height: paper.height - margins.top - margins.bottom,
  };
  const layout: SheetPrintLayout = { paper: { width: paper.width, height: paper.height }, landscape: paper.landscape, margins, scale: 1, headings: false, headingWidth: 0, headingHeight: 0, gridLines: false, pages: [page] };
  if (stated && stated.useFirstPageNumber === true && typeof stated.firstPageNumber === 'number' && Number.isFinite(stated.firstPageNumber) && Math.abs(stated.firstPageNumber) <= 32767) layout.firstPageNumber = Math.floor(stated.firstPageNumber);
  return layout;
}

/**
 * What "&F" in a header or footer stands for in a job: the name the call gives (an empty one too),
 * else the name the view was given. Printing, the print preview and the PDF all go by it, so that
 * what is printed is what is previewed.
 */
export function jobFileName(given: string | undefined, ofView: string | undefined): string | undefined {
  return given !== undefined ? given : ofView;
}

/**
 * The sheets a job is of, in order: the one shown ('active', the default),
 * every visible one ('all'), or those of a list as far as the workbook still
 * has them. What printing, a PDF and the print preview all go by.
 */
export function jobSheets(workbook: Workbook, which: 'active' | 'all' | readonly number[] | undefined, shown: number): number[] {
  const sheets: number[] = [];
  if (which === 'all') {
    for (let i = 0; i < workbook.sheets.length; i++) if (workbook.sheets[i].state === 'visible') sheets.push(i);
  } else if (which !== undefined && which !== 'active') {
    for (let i = 0; i < which.length; i++) if (which[i] < workbook.sheets.length) sheets.push(which[i]);
  } else sheets.push(shown);
  return sheets;
}

/**
 * The sheets the print preview shows, as a list its Print and PDF hand on as they are: 'active' is
 * the sheet the preview was opened on (followed wherever it was moved; the sheet shown when it is
 * gone), not the sheet that happens to be shown later; 'all' the visible ones.
 */
export function previewSheets(workbook: Workbook, which: 'active' | 'all', opened: Sheet | undefined, shown: number): number[] {
  if (which === 'all') return jobSheets(workbook, 'all', shown);
  const at = opened ? workbook.sheets.indexOf(opened) : -1;
  return [at >= 0 ? at : shown];
}

/**
 * The pages of a print job over the given sheets. A chart sheet is one page
 * (see chartSheetLayout). Sheets that are neither worksheets nor chart
 * sheets, chart sheets whose chart is not drawn, and sheets with nothing to
 * print are left out; a job without any page throws a DocumentPrintError
 * 'empty', one beyond the limit 'too-long'.
 *
 * Pages are numbered through the whole job; a sheet that states the number
 * of its first page starts there. UNSURE: Excel numbers the same way when a
 * whole workbook is printed, as far as it could be told without Excel.
 */
export function planSheetPages(workbook: Workbook, sheets: readonly number[], options: SheetPlanOptions = {}): SheetJobPlan {
  const limit = typeof options.maxPages === 'number' && options.maxPages >= 1 ? Math.floor(options.maxPages) : MAX_SHEET_PAGES;
  const pages: SheetJobPage[] = [];
  const printed: number[] = [];
  let next = 1;
  for (let i = 0; i < sheets.length; i++) {
    const index = sheets[i];
    const sheet = workbook.sheets[index];
    const chart = sheet ? chartOfSheet(sheet) : undefined;
    if (sheet && chart) {
      const layout = chartSheetLayout(sheet, options.defaultPaper);
      if (layout.firstPageNumber !== undefined) next = layout.firstPageNumber;
      pages.push({ sheet: index, layout, page: layout.pages[0], index: 0, number: next++, chart });
      printed.push(index);
      if (pages.length > limit) break;
      continue;
    }
    if (!sheet || sheet.kind !== 'worksheet') continue;
    const metrics = options.metrics ? options.metrics(sheet, index) : createSheetMetrics(sheet);
    const ranges = printRanges(workbook, index);
    let layout: SheetPrintLayout;
    try {
      layout = paginateSheet(sheet, {
        metrics,
        styles: workbook.styles,
        areas: ranges.areas,
        titleRows: ranges.titleRows,
        titleCols: ranges.titleCols,
        drawings: options.drawings ? options.drawings(sheet, index, metrics) : undefined,
        defaultPaper: options.defaultPaper,
        // One page more than there is room for shows that the job is too long.
        maxPages: Math.max(limit - pages.length, 0) + 1,
      });
    } catch (error) {
      if (error instanceof SheetPagesError) {
        if (error.code === 'empty') continue;
        throw new DocumentPrintError('too-long', `The workbook has more than ${limit} pages to print`);
      }
      throw error;
    }
    if (layout.firstPageNumber !== undefined) next = layout.firstPageNumber;
    for (let p = 0; p < layout.pages.length; p++) pages.push({ sheet: index, layout, page: layout.pages[p], index: p, number: next++ });
    printed.push(index);
    if (pages.length > limit) break;
  }
  if (pages.length === 0) throw new DocumentPrintError('empty', 'There is nothing to print');
  if (pages.length > limit) throw new DocumentPrintError('too-long', `The workbook has more than ${limit} pages to print`);
  return { pages, sheets: printed };
}

// ---------------------------------------------------------------------------
// Painting

export interface SheetPagePaint {
  workbook: Workbook;
  job: SheetJobPage;
  /** Pages of the whole job: what "&N" shows. */
  total: number;
  /** Row and column sizes of the sheet; not needed for the page of a chart sheet. */
  metrics: SheetMetrics;
  /** What draws the chart of a chart sheet; one is made for the page when missing. */
  charts?: Pick<ChartRenderer, 'paint'>;
  /** Width of the canvas in pixels; the height follows from the paper. */
  pixelWidth: number;
  locale?: FormatLocale;
  conditional?: ConditionalSource;
  drawings?: GridDrawingSource;
  fileName?: string;
  /** The moment "&D" and "&T" show. */
  now: Date;
}

/** Margins of a header that is not aligned with the page margins: 0.7 inch (UNSURE: Excel's value is not documented). */
const FREE_HEADER_MARGIN = 0.7 * 96;
const LINE_HEIGHT = 1.2;
const DEFAULT_POINTS = 11;
/** How much smaller text above or below the line is, and how far it is moved. */
const SCRIPT_SIZE = 0.65;
const SCRIPT_SHIFT = 0.35;

/** A date as a serial number of the sheet's calendar, in local time. */
function serialOf(date: Date, date1904: boolean): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds());
  const days = utc / 86400000 + 25569;
  return date1904 ? days - 1462 : days;
}

/** What the fields of a header or footer show on a page. */
export function headerFooterValues(paint: Pick<SheetPagePaint, 'workbook' | 'job' | 'total' | 'locale' | 'fileName' | 'now'>): HeaderFooterValues {
  const { workbook, job } = paint;
  const sheet = workbook.sheets[job.sheet];
  const date1904 = workbook.date1904 === true;
  let date = '';
  let time = '';
  if (paint.now instanceof Date && !isNaN(paint.now.getTime())) {
    const serial = serialOf(paint.now, date1904);
    // Built-in format 14 is the short date of the system; the time is its short time.
    date = formatValue(serial, 'm/d/yyyy', { locale: paint.locale, date1904, builtinId: 14 }).text;
    time = formatValue(serial, (paint.locale && paint.locale.shortTime) || 'h:mm', { locale: paint.locale, date1904 }).text;
  }
  const name = typeof paint.fileName === 'string' ? paint.fileName.slice(0, 512) : '';
  // No path is known in a page: "&Z" is empty.
  return { page: job.number, pages: paint.total, date, time, file: name, path: '', sheet: sheet ? sheet.name : '' };
}

/** The code of the header and of the footer that belong on a page. */
export function headerFooterFor(setup: HeaderFooter | undefined, firstOfSheet: boolean, pageNumber: number): { header: string | undefined; footer: string | undefined } {
  if (!setup) return { header: undefined, footer: undefined };
  if (setup.differentFirst === true && firstOfSheet) return { header: setup.firstHeader, footer: setup.firstFooter };
  if (setup.differentOddEven === true && pageNumber % 2 === 0) return { header: setup.evenHeader, footer: setup.evenFooter };
  return { header: setup.oddHeader, footer: setup.oddFooter };
}

interface SectionBox {
  left: number;
  right: number;
  /** Top of the first line (header) or bottom of the last one (footer). */
  edge: number;
  footer: boolean;
}

function pieceFont(format: HeaderFooterFormat, base: Font, scale: number): { css: string; px: number } {
  const points = (format.size !== undefined ? format.size : base.size !== undefined ? base.size : DEFAULT_POINTS) * (format.vertical ? SCRIPT_SIZE : 1) * scale;
  const font: Font = {
    name: format.font !== undefined ? format.font : base.name,
    size: points,
    bold: format.bold === true,
    italic: format.italic === true,
    strike: false,
    underline: 'none',
  };
  return { css: cssFont(font), px: fontPixels(font) };
}

function paintSection(ctx: CanvasRenderingContext2D, runs: HeaderFooterRun[], align: 'left' | 'center' | 'right', box: SectionBox, values: HeaderFooterValues, base: Font, scale: number): void {
  if (runs.length === 0) return;
  const lines = headerFooterLines(runs, values);
  if (lines.length === 0) return;
  const basePx = fontPixels({ size: (base.size !== undefined ? base.size : DEFAULT_POINTS) * scale });
  interface Measured {
    piece: HeaderFooterPiece;
    css: string;
    px: number;
    width: number;
  }
  const measured: { pieces: Measured[]; width: number; height: number; px: number }[] = [];
  let total = 0;
  for (let i = 0; i < lines.length; i++) {
    const pieces: Measured[] = [];
    let width = 0;
    let px = lines[i].length === 0 ? basePx : 0;
    for (let p = 0; p < lines[i].length; p++) {
      const piece = lines[i][p];
      const font = pieceFont(piece.format, base, scale);
      ctx.font = font.css;
      const w = ctx.measureText(piece.text).width;
      pieces.push({ piece, css: font.css, px: font.px, width: w });
      width += w;
      // A piece above or below the line does not make the line higher.
      px = Math.max(px, piece.format.vertical ? font.px / SCRIPT_SIZE : font.px);
    }
    const height = px * LINE_HEIGHT;
    measured.push({ pieces, width, height, px });
    total += height;
  }
  let y = box.footer ? box.edge - total : box.edge;
  for (let i = 0; i < measured.length; i++) {
    const line = measured[i];
    const baseline = y + line.height - line.px * 0.3;
    let x = align === 'left' ? box.left : align === 'right' ? box.right - line.width : (box.left + box.right - line.width) / 2;
    for (let p = 0; p < line.pieces.length; p++) {
      const { piece, css, px, width } = line.pieces[p];
      const format = piece.format;
      const at = format.vertical === 'super' ? baseline - line.px * SCRIPT_SHIFT : format.vertical === 'sub' ? baseline + line.px * SCRIPT_SHIFT * 0.4 : baseline;
      ctx.font = css;
      // parseHeaderFooter() only lets "#RRGGBB" through.
      ctx.fillStyle = format.color !== undefined ? format.color : '#000000';
      ctx.fillText(piece.text, x, at);
      const thickness = Math.max(px / 14, 0.5);
      if (format.underline) {
        const under = at + Math.max(px * 0.1, 1);
        ctx.fillRect(x, under, width, thickness);
        if (format.underline === 2) ctx.fillRect(x, under + thickness * 2, width, thickness);
      }
      if (format.strike) ctx.fillRect(x, at - px * 0.3, width, thickness);
      x += width;
    }
    y += line.height;
  }
}

/** The header and the footer of a page, in paper px under the transform of the caller. */
function paintHeaderFooter(ctx: CanvasRenderingContext2D, paint: SheetPagePaint): void {
  const { workbook, job } = paint;
  const sheet = workbook.sheets[job.sheet];
  const setup = sheet && sheet.pageSetup ? sheet.pageSetup.headerFooter : undefined;
  const codes = headerFooterFor(setup, job.index === 0, job.number);
  if (!codes.header && !codes.footer) return;
  const layout = job.layout;
  const values = headerFooterValues(paint);
  const base = resolveCellStyle(workbook.styles, 0).font;
  const scale = setup && setup.scaleWithDoc === false ? 1 : layout.scale;
  const aligned = !setup || setup.alignWithMargins !== false;
  const left = aligned ? layout.margins.left : Math.min(FREE_HEADER_MARGIN, layout.paper.width / 4);
  const right = layout.paper.width - (aligned ? layout.margins.right : Math.min(FREE_HEADER_MARGIN, layout.paper.width / 4));
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  const both: [string | undefined, boolean][] = [
    [codes.header, false],
    [codes.footer, true],
  ];
  for (let i = 0; i < both.length; i++) {
    const code = both[i][0];
    if (!code) continue;
    const footer = both[i][1];
    const sections = parseHeaderFooter(code);
    const box: SectionBox = { left, right, edge: footer ? layout.paper.height - layout.margins.footer : layout.margins.header, footer };
    paintSection(ctx, sections.left, 'left', box, values, base, scale);
    paintSection(ctx, sections.center, 'center', box, values, base, scale);
    paintSection(ctx, sections.right, 'right', box, values, base, scale);
  }
}

/**
 * Paints one page of a job on a canvas of `pixelWidth` pixels across: white
 * paper, the cells at their place and scale, the header and the footer.
 * Every pixel of the canvas is written.
 */
export function paintSheetPage(ctx: CanvasRenderingContext2D, paint: SheetPagePaint): void {
  const { workbook, job } = paint;
  const { layout, page } = job;
  const sheet = workbook.sheets[job.sheet];
  const k = paint.pixelWidth / layout.paper.width;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = 1;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, Math.ceil(layout.paper.width * k) + 1, Math.ceil(layout.paper.height * k) + 1);
  if (!sheet) return;

  if (job.chart) {
    // A chart sheet: the chart in what the margins leave of the paper, laid out in px of the paper.
    const charts = paint.charts || new ChartRenderer({ workbook, locale: () => paint.locale });
    ctx.save();
    try {
      ctx.setTransform(k, 0, 0, k, 0, 0);
      ctx.beginPath();
      ctx.rect(page.x, page.y, page.width, page.height);
      ctx.clip();
      charts.paint(ctx, job.chart, page.x, page.y, Math.floor(page.width), Math.floor(page.height));
    } finally {
      ctx.restore();
    }
    ctx.save();
    try {
      ctx.setTransform(k, 0, 0, k, 0, 0);
      paintHeaderFooter(ctx, paint);
    } finally {
      ctx.restore();
    }
    return;
  }

  const s = layout.scale;
  const bands = (title: GridBand | undefined, own: GridBand): GridBand[] => (title ? [title, own] : [own]);
  ctx.save();
  try {
    ctx.setTransform(k * s, 0, 0, k * s, k * page.x, k * page.y);
    // What is wider or higher than the page ends at the margin.
    ctx.beginPath();
    ctx.rect(0, 0, page.width / s, page.height / s);
    ctx.clip();
    renderGridPrint(ctx, {
      workbook,
      sheet,
      metrics: { rows: paint.metrics.rows, cols: paint.metrics.cols, headerWidth: layout.headingWidth, headerHeight: layout.headingHeight },
      rows: bands(page.titleRows, page.rows),
      cols: bands(page.titleCols, page.cols),
      headings: layout.headings,
      gridLines: layout.gridLines,
      locale: paint.locale,
      conditional: paint.conditional,
      drawings: paint.drawings,
      scale: k * s,
    });
  } finally {
    ctx.restore();
  }

  ctx.save();
  try {
    ctx.setTransform(k, 0, 0, k, 0, 0);
    paintHeaderFooter(ctx, paint);
  } finally {
    ctx.restore();
  }
}
