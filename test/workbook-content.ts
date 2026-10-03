// What a workbook model holds, as plain data, for comparing a model with the
// one read back from a file written from it (buildXlsx). Two models with the
// same content give the same text, whatever their table indexes are: formats
// are compared by what they resolve to, shared formulas by the text each cell
// has, pictures by their bytes.
//
// The two sides are not treated alike. The model that was written is turned
// into what the file is EXPECTED to read back as (`expected: true`): where the
// writer adds something the model does not state (a default row height, the
// "Normal" cell style, the name of an object, the box of a comment), or leaves
// something out with a warning (the placeholder of a chart, the replies of a
// conversation), the expected value is written out here, one rule at a time.
// The model read back is taken as it is. So a field the writer loses shows up
// as a difference; nothing is passed over on both sides.

import { shiftFormula } from '../src/calc/refs';
import { createSheetMetrics, type SheetMetrics } from '../src/layout/sheet-metrics';
import {
  MAX_COLS,
  MAX_ROWS,
  cellName,
  parseRange,
  type Cell,
  type CellFormat,
  type DrawingAnchor,
  type DrawingMarker,
  type DrawingObject,
  type DrawingText,
  type DrawingTextRun,
  type Sheet,
  type SheetComment,
  type StyleSheet,
  type Workbook,
} from '../src/model/sheet/types';
import { formatCodeOf } from '../src/model/sheet/color';
import { anchorRect } from '../src/render/sheet-drawings';
import { localName, type XmlElement } from '../src/xml';

const EMU = 9525;

/** JSON with the keys of every object sorted and undefined members left out. */
export function canonical(value: unknown): string {
  return JSON.stringify(value, (_key, item) => {
    if (item instanceof Map) return [...item.entries()];
    if (item && typeof item === 'object' && !Array.isArray(item)) {
      const sorted: Record<string, unknown> = {};
      for (const key of Object.keys(item).sort()) if (item[key] !== undefined) sorted[key] = item[key];
      return sorted;
    }
    return item;
  });
}

function formatContent(styles: StyleSheet, xf: CellFormat | undefined, parents: boolean): unknown {
  if (!xf) return null;
  return {
    font: styles.fonts[xf.fontId],
    fill: styles.fills[xf.fillId],
    border: styles.borders[xf.borderId],
    numFmtId: xf.numFmtId,
    numFmt: formatCodeOf(styles, xf.numFmtId),
    alignment: xf.alignment,
    protection: xf.protection,
    quotePrefix: xf.quotePrefix,
    parent: parents && xf.xfId !== undefined ? formatContent(styles, styles.cellStyleXfs[xf.xfId], false) : undefined,
  };
}

/** What each cell format of a workbook resolves to, as text, by index; an index outside the table is the default format. */
function formatKeys(styles: StyleSheet): (index: number | undefined) => string {
  const keys: string[] = [];
  return (index) => {
    const at = index !== undefined && Number.isInteger(index) && index >= 0 && index < styles.cellXfs.length ? index : 0;
    return keys[at] !== undefined ? keys[at] : (keys[at] = canonical(formatContent(styles, styles.cellXfs[at], true)));
  };
}

/** An element kept as it was read, without the prefix of its names: a differential format means the same under any prefix. */
function elementContent(el: XmlElement): unknown {
  return {
    name: localName(el.name),
    attrs: [...el.attrs].filter(([name]) => name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:').sort(),
    children: el.children.filter((node): node is XmlElement => node.type === 'element').map(elementContent),
  };
}

/** The text of the formula of every cell, shared formulas written out, as the writer writes them out where a file cannot share them. */
function formulaContent(cell: Cell, row: number, col: number, masters: Map<number, { row: number; col: number; text: string }>): unknown {
  const formula = cell.formula;
  if (!formula) return undefined;
  if (formula.type !== 'shared') return { text: formula.text, type: formula.type, ref: formula.ref, attrs: formula.attrs && [...formula.attrs].sort() };
  let text = formula.text;
  if (text === '') {
    const master = formula.si === undefined ? undefined : masters.get(formula.si);
    // A cell of a group without a first cell has no formula in the file.
    if (!master) return undefined;
    text = shiftFormula(master.text, row - master.row, col - master.col);
  }
  return { text, type: 'normal' };
}

function roundAll<T extends object>(value: T, scale: number): Record<string, number> {
  const out: Record<string, number> = {};
  for (const [key, item] of Object.entries(value)) out[key] = Math.round((item as number) * scale) / scale;
  return out;
}

const round = (value: number, scale: number) => Math.round(value * scale) / scale;

/** The runs of a shape's text as a file has them: a line break is a run of its own, with the format of the run it was in. */
function textContent(text: DrawingText | undefined, expected: boolean): unknown {
  if (!text) return undefined;
  const paragraphs = text.paragraphs.map((paragraph) => {
    const runs: DrawingTextRun[] = [];
    for (const run of paragraph.runs) {
      if (!expected) {
        runs.push(run);
        continue;
      }
      const lines = String(run.text).split('\n');
      lines.forEach((line, i) => {
        if (i > 0) runs.push({ ...run, text: '\n' });
        if (line !== '') runs.push({ ...run, text: line });
      });
    }
    return {
      align: paragraph.align,
      runs: runs.map((run) => ({
        text: run.text,
        font: run.font === '' || (expected && run.font !== undefined && run.font.charAt(0) === '+') ? undefined : run.font,
        size: run.size === undefined ? undefined : round(run.size, 100),
        bold: run.bold,
        italic: run.italic,
        underline: run.underline,
        strike: run.strike,
        color: run.color === undefined ? undefined : run.color.toUpperCase(),
      })),
    };
  });
  // A text without a character is no text in a file.
  if (expected && !paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text !== '\n'))) return undefined;
  return {
    anchor: text.anchor === 'top' ? undefined : text.anchor,
    wrap: text.wrap === false ? false : undefined,
    // A file always states the space around a text; the model's default is 0.1 inch at the sides and 0.05 above and below.
    insets: (text.insets || (expected ? [9.6, 4.8, 9.6, 4.8] : undefined))?.map((inset) => round(inset, 100)),
    paragraphs,
  };
}

interface ObjectContext {
  digests: Map<string, string> | undefined;
  expected: boolean;
  /** Objects written so far in the drawing: the writer numbers the names it gives. */
  count: number;
}

/** `undefined` for what the writer leaves out (with a warning): the placeholder of a chart, a group member without a place. */
function objectContent(object: DrawingObject, context: ObjectContext, depth: number): unknown {
  if (object.type === 'placeholder') return context.expected ? undefined : { type: 'placeholder', what: object.what };
  const id = ++context.count;
  const given = (kind: string) => (object.name !== undefined && object.name !== '' ? object.name : context.expected ? `${kind} ${id}` : undefined);
  const base = { rotation: object.rotation || undefined, flipH: object.flipH || undefined, flipV: object.flipV || undefined, hidden: object.hidden || undefined, description: object.description || undefined };
  if (object.type === 'picture') {
    // A picture whose bytes cannot be read is written as a frame without a picture; the address of a linked one is not in the model.
    const media = object.media !== undefined && context.digests ? context.digests.get(object.media) : undefined;
    return { ...base, type: 'picture', name: given('Picture'), media, external: context.expected ? undefined : object.external, crop: object.crop && roundAll(object.crop, 1e5) };
  }
  if (object.type === 'shape') {
    const text = textContent(object.text, context.expected);
    return {
      ...base,
      type: 'shape',
      name: given(object.connector && !object.text ? 'Connector' : 'Shape'),
      geometry: object.geometry,
      connector: object.connector || undefined,
      fill: object.fill,
      line: object.line && { color: object.line.color, dash: object.line.dash || undefined, width: Math.min(Math.max(round(object.line.width, 100), 0.75), 32) },
      text,
    };
  }
  const children: unknown[] = [];
  if (depth < 16) {
    for (const child of object.children) {
      if (context.expected && !child.box) continue;
      const content = objectContent(child, context, depth + 1);
      if (content !== undefined) children.push({ box: child.box && roundAll(child.box, 1e3), object: content });
    }
  }
  return { ...base, type: 'group', name: given('Group'), children };
}

/** The place of an anchor in whole pixels: a file states offsets in EMU, an xls model in shares of the cell. */
function rectContent(anchor: DrawingAnchor | undefined, metrics: SheetMetrics): unknown {
  if (!anchor) return undefined;
  const rect = anchorRect(anchor, metrics);
  return [Math.round(rect.x), Math.round(rect.y), Math.round(rect.width), Math.round(rect.height)];
}

/** The box of a comment as a file states it: cells and whole pixels into them. */
function boxContent(comment: SheetComment, metrics: SheetMetrics, expected: boolean): unknown {
  const corner = (marker: DrawingMarker) => {
    const width = metrics.cols.size(marker.col);
    const height = metrics.rows.size(marker.row);
    const share = (value: number) => (value > 0 ? (value < 1 ? value : 1) : 0);
    const x = marker.colShare !== undefined ? share(marker.colShare) * width : expected ? Math.min(marker.colOff / EMU, width) : marker.colOff / EMU;
    const y = marker.rowShare !== undefined ? share(marker.rowShare) * height : expected ? Math.min(marker.rowOff / EMU, height) : marker.rowOff / EMU;
    return [marker.col, Math.round(x), marker.row, Math.round(y)];
  };
  const box = comment.box;
  if (box && box.kind === 'twoCell' && box.from && box.to) return [...corner(box.from), ...corner(box.to)];
  // A comment the model gives no box gets the one applications give a new comment: beside its cell.
  if (expected) return [Math.min(comment.col + 1, MAX_COLS - 1), 15, Math.max(comment.row - 1, 0), 10, Math.min(comment.col + 3, MAX_COLS - 1), 15, Math.min(comment.row + 3, MAX_ROWS - 1), 4];
  return undefined;
}

function commentContent(comment: SheetComment, metrics: SheetMetrics, expected: boolean): unknown {
  let runs = comment.runs;
  // The reader of a file takes the formatting of a comment only when it has more than one run and one of them
  // states a font (the font of a comment that is one run is checked in the file itself, by a test).
  if (expected && runs && !(runs.length > 1 && runs.some((run) => run.font))) runs = undefined;
  return {
    row: comment.row,
    col: comment.col,
    author: comment.author,
    text: comment.text,
    runs,
    visible: comment.visible || undefined,
    box: boxContent(comment, metrics, expected),
    // A conversation is written as a note with its first message (and a warning).
    threaded: expected ? undefined : comment.threaded,
    replies: expected ? undefined : comment.replies,
  };
}

function oneCell(ref: string | undefined): string | undefined {
  const range = ref === undefined ? undefined : parseRange(ref);
  return range ? cellName(range.top, range.left) : undefined;
}

function sheetContent(sheet: Sheet, format: (index: number | undefined) => string, digests: Map<string, string> | undefined, expected: boolean): unknown {
  const metrics = createSheetMetrics(sheet);
  const indexes = [...sheet.rows.keys()].filter((index) => Number.isInteger(index) && index >= 0 && index < MAX_ROWS).sort((a, b) => a - b);
  const masters = new Map<number, { row: number; col: number; text: string }>();
  for (const index of indexes) {
    sheet.rows.get(index)!.cells.forEach((cell, col) => {
      const formula = cell.formula;
      if (formula && formula.type === 'shared' && formula.text !== '' && formula.si !== undefined && !masters.has(formula.si)) masters.set(formula.si, { row: index, col, text: formula.text });
    });
  }
  const plain = format(0);
  const rows: unknown[] = [];
  for (const index of indexes) {
    const row = sheet.rows.get(index)!;
    const cells: unknown[] = [];
    row.cells.forEach((cell, col) => {
      if (col >= MAX_COLS) return;
      const formula = formulaContent(cell, index, col, masters);
      const style = format(cell.style);
      // A blank cell without a format or a formula says nothing; the writer leaves it out.
      if (cell.kind === 'blank' && !formula && style === plain) return;
      cells.push([col, cell.kind, cell.value, formula, style, cell.runs, cell.ph]);
    });
    const props = [row.height, row.customHeight || undefined, row.hidden || undefined, row.style === undefined ? undefined : format(row.style), row.outlineLevel || undefined, row.collapsed || undefined];
    // Likewise a row with neither cells nor properties.
    if (cells.length > 0 || props.some((prop) => prop !== undefined)) rows.push([index, props, cells]);
  }

  const view = { ...sheet.view };
  if (expected) {
    view.activeCell = oneCell(view.activeCell);
    view.firstVisibleCell = oneCell(view.firstVisibleCell);
    view.topLeftCell = view.frozenRows > 0 || view.frozenCols > 0 ? oneCell(view.topLeftCell) : undefined;
    if ((view.frozenRows > 0 || view.frozenCols > 0) && view.topLeftCell === undefined) {
      // A file states where the scrollable pane starts; the writer works it out where the model does not say.
      const first = view.firstVisibleCell === undefined ? undefined : parseRange(view.firstVisibleCell);
      view.topLeftCell = cellName(Math.min((first ? first.top : 0) + view.frozenRows, MAX_ROWS - 1), Math.min((first ? first.left : 0) + view.frozenCols, MAX_COLS - 1));
    }
  }
  const sized = sheet.baseColWidth !== undefined || sheet.defaultColWidth !== undefined;
  const objects: ObjectContext = { digests, expected, count: 0 };
  const drawings: unknown[] = [];
  for (const drawing of sheet.drawings || []) {
    const object = objectContent(drawing.object, objects, 0);
    if (object !== undefined) drawings.push({ rect: rectContent(drawing.anchor, metrics), kind: drawing.anchor.kind, editAs: drawing.anchor.editAs, object });
  }
  return {
    name: sheet.name,
    state: sheet.state,
    view,
    tabColor: sheet.tabColor,
    baseColWidth: sheet.baseColWidth,
    defaultColWidth: sheet.defaultColWidth,
    // The element that states the widths has to state a height: 15 points where the model has none.
    defaultRowHeight: sheet.defaultRowHeight === undefined && expected && sized ? 15 : sheet.defaultRowHeight,
    columns: sheet.columns.map((range) => ({ ...range, style: range.style === undefined ? undefined : format(range.style) })),
    merges: sheet.merges,
    // A link without a target leads nowhere and is not written (with a warning); the relationship id is the file's own business.
    links: sheet.hyperlinks.filter((link) => !expected || link.target || link.location).map((link) => [link.range, link.target || undefined, link.location || undefined, link.display, link.tooltip]),
    rows,
    conditionalFormats: sheet.conditionalFormats,
    dataValidations: sheet.dataValidations,
    autoFilter: sheet.autoFilter && { range: sheet.autoFilter.range, columns: sheet.autoFilter.columns.map((column) => ({ ...column, raw: undefined })) },
    drawings,
    comments: (sheet.comments || []).map((comment) => commentContent(comment, metrics, expected)),
  };
}

function writtenNames(workbook: Workbook, count: number): Workbook['definedNames'] {
  const scopes = new Set<string>();
  return workbook.definedNames.filter((name) => {
    if (name.name === '' || name.value === '' || (name.localSheet !== undefined && !(name.localSheet >= 0 && name.localSheet < count))) return false;
    const scope = `${name.localSheet === undefined ? '' : name.localSheet}!${name.name.toLowerCase()}`;
    if (scopes.has(scope)) return false;
    scopes.add(scope);
    return true;
  });
}

export interface ContentOptions {
  /** Digests of the pictures by their key in the model (see mediaDigests); without them pictures are compared by presence. */
  digests?: Map<string, string>;
  /** The model is one that was written: its content is what the file made of it is expected to read back as. */
  expected?: boolean;
}

/** The content of a workbook as plain data; compare two with canonical(). */
export function workbookContent(workbook: Workbook, options: ContentOptions = {}): { book: unknown; sheets: unknown[] } {
  const expected = options.expected === true;
  const styles = workbook.styles;
  const format = formatKeys(styles);
  const count = workbook.sheets.length;
  const hidden = (index: number) => workbook.sheets[index] !== undefined && workbook.sheets[index].state !== 'visible';
  let active = workbook.activeSheet;
  if (expected) {
    // A file opens on a sheet that is shown: the first visible one when the model's active sheet is hidden.
    const shown = workbook.sheets.findIndex((sheet) => sheet.state === 'visible');
    if (!(active >= 0 && active < count)) active = 0;
    if (shown >= 0 && hidden(active)) active = shown;
  }
  // A workbook without named cell styles gets the one every file has.
  const cellStyles = styles.cellStyles.length === 0 && expected ? [{ name: 'Normal', xfId: 0, builtinId: 0 }] : styles.cellStyles;
  return {
    book: {
      date1904: workbook.date1904,
      activeSheet: active,
      // A name without a value, of a sheet that does not exist, or stated a second time in its scope is not written (with a warning).
      names: expected ? writtenNames(workbook, count) : workbook.definedNames,
      palette: styles.indexedColors,
      theme: styles.themeColors.length > 0 ? styles.themeColors : undefined,
      cellStyles: cellStyles.map((style) => ({ name: style.name, builtinId: style.builtinId, format: canonical(formatContent(styles, styles.cellStyleXfs[style.xfId], false)) })),
      dxfs: styles.dxfs.map(elementContent),
      // Chart sheets and macro sheets are written as empty worksheets (with a warning), so the kind is not compared.
      sheets: workbook.sheets.map((sheet) => [sheet.name, sheet.state]),
    },
    sheets: workbook.sheets.map((sheet) => sheetContent(sheet, format, options.digests, expected)),
  };
}

function digest(bytes: Uint8Array): string {
  // FNV-1a, with the length: enough to tell pictures apart.
  let hash = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) hash = Math.imul(hash ^ bytes[i], 0x01000193) >>> 0;
  return `${bytes.length}:${hash.toString(16)}`;
}

/** A digest of every picture the drawings of a workbook show; a picture whose bytes cannot be read has none. */
export async function mediaDigests(workbook: Workbook): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const visit = async (object: DrawingObject, depth: number): Promise<void> => {
    if (object.type === 'group' && depth < 16) for (const child of object.children) await visit(child, depth + 1);
    if (object.type !== 'picture' || object.media === undefined || out.has(object.media) || !workbook.media) return;
    try {
      out.set(object.media, digest(await workbook.media.readPart(object.media)));
    } catch (error) {
      // Compared as a frame without a picture.
    }
  };
  for (const sheet of workbook.sheets) for (const drawing of sheet.drawings || []) await visit(drawing.object, 0);
  return out;
}

/**
 * The first difference between a workbook that was written with buildXlsx() and the one read back
 * from the file, in words; undefined when the file holds what the model is expected to give.
 */
export async function workbookDifference(written: Workbook, readBack: Workbook): Promise<string | undefined> {
  const a = workbookContent(written, { digests: await mediaDigests(written), expected: true });
  const b = workbookContent(readBack, { digests: await mediaDigests(readBack) });
  if (canonical(a.book) !== canonical(b.book)) return `workbook: ${firstDifference(a.book, b.book)}`;
  for (let i = 0; i < Math.max(a.sheets.length, b.sheets.length); i++) {
    if (canonical(a.sheets[i]) !== canonical(b.sheets[i])) return `sheet ${i + 1}: ${firstDifference(a.sheets[i], b.sheets[i])}`;
  }
  return undefined;
}

/** Path and the two values of the first place where two plain values differ. */
function firstDifference(a: unknown, b: unknown, path = ''): string {
  if (canonical(a) === canonical(b)) return '';
  if (a && b && typeof a === 'object' && typeof b === 'object' && Array.isArray(a) === Array.isArray(b)) {
    const keys = Array.isArray(a) ? Array.from({ length: Math.max(a.length, (b as unknown[]).length) }, (_, i) => String(i)) : [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
    for (const key of keys) {
      const left = (a as Record<string, unknown>)[key];
      const right = (b as Record<string, unknown>)[key];
      if (canonical(left) !== canonical(right)) return firstDifference(left, right, `${path}/${key}`);
    }
  }
  const short = (value: unknown) => {
    const text = canonical(value);
    return text === undefined ? 'undefined' : text.length > 160 ? `${text.slice(0, 160)}...` : text;
  };
  return `${path || '/'}: ${short(a)} != ${short(b)}`;
}
