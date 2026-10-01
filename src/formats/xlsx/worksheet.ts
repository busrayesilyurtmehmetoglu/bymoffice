// Worksheet part (xl/worksheets/sheetN.xml) -> sheet model.
//
// The part is scanned as a stream, never as a tree: a sheet can hold a million
// cells. Elements the model has no place for are skipped; they stay in the
// source package and are written back from there.

import {
  MAX_COLS,
  MAX_ROWS,
  cellName,
  parseCellRef,
  parseRange,
  type Cell,
  type CellFormula,
  type ColumnRange,
  type Hyperlink,
  type Row,
  type SharedString,
  type Sheet,
  type SheetView,
} from '../../model/sheet/types';
import { XmlScanner } from '../../xml-scan';
import { decodeXstring, parseColor, parseStringItem } from './shared-strings';

const ROOT_NAMES = ['worksheet', 'chartsheet', 'dialogsheet', 'macrosheet'];

/** Plain non-negative integer, or -1. */
function parseUint(text: string | undefined): number {
  if (text === undefined) return -1;
  const n = text.length;
  if (n === 0 || n > 9) return -1;
  let value = 0;
  for (let i = 0; i < n; i++) {
    const c = text.charCodeAt(i);
    if (c < 48 || c > 57) return -1;
    value = value * 10 + (c - 48);
  }
  return value;
}

// xsd:double as files write it. Number() alone would also take "0x1F", "0b11" and "".
const DECIMAL = /^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$/;

function parseFinite(text: string | undefined): number | undefined {
  if (text === undefined || !DECIMAL.test(text)) return undefined;
  const value = Number(text);
  return isFinite(value) ? value : undefined;
}

/** File content quoted in an error message, cut so a hostile value cannot produce a huge message. */
const quote = (text: string) => (text.length > 40 ? `"${text.slice(0, 40)}…"` : `"${text}"`);

function parseBool(text: string | undefined, fallback: boolean): boolean {
  if (text === '1' || text === 'true') return true;
  if (text === '0' || text === 'false') return false;
  return fallback;
}

/**
 * Column index of a cell reference such as "B7" or "$B$7", or -1 when the
 * text is not a reference inside the sheet limits. Same rules as parseCellRef,
 * without allocating a result per cell.
 */
function columnOfRef(ref: string): number {
  let i = 0;
  const n = ref.length;
  if (ref.charCodeAt(i) === 36) i++;
  let col = 0;
  const lettersStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    if (col > MAX_COLS) return -1;
  }
  if (i === lettersStart) return -1;
  if (ref.charCodeAt(i) === 36) i++;
  let row = 0;
  const digitsStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i);
    if (c < 48 || c > 57) return -1;
    row = row * 10 + (c - 48);
    if (row > MAX_ROWS) return -1;
  }
  if (i === digitsStart || row === 0) return -1;
  return col - 1;
}

/** With a <sheetView> open tag current: reads the view and consumes the element. */
function parseSheetView(sc: XmlScanner, view: SheetView): void {
  view.showGridLines = parseBool(sc.attr('showGridLines'), true);
  view.showRowColHeaders = parseBool(sc.attr('showRowColHeaders'), true);
  view.showZeros = parseBool(sc.attr('showZeros'), true);
  view.rightToLeft = parseBool(sc.attr('rightToLeft'), false);
  view.tabSelected = parseBool(sc.attr('tabSelected'), false);
  const firstVisible = sc.attr('topLeftCell');
  if (firstVisible !== undefined && parseCellRef(firstVisible)) view.firstVisibleCell = firstVisible;
  const zoom = parseFinite(sc.attr('zoomScale'));
  if (zoom !== undefined && zoom >= 10 && zoom <= 400) view.zoomScale = zoom;
  if (sc.selfClosing) return;

  // The active cell is the one of the selection that belongs to the active pane.
  let activePane = 'topLeft';
  let firstActive: string | undefined;
  let paneActive: string | undefined;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local === 'pane') {
      activePane = sc.attr('activePane') || 'topLeft';
      const state = sc.attr('state');
      if (state === 'frozen' || state === 'frozenSplit') {
        const x = parseFinite(sc.attr('xSplit'));
        const y = parseFinite(sc.attr('ySplit'));
        if (x !== undefined && x > 0) view.frozenCols = Math.min(Math.floor(x), MAX_COLS - 1);
        if (y !== undefined && y > 0) view.frozenRows = Math.min(Math.floor(y), MAX_ROWS - 1);
        const topLeft = sc.attr('topLeftCell');
        if (topLeft !== undefined && parseCellRef(topLeft)) view.topLeftCell = topLeft;
      }
    } else if (sc.local === 'selection') {
      const active = sc.attr('activeCell');
      if (active !== undefined && parseCellRef(active)) {
        if (firstActive === undefined) firstActive = active;
        if (paneActive === undefined && (sc.attr('pane') || 'topLeft') === activePane) paneActive = active;
      }
    }
    sc.skip();
  }
  const active = paneActive !== undefined ? paneActive : firstActive;
  if (active !== undefined) view.activeCell = active;
}

function parseSheetViews(sc: XmlScanner, sheet: Sheet): void {
  if (sc.selfClosing) return;
  let first = true;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    // Only the first view is modelled.
    if (sc.local === 'sheetView' && first) {
      first = false;
      parseSheetView(sc, sheet.view);
    } else sc.skip();
  }
}

function parseSheetPr(sc: XmlScanner, sheet: Sheet): void {
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local === 'tabColor') {
      const color = parseColor(sc);
      if (color) sheet.tabColor = color;
    } else sc.skip();
  }
}

/**
 * Adds a column range, keeping the list sorted and free of overlaps. Where the
 * new range overlaps earlier ones the earlier ones win and only the uncovered
 * parts are added.
 */
function addColumnRange(columns: ColumnRange[], range: ColumnRange): void {
  const last = columns[columns.length - 1];
  if (!last || range.min > last.max) {
    columns.push(range);
    return;
  }
  // First existing range that ends at or after the start of the new one.
  let lo = 0;
  let hi = columns.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (columns[mid].max < range.min) lo = mid + 1;
    else hi = mid;
  }
  let at = lo;
  let start = range.min;
  while (start <= range.max) {
    const current = columns[at];
    if (!current || current.min > range.max) {
      columns.splice(at, 0, { ...range, min: start, max: range.max });
      return;
    }
    if (current.min > start) {
      columns.splice(at, 0, { ...range, min: start, max: current.min - 1 });
      at++;
    }
    start = current.max + 1;
    at++;
  }
}

function parseCols(sc: XmlScanner, sheet: Sheet, seen: { count: number }): void {
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local !== 'col') {
      sc.skip();
      continue;
    }
    // A valid sheet cannot state more ranges than it has columns; this bounds the work on a hostile file.
    if (++seen.count > MAX_COLS) throw new Error(`Invalid worksheet: more than ${MAX_COLS} <col> elements`);
    let min = parseUint(sc.attr('min'));
    let max = parseUint(sc.attr('max'));
    if (min >= 0 && max >= 0) {
      if (min < 1) min = 1;
      if (max > MAX_COLS) max = MAX_COLS;
      if (min <= max) {
        const range: ColumnRange = {
          min: min - 1,
          max: max - 1,
          customWidth: parseBool(sc.attr('customWidth'), false),
          hidden: parseBool(sc.attr('hidden'), false),
          outlineLevel: Math.max(parseUint(sc.attr('outlineLevel')), 0),
          collapsed: parseBool(sc.attr('collapsed'), false),
        };
        const width = parseFinite(sc.attr('width'));
        if (width !== undefined && width >= 0) range.width = width;
        const style = parseUint(sc.attr('style'));
        if (style >= 0) range.style = style;
        addColumnRange(sheet.columns, range);
      }
    }
    sc.skip();
  }
}

/** With an <f> open tag current: reads the formula and consumes the element. */
function parseFormula(sc: XmlScanner): CellFormula {
  const formula: CellFormula = { text: '', type: 'normal' };
  for (const [name, value] of sc.attrs()) {
    if (name === 't' && (value === 'normal' || value === 'array' || value === 'shared' || value === 'dataTable')) {
      formula.type = value;
    } else if (name === 'ref') {
      formula.ref = value;
    } else if (name === 'si' && parseUint(value) >= 0) {
      formula.si = parseUint(value);
    } else {
      if (!formula.attrs) formula.attrs = new Map();
      formula.attrs.set(name, value);
    }
  }
  // Formula text is taken literally: "_x0041_" inside a formula is a name, not an escape.
  formula.text = sc.readText();
  return formula;
}

/**
 * With a <c> open tag current: reads the cell into `cells` and consumes the
 * element. Returns the cell's column.
 */
function parseCell(sc: XmlScanner, cells: Cell[], rowIndex: number, prevCol: number, shared: SharedString[]): number {
  const ref = sc.attr('r');
  let col: number;
  if (ref === undefined) {
    col = prevCol + 1;
    if (col >= MAX_COLS) throw new Error(`Invalid worksheet: a cell in row ${rowIndex + 1} is past the last column`);
  } else {
    col = columnOfRef(ref);
    if (col < 0) throw new Error(`Invalid worksheet: cell reference ${quote(ref)} is not inside the sheet limits`);
  }
  const styleAttr = sc.attr('s');
  const style = styleAttr === undefined ? 0 : Math.max(parseUint(styleAttr), 0);
  const type = sc.attr('t');

  let v: string | undefined;
  let formula: CellFormula | undefined;
  let inline: SharedString | undefined;
  if (!sc.selfClosing) {
    while (sc.next()) {
      if (sc.kind === 'close') break;
      if (sc.kind !== 'open') continue;
      if (sc.local === 'v') v = sc.readText();
      else if (sc.local === 'f') formula = parseFormula(sc);
      else if (sc.local === 'is') inline = parseStringItem(sc);
      else sc.skip();
    }
  }

  let cell: Cell;
  if (type === undefined || type === 'n') {
    if (v === undefined || v === '') cell = { kind: 'blank', value: null, style };
    else {
      const value = parseFinite(v);
      if (value === undefined) {
        throw new Error(`Invalid worksheet: ${quote(v)} in cell ${cellName(rowIndex, col)} is not a number`);
      }
      cell = { kind: 'number', value, style };
    }
  } else if (type === 's') {
    if (v === undefined || v === '') cell = { kind: 'blank', value: null, style };
    else {
      const index = parseUint(v.trim());
      if (index < 0 || index >= shared.length) {
        throw new Error(
          `Invalid worksheet: shared string index ${quote(v)} in cell ${cellName(rowIndex, col)} is out of range (${shared.length} strings)`,
        );
      }
      const item = shared[index];
      cell = { kind: 'string', value: item.text, style };
      if (item.runs) cell.runs = item.runs;
    }
  } else if (type === 'b') {
    if (v === undefined || v === '') cell = { kind: 'blank', value: null, style };
    else {
      const text = v.trim();
      cell = { kind: 'boolean', value: text === '1' || text === 'true', style };
    }
  } else if (type === 'e') {
    cell = { kind: 'error', value: v === undefined ? '' : v, style };
  } else if (type === 'd') {
    if (v === undefined || v === '') cell = { kind: 'blank', value: null, style };
    else cell = { kind: 'date', value: v, style };
  } else if (type === 'inlineStr' && inline) {
    cell = { kind: 'string', value: inline.text, style };
    if (inline.runs) cell.runs = inline.runs;
  } else {
    // "str" (formula result), an inline string without <is>, or a type this reader does not know: text as written.
    cell = { kind: 'string', value: v === undefined ? '' : decodeXstring(v), style };
  }
  if (formula) cell.formula = formula;

  // A repeated cell replaces the earlier one.
  cells[col] = cell;
  return col;
}

function parseRowCells(sc: XmlScanner, row: Row, rowIndex: number, shared: SharedString[], sheet: Sheet): void {
  let prevCol = -1;
  let lastCol = -1;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local === 'c') {
      prevCol = parseCell(sc, row.cells, rowIndex, prevCol, shared);
      if (prevCol > lastCol) lastCol = prevCol;
    } else sc.skip();
  }
  if (lastCol >= 0) {
    if (rowIndex + 1 > sheet.rowCount) sheet.rowCount = rowIndex + 1;
    if (lastCol + 1 > sheet.colCount) sheet.colCount = lastCol + 1;
  }
}

function parseSheetData(sc: XmlScanner, shared: SharedString[], sheet: Sheet, state: { prevRow: number }): void {
  if (sc.selfClosing) return;
  const rows = sheet.rows;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local !== 'row') {
      sc.skip();
      continue;
    }

    const r = sc.attr('r');
    let rowIndex: number;
    if (r === undefined) rowIndex = state.prevRow + 1;
    else {
      rowIndex = parseUint(r) - 1;
      if (rowIndex < 0) throw new Error(`Invalid worksheet: row number ${quote(r)} is not inside the sheet limits`);
    }
    if (rowIndex >= MAX_ROWS) throw new Error(`Invalid worksheet: row ${rowIndex + 1} is past the last row`);
    state.prevRow = rowIndex;

    // A repeated row adds to the earlier one; its own properties replace the earlier ones.
    let row = rows.get(rowIndex);
    if (!row) {
      row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
      rows.set(rowIndex, row);
    }
    const height = parseFinite(sc.attr('ht'));
    if (height !== undefined && height >= 0) row.height = height;
    row.customHeight = parseBool(sc.attr('customHeight'), false);
    row.hidden = parseBool(sc.attr('hidden'), false);
    row.outlineLevel = Math.max(parseUint(sc.attr('outlineLevel')), 0);
    row.collapsed = parseBool(sc.attr('collapsed'), false);
    if (parseBool(sc.attr('customFormat'), false)) {
      const style = parseUint(sc.attr('s'));
      if (style >= 0) row.style = style;
    }

    if (!sc.selfClosing) parseRowCells(sc, row, rowIndex, shared, sheet);
  }
}

function parseMergeCells(sc: XmlScanner, sheet: Sheet): void {
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local === 'mergeCell') {
      const ref = sc.attr('ref');
      const range = ref === undefined ? undefined : parseRange(ref);
      if (range) sheet.merges.push(range);
    }
    sc.skip();
  }
}

function parseHyperlinks(sc: XmlScanner, sheet: Sheet): void {
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local === 'hyperlink') {
      let range: Hyperlink['range'] | undefined;
      const link: Partial<Hyperlink> = {};
      for (const [name, value] of sc.attrs()) {
        if (name === 'ref') range = parseRange(value);
        else if (name === 'location') link.location = value;
        else if (name === 'display') link.display = value;
        else if (name === 'tooltip') link.tooltip = value;
        // The relationship id is "r:id", but the prefix is the producer's choice.
        else if (name.slice(name.indexOf(':') + 1) === 'id') link.rId = value;
      }
      if (range) sheet.hyperlinks.push({ ...link, range });
    }
    sc.skip();
  }
}

/** Moves to the root element and checks its local name. */
function openRoot(sc: XmlScanner): void {
  while (sc.next()) {
    if (sc.kind === 'open') break;
  }
  if (ROOT_NAMES.indexOf(sc.local) < 0) {
    throw new Error(`Invalid worksheet: root element is <${sc.name}>, expected <worksheet>`);
  }
}

/** Fills `sheet` (created with createSheet) from a worksheet part. */
export function parseWorksheet(bytes: Uint8Array, sharedStrings: SharedString[], sheet: Sheet): void {
  const sc = new XmlScanner(bytes);
  openRoot(sc);

  const rowState = { prevRow: -1 };
  const colState = { count: 0 };
  if (!sc.selfClosing) {
    while (sc.next()) {
      if (sc.kind === 'close') break;
      if (sc.kind !== 'open') continue;
      switch (sc.local) {
        case 'dimension': {
          const ref = sc.attr('ref');
          const range = ref === undefined ? undefined : parseRange(ref);
          if (range) sheet.dimension = range;
          sc.skip();
          break;
        }
        case 'sheetViews':
          parseSheetViews(sc, sheet);
          break;
        case 'sheetFormatPr': {
          const rowHeight = parseFinite(sc.attr('defaultRowHeight'));
          if (rowHeight !== undefined && rowHeight >= 0) sheet.defaultRowHeight = rowHeight;
          const colWidth = parseFinite(sc.attr('defaultColWidth'));
          if (colWidth !== undefined && colWidth >= 0) sheet.defaultColWidth = colWidth;
          const baseWidth = parseFinite(sc.attr('baseColWidth'));
          if (baseWidth !== undefined && baseWidth >= 0) sheet.baseColWidth = baseWidth;
          sc.skip();
          break;
        }
        case 'sheetPr':
          parseSheetPr(sc, sheet);
          break;
        case 'cols':
          parseCols(sc, sheet, colState);
          break;
        case 'sheetData':
          parseSheetData(sc, sharedStrings, sheet, rowState);
          break;
        case 'mergeCells':
          parseMergeCells(sc, sheet);
          break;
        case 'hyperlinks':
          parseHyperlinks(sc, sheet);
          break;
        default:
          // Conditional formats, validations, drawings, ...: kept in the source part.
          sc.skip();
      }
    }
  }
  // Anything after the root is still checked for well-formedness.
  while (sc.next());
}
