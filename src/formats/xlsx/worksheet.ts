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
  type AutoFilter,
  type Cell,
  type CellFormula,
  type CellRange,
  type ColumnRange,
  type ConditionalColorScale,
  type ConditionalDataBar,
  type ConditionalFormat,
  type ConditionalIconSet,
  type ConditionalRule,
  type ConditionalValue,
  type DataValidation,
  type FilterColumn,
  type FilterCondition,
  type FilterDateGroup,
  type HeaderFooter,
  type Hyperlink,
  type PageBreak,
  type PageSettings,
  type PrintOptions,
  type Row,
  type SharedString,
  type Sheet,
  type SheetPageSetup,
  type SheetView,
  type SortCondition,
  type SortState,
} from '../../model/sheet/types';
import { childElements, localName, parseXml, serializeXml, type XmlElement } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { decodeXstring, parseColor, parseStringItem } from './shared-strings';
import { parseColor as parseColorElement, parseDifferentialFormat } from './styles';

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

const isXmlSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * Whether the text is an xsd:double as files write it: an optional sign, digits with an
 * optional point (at least one digit), an optional exponent, white space around. Number()
 * alone would also take "0x1F", "0b11" and "". Read by hand, one pass: a regular
 * expression for this backtracks quadratically over a long run of digits that ends in
 * something else.
 */
function isDecimal(text: string): boolean {
  const n = text.length;
  let i = 0;
  while (i < n && isXmlSpace(text.charCodeAt(i))) i++;
  let c = text.charCodeAt(i);
  if (c === 43 || c === 45) i++;
  let digits = 0;
  while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) {
    i++;
    digits++;
  }
  if (i < n && text.charCodeAt(i) === 46) {
    i++;
    while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) {
      i++;
      digits++;
    }
  }
  if (digits === 0) return false;
  c = text.charCodeAt(i);
  if (c === 69 || c === 101) {
    i++;
    c = text.charCodeAt(i);
    if (c === 43 || c === 45) i++;
    const exponent = i;
    while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) i++;
    if (i === exponent) return false;
  }
  while (i < n && isXmlSpace(text.charCodeAt(i))) i++;
  return i === n;
}

function parseFinite(text: string | undefined): number | undefined {
  if (text === undefined || !isDecimal(text)) return undefined;
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

/** Entries of Sheet.unmodelled that are kept; a part with more is damaged or made to be large. */
const MAX_UNMODELLED = 256;

/** Notes the name of something in the part that the model has no field for (see Sheet.unmodelled). */
function unmodelled(sheet: Sheet, name: string): void {
  const list = sheet.unmodelled || (sheet.unmodelled = []);
  if (list.length < MAX_UNMODELLED) list.push(name);
}

/** The attributes of the current element that belong to no namespace and declare none, as the file has them. */
function plainAttrs(sc: XmlScanner): [string, string][] {
  const out: [string, string][] = [];
  for (const [name, value] of sc.attrs()) if (name.indexOf(':') < 0 && name !== 'xmlns' && out.length < 64) out.push([name, value]);
  return out;
}

/** Ranges of one sheet that are read; more than any sheet has. */
const MAX_PROTECTED_RANGES = 10000;

function parseProtectedRanges(sc: XmlScanner, sheet: Sheet): void {
  const ranges: NonNullable<NonNullable<Sheet['protection']>['ranges']> = [];
  eachChild(sc, () => {
    if (sc.local !== 'protectedRange' || ranges.length >= MAX_PROTECTED_RANGES) return sc.skip();
    const range: (typeof ranges)[number] = { attrs: plainAttrs(sc) };
    eachChild(sc, () => {
      if (sc.local !== 'securityDescriptor') return sc.skip();
      const text = sc.readText();
      if (!range.descriptors) range.descriptors = [];
      if (range.descriptors.length < 64) range.descriptors.push(text);
    });
    ranges.push(range);
  });
  if (ranges.length === 0) return;
  const protection = sheet.protection || (sheet.protection = {});
  if (!protection.ranges) protection.ranges = ranges;
}

function parseSheetPr(sc: XmlScanner, sheet: Sheet): void {
  // What the element says beside the colour of the tab and the page: the code name, how the sheet is published, ...
  for (const [name] of sc.attrs()) if (name !== 'filterMode' && name.slice(0, 5) !== 'xmlns') unmodelled(sheet, name);
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local !== 'tabColor' && sc.local !== 'pageSetUpPr') unmodelled(sheet, sc.local);
    if (sc.local === 'tabColor') {
      const color = parseColor(sc);
      if (color) sheet.tabColor = color;
    } else if (sc.local === 'pageSetUpPr') {
      const fit = parseFlag(sc.attr('fitToPage'));
      if (fit !== undefined) pageSetupOf(sheet).fitToPage = fit;
      const breaks = parseFlag(sc.attr('autoPageBreaks'));
      if (breaks !== undefined) pageSetupOf(sheet).autoPageBreaks = breaks;
      sc.skip();
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
  // Metadata indexes and the phonetic flag are rare; they are only looked for in a tag that has more
  // attributes than the ones found so far, and only added to the cells that have them.
  let cm = -1;
  let vm = -1;
  let ph = false;
  if (sc.attrCount > (ref === undefined ? 0 : 1) + (styleAttr === undefined ? 0 : 1) + (type === undefined ? 0 : 1)) {
    cm = parseUint(sc.attr('cm'));
    vm = parseUint(sc.attr('vm'));
    ph = parseBool(sc.attr('ph'), false);
  }

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
      cell.si = index;
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
  if (cm >= 0) cell.cm = cm;
  if (vm >= 0) cell.vm = vm;
  if (ph) cell.ph = true;

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

/** Most distinct attribute lists kept for sharing between rows; a sheet normally has one or two. */
const MAX_SHARED_ROW_ATTRS = 256;

function parseSheetData(sc: XmlScanner, shared: SharedString[], sheet: Sheet, state: { prevRow: number }): void {
  if (sc.selfClosing) return;
  const rows = sheet.rows;
  // Nearly every row of a file carries the same extra attributes (x14ac:dyDescent); equal lists are one object.
  const attrLists = new Map<string, [string, string][]>();
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    if (sc.local !== 'row') {
      sc.skip();
      continue;
    }

    let r: string | undefined;
    let ht: string | undefined;
    let s: string | undefined;
    let customFormat = false;
    let customHeight = false;
    let hidden = false;
    let collapsed = false;
    let outlineLevel = 0;
    let extra: [string, string][] | undefined;
    for (const pair of sc.attrs()) {
      const value = pair[1];
      switch (pair[0]) {
        case 'r':
          r = value;
          break;
        case 'spans':
          // A hint for readers; not kept, it is stale after any edit.
          break;
        case 's':
          s = value;
          break;
        case 'customFormat':
          customFormat = parseBool(value, false);
          break;
        case 'ht':
          ht = value;
          break;
        case 'customHeight':
          customHeight = parseBool(value, false);
          break;
        case 'hidden':
          hidden = parseBool(value, false);
          break;
        case 'outlineLevel':
          outlineLevel = Math.max(parseUint(value), 0);
          break;
        case 'collapsed':
          collapsed = parseBool(value, false);
          break;
        default:
          if (extra) extra.push(pair);
          else extra = [pair];
      }
    }

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
      row = { customHeight, hidden, outlineLevel, collapsed, cells: [] };
      rows.set(rowIndex, row);
    } else {
      row.customHeight = customHeight;
      row.hidden = hidden;
      row.outlineLevel = outlineLevel;
      row.collapsed = collapsed;
      delete row.attrs;
    }
    const height = parseFinite(ht);
    if (height !== undefined && height >= 0) row.height = height;
    if (customFormat) {
      const style = parseUint(s);
      if (style >= 0) row.style = style;
    } else if (s !== undefined) {
      // Without customFormat the index is not the row's format; it is kept only to be written back.
      if (extra) extra.unshift(['s', s]);
      else extra = [['s', s]];
    }
    if (extra) {
      let key = '';
      for (let i = 0; i < extra.length; i++) key += extra[i][0] + '\u0000' + extra[i][1] + '\u0000';
      const known = attrLists.get(key);
      if (known) extra = known;
      else if (attrLists.size < MAX_SHARED_ROW_ATTRS) attrLists.set(key, extra);
      row.attrs = extra;
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

// ---------------------------------------------------------------------------
// Conditional formats
//
// The elements are small, so each one is read as a tree. Nothing here is written
// back: the writer copies the source elements.

const firstElement = (parent: XmlElement, local: string): XmlElement | undefined => childElements(parent, local)[0];

function elementText(el: XmlElement): string {
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

/** One reference of an sqref list: "A1", "A1:B2", and the whole-column and whole-row forms "A:C" and "1:5". */
function parseSqrefItem(ref: string): CellRange | undefined {
  const range = parseRange(ref);
  if (range) return range;
  const cols = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(ref);
  if (cols) {
    const a = parseCellRef(cols[1] + '1');
    const b = parseCellRef(cols[2] + '1');
    if (a && b) return { top: 0, left: Math.min(a.col, b.col), bottom: MAX_ROWS - 1, right: Math.max(a.col, b.col) };
  }
  const rows = /^\$?(\d{1,7}):\$?(\d{1,7})$/.exec(ref);
  if (rows) {
    const a = Number(rows[1]);
    const b = Number(rows[2]);
    if (a >= 1 && b >= 1 && a <= MAX_ROWS && b <= MAX_ROWS) return { top: Math.min(a, b) - 1, left: 0, bottom: Math.max(a, b) - 1, right: MAX_COLS - 1 };
  }
  return undefined;
}

/** The ranges of an sqref list; references that are not ranges inside the sheet are left out. */
function parseSqref(text: string | undefined): CellRange[] {
  const ranges: CellRange[] = [];
  if (!text) return ranges;
  const items = text.trim().split(/\s+/);
  for (let i = 0; i < items.length; i++) {
    const range = items[i] ? parseSqrefItem(items[i]) : undefined;
    if (range) ranges.push(range);
  }
  return ranges;
}

const flagAttr = (el: XmlElement, name: string, fallback: boolean) => parseBool(el.attrs.get(name), fallback);

/** A <cfvo>: its value is the "val" attribute, or an <xm:f> child in the extension list. */
function parseConditionalValue(el: XmlElement): ConditionalValue {
  const value: ConditionalValue = { type: el.attrs.get('type') || 'num', gte: flagAttr(el, 'gte', true) };
  const formula = firstElement(el, 'f');
  const val = formula ? elementText(formula) : el.attrs.get('val');
  if (val !== undefined) value.value = val;
  return value;
}

const conditionalValues = (el: XmlElement) => childElements(el, 'cfvo').map(parseConditionalValue);

function parseColorScale(el: XmlElement): ConditionalColorScale {
  const colors: ConditionalColorScale['colors'] = [];
  for (const child of childElements(el, 'color')) colors.push(parseColorElement(child) || {});
  return { values: conditionalValues(el), colors };
}

const percentAttr = (el: XmlElement, name: string, fallback: number): number => {
  const value = parseFinite(el.attrs.get(name));
  return value !== undefined && value >= 0 && value <= 100 ? value : fallback;
};

/** Reads the properties only <x14:dataBar> states into `bar`. */
function readDataBarExtension(el: XmlElement, bar: ConditionalDataBar): void {
  bar.extended = true;
  // UNSURE: the schema of the extension gives the same defaults as the 2007 element, 10 and 90.
  // Excel always writes both (0 and 100 for a bar made in 2010 or later), so the defaults only
  // matter for files of other producers.
  bar.minLength = percentAttr(el, 'minLength', 10);
  bar.maxLength = percentAttr(el, 'maxLength', 90);
  bar.gradient = flagAttr(el, 'gradient', true);
  bar.negativeFillSameAsPositive = flagAttr(el, 'negativeBarColorSameAsPositive', false);
  bar.negativeBorderSameAsPositive = flagAttr(el, 'negativeBarBorderColorSameAsPositive', true);
  const axis = el.attrs.get('axisPosition');
  bar.axisPosition = axis === 'middle' || axis === 'none' ? axis : 'automatic';
  const direction = el.attrs.get('direction');
  bar.direction = direction === 'leftToRight' || direction === 'rightToLeft' ? direction : 'context';
  const values = conditionalValues(el);
  if (values[0]) bar.min = values[0];
  if (values[1]) bar.max = values[1];
  const fill = parseColorElement(firstElement(el, 'fillColor'));
  if (fill) bar.color = fill;
  // The border colour only counts when the bar is said to have a border.
  const border = flagAttr(el, 'border', false) ? parseColorElement(firstElement(el, 'borderColor')) : undefined;
  if (border) bar.borderColor = border;
  const negativeFill = parseColorElement(firstElement(el, 'negativeFillColor'));
  if (negativeFill) bar.negativeFillColor = negativeFill;
  const negativeBorder = parseColorElement(firstElement(el, 'negativeBorderColor'));
  if (negativeBorder) bar.negativeBorderColor = negativeBorder;
  const axisColor = parseColorElement(firstElement(el, 'axisColor'));
  if (axisColor) bar.axisColor = axisColor;
}

function parseDataBar(el: XmlElement, extension: boolean): ConditionalDataBar {
  const values = conditionalValues(el);
  const bar: ConditionalDataBar = {
    min: values[0] || { type: 'min', gte: true },
    max: values[1] || { type: 'max', gte: true },
    showValue: flagAttr(el, 'showValue', true),
    minLength: percentAttr(el, 'minLength', 10),
    maxLength: percentAttr(el, 'maxLength', 90),
    extended: false,
    gradient: true,
    negativeFillSameAsPositive: false,
    negativeBorderSameAsPositive: true,
    axisPosition: 'automatic',
    direction: 'context',
  };
  const color = parseColorElement(firstElement(el, 'color'));
  if (color) bar.color = color;
  if (extension) readDataBarExtension(el, bar);
  return bar;
}

function parseIconSet(el: XmlElement): ConditionalIconSet {
  const icons: ConditionalIconSet = {
    set: el.attrs.get('iconSet') || '3TrafficLights1',
    values: conditionalValues(el),
    showValue: flagAttr(el, 'showValue', true),
    reverse: flagAttr(el, 'reverse', false),
  };
  const custom = childElements(el, 'cfIcon');
  if (custom.length > 0) {
    icons.custom = custom.map((icon) => ({ set: icon.attrs.get('iconSet') || '', id: Math.max(parseUint(icon.attrs.get('iconId')), 0) }));
  }
  return icons;
}

/** Reads a <cfRule>; `extension` is set for an <x14:cfRule>. The element of a data bar is handed back for linking. */
function parseConditionalRule(el: XmlElement, extension: boolean): ConditionalRule {
  const priority = parseUint(el.attrs.get('priority'));
  const rule: ConditionalRule = {
    type: el.attrs.get('type') || '',
    priority: priority >= 0 ? priority : Infinity,
    stopIfTrue: flagAttr(el, 'stopIfTrue', false),
    formulas: [],
    percent: flagAttr(el, 'percent', false),
    bottom: flagAttr(el, 'bottom', false),
    aboveAverage: flagAttr(el, 'aboveAverage', true),
    equalAverage: flagAttr(el, 'equalAverage', false),
  };
  const dxfId = parseUint(el.attrs.get('dxfId'));
  if (dxfId >= 0) rule.dxfId = dxfId;
  const operator = el.attrs.get('operator');
  if (operator !== undefined) rule.operator = operator;
  const text = el.attrs.get('text');
  if (text !== undefined) rule.text = text;
  const timePeriod = el.attrs.get('timePeriod');
  if (timePeriod !== undefined) rule.timePeriod = timePeriod;
  const rank = parseUint(el.attrs.get('rank'));
  if (rank >= 0) rule.rank = rank;
  const stdDev = parseUint(el.attrs.get('stdDev'));
  if (stdDev >= 0) rule.stdDev = stdDev;
  if (extension) {
    rule.extended = true;
    const id = el.attrs.get('id');
    if (id !== undefined) rule.extId = id;
  }

  for (const child of el.children) {
    if (child.type !== 'element') continue;
    switch (localName(child.name)) {
      // <formula> in the main namespace, <xm:f> in the extension list. Taken literally, as cell formulas are.
      case 'formula':
      case 'f':
        rule.formulas.push(elementText(child));
        break;
      case 'colorScale':
        rule.colorScale = parseColorScale(child);
        break;
      case 'dataBar':
        rule.dataBar = parseDataBar(child, extension);
        break;
      case 'iconSet':
        rule.iconSet = parseIconSet(child);
        break;
      case 'dxf':
        rule.dxf = parseDifferentialFormat(child);
        break;
      case 'extLst':
        // <ext><x14:id>{...}</x14:id></ext> names the rule's counterpart in the extension list.
        for (const ext of childElements(child, 'ext')) {
          const id = firstElement(ext, 'id');
          if (id) rule.extId = elementText(id).trim();
        }
        break;
    }
  }
  return rule;
}

/** With a <conditionalFormatting> open tag current (either namespace): reads it and consumes the element. */
function parseConditionalFormatting(sc: XmlScanner, extension: boolean, index = -1): ConditionalFormat {
  const el = parseXml(sc.readRaw()).root;
  // The range is an attribute in the main namespace and an <xm:sqref> child in the extension list.
  const sqref = firstElement(el, 'sqref');
  const format: ConditionalFormat = { ranges: parseSqref(sqref ? elementText(sqref) : el.attrs.get('sqref')), rules: [] };
  // Where the element and its rules are in the part: how the writer finds them again.
  if (index >= 0) format.src = index;
  childElements(el, 'cfRule').forEach((child, at) => {
    const rule = parseConditionalRule(child, extension);
    if (index >= 0) rule.src = [index, at];
    format.rules.push(rule);
  });
  return format;
}

// ---------------------------------------------------------------------------
// Filter and sort state
//
// Read as a tree like the conditional formats. Unlike them the filter is kept up to
// date by edits, and written from the model once it was changed; what the model has
// no field for is kept as source text beside the fields.

const sourceText = (root: XmlElement) => serializeXml({ declaration: null, leading: [], root, trailing: [] });

/** The attributes of an element that are not in `known`, in document order. */
function otherAttrs(el: XmlElement, known: string[]): [string, string][] | undefined {
  let out: [string, string][] | undefined;
  for (const [name, value] of el.attrs) {
    if (known.indexOf(name) >= 0) continue;
    if (out) out.push([name, value]);
    else out = [[name, value]];
  }
  return out;
}

function parseDateGroup(el: XmlElement): FilterDateGroup | undefined {
  const year = parseUint(el.attrs.get('year'));
  if (year < 0) return undefined;
  const group: FilterDateGroup = { grouping: el.attrs.get('dateTimeGrouping') || 'year', year };
  for (const part of ['month', 'day', 'hour', 'minute', 'second'] as const) {
    const value = parseUint(el.attrs.get(part));
    if (value >= 0) group[part] = value;
  }
  return group;
}

/** Reads a <filterColumn>; undefined when it names no column of the sheet. */
function parseFilterColumn(el: XmlElement): FilterColumn | undefined {
  const colId = parseUint(el.attrs.get('colId'));
  if (colId < 0 || colId >= MAX_COLS) return undefined;
  const column: FilterColumn = { colId, kind: 'none', raw: sourceText(el) };
  if (flagAttr(el, 'hiddenButton', false) || !flagAttr(el, 'showButton', true)) column.hiddenButton = true;
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    const local = localName(child.name);
    // More than one criterion is not valid; what comes after the first is not understood.
    if (local !== 'extLst' && column.kind !== 'none') {
      column.kind = 'unknown';
      continue;
    }
    switch (local) {
      case 'filters': {
        column.kind = 'values';
        column.values = [];
        if (flagAttr(child, 'blank', false)) column.blank = true;
        for (const item of child.children) {
          if (item.type !== 'element') continue;
          const name = localName(item.name);
          if (name === 'filter') {
            const value = item.attrs.get('val');
            if (value !== undefined) column.values.push(decodeXstring(value));
          } else if (name === 'dateGroupItem') {
            const group = parseDateGroup(item);
            if (group) (column.dateGroups || (column.dateGroups = [])).push(group);
          }
        }
        break;
      }
      case 'customFilters': {
        column.kind = 'custom';
        column.and = flagAttr(child, 'and', false);
        const conditions: FilterCondition[] = [];
        for (const item of childElements(child, 'customFilter')) {
          conditions.push({ operator: item.attrs.get('operator') || 'equal', value: decodeXstring(item.attrs.get('val') || '') });
        }
        column.conditions = conditions;
        break;
      }
      case 'top10': {
        const value = parseFinite(child.attrs.get('val'));
        column.kind = value === undefined ? 'unknown' : 'top10';
        column.top = flagAttr(child, 'top', true);
        column.percent = flagAttr(child, 'percent', false);
        if (value !== undefined) column.value = value;
        break;
      }
      case 'dynamicFilter':
        column.kind = 'dynamic';
        column.dynamicType = child.attrs.get('type') || '';
        break;
      case 'colorFilter':
        column.kind = 'color';
        break;
      case 'iconFilter':
        column.kind = 'icon';
        break;
      case 'extLst':
        break;
      default:
        column.kind = 'unknown';
    }
  }
  return column;
}

function parseSortState(el: XmlElement): SortState | undefined {
  const range = parseRange(el.attrs.get('ref') || '');
  if (!range) return undefined;
  const state: SortState = { range, conditions: [] };
  if (flagAttr(el, 'caseSensitive', false)) state.caseSensitive = true;
  if (flagAttr(el, 'columnSort', false)) state.columnSort = true;
  const attrs = otherAttrs(el, ['ref', 'caseSensitive', 'columnSort']);
  if (attrs) state.attrs = attrs;
  let extra = '';
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    if (localName(child.name) !== 'sortCondition') {
      extra += sourceText(child);
      continue;
    }
    const ref = parseRange(child.attrs.get('ref') || '');
    if (!ref) continue;
    const condition: SortCondition = { range: ref, descending: flagAttr(child, 'descending', false) };
    // Anything beyond the two attributes the model holds: the element is kept to be written back.
    if (child.attrs.size > (child.attrs.has('descending') ? 2 : 1) || child.children.length > 0) condition.raw = sourceText(child);
    state.conditions.push(condition);
  }
  if (extra !== '') state.extra = extra;
  return state;
}

/** With an <autoFilter> open tag current: reads the filter and consumes the element. */
function parseAutoFilter(sc: XmlScanner): AutoFilter | undefined {
  const el = parseXml(sc.readRaw()).root;
  const range = parseRange(el.attrs.get('ref') || '');
  if (!range) return undefined;
  const filter: AutoFilter = { range, columns: [] };
  const attrs = otherAttrs(el, ['ref']);
  if (attrs) filter.attrs = attrs;
  const seen = new Set<number>();
  let extra = '';
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    const local = localName(child.name);
    if (local === 'filterColumn') {
      const column = parseFilterColumn(child);
      // A column stated twice, or one that is not a column: kept as text, never applied.
      if (column && !seen.has(column.colId)) {
        seen.add(column.colId);
        filter.columns.push(column);
        continue;
      }
    } else if (local === 'sortState' && !filter.sortState) {
      const state = parseSortState(child);
      if (state) {
        filter.sortState = state;
        continue;
      }
    }
    extra += sourceText(child);
  }
  if (extra !== '') filter.extra = extra;
  return filter;
}

// ---------------------------------------------------------------------------
// Data validations
//
// A view for reading, like the conditional formats: the writer copies the source elements.

/** Most validations read from one sheet; the format allows 65,534. */
const MAX_VALIDATIONS = 65534;

function parseDataValidation(el: XmlElement, extension: boolean): DataValidation | undefined {
  // The range is an attribute in the main namespace and an <xm:sqref> child in the extension list.
  const sqref = firstElement(el, 'sqref');
  const ranges = parseSqref(sqref ? elementText(sqref) : el.attrs.get('sqref'));
  if (ranges.length === 0) return undefined;
  const style = el.attrs.get('errorStyle');
  const rule: DataValidation = {
    ranges,
    type: el.attrs.get('type') || 'none',
    operator: el.attrs.get('operator') || 'between',
    allowBlank: flagAttr(el, 'allowBlank', false),
    hideDropDown: flagAttr(el, 'showDropDown', false),
    showInputMessage: flagAttr(el, 'showInputMessage', false),
    showErrorMessage: flagAttr(el, 'showErrorMessage', false),
    errorStyle: style === 'warning' || style === 'information' ? style : 'stop',
  };
  for (const name of ['promptTitle', 'prompt', 'errorTitle', 'error'] as const) {
    const value = el.attrs.get(name);
    // Like cell text, these may hold characters written as _xHHHH_.
    if (value !== undefined) rule[name] = decodeXstring(value);
  }
  for (const name of ['formula1', 'formula2'] as const) {
    const child = firstElement(el, name);
    if (!child) continue;
    // <formula1>text</formula1>, or <x14:formula1><xm:f>text</xm:f></x14:formula1>.
    const inner = firstElement(child, 'f');
    rule[name] = elementText(inner || child);
  }
  if (extension) rule.extended = true;
  return rule;
}

/** With a <dataValidations> open tag current (either namespace): reads its rules and consumes the element. */
function parseDataValidations(sc: XmlScanner, out: DataValidation[], extension: boolean, seen?: { count: number }): void {
  eachChild(sc, () => {
    if (sc.local !== 'dataValidation') return sc.skip();
    // Every element counts, read or not: the number is how the writer finds the element again.
    const index = seen ? seen.count++ : -1;
    if (out.length >= MAX_VALIDATIONS) return sc.skip();
    const rule = parseDataValidation(parseXml(sc.readRaw()).root, extension);
    if (!rule) return;
    if (index >= 0) rule.src = index;
    out.push(rule);
  });
}

// ---------------------------------------------------------------------------
// Page setup
//
// A view for reading, like the conditional formats: the writer copies the source
// elements, with the attributes the model has no field for (the relationship of the
// printer settings among them). Of an element that is stated twice the first one counts.

const pageSetupOf = (sheet: Sheet): SheetPageSetup => sheet.pageSetup || (sheet.pageSetup = {});

/** An xsd:boolean, or undefined for an attribute that is missing or is none. */
function parseFlag(text: string | undefined): boolean | undefined {
  if (text === '1' || text === 'true') return true;
  if (text === '0' || text === 'false') return false;
  return undefined;
}

/** An xsd:unsignedInt, or -1. */
function parseUint32(text: string | undefined): number {
  if (text === undefined || text.length === 0 || text.length > 10) return -1;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c < 48 || c > 57) return -1;
  }
  const value = Number(text);
  return value <= 0xffffffff ? value : -1;
}

/** The margins applications give a new sheet; they stand in for an attribute that is missing or is no margin. */
const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

/** With a <pageMargins> open tag current. */
function parsePageMargins(sc: XmlScanner): SheetPageSetup['margins'] {
  const margins = { ...DEFAULT_MARGINS };
  for (const name of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
    const value = parseFinite(sc.attr(name));
    if (value !== undefined && value >= 0) margins[name] = value;
  }
  return margins;
}

/** A length with its unit, as the size of a paper without a code is stated ("210mm", "8.5in"). */
const PAPER_MEASURE = /^[0-9]{1,9}(\.[0-9]{1,9})?(mm|cm|in|pt|pc|pi)$/;

/** With a <pageSetup> open tag current. Attributes that hold no value of their type state nothing. */
function parsePageSettings(sc: XmlScanner): PageSettings {
  const page: PageSettings = {};
  for (const [name, value] of sc.attrs()) {
    switch (name) {
      case 'paperSize':
      case 'fitToWidth':
      case 'fitToHeight':
      case 'firstPageNumber':
      case 'horizontalDpi':
      case 'verticalDpi':
      case 'copies': {
        const number = parseUint32(value);
        if (number >= 0) page[name] = number;
        break;
      }
      case 'scale': {
        const scale = parseUint32(value);
        if (scale >= 10 && scale <= 400) page.scale = scale;
        break;
      }
      case 'paperWidth':
      case 'paperHeight':
        if (PAPER_MEASURE.test(value)) page[name] = value;
        break;
      case 'orientation':
        if (value === 'default' || value === 'portrait' || value === 'landscape') page.orientation = value;
        break;
      case 'pageOrder':
        if (value === 'downThenOver' || value === 'overThenDown') page.pageOrder = value;
        break;
      case 'cellComments':
        if (value === 'none' || value === 'asDisplayed' || value === 'atEnd') page.cellComments = value;
        break;
      case 'errors':
        if (value === 'displayed' || value === 'blank' || value === 'dash' || value === 'NA') page.errors = value;
        break;
      case 'useFirstPageNumber':
      case 'blackAndWhite':
      case 'draft':
      case 'usePrinterDefaults': {
        const flag = parseFlag(value);
        if (flag !== undefined) page[name] = flag;
        break;
      }
    }
  }
  return page;
}

/** With a <printOptions> open tag current. */
function parsePrintOptions(sc: XmlScanner): PrintOptions {
  const print: PrintOptions = {};
  for (const name of ['horizontalCentered', 'verticalCentered', 'headings', 'gridLines', 'gridLinesSet'] as const) {
    const flag = parseFlag(sc.attr(name));
    if (flag !== undefined) print[name] = flag;
  }
  return print;
}

/** With a <headerFooter> open tag current: reads the texts and consumes the element. */
function parseHeaderFooter(sc: XmlScanner): HeaderFooter {
  const out: HeaderFooter = {};
  for (const name of ['differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins'] as const) {
    const flag = parseFlag(sc.attr(name));
    if (flag !== undefined) out[name] = flag;
  }
  eachChild(sc, () => {
    const local = sc.local;
    if (local !== 'oddHeader' && local !== 'oddFooter' && local !== 'evenHeader' && local !== 'evenFooter' && local !== 'firstHeader' && local !== 'firstFooter') return sc.skip();
    // Like cell text, these may hold characters written as _xHHHH_. An empty one is none.
    const text = decodeXstring(sc.readText());
    if (text !== '' && out[local] === undefined) out[local] = text;
  });
  return out;
}

/**
 * With a <rowBreaks> or <colBreaks> open tag current: reads the breaks and consumes the
 * element. A break that is not inside the sheet, and a second one at the same place, is left out.
 */
function parsePageBreaks(sc: XmlScanner, rows: boolean): PageBreak[] {
  const limit = rows ? MAX_ROWS : MAX_COLS;
  const crossLimit = rows ? MAX_COLS : MAX_ROWS;
  const breaks: PageBreak[] = [];
  const seen = new Set<number>();
  let sorted = true;
  eachChild(sc, () => {
    if (sc.local === 'brk') {
      const at = parseUint(sc.attr('id'));
      const min = Math.max(parseUint(sc.attr('min')), 0);
      const max = Math.max(parseUint(sc.attr('max')), 0);
      if (at >= 1 && at < limit && min <= max && max < crossLimit && !seen.has(at)) {
        seen.add(at);
        if (breaks.length > 0 && breaks[breaks.length - 1].at > at) sorted = false;
        const item: PageBreak = { at, min, max, manual: parseBool(sc.attr('man'), false) };
        if (parseBool(sc.attr('pt'), false)) item.pivot = true;
        breaks.push(item);
      }
    }
    sc.skip();
  });
  if (!sorted) breaks.sort((a, b) => a.at - b.at);
  return breaks;
}

/**
 * With an open tag current: calls `visit` at the open tag of every child element and
 * consumes the element. `visit` has to consume the child it is called for (skip() does).
 */
function eachChild(sc: XmlScanner, visit: () => void): void {
  if (sc.selfClosing) return;
  while (sc.next()) {
    if (sc.kind === 'close') return;
    if (sc.kind === 'open') visit();
  }
}

/** With the worksheet's <extLst> open tag current: reads the conditional formats and the validations of the 2010 extension. */
function parseSheetExtensions(sc: XmlScanner, extended: ConditionalFormat[], validations: DataValidation[], sheet: Sheet): void {
  eachChild(sc, () => {
    if (sc.local !== 'ext') return sc.skip();
    eachChild(sc, () => {
      if (sc.local === 'dataValidations') return parseDataValidations(sc, validations, true);
      if (sc.local !== 'conditionalFormattings') {
        // Sparklines, slicers, ...: kept in the source part.
        unmodelled(sheet, sc.local);
        return sc.skip();
      }
      eachChild(sc, () => {
        if (sc.local === 'conditionalFormatting') extended.push(parseConditionalFormatting(sc, true));
        else sc.skip();
      });
    });
  });
}

/**
 * Joins the rules of the extension list with the ones in the main part. A data bar is
 * written twice: once as Excel 2007 reads it and once with what later versions added,
 * tied by an id. Rules that exist only in the extension list are added as they are.
 */
function linkConditionalFormats(main: ConditionalFormat[], extended: ConditionalFormat[]): ConditionalFormat[] {
  if (extended.length === 0) return main;
  const byId = new Map<string, ConditionalRule>();
  for (const format of main) {
    for (const rule of format.rules) if (rule.extId !== undefined && !byId.has(rule.extId)) byId.set(rule.extId, rule);
  }
  const out = main.slice();
  for (const format of extended) {
    const own: ConditionalRule[] = [];
    for (const rule of format.rules) {
      const base = rule.extId === undefined ? undefined : byId.get(rule.extId);
      if (!base) {
        own.push(rule);
        continue;
      }
      if (base.dataBar && rule.dataBar) {
        // The colour of the bar is stated by the main rule; the extension only repeats it sometimes.
        const color = base.dataBar.color || rule.dataBar.color;
        const showValue = base.dataBar.showValue && rule.dataBar.showValue;
        base.dataBar = { ...rule.dataBar, showValue };
        if (color) base.dataBar.color = color;
      } else if (base.iconSet && rule.iconSet) {
        base.iconSet = rule.iconSet;
      }
    }
    if (own.length > 0 && format.ranges.length > 0) out.push({ ranges: format.ranges, rules: own });
  }
  return out;
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
  const conditional: ConditionalFormat[] = [];
  const extended: ConditionalFormat[] = [];
  const validations: DataValidation[] = [];
  let formatsSeen = 0;
  const validationsSeen = { count: 0 };
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
          // How deep the rows and columns are grouped is not in the model.
          if (/^[1-9]/.test(sc.attr('outlineLevelRow') || '') || /^[1-9]/.test(sc.attr('outlineLevelCol') || '')) unmodelled(sheet, 'outlineLevels');
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
        case 'conditionalFormatting':
          conditional.push(parseConditionalFormatting(sc, false, formatsSeen++));
          break;
        case 'autoFilter': {
          const filter = parseAutoFilter(sc);
          if (filter && !sheet.autoFilter) sheet.autoFilter = filter;
          break;
        }
        case 'dataValidations':
          parseDataValidations(sc, validations, false, validationsSeen);
          break;
        case 'extLst':
          parseSheetExtensions(sc, extended, validations, sheet);
          break;
        case 'pageMargins':
          if (!pageSetupOf(sheet).margins) pageSetupOf(sheet).margins = parsePageMargins(sc);
          sc.skip();
          break;
        case 'pageSetup':
          if (!pageSetupOf(sheet).page) pageSetupOf(sheet).page = parsePageSettings(sc);
          sc.skip();
          break;
        case 'printOptions':
          if (!pageSetupOf(sheet).print) pageSetupOf(sheet).print = parsePrintOptions(sc);
          sc.skip();
          break;
        case 'headerFooter':
          if (pageSetupOf(sheet).headerFooter) sc.skip();
          else pageSetupOf(sheet).headerFooter = parseHeaderFooter(sc);
          break;
        case 'rowBreaks':
        case 'colBreaks': {
          const rows = sc.local === 'rowBreaks';
          const setup = pageSetupOf(sheet);
          if (rows ? setup.rowBreaks : setup.colBreaks) sc.skip();
          else {
            const breaks = parsePageBreaks(sc, rows);
            // An element without a break says nothing.
            if (breaks.length > 0) setup[rows ? 'rowBreaks' : 'colBreaks'] = breaks;
            else if (Object.keys(setup).length === 0) delete sheet.pageSetup;
          }
          break;
        }
        case 'sheetProtection': {
          const protection = sheet.protection || (sheet.protection = {});
          if (!protection.attrs) protection.attrs = plainAttrs(sc);
          sc.skip();
          break;
        }
        case 'protectedRanges':
          parseProtectedRanges(sc, sheet);
          break;
        case 'legacyDrawingHF':
        case 'picture':
          // The pictures of headers and footers and the picture behind the cells: their parts are of kinds the
          // drawings and the notes have too (a VML drawing, an image), so it is the element that says they are there.
          unmodelled(sheet, sc.local);
          sc.skip();
          break;
        case 'drawing':
        case 'legacyDrawing':
        case 'tableParts':
        case 'oleObjects':
        case 'controls':
          // What these lead to is in parts of its own, and is known by them (Sheet.anchored, the drawings, the comments).
          sc.skip();
          break;
        default:
          // Kept in the source part; its name is noted for what a copy of the sheet has to say.
          unmodelled(sheet, sc.local);
          sc.skip();
      }
    }
  }
  const formats = linkConditionalFormats(conditional, extended).filter((format) => format.ranges.length > 0 && format.rules.length > 0);
  if (formats.length > 0) sheet.conditionalFormats = formats;
  if (validations.length > 0) sheet.dataValidations = validations;
  // Anything after the root is still checked for well-formedness.
  while (sc.next());
}
