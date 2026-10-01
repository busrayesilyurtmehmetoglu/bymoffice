// Writes an edited spreadsheet model back into its xlsx package.
//
// Only what changed is produced again. A sheet that was not modified keeps its
// part untouched. In a modified sheet, the rows that were not edited and every
// element the model does not hold are copied from the original text as they
// are; rows that were edited are written from the model. Parts are always
// regenerated from the bytes the file had when it was opened, never from the
// output of an earlier save: the model records its changes relative to those.

import { adjustAnchoredFormula, adjustFormula, adjustRange, adjustReferenceText, type StructuralEdit } from '../../calc/refs';
import {
  MAX_COLS,
  MAX_ROWS,
  cellName,
  columnName,
  getCell,
  parseCellRef,
  parseRange,
  rangeName,
  type Alignment,
  type Border,
  type BorderEdge,
  type Cell,
  type CellFormat,
  type CellRange,
  type CellFormula,
  type Color,
  type ColumnRange,
  type Fill,
  type Font,
  type Protection,
  type RichRun,
  type Row,
  type SharedString,
  type Sheet,
  type SheetStructuralEdit,
  type StyleSheet,
  type Workbook,
} from '../../model/sheet/types';
import { cellDisplayText } from '../../format/cell-format';
import type { OpcPackage } from '../../opc';
import { localName, parseXml, serializeXml, type XmlDocument, type XmlElement, type XmlNode } from '../../xml';
import { XmlScanner } from '../../xml-scan';

const encoder = new TextEncoder();

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_REL_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';

// --- escaping ----------------------------------------------------------------

/** Characters XML 1.0 cannot carry at all, and halves of surrogate pairs (checked for a partner separately). */
const NEEDS_CARE = /[\u0000-\u0008\u000B-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/;

const isHex = (c: number) => (c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102);

/**
 * Kind of the UTF-16 unit at `i`: 0 an ordinary character, 1 the first half of
 * a valid surrogate pair (the next unit belongs to it), 2 a character that
 * cannot be written as it is (control character, U+FFFE, U+FFFF, lone surrogate).
 */
function unitKind(text: string, i: number): 0 | 1 | 2 {
  const c = text.charCodeAt(i);
  if (c < 32) return c === 9 || c === 10 ? 0 : 2;
  if (c < 0xd800) return 0;
  if (c <= 0xdbff) {
    const next = text.charCodeAt(i + 1);
    return next >= 0xdc00 && next <= 0xdfff ? 1 : 2;
  }
  if (c <= 0xdfff || c === 0xfffe || c === 0xffff) return 2;
  return 0;
}

const hex4 = (code: number) => ('000' + code.toString(16).toUpperCase()).slice(-4);

/**
 * The inverse of decodeXstring(): characters XML cannot carry (and carriage
 * returns, which XML would turn into line feeds) become "_xHHHH_", and text
 * that already looks like such an escape gets its underscore escaped. The
 * result still needs XML escaping.
 */
export function encodeXstring(text: string): string {
  if (text.indexOf('_x') < 0 && !NEEDS_CARE.test(text)) return text;
  const n = text.length;
  let out = '';
  let copied = 0;
  for (let i = 0; i < n; i++) {
    const c = text.charCodeAt(i);
    if (c === 95) {
      // "_xHHHH" reads as an escape when an underscore follows, be it one of the text or the start of the next escape.
      if (
        text.charCodeAt(i + 1) === 120 &&
        i + 6 <= n &&
        isHex(text.charCodeAt(i + 2)) &&
        isHex(text.charCodeAt(i + 3)) &&
        isHex(text.charCodeAt(i + 4)) &&
        isHex(text.charCodeAt(i + 5)) &&
        i + 6 < n &&
        (text.charCodeAt(i + 6) === 95 || unitKind(text, i + 6) === 2)
      ) {
        out += text.slice(copied, i) + '_x005F_';
        copied = i + 1;
      }
      continue;
    }
    if (c >= 32 && c < 0xd800) continue;
    const kind = unitKind(text, i);
    if (kind === 1) i++;
    else if (kind === 2) {
      out += text.slice(copied, i) + '_x' + hex4(c) + '_';
      copied = i + 1;
    }
  }
  return copied === 0 ? text : out + text.slice(copied);
}

/** Replaces what XML cannot carry with U+FFFD, for text that has no escape of its own (formulas, attributes). */
function replaceUnwritable(text: string): string {
  if (!NEEDS_CARE.test(text)) return text;
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const kind = unitKind(text, i);
    if (kind === 1) out += text[i] + text[++i];
    else if (kind === 2 && text.charCodeAt(i) !== 13) out += '\uFFFD';
    else out += text[i];
  }
  return out;
}

const TEXT_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#13;' };
const ATTR_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '"': '&quot;', '\t': '&#9;', '\n': '&#10;', '\r': '&#13;' };

function escapeText(text: string): string {
  text = replaceUnwritable(text);
  return /[&<>\r]/.test(text) ? text.replace(/[&<>\r]/g, (ch) => TEXT_ESCAPES[ch]) : text;
}

function escapeAttr(text: string): string {
  text = replaceUnwritable(text);
  return /[&<"\t\n\r]/.test(text) ? text.replace(/[&<"\t\n\r]/g, (ch) => ATTR_ESCAPES[ch]) : text;
}

/** Element text of an ST_Xstring value: "_xHHHH_" escapes first, then XML escapes. */
function escapeXstring(text: string): string {
  const encoded = encodeXstring(text);
  return /[&<>]/.test(encoded) ? encoded.replace(/[&<>]/g, (ch) => TEXT_ESCAPES[ch]) : encoded;
}

// --- source text ----------------------------------------------------------------

/** The part as text, the way XmlScanner sees it: byte order mark removed, line endings as "\n". */
function decodeSource(bytes: Uint8Array): string {
  let s: string;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) s = new TextDecoder('utf-16le').decode(bytes);
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) s = new TextDecoder('utf-16be').decode(bytes);
  else s = new TextDecoder('utf-8').decode(bytes);
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
  if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n');
  return s;
}

const malformed = (): never => {
  throw new Error('Invalid worksheet: the part ends unexpectedly');
};

/**
 * Offset of the next start or end tag at or after `from`; text, comments,
 * processing instructions and CDATA sections before it are passed over. The
 * scanner walks the same text and reports what is wrong with it, so this only
 * has to find positions.
 */
function nextTag(s: string, from: number): number {
  let i = from;
  for (;;) {
    i = s.indexOf('<', i);
    if (i < 0) return malformed();
    const c = s.charCodeAt(i + 1);
    let end: number;
    if (c === 63) end = s.indexOf('?>', i) + 2;
    else if (c !== 33) return i;
    else if (s.startsWith('<!--', i)) end = s.indexOf('-->', i + 4) + 3;
    else if (s.startsWith('<![CDATA[', i)) end = s.indexOf(']]>', i + 9) + 3;
    else return malformed();
    if (end <= i) return malformed();
    i = end;
  }
}

/** Offset just past the ">" of the start tag at `i`. */
function startTagEnd(s: string, i: number): number {
  const n = s.length;
  let j = i + 1;
  while (j < n) {
    const c = s.charCodeAt(j);
    if (c === 62) return j + 1;
    if (c === 34 || c === 39) {
      j = s.indexOf(c === 34 ? '"' : "'", j + 1);
      if (j < 0) return malformed();
    }
    j++;
  }
  return malformed();
}

/** The output is UTF-8 whatever the source was; a declaration that names another encoding is corrected. */
function fixDeclaration(prolog: string): string {
  return prolog.replace(/^(<\?xml[^?]*?encoding\s*=\s*)(["'])([^"']*)\2/, (all, head: string, quote: string, name: string) =>
    /^utf-?8$/i.test(name) ? all : `${head}${quote}UTF-8${quote}`,
  );
}

/** Plain non-negative integer, or -1. */
function parseUint(text: string | undefined): number {
  return text !== undefined && /^[0-9]{1,9}$/.test(text) ? Number(text) : -1;
}

// --- cells and rows --------------------------------------------------------------

interface CellWriter {
  /** Namespace prefix of the sheet's elements with its colon, usually empty. */
  prefix: string;
  shared: SharedString[];
  /** Column ranges that give their cells a format other than the default one. */
  styledColumns: ColumnRange[];
  /**
   * Cells written in place of the model's, by row and column: the headers of
   * a table whose column names had to be made valid.
   */
  overrides?: Map<number, Map<number, Cell>>;
}

/** Whether a position without a cell would take a format from its column. */
function columnHasStyle(writer: CellWriter, col: number): boolean {
  for (const range of writer.styledColumns) {
    if (range.min > col) break;
    if (range.max >= col) return true;
  }
  return false;
}

const COLUMN_NAMES: string[] = [];

function columnLetters(col: number): string {
  return COLUMN_NAMES[col] || (COLUMN_NAMES[col] = columnName(col));
}

function colorAttrs(color: Color): string {
  let out = '';
  if (color.auto !== undefined) out += ` auto="${color.auto ? 1 : 0}"`;
  if (color.indexed !== undefined) out += ` indexed="${color.indexed}"`;
  if (color.rgb !== undefined) out += ` rgb="${escapeAttr(color.rgb)}"`;
  if (color.theme !== undefined) out += ` theme="${color.theme}"`;
  if (color.tint !== undefined) out += ` tint="${color.tint}"`;
  return out;
}

/**
 * <rPr> for a run font: the inverse of parseRunFont(). Children are in the
 * order CT_RPrElt requires. `prefix` is the namespace prefix with its colon.
 */
export function runPropertiesXml(font: Font, prefix = ''): string {
  let out = '';
  if (font.name !== undefined) out += `<${prefix}rFont val="${escapeAttr(font.name)}"/>`;
  if (font.charset !== undefined) out += `<${prefix}charset val="${font.charset}"/>`;
  if (font.family !== undefined) out += `<${prefix}family val="${font.family}"/>`;
  if (font.bold) out += `<${prefix}b/>`;
  if (font.italic) out += `<${prefix}i/>`;
  if (font.strike) out += `<${prefix}strike/>`;
  if (font.color) out += `<${prefix}color${colorAttrs(font.color)}/>`;
  if (font.size !== undefined) out += `<${prefix}sz val="${font.size}"/>`;
  if (font.underline === 'single') out += `<${prefix}u/>`;
  else if (font.underline && font.underline !== 'none') out += `<${prefix}u val="${font.underline}"/>`;
  if (font.vertAlign !== undefined) out += `<${prefix}vertAlign val="${font.vertAlign}"/>`;
  if (font.scheme !== undefined) out += `<${prefix}scheme val="${font.scheme}"/>`;
  return out === '' ? `<${prefix}rPr/>` : `<${prefix}rPr>${out}</${prefix}rPr>`;
}

function textElement(text: string, prefix: string): string {
  // Without the attribute a reader may trim the text or fold its line breaks.
  const preserve = /^\s|\s$|[\n\t]/.test(text) ? ' xml:space="preserve"' : '';
  return `<${prefix}t${preserve}>${escapeXstring(text)}</${prefix}t>`;
}

function inlineString(text: string, runs: RichRun[] | undefined, prefix: string): string {
  let body = '';
  if (runs && runs.length > 0) {
    let joined = '';
    for (const run of runs) joined += run.text;
    // Runs that no longer spell the cell's text belong to an earlier value.
    if (joined === text) {
      for (const run of runs) {
        body += `<${prefix}r>${run.font ? runPropertiesXml(run.font, prefix) : ''}${textElement(run.text, prefix)}</${prefix}r>`;
      }
    }
  }
  if (body === '') body = textElement(text, prefix);
  return `<${prefix}is>${body}</${prefix}is>`;
}

function formulaXml(formula: CellFormula, prefix: string): string {
  let out = `<${prefix}f`;
  if (formula.type !== 'normal') out += ` t="${formula.type}"`;
  if (formula.ref !== undefined) out += ` ref="${escapeAttr(formula.ref)}"`;
  if (formula.si !== undefined) out += ` si="${formula.si}"`;
  if (formula.attrs) for (const [name, value] of formula.attrs) out += ` ${name}="${escapeAttr(value)}"`;
  // Formula text is literal: "_x0041_" in a formula is a name, not an escape.
  return formula.text === '' ? out + '/>' : `${out}>${escapeText(formula.text)}</${prefix}f>`;
}

/**
 * One <c> element, or '' for a cell that says nothing: blank, default format,
 * no formula. Where the row or the column has a format (`inherits`), such a
 * cell does say something: that the position has the default format instead.
 */
function cellXml(writer: CellWriter, cell: Cell, rowNumber: number, col: number, inherits: boolean): string {
  const prefix = writer.prefix;
  const formula = cell.formula;
  if (cell.kind === 'blank' && !formula && !cell.style && cell.cm === undefined && cell.vm === undefined && !cell.ph) {
    return inherits || columnHasStyle(writer, col) ? `<${prefix}c r="${columnLetters(col)}${rowNumber}"/>` : '';
  }

  let out = `<${prefix}c r="${columnLetters(col)}${rowNumber}"`;
  if (cell.style) out += ` s="${cell.style}"`;
  if (cell.cm !== undefined) out += ` cm="${cell.cm}"`;
  if (cell.vm !== undefined) out += ` vm="${cell.vm}"`;
  if (cell.ph) out += ' ph="1"';

  let body = formula ? formulaXml(formula, prefix) : '';
  const value = cell.value;
  switch (cell.kind) {
    case 'number':
      if (typeof value === 'number' && isFinite(value)) body += `<${prefix}v>${String(value)}</${prefix}v>`;
      else {
        // Not a number a file can hold; what a spreadsheet shows for such a result.
        out += ' t="e"';
        body += `<${prefix}v>#NUM!</${prefix}v>`;
      }
      break;
    case 'boolean':
      out += ' t="b"';
      body += `<${prefix}v>${value ? 1 : 0}</${prefix}v>`;
      break;
    case 'error':
      out += ' t="e"';
      body += `<${prefix}v>${escapeText(String(value))}</${prefix}v>`;
      break;
    case 'date':
      out += ' t="d"';
      body += `<${prefix}v>${escapeText(String(value))}</${prefix}v>`;
      break;
    case 'string': {
      const text = value === null ? '' : String(value);
      const si = cell.si;
      if (formula) {
        out += ' t="str"';
        body += `<${prefix}v>${escapeXstring(text)}</${prefix}v>`;
      } else if (si !== undefined && writer.shared[si] !== undefined && writer.shared[si].text === text) {
        // Still the string it was read as: refer to the original, which keeps its phonetic text too.
        out += ' t="s"';
        body += `<${prefix}v>${si}</${prefix}v>`;
      } else {
        out += ' t="inlineStr"';
        body += inlineString(text, cell.runs, prefix);
      }
      break;
    }
    default:
      // blank: a format, a formula without a result yet, or metadata only.
      break;
  }
  return body === '' ? out + '/>' : `${out}>${body}</${prefix}c>`;
}

/**
 * One <row> element, or '' for a row with no cells to write and no properties
 * of its own. `row` is undefined for a row that only has cells the writer
 * itself supplies (see CellWriter.overrides).
 */
function rowXml(writer: CellWriter, index: number, row: Row | undefined): string {
  const prefix = writer.prefix;
  const rowNumber = index + 1;
  let cells = '';
  let list: Cell[] = row ? row.cells : [];
  const overrides = writer.overrides && writer.overrides.get(index);
  if (overrides) {
    list = list.slice();
    for (const [col, cell] of overrides) list[col] = cell;
  }
  const styled = !!(row && row.style);
  // forEach passes over the holes of a sparse row.
  list.forEach((cell, col) => {
    if (col < MAX_COLS) cells += cellXml(writer, cell, rowNumber, col, styled);
  });

  let attrs = '';
  if (row) {
    if (row.style !== undefined) attrs += ` s="${row.style}" customFormat="1"`;
    if (row.height !== undefined) attrs += ` ht="${row.height}"`;
    if (row.customHeight) attrs += ' customHeight="1"';
    if (row.hidden) attrs += ' hidden="1"';
    if (row.outlineLevel) attrs += ` outlineLevel="${row.outlineLevel}"`;
    if (row.collapsed) attrs += ' collapsed="1"';
    if (row.attrs) {
      for (const [name, value] of row.attrs) {
        // An "s" kept from a row without customFormat gives way to a format the row has since been given.
        if (name === 's' && row.style !== undefined) continue;
        attrs += ` ${name}="${escapeAttr(value)}"`;
      }
    }
  }

  if (cells === '') return attrs === '' ? '' : `<${prefix}row r="${rowNumber}"${attrs}/>`;
  return `<${prefix}row r="${rowNumber}"${attrs}>${cells}</${prefix}row>`;
}

/** Every row of the model, in order. */
function allRowsXml(writer: CellWriter, sheet: Sheet, out: string[]): void {
  const indexes = Array.from(sheet.rows.keys());
  if (writer.overrides) for (const index of writer.overrides.keys()) if (!sheet.rows.has(index)) indexes.push(index);
  indexes.sort((a, b) => a - b);
  for (const index of indexes) {
    if (index < 0 || index >= MAX_ROWS) continue;
    const xml = rowXml(writer, index, sheet.rows.get(index));
    if (xml !== '') out.push(xml);
  }
}

/** The range that holds every cell of the model; "A1" for a sheet without cells. */
function usedRange(sheet: Sheet): string {
  let top = -1;
  let bottom = -1;
  let left = MAX_COLS;
  let right = -1;
  for (const [index, row] of sheet.rows) {
    const cells = row.cells;
    let last = cells.length - 1;
    while (last >= 0 && !(last in cells)) last--;
    if (last < 0) continue;
    let first = 0;
    while (!(first in cells)) first++;
    if (top < 0 || index < top) top = index;
    if (index > bottom) bottom = index;
    if (first < left) left = first;
    if (last > right) right = last;
  }
  return top < 0 ? 'A1' : rangeName({ top, left, bottom, right });
}

// --- sheet data -------------------------------------------------------------------

interface Emitted {
  /** Text to put in place of the element. */
  text: string;
  /** Offset just past the element in the source. */
  end: number;
}

/**
 * With the <sheetData> open tag current (it starts at `start`): produces the
 * element and consumes it. While `full` is false, rows that were not changed
 * are copied from the source as they are.
 */
function writeSheetData(sc: XmlScanner, s: string, start: number, sheet: Sheet, writer: CellWriter, full: boolean): Emitted {
  const tagEnd = startTagEnd(s, start);
  const selfClosing = sc.selfClosing;
  const name = sc.name;
  const openTag = selfClosing ? s.slice(start, tagEnd - 2).replace(/\s+$/, '') + '>' : s.slice(start, tagEnd);

  const overrides = writer.overrides;
  let changedRows = sheet.changedRows;
  if (overrides && overrides.size > 0) {
    changedRows = new Set(changedRows);
    for (const index of overrides.keys()) changedRows.add(index);
  }
  // Changed rows that still have something to write, in order; the others were removed.
  const pending: number[] = [];
  if (changedRows) {
    for (const index of changedRows) if (sheet.rows.has(index) || (overrides && overrides.has(index))) pending.push(index);
    pending.sort((a, b) => a - b);
  }
  let next = 0;
  const emitModelRow = (index: number, body: string[]) => {
    const xml = rowXml(writer, index, sheet.rows.get(index));
    if (xml !== '') body.push(xml);
  };

  let body: string[] = [];
  // Elements other than rows are not valid here; whatever they are, they are kept.
  const others: string[] = [];
  let safe = !full;
  let previous = -1;
  let end = tagEnd;
  if (!selfClosing) {
    let cursor = tagEnd;
    while (sc.next()) {
      if (sc.kind === 'close') break;
      if (sc.kind !== 'open') continue;
      const at = nextTag(s, cursor);
      const isRow = sc.local === 'row';
      const index = isRow ? parseUint(sc.attr('r')) - 1 : -1;
      const raw = sc.readRaw();
      const gap = s.slice(cursor, at);
      cursor = at + raw.length;
      if (!isRow) {
        others.push(raw);
        if (safe) body.push(gap, raw);
        continue;
      }
      if (!safe) continue;
      // Copying rows as they are needs every row to say where it is, once, in ascending order.
      if (index < 0 || index <= previous) {
        safe = false;
        continue;
      }
      previous = index;
      while (next < pending.length && pending[next] < index) emitModelRow(pending[next++], body);
      body.push(gap);
      if (changedRows && changedRows.has(index)) {
        if (next < pending.length && pending[next] === index) next++;
        emitModelRow(index, body);
      } else body.push(raw);
    }
    const close = nextTag(s, cursor);
    if (safe) {
      while (next < pending.length) emitModelRow(pending[next++], body);
      body.push(s.slice(cursor, close));
    }
    end = s.indexOf('>', close) + 1;
    if (end <= close) malformed();
  } else if (safe) {
    while (next < pending.length) emitModelRow(pending[next++], body);
  }

  if (!safe) {
    body = [];
    allRowsXml(writer, sheet, body);
    for (const raw of others) body.push(raw);
  }
  const inner = body.join('');
  const text = inner === '' ? openTag.slice(0, -1) + '/>' : `${openTag}${inner}</${name}>`;
  return { text, end };
}

// --- parts of the sheet outside the model, after rows or columns moved --------------------------

const KEEP = 0;
const DROP = 1;
const DROP_PARENT = 2;

interface Adjuster {
  /** Name of the sheet the part belongs to: references without a sheet name are its own. */
  sheetName: string;
  /** Rows and columns inserted into or deleted from this sheet, in order. */
  own: StructuralEdit[];
  /** The same for the other sheets of the workbook; only formulas that name those sheets follow them. */
  foreign: StructuralEdit[];
  /** Lower-case spellings of the other edited sheets' names; text that has none of them needs no look. */
  needles: string[];
  /** Set when anything was altered; an element that was not is copied from the source instead. */
  changed: boolean;
}

/** "B2:B2" -> "B2" for each reference of a list: a range that shrank to one cell is written as that cell. */
function collapseSingleCells(list: string): string {
  if (list.indexOf(':') < 0) return list;
  const refs = list.split(' ');
  for (let i = 0; i < refs.length; i++) {
    const colon = refs[i].indexOf(':');
    if (colon > 0 && refs[i].slice(0, colon) === refs[i].slice(colon + 1)) refs[i] = refs[i].slice(0, colon);
  }
  return refs.join(' ');
}

/** A reference list of this sheet (sqref, ref) after its own edits; '' when all of it was deleted. */
function adjustList(text: string, adjuster: Adjuster): string {
  let next = text;
  for (const edit of adjuster.own) {
    if (next === '') break;
    next = adjustReferenceText(next, adjuster.sheetName, edit);
  }
  return next === text ? text : collapseSingleCells(next);
}

/** A formula that stands on its own (not relative to a cell) after every edit of the workbook. */
function adjustFormulaText(text: string, adjuster: Adjuster): string {
  for (const edit of adjuster.own) text = adjustFormula(text, adjuster.sheetName, edit);
  for (const edit of adjuster.foreign) text = adjustFormula(text, adjuster.sheetName, edit);
  return text;
}

/** Where one cell of this sheet ends up, as text; '' when it was deleted, the text itself when it is no cell. */
function adjustCellRef(text: string, adjuster: Adjuster): string {
  const cell = parseCellRef(text);
  if (!cell) return text;
  let row = cell.row;
  let col = cell.col;
  for (const edit of adjuster.own) {
    const moved = adjustRange({ top: row, bottom: row, left: col, right: col }, edit);
    if (!moved) return '';
    row = moved.top;
    col = moved.left;
  }
  return row === cell.row && col === cell.col ? text : cellName(row, col);
}

const FORMULA_ELEMENTS = ['formula', 'formula1', 'formula2', 'f'];

const elementCount = (el: XmlElement) => el.children.reduce((count, node) => (node.type === 'element' ? count + 1 : count), 0);

const isTextOnly = (el: XmlElement) => el.children.length > 0 && el.children.every((node) => node.type === 'text');

function plainText(el: XmlElement): string {
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

/** Page breaks: <brk id> is the row (or column) the break follows; min and max bound it on the other axis. */
function adjustBreaks(el: XmlElement, rows: boolean, adjuster: Adjuster): number {
  const axis = rows ? 'row' : 'col';
  const crossLimit = rows ? MAX_COLS : MAX_ROWS;
  const kept: XmlNode[] = [];
  let altered = false;
  for (const node of el.children) {
    if (node.type !== 'element' || localName(node.name) !== 'brk') {
      kept.push(node);
      continue;
    }
    let id = parseUint(node.attrs.get('id'));
    let min = parseUint(node.attrs.get('min'));
    let max = parseUint(node.attrs.get('max'));
    let gone = false;
    for (const edit of adjuster.own) {
      if (edit.axis === axis) {
        if (id < 1) continue;
        const moved = adjustRange(rows ? { top: id - 1, bottom: id - 1, left: 0, right: 0 } : { top: 0, bottom: 0, left: id - 1, right: id - 1 }, edit);
        if (!moved) gone = true;
        else id = (rows ? moved.top : moved.left) + 1;
      } else if (min >= 0 && max >= min && !(min === 0 && max >= crossLimit - 1)) {
        // A break that spans the whole sheet keeps doing so.
        const moved = adjustRange(rows ? { top: 0, bottom: 0, left: min, right: max } : { top: min, bottom: max, left: 0, right: 0 }, edit);
        if (!moved) gone = true;
        else {
          min = rows ? moved.left : moved.top;
          max = rows ? moved.right : moved.bottom;
        }
      }
      if (gone) break;
    }
    if (gone) {
      altered = true;
      continue;
    }
    for (const [attr, value] of [['id', id], ['min', min], ['max', max]] as [string, number][]) {
      if (value >= 0 && node.attrs.has(attr) && node.attrs.get(attr) !== String(value)) {
        node.attrs.set(attr, String(value));
        altered = true;
      }
    }
    kept.push(node);
  }
  if (!altered) return KEEP;
  adjuster.changed = true;
  el.children = kept;
  const breaks = kept.filter((node): node is XmlElement => node.type === 'element');
  if (breaks.length === 0) return DROP;
  if (el.attrs.has('count')) el.attrs.set('count', String(breaks.length));
  if (el.attrs.has('manualBreakCount')) {
    el.attrs.set('manualBreakCount', String(breaks.filter((brk) => brk.attrs.get('man') === '1' || brk.attrs.get('man') === 'true').length));
  }
  return KEEP;
}

/** A place a formula is kept in: the text of an element, or an attribute. */
interface FormulaSlot {
  get(): string;
  set(text: string): void;
}

/** The formulas of a conditional format or a validation: <formula>, <formula1>, <xm:f>, and thresholds given as formulas. */
function formulaSlots(el: XmlElement, out: FormulaSlot[]): void {
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (FORMULA_ELEMENTS.indexOf(local) >= 0 && isTextOnly(node)) {
      out.push({
        get: () => plainText(node),
        set: (text) => {
          node.children = [{ type: 'text', value: text }];
        },
      });
    } else if (local === 'cfvo' && node.attrs.get('type') === 'formula' && node.attrs.has('val')) {
      out.push({ get: () => node.attrs.get('val') as string, set: (text) => node.attrs.set('val', text) });
    }
    if (local !== 'sqref') formulaSlots(node, out);
  }
}

/** Top-left cell of the first range of a reference list. */
function anchorOf(list: string): { row: number; col: number } | undefined {
  const first = list.trim().split(/\s+/)[0];
  const range = first ? parseRange(first) : undefined;
  return range ? { row: range.top, col: range.left } : undefined;
}

/**
 * A conditional format or a data validation. Its formulas are written for
 * the top-left cell of its range: a relative reference means "the cell this
 * far from the one the rule is applied to". So what a relative reference has
 * to follow is that corner, not the rows and columns it happens to name;
 * only references anchored with "$" follow those. Returns undefined for an
 * element that has no range of its own.
 */
function adjustRule(el: XmlElement, adjuster: Adjuster): number | undefined {
  // The range is an attribute in the main namespace and an <xm:sqref> child in the extension list.
  const child = el.attrs.has('sqref') ? undefined : childrenOf(el, 'sqref').filter(isTextOnly)[0];
  if (!el.attrs.has('sqref') && !child) return undefined;
  const original = child ? plainText(child) : (el.attrs.get('sqref') as string);
  const slots: FormulaSlot[] = [];
  formulaSlots(el, slots);
  const before = slots.map((slot) => slot.get());
  const texts = before.slice();
  const name = adjuster.sheetName;

  let list = original;
  for (const edit of adjuster.own) {
    const from = anchorOf(list);
    const next = adjustReferenceText(list, name, edit);
    if (next === '') {
      adjuster.changed = true;
      return DROP;
    }
    const to = anchorOf(next);
    for (let i = 0; i < texts.length; i++) {
      texts[i] = from && to ? adjustAnchoredFormula(texts[i], name, edit, from, to) : adjustFormula(texts[i], name, edit);
    }
    list = next;
  }
  // Edits of other sheets move nothing here; only what points at them follows.
  const anchor = anchorOf(list);
  for (const edit of adjuster.foreign) {
    for (let i = 0; i < texts.length; i++) {
      texts[i] = anchor ? adjustAnchoredFormula(texts[i], name, edit, anchor, anchor) : adjustFormula(texts[i], name, edit);
    }
  }

  if (list !== original) {
    list = collapseSingleCells(list);
    if (child) child.children = [{ type: 'text', value: list }];
    else el.attrs.set('sqref', list);
    adjuster.changed = true;
  }
  for (let i = 0; i < texts.length; i++) {
    if (texts[i] === before[i]) continue;
    slots[i].set(texts[i]);
    adjuster.changed = true;
  }
  return KEEP;
}

const childrenOf = (parent: XmlElement, local: string): XmlElement[] =>
  parent.children.filter((node): node is XmlElement => node.type === 'element' && localName(node.name) === local);

/**
 * A filter: its range, and the columns its criteria apply to, which are
 * counted from the left edge of that range. Returns DROP when the range is gone.
 */
function adjustAutoFilter(el: XmlElement, adjuster: Adjuster): number {
  const original = el.attrs.get('ref');
  if (original === undefined) return KEEP;
  let ref = original;
  for (const edit of adjuster.own) {
    const from = parseRange(ref);
    const next = adjustReferenceText(ref, adjuster.sheetName, edit);
    if (next === '') {
      adjuster.changed = true;
      return DROP;
    }
    const to = parseRange(next);
    if (edit.axis === 'col' && from && to) {
      el.children = el.children.filter((node) => {
        if (node.type !== 'element' || localName(node.name) !== 'filterColumn') return true;
        const id = parseUint(node.attrs.get('colId'));
        if (id < 0) return true;
        const moved = adjustRange({ top: 0, bottom: 0, left: from.left + id, right: from.left + id }, edit);
        // The column the criterion was on is gone, and the criterion with it.
        if (!moved || moved.left < to.left || moved.left > to.right) {
          adjuster.changed = true;
          return false;
        }
        const nextId = String(moved.left - to.left);
        if (nextId !== node.attrs.get('colId')) {
          node.attrs.set('colId', nextId);
          adjuster.changed = true;
        }
        return true;
      });
    }
    ref = next;
  }
  if (ref !== original) {
    el.attrs.set('ref', collapseSingleCells(ref));
    adjuster.changed = true;
  }
  return KEEP;
}

/** Whether <dataRef sheet="..."> names the sheet being written. */
const sameSheetName = (a: string | undefined, b: string) => a !== undefined && a.toLowerCase() === b.toLowerCase();

/**
 * Makes the references inside an element follow the structural edits.
 * Returns DROP when the element has nothing left to apply to, DROP_PARENT
 * when that is true of the element around it.
 *
 * Only what is known to be a reference is touched. An attribute called "ref"
 * on an element this function does not know is left alone: moving something
 * that is not a cell range would do more harm than leaving a range behind.
 */
function adjustElement(el: XmlElement, adjuster: Adjuster): number {
  const local = localName(el.name);
  if (local === 'rowBreaks' || local === 'colBreaks') return adjustBreaks(el, local === 'rowBreaks', adjuster);
  if (local === 'conditionalFormatting' || local === 'dataValidation') {
    const result = adjustRule(el, adjuster);
    if (result !== undefined) return result;
  }
  if (local === 'autoFilter' && adjustAutoFilter(el, adjuster) === DROP) return DROP;

  for (const [name, value] of el.attrs) {
    const attr = localName(name);
    let next = value;
    if (attr === 'sqref') next = adjustList(value, adjuster);
    else if (attr === 'activeCell' && local === 'selection') next = adjustList(value, adjuster);
    else if (attr === 'ref' && (local === 'sortState' || local === 'sortCondition')) next = adjustList(value, adjuster);
    // A consolidation source names its sheet; only one on this sheet moves with it.
    else if (attr === 'ref' && local === 'dataRef' && sameSheetName(el.attrs.get('sheet'), adjuster.sheetName)) next = adjustList(value, adjuster);
    else if (attr === 'r' && (local === 'cellWatch' || local === 'inputCells')) next = adjustCellRef(value, adjuster);
    else continue;
    if (next === '') {
      // A view always has a selection. Anything else whose range is gone is gone with it:
      // an empty range is not valid.
      if (local === 'selection') next = 'A1';
      else {
        adjuster.changed = true;
        return DROP;
      }
    }
    if (next !== value) {
      el.attrs.set(name, next);
      adjuster.changed = true;
    }
  }

  const isList = local === 'sqref';
  if ((isList || FORMULA_ELEMENTS.indexOf(local) >= 0) && isTextOnly(el)) {
    const text = plainText(el);
    const next = isList ? adjustList(text, adjuster) : adjustFormulaText(text, adjuster);
    if (next !== text) {
      el.children = [{ type: 'text', value: next }];
      adjuster.changed = true;
    }
    if (isList && next === '') return DROP_PARENT;
    return KEEP;
  }

  let dropped = false;
  const kept: XmlNode[] = [];
  for (const node of el.children) {
    if (node.type === 'element') {
      const result = adjustElement(node, adjuster);
      if (result === DROP_PARENT) {
        adjuster.changed = true;
        return DROP;
      }
      if (result === DROP) {
        dropped = true;
        continue;
      }
    }
    kept.push(node);
  }
  if (dropped) {
    el.children = kept;
    const left = elementCount(el);
    // A container whose entries are all gone would be empty, which the schema does not allow.
    if (left === 0) return DROP;
    if (el.attrs.has('count')) el.attrs.set('count', String(left));
  }
  return KEEP;
}

const elementText = (root: XmlElement) => serializeXml({ declaration: null, leading: [], root, trailing: [] });

/**
 * The text that takes the place of one raw element after the edits: '' when
 * it is dropped, the text itself when nothing in it changed.
 */
function adjustRaw(raw: string, name: string, adjuster: Adjuster): string {
  if (adjuster.own.length === 0) {
    // Only formulas that name another edited sheet can change; most elements name none.
    const lower = raw.toLowerCase();
    if (!adjuster.needles.some((needle) => lower.indexOf(needle) >= 0)) return raw;
  }
  adjuster.changed = false;
  try {
    const root = parseXml(raw).root;
    if (adjustElement(root, adjuster) !== KEEP) return '';
    return adjuster.changed ? elementText(root) : raw;
  } catch (error) {
    // The tree is walked by recursion; content nested without end must fail as bad content, not as a crash.
    if (error instanceof RangeError) throw new Error(`Cannot write the worksheet: <${name}> is nested too deeply to adjust`);
    throw error;
  }
}

// --- regenerated parts of the sheet ---------------------------------------------------

/** Children of CT_Worksheet in schema order. */
const WORKSHEET_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
];

/** The attributes of an element other than the ones the model holds, as they go into a start tag. */
function extraAttrs(el: XmlElement, known: string[], knownLocal?: string): string {
  let out = '';
  for (const [name, value] of el.attrs) {
    if (known.indexOf(name) >= 0 || (knownLocal !== undefined && localName(name) === knownLocal)) continue;
    out += ` ${name}="${escapeAttr(value)}"`;
  }
  return out;
}

const COL_ATTRS = ['min', 'max', 'width', 'customWidth', 'hidden', 'style', 'outlineLevel', 'collapsed'];
const LINK_ATTRS = ['ref', 'location', 'display', 'tooltip'];

interface OriginalColumns {
  min: number;
  max: number;
  extra: string;
}

/**
 * The <col> elements of the source with the attributes the model does not
 * hold (bestFit, phonetic, ...), at the place the edits moved them to.
 */
function originalColumns(raw: string | undefined, edits: StructuralEdit[]): OriginalColumns[] {
  const out: OriginalColumns[] = [];
  if (raw === undefined) return out;
  for (const col of childrenOf(parseXml(raw).root, 'col')) {
    const extra = extraAttrs(col, COL_ATTRS);
    const min = parseUint(col.attrs.get('min'));
    const max = parseUint(col.attrs.get('max'));
    if (extra === '' || min < 1 || max < min) continue;
    let range: CellRange | undefined = { top: 0, bottom: 0, left: min - 1, right: Math.min(max, MAX_COLS) - 1 };
    for (const edit of edits) if (range) range = adjustRange(range, edit);
    if (range) out.push({ min: range.left, max: range.right, extra });
  }
  return out;
}

function colsXml(sheet: Sheet, prefix: string, originals: OriginalColumns[]): string {
  let out = '';
  for (const range of sheet.columns) {
    if (range.min > range.max || range.min < 0 || range.min >= MAX_COLS) continue;
    out += `<${prefix}col min="${range.min + 1}" max="${Math.min(range.max, MAX_COLS - 1) + 1}"`;
    if (range.width !== undefined) out += ` width="${range.width}"`;
    if (range.customWidth) out += ' customWidth="1"';
    if (range.hidden) out += ' hidden="1"';
    if (range.style !== undefined) out += ` style="${range.style}"`;
    if (range.outlineLevel) out += ` outlineLevel="${range.outlineLevel}"`;
    if (range.collapsed) out += ' collapsed="1"';
    // A range that lies within an original one is that one, or a piece of it.
    for (const original of originals) {
      if (original.min <= range.min && original.max >= range.max) {
        out += original.extra;
        break;
      }
    }
    out += '/>';
  }
  return out === '' ? '' : `<${prefix}cols>${out}</${prefix}cols>`;
}

function mergeCellsXml(sheet: Sheet, prefix: string): string {
  if (sheet.merges.length === 0) return '';
  let out = `<${prefix}mergeCells count="${sheet.merges.length}">`;
  for (const range of sheet.merges) out += `<${prefix}mergeCell ref="${rangeName(range)}"/>`;
  return `${out}</${prefix}mergeCells>`;
}

interface OriginalLink {
  /** Where the link is after the edits. */
  range: CellRange;
  rId: string | undefined;
  extra: string;
  used: boolean;
}

/** The <hyperlink> elements of the source that carry attributes the model does not hold (xr:uid, ...). */
function originalLinks(raw: string | undefined, edits: StructuralEdit[]): OriginalLink[] {
  const out: OriginalLink[] = [];
  if (raw === undefined) return out;
  for (const link of childrenOf(parseXml(raw).root, 'hyperlink')) {
    const extra = extraAttrs(link, LINK_ATTRS, 'id');
    const ref = link.attrs.get('ref');
    let range = ref === undefined ? undefined : parseRange(ref);
    if (extra === '' || !range) continue;
    for (const edit of edits) if (range) range = adjustRange(range, edit);
    if (!range) continue;
    let rId: string | undefined;
    for (const [name, value] of link.attrs) if (localName(name) === 'id') rId = value;
    out.push({ range, rId, extra, used: false });
  }
  return out;
}

/** `idAttr` is the relationship id attribute with whatever the element needs to declare its prefix. */
function hyperlinksXml(sheet: Sheet, prefix: string, idAttr: (id: string) => string, originals: OriginalLink[]): string {
  if (sheet.hyperlinks.length === 0) return '';
  let out = `<${prefix}hyperlinks>`;
  for (const link of sheet.hyperlinks) {
    out += `<${prefix}hyperlink ref="${rangeName(link.range)}"`;
    if (link.rId !== undefined) out += idAttr(link.rId);
    if (link.location !== undefined) out += ` location="${escapeAttr(link.location)}"`;
    if (link.display !== undefined) out += ` display="${escapeAttr(link.display)}"`;
    if (link.tooltip !== undefined) out += ` tooltip="${escapeAttr(link.tooltip)}"`;
    // The same place and the same target: the link the source had there.
    for (const original of originals) {
      const r = original.range;
      if (original.used || r.top !== link.range.top || r.left !== link.range.left || r.bottom !== link.range.bottom || r.right !== link.range.right) continue;
      if (original.rId !== link.rId) continue;
      original.used = true;
      out += original.extra;
      break;
    }
    out += '/>';
  }
  return `${out}</${prefix}hyperlinks>`;
}

/** Moves to the root element and returns its local name. */
function openRoot(sc: XmlScanner): string {
  while (sc.next()) {
    if (sc.kind === 'open') break;
  }
  return sc.local;
}

/** Stand in the output for elements that are written last, once the source ones were seen. */
const COLS_MARKER = '\u0000cols';
const LINKS_MARKER = '\u0000hyperlinks';

interface SheetRewrite {
  /** Edits of the other sheets of the workbook, which formulas that name those sheets follow. */
  foreign?: StructuralEdit[];
  /**
   * The sheet itself was not modified: its rows and everything the model
   * holds are copied as they are, and only raw content that refers to an
   * edited sheet is adjusted.
   */
  rawOnly?: boolean;
  overrides?: Map<number, Map<number, Cell>>;
}

interface ChildState {
  sheetData: number;
  linkPrefix: string | undefined;
  colsRaw: string | undefined;
  linksRaw: string | undefined;
  changed: boolean;
}

/**
 * With a child of the worksheet root current: the text that takes its place,
 * and the element consumed.
 */
function writeChild(
  sc: XmlScanner,
  s: string,
  start: number,
  sheet: Sheet,
  writer: CellWriter,
  adjuster: Adjuster | undefined,
  regenerated: Map<string, string>,
  state: ChildState,
  rawOnly: boolean,
): Emitted {
  const local = sc.local;
  const name = sc.name;
  if (local === 'sheetData') {
    state.sheetData++;
    if (state.sheetData > 1) throw new Error('Cannot write the worksheet: it has more than one <sheetData>');
    if (rawOnly) {
      const rows = sc.readRaw();
      return { text: rows, end: start + rows.length };
    }
    return writeSheetData(sc, s, start, sheet, writer, adjuster !== undefined && adjuster.own.length > 0);
  }
  const currentRef = local === 'dimension' ? sc.attr('ref') : undefined;
  const raw = sc.readRaw();
  const end = start + raw.length;

  if (local === 'dimension') {
    const ref = rawOnly ? currentRef : usedRange(sheet);
    if (ref === currentRef) return { text: raw, end };
    const root = parseXml(raw).root;
    root.attrs.set('ref', ref as string);
    return { text: elementText(root), end };
  }

  const replacement = regenerated.get(local);
  if (replacement !== undefined) {
    if (local === 'cols') state.colsRaw = raw;
    if (local === 'hyperlinks') {
      state.linksRaw = raw;
      if (state.linkPrefix === undefined) {
        // Without a prefix declared on the root, the one the original links use is the one in scope.
        const match = /\s([^\s=:"'<>]+):id\s*=/.exec(raw);
        if (match) state.linkPrefix = match[1];
      }
    }
    regenerated.delete(local);
    return { text: replacement, end };
  }
  if (!adjuster) return { text: raw, end };
  const text = adjustRaw(raw, name, adjuster);
  if (text !== raw) state.changed = true;
  return { text, end };
}

/** Lower-case spellings under which a sheet name may appear in the raw text of a part. */
function nameNeedles(edits: StructuralEdit[]): string[] {
  const out: string[] = [];
  for (const edit of edits) {
    const lower = edit.sheetName.toLowerCase();
    for (const spelling of [lower, escapeText(lower), lower.replace(/'/g, "''"), escapeText(lower.replace(/'/g, "''")), escapeAttr(lower)]) {
      if (out.indexOf(spelling) < 0) out.push(spelling);
    }
  }
  return out;
}

/** The sheet part with the model and the edits of the workbook applied; undefined when `rawOnly` and nothing changed. */
function rewriteSheet(original: Uint8Array, sheet: Sheet, workbook: Workbook, options: SheetRewrite): Uint8Array | undefined {
  const rawOnly = !!options.rawOnly;
  const s = decodeSource(original);
  const sc = new XmlScanner(s);
  const rootLocal = openRoot(sc);
  if (rootLocal !== 'worksheet' && rootLocal !== 'dialogsheet' && rootLocal !== 'macrosheet') {
    throw new Error(`Cannot write sheet "${sheet.name}": its part has the root element <${sc.name}>, not a worksheet`);
  }
  const rootName = sc.name;
  const prefix = rootName.slice(0, rootName.length - rootLocal.length);
  let relPrefix: string | undefined;
  for (const [name, value] of sc.attrs()) {
    if (name.slice(0, 6) === 'xmlns:' && (value === NS_REL || value === NS_REL_STRICT)) relPrefix = name.slice(6);
  }
  if (sc.selfClosing) throw new Error(`Cannot write sheet "${sheet.name}": its part has no <sheetData>`);

  const rootStart = nextTag(s, 0);
  let cursor = startTagEnd(s, rootStart);
  const out: string[] = [fixDeclaration(s.slice(0, rootStart)), s.slice(rootStart, cursor)];

  const writer: CellWriter = { prefix, shared: workbook.sharedStrings, styledColumns: sheet.columns.filter((range) => !!range.style), overrides: options.overrides };
  const own: StructuralEdit[] = rawOnly ? [] : (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: sheet.name }));
  const foreign = options.foreign || [];
  const structural = own.length > 0;
  const adjuster: Adjuster | undefined =
    structural || foreign.length > 0 ? { sheetName: sheet.name, own, foreign, needles: nameNeedles(foreign), changed: false } : undefined;

  // These come from the model once it changed them, or once rows or columns moved; they are put at
  // the place the schema gives them when the source has none.
  const regenerated = new Map<string, string>();
  if (!rawOnly) {
    if (structural || sheet.columnsModified) regenerated.set('cols', sheet.columns.length > 0 ? COLS_MARKER : '');
    if (structural || sheet.mergesModified) regenerated.set('mergeCells', mergeCellsXml(sheet, prefix));
    if (structural || sheet.hyperlinksModified) regenerated.set('hyperlinks', sheet.hyperlinks.length > 0 ? LINKS_MARKER : '');
  }
  const state: ChildState = { sheetData: 0, linkPrefix: relPrefix, colsRaw: undefined, linksRaw: undefined, changed: false };
  /** Writes the regenerated elements that belong before the element of rank `rank` and were not in the source. */
  const flush = (rank: number) => {
    if (regenerated.size === 0) return;
    for (const local of ['cols', 'mergeCells', 'hyperlinks']) {
      const text = regenerated.get(local);
      if (text === undefined || WORKSHEET_ORDER.indexOf(local) >= rank) continue;
      out.push(text);
      regenerated.delete(local);
    }
  };

  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    const start = nextTag(s, cursor);
    const rank = WORKSHEET_ORDER.indexOf(sc.local);
    if (rank >= 0) flush(rank);
    out.push(s.slice(cursor, start));
    const emitted = writeChild(sc, s, start, sheet, writer, adjuster, regenerated, state, rawOnly);
    out.push(emitted.text);
    cursor = emitted.end;
  }
  if (state.sheetData === 0) throw new Error(`Cannot write sheet "${sheet.name}": its part has no <sheetData>`);
  flush(WORKSHEET_ORDER.length);
  out.push(s.slice(cursor));
  // What follows the root is still checked for well-formedness.
  while (sc.next());
  if (rawOnly && !state.changed) return undefined;

  // Written last: what the source elements held (unknown attributes, the prefix of relationship ids)
  // is only known once they were passed.
  const colsAt = out.indexOf(COLS_MARKER);
  if (colsAt >= 0) out[colsAt] = colsXml(sheet, prefix, originalColumns(state.colsRaw, own));
  const linksAt = out.indexOf(LINKS_MARKER);
  if (linksAt >= 0) {
    out[linksAt] = hyperlinksXml(
      sheet,
      prefix,
      (id) => (state.linkPrefix !== undefined ? ` ${state.linkPrefix}:id="${escapeAttr(id)}"` : ` xmlns:r="${NS_REL}" r:id="${escapeAttr(id)}"`),
      originalLinks(state.linksRaw, own),
    );
  }
  return encoder.encode(out.join(''));
}

/** One worksheet part regenerated from its original bytes and the model. */
export function writeWorksheet(original: Uint8Array, sheet: Sheet, workbook: Workbook): Uint8Array {
  return rewriteSheet(original, sheet, workbook, {}) as Uint8Array;
}

// --- styles -----------------------------------------------------------------------

type Attrs = [name: string, value: string][];

const element = (name: string, attrs: Attrs = [], children: XmlNode[] = []): XmlElement => ({ type: 'element', name, attrs: new Map(attrs), children });

const childrenNamed = (parent: XmlElement, local: string): XmlElement[] =>
  parent.children.filter((node): node is XmlElement => node.type === 'element' && localName(node.name) === local);

const prefixOf = (el: XmlElement) => el.name.slice(0, el.name.length - localName(el.name).length);

/** Puts a new child where the schema order has it: before the first child that comes later in `order`. */
function insertInOrder(parent: XmlElement, child: XmlElement, order: string[]): void {
  const rank = order.indexOf(localName(child.name));
  let at = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const node = parent.children[i];
    if (node.type === 'element' && order.indexOf(localName(node.name)) > rank) {
      at = i;
      break;
    }
  }
  parent.children.splice(at, 0, child);
}

const STYLESHEET_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

function sameRecord(a: object | undefined, b: object | undefined): boolean {
  const left = (a || {}) as Record<string, unknown>;
  const right = (b || {}) as Record<string, unknown>;
  for (const key of Object.keys(left)) if (left[key] !== right[key]) return false;
  for (const key of Object.keys(right)) if (left[key] !== right[key]) return false;
  return true;
}

function cloneElement(el: XmlElement): XmlElement {
  return {
    type: 'element',
    name: el.name,
    attrs: new Map(el.attrs),
    children: el.children.map((node) => (node.type === 'element' ? cloneElement(node) : { ...node })),
  };
}

function colorElement(name: string, color: Color): XmlElement {
  const attrs: Attrs = [];
  if (color.auto !== undefined) attrs.push(['auto', color.auto ? '1' : '0']);
  if (color.indexed !== undefined) attrs.push(['indexed', String(color.indexed)]);
  if (color.rgb !== undefined) attrs.push(['rgb', color.rgb]);
  if (color.theme !== undefined) attrs.push(['theme', String(color.theme)]);
  if (color.tint !== undefined) attrs.push(['tint', String(color.tint)]);
  return element(name, attrs);
}

/** <font> of the style sheet; children in the order CT_Font gives them. */
function fontElement(font: Font, prefix: string): XmlElement {
  const children: XmlNode[] = [];
  const add = (local: string, val?: string) => children.push(element(prefix + local, val === undefined ? [] : [['val', val]]));
  if (font.bold) add('b');
  if (font.italic) add('i');
  if (font.strike) add('strike');
  if (font.underline === 'single') add('u');
  else if (font.underline && font.underline !== 'none') add('u', font.underline);
  if (font.vertAlign !== undefined) add('vertAlign', font.vertAlign);
  if (font.size !== undefined) add('sz', String(font.size));
  if (font.color) children.push(colorElement(`${prefix}color`, font.color));
  if (font.name !== undefined) add('name', font.name);
  if (font.family !== undefined) add('family', String(font.family));
  if (font.charset !== undefined) add('charset', String(font.charset));
  if (font.scheme !== undefined) add('scheme', font.scheme);
  return element(`${prefix}font`, [], children);
}

function fillElement(fill: Fill, prefix: string): XmlElement {
  // A gradient is kept as it was read.
  if (fill.type === 'gradient') return element(`${prefix}fill`, [], [cloneElement(fill.raw)]);
  const children: XmlNode[] = [];
  if (fill.fg) children.push(colorElement(`${prefix}fgColor`, fill.fg));
  if (fill.bg) children.push(colorElement(`${prefix}bgColor`, fill.bg));
  return element(`${prefix}fill`, [], [element(`${prefix}patternFill`, [['patternType', fill.pattern]], children)]);
}

function borderElement(border: Border, prefix: string): XmlElement {
  const edge = (local: string, value: BorderEdge | undefined) =>
    value ? element(prefix + local, [['style', value.style]], value.color ? [colorElement(`${prefix}color`, value.color)] : []) : element(prefix + local);
  const attrs: Attrs = [];
  if (border.diagonalUp) attrs.push(['diagonalUp', '1']);
  if (border.diagonalDown) attrs.push(['diagonalDown', '1']);
  return element(`${prefix}border`, attrs, [
    edge('left', border.left),
    edge('right', border.right),
    edge('top', border.top),
    edge('bottom', border.bottom),
    edge('diagonal', border.diagonal),
  ]);
}

const ALIGNMENT_ATTRS = ['horizontal', 'vertical', 'textRotation', 'wrapText', 'indent', 'shrinkToFit', 'readingOrder'];
const PROTECTION_ATTRS = ['locked', 'hidden'];

function alignmentElement(alignment: Alignment, prefix: string): XmlElement {
  const attrs: Attrs = [];
  if (alignment.horizontal !== undefined) attrs.push(['horizontal', alignment.horizontal]);
  if (alignment.vertical !== undefined) attrs.push(['vertical', alignment.vertical]);
  if (alignment.textRotation !== undefined) attrs.push(['textRotation', String(alignment.textRotation)]);
  if (alignment.wrapText !== undefined) attrs.push(['wrapText', alignment.wrapText ? '1' : '0']);
  if (alignment.indent !== undefined) attrs.push(['indent', String(alignment.indent)]);
  if (alignment.shrinkToFit !== undefined) attrs.push(['shrinkToFit', alignment.shrinkToFit ? '1' : '0']);
  if (alignment.readingOrder !== undefined) attrs.push(['readingOrder', String(alignment.readingOrder)]);
  return element(`${prefix}alignment`, attrs);
}

function protectionElement(protection: Protection, prefix: string): XmlElement {
  const attrs: Attrs = [];
  if (protection.locked !== undefined) attrs.push(['locked', protection.locked ? '1' : '0']);
  if (protection.hidden !== undefined) attrs.push(['hidden', protection.hidden ? '1' : '0']);
  return element(`${prefix}protection`, attrs);
}

/** <xf> for a format that has no element in the file to start from. */
function formatElement(format: CellFormat, prefix: string): XmlElement {
  const attrs: Attrs = [
    ['numFmtId', String(format.numFmtId)],
    ['fontId', String(format.fontId)],
    ['fillId', String(format.fillId)],
    ['borderId', String(format.borderId)],
  ];
  if (format.xfId !== undefined) attrs.push(['xfId', String(format.xfId)]);
  if (format.quotePrefix) attrs.push(['quotePrefix', '1']);
  if (format.numFmtId) attrs.push(['applyNumberFormat', '1']);
  if (format.fontId) attrs.push(['applyFont', '1']);
  if (format.fillId) attrs.push(['applyFill', '1']);
  if (format.borderId) attrs.push(['applyBorder', '1']);
  const children: XmlNode[] = [];
  if (format.alignment) {
    attrs.push(['applyAlignment', '1']);
    children.push(alignmentElement(format.alignment, prefix));
  }
  if (format.protection) {
    attrs.push(['applyProtection', '1']);
    children.push(protectionElement(format.protection, prefix));
  }
  return element(`${prefix}xf`, attrs, children);
}

const XF_IDS: [field: 'numFmtId' | 'fontId' | 'fillId' | 'borderId', apply: string][] = [
  ['numFmtId', 'applyNumberFormat'],
  ['fontId', 'applyFont'],
  ['fillId', 'applyFill'],
  ['borderId', 'applyBorder'],
];

/** How much of a format an existing one shares; -1 when it cannot serve as its starting point. */
function closeness(base: CellFormat, format: CellFormat): number {
  if (base.xfId !== format.xfId || base.quotePrefix !== format.quotePrefix) return -1;
  let score = 0;
  for (const [field] of XF_IDS) if (base[field] === format[field]) score++;
  if (sameRecord(base.alignment, format.alignment)) score++;
  if (sameRecord(base.protection, format.protection)) score++;
  return score;
}

/** Brings the <alignment> or <protection> child of a copied <xf> in line with the model, keeping attributes the model does not hold. */
function syncFormatChild(xf: XmlElement, wanted: XmlElement | undefined, known: string[], apply: string): void {
  const local = wanted ? localName(wanted.name) : apply === 'applyAlignment' ? 'alignment' : 'protection';
  const current = childrenNamed(xf, local)[0];
  if (!wanted) {
    if (current) xf.children = xf.children.filter((node) => node !== current);
    xf.attrs.delete(apply);
    return;
  }
  if (!current) insertInOrder(xf, wanted, ['alignment', 'protection', 'extLst']);
  else {
    for (const name of known) {
      const value = wanted.attrs.get(name);
      if (value === undefined) current.attrs.delete(name);
      else current.attrs.set(name, value);
    }
  }
  xf.attrs.set(apply, '1');
}

/**
 * <xf> for a format the editor added. A format is made from an existing one
 * by changing some of its properties, so the element starts as a copy of the
 * closest original: that keeps what the model does not hold (apply flags,
 * alignment details). What differs is set and flagged as applied.
 */
function addedFormatElement(format: CellFormat, originals: XmlElement[], formats: CellFormat[], prefix: string): XmlElement {
  let best = -1;
  let bestScore = 0;
  for (let j = 0; j < originals.length; j++) {
    const score = closeness(formats[j], format);
    if (score > bestScore) {
      best = j;
      bestScore = score;
      if (score === 6) break;
    }
  }
  if (best < 0) return formatElement(format, prefix);
  const base = formats[best];
  const xf = cloneElement(originals[best]);
  for (const [field, apply] of XF_IDS) {
    if (base[field] === format[field]) continue;
    xf.attrs.set(field, String(format[field]));
    xf.attrs.set(apply, '1');
  }
  if (!sameRecord(base.alignment, format.alignment)) {
    syncFormatChild(xf, format.alignment && alignmentElement(format.alignment, prefix), ALIGNMENT_ATTRS, 'applyAlignment');
  }
  if (!sameRecord(base.protection, format.protection)) {
    syncFormatChild(xf, format.protection && protectionElement(format.protection, prefix), PROTECTION_ATTRS, 'applyProtection');
  }
  return xf;
}

/**
 * Appends the records the model has beyond those of the file to one of the
 * style tables, creating the table when the file has none. The model's list
 * is the file's list with new entries at the end, so positions line up.
 */
function appendRecords(root: XmlElement, table: string, record: string, count: number, make: (index: number, existing: XmlElement[]) => XmlElement): void {
  let container = childrenNamed(root, table)[0];
  const existing = container ? childrenNamed(container, record) : [];
  if (count <= existing.length) return;
  if (!container) {
    container = element(prefixOf(root) + table);
    insertInOrder(root, container, STYLESHEET_ORDER);
  }
  for (let i = existing.length; i < count; i++) container.children.push(make(i, existing));
  container.attrs.set('count', String(childrenNamed(container, record).length));
}

/**
 * The frame of a style sheet for a workbook that came without one. The font,
 * fill, border and format tables are filled in from the model, so that its
 * indexes mean the same in the file.
 */
function minimalStyles(): XmlDocument {
  const root = element('styleSheet', [['xmlns', NS_MAIN]], [
    element('cellStyleXfs', [['count', '1']], [
      element('xf', [['numFmtId', '0'], ['fontId', '0'], ['fillId', '0'], ['borderId', '0']]),
    ]),
    element('cellStyles', [['count', '1']], [element('cellStyle', [['name', 'Normal'], ['xfId', '0'], ['builtinId', '0']])]),
  ]);
  return { declaration: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', leading: [], root, trailing: [] };
}

/** styles.xml with the records the model added appended; the rest stays as it was. */
export function writeStyles(original: Uint8Array | undefined, styles: StyleSheet): Uint8Array {
  if (original !== undefined && !styles.modified) return original;
  const doc = original !== undefined ? parseXml(original) : minimalStyles();
  const root = doc.root;
  if (localName(root.name) !== 'styleSheet') throw new Error(`Cannot write the styles: the part has the root element <${root.name}>`);
  const prefix = prefixOf(root);

  // Where the file has no table at all, the model's first entry is the default the reader assumed;
  // it is written too, since the formats refer to it.
  appendRecords(root, 'fonts', 'font', styles.fonts.length, (i) => fontElement(styles.fonts[i], prefix));
  appendRecords(root, 'fills', 'fill', styles.fills.length, (i) => fillElement(styles.fills[i], prefix));
  appendRecords(root, 'borders', 'border', styles.borders.length, (i) => borderElement(styles.borders[i], prefix));
  appendRecords(root, 'cellXfs', 'xf', styles.cellXfs.length, (i, existing) => addedFormatElement(styles.cellXfs[i], existing, styles.cellXfs, prefix));

  // Number formats the model has and the file does not.
  let numFmts = childrenNamed(root, 'numFmts')[0];
  const known = new Set<number>();
  if (numFmts) for (const numFmt of childrenNamed(numFmts, 'numFmt')) known.add(parseUint(numFmt.attrs.get('numFmtId')));
  let added = false;
  for (const [id, code] of styles.numFmts) {
    if (known.has(id)) continue;
    if (!numFmts) {
      numFmts = element(`${prefix}numFmts`);
      insertInOrder(root, numFmts, STYLESHEET_ORDER);
    }
    numFmts.children.push(element(`${prefix}numFmt`, [['numFmtId', String(id)], ['formatCode', code]]));
    added = true;
  }
  if (numFmts && added) numFmts.attrs.set('count', String(childrenNamed(numFmts, 'numFmt').length));

  return encoder.encode(serializeXml(doc));
}

// --- parts placed by cell position: pictures and comments ------------------------------------

/** Where a row or column index ends up; `snapped` when it was deleted and now stands for the place of the deletion. */
function shiftIndex(value: number, edit: SheetStructuralEdit): { value: number; snapped: boolean } {
  const limit = edit.axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (edit.kind === 'insert') return { value: value >= edit.at ? Math.min(value + edit.count, limit - 1) : value, snapped: false };
  const end = edit.at + edit.count - 1;
  if (value > end) return { value: value - edit.count, snapped: false };
  if (value >= edit.at) return { value: edit.at, snapped: true };
  return { value, snapped: false };
}

const childNamed = (parent: XmlElement | undefined, local: string): XmlElement | undefined => (parent ? childrenNamed(parent, local)[0] : undefined);

function textOf(el: XmlElement | undefined): string | undefined {
  if (!el) return undefined;
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

function setText(el: XmlElement, text: string): boolean {
  if (textOf(el) === text) return false;
  el.children = [{ type: 'text', value: text }];
  return true;
}

/**
 * Moves one drawing anchor. An object that moves and sizes with cells has
 * both corners adjusted; one that only moves keeps its size; an absolute one
 * stays. A corner whose row or column was deleted goes to where the deletion
 * was, so an object whose rows are all gone is left with no height there.
 */
function adjustAnchor(anchor: XmlElement, edits: SheetStructuralEdit[]): boolean {
  const twoCell = localName(anchor.name) === 'twoCellAnchor';
  const editAs = anchor.attrs.get('editAs');
  if (twoCell && editAs === 'absolute') return false;
  const from = childNamed(anchor, 'from');
  const to = twoCell ? childNamed(anchor, 'to') : undefined;
  if (!from) return false;
  let changed = false;
  for (const edit of edits) {
    const axis = edit.axis;
    const fromEl = childNamed(from, axis);
    const before = parseUint((textOf(fromEl) || '').trim());
    if (!fromEl || before < 0) continue;
    const moved = shiftIndex(before, edit);
    if (setText(fromEl, String(moved.value))) changed = true;
    const fromOff = childNamed(from, `${axis}Off`);
    if (moved.snapped && fromOff && setText(fromOff, '0')) changed = true;
    const toEl = childNamed(to, axis);
    const toBefore = parseUint((textOf(toEl) || '').trim());
    if (!to || !toEl || toBefore < 0) continue;
    if (editAs === 'oneCell') {
      // Moves with its top-left cell and keeps its size.
      const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
      if (setText(toEl, String(Math.max(0, Math.min(toBefore + moved.value - before, limit - 1))))) changed = true;
    } else {
      const end = shiftIndex(toBefore, edit);
      if (setText(toEl, String(end.value))) changed = true;
      const toOff = childNamed(to, `${axis}Off`);
      if (end.snapped && toOff && setText(toOff, '0')) changed = true;
    }
  }
  return changed;
}

function adjustAnchorsIn(el: XmlElement, edits: SheetStructuralEdit[]): boolean {
  let changed = false;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
      if (adjustAnchor(node, edits)) changed = true;
    } else if (local === 'AlternateContent' || local === 'Choice' || local === 'Fallback') {
      // Newer kinds of objects are wrapped together with a fallback picture; both carry an anchor.
      if (adjustAnchorsIn(node, edits)) changed = true;
    }
  }
  return changed;
}

/** Shapes and connectors may show the content of a cell: textlink="Sheet1!$A$5" is a formula. */
function adjustTextLinks(el: XmlElement, edits: SheetStructuralEdit[], sheetName: string): boolean {
  let changed = false;
  const link = el.attrs.get('textlink');
  if (link !== undefined && link !== '') {
    let next = link;
    for (const edit of edits) next = adjustFormula(next, sheetName, { ...edit, sheetName });
    if (next !== link) {
      el.attrs.set('textlink', next);
      changed = true;
    }
  }
  for (const node of el.children) if (node.type === 'element' && adjustTextLinks(node, edits, sheetName)) changed = true;
  return changed;
}

/** A drawing part with its anchors moved; the same bytes when nothing moved. */
function adjustDrawing(bytes: Uint8Array, edits: SheetStructuralEdit[], sheetName: string): Uint8Array {
  const doc = parseXml(bytes);
  const moved = adjustAnchorsIn(doc.root, edits);
  let linked: boolean;
  try {
    linked = adjustTextLinks(doc.root, edits, sheetName);
  } catch (error) {
    if (error instanceof RangeError) throw new Error('Cannot write the drawing: its content is nested too deeply to adjust');
    throw error;
  }
  return moved || linked ? encoder.encode(serializeXml(doc)) : bytes;
}

/** Where a single cell ends up; undefined when it was deleted. */
function moveCell(row: number, col: number, edits: SheetStructuralEdit[]): { row: number; col: number } | undefined {
  for (const edit of edits) {
    const moved = adjustRange({ top: row, bottom: row, left: col, right: col }, edit);
    if (!moved) return undefined;
    row = moved.top;
    col = moved.left;
  }
  return { row, col };
}

/** A comments part with each comment at its cell's new place; comments of deleted cells are removed. */
function adjustComments(bytes: Uint8Array, edits: SheetStructuralEdit[]): Uint8Array {
  const doc = parseXml(bytes);
  let changed = false;
  for (const list of childrenNamed(doc.root, 'commentList')) {
    list.children = list.children.filter((node) => {
      if (node.type !== 'element' || localName(node.name) !== 'comment') return true;
      const ref = node.attrs.get('ref');
      const cell = ref === undefined ? undefined : parseCellRef(ref);
      if (!cell) return true;
      const moved = moveCell(cell.row, cell.col, edits);
      if (!moved) {
        changed = true;
        return false;
      }
      const next = cellName(moved.row, moved.col);
      if (next !== ref) {
        node.attrs.set('ref', next);
        changed = true;
      }
      return true;
    });
  }
  return changed ? encoder.encode(serializeXml(doc)) : bytes;
}

interface Note {
  row: number;
  col: number;
  /** Text of <x:Anchor>: left column, offset, top row, offset, right column, offset, bottom row, offset. */
  anchor: string | undefined;
}

/** A comment box of the legacy drawing at its new place; undefined when its cell was deleted. */
function adjustNote(note: Note, edits: SheetStructuralEdit[]): Note | undefined {
  const cell = moveCell(note.row, note.col, edits);
  if (!cell) return undefined;
  let anchor = note.anchor;
  if (anchor !== undefined) {
    const numbers = anchor.split(',').map((part) => parseUint(part.trim()));
    if (numbers.length === 8 && numbers.every((n) => n >= 0)) {
      for (const edit of edits) {
        for (const at of edit.axis === 'col' ? [0, 4] : [2, 6]) {
          const moved = shiftIndex(numbers[at], edit);
          numbers[at] = moved.value;
          if (moved.snapped) numbers[at + 1] = 0;
        }
      }
      const lead = /^\s*/.exec(anchor) as RegExpExecArray;
      const trail = /\s*$/.exec(anchor) as RegExpExecArray;
      const next = lead[0] + numbers.join(', ') + trail[0];
      // Only the numbers matter; an anchor that did not move keeps its own spacing.
      if (next.replace(/\s/g, '') !== anchor.replace(/\s/g, '')) anchor = next;
    }
  }
  return { row: cell.row, col: cell.col, anchor };
}

/** The legacy drawing as a tree. Returns undefined when nothing changed. */
function adjustVmlTree(doc: XmlDocument, edits: SheetStructuralEdit[]): XmlDocument | undefined {
  let changed = false;
  doc.root.children = doc.root.children.filter((node) => {
    if (node.type !== 'element' || localName(node.name) !== 'shape') return true;
    const data = childNamed(node, 'ClientData');
    // Only comment boxes are known to sit on a cell this way; form controls and the rest stay.
    if (!data || data.attrs.get('ObjectType') !== 'Note') return true;
    const rowEl = childNamed(data, 'Row');
    const colEl = childNamed(data, 'Column');
    const anchorEl = childNamed(data, 'Anchor');
    const row = parseUint((textOf(rowEl) || '').trim());
    const col = parseUint((textOf(colEl) || '').trim());
    if (!rowEl || !colEl || row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return true;
    const moved = adjustNote({ row, col, anchor: textOf(anchorEl) }, edits);
    if (!moved) {
      changed = true;
      return false;
    }
    if (moved.row !== row && setText(rowEl, String(moved.row))) changed = true;
    if (moved.col !== col && setText(colEl, String(moved.col))) changed = true;
    if (anchorEl && moved.anchor !== undefined && setText(anchorEl, moved.anchor)) changed = true;
    return true;
  });
  return changed ? doc : undefined;
}

/**
 * The same on the text, for a legacy drawing that is not well-formed XML
 * (VML is older than the rule that it has to be). Only the three elements of
 * a comment box are touched; everything else is copied.
 */
function adjustVmlText(text: string, edits: SheetStructuralEdit[]): string {
  let out = '';
  let copied = 0;
  let i = 0;
  for (;;) {
    const start = text.indexOf('<v:shape', i);
    if (start < 0) break;
    const after = text.charCodeAt(start + 8);
    // "<v:shapetype" and the like are other elements.
    if (!(after === 32 || after === 62 || after === 9 || after === 10 || after === 13)) {
      i = start + 8;
      continue;
    }
    const close = text.indexOf('</v:shape>', start);
    if (close < 0) break;
    const end = close + 10;
    i = end;
    const block = text.slice(start, end);
    if (!/<x:ClientData[^>]*ObjectType\s*=\s*["']Note["']/.test(block)) continue;
    const rowMatch = /<x:Row>\s*(\d{1,7})\s*<\/x:Row>/.exec(block);
    const colMatch = /<x:Column>\s*(\d{1,5})\s*<\/x:Column>/.exec(block);
    if (!rowMatch || !colMatch) continue;
    const anchorMatch = /<x:Anchor>([^<]*)<\/x:Anchor>/.exec(block);
    const moved = adjustNote({ row: Number(rowMatch[1]), col: Number(colMatch[1]), anchor: anchorMatch ? anchorMatch[1] : undefined }, edits);
    let next = '';
    if (moved) {
      next = block.replace(rowMatch[0], `<x:Row>${moved.row}</x:Row>`).replace(colMatch[0], `<x:Column>${moved.col}</x:Column>`);
      if (anchorMatch && moved.anchor !== undefined) next = next.replace(anchorMatch[0], `<x:Anchor>${moved.anchor}</x:Anchor>`);
      if (next === block) continue;
    }
    out += text.slice(copied, start) + next;
    copied = end;
  }
  return copied === 0 ? text : out + text.slice(copied);
}

/** A legacy drawing part (the boxes of comments) moved along; the same bytes when nothing moved. */
function adjustVml(bytes: Uint8Array, edits: SheetStructuralEdit[]): Uint8Array {
  let doc: XmlDocument | undefined;
  try {
    doc = parseXml(bytes);
  } catch (error) {
    doc = undefined;
  }
  if (doc) {
    const next = adjustVmlTree(doc, edits);
    return next ? encoder.encode(serializeXml(next)) : bytes;
  }
  const text = decodeSource(bytes);
  const next = adjustVmlText(text, edits);
  return next === text ? bytes : encoder.encode(next);
}

/** Parts of a sheet that follow its rows and columns, by the last segment of their relationship type. */
const ANCHORED_PARTS: Record<string, (bytes: Uint8Array, edits: SheetStructuralEdit[], sheetName: string) => Uint8Array> = {
  drawing: adjustDrawing,
  comments: adjustComments,
  vmlDrawing: adjustVml,
};

// --- the package --------------------------------------------------------------------

/** Children of CT_Workbook in schema order. */
const WORKBOOK_ORDER = [
  'fileVersion',
  'fileSharing',
  'workbookPr',
  'workbookProtection',
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
];

/** Brings workbook.xml up to date: defined names follow the model, and formulas are recalculated on load. */
function updateWorkbookXml(doc: XmlDocument, workbook: Workbook): void {
  const root = doc.root;
  const prefix = prefixOf(root);
  const container = childrenNamed(root, 'definedNames')[0];
  const elements = container ? childrenNamed(container, 'definedName') : [];
  const used = new Set<XmlElement>();
  const matches = (el: XmlElement, name: string, localSheet: number | undefined) => {
    if (used.has(el) || (el.attrs.get('name') || '') !== name) return false;
    const local = el.attrs.get('localSheetId');
    return (local !== undefined && /^\d{1,6}$/.test(local) ? Number(local) : undefined) === localSheet;
  };
  workbook.definedNames.forEach((name, index) => {
    // The lists normally line up; otherwise the name and its scope identify the element.
    let target: XmlElement | undefined = elements[index];
    if (!target || !matches(target, name.name, name.localSheet)) target = elements.find((el) => matches(el, name.name, name.localSheet));
    if (!target) throw new Error(`Cannot write the workbook: the defined name "${name.name}" is not in the file`);
    used.add(target);
    if (textOf(target) !== name.value) target.children = [{ type: 'text', value: name.value }];
  });

  // Results stored in the file may be stale after an edit; the application recalculates when it opens the file.
  let calcPr = childrenNamed(root, 'calcPr')[0];
  if (!calcPr) {
    calcPr = element(`${prefix}calcPr`);
    insertInOrder(root, calcPr, WORKBOOK_ORDER);
  }
  calcPr.attrs.set('fullCalcOnLoad', '1');
}

function resolveTarget(source: string, target: string): string {
  const base = target.charAt(0) === '/' ? target : `${source.slice(0, source.lastIndexOf('/'))}/${target}`;
  const segments: string[] = [];
  for (const segment of base.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return `/${segments.join('/')}`;
}

const endsWith = (text: string, suffix: string) => text.slice(-suffix.length) === suffix;

const relsPartOf = (part: string) => `${part.slice(0, part.lastIndexOf('/'))}/_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;

interface SaveState {
  /** The bytes each rewritten part had when the file was opened. null stands for a part the file did not have. */
  originals: Map<string, Uint8Array | null>;
  /**
   * Per sheet part: the parts around it (drawings, comments, relationships)
   * that the last save wrote differently from the original. A save that no
   * longer changes them has to put the original back.
   */
  moved: Map<string, string[]>;
}

const SAVES = new WeakMap<Workbook, SaveState>();

type PartWrite = [part: string, data: Uint8Array | XmlDocument, contentType?: string];

/**
 * Brings the names of a table's columns in line with its header cells, the
 * way the application does: a column is named by the text of its header; an
 * empty header is named "Column<N>"; a name already taken gets a number
 * appended. Returns undefined for a table without a header row or whose
 * header row was not edited; otherwise whether the table changed, and the
 * header cells that have to be written with the name the column got.
 */
function syncTableColumns(table: XmlElement, sheet: Sheet, workbook: Workbook): { changed: boolean; row: number; cells: Map<number, Cell> } | undefined {
  if (localName(table.name) !== 'table' || table.attrs.get('headerRowCount') === '0') return undefined;
  const range = parseRange(table.attrs.get('ref') || '');
  if (!range || !sheet.changedRows || !sheet.changedRows.has(range.top)) return undefined;
  const columns = childrenNamed(childNamed(table, 'tableColumns') || table, 'tableColumn');
  const taken = new Set<string>();
  const cells = new Map<number, Cell>();
  let changed = false;
  for (let i = 0; i < columns.length && range.left + i <= range.right; i++) {
    const col = range.left + i;
    const cell = getCell(sheet, range.top, col);
    const text = cell ? cellDisplayText(workbook, cell) : '';
    const base = text === '' ? `Column${i + 1}` : text;
    let name = base;
    for (let n = 2; taken.has(name.toLowerCase()); n++) name = base + n;
    taken.add(name.toLowerCase());
    // A header that is not text, or not the name the column got, is written as that name.
    if (!cell || cell.kind !== 'string' || cell.formula || cell.value !== name) cells.set(col, { kind: 'string', value: name, style: cell ? cell.style : 0 });
    if (columns[i].attrs.get('name') !== name) {
      columns[i].attrs.set('name', name);
      changed = true;
    }
  }
  return { changed, row: range.top, cells };
}

/** Brings the package up to date with the model. Does nothing when nothing was modified. */
export async function writeXlsx(pkg: OpcPackage, workbook: Workbook): Promise<void> {
  const modifiedSheets = workbook.sheets.filter((sheet) => sheet.modified && sheet.kind !== 'chartsheet');
  const stylesModified = !!workbook.styles.modified;
  if (modifiedSheets.length === 0 && !stylesModified) return;

  let saved = SAVES.get(workbook);
  if (!saved) {
    saved = { originals: new Map(), moved: new Map() };
    SAVES.set(workbook, saved);
  }
  const { originals, moved } = saved;
  // The first time a part is about to be rewritten, the package still holds it as the file had it.
  const originalOf = async (part: string): Promise<Uint8Array | undefined> => {
    const kept = originals.get(part);
    if (kept !== undefined) return kept === null ? undefined : kept;
    const bytes = pkg.hasPart(part) ? await pkg.readPart(part) : undefined;
    originals.set(part, bytes === undefined ? null : bytes);
    return bytes;
  };

  const main = await pkg.mainPart();
  if (!main) throw new Error('Cannot write the workbook: the package has no workbook part');
  const slash = main.lastIndexOf('/');
  const relsPart = relsPartOf(main);
  const relsBytes = await originalOf(relsPart);
  if (!relsBytes) throw new Error('Cannot write the workbook: its relationships part is missing');
  const rels = parseXml(relsBytes);
  const relationships = childrenNamed(rels.root, 'Relationship');
  const internal = (rel: XmlElement) => rel.attrs.get('TargetMode') !== 'External';
  const typeOf = (rel: XmlElement) => rel.attrs.get('Type') || '';
  const withType = (suffix: string) => relationships.find((rel) => endsWith(typeOf(rel), suffix) && internal(rel));
  let relsChanged = false;

  // Everything is produced before anything is put into the package: a failure leaves the package as it was.
  const writes: PartWrite[] = [];

  if (stylesModified) {
    const stylesRel = withType('/styles');
    const stylesPart = stylesRel ? resolveTarget(main, stylesRel.attrs.get('Target') || '') : `${main.slice(0, slash)}/styles.xml`;
    writes.push([stylesPart, writeStyles(await originalOf(stylesPart), workbook.styles), STYLES_CONTENT_TYPE]);
    if (!stylesRel) {
      let highest = 0;
      for (const rel of relationships) {
        const match = /^rId(\d{1,9})$/.exec(rel.attrs.get('Id') || '');
        if (match && Number(match[1]) > highest) highest = Number(match[1]);
      }
      // The same flavour of relationship types as the rest of the file (transitional or strict).
      const sample = withType('/worksheet');
      const sampleType = sample ? typeOf(sample) : '';
      const base = sampleType ? sampleType.slice(0, sampleType.lastIndexOf('/')) : NS_REL;
      rels.root.children.push(
        element(`${prefixOf(rels.root)}Relationship`, [
          ['Id', `rId${highest + 1}`],
          ['Type', `${base}/styles`],
          ['Target', stylesPart.slice(0, slash + 1) === main.slice(0, slash + 1) ? stylesPart.slice(slash + 1) : stylesPart],
        ]),
      );
      relsChanged = true;
    }
  }

  // Rows and columns moved in one sheet change what the others say about it.
  const editsOf = (sheet: Sheet): StructuralEdit[] => (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: sheet.name }));
  const edited = workbook.sheets.filter((sheet) => sheet.kind !== 'chartsheet' && sheet.structuralEdits && sheet.structuralEdits.length > 0);
  const foreignOf = (sheet: Sheet): StructuralEdit[] => {
    let out: StructuralEdit[] = [];
    for (const other of edited) if (other !== sheet) out = out.concat(editsOf(other));
    return out;
  };

  const movedNow = new Map<string, string[]>();
  for (const sheet of workbook.sheets) {
    if (sheet.kind === 'chartsheet' || sheet.modified) continue;
    // A sheet that was not modified itself: only its raw content that names an edited sheet changes
    // (the source of a validation list on another sheet, say). Its cells were adjusted through the
    // model, which marked the sheet as modified if any of them changed.
    const foreign = foreignOf(sheet);
    const before = moved.get(sheet.part) || [];
    if (foreign.length === 0 && before.length === 0) continue;
    const kept = originals.get(sheet.part);
    const bytes = kept || (pkg.hasPart(sheet.part) ? await pkg.readPart(sheet.part) : undefined);
    if (!bytes) continue;
    const next = foreign.length > 0 ? rewriteSheet(bytes, sheet, workbook, { foreign, rawOnly: true }) : undefined;
    if (next) {
      // From here on the package no longer holds the part as the file had it.
      originals.set(sheet.part, bytes);
      writes.push([sheet.part, next]);
      movedNow.set(sheet.part, [sheet.part]);
    } else {
      if (before.indexOf(sheet.part) >= 0 && kept) writes.push([sheet.part, kept]);
      movedNow.set(sheet.part, []);
    }
  }

  for (const sheet of modifiedSheets) {
    const bytes = await originalOf(sheet.part);
    if (!bytes) throw new Error(`Cannot write sheet "${sheet.name}": its part ${sheet.part} is missing`);
    const rewritten: string[] = [];
    const sheetRelsPart = relsPartOf(sheet.part);

    // Tables name their columns after their header cells; the two have to agree.
    let overrides: Map<number, Map<number, Cell>> | undefined;
    if (sheet.changedRows && sheet.changedRows.size > 0) {
      const tableRels = await originalOf(sheetRelsPart);
      for (const rel of tableRels ? childrenNamed(parseXml(tableRels).root, 'Relationship') : []) {
        if (!endsWith(typeOf(rel), '/table') || !internal(rel)) continue;
        const part = resolveTarget(sheet.part, rel.attrs.get('Target') || '');
        const original = await originalOf(part);
        if (!original) continue;
        const doc = parseXml(original);
        const headers = syncTableColumns(doc.root, sheet, workbook);
        if (!headers) continue;
        if (headers.changed) {
          writes.push([part, doc]);
          rewritten.push(part);
        }
        if (headers.cells.size > 0) {
          if (!overrides) overrides = new Map();
          const row = overrides.get(headers.row) || new Map<number, Cell>();
          for (const [col, cell] of headers.cells) row.set(col, cell);
          overrides.set(headers.row, row);
        }
      }
    }
    writes.push([sheet.part, rewriteSheet(bytes, sheet, workbook, { foreign: foreignOf(sheet), overrides }) as Uint8Array]);

    // After rows or columns moved: the pictures and comments of the sheet follow, and the
    // relationships of links that are gone are removed.
    const edits = sheet.structuralEdits && sheet.structuralEdits.length > 0 ? sheet.structuralEdits : undefined;
    const sheetRelsBytes = edits ? await originalOf(sheetRelsPart) : undefined;
    if (edits && sheetRelsBytes) {
      const sheetRels = parseXml(sheetRelsBytes);
      const linkIds = new Set<string>();
      for (const link of sheet.hyperlinks) if (link.rId !== undefined) linkIds.add(link.rId);
      let dropped = false;
      for (const rel of childrenNamed(sheetRels.root, 'Relationship')) {
        const type = typeOf(rel);
        const kind = type.slice(type.lastIndexOf('/') + 1);
        if (kind === 'hyperlink') {
          if (!linkIds.has(rel.attrs.get('Id') || '')) {
            sheetRels.root.children = sheetRels.root.children.filter((node) => node !== rel);
            dropped = true;
          }
          continue;
        }
        const adjust = Object.prototype.hasOwnProperty.call(ANCHORED_PARTS, kind) ? ANCHORED_PARTS[kind] : undefined;
        if (!adjust || !internal(rel)) continue;
        const part = resolveTarget(sheet.part, rel.attrs.get('Target') || '');
        const original = await originalOf(part);
        if (!original) continue;
        // One sheet's rows cannot move a part that another sheet shows as well.
        for (const other of workbook.sheets) {
          if (other === sheet || other.kind === 'chartsheet') continue;
          for (const otherRel of await pkg.relationships(other.part)) {
            if (otherRel.part === part) throw new Error(`Cannot write sheet "${sheet.name}": ${part} is shared with sheet "${other.name}"`);
          }
        }
        const adjusted = adjust(original, edits, sheet.name);
        if (adjusted !== original) {
          writes.push([part, adjusted]);
          rewritten.push(part);
        }
      }
      if (dropped) {
        writes.push([sheetRelsPart, sheetRels]);
        rewritten.push(sheetRelsPart);
      }
    }
    // What an earlier save moved and this one does not goes back to what the file had.
    for (const part of moved.get(sheet.part) || []) {
      const original = originals.get(part);
      if (rewritten.indexOf(part) < 0 && original) writes.push([part, original]);
    }
    movedNow.set(sheet.part, rewritten);
  }

  let calcChain: string | undefined;
  if (modifiedSheets.length > 0) {
    const mainBytes = await originalOf(main);
    if (!mainBytes) throw new Error('Cannot write the workbook: its part is missing');
    const doc = parseXml(mainBytes);
    updateWorkbookXml(doc, workbook);
    writes.push([main, doc]);

    // A calculation chain that no longer matches the formulas makes the application report the file
    // as damaged; without one it simply builds a new chain.
    const chainRel = withType('/calcChain');
    if (chainRel) {
      calcChain = resolveTarget(main, chainRel.attrs.get('Target') || '');
      rels.root.children = rels.root.children.filter((node) => node !== chainRel);
      relsChanged = true;
    }
  }

  for (const [part, data, contentType] of writes) pkg.writePart(part, data, contentType);
  for (const [part, list] of movedNow) moved.set(part, list);
  if (calcChain !== undefined) pkg.deletePart(calcChain);
  if (relsChanged) pkg.writePart(relsPart, rels);
}
