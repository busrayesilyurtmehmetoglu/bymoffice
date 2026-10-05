// Writes an edited spreadsheet model back into its xlsx package.
//
// Only what changed is produced again. A sheet that was not modified keeps its
// part untouched. In a modified sheet, the rows that were not edited and every
// element the model does not hold are copied from the original text as they
// are; rows that were edited are written from the model. Parts are always
// regenerated from the bytes the file had when it was opened, never from the
// output of an earlier save: the model records its changes relative to those.

import { adjustAnchoredFormula, adjustFormula, adjustRange, adjustReferenceText, renameSheetsInFormula, type StructuralEdit } from '../../calc/refs';
import {
  FILTER_NAME,
  MAX_COLS,
  MAX_ROWS,
  cellName,
  columnName,
  foldSheetName,
  getCell,
  parseCellRef,
  parseRange,
  rangeName,
  sheetNameMap,
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
  type FilterColumn,
  type Font,
  type Protection,
  type RichRun,
  type Row,
  type SharedString,
  type Sheet,
  type SheetNameMap,
  type SheetStructuralEdit,
  type SortState,
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
const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

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
export function replaceUnwritable(text: string): string {
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

export function escapeText(text: string): string {
  text = replaceUnwritable(text);
  return /[&<>\r]/.test(text) ? text.replace(/[&<>\r]/g, (ch) => TEXT_ESCAPES[ch]) : text;
}

export function escapeAttr(text: string): string {
  text = replaceUnwritable(text);
  return /[&<"\t\n\r]/.test(text) ? text.replace(/[&<"\t\n\r]/g, (ch) => ATTR_ESCAPES[ch]) : text;
}

/** Element text of an ST_Xstring value: "_xHHHH_" escapes first, then XML escapes. */
export function escapeXstring(text: string): string {
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

export interface CellWriter {
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
  /**
   * For a writer that makes the shared string table itself (a package built from the
   * model alone): gives a text that has no entry yet the index of one. Without it such
   * a text is written into the cell.
   */
  intern?(text: string, runs: RichRun[] | undefined): number;
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

export function colorAttrs(color: Color): string {
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

export function textElement(text: string, prefix: string): string {
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
      } else if (writer.intern) {
        out += ' t="s"';
        body += `<${prefix}v>${writer.intern(text, cell.runs)}</${prefix}v>`;
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
export function rowXml(writer: CellWriter, index: number, row: Row | undefined): string {
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
  /** Name the file has for the sheet the part belongs to: references without a sheet name are its own. */
  sheetName: string;
  /** Rows and columns inserted into or deleted from this sheet, in order. */
  own: StructuralEdit[];
  /** The same for the other sheets of the workbook; only formulas that name those sheets follow them. */
  foreign: StructuralEdit[];
  /** Folded spellings of the names of the other edited sheets and of the sheets that were renamed or deleted; text that has none of them needs no look. */
  needles: string[];
  /**
   * What the sheet names of the file stand for now, when sheets were renamed or deleted. Applied
   * to a formula last: the edits above are matched with the names the file has.
   */
  rename?: SheetNameMap;
  /** Set when anything was altered; an element that was not is copied from the source instead. */
  changed: boolean;
}

/**
 * "B2:B2" -> "B2" for each reference of a list: a range that shrank to one cell is written as
 * that cell. A whole column or row ("C:C", "4:4") has no shorter form: "C" alone is not a
 * reference, and a rule written with it is lost on the next read.
 */
function collapseSingleCells(list: string): string {
  if (list.indexOf(':') < 0) return list;
  const refs = list.split(' ');
  for (let i = 0; i < refs.length; i++) {
    const colon = refs[i].indexOf(':');
    if (colon <= 0) continue;
    const first = refs[i].slice(0, colon);
    if (first === refs[i].slice(colon + 1) && parseCellRef(first)) refs[i] = first;
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
  return adjuster.rename ? renameSheetsInFormula(text, adjuster.rename) : text;
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

/** One side of a whole-column or whole-row reference ("$B", "7") as a 0-based index; -1 for anything else. */
function axisIndex(text: string, rows: boolean): number {
  let i = text.charCodeAt(0) === 36 ? 1 : 0;
  const n = text.length;
  if (i >= n || n - i > 7) return -1;
  let value = 0;
  for (; i < n; i++) {
    const c = text.charCodeAt(i);
    if (rows) {
      if (c < 48 || c > 57) return -1;
      value = value * 10 + (c - 48);
    } else {
      const letter = c & ~32;
      if (letter < 65 || letter > 90) return -1;
      value = value * 26 + (letter - 64);
    }
  }
  return value >= 1 && value <= (rows ? MAX_ROWS : MAX_COLS) ? value - 1 : -1;
}

/**
 * Top-left cell of the first range of a reference list. A whole column starts at row 1 and a
 * whole row at column A, as the reader has them.
 */
function anchorOf(list: string): { row: number; col: number } | undefined {
  const first = list.trim().split(/\s+/)[0];
  if (!first) return undefined;
  const range = parseRange(first);
  if (range) return { row: range.top, col: range.left };
  const colon = first.indexOf(':');
  if (colon <= 0) return undefined;
  const a = first.slice(0, colon);
  const b = first.slice(colon + 1);
  const cols = [axisIndex(a, false), axisIndex(b, false)];
  if (cols[0] >= 0 && cols[1] >= 0) return { row: 0, col: Math.min(cols[0], cols[1]) };
  const rows = [axisIndex(a, true), axisIndex(b, true)];
  if (rows[0] >= 0 && rows[1] >= 0) return { row: Math.min(rows[0], rows[1]), col: 0 };
  return undefined;
}

/**
 * A conditional format or a data validation. Its formulas are written for
 * the top-left cell of its range: a relative reference means "the cell this
 * far from the one the rule is applied to". So what a relative reference has
 * to follow is that corner, not the rows and columns it happens to name;
 * only references anchored with "$" follow those. A relative reference the
 * move of the corner takes off the sheet comes back in at the opposite edge,
 * as it does when the rule is evaluated (see ConditionalFormats): what the
 * file says and what the view showed before saving stay the same thing.
 * Returns undefined for an element that has no range of its own.
 *
 * UNSURE: Excel is known to wrap such references in conditional formats; data
 * validations are treated the same way here.
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
      texts[i] = from && to ? adjustAnchoredFormula(texts[i], name, edit, from, to, true) : adjustFormula(texts[i], name, edit);
    }
    list = next;
  }
  // Edits of other sheets move nothing here; only what points at them follows.
  const anchor = anchorOf(list);
  for (const edit of adjuster.foreign) {
    for (let i = 0; i < texts.length; i++) {
      texts[i] = anchor ? adjustAnchoredFormula(texts[i], name, edit, anchor, anchor, true) : adjustFormula(texts[i], name, edit);
    }
  }
  // Last, as where the rules are evaluated: the sheets get the names they have now.
  if (adjuster.rename) for (let i = 0; i < texts.length; i++) texts[i] = renameSheetsInFormula(texts[i], adjuster.rename);

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
const sameSheetName = (a: string | undefined, b: string) => a !== undefined && foldSheetName(a) === foldSheetName(b);

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

  // Last, the name itself: a source on a sheet that was deleted is gone with it.
  if (local === 'dataRef' && adjuster.rename) {
    const named = el.attrs.get('sheet');
    const now = named === undefined ? undefined : adjuster.rename(named);
    if (now === null) {
      adjuster.changed = true;
      return DROP;
    }
    if (now !== undefined && now !== named) {
      el.attrs.set('sheet', now);
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
    // Only formulas that name another edited sheet, or a sheet that has another name now, can change; most
    // elements name none. The text is looked at as a reader of the file sees it: a name may be spelled with
    // character references ("R&#38;D", "&apos;It&apos;&apos;s&apos;").
    const folded = foldSheetName(plainOf(raw));
    if (!adjuster.needles.some((needle) => folded.indexOf(needle) >= 0)) return raw;
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

/**
 * A child element of a worksheet that the model does not keep up to date itself (a
 * conditional format, the data validations) after the rows and columns moved in the
 * workbook, by the rules a saved sheet follows: '' when nothing of it is left, the
 * text itself when nothing in it moved. `own` are the edits of the sheet the element
 * belongs to, `foreign` those of the other sheets.
 */
export function adjustSheetElement(raw: string, sheetName: string, own: StructuralEdit[], foreign: StructuralEdit[]): string {
  if (own.length === 0 && foreign.length === 0) return raw;
  const name = /^<([^\s/>]+)/.exec(raw);
  return adjustRaw(raw, name ? name[1] : '', { sheetName, own, foreign, needles: nameNeedles(foreign.map((edit) => edit.sheetName)), changed: false });
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

export function colsXml(sheet: Sheet, prefix: string, originals: OriginalColumns[]): string {
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

/** A source element with some of its attributes set; an attribute given as undefined is removed. */
function withAttrs(raw: string, attrs: [name: string, value: string | undefined][]): string {
  const root = parseXml(raw).root;
  let changed = false;
  for (const [name, value] of attrs) {
    if (root.attrs.get(name) === value) continue;
    if (value === undefined) root.attrs.delete(name);
    else root.attrs.set(name, value);
    changed = true;
  }
  return changed ? elementText(root) : raw;
}

function filterColumnXml(column: FilterColumn, prefix: string): string {
  // As the file had it, at the place the column is now.
  if (column.raw !== undefined) return withAttrs(column.raw, [['colId', String(column.colId)]]);
  let inner = '';
  switch (column.kind) {
    case 'values': {
      for (const value of column.values || []) inner += `<${prefix}filter val="${escapeAttr(encodeXstring(String(value)))}"/>`;
      for (const group of column.dateGroups || []) {
        inner += `<${prefix}dateGroupItem year="${group.year}"`;
        for (const part of ['month', 'day', 'hour', 'minute', 'second'] as const) if (group[part] !== undefined) inner += ` ${part}="${group[part]}"`;
        inner += ` dateTimeGrouping="${escapeAttr(group.grouping)}"/>`;
      }
      const open = `<${prefix}filters${column.blank ? ' blank="1"' : ''}`;
      inner = inner === '' ? `${open}/>` : `${open}>${inner}</${prefix}filters>`;
      break;
    }
    case 'custom': {
      for (const condition of column.conditions || []) {
        inner += `<${prefix}customFilter${condition.operator && condition.operator !== 'equal' ? ` operator="${escapeAttr(condition.operator)}"` : ''} val="${escapeAttr(encodeXstring(String(condition.value)))}"/>`;
      }
      if (inner !== '') inner = `<${prefix}customFilters${column.and ? ' and="1"' : ''}>${inner}</${prefix}customFilters>`;
      break;
    }
    case 'top10':
      if (column.value !== undefined) inner = `<${prefix}top10${column.top === false ? ' top="0"' : ''}${column.percent ? ' percent="1"' : ''} val="${column.value}"/>`;
      break;
    case 'dynamic':
      if (column.dynamicType) inner = `<${prefix}dynamicFilter type="${escapeAttr(column.dynamicType)}"/>`;
      break;
    default:
    // A colour or an icon criterion only exists as the element it was read from.
  }
  if (inner === '' && !column.hiddenButton) return '';
  const open = `<${prefix}filterColumn colId="${column.colId}"${column.hiddenButton ? ' hiddenButton="1"' : ''}`;
  return inner === '' ? `${open}/>` : `${open}>${inner}</${prefix}filterColumn>`;
}

function sortStateXml(state: SortState, prefix: string): string {
  let out = `<${prefix}sortState`;
  if (state.columnSort) out += ' columnSort="1"';
  if (state.caseSensitive) out += ' caseSensitive="1"';
  out += ` ref="${rangeName(state.range)}"`;
  for (const [name, value] of state.attrs || []) out += ` ${name}="${escapeAttr(value)}"`;
  let inner = '';
  for (const condition of state.conditions) {
    const ref = rangeName(condition.range);
    if (condition.raw !== undefined) inner += withAttrs(condition.raw, [['ref', ref], ['descending', condition.descending ? '1' : undefined]]);
    else inner += `<${prefix}sortCondition${condition.descending ? ' descending="1"' : ''} ref="${ref}"/>`;
  }
  if (state.extra) inner += state.extra;
  return inner === '' ? `${out}/>` : `${out}>${inner}</${prefix}sortState>`;
}

/** The filter of the sheet from the model: its range, the criteria of its columns and how it was last sorted. */
export function autoFilterXml(sheet: Sheet, prefix: string): string {
  const filter = sheet.autoFilter;
  if (!filter) return '';
  let out = `<${prefix}autoFilter ref="${rangeName(filter.range)}"`;
  for (const [name, value] of filter.attrs || []) out += ` ${name}="${escapeAttr(value)}"`;
  let inner = '';
  for (const column of filter.columns) inner += filterColumnXml(column, prefix);
  if (filter.sortState) inner += sortStateXml(filter.sortState, prefix);
  if (filter.extra) inner += filter.extra;
  return inner === '' ? `${out}/>` : `${out}>${inner}</${prefix}autoFilter>`;
}

export function mergeCellsXml(sheet: Sheet, prefix: string): string {
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
  /** The name the file has for the sheet, which its raw content and the edits are matched with; the sheet's own name without it. */
  fileName?: string;
  /** Sheets of the file that have another name now, or are gone: applied to raw formulas after the edits. */
  rename?: SheetRenames;
  /** What <sheetView tabSelected> has to say; undefined leaves it as it is. */
  tabSelected?: boolean;
  /**
   * The source is the skeleton of a sheet that is not in the file: everything the model
   * holds is written from it, whatever was marked as changed.
   */
  fresh?: boolean;
}

interface SheetRenames {
  map: SheetNameMap;
  /** The names of the file that the map answers for. */
  names: string[];
}

interface ChildState {
  sheetData: number;
  linkPrefix: string | undefined;
  colsRaw: string | undefined;
  linksRaw: string | undefined;
  changed: boolean;
  /** What <sheetPr filterMode> has to say once the filter was changed; undefined leaves it as it is. */
  filterMode: boolean | undefined;
  /** What <sheetView tabSelected> has to say; undefined leaves it as it is. */
  tabSelected: boolean | undefined;
  /** Every row is written from the model. */
  full: boolean;
}

const isTrueAttr = (value: string | undefined) => value === '1' || value === 'true';

/** <sheetViews> with the sheet marked as the selected one, or as not selected; the text itself when it says so already. */
function selectTab(raw: string, selected: boolean): string {
  const root = parseXml(raw).root;
  let changed = false;
  for (const view of childrenOf(root, 'sheetView')) {
    if (isTrueAttr(view.attrs.get('tabSelected')) === selected) continue;
    if (selected) view.attrs.set('tabSelected', '1');
    else view.attrs.delete('tabSelected');
    changed = true;
  }
  return changed ? elementText(root) : raw;
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
    return writeSheetData(sc, s, start, sheet, writer, state.full || (adjuster !== undefined && adjuster.own.length > 0));
  }
  const currentRef = local === 'dimension' ? sc.attr('ref') : undefined;
  const source = sc.readRaw();
  const end = start + source.length;
  let raw = source;
  if (local === 'sheetViews' && state.tabSelected !== undefined) {
    // The source has the element: it is brought in line instead of being made.
    regenerated.delete(local);
    raw = selectTab(source, state.tabSelected);
    if (raw !== source) state.changed = true;
  }

  if (local === 'dimension') {
    const ref = rawOnly ? currentRef : usedRange(sheet);
    if (ref === currentRef) return { text: raw, end };
    const root = parseXml(raw).root;
    root.attrs.set('ref', ref as string);
    return { text: elementText(root), end };
  }

  if (local === 'sheetPr' && state.filterMode !== undefined) {
    // The source has the element: it is brought in line instead of being replaced.
    regenerated.delete(local);
    const text = withAttrs(raw, [['filterMode', state.filterMode ? '1' : undefined]]);
    if (text !== raw) return { text, end };
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

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/**
 * Raw XML with its entities and character references read: what to look for a sheet name in.
 * Not a parse: a reference that is none stays as it is, which hides no name.
 */
function plainOf(raw: string): string {
  if (raw.indexOf('&') < 0) return raw;
  return raw.replace(/&(#[xX][0-9a-fA-F]{1,6}|#[0-9]{1,7}|amp|lt|gt|quot|apos);/g, (match, ref: string) => {
    if (ref.charCodeAt(0) !== 35) return ENTITIES[ref];
    const hex = ref.charCodeAt(1) === 120 || ref.charCodeAt(1) === 88;
    const code = hex ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : match;
  });
}

/** Folded spellings (see foldSheetName()) under which sheet names may appear in the text of a part once its entities are read (see plainOf()). */
function nameNeedles(names: string[]): string[] {
  const out: string[] = [];
  for (const name of names) {
    const lower = foldSheetName(name);
    for (const spelling of [lower, lower.replace(/'/g, "''")]) {
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
  const fresh = !!options.fresh;
  // The raw content names sheets as the file does, this one included.
  const fileName = options.fileName !== undefined ? options.fileName : sheet.name;
  const own: StructuralEdit[] = rawOnly || fresh ? [] : (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: fileName }));
  const foreign = options.foreign || [];
  const rename = options.rename;
  const structural = own.length > 0;
  const adjuster: Adjuster | undefined =
    !fresh && (structural || foreign.length > 0 || rename)
      ? {
          sheetName: fileName,
          own,
          foreign,
          needles: nameNeedles(foreign.map((edit) => edit.sheetName).concat(rename ? rename.names : [])),
          rename: rename ? rename.map : undefined,
          changed: false,
        }
      : undefined;

  // These come from the model once it changed them, or once rows or columns moved; they are put at
  // the place the schema gives them when the source has none.
  const regenerated = new Map<string, string>();
  if (!rawOnly) {
    if (fresh || structural || sheet.columnsModified) regenerated.set('cols', sheet.columns.length > 0 ? COLS_MARKER : '');
    if (fresh || structural || sheet.mergesModified) regenerated.set('mergeCells', mergeCellsXml(sheet, prefix));
    if (fresh || structural || sheet.hyperlinksModified) regenerated.set('hyperlinks', sheet.hyperlinks.length > 0 ? LINKS_MARKER : '');
  }
  // A filter that was changed is written from the model; one that was not is copied, and only follows moved rows and columns.
  let filterMode: boolean | undefined;
  if (!rawOnly && (fresh || sheet.autoFilterModified)) {
    regenerated.set('autoFilter', autoFilterXml(sheet, prefix));
    filterMode = !!sheet.autoFilter && sheet.autoFilter.columns.some((column) => column.kind !== 'none');
    // Applications mark a sheet whose filter hides rows; without a <sheetPr> to hold the mark, one is made.
    if (filterMode) regenerated.set('sheetPr', `<${prefix}sheetPr filterMode="1"/>`);
  }
  // A sheet without views is not the selected one; one that has to be gets a view that says so.
  if (options.tabSelected) regenerated.set('sheetViews', `<${prefix}sheetViews><${prefix}sheetView tabSelected="1" workbookViewId="0"/></${prefix}sheetViews>`);
  const state: ChildState = { sheetData: 0, linkPrefix: relPrefix, colsRaw: undefined, linksRaw: undefined, changed: false, filterMode, tabSelected: options.tabSelected, full: fresh };
  /** Writes the regenerated elements that belong before the element of rank `rank` and were not in the source. */
  const flush = (rank: number) => {
    if (regenerated.size === 0) return;
    for (const local of ['sheetPr', 'sheetViews', 'cols', 'autoFilter', 'mergeCells', 'hyperlinks']) {
      const text = regenerated.get(local);
      if (text === undefined || WORKSHEET_ORDER.indexOf(local) >= rank) continue;
      out.push(text);
      regenerated.delete(local);
      if (text !== '') state.changed = true;
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

export const STYLESHEET_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

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
export function fontElement(font: Font, prefix: string): XmlElement {
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

export function fillElement(fill: Fill, prefix: string): XmlElement {
  // A gradient is kept as it was read.
  if (fill.type === 'gradient') return element(`${prefix}fill`, [], [cloneElement(fill.raw)]);
  const children: XmlNode[] = [];
  if (fill.fg) children.push(colorElement(`${prefix}fgColor`, fill.fg));
  if (fill.bg) children.push(colorElement(`${prefix}bgColor`, fill.bg));
  return element(`${prefix}fill`, [], [element(`${prefix}patternFill`, [['patternType', fill.pattern]], children)]);
}

export function borderElement(border: Border, prefix: string): XmlElement {
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
export function formatElement(format: CellFormat, prefix: string): XmlElement {
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

/** What happened to the other sheets of the workbook, for a part that names them. */
interface OtherSheets {
  /** Rows and columns moved in the other sheets, under the names the file has for those. */
  foreign: StructuralEdit[];
  /** What the sheet names of the file stand for now, when sheets were renamed or deleted. */
  rename?: SheetNameMap;
}

/**
 * Shapes and connectors may show the content of a cell: textlink="Sheet1!$A$5" is a formula. It
 * follows the rows and columns of its own sheet and of the sheets it names (all named as the
 * file names them), and then names the sheets as they are called now.
 */
function adjustTextLinks(el: XmlElement, edits: SheetStructuralEdit[], sheetName: string, others: OtherSheets | undefined): boolean {
  let changed = false;
  const link = el.attrs.get('textlink');
  if (link !== undefined && link !== '') {
    let next = link;
    for (const edit of edits) next = adjustFormula(next, sheetName, { ...edit, sheetName });
    if (others) {
      for (const edit of others.foreign) next = adjustFormula(next, sheetName, edit);
      if (others.rename) next = renameSheetsInFormula(next, others.rename);
    }
    if (next !== link) {
      el.attrs.set('textlink', next);
      changed = true;
    }
  }
  for (const node of el.children) if (node.type === 'element' && adjustTextLinks(node, edits, sheetName, others)) changed = true;
  return changed;
}

/** A drawing part with its anchors moved and its text links brought up to date; the same bytes when nothing changed. */
function adjustDrawing(bytes: Uint8Array, edits: SheetStructuralEdit[], sheetName: string, others?: OtherSheets): Uint8Array {
  // Without rows or columns moved in the sheet itself, only a link to a cell can change.
  if (edits.length === 0 && (!others || decodeSource(bytes).indexOf('textlink') < 0)) return bytes;
  const doc = parseXml(bytes);
  const moved = adjustAnchorsIn(doc.root, edits);
  let linked: boolean;
  try {
    linked = adjustTextLinks(doc.root, edits, sheetName, others);
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
/** The elements of a table part that hold formulas. */
const TABLE_FORMULAS = ['calculatedColumnFormula', 'totalsRowFormula'];

/**
 * The formulas of a table (of a calculated column, of the totals row) after rows and columns
 * moved in the other sheets and sheets were renamed or deleted. Returns whether any changed.
 * (Rows and columns do not move in the sheet of a table itself: the model refuses that.)
 */
function adjustTableFormulas(el: XmlElement, sheetName: string, others: OtherSheets, depth = 0): boolean {
  let changed = false;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (TABLE_FORMULAS.indexOf(localName(node.name)) >= 0 && isTextOnly(node)) {
      const text = plainText(node);
      let next = text;
      for (const edit of others.foreign) next = adjustFormula(next, sheetName, edit);
      if (others.rename) next = renameSheetsInFormula(next, others.rename);
      if (next !== text) {
        node.children = [{ type: 'text', value: next }];
        changed = true;
      }
    } else if (depth < 8 && adjustTableFormulas(node, sheetName, others, depth + 1)) changed = true;
  }
  return changed;
}

const ANCHORED_PARTS: Record<string, (bytes: Uint8Array, edits: SheetStructuralEdit[], sheetName: string, others?: OtherSheets) => Uint8Array> = {
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

const isFilterName = (name: string) => name.toLowerCase() === FILTER_NAME.toLowerCase();

/** A sheet as the workbook part of the file lists it. */
interface FileSheet {
  element: XmlElement;
  name: string;
  /** The part the sheet is in; '' when the file does not say. */
  part: string;
  /** Its entry in the relationships of the workbook part. */
  rel: XmlElement | undefined;
}

/** The sheets of the file in its order, each with the part its relationship leads to. */
function fileSheetsOf(root: XmlElement, relationships: XmlElement[], main: string): FileSheet[] {
  const byId = new Map<string, XmlElement>();
  for (const rel of relationships) byId.set(rel.attrs.get('Id') || '', rel);
  const list = childNamed(root, 'sheets');
  const out: FileSheet[] = [];
  for (const el of list ? childrenNamed(list, 'sheet') : []) {
    let id: string | undefined;
    for (const [name, value] of el.attrs) {
      if (name.indexOf(':') >= 0 && localName(name) === 'id') {
        id = value;
        break;
      }
    }
    const rel = id === undefined ? undefined : byId.get(id);
    const internal = rel && rel.attrs.get('TargetMode') !== 'External';
    out.push({ element: el, name: el.attrs.get('name') || '', part: rel && internal ? resolveTarget(main, rel.attrs.get('Target') || '') : '', rel });
  }
  return out;
}

/** What a save does to the workbook part beyond the defined names. */
interface BookChange {
  file: FileSheet[];
  /** Sheets have other names than in the file. */
  renamed: boolean;
  /** The sheets are not the ones of the file, one for one (a sheet was added or deleted): whatever counts sheets is brought up to date. */
  reordered: boolean;
  /** The relationship made for each sheet that is not in the file. */
  relIds: Map<Sheet, string>;
  /** The namespace relationship ids are in, in this flavour of the format (transitional or strict). */
  relNamespace: string;
  /** Index of the sheet the file opens on. */
  active: number;
  /** Stored results may be stale: the application is asked to calculate when it opens the file. */
  recalculate: boolean;
}

const MAX_SHEET_ID = 65534;

/**
 * Brings <sheets> in line with the model: the elements of the file in the order of the model,
 * each with the name its sheet has now and everything else it had; none for a sheet that was
 * deleted; a new one for a sheet that is not in the file. Returns the sheetId of each sheet.
 */
function updateSheetList(root: XmlElement, workbook: Workbook, change: BookChange): number[] {
  const list = childNamed(root, 'sheets');
  if (!list) throw new Error('Cannot write the workbook: its part has no list of sheets');
  const listed = new Map<string, XmlElement>();
  for (const sheet of change.file) if (sheet.part !== '') listed.set(sheet.part, sheet.element);
  const kept = workbook.sheets.map((sheet) => (sheet.part === '' ? undefined : listed.get(sheet.part)));
  const taken = new Set<number>();
  for (const el of kept) if (el) taken.add(parseUint(el.attrs.get('sheetId')));

  // The prefix relationship ids have where the root declares one; a new element declares its own otherwise.
  let idAttr = 'r:id';
  let declared = false;
  for (const [name, value] of root.attrs) {
    if (!declared && name.slice(0, 6) === 'xmlns:' && (value === NS_REL || value === NS_REL_STRICT)) {
      idAttr = `${name.slice(6)}:id`;
      declared = true;
    }
  }
  const sample = change.file.length > 0 ? change.file[0].element : undefined;
  const prefix = sample ? prefixOf(sample) : prefixOf(list);

  const ids: number[] = [];
  const elements = workbook.sheets.map((sheet, index) => {
    const el = kept[index];
    if (el) {
      if (el.attrs.get('name') !== sheet.name) el.attrs.set('name', sheet.name);
      ids.push(parseUint(el.attrs.get('sheetId')));
      return el;
    }
    if (sheet.part !== '') throw new Error(`Cannot write sheet "${sheet.name}": the file has no sheet in ${sheet.part}`);
    // The id the model gave the sheet, unless a sheet of the file has it.
    let id = sheet.sheetId;
    if (!(id >= 1 && id <= MAX_SHEET_ID && Math.floor(id) === id) || taken.has(id)) {
      id = 1;
      while (taken.has(id)) id++;
    }
    taken.add(id);
    ids.push(id);
    const attrs: Attrs = [
      ['name', sheet.name],
      ['sheetId', String(id)],
    ];
    if (sheet.state === 'hidden' || sheet.state === 'veryHidden') attrs.push(['state', sheet.state]);
    if (!declared) attrs.push(['xmlns:r', change.relNamespace]);
    attrs.push([idAttr, change.relIds.get(sheet) || '']);
    return element(`${prefix}sheet`, attrs);
  });
  if (!change.reordered) return ids;

  // Whatever stands before an element (white space, a comment) stays before it.
  const leading = new Map<XmlNode, XmlNode[]>();
  const sheetElements = new Set<XmlNode>(change.file.map((sheet) => sheet.element));
  const others: XmlNode[] = [];
  let pending: XmlNode[] = [];
  for (const node of list.children) {
    if (sheetElements.has(node)) {
      leading.set(node, pending);
      pending = [];
    } else if (node.type === 'element') {
      // Not a sheet, whatever it is: kept after the sheets.
      others.push(...pending, node);
      pending = [];
    } else pending.push(node);
  }
  const children: XmlNode[] = [];
  for (const el of elements) children.push(...(leading.get(el) || []), el);
  list.children = children.concat(others, pending);
  return ids;
}

/**
 * The views of the workbook after sheets were added or deleted: the sheet shown is the active
 * one of the model, and the first tab shown stays the sheet it was (or what is at its place).
 */
function updateBookViews(root: XmlElement, workbook: Workbook, change: BookChange, sheetIds: number[]): void {
  const prefix = prefixOf(root);
  const count = workbook.sheets.length;
  const indexOfPart = new Map<string, number>();
  workbook.sheets.forEach((sheet, index) => {
    if (sheet.part !== '') indexOfPart.set(sheet.part, index);
  });
  let views = childNamed(root, 'bookViews');
  if (!views && change.active > 0) {
    views = element(`${prefix}bookViews`, [], [element(`${prefix}workbookView`)]);
    insertInOrder(root, views, WORKBOOK_ORDER);
  }
  for (const view of views ? childrenNamed(views, 'workbookView') : []) {
    // UNSURE: a second window of the workbook shows its own sheet; here every window opens on the active sheet of the model.
    const stated = view.attrs.get('activeTab');
    if ((stated === undefined ? '0' : stated) !== String(change.active)) {
      if (change.active === 0) view.attrs.delete('activeTab');
      else view.attrs.set('activeTab', String(change.active));
    }
    const first = view.attrs.get('firstSheet');
    if (first !== undefined) {
      const was = parseUint(first);
      const sheet = was >= 0 ? change.file[was] : undefined;
      const now = sheet ? indexOfPart.get(sheet.part) : undefined;
      const next = now !== undefined ? now : Math.max(0, Math.min(was, count - 1));
      if (String(next) !== first) view.attrs.set('firstSheet', String(next));
    }
  }
  // A custom view names the sheet it shows by its sheetId; one that showed a deleted sheet shows the active one.
  const custom = childNamed(root, 'customWorkbookViews');
  for (const view of custom ? childrenNamed(custom, 'customWorkbookView') : []) {
    const shown = view.attrs.get('activeSheetId');
    if (shown !== undefined && sheetIds.indexOf(parseUint(shown)) < 0 && sheetIds[change.active] >= 0) view.attrs.set('activeSheetId', String(sheetIds[change.active]));
  }
}

/** The scope of a <definedName> whose sheet was deleted, and of one whose sheet the file does not have. */
const SCOPE_GONE = -1;
const SCOPE_UNKNOWN = -2;

/**
 * Brings workbook.xml up to date: the list of sheets and what counts them follow the model, defined
 * names follow the model, and formulas are recalculated on load.
 */
function updateWorkbookXml(doc: XmlDocument, workbook: Workbook, change: BookChange): void {
  const root = doc.root;
  const prefix = prefixOf(root);
  if (change.renamed || change.reordered) {
    const sheetIds = updateSheetList(root, workbook, change);
    if (change.reordered) updateBookViews(root, workbook, change, sheetIds);
  }

  // localSheetId counts the sheets of the file; the model counts its own. A name is found by
  // the sheet it belongs to: the id of the file, the part of that sheet, the index it has now.
  const indexOfPart = new Map<string, number>();
  workbook.sheets.forEach((sheet, index) => {
    if (sheet.part !== '') indexOfPart.set(sheet.part, index);
  });
  let container: XmlElement | undefined = childrenNamed(root, 'definedNames')[0];
  const elements = container ? childrenNamed(container, 'definedName') : [];
  const scopes = new Map<XmlElement, number | undefined>();
  for (const el of elements) {
    const local = el.attrs.get('localSheetId');
    let scope = local !== undefined && /^\d{1,6}$/.test(local) ? Number(local) : undefined;
    if (scope !== undefined && change.reordered) {
      const sheet = change.file[scope];
      const now = sheet && sheet.part !== '' ? indexOfPart.get(sheet.part) : undefined;
      scope = !sheet || sheet.part === '' ? SCOPE_UNKNOWN : now === undefined ? SCOPE_GONE : now;
    }
    scopes.set(el, scope);
  }
  const used = new Set<XmlElement>();
  const matches = (el: XmlElement, name: string, localSheet: number | undefined) => !used.has(el) && (el.attrs.get('name') || '') === name && scopes.get(el) === localSheet;
  workbook.definedNames.forEach((name, index) => {
    // The lists normally line up; otherwise the name and its scope identify the element.
    let target: XmlElement | undefined = elements[index];
    if (!target || !matches(target, name.name, name.localSheet)) target = elements.find((el) => matches(el, name.name, name.localSheet));
    // A name of a sheet the file does not have (a damaged id) is known by its name alone.
    if (!target && name.localSheet !== undefined) target = elements.find((el) => !used.has(el) && scopes.get(el) === SCOPE_UNKNOWN && (el.attrs.get('name') || '') === name.name);
    if (!target) {
      // The only name the model adds is the range of a filter that was switched on; any other
      // name that is not in the file is a mistake of the caller.
      if (!isFilterName(name.name)) throw new Error(`Cannot write the workbook: the defined name "${name.name}" is not in the file`);
      const attrs: Attrs = [['name', name.name]];
      if (name.localSheet !== undefined) attrs.push(['localSheetId', String(name.localSheet)]);
      if (name.hidden) attrs.push(['hidden', '1']);
      target = element(`${prefix}definedName`, attrs, [{ type: 'text', value: name.value }]);
      if (!container) {
        container = element(`${prefix}definedNames`);
        insertInOrder(root, container, WORKBOOK_ORDER);
      }
      container.children.push(target);
      used.add(target);
      return;
    }
    used.add(target);
    if (textOf(target) !== name.value) target.children = [{ type: 'text', value: name.value }];
    const scope = scopes.get(target);
    if (name.localSheet !== undefined && scope !== undefined && scope >= 0 && target.attrs.get('localSheetId') !== String(name.localSheet)) target.attrs.set('localSheetId', String(name.localSheet));
  });
  // The range of a filter that was switched off goes with it, and the names of a deleted sheet with their sheet; an empty list is not valid.
  const gone = (el: XmlElement) => !used.has(el) && (isFilterName(el.attrs.get('name') || '') || scopes.get(el) === SCOPE_GONE);
  if (container && elements.some(gone)) {
    const list = container;
    list.children = list.children.filter((node) => node.type !== 'element' || localName(node.name) !== 'definedName' || !gone(node));
    if (!list.children.some((node) => node.type === 'element')) root.children = root.children.filter((node) => node !== list);
  }

  if (!change.recalculate) return;
  // Results stored in the file may be stale after an edit; the application recalculates when it opens the file.
  let calcPr = childrenNamed(root, 'calcPr')[0];
  if (!calcPr) {
    calcPr = element(`${prefix}calcPr`);
    insertInOrder(root, calcPr, WORKBOOK_ORDER);
  }
  calcPr.attrs.set('fullCalcOnLoad', '1');
}

/**
 * The extended properties part (docProps/app.xml) lists the names of the sheets. The list is
 * brought up to date where that is safe: where one of its groups holds exactly the names the
 * worksheets have in the file. Returns undefined to leave the part as it is.
 */
function updateAppXml(bytes: Uint8Array, fileNames: string[], names: string[]): Uint8Array | undefined {
  if (fileNames.length === 0 || names.length === 0) return undefined;
  let doc: XmlDocument;
  try {
    doc = parseXml(bytes);
  } catch (error) {
    return undefined;
  }
  const elementsOf = (el: XmlElement | undefined) => (el ? el.children.filter((node): node is XmlElement => node.type === 'element') : []);
  const titlesVector = childNamed(childNamed(doc.root, 'TitlesOfParts'), 'vector');
  const pairsVector = childNamed(childNamed(doc.root, 'HeadingPairs'), 'vector');
  if (!titlesVector || !pairsVector) return undefined;
  const titles = elementsOf(titlesVector);
  if (titles.some((el) => localName(el.name) !== 'lpstr' || el.children.some((node) => node.type !== 'text'))) return undefined;
  // Pairs of a heading and the number of titles under it, in the order of the titles.
  const variants = elementsOf(pairsVector);
  let offset = 0;
  let found: { at: number; count: XmlElement } | undefined;
  for (let i = 0; i + 1 < variants.length; i += 2) {
    const count = elementsOf(variants[i + 1])[0];
    const size = count && localName(count.name) === 'i4' ? parseUint(plainText(count).trim()) : -1;
    if (size < 0) return undefined;
    if (!found && size === fileNames.length && fileNames.every((name, k) => titles[offset + k] !== undefined && plainText(titles[offset + k]) === name)) found = { at: offset, count };
    offset += size;
  }
  if (!found || offset !== titles.length || variants.length % 2 !== 0) return undefined;
  const at = found.at;
  const made = names.map((name) => element(titles[at].name, [], [{ type: 'text', value: name }]));
  const next = titles.slice(0, at).concat(made, titles.slice(at + fileNames.length));
  titlesVector.children = next;
  if (titlesVector.attrs.has('size')) titlesVector.attrs.set('size', String(next.length));
  found.count.children = [{ type: 'text', value: String(names.length) }];
  return encoder.encode(serializeXml(doc));
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

/** "/xl/_rels/workbook.xml.rels": a part that holds the relationships of another. */
const isRelsPart = (part: string) => /(^|\/)_rels\/[^/]*\.rels$/i.test(part);

const relsPartOf = (part: string) => `${part.slice(0, part.lastIndexOf('/'))}/_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;

interface SaveState {
  /** The bytes each rewritten part had when the file was opened. null stands for a part the file did not have. */
  originals: Map<string, Uint8Array | null>;
  /**
   * Per sheet part: the parts around it (drawings, comments, relationships)
   * that the last save wrote differently from the original. A save that no
   * longer changes them has to put the original back. Under BOOK: the parts
   * of the workbook itself (its part, its relationships, the list of titles).
   */
  moved: Map<string, string[]>;
  /**
   * Parts of the file that the last save took out of the package (those of deleted sheets), each
   * with its content type. Their bytes are in `originals`: a sheet that comes back by undo comes
   * back with every part it had.
   */
  removed: Map<string, string | undefined>;
  /**
   * The sheets that are not in the file, each with the part the last save wrote it to. The model
   * keeps saying that such a sheet has no part (Sheet.part stays ''): every save starts from the
   * file as it was opened, in which the sheet is not.
   */
  added: Map<Sheet, string>;
}

const SAVES = new WeakMap<Workbook, SaveState>();

/** The key of the workbook's own parts in SaveState.moved; no sheet part has it. */
const BOOK = '';

/** More parts than any package has; following relationships stops there. */
const MAX_RELATED_PARTS = 1000000;

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

/** Index of the sheet a saved file opens on: the active one of the model, or the first visible one when that is hidden. */
function activeSheetOf(workbook: Workbook): number {
  const sheets = workbook.sheets;
  let active = workbook.activeSheet;
  if (!(active >= 0 && active < sheets.length && Math.floor(active) === active)) active = 0;
  if (sheets[active] && sheets[active].state !== 'visible') {
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].state === 'visible') return i;
    }
  }
  return active;
}

/**
 * Brings the package up to date with the model. Does nothing when nothing was modified.
 *
 * Sheets that were inserted, deleted or renamed (Workbook.sheetsModified):
 * - A sheet that is not in the file (Sheet.part === '') is written to a part of its own, the
 *   first free /xl/worksheets/sheetN.xml. The model is not told: the part is remembered here,
 *   and the next save writes the sheet to the same part. Once the sheet is gone from the model
 *   (undo, or deleted), the next save takes the part out again.
 * - A sheet of the file that is gone from the model loses its part and every part only it
 *   leads to (its relationships, drawing, comments, printer settings, pictures no other part
 *   shows). Their bytes are kept: when the sheet is back in the model, the next save puts them
 *   back as they were.
 * - Content copied from the file names sheets as the file does. Rows and columns moved are
 *   matched with those names first; then the sheets get the names they have now.
 */
export async function writeXlsx(pkg: OpcPackage, workbook: Workbook): Promise<void> {
  const modifiedSheets = workbook.sheets.filter((sheet) => sheet.modified && sheet.kind !== 'chartsheet');
  const stylesModified = !!workbook.styles.modified;
  const sheetsModified = !!workbook.sheetsModified || workbook.sheets.some((sheet) => sheet.part === '');
  if (modifiedSheets.length === 0 && !stylesModified && !sheetsModified) return;

  let saved = SAVES.get(workbook);
  if (!saved) {
    saved = { originals: new Map(), moved: new Map(), removed: new Map(), added: new Map() };
    SAVES.set(workbook, saved);
  }
  const { originals, moved, removed, added } = saved;
  // The first time a part is about to be rewritten, the package still holds it as the file had it.
  const originalOf = async (part: string): Promise<Uint8Array | undefined> => {
    const kept = originals.get(part);
    if (kept !== undefined) return kept === null ? undefined : kept;
    const bytes = pkg.hasPart(part) ? await pkg.readPart(part) : undefined;
    originals.set(part, bytes === undefined ? null : bytes);
    return bytes;
  };
  // The same for a part that is only looked at: its bytes are not kept.
  const peek = async (part: string): Promise<Uint8Array | undefined> => {
    const kept = originals.get(part);
    if (kept !== undefined) return kept === null ? undefined : kept;
    return pkg.hasPart(part) ? pkg.readPart(part) : undefined;
  };
  const keep = (part: string, bytes: Uint8Array) => {
    if (!originals.has(part)) originals.set(part, bytes);
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
  let highest = 0;
  for (const rel of relationships) {
    const match = /^rId(\d{1,9})$/.exec(rel.attrs.get('Id') || '');
    if (match && Number(match[1]) > highest) highest = Number(match[1]);
  }
  // The same flavour of relationship types as the rest of the file (transitional or strict).
  const sampleRel = withType('/worksheet');
  const sampleType = sampleRel ? typeOf(sampleRel) : '';
  const relBase = sampleType ? sampleType.slice(0, sampleType.lastIndexOf('/')) : NS_REL;
  /** Adds a relationship from the workbook to one of its parts and returns its id. */
  const relate = (kind: string, part: string): string => {
    const id = `rId${++highest}`;
    rels.root.children.push(
      element(`${prefixOf(rels.root)}Relationship`, [
        ['Id', id],
        ['Type', `${relBase}/${kind}`],
        ['Target', part.slice(0, slash + 1) === main.slice(0, slash + 1) ? part.slice(slash + 1) : part],
      ]),
    );
    relsChanged = true;
    return id;
  };

  const mainBytes = await originalOf(main);
  if (!mainBytes) throw new Error('Cannot write the workbook: its part is missing');
  const book = parseXml(mainBytes);
  const file = fileSheetsOf(book.root, relationships, main);
  const listed = new Map<string, FileSheet>();
  for (const sheet of file) if (sheet.part !== '') listed.set(sheet.part, sheet);

  // What became of the sheets of the file. A sheet is known by its part, which never changes.
  let reordered = false;
  let renamed = false;
  const present = new Set<string>();
  if (sheetsModified) {
    reordered = workbook.sheets.length !== file.length;
    workbook.sheets.forEach((sheet, index) => {
      if (sheet.part === '') {
        reordered = true;
        return;
      }
      const inFile = listed.get(sheet.part);
      if (!inFile) throw new Error(`Cannot write sheet "${sheet.name}": the file has no sheet in ${sheet.part}`);
      present.add(sheet.part);
      if (!file[index] || file[index].part !== sheet.part) reordered = true;
      if (inFile.name !== sheet.name) renamed = true;
    });
  }
  const gone = sheetsModified ? file.filter((sheet) => sheet.part !== '' && !present.has(sheet.part)) : [];
  /** The name the file has for a sheet; undefined for a sheet that is not in the file. */
  const fileNameOf = (sheet: Sheet): string | undefined => {
    if (sheet.part === '') return undefined;
    const inFile = listed.get(sheet.part);
    return inFile ? inFile.name : sheet.name;
  };
  // Content copied from the file names sheets as the file does: what those names stand for now.
  let rename: SheetRenames | undefined;
  if (renamed || gone.length > 0) {
    const names = new Map<string, string | null>();
    const list: string[] = [];
    for (const sheet of gone) {
      names.set(foldSheetName(sheet.name), null);
      list.push(sheet.name);
    }
    for (const sheet of workbook.sheets) {
      const name = fileNameOf(sheet);
      if (name === undefined || name === sheet.name) continue;
      names.set(foldSheetName(name), sheet.name);
      list.push(name);
    }
    const map: SheetNameMap = (name) => names.get(foldSheetName(name));
    // Which sheet takes the place of a deleted one at the end of a 3D reference is what the model noted when it was deleted.
    const noted = sheetNameMap(workbook);
    if (noted && noted.runs) map.runs = noted.runs;
    rename = { map, names: list };
  }
  const active = activeSheetOf(workbook);

  // Everything is produced before anything is put into the package: a failure leaves the package as it was.
  const writes: PartWrite[] = [];

  if (stylesModified) {
    const stylesRel = withType('/styles');
    const stylesPart = stylesRel ? resolveTarget(main, stylesRel.attrs.get('Target') || '') : `${main.slice(0, slash)}/styles.xml`;
    writes.push([stylesPart, writeStyles(await originalOf(stylesPart), workbook.styles), STYLES_CONTENT_TYPE]);
    if (!stylesRel) relate('styles', stylesPart);
  }

  // Rows and columns moved in one sheet change what the others say about it. What the others say is
  // content of the file, which knows the sheet by the name the file has for it; nothing in the file
  // refers to a sheet that is not in it.
  const editsOf = (sheet: Sheet): StructuralEdit[] => {
    const name = fileNameOf(sheet);
    return name === undefined ? [] : (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: name }));
  };
  const edited = workbook.sheets.filter((sheet) => sheet.kind !== 'chartsheet' && sheet.part !== '' && sheet.structuralEdits && sheet.structuralEdits.length > 0);
  const foreignOf = (sheet: Sheet): StructuralEdit[] => {
    let out: StructuralEdit[] = [];
    for (const other of edited) if (other !== sheet) out = out.concat(editsOf(other));
    return out;
  };

  const movedNow = new Map<string, string[]>();
  for (let index = 0; index < workbook.sheets.length; index++) {
    const sheet = workbook.sheets[index];
    if (sheet.kind === 'chartsheet' || sheet.part === '') continue;
    const fileName = fileNameOf(sheet) as string;
    const foreign = foreignOf(sheet);
    const before = moved.get(sheet.part) || [];
    // Once sheets were added or deleted, the one sheet marked as selected is the active one.
    const tabSelected = reordered && !!sheet.view.tabSelected !== (index === active) ? index === active : undefined;
    if (!sheet.modified && foreign.length === 0 && !rename && tabSelected === undefined && before.length === 0) continue;
    const rewritten: string[] = [];
    const sheetRelsPart = relsPartOf(sheet.part);
    const others: OtherSheets | undefined = foreign.length > 0 || rename ? { foreign, rename: rename ? rename.map : undefined } : undefined;

    // Tables name their columns after their header cells; the two have to agree. Their formulas
    // name sheets as the file does and follow what happened to those.
    let overrides: Map<number, Map<number, Cell>> | undefined;
    const headersChanged = !!sheet.modified && !!sheet.changedRows && sheet.changedRows.size > 0;
    const tableRels = headersChanged || others ? await peek(sheetRelsPart) : undefined;
    for (const rel of tableRels ? childrenNamed(parseXml(tableRels).root, 'Relationship') : []) {
      if (!endsWith(typeOf(rel), '/table') || !internal(rel)) continue;
      const part = resolveTarget(sheet.part, rel.attrs.get('Target') || '');
      const original = await peek(part);
      if (!original) continue;
      const doc = parseXml(original);
      const headers = headersChanged ? syncTableColumns(doc.root, sheet, workbook) : undefined;
      const formulas = others ? adjustTableFormulas(doc.root, fileName, others) : false;
      if ((headers && headers.changed) || formulas) {
        keep(part, original);
        writes.push([part, doc]);
        rewritten.push(part);
      }
      if (headers && headers.cells.size > 0) {
        if (!overrides) overrides = new Map();
        const row = overrides.get(headers.row) || new Map<number, Cell>();
        for (const [col, cell] of headers.cells) row.set(col, cell);
        overrides.set(headers.row, row);
      }
    }

    if (!sheet.modified) {
      // A sheet that was not modified itself: only its raw content that names an edited or a renamed
      // sheet changes (the source of a validation list on another sheet, say). Its cells were adjusted
      // through the model, which marked the sheet as modified if any of them changed.
      const bytes = await peek(sheet.part);
      const next = bytes && (foreign.length > 0 || rename || tabSelected !== undefined) ? rewriteSheet(bytes, sheet, workbook, { foreign, rawOnly: true, fileName, rename, tabSelected }) : undefined;
      if (bytes && next) {
        // From here on the package no longer holds the part as the file had it.
        keep(sheet.part, bytes);
        writes.push([sheet.part, next]);
        rewritten.push(sheet.part);
      }
    } else {
      const bytes = await originalOf(sheet.part);
      if (!bytes) throw new Error(`Cannot write sheet "${sheet.name}": its part ${sheet.part} is missing`);
      writes.push([sheet.part, rewriteSheet(bytes, sheet, workbook, { foreign, overrides, fileName, rename, tabSelected }) as Uint8Array]);
      rewritten.push(sheet.part);
    }

    // After rows or columns moved: the pictures and comments of the sheet follow, and the
    // relationships of links that are gone are removed. After rows or columns moved in another
    // sheet, or sheets were renamed or deleted: the shapes that show a cell of such a sheet follow.
    const edits = sheet.modified && sheet.structuralEdits && sheet.structuralEdits.length > 0 ? sheet.structuralEdits : undefined;
    const sheetRelsBytes = edits || others ? await peek(sheetRelsPart) : undefined;
    if (sheetRelsBytes) {
      const sheetRels = parseXml(sheetRelsBytes);
      const linkIds = new Set<string>();
      for (const link of sheet.hyperlinks) if (link.rId !== undefined) linkIds.add(link.rId);
      let dropped = false;
      for (const rel of childrenNamed(sheetRels.root, 'Relationship')) {
        const type = typeOf(rel);
        const kind = type.slice(type.lastIndexOf('/') + 1);
        if (kind === 'hyperlink') {
          if (edits && !linkIds.has(rel.attrs.get('Id') || '')) {
            sheetRels.root.children = sheetRels.root.children.filter((node) => node !== rel);
            dropped = true;
          }
          continue;
        }
        const adjust = Object.prototype.hasOwnProperty.call(ANCHORED_PARTS, kind) ? ANCHORED_PARTS[kind] : undefined;
        // Of these parts only a drawing names sheets.
        if (!adjust || !internal(rel) || (!edits && kind !== 'drawing')) continue;
        const part = resolveTarget(sheet.part, rel.attrs.get('Target') || '');
        const original = await peek(part);
        if (!original) continue;
        if (edits) {
          // One sheet's rows cannot move a part that another sheet shows as well.
          for (const other of workbook.sheets) {
            if (other === sheet || other.kind === 'chartsheet' || other.part === '') continue;
            for (const otherRel of await pkg.relationships(other.part)) {
              if (otherRel.part === part) throw new Error(`Cannot write sheet "${sheet.name}": ${part} is shared with sheet "${other.name}"`);
            }
          }
        }
        const adjusted = adjust(original, edits || [], fileName, others);
        if (adjusted !== original) {
          keep(part, original);
          writes.push([part, adjusted]);
          rewritten.push(part);
        }
      }
      if (dropped) {
        keep(sheetRelsPart, sheetRelsBytes);
        writes.push([sheetRelsPart, sheetRels]);
        rewritten.push(sheetRelsPart);
      }
    }
    // What an earlier save moved and this one does not goes back to what the file had.
    for (const part of before) {
      const original = originals.get(part);
      if (rewritten.indexOf(part) < 0 && original) writes.push([part, original]);
    }
    movedNow.set(sheet.part, rewritten);
  }

  // Sheets that are not in the file: each is written to a part of its own, the one the last save gave it.
  const addedNow = new Map<Sheet, string>();
  const relIds = new Map<Sheet, string>();
  let taken: Set<string> | undefined;
  for (let index = 0; index < workbook.sheets.length; index++) {
    const sheet = workbook.sheets[index];
    if (sheet.part !== '') continue;
    if (sheet.kind !== 'worksheet') throw new Error(`Cannot write sheet "${sheet.name}": a ${sheet.kind} that is not in the file cannot be written`);
    let part = added.get(sheet);
    if (part === undefined) {
      if (!taken) {
        // Part names are told apart without regard to case. A part that was taken out may come back.
        taken = new Set();
        for (const name of pkg.partNames()) taken.add(name.toLowerCase());
        for (const [name, bytes] of originals) if (bytes !== null) taken.add(name.toLowerCase());
        for (const name of removed.keys()) taken.add(name.toLowerCase());
        for (const name of added.values()) taken.add(name.toLowerCase());
      }
      for (let n = 1; part === undefined; n++) {
        const candidate = `${main.slice(0, slash)}/worksheets/sheet${n}.xml`;
        // Relationships left behind under the name would become those of the new sheet.
        if (!taken.has(candidate.toLowerCase()) && !taken.has(relsPartOf(candidate).toLowerCase())) part = candidate;
      }
      taken.add(part.toLowerCase());
    }
    const rootPrefix = prefixOf(book.root);
    const mainNamespace = book.root.attrs.get(rootPrefix === '' ? 'xmlns' : `xmlns:${rootPrefix.slice(0, -1)}`) || NS_MAIN;
    const skeleton =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${escapeAttr(mainNamespace)}" xmlns:r="${escapeAttr(relBase)}">` +
      `<dimension ref="A1"/><sheetViews><sheetView${index === active ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews><sheetData/></worksheet>`;
    // A link to another file belongs to the relationships of the sheet it was made in; only places in the workbook are written.
    const hyperlinks = sheet.hyperlinks.filter((link) => link.location !== undefined).map((link) => (link.rId === undefined ? link : { ...link, rId: undefined }));
    const bytes = rewriteSheet(encoder.encode(skeleton), { ...sheet, hyperlinks }, workbook, { fresh: true }) as Uint8Array;
    if (!originals.has(part)) originals.set(part, null);
    writes.push([part, bytes, WORKSHEET_CONTENT_TYPE]);
    addedNow.set(sheet, part);
    relIds.set(sheet, relate('worksheet', part));
  }
  const deletions: string[] = [];
  for (const [sheet, part] of added) if (addedNow.get(sheet) !== part) deletions.push(part);

  // Sheets of the file that are gone: their parts, and what only they lead to, leave the package.
  const wanted = new Set<string>();
  if (gone.length > 0) {
    const actual = new Map<string, string>();
    for (const name of removed.keys()) actual.set(name.toLowerCase(), name);
    for (const name of pkg.partNames()) actual.set(name.toLowerCase(), name);
    /** The parts a part leads to in the file as it was opened. */
    const targetsOf = async (source: string): Promise<string[]> => {
      const bytes = await peek(relsPartOf(source));
      const out: string[] = [];
      if (!bytes) return out;
      let doc: XmlDocument;
      try {
        doc = parseXml(bytes);
      } catch (error) {
        // Relationships that cannot be read lead nowhere: what they may point at stays.
        return out;
      }
      for (const rel of childrenNamed(doc.root, 'Relationship')) {
        if (!internal(rel)) continue;
        const target = actual.get(resolveTarget(source, rel.attrs.get('Target') || '').toLowerCase());
        // A relationships part is not content: it goes with the part it belongs to and with nothing else.
        if (target !== undefined && !isRelsPart(target)) out.push(target);
      }
      return out;
    };
    /** Every part that can be reached from `from` without passing through a part of `blocked`. */
    const reach = async (from: string[], blocked: Set<string>): Promise<Set<string>> => {
      const seen = new Set<string>(from);
      const queue = from.slice();
      for (let i = 0; i < queue.length; i++) {
        if (i >= MAX_RELATED_PARTS) throw new Error('Cannot write the workbook: its parts refer to each other without end');
        for (const target of await targetsOf(queue[i])) {
          if (seen.has(target) || blocked.has(target)) continue;
          seen.add(target);
          queue.push(target);
        }
      }
      return seen;
    };
    const roots: string[] = [];
    for (const sheet of gone) {
      const name = actual.get(sheet.part.toLowerCase());
      if (name !== undefined) roots.push(name);
    }
    const doomed = await reach(roots, new Set());
    // What the rest of the package shows as well (a picture two drawings share) stays.
    const kept = await reach(['/'], new Set(roots));
    // Whatever the relationships of the deleted sheet claim, the package keeps what it stands on:
    // the workbook part and the parts of the sheets that stay.
    kept.add(actual.get(main.toLowerCase()) || main);
    for (const sheet of file) {
      if (sheet.part === '' || !present.has(sheet.part)) continue;
      kept.add(actual.get(sheet.part.toLowerCase()) || sheet.part);
    }
    for (const part of doomed) if (!kept.has(part) && !isRelsPart(part) && part.toLowerCase() !== '/[content_types].xml') wanted.add(part);
    for (const sheet of gone) {
      if (sheet.rel) rels.root.children = rels.root.children.filter((node) => node !== sheet.rel);
      relsChanged = true;
      // What an earlier save wrote differently around the sheet and stays in the package goes back to what the file had.
      for (const part of moved.get(sheet.part) || []) {
        const original = originals.get(part);
        if (original && !wanted.has(actual.get(part.toLowerCase()) || part)) writes.push([part, original]);
      }
      movedNow.set(sheet.part, []);
    }
  }

  // A calculation chain that no longer matches the formulas makes the application report the file
  // as damaged; without one it simply builds a new chain.
  const recalculate = modifiedSheets.length > 0 || reordered;
  let calcChain: string | undefined;
  const chainRel = recalculate ? withType('/calcChain') : undefined;
  if (chainRel) {
    const chain = resolveTarget(main, chainRel.attrs.get('Target') || '');
    rels.root.children = rels.root.children.filter((node) => node !== chainRel);
    relsChanged = true;
    // A modified sheet stays modified, and the chain stays out with it. Without one, undoing may
    // bring back the sheets of the file, and the chain with them.
    if (modifiedSheets.length > 0) calcChain = chain;
    else if (pkg.hasPart(chain) || removed.has(chain)) wanted.add(chain);
  }

  // Parts that were taken out and belong to the workbook again go back in first, as the file had them.
  const restores: PartWrite[] = [];
  const restored: string[] = [];
  for (const [part, contentType] of removed) {
    if (wanted.has(part) || part === calcChain) continue;
    const bytes = originals.get(part);
    if (bytes) restores.push([part, bytes, contentType]);
    const partRels = originals.get(relsPartOf(part));
    if (partRels) restores.push([relsPartOf(part), partRels]);
    restored.push(part);
  }
  const removals: [part: string, contentType: string | undefined][] = [];
  for (const part of wanted) {
    if (removed.has(part)) continue;
    await originalOf(part);
    await originalOf(relsPartOf(part));
    removals.push([part, pkg.contentTypeOf(part)]);
  }

  const bookParts: string[] = [];
  if (recalculate || renamed) {
    updateWorkbookXml(book, workbook, { file, renamed, reordered, relIds, relNamespace: relBase, active, recalculate });
    writes.push([main, book]);
    bookParts.push(main);
    if (reordered || renamed) {
      // The list of titles in the extended properties names the sheets as well.
      let appPart: string | undefined;
      for (const rel of await pkg.relationships('/')) if (endsWith(rel.type, '/extended-properties') && rel.part) appPart = rel.part;
      const appBytes = appPart === undefined ? undefined : await peek(appPart);
      const isWorksheet = (sheet: FileSheet) => !!sheet.rel && endsWith(typeOf(sheet.rel), '/worksheet');
      const app = appBytes
        ? updateAppXml(
            appBytes,
            file.filter(isWorksheet).map((sheet) => sheet.name),
            workbook.sheets.filter((sheet) => sheet.kind === 'worksheet').map((sheet) => sheet.name),
          )
        : undefined;
      if (appPart !== undefined && appBytes && app) {
        keep(appPart, appBytes);
        writes.push([appPart, app]);
        bookParts.push(appPart);
      }
    }
  }
  if (relsChanged) {
    writes.push([relsPart, rels]);
    bookParts.push(relsPart);
  }
  for (const part of moved.get(BOOK) || []) {
    const original = originals.get(part);
    if (bookParts.indexOf(part) < 0 && original) writes.push([part, original]);
  }
  movedNow.set(BOOK, bookParts);

  for (const [part, data, contentType] of restores) pkg.writePart(part, data, contentType);
  for (const [part, data, contentType] of writes) pkg.writePart(part, data, contentType);
  for (const [part] of removals) pkg.deletePart(part);
  for (const part of deletions) pkg.deletePart(part);
  if (calcChain !== undefined) pkg.deletePart(calcChain);

  for (const [part, list] of movedNow) moved.set(part, list);
  for (const part of restored) removed.delete(part);
  for (const [part, contentType] of removals) removed.set(part, contentType);
  if (calcChain !== undefined && removed.has(calcChain)) {
    // Gone for good from here on: its bytes need not be kept.
    removed.delete(calcChain);
    originals.delete(calcChain);
  }
  added.clear();
  for (const [sheet, part] of addedNow) added.set(sheet, part);
}
