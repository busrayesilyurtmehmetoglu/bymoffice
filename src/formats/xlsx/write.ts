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
  type DrawingAnchor,
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
import { adjustDrawingAnchor, commentCellNow, commentEdits } from '../../render/sheet-drawings';
import { MAX_SHEET_COMMENTS, readText, vmlRoot } from './comments';
import { decodeXstring } from './shared-strings';
import { parseColor } from './styles';
import { freshSheetParts } from './build';

const encoder = new TextEncoder();

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_REL_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const WORKSHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const VML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
const NS_VML = 'urn:schemas-microsoft-com:vml';
const NS_VML_OFFICE = 'urn:schemas-microsoft-com:office:office';
const NS_VML_EXCEL = 'urn:schemas-microsoft-com:office:excel';

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
 * belongs to, `foreign` those of the other sheets, all under the names the file has for the
 * sheets (`sheetName` as well). `rename` is what those names stand for now, with the names it
 * answers for; it is applied last.
 */
export function adjustSheetElement(raw: string, sheetName: string, own: StructuralEdit[], foreign: StructuralEdit[], rename?: { map: SheetNameMap; names: string[] }): string {
  if (own.length === 0 && foreign.length === 0 && !rename) return raw;
  const name = /^<([^\s/>]+)/.exec(raw);
  return adjustRaw(raw, name ? name[1] : '', {
    sheetName,
    own,
    foreign,
    needles: nameNeedles(foreign.map((edit) => edit.sheetName).concat(rename ? rename.names : [])),
    rename: rename ? rename.map : undefined,
    changed: false,
  });
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
  /** The windows of the workbook (by workbookViewId) whose views `tabSelected` is for; all of them without it. */
  tabWindows?: Set<number>;
  /**
   * The source is the skeleton of a sheet that is not in the file: everything the model
   * holds is written from it, whatever was marked as changed.
   */
  fresh?: boolean;
  /**
   * What <legacyDrawing> has to say once the notes of the sheet were changed: the id of the
   * relationship to the drawing with their boxes, or null for a sheet that has none any more.
   * Undefined leaves the element as it is.
   */
  legacyDrawing?: string | null;
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
  tabWindows: Set<number> | undefined;
  /** Every row is written from the model. */
  full: boolean;
}

const isTrueAttr = (value: string | undefined) => value === '1' || value === 'true';

/** <sheetViews> with the sheet marked as the selected one, or as not selected; the text itself when it says so already. */
function selectTab(raw: string, selected: boolean, windows?: Set<number>): string {
  const root = parseXml(raw).root;
  let changed = false;
  for (const view of childrenOf(root, 'sheetView')) {
    // Only the views of these windows of the workbook, when some are named.
    if (windows && !windows.has(Math.max(parseUint(view.attrs.get('workbookViewId')), 0))) continue;
    if (isTrueAttr(view.attrs.get('tabSelected')) === selected) continue;
    if (selected) view.attrs.set('tabSelected', '1');
    else view.attrs.delete('tabSelected');
    changed = true;
  }
  return changed ? elementText(root) : raw;
}

/**
 * The part of a chart sheet with its views marked as those of the selected sheet, or as not
 * selected; undefined when the part says so already, or cannot be read. A chart sheet is not
 * in the model: its part is the file's, and only the <sheetViews> element is written anew
 * here (added behind <sheetPr> when the sheet has none and becomes the selected one). Every
 * other byte stays what it was. A part in UTF-16 is left alone: writing it would change all of it.
 */
function selectChartTab(original: Uint8Array, selected: boolean): Uint8Array | undefined {
  if ((original[0] === 0xff && original[1] === 0xfe) || (original[0] === 0xfe && original[1] === 0xff)) return undefined;
  try {
    const whole = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(original);
    const mark = whole.charCodeAt(0) === 0xfeff ? whole.charAt(0) : '';
    const s = mark === '' ? whole : whole.slice(1);
    const sc = new XmlScanner(s);
    if (openRoot(sc) !== 'chartsheet') return undefined;
    const prefix = sc.name.slice(0, sc.name.length - sc.local.length);
    const views = `<${prefix}sheetViews><${prefix}sheetView tabSelected="1" workbookViewId="0"/></${prefix}sheetViews>`;
    const rootStart = nextTag(s, 0);
    let cursor = startTagEnd(s, rootStart);
    // A root without content has no views to change, and none are made up for it.
    if (sc.selfClosing) return undefined;
    let insertAt = cursor;
    while (sc.next()) {
      if (sc.kind === 'close') break;
      if (sc.kind !== 'open') continue;
      const start = nextTag(s, cursor);
      const local = sc.local;
      const raw = sc.readRaw();
      const end = start + raw.length;
      if (local === 'sheetViews') {
        let next = selectTab(raw, selected);
        // A list without a view cannot say that the sheet is selected.
        if (selected && next === raw && !/tabSelected/.test(raw)) next = views;
        return next === raw ? undefined : encoder.encode(mark + s.slice(0, start) + next + s.slice(end));
      }
      if (local === 'sheetPr') insertAt = end;
      cursor = end;
    }
    return selected ? encoder.encode(mark + s.slice(0, insertAt) + views + s.slice(insertAt)) : undefined;
  } catch (error) {
    // A chart sheet that cannot be read stays as it is.
    return undefined;
  }
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
    raw = selectTab(source, state.tabSelected, state.tabWindows);
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
    if (local === 'legacyDrawing' && replacement !== raw) state.changed = true;
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
  // The boxes of notes: the element is put where the schema has it when the sheet got its first note.
  if (options.legacyDrawing !== undefined) {
    const id = options.legacyDrawing;
    const attr = relPrefix !== undefined ? `${relPrefix}:id` : `xmlns:r="${NS_REL}" r:id`;
    regenerated.set('legacyDrawing', id === null ? '' : `<${prefix}legacyDrawing ${attr}="${escapeAttr(id)}"/>`);
  }
  // A sheet without views is not the selected one; one that has to be gets a view that says so.
  if (options.tabSelected) regenerated.set('sheetViews', `<${prefix}sheetViews><${prefix}sheetView tabSelected="1" workbookViewId="0"/></${prefix}sheetViews>`);
  const state: ChildState = { sheetData: 0, linkPrefix: relPrefix, colsRaw: undefined, linksRaw: undefined, changed: false, filterMode, tabSelected: options.tabSelected, tabWindows: options.tabWindows, full: fresh };
  /** Writes the regenerated elements that belong before the element of rank `rank` and were not in the source. */
  const flush = (rank: number) => {
    if (regenerated.size === 0) return;
    for (const local of ['sheetPr', 'sheetViews', 'cols', 'autoFilter', 'mergeCells', 'hyperlinks', 'legacyDrawing']) {
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
    const rank = worksheetRank(s, start, sc.local);
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

/**
 * A link to a place in the workbook may be written as the target of a relationship: "#Sheet2!A1"
 * with no target mode (the link of a shape or a picture is always written so; a cell link may be,
 * in place of a location). The place is a formula and follows what a text link follows: the rows
 * and columns of its own sheet and of the sheets it names, and then the names of the sheets. The
 * relationship itself always stays: the part it belongs to names it by its id.
 * UNSURE: a place on a sheet that is gone is written as "#REF!", a target that leads nowhere;
 * what the application writes there could not be checked.
 */
function adjustPlaceTargets(rels: XmlElement, edits: SheetStructuralEdit[], sheetName: string, others: OtherSheets | undefined): boolean {
  let changed = false;
  for (const rel of childrenNamed(rels, 'Relationship')) {
    const target = rel.attrs.get('Target') || '';
    if (target.length < 2 || target.charAt(0) !== '#' || rel.attrs.get('TargetMode') === 'External' || !endsWith(rel.attrs.get('Type') || '', '/hyperlink')) continue;
    let next = target.slice(1);
    for (const edit of edits) next = adjustFormula(next, sheetName, { ...edit, sheetName });
    if (others) {
      for (const edit of others.foreign) next = adjustFormula(next, sheetName, edit);
      if (others.rename) next = renameSheetsInFormula(next, others.rename);
    }
    // An error value starts with the same sign; the target is not given a second one.
    next = next.charAt(0) === '#' ? next : `#${next}`;
    if (next !== target) {
      rel.attrs.set('Target', next);
      changed = true;
    }
  }
  return changed;
}

/** True for relationships that may hold a place in the workbook as a target. */
const mayHoldPlaces = (bytes: Uint8Array) => /Target\s*=\s*["']#/.test(decodeSource(bytes));

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

/**
 * What became of the comments of the file once a sort moved notes (SheetComment.from): rows
 * and columns that were inserted or deleted move every comment alike, a sort moves each one
 * on its own, so each comment of the file is looked up by the cell the file has it on.
 */
interface NoteMoves {
  /** Cell of the file -> where the comment is now, with the numbers of its box (null: its cell was deleted). Only comments a sort moved. */
  moved: Map<number, { row: number; col: number; anchor: number[] | undefined } | null>;
  /** Cells of the file whose comments were not moved by a sort and are still in the model: the edits place them. */
  known: Set<number>;
}

/**
 * A comments part with each comment at the cell it is on now, for a sheet in which a sort moved
 * notes. A comment a sort moved gets the cell the model has for it; one that is no longer in
 * the model is removed, and its cell (of the file) is added to `gone`; every other one is moved
 * by the edits, as adjustComments() does. `seen` holds the cells of the file read so far: a
 * second comment on a cell is not one of the model and only follows the edits.
 */
function placeComments(bytes: Uint8Array, edits: SheetStructuralEdit[], moves: NoteMoves, gone: Set<number>, seen: Set<number>): Uint8Array {
  const doc = parseXml(bytes);
  let changed = false;
  for (const list of childrenNamed(doc.root, 'commentList')) {
    list.children = list.children.filter((node) => {
      if (node.type !== 'element' || localName(node.name) !== 'comment') return true;
      const ref = node.attrs.get('ref');
      const cell = ref === undefined ? undefined : parseCellRef(ref);
      if (!cell) return true;
      const key = cellKey(cell.row, cell.col);
      const first = !seen.has(key);
      seen.add(key);
      let target: { row: number; col: number } | undefined | null;
      if (first && moves.moved.has(key)) target = moves.moved.get(key);
      else if (first && !moves.known.has(key)) target = null;
      else target = moveCell(cell.row, cell.col, edits);
      if (!target) {
        if (first) gone.add(key);
        changed = true;
        return false;
      }
      const next = cellName(target.row, target.col);
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

/** What becomes of a comment box of a legacy drawing; undefined takes it out. */
type PlaceNote = (note: Note) => Note | undefined;

/**
 * The rule for the boxes of a sheet in which a sort moved notes: the box of a moved note goes
 * to the cell of its note with the place the model has for it, the box of a note that went
 * (`gone`, by the cells of the file) goes too, and every other box follows the edits. One box
 * per cell is the box of the note; a second one follows the edits.
 */
function placeNoteBy(edits: SheetStructuralEdit[], moves: NoteMoves, gone: Set<number>): PlaceNote {
  const seen = new Set<number>();
  return (note) => {
    const key = cellKey(note.row, note.col);
    const first = !seen.has(key);
    seen.add(key);
    if (first && moves.moved.has(key)) {
      const target = moves.moved.get(key);
      if (!target) return undefined;
      let anchor = note.anchor;
      if (anchor !== undefined && target.anchor) {
        const lead = /^\s*/.exec(anchor) as RegExpExecArray;
        const trail = /\s*$/.exec(anchor) as RegExpExecArray;
        const next = lead[0] + target.anchor.join(', ') + trail[0];
        if (next.replace(/\s/g, '') !== anchor.replace(/\s/g, '')) anchor = next;
      }
      return { row: target.row, col: target.col, anchor };
    }
    if (first && gone.has(key)) return undefined;
    return adjustNote(note, edits);
  };
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
function adjustVmlTree(doc: XmlDocument, place: PlaceNote): XmlDocument | undefined {
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
    const moved = place({ row, col, anchor: textOf(anchorEl) });
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
function adjustVmlText(text: string, place: PlaceNote): string {
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
    const moved = place({ row: Number(rowMatch[1]), col: Number(colMatch[1]), anchor: anchorMatch ? anchorMatch[1] : undefined });
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
  return placeVml(bytes, (note) => adjustNote(note, edits));
}

/** A legacy drawing with each comment box where `place` puts it; the same bytes when nothing changed. */
function placeVml(bytes: Uint8Array, place: PlaceNote): Uint8Array {
  let doc: XmlDocument | undefined;
  try {
    doc = parseXml(bytes);
  } catch (error) {
    doc = undefined;
  }
  if (doc) {
    const next = adjustVmlTree(doc, place);
    return next ? encoder.encode(serializeXml(next)) : bytes;
  }
  const text = decodeSource(bytes);
  const next = adjustVmlText(text, place);
  return next === text ? bytes : encoder.encode(next);
}

/** Whether a legacy drawing draws nothing: it parses and holds nothing but what a drawing declares. */
function vmlDrawsNothing(bytes: Uint8Array): boolean {
  try {
    return parseXml(bytes).root.children.every((node) => node.type !== 'element' || localName(node.name) === 'shapelayout' || localName(node.name) === 'shapetype');
  } catch (error) {
    return false;
  }
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

// --- notes: comments added, edited and deleted ---------------------------------------------------
//
// The comment parts of a sheet are brought in line with the model by cutting and pasting in
// their text: a comment or a box that did not change keeps every byte it has. First the rows
// and columns moved since the file was read are applied, as for a sheet whose notes were not
// touched (adjustComments(), adjustVml()); then, by the cells the comments are on now, what
// the model no longer has is cut out, a text that is another one is replaced, and the notes
// added since are appended.

/** Shape ids one <o:idmap> block of a legacy drawing stands for. */
export const VML_BLOCK = 1024;

/** The namespace prefixes of a legacy drawing, without their colons. */
export interface VmlPrefixes {
  v: string;
  o: string;
  x: string;
}

const VML_PREFIXES: VmlPrefixes = { v: 'v', o: 'o', x: 'x' };

/** The shape type the boxes of notes are of (a text box), as applications declare it once per drawing. */
export function noteShapeTypeXml(p: VmlPrefixes = VML_PREFIXES): string {
  return (
    `<${p.v}:shapetype id="_x0000_t202" coordsize="21600,21600" ${p.o}:spt="202" path="m,l,21600r21600,l21600,xe">` +
    `<${p.v}:stroke joinstyle="miter"/><${p.v}:path gradientshapeok="t" ${p.o}:connecttype="rect"/></${p.v}:shapetype>`
  );
}

/** <x:Anchor> of the box of a note the model gives no place for: beside its cell, where applications put a new one. */
export function defaultNoteAnchor(row: number, col: number): number[] {
  return [Math.min(col + 1, MAX_COLS - 1), 15, Math.max(row - 1, 0), 10, Math.min(col + 3, MAX_COLS - 1), 15, Math.min(row + 3, MAX_ROWS - 1), 4];
}

/**
 * The box of a note in a legacy drawing. `anchor` is the eight numbers of <x:Anchor>;
 * `declare` is put into the start tag (namespace declarations a drawing lacks).
 */
export function noteShapeXml(id: number, row: number, col: number, anchor: number[], visible: boolean, p: VmlPrefixes = VML_PREFIXES, declare = ''): string {
  return (
    `<${p.v}:shape${declare} id="_x0000_s${id}" type="#_x0000_t202" style="position:absolute;visibility:${visible ? 'visible' : 'hidden'}" fillcolor="#ffffe1" ${p.o}:insetmode="auto">` +
    `<${p.v}:fill color2="#ffffe1"/><${p.v}:shadow on="t" color="black" obscured="t"/><${p.v}:path ${p.o}:connecttype="none"/>` +
    `<${p.v}:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></${p.v}:textbox>` +
    `<${p.x}:ClientData ObjectType="Note"><${p.x}:MoveWithCells/><${p.x}:SizeWithCells/><${p.x}:Anchor>${anchor.join(', ')}</${p.x}:Anchor><${p.x}:AutoFill>False</${p.x}:AutoFill>` +
    `<${p.x}:Row>${row}</${p.x}:Row><${p.x}:Column>${col}</${p.x}:Column>${visible ? `<${p.x}:Visible/>` : ''}</${p.x}:ClientData></${p.v}:shape>`
  );
}

interface AddedNote {
  row: number;
  col: number;
  text: string;
  author: string;
}

/** What the model says about the notes of a sheet, by the cells they are on now. */
interface NotePlan {
  /** Comments of the file that are still there: cell key -> their text, or null for one that is left as it is (a conversation). */
  kept: Map<number, string | null>;
  /** Notes added since the sheet was read, in the order of the model. */
  added: AddedNote[];
  /** Set when a sort moved a comment of the file: where the comments of the file are found by the cells the file has them on. */
  moves?: NoteMoves;
}

const EMU_PER_PIXEL = 9525;

/** The eight numbers of <x:Anchor> for the box of a note, when the model places it between two cells. */
function anchorNumbers(box: DrawingAnchor | undefined): number[] | undefined {
  if (!box || box.kind !== 'twoCell' || !box.from || !box.to) return undefined;
  const numbers = [box.from.col, Math.round(box.from.colOff / EMU_PER_PIXEL), box.from.row, Math.round(box.from.rowOff / EMU_PER_PIXEL), box.to.col, Math.round(box.to.colOff / EMU_PER_PIXEL), box.to.row, Math.round(box.to.rowOff / EMU_PER_PIXEL)];
  return numbers.every((n) => n >= 0 && Math.floor(n) === n) ? numbers : undefined;
}

const cellKey = (row: number, col: number): number => row * MAX_COLS + col;

function notePlanOf(sheet: Sheet): NotePlan {
  const plan: NotePlan = { kept: new Map(), added: [] };
  const seen = new Set<number>();
  const edits = sheet.structuralEdits;
  const known = new Set<number>();
  let moved: NoteMoves['moved'] | undefined;
  for (const comment of sheet.comments || []) {
    if (!comment || !(comment.row >= 0 && comment.row < MAX_ROWS && comment.col >= 0 && comment.col < MAX_COLS)) continue;
    const cell = commentCellNow(comment, edits);
    // A comment of the file that a sort moved is known by the cell the file has it on.
    const from = comment.since !== undefined ? comment.from : undefined;
    const ofFile = comment.since === undefined || from !== undefined;
    if (from) {
      const box = cell && comment.box ? anchorNumbers(adjustDrawingAnchor(comment.box, commentEdits(comment, edits))) : undefined;
      (moved || (moved = new Map())).set(cellKey(from.row, from.col), cell ? { row: cell.row, col: cell.col, anchor: box } : null);
    } else if (ofFile) known.add(cellKey(comment.row, comment.col));
    if (!cell) continue;
    const key = cellKey(cell.row, cell.col);
    // As where comments are shown: the first one of a cell is the one that counts.
    if (seen.has(key)) continue;
    seen.add(key);
    if (ofFile) plan.kept.set(key, comment.threaded ? null : String(comment.text));
    else plan.added.push({ row: cell.row, col: cell.col, text: String(comment.text), author: comment.author || '' });
  }
  if (moved) plan.moves = { moved, known };
  return plan;
}

interface RawChild {
  local: string;
  start: number;
  end: number;
  /** The attribute that was asked for. */
  value: string | undefined;
}

interface RawElement {
  name: string;
  /** Offset of its start tag, and just past it. */
  start: number;
  openEnd: number;
  /** Offset of its end tag; `openEnd` for an element without one. */
  closeStart: number;
  selfClosing: boolean;
  children: RawChild[];
}

/** Where the children of the first element of `s` are in it, without building a tree. */
function rawElement(s: string, attr?: string): RawElement {
  const sc = new XmlScanner(s);
  openRoot(sc);
  const name = sc.name;
  const start = nextTag(s, 0);
  let cursor = startTagEnd(s, start);
  const openEnd = cursor;
  const children: RawChild[] = [];
  if (sc.selfClosing) return { name, start, openEnd, closeStart: openEnd, selfClosing: true, children };
  while (sc.next()) {
    if (sc.kind === 'close') break;
    if (sc.kind !== 'open') continue;
    const at = nextTag(s, cursor);
    const local = sc.local;
    const value = attr === undefined ? undefined : sc.attr(attr);
    const raw = sc.readRaw();
    cursor = at + raw.length;
    children.push({ local, start: at, end: cursor, value });
  }
  return { name, start, openEnd, closeStart: nextTag(s, cursor), selfClosing: false, children };
}

interface Splice {
  start: number;
  end: number;
  text: string;
}

/** Text put inside an element behind what it holds; an element without an end tag gets one. `base` is where the element's text starts. */
function appendInside(el: RawElement, base: number, text: string): Splice {
  if (!el.selfClosing) return { start: base + el.closeStart, end: base + el.closeStart, text };
  // "<name .../>": the slash goes, and the end tag comes behind the new content.
  return { start: base + el.openEnd - 2, end: base + el.openEnd, text: `>${text}</${el.name}>` };
}

function applySplices(s: string, splices: Splice[]): string {
  if (splices.length === 0) return s;
  // Splices at one offset keep the order they were made in.
  const order = splices.map((splice, index) => ({ splice, index })).sort((a, b) => a.splice.start - b.splice.start || a.index - b.index);
  let out = '';
  let cursor = 0;
  for (const { splice } of order) {
    if (splice.start < cursor) throw new Error('Cannot write the comments: their parts cannot be read');
    out += s.slice(cursor, splice.start) + splice.text;
    cursor = splice.end;
  }
  return out + s.slice(cursor);
}

/**
 * Applies splices whose offsets are those of the part as the scanner sees it (decodeSource():
 * no byte order mark, line breaks as "\n") to the part as it is, so that everything between
 * the splices keeps its bytes, line breaks included. The result is UTF-8; a declaration that
 * names another encoding is corrected.
 */
function spliceExact(source: Uint8Array, splices: Splice[]): Uint8Array {
  const exact = decodeExact(source);
  const mark = exact.charCodeAt(0) === 0xfeff ? 1 : 0;
  // Offsets (in the scanner's text) of the line breaks that are two characters in the part.
  const pairs: number[] = [];
  for (let i = exact.indexOf('\r\n', mark); i >= 0; i = exact.indexOf('\r\n', i + 2)) pairs.push(i - mark - pairs.length);
  const toExact = (offset: number): number => {
    let low = 0;
    let high = pairs.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (pairs[middle] < offset) low = middle + 1;
      else high = middle;
    }
    return mark + offset + low;
  };
  const out = applySplices(
    exact,
    splices.map((splice) => ({ start: toExact(splice.start), end: toExact(splice.end), text: splice.text })),
  );
  return encoder.encode(mark ? out : fixDeclaration(out));
}

/** The place of a child of a worksheet in the order of the schema; -1 for one the schema does not name. */
function worksheetRank(s: string, start: number, local: string): number {
  if (local !== 'AlternateContent') return WORKSHEET_ORDER.indexOf(local);
  // A wrapper stands where what it wraps would: the first element of its first choice.
  const match = /^<[^>]*>\s*<[^>]*>\s*<(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/.exec(s.slice(start, start + 4096));
  return match ? WORKSHEET_ORDER.indexOf(match[1]) : -1;
}

/**
 * A sheet part with nothing changed but what <legacyDrawing> says (see SheetRewrite.legacyDrawing):
 * every other byte stays. Undefined when the part already says so.
 */
function setLegacyDrawing(original: Uint8Array, id: string | null, sheetName: string): Uint8Array | undefined {
  const s = decodeSource(original);
  const root = rawElement(s);
  const rootLocal = localName(root.name);
  if (rootLocal !== 'worksheet' && rootLocal !== 'dialogsheet' && rootLocal !== 'macrosheet') {
    throw new Error(`Cannot write sheet "${sheetName}": its part has the root element <${root.name}>, not a worksheet`);
  }
  const prefix = root.name.slice(0, root.name.length - rootLocal.length);
  const declared = new RegExp(`\\sxmlns:([A-Za-z_][\\w.-]*)\\s*=\\s*["'](?:${escapeRegExp(NS_REL)}|${escapeRegExp(NS_REL_STRICT)})["']`).exec(s.slice(root.start, root.openEnd));
  const attr = declared ? `${declared[1]}:id` : `xmlns:r="${NS_REL}" r:id`;
  const text = id === null ? '' : `<${prefix}legacyDrawing ${attr}="${escapeAttr(id)}"/>`;
  const rank = WORKSHEET_ORDER.indexOf('legacyDrawing');
  const present = root.children.find((child) => child.local === 'legacyDrawing');
  if (present) {
    if (s.slice(present.start, present.end) === text) return undefined;
    return spliceExact(original, [{ start: present.start, end: present.end, text }]);
  }
  if (id === null) return undefined;
  if (root.selfClosing) throw new Error(`Cannot write sheet "${sheetName}": its part has no <sheetData>`);
  const behind = root.children.find((child) => worksheetRank(s, child.start, child.local) > rank);
  const at = behind ? behind.start : root.closeStart;
  return spliceExact(original, [{ start: at, end: at, text }]);
}

const sameColor = (a: Color | undefined, b: Color | undefined): boolean =>
  (!a && !b) || (!!a && !!b && (a.rgb === undefined ? undefined : a.rgb.toUpperCase()) === (b.rgb === undefined ? undefined : b.rgb.toUpperCase()) && a.theme === b.theme && a.indexed === b.indexed && a.tint === b.tint && a.auto === b.auto);

/**
 * A sheet part with nothing changed but the colour of its tab (<tabColor> in <sheetPr>):
 * every other byte stays, the other attributes and children of <sheetPr> included. Undefined
 * when the part says what the model says already, so that a colour of the file (by theme, by
 * index, with a tint) is only written anew once it is another one.
 */
function setTabColorIn(original: Uint8Array, color: Color | undefined, sheetName: string): Uint8Array | undefined {
  const s = decodeSource(original);
  const root = rawElement(s);
  const rootLocal = localName(root.name);
  if (rootLocal !== 'worksheet' && rootLocal !== 'dialogsheet' && rootLocal !== 'macrosheet') {
    throw new Error(`Cannot write sheet "${sheetName}": its part has the root element <${root.name}>, not a worksheet`);
  }
  const prefix = root.name.slice(0, root.name.length - rootLocal.length);
  const text = color ? `<${prefix}tabColor${colorAttrs(color)}/>` : '';
  const properties = root.children.find((child) => child.local === 'sheetPr');
  if (!properties) {
    if (!color) return undefined;
    // A root without content gets its first child, and an end tag with it.
    if (root.selfClosing) return spliceExact(original, [appendInside(root, 0, `<${prefix}sheetPr>${text}</${prefix}sheetPr>`)]);
    // <sheetPr> is the first child of a worksheet.
    return spliceExact(original, [{ start: root.openEnd, end: root.openEnd, text: `<${prefix}sheetPr>${text}</${prefix}sheetPr>` }]);
  }
  const raw = s.slice(properties.start, properties.end);
  const element = rawElement(raw);
  const present = element.children.find((child) => child.local === 'tabColor');
  if (present) {
    if (sameColor(parseColor(parseXml(raw.slice(present.start, present.end)).root), color)) return undefined;
    return spliceExact(original, [{ start: properties.start + present.start, end: properties.start + present.end, text }]);
  }
  if (!color) return undefined;
  // <tabColor> is the first child of <sheetPr>.
  if (element.selfClosing) return spliceExact(original, [appendInside(element, properties.start, text)]);
  return spliceExact(original, [{ start: properties.start + element.openEnd, end: properties.start + element.openEnd, text }]);
}

/** The comments read so far of a sheet, over all of its comment parts: what the reader took as the comments of the sheet. */
interface CommentsScan {
  seen: Set<number>;
  known: number;
}

/** Indexes of authors beyond this are not made room for (the part is damaged). */
const MAX_AUTHOR_INDEX = 65535;

interface CommentsEdit {
  bytes: Uint8Array;
  changed: boolean;
  /** Comments the part holds afterwards. */
  left: number;
}

/**
 * A comments part brought in line with the notes of the model. `source` has its comments at
 * the cells they are on now. The cells of the comments that were cut out are added to `removed`.
 */
function editComments(source: Uint8Array, plan: NotePlan, added: AddedNote[], scan: CommentsScan, removed: Set<number>): CommentsEdit {
  const s = decodeSource(source);
  const root = rawElement(s);
  const prefix = root.name.slice(0, root.name.length - localName(root.name).length);
  const splices: Splice[] = [];
  let left = 0;
  /** The highest index of an author a comment of the part names, listed or not. */
  let highestAuthor = -1;
  let authors: RawChild | undefined;
  let firstList: RawChild | undefined;
  let extensions: RawChild | undefined;
  for (const part of root.children) {
    if (part.local === 'authors' && !authors) authors = part;
    if (part.local === 'extLst' && !extensions) extensions = part;
    if (part.local !== 'commentList') continue;
    if (!firstList) firstList = part;
    const listText = s.slice(part.start, part.end);
    const list = rawElement(listText, 'ref');
    for (const item of list.children) {
      if (item.local !== 'comment') continue;
      const named = /\sauthorId\s*=\s*["'](\d{1,5})["']/.exec(listText.slice(item.start, startTagEnd(listText, item.start)));
      if (named && Number(named[1]) > highestAuthor && Number(named[1]) <= MAX_AUTHOR_INDEX) highestAuthor = Number(named[1]);
      const cell = item.value === undefined ? undefined : parseCellRef(item.value);
      const key = cell ? cellKey(cell.row, cell.col) : -1;
      // What the reader did not take as a comment of the sheet (no cell, a second one on a cell, one too many) stays as it is.
      if (!cell || scan.seen.has(key) || scan.known >= MAX_SHEET_COMMENTS) {
        left++;
        continue;
      }
      scan.seen.add(key);
      scan.known++;
      const base = part.start + item.start;
      if (!plan.kept.has(key)) {
        splices.push({ start: base, end: part.start + item.end, text: '' });
        removed.add(key);
        continue;
      }
      left++;
      const wanted = plan.kept.get(key);
      if (wanted === null || wanted === undefined) continue;
      const itemText = listText.slice(item.start, item.end);
      const element = parseXml(itemText).root;
      if (readText(childNamed(element, 'text')).text === wanted) continue;
      const body = `<${prefix}text>${textElement(wanted, prefix)}</${prefix}text>`;
      const inner = rawElement(itemText);
      const text = inner.children.find((child) => child.local === 'text');
      if (text) splices.push({ start: base + text.start, end: base + text.end, text: body });
      else if (inner.selfClosing) splices.push(appendInside(inner, base, body));
      // <text> is the first child of a comment.
      else splices.push({ start: base + inner.openEnd, end: base + inner.openEnd, text: body });
    }
  }

  if (added.length > 0) {
    const names: string[] = [];
    let authorsEl: RawElement | undefined;
    if (authors) {
      authorsEl = rawElement(s.slice(authors.start, authors.end));
      const parsed = parseXml(s.slice(authors.start, authors.end)).root;
      for (const author of childrenNamed(parsed, 'author')) names.push(decodeXstring(textOf(author) || ''));
    }
    const known = names.length;
    // Comments may name authors the list does not have (they are then by nobody). A new author is listed behind
    // every index that is in use, so that none of those comments becomes one of the new author.
    const padding = Math.max(highestAuthor + 1 - known, 0);
    const fresh: string[] = [];
    let items = '';
    for (const note of added) {
      let id = names.indexOf(note.author);
      if (id < 0) {
        id = fresh.indexOf(note.author);
        if (id < 0) {
          id = fresh.length;
          fresh.push(note.author);
        }
        id += known + padding;
      }
      items += `<${prefix}comment ref="${cellName(note.row, note.col)}" authorId="${id}" shapeId="0"><${prefix}text>${textElement(note.text, prefix)}</${prefix}text></${prefix}comment>`;
      left++;
    }
    let newAuthors = '';
    if (fresh.length > 0) for (let i = 0; i < padding; i++) newAuthors += `<${prefix}author></${prefix}author>`;
    for (const name of fresh) newAuthors += `<${prefix}author>${escapeXstring(name)}</${prefix}author>`;
    if (authors && authorsEl) {
      if (newAuthors !== '') splices.push(appendInside(authorsEl, authors.start, newAuthors));
    } else {
      // The list of authors is the first child of the part.
      const at = root.selfClosing ? -1 : root.openEnd;
      newAuthors = `<${prefix}authors>${newAuthors}</${prefix}authors>`;
      if (at >= 0) splices.push({ start: at, end: at, text: newAuthors });
      else items = newAuthors + `<${prefix}commentList>${items}</${prefix}commentList>`;
    }
    if (firstList) splices.push(appendInside(rawElement(s.slice(firstList.start, firstList.end)), firstList.start, items));
    else if (root.selfClosing) splices.push(appendInside(root, 0, items));
    else {
      const at = extensions ? extensions.start : root.closeStart;
      splices.push({ start: at, end: at, text: `<${prefix}commentList>${items}</${prefix}commentList>` });
    }
  }
  if (splices.length === 0) return { bytes: source, changed: false, left };
  return { bytes: spliceExact(source, splices), changed: true, left };
}

/** A part as text with nothing brought in line: every line break as it is, a byte order mark kept. */
function decodeExact(bytes: Uint8Array): string {
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) return decodeSource(bytes);
  try {
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch (error) {
    return new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
  }
}

const escapeRegExp = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** The prefix a namespace is declared with in a text, without parsing it. */
function declaredPrefix(text: string, namespace: string): string | undefined {
  const match = new RegExp(`xmlns:([A-Za-z_][\\w.-]*)\\s*=\\s*["']${escapeRegExp(namespace)}["']`).exec(text);
  return match ? match[1] : undefined;
}

/** Blocks of shape ids a legacy drawing uses: those its <o:idmap> names and those its shapes are in. */
function vmlBlocks(text: string): number[] {
  const out: number[] = [];
  const map = /:idmap\b[^>]*?\bdata\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(text);
  if (map) {
    for (const piece of (map[1] !== undefined ? map[1] : map[2]).split(',')) {
      const block = parseUint(piece.trim());
      if (block > 0) out.push(block);
    }
  }
  const ids = /_x0000_s(\d{1,9})(?!\d)/g;
  let match: RegExpExecArray | null;
  let guard = 0;
  while ((match = ids.exec(text)) && guard++ < 10000000) out.push(Math.floor(Number(match[1]) / VML_BLOCK));
  return out;
}

interface VmlEdit {
  bytes: Uint8Array;
  changed: boolean;
  /** Nothing is drawn by the part any more: it holds no shape, only what a drawing declares. */
  empty: boolean;
}

/**
 * A legacy drawing brought in line with the notes of the model: the boxes of the notes on the
 * cells of `removed` are cut out, and a box is appended for every note of `added`. Like
 * adjustVmlText() this works on the text, so that a part that is not well-formed is no
 * obstacle and every other shape (a form control, a box that stays) keeps its bytes.
 * `nextBlock` gives a block of shape ids that no drawing of the workbook uses.
 */
function editVml(source: Uint8Array, removed: Set<number>, added: AddedNote[], nextBlock: () => number): VmlEdit {
  let text = decodeExact(source);
  const declared = { v: declaredPrefix(text, NS_VML), o: declaredPrefix(text, NS_VML_OFFICE), x: declaredPrefix(text, NS_VML_EXCEL) };
  const p: VmlPrefixes = { v: declared.v || 'v', o: declared.o || 'o', x: declared.x || 'x' };
  let changed = false;

  if (removed.size > 0) {
    const open = `<${p.v}:shape`;
    const close = `</${p.v}:shape>`;
    const isNote = new RegExp(`<${escapeRegExp(p.x)}:ClientData[^>]*ObjectType\\s*=\\s*["']Note["']`);
    const rowOf = new RegExp(`<${escapeRegExp(p.x)}:Row>\\s*(\\d{1,7})\\s*</${escapeRegExp(p.x)}:Row>`);
    const colOf = new RegExp(`<${escapeRegExp(p.x)}:Column>\\s*(\\d{1,5})\\s*</${escapeRegExp(p.x)}:Column>`);
    const gone = new Set<number>();
    let out = '';
    let copied = 0;
    let i = 0;
    for (;;) {
      const start = text.indexOf(open, i);
      if (start < 0) break;
      const after = text.charCodeAt(start + open.length);
      // "<v:shapetype" and the like are other elements.
      if (!(after === 32 || after === 62 || after === 9 || after === 10 || after === 13)) {
        i = start + open.length;
        continue;
      }
      const end = text.indexOf(close, start);
      if (end < 0) break;
      i = end + close.length;
      const block = text.slice(start, i);
      if (!isNote.test(block)) continue;
      const row = rowOf.exec(block);
      const col = colOf.exec(block);
      if (!row || !col) continue;
      const key = cellKey(Number(row[1]), Number(col[1]));
      // One box per note; a second box on the cell is not the one of the note that went.
      if (!removed.has(key) || gone.has(key)) continue;
      gone.add(key);
      out += text.slice(copied, start);
      copied = i;
    }
    if (copied > 0) {
      text = out + text.slice(copied);
      changed = true;
    }
  }

  if (added.length > 0) {
    // The reader marks a sheet with such a drawing, and the operations refuse its notes: this is the last defence.
    const root = vmlRoot(text);
    if (!root) throw new Error('Cannot write the notes: the legacy drawing of the sheet cannot be read');
    const rootEnd = root.openEnd;
    const closeAt = root.closeStart;
    // Shape ids come in blocks; each block a drawing uses is named in its <o:idmap>.
    // Block n stands for the ids n*1024 .. n*1024+1023, of which the first is not used.
    const map = new RegExp(`<${escapeRegExp(p.o)}:idmap\\b[^>]*?\\bdata\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(text);
    const listed: number[] = [];
    if (map) {
      for (const piece of (map[1] !== undefined ? map[1] : map[2]).split(',')) {
        const number = parseUint(piece.trim());
        if (number > 0) listed.push(number);
      }
    }
    let block = listed.length > 0 ? listed[listed.length - 1] : -1;
    let next = 0;
    if (block > 0) {
      next = block * VML_BLOCK + 1;
      const ids = /_x0000_s(\d{1,9})(?!\d)/g;
      let match: RegExpExecArray | null;
      let guard = 0;
      while ((match = ids.exec(text)) && guard++ < 10000000) {
        const id = Number(match[1]);
        if (Math.floor(id / VML_BLOCK) === block && id >= next) next = id + 1;
      }
    }
    const fresh: number[] = [];
    // Prefixes the drawing does not declare are declared on each box.
    let declare = '';
    if (!declared.v) declare += ` xmlns:${p.v}="${NS_VML}"`;
    if (!declared.o) declare += ` xmlns:${p.o}="${NS_VML_OFFICE}"`;
    if (!declared.x) declare += ` xmlns:${p.x}="${NS_VML_EXCEL}"`;
    let shapes = '';
    for (const note of added) {
      if (block <= 0 || next >= (block + 1) * VML_BLOCK) {
        block = nextBlock();
        fresh.push(block);
        next = block * VML_BLOCK + 1;
      }
      shapes += noteShapeXml(next++, note.row, note.col, defaultNoteAnchor(note.row, note.col), false, p, declare);
    }
    if (text.indexOf('id="_x0000_t202"') < 0 && text.indexOf("id='_x0000_t202'") < 0) shapes = noteShapeTypeXml(p).replace(`<${p.v}:shapetype`, `<${p.v}:shapetype${declare}`) + shapes;
    const splices: Splice[] = [{ start: closeAt, end: closeAt, text: shapes }];
    if (fresh.length > 0) {
      if (map) {
        const value = map[1] !== undefined ? map[1] : map[2];
        const at = map.index + map[0].length - 1 - value.length;
        splices.push({ start: at, end: at + value.length, text: listed.concat(fresh).join(',') });
      } else {
        const layout = `<${p.o}:shapelayout${declare} ${p.v}:ext="edit"><${p.o}:idmap ${p.v}:ext="edit" data="${fresh.join(',')}"/></${p.o}:shapelayout>`;
        splices.push({ start: rootEnd, end: rootEnd, text: layout });
      }
    }
    text = applySplices(text, splices);
    changed = true;
  }
  if (!changed) return { bytes: source, changed: false, empty: false };
  // The bytes are UTF-8 from here on, whatever the part was in: its declaration has to say so.
  if (text.charCodeAt(0) !== 0xfeff) text = fixDeclaration(text);

  // A drawing that draws nothing is not kept: only what parses and holds nothing but declarations counts as that.
  let empty = false;
  try {
    empty = parseXml(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text).root.children.every((node) => node.type !== 'element' || localName(node.name) === 'shapelayout' || localName(node.name) === 'shapetype');
  } catch (error) {
    empty = false;
  }
  return { bytes: encoder.encode(text), changed: true, empty };
}

/** The id of the relationship <legacyDrawing> of a sheet names; undefined for a sheet without the element. */
function legacyDrawingId(sheetXml: string): string | undefined {
  const match = /<(?:[A-Za-z_][\w.-]*:)?legacyDrawing(?=[\s/>])[^>]*?\b[A-Za-z_][\w.-]*:id\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(sheetXml);
  if (!match) return undefined;
  return plainOf(match[1] !== undefined ? match[1] : match[2]);
}

/** The target of a relationship from one part to another, relative the way applications write it. */
function relativeTarget(from: string, to: string): string {
  const base = from.split('/').slice(1, -1);
  const target = to.split('/').slice(1);
  let same = 0;
  while (same < base.length && same < target.length - 1 && base[same] === target[same]) same++;
  let out = '';
  for (let i = same; i < base.length; i++) out += '../';
  return out + target.slice(same).join('/');
}

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
  /** Sheets of the file are shown that the file hides, or the other way round. */
  restated: boolean;
  /** The sheets are not the ones of the file, one for one (a sheet was added or deleted): whatever counts sheets is brought up to date. */
  reordered: boolean;
  /** The relationship made for each sheet that is not in the file. */
  relIds: Map<Sheet, string>;
  /** The namespace relationship ids are in, in this flavour of the format (transitional or strict). */
  relNamespace: string;
  /** Index of the sheet the file opens on. */
  active: number;
  /** The sheets are those of the file, and the model opens on another one of them than the file does. */
  activeMoved: boolean;
  /** Stored results may be stale: the application is asked to calculate when it opens the file. */
  recalculate: boolean;
  /**
   * Windows of the workbook (by their place among the <workbookView> elements) that open on a
   * sheet which is hidden now: each opens on the active sheet instead. Only where the sheets are
   * those of the file in its order.
   */
  hiddenWindows?: Set<number>;
}

const MAX_SHEET_ID = 65534;

/** Whether the file shows a sheet, as the reader takes <sheet state>. */
function fileStateOf(el: XmlElement): string {
  const state = el.attrs.get('state');
  return state === 'hidden' || state === 'veryHidden' ? state : 'visible';
}

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
      if (fileStateOf(el) !== sheet.state) {
        if (sheet.state === 'visible') el.attrs.delete('state');
        else el.attrs.set('state', sheet.state);
      }
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
 * The views of the workbook after sheets were added or deleted, or another sheet is the one
 * shown: the sheet shown is the active one of the model, and the first tab shown stays the
 * sheet it was (or what is at its place).
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
  if (change.renamed || change.reordered || change.restated) {
    const sheetIds = updateSheetList(root, workbook, change);
    if (change.reordered) updateBookViews(root, workbook, change, sheetIds);
  }
  // The same sheets, another one shown: no custom view has lost its sheet, so none is looked at.
  if (change.activeMoved) updateBookViews(root, workbook, change, []);
  if (change.restated && !change.reordered) {
    // No window opens on a sheet that is hidden now, and none has one as the first tab it shows.
    const views = childNamed(root, 'bookViews');
    const sheets = workbook.sheets;
    let firstVisible = 0;
    while (firstVisible < sheets.length - 1 && sheets[firstVisible].state !== 'visible') firstVisible++;
    (views ? childrenNamed(views, 'workbookView') : []).forEach((view, index) => {
      if (change.hiddenWindows && change.hiddenWindows.has(index)) {
        if (change.active === 0) view.attrs.delete('activeTab');
        else view.attrs.set('activeTab', String(change.active));
      }
      const first = view.attrs.get('firstSheet');
      const at = first === undefined ? -1 : parseUint(first);
      if (at >= 0 && at < sheets.length && sheets[at].state !== 'visible' && String(firstVisible) !== first) view.attrs.set('firstSheet', String(firstVisible));
    });
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
      // The names the model adds are the range of a filter that was switched on and the names of a sheet
      // that is not in the file (a copy brings the names of its sheet); any other name that is not in the
      // file is a mistake of the caller.
      const ofNewSheet = name.localSheet !== undefined && !!workbook.sheets[name.localSheet] && workbook.sheets[name.localSheet].part === '';
      if (!isFilterName(name.name) && !ofNewSheet) throw new Error(`Cannot write the workbook: the defined name "${name.name}" is not in the file`);
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
  /**
   * Per sheet: the parts the last save made for its notes that the file does not have (a
   * comments part, a legacy drawing, the relationships of the sheet). A save that no longer
   * needs one takes it out again; one that does writes it under the same name.
   */
  notes: Map<Sheet, NoteParts>;
}

/** Names of the parts made for the notes of a sheet; each is of a part the file does not have. */
interface NoteParts {
  comments?: string;
  vml?: string;
  rels?: string;
}

/** What a save does around a sheet for its notes. */
interface NoteWrites {
  /** Parts of the file with what they become; null takes the part out. */
  parts: Map<string, Uint8Array | null>;
  /** Parts the file does not have. */
  fresh: PartWrite[];
  /** Relationships of the sheet to add, and the ids of those to drop. */
  relate: { id: string; kind: string; part: string }[];
  unrelate: Set<string>;
  /** See SheetRewrite.legacyDrawing. */
  legacy: string | null | undefined;
  names: NoteParts;
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

/** The active sheet of the model as the reader gives it: an index that is none stands for the first sheet. */
function modelActiveOf(workbook: Workbook): number {
  const active = workbook.activeSheet;
  return active >= 0 && active < workbook.sheets.length && Math.floor(active) === active ? active : 0;
}

/** The sheet the file says it opens on, as the reader takes it (see readWorkbook): activeTab of the first view. */
function fileActiveOf(root: XmlElement, count: number): number {
  const views = childNamed(root, 'bookViews');
  const view = views ? childrenNamed(views, 'workbookView')[0] : undefined;
  const stated = (view && view.attrs.get('activeTab')) || '';
  const active = /^\d{1,6}$/.test(stated) ? Number(stated) : 0;
  return active < count ? active : 0;
}

/**
 * Brings the package up to date with the model. Does nothing when nothing was modified.
 *
 * The sheet the file opens on (Workbook.activeSheet) is written when anything else is: see
 * `activeMoved` below.
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
  const notesModified = workbook.sheets.some((sheet) => sheet.commentsModified && sheet.kind !== 'chartsheet');
  if (modifiedSheets.length === 0 && !stylesModified && !sheetsModified && !notesModified) return;

  let saved = SAVES.get(workbook);
  if (!saved) {
    saved = { originals: new Map(), moved: new Map(), removed: new Map(), added: new Map(), notes: new Map() };
    SAVES.set(workbook, saved);
  }
  const { originals, moved, removed, added, notes: noted } = saved;
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
  let restated = false;
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
      if (fileStateOf(inFile.element) !== sheet.state) restated = true;
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
  // Names of parts that are taken: part names are told apart without regard to case, and a part
  // that was taken out may come back.
  let taken: Set<string> | undefined;
  const takenNames = (): Set<string> => {
    if (!taken) {
      taken = new Set();
      for (const name of pkg.partNames()) taken.add(name.toLowerCase());
      for (const [name, bytes] of originals) if (bytes !== null) taken.add(name.toLowerCase());
      for (const name of removed.keys()) taken.add(name.toLowerCase());
      for (const name of added.values()) taken.add(name.toLowerCase());
      for (const names of noted.values()) for (const name of [names.comments, names.vml, names.rels]) if (name !== undefined) taken.add(name.toLowerCase());
    }
    return taken;
  };
  /** The first free name of a numbered part: "/xl/comments" + n + ".xml". */
  const freeName = (head: string, tail: string): string => {
    const names = takenNames();
    for (let n = 1; ; n++) {
      const candidate = `${head}${n}${tail}`;
      if (names.has(candidate.toLowerCase())) continue;
      names.add(candidate.toLowerCase());
      return candidate;
    }
  };
  // Parts taken out of the package by this save; see `removed`.
  const wanted = new Set<string>();
  // A relationship may spell the name of a part in another case than the archive does. What is
  // remembered about a part (its bytes, that it was taken out) is kept under the name of the
  // archive, so that a part that comes back comes back as the entry it was.
  let spelled: Map<string, string> | undefined;
  const actualOf = (part: string): string => {
    if (!spelled) {
      spelled = new Map();
      for (const name of removed.keys()) spelled.set(name.toLowerCase(), name);
      for (const name of pkg.partNames()) spelled.set(name.toLowerCase(), name);
    }
    return spelled.get(part.toLowerCase()) || part;
  };

  // The highest block of shape ids a legacy drawing of the file uses; new boxes get their ids from the blocks behind it.
  let topBlock: number | undefined;
  const useBlocks = async (): Promise<() => number> => {
    if (topBlock === undefined) {
      let top = 0;
      const names = pkg.partNames().concat([...removed.keys()]);
      for (const name of names) {
        if (!/\.vml$/i.test(name)) continue;
        const bytes = await peek(name);
        if (!bytes) continue;
        for (const block of vmlBlocks(decodeExact(bytes))) if (block > top) top = block;
      }
      topBlock = top;
    }
    return () => {
      topBlock = (topBlock as number) + 1;
      return topBlock;
    };
  };

  const mainPrefix = prefixOf(book.root);
  const mainNamespace = book.root.attrs.get(mainPrefix === '' ? 'xmlns' : `xmlns:${mainPrefix.slice(0, -1)}`) || NS_MAIN;
  const notesNow = new Map<Sheet, NoteParts>();
  /**
   * What the notes of a sheet whose notes were changed ask of the parts around it; undefined
   * when the parts of the file say what the model says (nothing was changed after all, or
   * everything was undone). `sheetXml` and `sheetRels` are the sheet as the file has it
   * (undefined for a sheet that is not in the file).
   */
  const notesOf = async (sheet: Sheet, sheetPart: string, sheetXml: string | undefined, sheetRels: Uint8Array | undefined, edits: SheetStructuralEdit[] | undefined): Promise<NoteWrites | undefined> => {
    const plan = notePlanOf(sheet);
    const rels = sheetRels ? childrenNamed(parseXml(sheetRels).root, 'Relationship') : [];
    const out: NoteWrites = { parts: new Map(), fresh: [], relate: [], unrelate: new Set(), legacy: undefined, names: {} };
    const before = noted.get(sheet) || {};
    const ids = new Set<string>();
    let highestId = 0;
    for (const rel of rels) {
      const id = rel.attrs.get('Id') || '';
      ids.add(id);
      const match = /^rId(\d{1,9})$/.exec(id);
      if (match && Number(match[1]) > highestId) highestId = Number(match[1]);
    }
    const newId = (): string => {
      for (;;) {
        const id = `rId${++highestId}`;
        if (!ids.has(id)) {
          ids.add(id);
          return id;
        }
      }
    };
    const dir = main.slice(0, slash);

    // The comments: what the file has, moved along with the rows and columns, then brought in line with the model.
    const scan: CommentsScan = { seen: new Set(), known: 0 };
    const removedCells = new Set<number>();
    // Once a sort moved notes, each comment of the file is placed on its own (see NoteMoves), and the parts
    // are written by this save whatever they come out as: the edits alone would put the comments elsewhere.
    const moves = plan.moves;
    const goneFile = new Set<number>();
    const seenFile = new Set<number>();
    let changed = false;
    let first = true;
    for (const rel of rels) {
      if (!endsWith(typeOf(rel), '/comments') || !internal(rel)) continue;
      const part = actualOf(resolveTarget(sheetPart, rel.attrs.get('Target') || ''));
      const original = await peek(part);
      if (!original) continue;
      const staged = moves ? placeComments(original, edits || [], moves, goneFile, seenFile) : edits ? adjustComments(original, edits) : original;
      const edit = editComments(staged, plan, first ? plan.added : [], scan, removedCells);
      first = false;
      if (!edit.changed && !moves) continue;
      changed = true;
      if (edit.left === 0) {
        out.parts.set(part, null);
        out.unrelate.add(rel.attrs.get('Id') || '');
      } else out.parts.set(part, edit.bytes);
    }
    if (first && plan.added.length > 0) {
      const part = before.comments || freeName(`${dir}/comments`, '.xml');
      const skeleton = encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${escapeAttr(mainNamespace)}"><authors></authors><commentList></commentList></comments>`);
      const edit = editComments(skeleton, plan, plan.added, { seen: new Set(), known: 0 }, removedCells);
      out.fresh.push([part, edit.bytes, COMMENTS_CONTENT_TYPE]);
      out.relate.push({ id: newId(), kind: 'comments', part });
      out.names.comments = part;
      changed = true;
    }

    // Their boxes, in the drawing <legacyDrawing> names.
    const legacyId = sheetXml === undefined ? undefined : legacyDrawingId(sheetXml);
    const vmlRel = legacyId === undefined ? undefined : rels.find((rel) => rel.attrs.get('Id') === legacyId && internal(rel));
    const vmlPart = vmlRel ? actualOf(resolveTarget(sheetPart, vmlRel.attrs.get('Target') || '')) : undefined;
    const vmlBytes = vmlPart === undefined ? undefined : await peek(vmlPart);
    if (vmlPart !== undefined && vmlBytes && vmlRel) {
      if (removedCells.size > 0 || plan.added.length > 0 || moves) {
        const staged = moves ? placeVml(vmlBytes, placeNoteBy(edits || [], moves, goneFile)) : edits ? adjustVml(vmlBytes, edits) : vmlBytes;
        const edit = editVml(staged, removedCells, plan.added, await useBlocks());
        if (edit.changed || moves) {
          changed = true;
          if (edit.changed ? edit.empty : staged !== vmlBytes && vmlDrawsNothing(staged)) {
            out.parts.set(vmlPart, null);
            out.unrelate.add(vmlRel.attrs.get('Id') || '');
            out.legacy = null;
          } else out.parts.set(vmlPart, edit.bytes);
        }
      }
    } else if (plan.added.length > 0) {
      const part = before.vml || freeName(`${dir}/drawings/vmlDrawing`, '.vml');
      const skeleton = encoder.encode(`<xml xmlns:v="${NS_VML}" xmlns:o="${NS_VML_OFFICE}" xmlns:x="${NS_VML_EXCEL}"></xml>`);
      const edit = editVml(skeleton, new Set(), plan.added, await useBlocks());
      out.fresh.push([part, edit.bytes, VML_CONTENT_TYPE]);
      const id = newId();
      out.relate.push({ id, kind: 'vmlDrawing', part });
      out.legacy = id;
      out.names.vml = part;
      changed = true;
    }
    if (!changed) return undefined;
    // One sheet's notes cannot change a part that another sheet shows as well.
    for (const part of out.parts.keys()) {
      for (const other of workbook.sheets) {
        if (other === sheet || other.kind === 'chartsheet' || other.part === '') continue;
        for (const otherRel of await pkg.relationships(other.part)) {
          if (otherRel.part !== null && actualOf(otherRel.part) === part) throw new Error(`Cannot write sheet "${sheet.name}": ${part} is shared with sheet "${other.name}"`);
        }
      }
    }
    return out;
  };
  /** The relationships of a sheet with what its notes ask for; undefined when they ask for nothing. */
  const relateNotes = (doc: XmlDocument, sheetPart: string, notes: NoteWrites): boolean => {
    let changed = false;
    if (notes.unrelate.size > 0) {
      const kept = doc.root.children.filter((node) => !(node.type === 'element' && localName(node.name) === 'Relationship' && notes.unrelate.has(node.attrs.get('Id') || '')));
      if (kept.length !== doc.root.children.length) {
        doc.root.children = kept;
        changed = true;
      }
    }
    for (const rel of notes.relate) {
      doc.root.children.push(
        element(`${prefixOf(doc.root)}Relationship`, [
          ['Id', rel.id],
          ['Type', `${relBase}/${rel.kind}`],
          ['Target', relativeTarget(sheetPart, rel.part)],
        ]),
      );
      changed = true;
    }
    return changed;
  };
  const emptyRels = (): XmlDocument => parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');


  // The notes of the sheets of the file are looked at first: whether anything of them is
  // written decides, with the rest, whether the file counts as edited.
  const fileNotes = new Map<Sheet, NoteWrites>();
  for (const sheet of workbook.sheets) {
    if (!sheet.commentsModified || sheet.kind === 'chartsheet' || sheet.part === '') continue;
    const bytes = await peek(sheet.part);
    if (!bytes) continue;
    const edits = sheet.modified && sheet.structuralEdits && sheet.structuralEdits.length > 0 ? sheet.structuralEdits : undefined;
    const notes = await notesOf(sheet, sheet.part, decodeSource(bytes), await peek(relsPartOf(sheet.part)), edits);
    if (notes) fileNotes.set(sheet, notes);
  }

  const active = activeSheetOf(workbook);
  // The sheet the file opens on follows the model, but only along with something else that is
  // written: a file that was looked at and not edited is saved byte for byte, whatever sheet was
  // shown last. A list of sheets that is marked as changed and is the list of the file again (a
  // sheet inserted and taken back) is no change. The model is compared as the reader made it
  // from the file, so that a file whose active sheet is hidden is not corrected on the way.
  // Not in a file with several windows: each has its own active sheet and its own view of every
  // sheet, and which of them the user looked at is not known. (Once sheets are added or deleted
  // every window has to be brought in line anyway, and all of them open on the active sheet.)
  // And only when the sheet that would be written is another one than the file has: a hidden
  // sheet as the active one of the model stands for the first visible one.
  // The colours of tabs: a sheet whose part says what the model says is not written for it.
  // A colour that is the one the sheet came with again (undone, or set back) costs nothing: the part is not
  // even read. Any other is put into the part here, once.
  const tabbed = new Map<Sheet, Uint8Array>();
  for (const sheet of workbook.sheets) {
    if (!sheet.tabColorModified || sheet.kind !== 'worksheet' || sheet.part === '') continue;
    if (sheet.tabColorRead && sameColor(sheet.tabColorRead.color, sheet.tabColor)) continue;
    const bytes = await originalOf(sheet.part);
    const coloured = bytes ? setTabColorIn(bytes, sheet.tabColor, sheet.name) : undefined;
    if (coloured) tabbed.set(sheet, coloured);
  }

  const edited = modifiedSheets.length > 0 || stylesModified || reordered || renamed || restated || fileNotes.size > 0 || tabbed.size > 0;
  const fileActive = fileActiveOf(book.root, file.length);
  const bookViews = childNamed(book.root, 'bookViews');
  const windows = bookViews ? childrenNamed(bookViews, 'workbookView').length : 0;
  const activeMoved = edited && !reordered && windows <= 1 && modelActiveOf(workbook) !== fileActive && active !== fileActive;
  // A file with several windows keeps the sheet of each - but not a sheet that was hidden since: such a
  // window opens on the active sheet of the model (the first visible one when that is hidden itself).
  let hiddenWindows: Set<number> | undefined;
  if (restated && !reordered && windows > 1 && bookViews) {
    childrenNamed(bookViews, 'workbookView').forEach((view, index) => {
      const stated = view.attrs.get('activeTab') || '';
      const at = /^\d{1,6}$/.test(stated) && Number(stated) < file.length ? Number(stated) : 0;
      const sheet = workbook.sheets[at];
      if (sheet && sheet.state !== 'visible' && fileStateOf(file[at].element) === 'visible') (hiddenWindows || (hiddenWindows = new Set())).add(index);
    });
  }

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
  const moving = workbook.sheets.filter((sheet) => sheet.kind !== 'chartsheet' && sheet.part !== '' && sheet.structuralEdits && sheet.structuralEdits.length > 0);
  const foreignOf = (sheet: Sheet): StructuralEdit[] => {
    let out: StructuralEdit[] = [];
    for (const other of moving) if (other !== sheet) out = out.concat(editsOf(other));
    return out;
  };

  const movedNow = new Map<string, string[]>();
  for (let index = 0; index < workbook.sheets.length; index++) {
    const sheet = workbook.sheets[index];
    if (sheet.part === '') continue;
    const before = moved.get(sheet.part) || [];
    if (sheet.kind === 'chartsheet') {
      // A chart sheet is not in the model. Its part only ever changes in what it says about being the selected sheet.
      const changed: string[] = [];
      if (reordered || activeMoved) {
        const bytes = await peek(sheet.part);
        const next = bytes ? selectChartTab(bytes, index === active) : undefined;
        if (bytes && next) {
          keep(sheet.part, bytes);
          writes.push([sheet.part, next]);
          changed.push(sheet.part);
        }
      }
      for (const part of before) {
        const original = originals.get(part);
        if (changed.indexOf(part) < 0 && original) writes.push([part, original]);
      }
      movedNow.set(sheet.part, changed);
      continue;
    }
    const fileName = fileNameOf(sheet) as string;
    const foreign = foreignOf(sheet);
    // Once sheets were added or deleted, the one sheet marked as selected is the active one.
    // The same when another sheet of the file is the active one: the sheet that was marked and the one that is now are written.
    // In a file with several windows a sheet has a view for each, and the model knows one of them:
    // every sheet is looked at, and written when one of its views says something else.
    let tabSelected = (reordered || activeMoved) && !!sheet.view.tabSelected !== (index === active) ? index === active : undefined;
    if (reordered && windows > 1) tabSelected = index === active;
    // Windows that showed a sheet which is hidden now: that sheet is the selected one of no window, and
    // the active sheet is the selected one of those windows.
    let tabWindows: Set<number> | undefined;
    if (hiddenWindows) {
      if (sheet.state !== 'visible') tabSelected = false;
      else if (index === active) {
        tabSelected = true;
        tabWindows = hiddenWindows;
      }
    }
    // Notes that were added, edited or deleted: the comment parts of the sheet follow the model.
    const notes = fileNotes.get(sheet);
    const tab = tabbed.has(sheet);
    if (!sheet.modified && foreign.length === 0 && !rename && tabSelected === undefined && before.length === 0 && !notes && !tab) continue;
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

    const edits = sheet.modified && sheet.structuralEdits && sheet.structuralEdits.length > 0 ? sheet.structuralEdits : undefined;
    const legacyDrawing = notes ? notes.legacy : undefined;
    if (!sheet.modified) {
      // A sheet that was not modified itself: only its raw content that names an edited or a renamed
      // sheet changes (the source of a validation list on another sheet, say). Its cells were adjusted
      // through the model, which marked the sheet as modified if any of them changed.
      const bytes = await peek(sheet.part);
      let next: Uint8Array | undefined;
      // Only the notes of the sheet changed: every byte of the part stays but the element that names their boxes.
      if (bytes && legacyDrawing !== undefined && foreign.length === 0 && !rename && tabSelected === undefined) next = setLegacyDrawing(bytes, legacyDrawing, sheet.name);
      else if (bytes && (foreign.length > 0 || rename || tabSelected !== undefined || legacyDrawing !== undefined)) {
        next = rewriteSheet(bytes, sheet, workbook, { foreign, rawOnly: true, fileName, rename, tabSelected, tabWindows, legacyDrawing });
      }
      // The colour of the tab, in the part as it is by now.
      if (bytes && tab) next = next ? setTabColorIn(next, sheet.tabColor, sheet.name) || next : tabbed.get(sheet);
      if (bytes && next) {
        // From here on the package no longer holds the part as the file had it.
        keep(sheet.part, bytes);
        writes.push([sheet.part, next]);
        rewritten.push(sheet.part);
      }
    } else {
      const bytes = await originalOf(sheet.part);
      if (!bytes) throw new Error(`Cannot write sheet "${sheet.name}": its part ${sheet.part} is missing`);
      const written = rewriteSheet(bytes, sheet, workbook, { foreign, overrides, fileName, rename, tabSelected, tabWindows, legacyDrawing }) as Uint8Array;
      writes.push([sheet.part, (tab && setTabColorIn(written, sheet.tabColor, sheet.name)) || written]);
      rewritten.push(sheet.part);
    }

    // After rows or columns moved: the pictures and comments of the sheet follow, and the
    // relationships of links that are gone are removed. After rows or columns moved in another
    // sheet, or sheets were renamed or deleted: the shapes that show a cell of such a sheet follow.
    const sheetRelsBytes = edits || others || notes ? await peek(sheetRelsPart) : undefined;
    if (sheetRelsBytes || notes) {
      const sheetRels = sheetRelsBytes ? parseXml(sheetRelsBytes) : emptyRels();
      const linkIds = new Set<string>();
      for (const link of sheet.hyperlinks) if (link.rId !== undefined) linkIds.add(link.rId);
      let dropped = false;
      // Links of cells that are written as a target ("#Sheet2!A1") instead of a location.
      if (sheetRelsBytes && (edits || others) && mayHoldPlaces(sheetRelsBytes) && adjustPlaceTargets(sheetRels.root, edits || [], fileName, others)) dropped = true;
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
        const part = actualOf(resolveTarget(sheet.part, rel.attrs.get('Target') || ''));
        // Written below, moved along and with the notes of the model.
        if (notes && notes.parts.has(part)) continue;
        const original = await peek(part);
        if (!original) continue;
        if (edits) {
          // One sheet's rows cannot move a part that another sheet shows as well.
          for (const other of workbook.sheets) {
            if (other === sheet || other.kind === 'chartsheet' || other.part === '') continue;
            for (const otherRel of await pkg.relationships(other.part)) {
              if (otherRel.part !== null && actualOf(otherRel.part) === part) throw new Error(`Cannot write sheet "${sheet.name}": ${part} is shared with sheet "${other.name}"`);
            }
          }
        }
        const adjusted = adjust(original, edits || [], fileName, others);
        if (adjusted !== original) {
          keep(part, original);
          writes.push([part, adjusted]);
          rewritten.push(part);
        }
        if (kind !== 'drawing') continue;
        // The links of its shapes and pictures to places in the workbook are in its relationships.
        const partRelsName = relsPartOf(part);
        const partRelsBytes = await peek(partRelsName);
        if (!partRelsBytes || !mayHoldPlaces(partRelsBytes)) continue;
        const partRels = parseXml(partRelsBytes);
        if (adjustPlaceTargets(partRels.root, edits || [], fileName, others)) {
          keep(partRelsName, partRelsBytes);
          writes.push([partRelsName, partRels]);
          rewritten.push(partRelsName);
        }
      }
      if (notes) {
        for (const [part, data] of notes.parts) {
          const original = await originalOf(part);
          if (data === null) wanted.add(part);
          else if (original) {
            writes.push([part, data]);
            rewritten.push(part);
          }
        }
        for (const [part, data, contentType] of notes.fresh) {
          if (!originals.has(part)) originals.set(part, null);
          writes.push([part, data, contentType]);
        }
        if (relateNotes(sheetRels, sheet.part, notes)) dropped = true;
        if (!sheetRelsBytes && dropped) {
          // The file has no relationships for the sheet: the part made for them goes when it is no longer needed.
          if (!originals.has(sheetRelsPart)) originals.set(sheetRelsPart, null);
          notes.names.rels = sheetRelsPart;
        }
        notesNow.set(sheet, notes.names);
      }
      if (dropped) {
        if (sheetRelsBytes) {
          keep(sheetRelsPart, sheetRelsBytes);
          rewritten.push(sheetRelsPart);
        }
        writes.push([sheetRelsPart, sheetRels]);
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
  for (let index = 0; index < workbook.sheets.length; index++) {
    const sheet = workbook.sheets[index];
    if (sheet.part !== '') continue;
    if (sheet.kind !== 'worksheet') throw new Error(`Cannot write sheet "${sheet.name}": a ${sheet.kind} that is not in the file cannot be written`);
    let part = added.get(sheet);
    if (part === undefined) {
      const names = takenNames();
      for (let n = 1; part === undefined; n++) {
        const candidate = `${main.slice(0, slash)}/worksheets/sheet${n}.xml`;
        // Relationships left behind under the name would become those of the new sheet.
        if (!names.has(candidate.toLowerCase()) && !names.has(relsPartOf(candidate).toLowerCase())) part = candidate;
      }
      names.add(part.toLowerCase());
    }
    // The notes of the sheet: all of them were added, and so are the parts that hold them.
    const notes = sheet.comments && sheet.comments.length > 0 ? await notesOf(sheet, part, undefined, undefined, undefined) : undefined;
    // What the model holds of the sheet beside its cells (a copy of a sheet has all of it): the colour of its
    // tab, how it is shown and printed, its default sizes. An empty new sheet has none of these.
    const around = freshSheetParts(sheet, index === active);
    const skeleton =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${escapeAttr(mainNamespace)}" xmlns:r="${escapeAttr(relBase)}">` +
      `${around.properties}<dimension ref="A1"/><sheetViews>${around.view}</sheetViews>${around.format}<sheetData/>${around.protection}${around.page}</worksheet>`;
    // A link to another file belongs to the relationships of the sheet it was made in; only places in the workbook are written.
    const hyperlinks = sheet.hyperlinks.filter((link) => link.location !== undefined).map((link) => (link.rId === undefined ? link : { ...link, rId: undefined }));
    const bytes = rewriteSheet(encoder.encode(skeleton), { ...sheet, hyperlinks }, workbook, { fresh: true, legacyDrawing: notes ? notes.legacy : undefined }) as Uint8Array;
    if (!originals.has(part)) originals.set(part, null);
    writes.push([part, bytes, WORKSHEET_CONTENT_TYPE]);
    if (notes) {
      const sheetRels = emptyRels();
      relateNotes(sheetRels, part, notes);
      const sheetRelsPart = relsPartOf(part);
      for (const [name, data, contentType] of notes.fresh.concat([[sheetRelsPart, sheetRels]])) {
        if (!originals.has(name)) originals.set(name, null);
        writes.push([name, data, contentType]);
      }
      notes.names.rels = sheetRelsPart;
      notesNow.set(sheet, notes.names);
    }
    addedNow.set(sheet, part);
    relIds.set(sheet, relate('worksheet', part));
  }
  const deletions: string[] = [];
  for (const [sheet, part] of added) if (addedNow.get(sheet) !== part) deletions.push(part);
  // Parts made for notes that are no longer there (undone, deleted, or their sheet is gone).
  const noteParts = new Set<string>();
  for (const names of notesNow.values()) for (const name of [names.comments, names.vml, names.rels]) if (name !== undefined) noteParts.add(name);
  for (const names of noted.values()) for (const name of [names.comments, names.vml, names.rels]) if (name !== undefined && !noteParts.has(name)) deletions.push(name);

  // Sheets of the file that are gone: their parts, and what only they lead to, leave the package.
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
    // A modified sheet of the file stays modified, and the chain stays out with it. Without one,
    // undoing may bring back the sheets of the file as they were, and the chain with them: a sheet
    // that is not in the file is modified from the start, and goes without a trace.
    if (modifiedSheets.some((sheet) => sheet.part !== '')) calcChain = chain;
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
  if (recalculate || renamed || activeMoved || restated) {
    updateWorkbookXml(book, workbook, { file, renamed, restated, reordered, relIds, relNamespace: relBase, active, activeMoved, recalculate, hiddenWindows });
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
  noted.clear();
  for (const [sheet, names] of notesNow) noted.set(sheet, names);
}
