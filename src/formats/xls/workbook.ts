// Reads an xls workbook (BIFF8, the format of Excel 97-2003, and BIFF5) into
// the spreadsheet model.
//
// The workbook is one stream of records: a "globals" substream (fonts, formats,
// sheet list, shared strings, names) and one substream per sheet. Nothing is
// written back: the model is for viewing, and for saving as xlsx later.

import type { CfbFile } from '../../cfb';
import {
  cellName,
  createSheet,
  rangeName,
  type Cell,
  type CellFormula,
  type CellRange,
  type ColumnRange,
  type DefinedName,
  type Font,
  type Row,
  type SharedString,
  type Sheet,
  type SheetKind,
  type SheetState,
  type StyleSheet,
  type Workbook,
} from '../../model/sheet/types';
import { isKnownCodepage } from './codepage';
import { errorText, formulaText, releaseFormulaBytes, type FormulaContext, type XlsSupBook, type XlsXti } from './formula';
import { readHyperlink, readHyperlinkTooltip } from './hyperlink';
import {
  REC_ARRAY,
  REC_BLANK,
  REC_BOF,
  REC_BOOLERR,
  REC_BOUNDSHEET,
  REC_CODEPAGE,
  REC_COLINFO,
  REC_CONTINUE,
  REC_CONTINUEFRT12,
  REC_DATEMODE,
  REC_DEFAULTROWHEIGHT,
  REC_DEFCOLWIDTH,
  REC_DIMENSION,
  REC_EOF,
  REC_EXTERNNAME,
  REC_EXTERNSHEET,
  REC_FILEPASS,
  REC_FONT,
  REC_FORMAT,
  REC_FORMULA,
  REC_HLINK,
  REC_HLINKTOOLTIP,
  REC_LABEL,
  REC_LABELSST,
  REC_MERGECELLS,
  REC_MULBLANK,
  REC_MULRK,
  REC_NAME,
  REC_NOTE,
  REC_NUMBER,
  REC_PALETTE,
  REC_PANE,
  REC_RK,
  REC_ROW,
  REC_RSTRING,
  REC_SCL,
  REC_SELECTION,
  REC_SHRFMLA,
  REC_SST,
  REC_STANDARDWIDTH,
  REC_STRING,
  REC_STYLE,
  REC_SUPBOOK,
  REC_TABLE,
  REC_THEME,
  REC_WINDOW1,
  REC_WINDOW2,
  REC_WSBOOL,
  REC_XF,
  REC_XFEXT,
  XlsUnsupportedError,
  corrupt,
  quote,
  readerAt,
  recordLength,
  recordStream,
  type RecordStream,
} from './records';
import { readByteString, readSst, readString16, readText16, readText8, toSharedString, type RawString } from './strings';
import { XlsStyles } from './styles';

/** Sheet limits of the format; a cell outside them makes the file invalid. */
export const XLS_MAX_ROWS = 65536;
export const XLS_MAX_COLS = 256;

const REC_SHEETEXT = 0x0862;

/** What reading left out or could not do, for diagnostics. None of it stops a file from opening. */
export interface XlsReadReport {
  /** 8 for the format of Excel 97-2003, 5 for the one of Excel 5 and 95. */
  biff: 5 | 8;
  codepage: number;
  /** Formula cells whose tokens could not be turned into text; they keep their value and have no formula. */
  untranslatedFormulas: number;
  /** Defined names left out for the same reason. */
  untranslatedNames: number;
  /** Why: reason -> number of formulas and names. */
  untranslatedReasons: Map<string, number>;
  /** Records that were not read: record id -> count. */
  skippedRecords: Map<number, number>;
  /** Cell comments (NOTE records); the model has no place for them. */
  comments: number;
}

export interface XlsReadResult {
  workbook: Workbook;
  report: XlsReadReport;
}

/** Built-in defined names, which the file stores as a number. */
const BUILTIN_NAMES: readonly string[] = [
  'Consolidate_Area',
  'Auto_Open',
  'Auto_Close',
  'Extract',
  'Database',
  'Criteria',
  'Print_Area',
  'Print_Titles',
  'Recorder',
  'Data_Form',
  'Auto_Activate',
  'Auto_Deactivate',
  'Sheet_Title',
  '_FilterDatabase',
];

interface BoundSheet {
  offset: number;
  state: SheetState;
  type: number;
  name: string;
}

interface RawName {
  name: string;
  hidden: boolean;
  /** 1-based sheet the name is local to; 0 for a workbook name. */
  sheet: number;
  /** Tokens, followed by the data that belongs to them. */
  data: Uint8Array;
  tokenSize: number;
}

interface Globals {
  biff8: boolean;
  codepage: number;
  date1904: boolean;
  activeTab: number;
  sheets: BoundSheet[];
  styles: StyleSheet;
  /** XF index of the file -> index into styles.cellXfs; the last entry answers indexes outside the table. */
  xfMap: number[];
  sharedStrings: SharedString[];
  rawNames: RawName[];
  /** The model's font for a font index of the file. */
  fontOf(index: number): Font | undefined;
  ctx: FormulaContext;
  report: XlsReadReport;
}

function count(map: Map<number, number>, id: number): void {
  map.set(id, (map.get(id) || 0) + 1);
}

/** SUPBOOK (BIFF8): a workbook that formulas refer to, or a marker for this workbook or for add-ins. */
function readSupBook(stream: RecordStream, offset: number, ordinal: number): XlsSupBook {
  const r = readerAt(stream, offset);
  const sheetCount = r.u16();
  const length = r.u16();
  if (r.remaining === 0 && length === 0x0401) return { kind: 'self', sheetNames: [], externNames: [], ordinal: 0 };
  if (r.remaining === 0 && length === 0x3a01) return { kind: 'addin', sheetNames: [], externNames: [], ordinal: 0 };
  // The path of the workbook; the model has no place for it.
  const wide = (r.u8() & 1) !== 0;
  r.chars(length, wide);
  const sheetNames: string[] = [];
  for (let i = 0; i < sheetCount && !r.done; i++) sheetNames.push(readString16(r).text);
  return { kind: 'external', sheetNames, externNames: [], ordinal };
}

/** EXTERNSHEET (BIFF5): one record per sheet or workbook that formulas refer to. */
function readExternSheet5(stream: RecordStream, offset: number, codepage: number, ordinal: number): XlsSupBook {
  const r = readerAt(stream, offset);
  const length = r.u8();
  const code = r.remaining > 0 ? r.u8() : 0;
  if (code === 0x03) {
    // A sheet of this workbook, by name.
    return { kind: 'self', sheetNames: [readByteString(r, Math.min(length, r.remaining), codepage)], externNames: [], ordinal: 0 };
  }
  if (code === 0x02) return { kind: 'self', sheetNames: [], externNames: [], ordinal: 0 };
  if (code === 0x04 || code === 0x3a) return { kind: 'addin', sheetNames: [], externNames: [], ordinal: 0 };
  return { kind: 'external', sheetNames: [], externNames: [], ordinal };
}

function readName(stream: RecordStream, offset: number, biff8: boolean, codepage: number): RawName {
  const r = readerAt(stream, offset);
  const flags = r.u16();
  r.skip(1);
  const nameLength = r.u8();
  const tokenSize = r.u16();
  r.skip(2);
  const sheet = r.u16();
  r.skip(4);
  let name: string;
  if (biff8) {
    const wide = (r.u8() & 1) !== 0;
    name = r.chars(nameLength, wide);
  } else name = readByteString(r, nameLength, codepage);
  if (flags & 0x0020) {
    // A built-in name is stored as one character: its number.
    const builtin = BUILTIN_NAMES[name.charCodeAt(0)];
    name = builtin === undefined ? `_xlnm.Builtin_${name.charCodeAt(0)}` : `_xlnm.${builtin}`;
  }
  return { name, hidden: (flags & 0x0001) !== 0, sheet, data: r.rest(), tokenSize };
}

function readGlobals(stream: RecordStream, streamName: string): Globals {
  const { bytes, view } = stream;
  const size = bytes.length;
  if (size < 4) throw corrupt('the workbook stream is empty');
  const firstId = view.getUint16(0, true);
  if (firstId === 0x0009 || firstId === 0x0209 || firstId === 0x0409) {
    throw new XlsUnsupportedError('Excel files older than Excel 5 are not supported');
  }
  const bofLength = recordLength(stream, 0);
  if (firstId !== REC_BOF || bofLength < 4) throw corrupt('the workbook stream does not start with a BOF record');
  const version = view.getUint16(4, true);
  // Some writers leave the version empty; the name of the stream then tells the format.
  const biff8 = version === 0x0600 || (version !== 0x0500 && streamName === 'Workbook');
  if (view.getUint16(6, true) !== 0x0005) throw new XlsUnsupportedError('The xls file does not start with workbook globals');

  const report: XlsReadReport = {
    biff: biff8 ? 8 : 5,
    codepage: 1252,
    untranslatedFormulas: 0,
    untranslatedNames: 0,
    untranslatedReasons: new Map(),
    skippedRecords: new Map(),
    comments: 0,
  };
  let codepage = 1252;
  let date1904 = false;
  let activeTab = 0;
  let sstOffset = -1;
  const sheets: BoundSheet[] = [];
  const styleRecords = new XlsStyles(biff8, () => codepage);
  const rawNames: RawName[] = [];
  const supBooks: XlsSupBook[] = [];
  const xti: XlsXti[] = [];
  let externalBooks = 0;

  let p = 4 + bofLength;
  let depth = 0;
  for (;;) {
    if (p >= size) throw corrupt('the workbook globals are not terminated');
    const length = recordLength(stream, p);
    const id = view.getUint16(p, true);
    const data = p + 4;
    if (id === REC_EOF) {
      if (depth === 0) break;
      depth--;
    } else if (id === REC_BOF) depth++;
    else if (depth === 0) {
      switch (id) {
        case REC_FILEPASS:
          throw new XlsUnsupportedError('Password-protected xls files are not supported');
        case REC_CODEPAGE:
          if (length >= 2) {
            codepage = view.getUint16(data, true);
            report.codepage = codepage;
          }
          break;
        case REC_DATEMODE:
          date1904 = length >= 2 && view.getUint16(data, true) === 1;
          break;
        case REC_WINDOW1:
          if (length >= 12) activeTab = view.getUint16(data + 10, true);
          break;
        case REC_FONT:
        case 0x0231:
          styleRecords.addFont(readerAt(stream, p));
          break;
        case REC_FORMAT:
          styleRecords.addFormat(readerAt(stream, p));
          break;
        case REC_XF:
          styleRecords.addXf(readerAt(stream, p));
          break;
        case REC_XFEXT:
          styleRecords.addXfExt(readerAt(stream, p));
          break;
        case REC_STYLE:
          styleRecords.addStyle(readerAt(stream, p));
          break;
        case REC_PALETTE:
          styleRecords.addPalette(readerAt(stream, p));
          break;
        case REC_THEME:
          styleRecords.addTheme(readerAt(stream, p));
          break;
        case REC_BOUNDSHEET: {
          const r = readerAt(stream, p);
          const offset = r.u32();
          const visibility = r.u8() & 0x03;
          const type = r.u8();
          sheets.push({
            offset,
            state: visibility === 1 ? 'hidden' : visibility === 2 ? 'veryHidden' : 'visible',
            type,
            name: readText8(r, biff8, codepage),
          });
          break;
        }
        case REC_SST:
          sstOffset = p;
          break;
        case REC_SUPBOOK: {
          let book: XlsSupBook;
          try {
            book = readSupBook(stream, p, externalBooks + 1);
          } catch {
            // A source this reader cannot make out still takes its place: EXTERNSHEET refers to it by position.
            book = { kind: 'external', sheetNames: [], externNames: [], ordinal: externalBooks + 1 };
          }
          if (book.kind === 'external') externalBooks++;
          supBooks.push(book);
          break;
        }
        case REC_EXTERNSHEET:
          if (biff8) {
            const r = readerAt(stream, p);
            const entries = r.u16();
            for (let i = 0; i < entries && r.remaining >= 6; i++) xti.push({ book: r.u16(), first: r.i16(), last: r.i16() });
          } else {
            const book = readExternSheet5(stream, p, codepage, externalBooks + 1);
            if (book.kind === 'external') externalBooks++;
            supBooks.push(book);
          }
          break;
        case REC_EXTERNNAME: {
          // Belongs to the workbook named last. Only its name is needed: formulas refer to it by position.
          const book = supBooks[supBooks.length - 1];
          if (!book) break;
          let name = '';
          try {
            const r = readerAt(stream, p);
            r.skip(6);
            name = readText8(r, biff8, codepage);
          } catch {
            // A DDE or OLE item laid out differently: it keeps its position, without a name.
          }
          book.externNames.push(name);
          break;
        }
        case REC_NAME:
          rawNames.push(readName(stream, p, biff8, codepage));
          break;
        case REC_CONTINUE:
        case REC_CONTINUEFRT12:
          break;
        default:
          count(report.skippedRecords, id);
      }
    }
    p = data + length;
  }

  const built = styleRecords.build();
  const fontOf = (index: number) => built.styles.fonts[styleRecords.fontIndex(index)];
  const sharedStrings = sstOffset < 0 ? [] : readSst(readerAt(stream, sstOffset)).map((raw) => toSharedString(raw, fontOf));

  const ctx: FormulaContext = {
    biff8,
    codepage: isKnownCodepage(codepage) ? codepage : 1252,
    sheetNames: sheets.map((sheet) => sheet.name),
    names: rawNames.map((name) => name.name),
    xti,
    supBooks,
    failures: report.untranslatedReasons,
  };
  return { biff8, codepage: ctx.codepage, date1904, activeTab, sheets, styles: built.styles, xfMap: built.xfMap, sharedStrings, rawNames, fontOf, ctx, report };
}

/** The defined names of the workbook; a name whose formula cannot be turned into text is left out. */
function definedNamesOf(globals: Globals): DefinedName[] {
  const names: DefinedName[] = [];
  for (const raw of globals.rawNames) {
    // A name without a formula stands for a function of an add-in or a macro; it defines nothing.
    if (raw.tokenSize === 0 || raw.tokenSize > raw.data.length) continue;
    const value = formulaText(globals.ctx, raw.data, 0, raw.tokenSize, raw.tokenSize, raw.data.length, 0, 0);
    if (value === undefined) {
      globals.report.untranslatedNames++;
      continue;
    }
    const name: DefinedName = { name: raw.name, value, localSheet: undefined, hidden: raw.hidden };
    if (raw.sheet >= 1 && raw.sheet <= globals.sheets.length) name.localSheet = raw.sheet - 1;
    names.push(name);
  }
  return names;
}

/**
 * Adds a column range, keeping the list sorted and free of overlaps. Where the new range
 * overlaps earlier ones the earlier ones win and only the uncovered parts are added.
 */
function addColumnRange(columns: ColumnRange[], range: ColumnRange): void {
  const last = columns[columns.length - 1];
  if (!last || range.min > last.max) {
    columns.push(range);
    return;
  }
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

const rkScratch = new DataView(new ArrayBuffer(8));

/** The number an RK value stands for: a 30-bit integer or the upper 30 bits of a double, optionally divided by 100. */
export function decodeRk(rk: number): number {
  let value: number;
  if (rk & 0x02) value = rk >> 2;
  else {
    rkScratch.setUint32(0, 0, true);
    rkScratch.setUint32(4, (rk & 0xfffffffc) >>> 0, true);
    value = rkScratch.getFloat64(0, true);
  }
  return rk & 0x01 ? value / 100 : value;
}

/** A shared formula (SHRFMLA): tokens that every cell of the group reads from its own position. */
interface SharedFormula {
  bytes: Uint8Array;
  start: number;
  end: number;
  extraEnd: number;
  range: CellRange;
  si: number;
  /** The first cell of the group has been given the range. */
  placed: boolean;
}

/** An ARRAY or TABLE record: the formula belongs to the first cell of the range. */
interface RangeFormula {
  formula: CellFormula | undefined;
  range: CellRange;
}

const key = (row: number, col: number) => row * XLS_MAX_COLS + col;

function readSheet(stream: RecordStream, start: number, sheet: Sheet, globals: Globals): void {
  const { bytes, view } = stream;
  const size = bytes.length;
  const { biff8, codepage, xfMap, sharedStrings, ctx, report } = globals;
  const xfCount = xfMap.length - 1;
  const rows = sheet.rows;
  const name = quote(sheet.name);

  let lastRowIndex = -1;
  let lastRow: Row | undefined;
  const rowOf = (index: number): Row => {
    if (index === lastRowIndex) return lastRow as Row;
    let row = rows.get(index);
    if (!row) {
      row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
      rows.set(index, row);
    }
    lastRowIndex = index;
    lastRow = row;
    return row;
  };
  const put = (row: number, col: number, cell: Cell): void => {
    if (col >= XLS_MAX_COLS) throw corrupt(`a cell of sheet ${name} is in column ${col + 1}, past the last column`);
    // A repeated cell replaces the earlier one.
    rowOf(row).cells[col] = cell;
    if (row >= sheet.rowCount) sheet.rowCount = row + 1;
    if (col >= sheet.colCount) sheet.colCount = col + 1;
  };
  const styleOf = (xf: number): number => xfMap[xf < xfCount ? xf : xfCount];
  const short = (what: string): Error => corrupt(`a ${what} record of sheet ${name} is too short`);

  // Cells whose formula is "the one of the group that starts at ...": resolved when the whole sheet is read.
  const linkedCells: Cell[] = [];
  const linkedAt: number[] = [];
  const linkedTo: number[] = [];
  const shared = new Map<number, SharedFormula>();
  const ranged = new Map<number, RangeFormula>();
  let nextShared = 0;
  /** The group the last FORMULA record points to; SHRFMLA, ARRAY and TABLE records define that group. */
  let lastLink = -1;
  let lastFormulaAt = -1;
  /** The formula cell whose text result is in the STRING record that follows. */
  let awaitingString: Cell | undefined;

  let frozen = false;
  let activePane = 3;
  let firstActive: string | undefined;
  let paneActive: string | undefined;
  let columnRecords = 0;
  let defaultWidthChars = -1;
  let standardWidth = false;
  const tooltips: { top: number; left: number; bottom: number; right: number; text: string }[] = [];

  let p = start + 4 + recordLength(stream, start);
  let depth = 0;
  for (;;) {
    if (p + 4 > size) throw corrupt(`sheet ${name} is not terminated`);
    const id = view.getUint16(p, true);
    const length = view.getUint16(p + 2, true);
    const d = p + 4;
    const next = d + length;
    if (next > size) throw corrupt(`the record at offset ${p} runs past the end of the stream`);
    p = next;

    // An embedded chart is a substream of its own inside the sheet.
    if (depth > 0) {
      if (id === REC_BOF) depth++;
      else if (id === REC_EOF) depth--;
      continue;
    }

    switch (id) {
      case REC_EOF:
        break;
      case REC_BOF:
        depth++;
        continue;
      case REC_NUMBER: {
        if (length < 14) throw short('NUMBER');
        const value = view.getFloat64(d + 6, true);
        const style = styleOf(view.getUint16(d + 4, true));
        // Not a number a cell can hold: what Excel shows for such a result.
        put(view.getUint16(d, true), view.getUint16(d + 2, true), isFinite(value) ? { kind: 'number', value, style } : { kind: 'error', value: '#NUM!', style });
        continue;
      }
      case REC_RK:
        if (length < 10) throw short('RK');
        put(view.getUint16(d, true), view.getUint16(d + 2, true), { kind: 'number', value: decodeRk(view.getInt32(d + 6, true)), style: styleOf(view.getUint16(d + 4, true)) });
        continue;
      case REC_MULRK: {
        if (length < 6) throw short('MULRK');
        const row = view.getUint16(d, true);
        let col = view.getUint16(d + 2, true);
        for (let q = d + 4; q + 6 <= next - 2; q += 6, col++) {
          put(row, col, { kind: 'number', value: decodeRk(view.getInt32(q + 2, true)), style: styleOf(view.getUint16(q, true)) });
        }
        continue;
      }
      case REC_BLANK:
        if (length < 6) throw short('BLANK');
        put(view.getUint16(d, true), view.getUint16(d + 2, true), { kind: 'blank', value: null, style: styleOf(view.getUint16(d + 4, true)) });
        continue;
      case REC_MULBLANK: {
        if (length < 6) throw short('MULBLANK');
        const row = view.getUint16(d, true);
        let col = view.getUint16(d + 2, true);
        for (let q = d + 4; q + 2 <= next - 2; q += 2, col++) put(row, col, { kind: 'blank', value: null, style: styleOf(view.getUint16(q, true)) });
        continue;
      }
      case REC_LABELSST: {
        if (length < 10) throw short('LABELSST');
        const row = view.getUint16(d, true);
        const col = view.getUint16(d + 2, true);
        const index = view.getUint32(d + 6, true);
        const item = sharedStrings[index];
        if (index >= sharedStrings.length || !item) {
          throw corrupt(`shared string index ${index} in cell ${cellName(row, col)} of sheet ${name} is out of range (${sharedStrings.length} strings)`);
        }
        const cell: Cell = { kind: 'string', value: item.text, style: styleOf(view.getUint16(d + 4, true)) };
        if (item.runs) cell.runs = item.runs;
        cell.si = index;
        put(row, col, cell);
        continue;
      }
      case REC_LABEL:
      case REC_RSTRING: {
        if (length < 8) throw short(id === REC_LABEL ? 'LABEL' : 'RSTRING');
        const r = readerAt(stream, d - 4);
        const row = r.u16();
        const col = r.u16();
        const style = styleOf(r.u16());
        const raw: RawString = readText16(r, biff8, codepage);
        if (id === REC_RSTRING && !raw.runs && r.remaining > 0) {
          // The formatting runs follow the text: 16-bit pairs in BIFF8, 8-bit pairs before.
          const runs: number[] = [];
          const pairs = biff8 ? r.u16() : r.u8();
          for (let i = 0; i < pairs && r.remaining >= (biff8 ? 4 : 2); i++) {
            if (biff8) runs.push(r.u16(), r.u16());
            else runs.push(r.u8(), r.u8());
          }
          raw.runs = runs;
        }
        const item = toSharedString(raw, globals.fontOf);
        const cell: Cell = { kind: 'string', value: item.text, style };
        if (item.runs) cell.runs = item.runs;
        put(row, col, cell);
        continue;
      }
      case REC_BOOLERR: {
        if (length < 8) throw short('BOOLERR');
        const style = styleOf(view.getUint16(d + 4, true));
        const value = bytes[d + 6];
        put(view.getUint16(d, true), view.getUint16(d + 2, true), bytes[d + 7] ? { kind: 'error', value: errorText(value) || '#N/A', style } : { kind: 'boolean', value: value !== 0, style });
        continue;
      }
      case REC_FORMULA: {
        if (length < 22) throw short('FORMULA');
        const row = view.getUint16(d, true);
        const col = view.getUint16(d + 2, true);
        const style = styleOf(view.getUint16(d + 4, true));
        let cell: Cell;
        awaitingString = undefined;
        if (view.getUint16(d + 12, true) === 0xffff) {
          const type = bytes[d + 6];
          if (type === 1) cell = { kind: 'boolean', value: bytes[d + 8] !== 0, style };
          else if (type === 2) cell = { kind: 'error', value: errorText(bytes[d + 8]) || '#N/A', style };
          else {
            // The text is in a STRING record after this one, except for an empty text.
            cell = { kind: 'string', value: '', style };
            if (type === 0) awaitingString = cell;
          }
        } else {
          const value = view.getFloat64(d + 6, true);
          cell = isFinite(value) ? { kind: 'number', value, style } : { kind: 'error', value: '#NUM!', style };
        }
        put(row, col, cell);

        // Tokens that do not fit in the record go on in CONTINUE records; then they are read from a copy.
        let source = bytes;
        let tokens = d + 22;
        let end = next;
        const tokenSize = view.getUint16(d + 20, true);
        if (tokens + tokenSize > next || (next + 4 <= size && view.getUint16(next, true) === REC_CONTINUE)) {
          source = readerAt(stream, d - 4).rest();
          tokens = 22;
          end = source.length;
          if (tokens + tokenSize > end) throw corrupt(`the formula of cell ${cellName(row, col)} in sheet ${name} is cut off`);
        }
        const first = tokenSize > 0 ? source[tokens] : 0;
        lastFormulaAt = key(row, col);
        if ((first === 0x01 || first === 0x02) && tokenSize >= 5) {
          // The formula is the one of a group: a shared formula, an array formula or a data table.
          const sourceView = source === bytes ? view : new DataView(source.buffer, source.byteOffset, source.byteLength);
          const masterRow = sourceView.getUint16(tokens + 1, true);
          const masterCol = sourceView.getUint16(tokens + 3, true);
          lastLink = masterCol < XLS_MAX_COLS ? key(masterRow, masterCol) : -1;
          linkedCells.push(cell);
          linkedAt.push(lastFormulaAt);
          linkedTo.push(lastLink);
        } else {
          lastLink = -1;
          const text = formulaText(ctx, source, tokens, tokens + tokenSize, tokens + tokenSize, end, row, col);
          if (text === undefined) report.untranslatedFormulas++;
          else cell.formula = { text, type: 'normal' };
        }
        continue;
      }
      case REC_STRING:
        if (awaitingString) {
          awaitingString.value = readText16(readerAt(stream, d - 4), biff8, codepage).text;
          awaitingString = undefined;
        }
        continue;
      case REC_SHRFMLA: {
        if (length < 10) throw short('SHRFMLA');
        const tokenSize = view.getUint16(d + 8, true);
        if (d + 10 + tokenSize > next) throw short('SHRFMLA');
        const range = { top: view.getUint16(d, true), bottom: view.getUint16(d + 2, true), left: bytes[d + 4], right: bytes[d + 5] };
        // The record follows the first formula of its group, which points to the cell the group is known by.
        const group = lastLink >= 0 ? lastLink : key(range.top, range.left);
        if (!shared.has(group)) shared.set(group, { bytes, start: d + 10, end: d + 10 + tokenSize, extraEnd: next, range, si: nextShared++, placed: false });
        continue;
      }
      case REC_ARRAY: {
        if (length < 14) throw short('ARRAY');
        const tokenSize = view.getUint16(d + 12, true);
        const range = { top: view.getUint16(d, true), bottom: view.getUint16(d + 2, true), left: bytes[d + 4], right: bytes[d + 5] };
        let source = bytes;
        let tokens = d + 14;
        let end = next;
        if (tokens + tokenSize > next || (next + 4 <= size && view.getUint16(next, true) === REC_CONTINUE)) {
          source = readerAt(stream, d - 4).rest();
          tokens = 14;
          end = source.length;
          if (tokens + tokenSize > end) throw short('ARRAY');
        }
        const group = lastLink >= 0 ? lastLink : key(range.top, range.left);
        const text = range.top <= range.bottom && range.left <= range.right ? formulaText(ctx, source, tokens, tokens + tokenSize, tokens + tokenSize, end, range.top, range.left) : undefined;
        if (!ranged.has(group)) ranged.set(group, { range, formula: text === undefined ? undefined : { text, type: 'array', ref: rangeName(range) } });
        continue;
      }
      case REC_TABLE: {
        if (length < 16) throw short('TABLE');
        const range = { top: view.getUint16(d, true), bottom: view.getUint16(d + 2, true), left: bytes[d + 4], right: bytes[d + 5] };
        const flags = view.getUint16(d + 6, true);
        const group = lastLink >= 0 ? lastLink : key(range.top, range.left);
        let formula: CellFormula | undefined;
        const col1 = view.getUint16(d + 10, true);
        const col2 = view.getUint16(d + 14, true);
        // A data table whose input cells were deleted cannot be written as one.
        if (!(flags & 0x30) && range.top <= range.bottom && range.left <= range.right && col1 < XLS_MAX_COLS && (!(flags & 0x08) || col2 < XLS_MAX_COLS)) {
          const attrs = new Map<string, string>();
          if (flags & 0x08) attrs.set('dt2D', '1');
          if (flags & 0x04) attrs.set('dtr', '1');
          attrs.set('r1', cellName(view.getUint16(d + 8, true), col1));
          if (flags & 0x08) attrs.set('r2', cellName(view.getUint16(d + 12, true), col2));
          formula = { text: '', type: 'dataTable', ref: rangeName(range), attrs };
        }
        if (!ranged.has(group)) ranged.set(group, { range, formula });
        continue;
      }
      case REC_ROW: {
        if (length < 16) throw short('ROW');
        const row = rowOf(view.getUint16(d, true));
        const height = view.getUint16(d + 6, true);
        const flags = view.getUint16(d + 12, true);
        // The top bit says the height is the default one, which is then not a property of the row.
        if (!(height & 0x8000)) row.height = height / 20;
        row.outlineLevel = flags & 0x07;
        row.collapsed = (flags & 0x10) !== 0;
        row.hidden = (flags & 0x20) !== 0;
        row.customHeight = (flags & 0x40) !== 0;
        if (flags & 0x80) row.style = styleOf(view.getUint16(d + 14, true) & 0x0fff);
        continue;
      }
      case REC_COLINFO: {
        if (length < 10) throw short('COLINFO');
        // A valid sheet cannot state more ranges than it has columns; this bounds the work on a hostile file.
        if (++columnRecords > XLS_MAX_COLS) continue;
        const min = view.getUint16(d, true);
        const max = Math.min(view.getUint16(d + 2, true), XLS_MAX_COLS - 1);
        if (min > max) continue;
        const flags = view.getUint16(d + 8, true);
        const range: ColumnRange = {
          min,
          max,
          width: view.getUint16(d + 4, true) / 256,
          customWidth: (flags & 0x02) !== 0,
          hidden: (flags & 0x01) !== 0,
          outlineLevel: (flags >> 8) & 0x07,
          collapsed: (flags & 0x1000) !== 0,
        };
        const style = styleOf(view.getUint16(d + 6, true));
        if (style !== 0) range.style = style;
        addColumnRange(sheet.columns, range);
        continue;
      }
      case REC_DEFCOLWIDTH:
        if (length >= 2) {
          defaultWidthChars = view.getUint16(d, true);
          sheet.baseColWidth = defaultWidthChars;
        }
        continue;
      case REC_STANDARDWIDTH:
        if (length >= 2) {
          sheet.defaultColWidth = view.getUint16(d, true) / 256;
          standardWidth = true;
        }
        continue;
      case REC_DEFAULTROWHEIGHT:
        if (length >= 4) sheet.defaultRowHeight = (view.getUint16(d + 2, true) & 0x7fff) / 20;
        continue;
      case REC_DIMENSION: {
        let top: number;
        let bottom: number;
        let left: number;
        let right: number;
        if (biff8) {
          if (length < 12) continue;
          top = view.getUint32(d, true);
          bottom = view.getUint32(d + 4, true);
          left = view.getUint16(d + 8, true);
          right = view.getUint16(d + 10, true);
        } else {
          if (length < 8) continue;
          top = view.getUint16(d, true);
          bottom = view.getUint16(d + 2, true);
          left = view.getUint16(d + 4, true);
          right = view.getUint16(d + 6, true);
        }
        // The last row and column are stored one past the end; an empty sheet has none.
        if (bottom > top && right > left && bottom <= XLS_MAX_ROWS && right <= XLS_MAX_COLS) sheet.dimension = { top, left, bottom: bottom - 1, right: right - 1 };
        continue;
      }
      case REC_MERGECELLS: {
        if (length < 2) continue;
        const entries = view.getUint16(d, true);
        for (let i = 0, q = d + 2; i < entries && q + 8 <= next; i++, q += 8) {
          const range = { top: view.getUint16(q, true), bottom: view.getUint16(q + 2, true), left: view.getUint16(q + 4, true), right: view.getUint16(q + 6, true) };
          if (range.top <= range.bottom && range.left <= range.right && range.right < XLS_MAX_COLS) sheet.merges.push(range);
        }
        continue;
      }
      case REC_HLINK:
        // Links come in more layouts than this reader knows; one it cannot read is left out, not the whole file.
        try {
          const link = readHyperlink(readerAt(stream, d - 4));
          if (link) sheet.hyperlinks.push(link);
        } catch {
          count(report.skippedRecords, id);
        }
        continue;
      case REC_HLINKTOOLTIP:
        if (length >= 10) tooltips.push(readHyperlinkTooltip(readerAt(stream, d - 4)));
        continue;
      case REC_WINDOW2: {
        if (length < 2) continue;
        const flags = view.getUint16(d, true);
        const viewOf = sheet.view;
        viewOf.showGridLines = (flags & 0x0002) !== 0;
        viewOf.showRowColHeaders = (flags & 0x0004) !== 0;
        frozen = (flags & 0x0008) !== 0;
        viewOf.showZeros = (flags & 0x0010) !== 0;
        viewOf.rightToLeft = (flags & 0x0040) !== 0;
        viewOf.tabSelected = (flags & 0x0200) !== 0;
        if (length >= 6) {
          const top = view.getUint16(d + 2, true);
          const left = view.getUint16(d + 4, true);
          if ((top > 0 || left > 0) && left < XLS_MAX_COLS) viewOf.firstVisibleCell = cellName(top, left);
        }
        continue;
      }
      case REC_PANE: {
        if (length < 9) continue;
        activePane = bytes[d + 8];
        if (frozen) {
          sheet.view.frozenCols = Math.min(view.getUint16(d, true), XLS_MAX_COLS - 1);
          sheet.view.frozenRows = Math.min(view.getUint16(d + 2, true), XLS_MAX_ROWS - 1);
          const left = view.getUint16(d + 6, true);
          if (left < XLS_MAX_COLS) sheet.view.topLeftCell = cellName(view.getUint16(d + 4, true), left);
        }
        continue;
      }
      case REC_SELECTION: {
        if (length < 5) continue;
        const col = view.getUint16(d + 3, true);
        if (col >= XLS_MAX_COLS) continue;
        const active = cellName(view.getUint16(d + 1, true), col);
        if (firstActive === undefined) firstActive = active;
        // The active cell is the one of the selection that belongs to the active pane.
        if (paneActive === undefined && bytes[d] === activePane) paneActive = active;
        continue;
      }
      case REC_SCL: {
        if (length < 4) continue;
        const denominator = view.getUint16(d + 2, true);
        const zoom = denominator > 0 ? Math.round((view.getUint16(d, true) * 100) / denominator) : 0;
        if (zoom >= 10 && zoom <= 400) sheet.view.zoomScale = zoom;
        continue;
      }
      case REC_WSBOOL:
        if (length >= 2 && view.getUint16(d, true) & 0x0010 && sheet.kind === 'worksheet') sheet.kind = 'dialogsheet';
        continue;
      case REC_SHEETEXT:
        if (length >= 20) {
          const color = view.getUint32(d + 16, true) & 0x7f;
          if (color >= 8 && color < 64) sheet.tabColor = { indexed: color };
        }
        continue;
      case REC_NOTE:
        report.comments++;
        continue;
      case REC_CONTINUE:
      case REC_CONTINUEFRT12:
        continue;
      default:
        count(report.skippedRecords, id);
        continue;
    }
    break;
  }

  // Formulas that belong to a group.
  for (let i = 0; i < linkedCells.length; i++) {
    const cell = linkedCells[i];
    const at = linkedAt[i];
    const row = Math.floor(at / XLS_MAX_COLS);
    const col = at % XLS_MAX_COLS;
    const group = shared.get(linkedTo[i]);
    if (group) {
      const text = formulaText(ctx, group.bytes, group.start, group.end, group.end, group.extraEnd, row, col);
      if (text === undefined) report.untranslatedFormulas++;
      else {
        // Every cell carries its own text, so nothing depends on the first cell of the group staying as it is.
        cell.formula = { text, type: 'shared', si: group.si };
        if (!group.placed) {
          group.placed = true;
          if (group.range.top <= group.range.bottom && group.range.left <= group.range.right) cell.formula.ref = rangeName(group.range);
        }
      }
      continue;
    }
    const range = ranged.get(linkedTo[i]);
    if (range) {
      // Only the first cell of an array formula or a data table carries it; the others just show their values.
      if (row === range.range.top && col === range.range.left) {
        if (range.formula) cell.formula = range.formula;
        else report.untranslatedFormulas++;
      }
      continue;
    }
    report.untranslatedFormulas++;
    countReason(report.untranslatedReasons, 'PtgExp without its group');
  }

  const active = paneActive !== undefined ? paneActive : firstActive;
  if (active !== undefined) sheet.view.activeCell = active;
  for (const tip of tooltips) {
    for (const link of sheet.hyperlinks) {
      const range = link.range;
      if (range.top === tip.top && range.left === tip.left && range.bottom === tip.bottom && range.right === tip.right && tip.text !== '') link.tooltip = tip.text;
    }
  }
  // The default width is stated in characters of the default font; a file stores widths with the padding
  // of 5 pixels, rounded up to a multiple of 8 pixels (8 characters are 64 pixels at 7 pixels a digit).
  if (!standardWidth && defaultWidthChars > 0 && defaultWidthChars !== 8 && defaultWidthChars <= 255) {
    sheet.defaultColWidth = Math.floor(((Math.ceil((defaultWidthChars * 7 + 5) / 8) * 8) / 7) * 256) / 256;
  }
}

function countReason(map: Map<string, number>, reason: string): void {
  map.set(reason, (map.get(reason) || 0) + 1);
}

const SHEET_KINDS: Readonly<Record<number, SheetKind>> = { 0: 'worksheet', 1: 'macrosheet', 2: 'chartsheet', 6: 'macrosheet' };

/** Reads the bytes of a Workbook (BIFF8) or Book (BIFF5) stream. `streamName` is the name the stream has in the file. */
export function readXlsStream(bytes: Uint8Array, streamName = 'Workbook'): XlsReadResult {
  const stream = recordStream(bytes);
  try {
    const globals = readGlobals(stream, streamName);
    if (globals.sheets.length === 0) throw corrupt('the workbook has no sheets');

    const sheets: Sheet[] = [];
    const seen = new Set<number>();
    globals.sheets.forEach((bound, index) => {
      const sheet = createSheet({ name: bound.name, sheetId: index + 1, state: bound.state, kind: SHEET_KINDS[bound.type] || 'worksheet', part: '' });
      sheets.push(sheet);
      // Each sheet has its own substream; one listed twice would be read into memory twice.
      if (seen.has(bound.offset)) throw corrupt(`sheet ${quote(bound.name)} shares its records with another sheet`);
      seen.add(bound.offset);
      if (bound.offset + 4 > bytes.length || stream.view.getUint16(bound.offset, true) !== REC_BOF || recordLength(stream, bound.offset) < 4) {
        throw corrupt(`sheet ${quote(bound.name)} does not start with a BOF record`);
      }
      const type = stream.view.getUint16(bound.offset + 6, true);
      if (type === 0x0020) sheet.kind = 'chartsheet';
      else if (type === 0x0040 || type === 0x0006) sheet.kind = 'macrosheet';
      // Charts are not cells, and macro sheets are never run: both stay empty.
      if (sheet.kind === 'worksheet') readSheet(stream, bound.offset, sheet, globals);
    });

    const workbook: Workbook = {
      sheets,
      styles: globals.styles,
      sharedStrings: globals.sharedStrings,
      definedNames: definedNamesOf(globals),
      date1904: globals.date1904,
      activeSheet: globals.activeTab < sheets.length ? globals.activeTab : 0,
    };
    return { workbook, report: globals.report };
  } finally {
    releaseFormulaBytes();
  }
}

/** Reads an xls file, and says what was left out. */
export function readXlsWithReport(cfb: CfbFile): XlsReadResult {
  // A file saved for both Excel 97 and Excel 95 has both streams; the newer one says more.
  const name = cfb.has('Workbook') ? 'Workbook' : cfb.has('Book') ? 'Book' : undefined;
  if (name === undefined) throw new XlsUnsupportedError('The compound file has no workbook stream');
  return readXlsStream(cfb.stream(name), name);
}

/** Reads an xls file into the spreadsheet model. */
export function readXls(cfb: CfbFile): Workbook {
  return readXlsWithReport(cfb).workbook;
}
