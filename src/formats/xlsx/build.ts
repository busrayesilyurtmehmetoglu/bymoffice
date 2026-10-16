// Builds a new xlsx package from a spreadsheet model alone: the way to save a
// workbook read from an xls file (or made in code) as xlsx.
//
// writeXlsx() brings a package up to date with its model and copies whatever
// the model does not hold from the original parts. Here there are no original
// parts: every part is written from the model, and what the model does not
// hold is not in the result. Cells, rows and the pieces of a sheet are written
// by the functions writeXlsx() uses for a sheet it regenerates.
//
// What the model holds and a new package cannot carry is never left out
// silently: it is noted in the report (see XlsxBuildReport).
//   - A chart sheet or a macro sheet has no content in the model; it is written
//     as an empty worksheet so that the names and positions of the sheets stay.
//   - The placeholder of a chart, a diagram or an embedded object is left out.
//   - A picture whose bytes cannot be had (a metafile of an xls file) is written
//     as a frame without a picture.
//   - Threaded comments are written as notes, without their replies.
//   - Of conditional formats, what only the 2010 extension can say is left out.
//   - Of the page setup, a value a file cannot state is left out.

import { adjustReferenceText, shiftFormula, type StructuralEdit } from '../../calc/refs';
import { createSheetMetrics, type SheetMetrics } from '../../layout/sheet-metrics';
import {
  MAX_COLS,
  MAX_ROWS,
  cellName,
  emptyStyleSheet,
  fileSheetName,
  isFileSheet,
  parseCellRef,
  parseRange,
  rangeName,
  sheetNameMap,
  type AutoFilter,
  type Cell,
  type CellFormat,
  type CellRange,
  type Color,
  type ColumnRange,
  type ConditionalRule,
  type ConditionalValue,
  type DataValidation,
  type DrawingAnchor,
  type DrawingMarker,
  type DrawingObject,
  type Fill,
  type PageBreak,
  type RichRun,
  type Row,
  type SharedString,
  type Sheet,
  type SheetComment,
  type SheetNameMap,
  type SheetPageSetup,
  type SheetProtection,
  type SheetStructuralEdit,
  type StyleSheet,
  type Workbook,
} from '../../model/sheet/types';
import { KEPT_MAIN, MAX_KEPT_DEPTH, unpackKept, type KeptAttr, type KeptNode, type RuleKept } from '../../model/sheet/types';
import { keptFor, keptNow, ruleHolder } from '../../ops/rules';
import { adjustCommentCell, adjustDrawingAnchor, commentEdits, drawingAnchorNow } from '../../render/sheet-drawings';
import { CHART_CONTENT_TYPE, chartCellsOf, chartSpaceXml, chartWritable } from './chart-write';
import { shapeElementXml, writtenAsConnector } from './shape-write';
import { pictureSource } from '../../model/sheet/pictures';
import { METADATA_CONTENT_TYPE, METADATA_RELATIONSHIP, newMetadataXml } from './metadata';
import { serializeXml, type XmlElement, type XmlNode } from '../../xml';
import { ZipArchive } from '../../zip';
import { readModel } from '../../model/edits';
import {
  STYLESHEET_ORDER,
  adjustSheetElement,
  planDxfs,
  autoFilterXml,
  borderElement,
  colorAttrs,
  colsXml,
  encodeXstring,
  escapeAttr as attrEscape,
  escapeText as textEscape,
  escapeXstring,
  fillElement,
  fontElement,
  formatElement,
  mergeCellsXml,
  replaceUnwritable,
  VML_BLOCK,
  defaultNoteAnchor,
  noteShapeTypeXml,
  noteShapeXml,
  rowXml,
  runPropertiesXml,
  textElement,
  type CellWriter,
} from './write';

const encoder = new TextEncoder();

/** In place of a name for a sheet that is not in the file: no formula of the file names it. */
const NOT_IN_FILE = '\u0000';

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_SHEET_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const TYPE_PREFIX = 'application/vnd.openxmlformats-officedocument.';

const EMU_PER_PIXEL = 9525;
/** A sheet is at most this wide or tall, in EMU; the reader takes no offset beyond it. */
const MAX_EMU = 2e12;
/** Groups inside groups that are written; deeper members are left out, as the reader leaves them out. */
const MAX_GROUP_DEPTH = 16;

/** What a build did; filled in when passed with the options. */
export interface XlsxBuildReport {
  /** Parts of the new package, in the order they were written. */
  parts: string[];
  /**
   * What the model holds and the file does not, or holds differently: one line per kind
   * of thing, with the number of times it happened. Empty when everything was written.
   */
  warnings: string[];
}

export interface BuildXlsxOptions {
  report?: XlsxBuildReport;
}

// ------------------------------------------------------------------ helpers

/**
 * Characters XML cannot carry (control characters, halves of surrogate pairs) that were
 * written as U+FFFD in an attribute or in text that has no escape of its own (a formula,
 * an error value). Counted
 * from the start of a build to its end, which is one run without a pause.
 */
let replacedCharacters = 0;

function counted(text: string): string {
  const clean = replaceUnwritable(text);
  if (clean !== text) replacedCharacters++;
  return clean;
}

const escapeAttr = (text: string) => attrEscape(counted(text));
const escapeText = (text: string) => textEscape(counted(text));

/** Counts what could not be written, by message, in the order it first happened. */
class Notes {
  private readonly counts = new Map<string, number>();

  add(message: string, count = 1): void {
    if (count > 0) this.counts.set(message, (this.counts.get(message) || 0) + count);
  }

  list(): string[] {
    const out: string[] = [];
    for (const [message, count] of this.counts) out.push(count === 1 ? message : `${message} (${count} times)`);
    return out;
  }
}

/** The relationships of one part of the new package. */
class Rels {
  private readonly entries: string[] = [];

  add(type: string, target: string, external = false): string {
    const id = `rId${this.entries.length + 1}`;
    this.entries.push(`<Relationship Id="${id}" Type="${escapeAttr(type)}" Target="${escapeAttr(target)}"${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }

  get size(): number {
    return this.entries.length;
  }

  text(): string {
    return `${XML_DECLARATION}<Relationships xmlns="${NS_PACKAGE_RELS}">${this.entries.join('')}</Relationships>`;
  }
}

const relsNameOf = (name: string) => `${name.slice(0, name.lastIndexOf('/') + 1)}_rels/${name.slice(name.lastIndexOf('/') + 1)}.rels`;

const quote = (text: string) => JSON.stringify(text.length > 40 ? `${text.slice(0, 40)}...` : text);

const isIndex = (value: unknown, limit: number): value is number => typeof value === 'number' && value >= 0 && value < limit && Math.floor(value) === value;

const validRange = (range: CellRange | undefined): range is CellRange =>
  !!range && isIndex(range.top, MAX_ROWS) && isIndex(range.bottom, MAX_ROWS) && isIndex(range.left, MAX_COLS) && isIndex(range.right, MAX_COLS) && range.top <= range.bottom && range.left <= range.right;

/** A name with a namespace prefix, other than the "xml:" every document knows. */
const isPrefixed = (name: string) => name.indexOf(':') >= 0 && name.slice(0, 4) !== 'xml:';

/** Source text of an element kept by the model that names no namespace prefix: it means the same in a new part. */
const isPortableText = (raw: string) => !/<\/?[A-Za-z_][\w.-]*:|\s[A-Za-z_][\w.-]*:[\w.-]+\s*=/.test(raw.replace(/\sxml:[\w.-]+\s*=/g, ' '));

/**
 * A copy of an element kept by the model as it was read, for a part whose
 * elements have no prefix: the prefix of the element is taken off its names.
 * Undefined when it holds names of another namespace, whose declaration is
 * not in the model.
 */
function portable(el: XmlElement, depth = 0): XmlElement | undefined {
  if (depth > 64) return undefined;
  const colon = el.name.indexOf(':');
  const prefix = colon < 0 ? '' : el.name.slice(0, colon + 1);
  const walk = (node: XmlElement, level: number): XmlElement | undefined => {
    if (level > 64 || node.name.slice(0, prefix.length) !== prefix) return undefined;
    const name = node.name.slice(prefix.length);
    if (name.indexOf(':') >= 0) return undefined;
    const attrs = new Map<string, string>();
    for (const [key, value] of node.attrs) {
      if (key === 'xmlns' || key.slice(0, 6) === 'xmlns:') continue;
      if (isPrefixed(key)) return undefined;
      attrs.set(key, value);
    }
    const children: XmlNode[] = [];
    for (const child of node.children) {
      if (child.type === 'element') {
        const copy = walk(child, level + 1);
        if (!copy) return undefined;
        children.push(copy);
      } else if (child.type === 'text') children.push({ ...child });
    }
    return { type: 'element', name, attrs, children };
  };
  return walk(el, depth);
}

/** Makes every attribute and text of a tree something XML can carry. */
function sanitize(el: XmlElement): XmlElement {
  for (const [name, value] of el.attrs) {
    const clean = counted(value);
    if (clean !== value) el.attrs.set(name, clean);
  }
  for (const child of el.children) {
    if (child.type === 'element') sanitize(child);
    else if (child.type === 'text') child.value = counted(child.value);
  }
  return el;
}

/**
 * A reference into another workbook ("[1]Sheet1!A1", "'[2]Other sheet'!A1") in a formula. What stands
 * in a text of the formula is none, nor is anything inside the brackets of a table reference
 * ("Table1[[#This Row],[1]]") or right after a name ("Table1[1]").
 */
function refersToWorkbook(formula: string): boolean {
  if (formula.indexOf('[') < 0) return false;
  let depth = 0;
  for (let i = 0; i < formula.length; i++) {
    const c = formula.charCodeAt(i);
    if (c === 34 && depth === 0) {
      // A text: up to the quote that closes it ("" inside it is a quote of the text).
      i = formula.indexOf('"', i + 1);
      while (i >= 0 && formula.charCodeAt(i + 1) === 34) i = formula.indexOf('"', i + 2);
      if (i < 0) return false;
    } else if (c === 91) {
      if (depth === 0 && !/[A-Za-z0-9_.\]]/.test(formula.charAt(i - 1)) && /^\[\d+\]/.test(formula.slice(i, i + 12))) return true;
      depth++;
    } else if (c === 93 && depth > 0) depth--;
  }
  return false;
}

/** "B2" for "B2", "$B$2" and "B2:C3": the one cell an attribute that takes a cell is given. */
function oneCell(ref: string | undefined): string | undefined {
  const range = typeof ref === 'string' ? parseRange(ref) : undefined;
  return range ? cellName(range.top, range.left) : undefined;
}

const element = (name: string, attrs: [string, string][] = [], children: XmlNode[] = []): XmlElement => ({ type: 'element', name, attrs: new Map(attrs), children });

// ------------------------------------------------------------------- styles

const isGray125 = (fill: Fill | undefined) => !!fill && fill.type === 'pattern' && fill.pattern === 'gray125';

/** xl/styles.xml from the model; the tables are in the order CT_Stylesheet gives them. */
function stylesXml(model: StyleSheet, notes: Notes, addedDxfs: XmlElement[] = []): string {
  const defaults = emptyStyleSheet();
  const fonts = model.fonts.length > 0 ? model.fonts : defaults.fonts;
  const borders = model.borders.length > 0 ? model.borders : defaults.borders;
  const styleXfs = model.cellStyleXfs.length > 0 ? model.cellStyleXfs : defaults.cellStyleXfs;
  const cellXfs = model.cellXfs.length > 0 ? model.cellXfs : defaults.cellXfs;

  // The first two fills are fixed in every application ("none" and "gray125"), whatever the file
  // says they are: a model whose second fill is something else gets the fixed one in between.
  // A model whose first fill is something else gets "none" before it as well; the indexes of the
  // formats follow, so that every format has the fill it had.
  const fills = model.fills.length > 0 ? model.fills.slice() : defaults.fills.slice();
  const lead = fills[0].type === 'pattern' && fills[0].pattern === 'none' ? 0 : 1;
  if (lead) fills.unshift({ type: 'pattern', pattern: 'none' });
  const gap = !isGray125(fills[1]);
  if (gap) fills.splice(1, 0, { type: 'pattern', pattern: 'gray125' });
  const fillId = (id: number) => (gap && id + lead >= 1 ? id + lead + 1 : id + lead);

  const table = (name: string, children: XmlNode[]) => element(name, [['count', String(children.length)]], children);
  const format = (xf: CellFormat, cell: boolean): XmlElement => {
    if (!isIndex(xf.fontId, fonts.length) || !isIndex(fillId(xf.fillId), fills.length) || !isIndex(xf.borderId, borders.length) || (cell && xf.xfId !== undefined && !isIndex(xf.xfId, styleXfs.length))) {
      notes.add('Styles: a format refers to a font, a fill, a border or a parent the tables do not have; the first one is used');
    }
    const fixed: CellFormat = {
      ...xf,
      numFmtId: isIndex(xf.numFmtId, 0x10000) ? xf.numFmtId : 0,
      fontId: isIndex(xf.fontId, fonts.length) ? xf.fontId : 0,
      fillId: isIndex(fillId(xf.fillId), fills.length) ? fillId(xf.fillId) : 0,
      borderId: isIndex(xf.borderId, borders.length) ? xf.borderId : 0,
    };
    if (cell) fixed.xfId = isIndex(xf.xfId, styleXfs.length) ? xf.xfId : 0;
    else delete fixed.xfId;
    return formatElement(fixed, '');
  };

  const children: XmlNode[] = [];
  const numFmts: XmlNode[] = [];
  for (const [id, code] of model.numFmts) {
    if (isIndex(id, 0x10000) && typeof code === 'string') numFmts.push(element('numFmt', [['numFmtId', String(id)], ['formatCode', code]]));
  }
  if (numFmts.length > 0) children.push(table('numFmts', numFmts));
  children.push(table('fonts', fonts.map((font) => fontElement(font, ''))));
  children.push(
    table(
      'fills',
      fills.map((fill) => {
        if (fill.type !== 'gradient') return fillElement(fill, '');
        const raw = portable(fill.raw);
        if (raw) return element('fill', [], [raw]);
        notes.add('Styles: a gradient fill uses a namespace the model does not hold and is written as no fill');
        return fillElement({ type: 'pattern', pattern: 'none' }, '');
      }),
    ),
  );
  children.push(table('borders', borders.map((border) => borderElement(border, ''))));
  children.push(table('cellStyleXfs', styleXfs.map((xf) => format(xf, false))));
  children.push(table('cellXfs', cellXfs.map((xf) => format(xf, true))));
  const cellStyles = model.cellStyles.length > 0 ? model.cellStyles : [{ name: 'Normal', xfId: 0, builtinId: 0 }];
  children.push(
    table(
      'cellStyles',
      cellStyles.map((style) => {
        const attrs: [string, string][] = [['name', style.name], ['xfId', String(isIndex(style.xfId, styleXfs.length) ? style.xfId : 0)]];
        if (style.builtinId !== undefined && isIndex(style.builtinId, 0x10000)) attrs.push(['builtinId', String(style.builtinId)]);
        return element('cellStyle', attrs);
      }),
    ),
  );
  if (model.dxfs.length > 0 || addedDxfs.length > 0) {
    children.push(
      table(
        'dxfs',
        // The records of the model, then those the rules that state their own format ask for (see planDxfs()).
        model.dxfs.concat(addedDxfs).map((dxf) => {
          const copy = portable(dxf);
          if (copy) return copy;
          // The record keeps its place: conditional formats refer to it by position.
          notes.add('Styles: a differential format uses a namespace the model does not hold and is written empty');
          return element('dxf');
        }),
      ),
    );
  }
  if (model.indexedColors && model.indexedColors.length > 0) {
    const palette = model.indexedColors.map((rgb) => element('rgbColor', [['rgb', `00${/^[0-9A-Fa-f]{6}$/.test(rgb) ? rgb.toUpperCase() : '000000'}`]]));
    children.push(element('colors', [], [element('indexedColors', [], palette)]));
  }
  children.sort((a, b) => STYLESHEET_ORDER.indexOf((a as XmlElement).name) - STYLESHEET_ORDER.indexOf((b as XmlElement).name));
  const root = sanitize(element('styleSheet', [['xmlns', NS_MAIN]], children));
  return XML_DECLARATION + serializeXml({ declaration: null, leading: [], root, trailing: [] });
}

/** Colours of the Office theme, in the order of StyleSheet.themeColors; they stand in for entries the model lacks. */
const OFFICE_COLORS = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];
/** The slots of a colour scheme in the order CT_ColorScheme gives them, with their index in StyleSheet.themeColors. */
const SCHEME_SLOTS: [name: string, index: number][] = [
  ['dk1', 1],
  ['lt1', 0],
  ['dk2', 3],
  ['lt2', 2],
  ['accent1', 4],
  ['accent2', 5],
  ['accent3', 6],
  ['accent4', 7],
  ['accent5', 8],
  ['accent6', 9],
  ['hlink', 10],
  ['folHlink', 11],
];

/**
 * A theme part with the colours of the model. Cell colours given as theme
 * indexes mean nothing without one. The model holds no more of a theme than
 * its colours; the rest is the least a theme has to have.
 */
function themeXml(styles: StyleSheet, notes: Notes): string {
  let colors = '';
  for (const [name, index] of SCHEME_SLOTS) {
    let rgb = styles.themeColors[index];
    if (typeof rgb !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(rgb)) {
      notes.add('Theme: a colour the model has no value for is written as the one of the Office theme');
      rgb = OFFICE_COLORS[index];
    }
    colors += `<a:${name}><a:srgbClr val="${rgb.toUpperCase()}"/></a:${name}>`;
  }
  const face = escapeAttr((styles.fonts[0] && styles.fonts[0].name) || 'Calibri');
  const font = `<a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (width: number) => `<a:ln w="${width}">${fill}</a:ln>`;
  const effect = '<a:effectStyle><a:effectLst/></a:effectStyle>';
  return (
    `${XML_DECLARATION}<a:theme xmlns:a="${NS_DRAWING}" name="Office Theme"><a:themeElements>` +
    `<a:clrScheme name="Office">${colors}</a:clrScheme>` +
    `<a:fontScheme name="Office"><a:majorFont>${font}</a:majorFont><a:minorFont>${font}</a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
    `<a:lnStyleLst>${line(9525)}${line(25400)}${line(38100)}</a:lnStyleLst>` +
    `<a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst>` +
    `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme>` +
    '</a:themeElements></a:theme>'
  );
}

// ----------------------------------------------------------- shared strings

/** The text of a string item or a comment: plain, or in runs when the runs spell the text. */
function richTextXml(text: string, runs: RichRun[] | undefined): string {
  if (runs && runs.length > 0) {
    let joined = '';
    for (const run of runs) joined += run.text;
    if (joined === text) {
      let out = '';
      for (const run of runs) out += `<r>${run.font ? runPropertiesXml(run.font) : ''}${textElement(run.text, '')}</r>`;
      return out;
    }
  }
  return textElement(text, '');
}

/** The shared strings of the new package: those of the model at their indexes, then the texts of cells that referred to none. */
class StringTable {
  readonly items: SharedString[];
  private readonly added = new Map<string, number>();
  /** Cells that refer to a string. */
  uses = 0;

  constructor(shared: SharedString[]) {
    // A hole in the list still takes its index.
    this.items = Array.from(shared, (item) => item || { text: '' });
  }

  intern(text: string, runs: RichRun[] | undefined): number {
    this.uses++;
    let rich = false;
    if (runs && runs.length > 0) {
      let joined = '';
      for (const run of runs) joined += run.text;
      rich = joined === text;
    }
    const key = rich ? `r${JSON.stringify(runs)}` : `t${text}`;
    let index = this.added.get(key);
    if (index === undefined) {
      index = this.items.length;
      this.items.push(rich ? { text, runs } : { text });
      this.added.set(key, index);
    }
    return index;
  }

  xml(): string {
    let out = `${XML_DECLARATION}<sst xmlns="${NS_MAIN}" count="${Math.max(this.uses, this.items.length)}" uniqueCount="${this.items.length}">`;
    for (const item of this.items) out += `<si>${richTextXml(typeof item.text === 'string' ? item.text : '', item.runs)}</si>`;
    return `${out}</sst>`;
  }
}

// ---------------------------------------------------------------- worksheet

interface SharedMaster {
  row: number;
  col: number;
  text: string;
  range: CellRange | undefined;
}

/** The cell that holds the text of each shared formula of a sheet: the first one in the order of the file. */
function sharedMastersOf(sheet: Sheet, rows: number[]): Map<number, SharedMaster> | undefined {
  let masters: Map<number, SharedMaster> | undefined;
  for (const index of rows) {
    const cells = (sheet.rows.get(index) as Row).cells;
    (cells.length > MAX_COLS ? cells.slice(0, MAX_COLS) : cells).forEach((cell, col) => {
      const formula = cell.formula;
      if (!formula || formula.type !== 'shared' || formula.si === undefined || formula.text === '') return;
      if (!masters) masters = new Map();
      if (masters.has(formula.si)) return;
      const range = formula.ref === undefined ? undefined : parseRange(formula.ref);
      masters.set(formula.si, { row: index, col, text: formula.text, range });
    });
  }
  return masters;
}

/**
 * The cell as a new package can hold it; the cell itself when it is fine as it is.
 *   - A shared formula is only kept where a file can state it: the first cell of the group gives
 *     the text and the range, the others come after it and lie inside that range. Any other cell
 *     of the group gets the text written out; one whose group has no text keeps its value only.
 *   - Metadata indexes point into a part the model does not hold.
 */
function writableCell(cell: Cell, row: number, col: number, masters: Map<number, SharedMaster> | undefined, notes: Notes): Cell {
  let out = cell;
  const formula = cell.formula;
  if (formula && formula.type === 'shared') {
    const master = formula.si === undefined || !masters ? undefined : masters.get(formula.si);
    const text = formula.text;
    if (!master) {
      if (text === '') {
        out = { ...cell };
        delete out.formula;
        notes.add('Formula: a cell of a shared formula whose first cell is missing keeps its value and has no formula');
      } else out = { ...cell, formula: { text, type: 'normal' } };
    } else if (master.row === row && master.col === col) {
      const range = master.range;
      if (!range || row < range.top || row > range.bottom || col < range.left || col > range.right) out = { ...cell, formula: { text, type: 'normal' } };
    } else {
      const range = master.range;
      const inside = !!range && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
      const after = row > master.row || (row === master.row && col > master.col);
      // A second cell with a text would start a group of the same number: it stands for itself.
      if (text !== '' || !inside || !after) out = { ...cell, formula: { text: text !== '' ? text : shiftFormula(master.text, row - master.row, col - master.col), type: 'normal' } };
    }
  }
  if (out.formula && out.formula.type !== 'shared') {
    const stated = out.formula;
    if (stated.text === '' && stated.type !== 'dataTable') {
      // An element without a text is no formula.
      out = { ...out };
      delete out.formula;
      notes.add('Formula: a formula without a text is left out; the cell keeps its value');
    } else if (stated.type !== 'normal' && (stated.ref === undefined || !parseRange(stated.ref))) {
      out = { ...out, formula: { ...stated, ref: cellName(row, col) } };
      notes.add('Formula: an array formula or a data table whose range is not one is given its own cell as the range');
    }
  }
  if (out.formula) {
    // Written by the writer of sheets, which replaces what XML cannot carry; here it is counted.
    counted(out.formula.text);
    if (out.formula.ref !== undefined) counted(out.formula.ref);
    if (out.formula.attrs) for (const value of out.formula.attrs.values()) counted(value);
    if (out.formula.text !== '' && refersToWorkbook(out.formula.text)) notes.add('Formula: a formula refers to another workbook; the link to that workbook is not in the model and is not written');
  }
  if (out.formula && out.formula.attrs) {
    let foreign = false;
    for (const name of out.formula.attrs.keys()) if (isPrefixed(name)) foreign = true;
    if (foreign) {
      const attrs = new Map<string, string>();
      for (const [name, value] of out.formula.attrs) {
        if (!isPrefixed(name)) attrs.set(name, value);
        else notes.add('Formula: an attribute of another namespace is left out');
      }
      out = { ...out, formula: { ...out.formula, attrs } };
    }
  }
  if (out.cm !== undefined || out.vm !== undefined) {
    out = { ...out };
    delete out.cm;
    delete out.vm;
    notes.add('Cell: metadata (a dynamic array or a rich value) is not in the model and its index is left out');
  }
  return out;
}

/** `selected`: the sheet is the one the file opens on, which is the one sheet marked as selected. */
function sheetViewXml(sheet: Sheet, selected: boolean): string {
  const view = sheet.view;
  let out = '<sheetView';
  if (!view.showGridLines) out += ' showGridLines="0"';
  if (!view.showRowColHeaders) out += ' showRowColHeaders="0"';
  if (!view.showZeros) out += ' showZeros="0"';
  if (view.rightToLeft) out += ' rightToLeft="1"';
  if (selected) out += ' tabSelected="1"';
  const firstVisible = oneCell(view.firstVisibleCell);
  if (firstVisible !== undefined) out += ` topLeftCell="${firstVisible}"`;
  if (view.zoomScale !== 100 && view.zoomScale >= 10 && view.zoomScale <= 400) out += ` zoomScale="${Math.round(view.zoomScale)}"`;
  out += ' workbookViewId="0"';

  let inner = '';
  const frozenCols = isIndex(view.frozenCols, MAX_COLS) ? view.frozenCols : 0;
  const frozenRows = isIndex(view.frozenRows, MAX_ROWS) ? view.frozenRows : 0;
  let pane: string | undefined;
  if (frozenCols > 0 || frozenRows > 0) {
    pane = frozenCols > 0 && frozenRows > 0 ? 'bottomRight' : frozenCols > 0 ? 'topRight' : 'bottomLeft';
    let topLeft = oneCell(view.topLeftCell);
    if (topLeft === undefined) {
      // The first cell that scrolls: right below and beside the frozen ones, which start where the view starts.
      const first = view.firstVisibleCell === undefined ? undefined : parseRange(view.firstVisibleCell);
      topLeft = cellName(Math.min((first ? first.top : 0) + frozenRows, MAX_ROWS - 1), Math.min((first ? first.left : 0) + frozenCols, MAX_COLS - 1));
    }
    inner += '<pane';
    if (frozenCols > 0) inner += ` xSplit="${frozenCols}"`;
    if (frozenRows > 0) inner += ` ySplit="${frozenRows}"`;
    inner += ` topLeftCell="${topLeft}" activePane="${pane}" state="frozen"/>`;
  }
  // The active cell is one cell: of a range, its first.
  const active = oneCell(view.activeCell);
  if (active !== undefined) {
    inner += `<selection${pane ? ` pane="${pane}"` : ''} activeCell="${active}" sqref="${active}"/>`;
  }
  return inner === '' ? `${out}/>` : `${out}>${inner}</sheetView>`;
}

/**
 * What the writer of a sheet that is in no file (see writeXlsx()) takes from here, as the
 * children of a worksheet in the default namespace: <sheetPr> ('' for none), the content of
 * <sheetViews>, <sheetFormatPr> ('' for none), and the elements that say how the sheet is
 * printed, which follow the cells.
 */
export function freshSheetParts(sheet: Sheet, selected: boolean): { properties: string; view: string; format: string; protection: string; page: string } {
  const notes = new Notes();
  const pageSetup = sheet.pageSetup ? pageSetupXml(sheet.pageSetup, notes) : undefined;
  const properties = (sheet.tabColor ? colorXml('tabColor', sheet.tabColor) : '') + (pageSetup ? pageSetup.properties : '');
  return {
    properties: properties === '' ? '' : `<sheetPr>${properties}</sheetPr>`,
    view: sheetViewXml(sheet, selected),
    format: sheetFormatXml(sheet),
    protection: protectionXml(sheet),
    page: pageSetup ? pageSetup.page + pageSetup.rowBreaks + pageSetup.colBreaks : '',
  };
}

const ATTR_NAME = /^[A-Za-z_][A-Za-z0-9_.-]*$/;

/** Attributes the model keeps as the file has them, as they go into a start tag: each once, under a name that is one. */
function keptAttrs(attrs: [string, string][] | undefined): string {
  let out = '';
  const seen = new Set<string>();
  for (const pair of attrs || []) {
    if (!pair || typeof pair[0] !== 'string' || typeof pair[1] !== 'string' || !ATTR_NAME.test(pair[0]) || pair[0] === 'xmlns' || seen.has(pair[0])) continue;
    seen.add(pair[0]);
    out += ` ${pair[0]}="${escapeAttr(pair[1])}"`;
  }
  return out;
}

/**
 * The protection of a sheet with its open ranges (<protectedRange sqref>) where the rows and
 * columns inserted and deleted since the sheet was read put them: Sheet.protection is as the
 * file has it, and the writer moves the ranges of the file by the same rule. A range all of
 * whose cells were deleted is gone; everything else of a range (its name, the hash of its
 * password, its security descriptors) is as it was. The protection itself when nothing moved.
 * What a copy of the sheet gets (copySheet()), and what a sheet written from the model says.
 */
export function protectionNow(protection: SheetProtection, edits: readonly SheetStructuralEdit[] | undefined, sheetName: string): SheetProtection {
  if (!protection.ranges || !edits || edits.length === 0) return protection;
  const ranges: NonNullable<SheetProtection['ranges']> = [];
  let changed = false;
  for (const range of protection.ranges) {
    let gone = false;
    let attrs = range.attrs;
    for (let i = 0; i < range.attrs.length && !gone; i++) {
      const [name, value] = range.attrs[i];
      if (name !== 'sqref' || typeof value !== 'string') continue;
      let next = value;
      for (const edit of edits) {
        if (next === '') break;
        next = adjustReferenceText(next, sheetName, { ...edit, sheetName });
      }
      if (next === value) continue;
      if (next === '') gone = true;
      else {
        if (attrs === range.attrs) attrs = range.attrs.slice();
        attrs[i] = [name, collapsedCells(next)];
      }
    }
    if (gone || attrs !== range.attrs) changed = true;
    if (!gone) ranges.push(attrs === range.attrs ? range : { ...range, attrs });
  }
  if (!changed) return protection;
  const out: SheetProtection = { ...protection };
  if (ranges.length > 0) out.ranges = ranges;
  else delete out.ranges;
  return out;
}

/** "B2:B2" -> "B2" for each reference of a list, as the writer states a range that shrank to one cell. */
function collapsedCells(list: string): string {
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

/** <sheetProtection> and <protectedRanges> from the model, the ranges where they are now: '' for a sheet that has neither. */
function protectionXml(sheet: Sheet): string {
  const protection = sheet.protection ? protectionNow(sheet.protection, sheet.structuralEdits, sheet.name) : undefined;
  if (!protection) return '';
  let out = protection.attrs ? `<sheetProtection${keptAttrs(protection.attrs)}/>` : '';
  let ranges = '';
  for (const range of protection.ranges || []) {
    if (!range || !Array.isArray(range.attrs)) continue;
    let inner = '';
    for (const text of range.descriptors || []) if (typeof text === 'string') inner += `<securityDescriptor>${escapeText(text)}</securityDescriptor>`;
    ranges += inner === '' ? `<protectedRange${keptAttrs(range.attrs)}/>` : `<protectedRange${keptAttrs(range.attrs)}>${inner}</protectedRange>`;
  }
  if (ranges !== '') out += `<protectedRanges>${ranges}</protectedRanges>`;
  return out;
}

function sheetFormatXml(sheet: Sheet): string {
  const positive = (value: number | undefined): value is number => typeof value === 'number' && value > 0 && value < Infinity;
  if (!positive(sheet.baseColWidth) && !positive(sheet.defaultColWidth) && !positive(sheet.defaultRowHeight)) return '';
  let out = '<sheetFormatPr';
  if (positive(sheet.baseColWidth)) out += ` baseColWidth="${Math.min(Math.round(sheet.baseColWidth), 255)}"`;
  if (positive(sheet.defaultColWidth)) out += ` defaultColWidth="${sheet.defaultColWidth}"`;
  // The height is the one attribute the element has to have: 15 points where the model states none.
  return `${out} defaultRowHeight="${positive(sheet.defaultRowHeight) ? sheet.defaultRowHeight : 15}"/>`;
}

const sqrefOf = (ranges: CellRange[]) => ranges.filter(validRange).map(rangeName).join(' ');

function conditionalValueXml(value: ConditionalValue, icons: boolean, prefix = ''): string {
  // The automatic limits of the 2010 extension are the lowest and the highest value here.
  const type = value.type === 'autoMin' ? 'min' : value.type === 'autoMax' ? 'max' : value.type;
  let out = `<${prefix}cfvo type="${escapeAttr(type)}"`;
  if (value.value !== undefined) out += ` val="${escapeAttr(value.value)}"`;
  if (icons && !value.gte) out += ' gte="0"';
  return `${out}/>`;
}

const colorXml = (name: string, color: Color) => `<${name}${colorAttrs(color)}/>`;

/** What the source element of a rule holds beside the fields of the model, for conditionalRuleXml(). */
export interface RuleExtra {
  /** Attributes of the <cfRule>, as they go into the start tag. */
  attrs: string;
  /** Children of the <cfRule>, as text. */
  children: string;
  /** Attributes of its <colorScale>, <dataBar> or <iconSet>, as they go into the start tag. */
  holder?: string;
  /**
   * Children of that element that are neither thresholds nor colours, as text: those before the
   * first <cfvo>, before the second, ..., and the last entry what stands at the end of the element.
   */
  between?: string[];
  /**
   * The threshold at `index` of its colour scale, data bar or icon set, given the element the
   * model alone would write: the source element where it says the same, or that element with
   * what changed.
   */
  cfvo?(index: number, written: string): string;
  /** The same for the colour at `index` of a scale or a bar. */
  color?(index: number, written: string): string;
}

const NS_XML = 'http://www.w3.org/XML/1998/namespace';
const NS_XMLNS = 'http://www.w3.org/2000/xmlns/';

/** The attributes conditionalRuleXml() and dataValidationXml() write themselves, per element. */
const WRITTEN_ATTRS = {
  cfRule: ['type', 'dxfId', 'priority', 'stopIfTrue', 'aboveAverage', 'percent', 'bottom', 'operator', 'text', 'timePeriod', 'rank', 'stdDev', 'equalAverage'],
  colorScale: [] as string[],
  dataBar: ['minLength', 'maxLength', 'showValue'],
  iconSet: ['iconSet', 'showValue', 'reverse'],
  cfvo: ['type', 'val'],
  iconCfvo: ['type', 'val', 'gte'],
  color: ['rgb', 'theme', 'indexed', 'tint', 'auto'],
  dataValidation: ['type', 'errorStyle', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorTitle', 'error', 'promptTitle', 'prompt', 'sqref'],
};

/** The part kept content is written into: how it writes its own names, and the namespaces its root declares. */
export interface KeptTarget {
  /** What the part writes before a name of its own namespace: '' or "x:". */
  prefix: string;
  /** The namespace of the part itself. */
  main: string;
  /** The namespace of names without a prefix at the root: '' when the root declares none. */
  plain: string;
  /** The prefix the root declares for a namespace, by namespace. */
  bound: Map<string, string>;
  /** Every prefix the root declares. */
  taken: Set<string>;
}

/** The target for a part with the root attributes `root` whose root element is written with `prefix`. */
export function keptTarget(prefix: string, root: readonly [string, string][]): KeptTarget {
  const target: KeptTarget = { prefix, main: '', plain: '', bound: new Map(), taken: new Set() };
  for (const [name, value] of root) {
    if (name === 'xmlns') target.plain = value;
    else if (name.slice(0, 6) === 'xmlns:') {
      target.taken.add(name.slice(6));
      if (!target.bound.has(value)) target.bound.set(value, name.slice(6));
    }
  }
  if (prefix === '') target.main = target.plain;
  else for (const [uri, bound] of target.bound) if (`${bound}:` === prefix) target.main = uri;
  if (target.main === '') target.main = NS_MAIN;
  return target;
}

/** The target a part made by this module is: its root declares its own namespace and that of relationships. */
const FRESH_TARGET = keptTarget('', [['xmlns', NS_MAIN], ['xmlns:r', NS_REL]]);

/** A name XML can state as a local name: what the reader took from a file is one, what a caller put into the model may not be. */
const isLocalName = (name: unknown): name is string => typeof name === 'string' && name !== '' && !/[\s<>&"'=/:!?;,()[\]{}#%$@^*+|\\~`]/.test(name) && !/^[\d.-]/.test(name);

/**
 * Writes what rules hold beside the fields of the model (RuleKept) as text for one part: names
 * under the prefixes the root of the part declares, and a namespace it does not declare under
 * a prefix that is declared on the element of the rule itself (see declarations()). Values go
 * through the escapes of attributes and text; nothing of it is ever taken as markup.
 */
class KeptWriter {
  private readonly local = new Map<string, string>();
  private made = 0;

  constructor(private readonly target: KeptTarget) {}

  private prefixOf(ns: string): string {
    const bound = this.target.bound.get(ns) || this.local.get(ns);
    if (bound !== undefined) return bound;
    let name: string;
    do name = `ns${++this.made}`;
    while (this.target.taken.has(name));
    this.local.set(ns, name);
    return name;
  }

  /** The namespaces used so far that the root does not declare, as attributes of the element of the rule. */
  declarations(): string {
    let out = '';
    for (const [ns, name] of this.local) out += ` xmlns:${name}="${attrEscape(ns)}"`;
    return out;
  }

  /**
   * Attributes as they go into a start tag. `known` are the attributes the model writes into that
   * tag itself: an element states an attribute once, so one of those is not written a second time
   * (the reader never keeps one; a model put together by hand might), and neither is one that is
   * listed twice or that would declare a namespace.
   */
  attrs(attrs: readonly KeptAttr[] | undefined, known: readonly string[] = []): string {
    let out = '';
    if (!Array.isArray(attrs)) return out;
    const seen = new Set<string>();
    for (const attr of attrs) {
      if (!Array.isArray(attr) || typeof attr[0] !== 'string' || !isLocalName(attr[1]) || typeof attr[2] !== 'string') continue;
      const [ns, name, value] = attr;
      if (ns === NS_XMLNS || (ns === '' && (name === 'xmlns' || known.indexOf(name) >= 0)) || seen.has(`${ns} ${name}`)) continue;
      seen.add(`${ns} ${name}`);
      let stated: string;
      if (ns === '') stated = name;
      else if (ns === NS_XML) stated = `xml:${name}`;
      else if (ns === KEPT_MAIN && this.target.prefix !== '') stated = this.target.prefix + name;
      else stated = `${this.prefixOf(ns === KEPT_MAIN ? this.target.main : ns)}:${name}`;
      out += ` ${stated}="${attrEscape(value)}"`;
    }
    return out;
  }

  /** One element; `plain` is the namespace names without a prefix are in where it stands. */
  node(node: KeptNode, plain: string, depth = 1): string {
    if (!node || typeof node !== 'object' || typeof node.ns !== 'string' || !isLocalName(node.name) || depth > MAX_KEPT_DEPTH) return '';
    const ns = node.ns === KEPT_MAIN ? this.target.main : node.ns;
    let name = node.name;
    let declared = '';
    if (node.ns === KEPT_MAIN && this.target.prefix !== '') name = this.target.prefix + name;
    else if (ns === NS_XML) name = `xml:${name}`;
    else if (node.ns === KEPT_MAIN || ns === '') {
      // A name without a prefix: the namespace such names are in is stated where it is another one.
      if (plain !== ns) {
        declared = ` xmlns="${attrEscape(ns)}"`;
        plain = ns;
      }
    } else name = `${this.prefixOf(ns)}:${name}`;
    let inner = '';
    if (Array.isArray(node.children)) for (const child of node.children) inner += typeof child === 'string' ? textEscape(child) : this.node(child, plain, depth + 1);
    const start = `<${name}${declared}${this.attrs(node.attrs)}`;
    return inner === '' ? `${start}/>` : `${start}>${inner}</${name}>`;
  }

  nodes(nodes: readonly KeptNode[] | undefined): string {
    let out = '';
    if (Array.isArray(nodes)) for (const node of nodes) out += this.node(node, this.target.plain);
    return out;
  }
}

/**
 * What a conditional rule carries beside its fields (`kept`, cut down to what the part can
 * state: see keptFor()) as conditionalRuleXml() takes it. What it holds of a scale, a bar or
 * an icon set goes with the rule while the rule states that element; what it holds of the
 * thresholds and the colours, while the rule has as many of them. Undefined for nothing.
 */
export function keptRuleExtra(rule: ConditionalRule, kept: RuleKept | undefined, target: KeptTarget): RuleExtra | undefined {
  const content = kept && typeof kept === 'object' && !kept.over ? unpackKept(kept.packed) : undefined;
  if (!content) return undefined;
  const writer = new KeptWriter(target);
  const extra: RuleExtra = { attrs: writer.attrs(content.attrs, WRITTEN_ATTRS.cfRule), children: writer.nodes(content.children) };
  const holder = content.holder;
  const now = ruleHolder(rule);
  if (holder && now && holder.kind === now.kind) {
    extra.holder = writer.attrs(holder.attrs, WRITTEN_ATTRS[now.kind]);
    if (Array.isArray(holder.between)) extra.between = holder.between.map((nodes) => writer.nodes(nodes));
    const thresholds = Array.isArray(holder.cfvo) && holder.cfvo.length === now.values ? holder.cfvo.map((item) => (item ? { attrs: writer.attrs(item.attrs, now.kind === 'iconSet' ? WRITTEN_ATTRS.iconCfvo : WRITTEN_ATTRS.cfvo), children: writer.nodes(item.children) } : undefined)) : undefined;
    const colors = Array.isArray(holder.colors) && holder.colors.length === now.colors ? holder.colors.map((item) => (item ? { attrs: writer.attrs(item.attrs, WRITTEN_ATTRS.color), children: writer.nodes(item.children) } : undefined)) : undefined;
    const added = (item: { attrs: string; children: string } | undefined, written: string) => {
      if (!item || written.slice(-2) !== '/>') return written;
      if (item.children === '') return `${written.slice(0, -2)}${item.attrs}/>`;
      const name = /^<([^\s/>]+)/.exec(written);
      return name ? `${written.slice(0, -2)}${item.attrs}>${item.children}</${name[1]}>` : written;
    };
    if (thresholds) extra.cfvo = (index, written) => added(thresholds[index], written);
    if (colors) extra.color = (index, written) => added(colors[index], written);
  }
  // Every piece is written by now: the namespaces they needed are known.
  extra.attrs = writer.declarations() + extra.attrs;
  return extra;
}

/** The same for a data validation, as dataValidationXml() takes it. */
export function keptValidationExtra(kept: RuleKept | undefined, target: KeptTarget): { attrs: string; children: string } | undefined {
  const content = kept && typeof kept === 'object' && !kept.over ? unpackKept(kept.packed) : undefined;
  if (!content) return undefined;
  const writer = new KeptWriter(target);
  const attrs = writer.attrs(content.attrs, WRITTEN_ATTRS.dataValidation);
  const children = writer.nodes(content.children);
  return { attrs: writer.declarations() + attrs, children };
}

/**
 * One <cfRule>, or '' for a rule only the 2010 extension can state. `dxfId` is the record of
 * the style sheet that holds its format (undefined: the rule gives none). `extra` is what the
 * source element of the rule holds beside the fields of the model (attributes as they go into
 * the start tag, children as text; see RuleExtra for what it holds of a scale, a bar or an icon set).
 */
export function conditionalRuleXml(rule: ConditionalRule, priority: number, dxfId: number | undefined, notes?: Notes, prefix = '', extra?: RuleExtra): string {
  if (rule.extended) {
    if (notes) notes.add('Conditional format: a rule of the 2010 extension is left out');
    return '';
  }
  let out = `<${prefix}cfRule type="${escapeAttr(rule.type)}"`;
  if (dxfId !== undefined) out += ` dxfId="${dxfId}"`;
  out += ` priority="${priority}"`;
  if (rule.stopIfTrue) out += ' stopIfTrue="1"';
  if (rule.type === 'aboveAverage') {
    if (!rule.aboveAverage) out += ' aboveAverage="0"';
    if (rule.equalAverage) out += ' equalAverage="1"';
    if (rule.stdDev !== undefined) out += ` stdDev="${rule.stdDev}"`;
  }
  if (rule.type === 'top10') {
    if (rule.percent) out += ' percent="1"';
    if (rule.bottom) out += ' bottom="1"';
  }
  if (rule.operator !== undefined) out += ` operator="${escapeAttr(rule.operator)}"`;
  if (rule.text !== undefined) out += ` text="${escapeAttr(rule.text)}"`;
  if (rule.timePeriod !== undefined) out += ` timePeriod="${escapeAttr(rule.timePeriod)}"`;
  if (rule.rank !== undefined) out += ` rank="${rule.rank}"`;
  if (extra) out += extra.attrs;

  // What the rule holds of a scale, a bar or an icon set is that of the one element it states (the first of these, where a model has more).
  const holder = rule.iconSet ? 'iconSet' : rule.colorScale ? 'colorScale' : rule.dataBar ? 'dataBar' : '';
  const more = (kind: string) => (extra && kind === holder ? extra : undefined);
  const thresholds = (kind: string, values: ConditionalValue[]) => {
    const from = more(kind);
    // The last entry is what stands at the end of the element, however many thresholds there are now.
    const before = from && from.between ? from.between.slice(0, -1) : [];
    let text = '';
    values.forEach((value, i) => {
      const written = conditionalValueXml(value, kind === 'iconSet', prefix);
      text += (before[i] || '') + (from && from.cfvo ? from.cfvo(i, written) : written);
    });
    // What stood before thresholds the element has no more.
    for (let i = values.length; i < before.length; i++) text += before[i];
    return text;
  };
  const colors = (kind: string, list: Color[]) => {
    const from = more(kind);
    let text = '';
    list.forEach((color, i) => {
      const written = colorXml(`${prefix}color`, color);
      text += from && from.color ? from.color(i, written) : written;
    });
    return text;
  };
  const start = (kind: string) => {
    const from = more(kind);
    return from && from.holder ? from.holder : '';
  };
  const end = (kind: string) => {
    const from = more(kind);
    return from && from.between && from.between.length > 0 ? from.between[from.between.length - 1] : '';
  };
  let inner = '';
  for (const formula of rule.formulas.slice(0, 3)) inner += `<${prefix}formula>${escapeText(formula)}</${prefix}formula>`;
  if (rule.colorScale) {
    inner += `<${prefix}colorScale${start('colorScale')}>${thresholds('colorScale', rule.colorScale.values)}${colors('colorScale', rule.colorScale.colors)}${end('colorScale')}</${prefix}colorScale>`;
  }
  if (rule.dataBar) {
    const bar = rule.dataBar;
    if (bar.extended && notes) notes.add('Conditional format: what the 2010 extension says about a data bar (border, axis, negative colours) is left out');
    inner += `<${prefix}dataBar${bar.minLength !== 10 ? ` minLength="${bar.minLength}"` : ''}${bar.maxLength !== 90 ? ` maxLength="${bar.maxLength}"` : ''}${bar.showValue ? '' : ' showValue="0"'}${start('dataBar')}>`;
    inner += thresholds('dataBar', [bar.min, bar.max]) + colors('dataBar', [bar.color || { rgb: 'FF638EC6' }]) + end('dataBar');
    inner += `</${prefix}dataBar>`;
  }
  if (rule.iconSet) {
    const icons = rule.iconSet;
    if (icons.custom && notes) notes.add('Conditional format: icons chosen one by one (2010 extension) are written as their set');
    inner += `<${prefix}iconSet iconSet="${escapeAttr(icons.set)}"${icons.showValue ? '' : ' showValue="0"'}${icons.reverse ? ' reverse="1"' : ''}${start('iconSet')}>`;
    inner += thresholds('iconSet', icons.values) + end('iconSet');
    inner += `</${prefix}iconSet>`;
  }
  if (extra) inner += extra.children;
  return inner === '' ? `${out}/>` : `${out}>${inner}</${prefix}cfRule>`;
}

/**
 * One <conditionalFormatting> element per format of the sheet that has something to say. A
 * format that is as the file has it still has to follow the rows, columns and sheets that
 * changed since (`moved`: '' when its cells are gone); a format added or changed since is
 * where it says (see conditionalFormatNow()). `dxfOf` says which record holds the format of a rule.
 */
function conditionalFormatsXml(sheet: Sheet, dxfOf: (rule: ConditionalRule) => number | undefined, notes: Notes, moved: (element: string) => string): string[] {
  const out: string[] = [];
  const formats = sheet.conditionalFormats;
  if (!formats || formats.length === 0) return out;
  // Priorities are unique in a sheet; a rule without one comes after those that have one.
  let highest = 0;
  for (const format of formats) for (const rule of format.rules) if (rule.priority < Infinity && rule.priority > highest) highest = Math.floor(rule.priority);
  for (const format of formats) {
    const sqref = sqrefOf(format.ranges);
    let rules = '';
    for (const rule of format.rules) {
      // What the element of the rule held beside the fields of the model, as far as a new part can state it.
      const fitted = keptFor(rule.extended ? undefined : keptNow(rule), undefined);
      if (fitted.lost) notes.add('Conditional format: what a rule holds beside its settings in a namespace of an extension is left out');
      rules += conditionalRuleXml(rule, rule.priority < Infinity && rule.priority >= 1 ? Math.floor(rule.priority) : ++highest, dxfOf(rule), notes, '', keptRuleExtra(rule, fitted.kept, FRESH_TARGET));
    }
    if (sqref === '' || rules === '') continue;
    const element = `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`;
    const text = format.live ? element : moved(element);
    if (text !== '') out.push(text);
  }
  return out;
}

/**
 * One <dataValidation> from the model. `extra` is what the source element of the rule holds
 * beside the fields of the model (attributes as they go into the start tag, children as text).
 */
export function dataValidationXml(validation: DataValidation, prefix = '', extra?: { attrs: string; children: string }): string {
  // The messages are ST_Xstring: what XML cannot carry is written as _xHHHH_, as the reader expects it.
  const message = (text: string) => escapeAttr(encodeXstring(text));
  let out = `<${prefix}dataValidation`;
  if (validation.type !== 'none') out += ` type="${escapeAttr(validation.type)}"`;
  if (validation.errorStyle !== 'stop') out += ` errorStyle="${validation.errorStyle}"`;
  if (validation.operator !== 'between') out += ` operator="${escapeAttr(validation.operator)}"`;
  if (validation.allowBlank) out += ' allowBlank="1"';
  if (validation.hideDropDown) out += ' showDropDown="1"';
  if (validation.showInputMessage) out += ' showInputMessage="1"';
  if (validation.showErrorMessage) out += ' showErrorMessage="1"';
  if (validation.errorTitle !== undefined) out += ` errorTitle="${message(validation.errorTitle)}"`;
  if (validation.error !== undefined) out += ` error="${message(validation.error)}"`;
  if (validation.promptTitle !== undefined) out += ` promptTitle="${message(validation.promptTitle)}"`;
  if (validation.prompt !== undefined) out += ` prompt="${message(validation.prompt)}"`;
  if (extra) out += extra.attrs;
  out += ` sqref="${sqrefOf(validation.ranges)}"`;
  let inner = '';
  if (validation.formula1 !== undefined) inner += `<${prefix}formula1>${escapeText(validation.formula1)}</${prefix}formula1>`;
  if (validation.formula2 !== undefined) inner += `<${prefix}formula2>${escapeText(validation.formula2)}</${prefix}formula2>`;
  if (extra) inner += extra.children;
  return inner === '' ? `${out}/>` : `${out}>${inner}</${prefix}dataValidation>`;
}

/**
 * <dataValidations> from the model ('' for a sheet without a rule), in the order of the model:
 * which rule a cell has is the first one that names it. A rule that is as the file has it
 * still has to follow the rows, columns and sheets that changed since (`moved`: '' when its
 * cells are gone); a rule added or changed since is where it says (see dataValidationNow()).
 * Rules of the extension list are among those of the file: written here, their source is
 * read where it is.
 */
function dataValidationsXml(sheet: Sheet, moved: (element: string) => string, notes: Notes): string {
  let rules = '';
  let count = 0;
  for (const validation of sheet.dataValidations || []) {
    if (!validation || !Array.isArray(validation.ranges) || sqrefOf(validation.ranges) === '') continue;
    // As for a conditional rule: what its element held beside the fields of the model.
    const fitted = keptFor(validation.kept, undefined);
    if (fitted.lost) notes.add('Data validation: what a rule holds beside its settings in a namespace of an extension is left out');
    const written = dataValidationXml(validation, '', keptValidationExtra(fitted.kept, FRESH_TARGET));
    const element = validation.live ? written : moved(written);
    if (element === '') continue;
    rules += element;
    count++;
  }
  return count === 0 ? '' : `<dataValidations count="${count}">${rules}</dataValidations>`;
}

/** The filter of a sheet without what only makes sense in the part it was read from. */
function portableFilter(filter: AutoFilter, notes: Notes): AutoFilter {
  const out: AutoFilter = { range: filter.range, columns: [] };
  for (const column of filter.columns) {
    if (column.raw !== undefined && isPortableText(column.raw)) out.columns.push(column);
    else {
      // Without its source element a column is what the model says about it; a colour or an icon criterion is not in the model.
      if (column.kind === 'color' || column.kind === 'icon' || column.kind === 'unknown') notes.add('Filter: a criterion the model only keeps as source text (a colour, an icon) is left out');
      const copy = { ...column };
      delete copy.raw;
      out.columns.push(copy);
    }
  }
  const sort = filter.sortState;
  if (sort && validRange(sort.range)) {
    out.sortState = {
      range: sort.range,
      conditions: sort.conditions.filter((condition) => validRange(condition.range)).map((condition) => (condition.raw !== undefined && !isPortableText(condition.raw) ? { range: condition.range, descending: condition.descending } : condition)),
      caseSensitive: sort.caseSensitive,
      columnSort: sort.columnSort,
    };
    if (sort.attrs) out.sortState.attrs = sort.attrs.filter(([name]) => !isPrefixed(name));
    if (sort.extra && isPortableText(sort.extra)) out.sortState.extra = sort.extra;
  }
  if (filter.attrs) out.attrs = filter.attrs.filter(([name]) => !isPrefixed(name));
  if (filter.extra && isPortableText(filter.extra)) out.extra = filter.extra;
  const dropped =
    (filter.attrs ? filter.attrs.length - (out.attrs as unknown[]).length : 0) +
    (filter.extra && out.extra === undefined ? 1 : 0) +
    (sort && sort.attrs && out.sortState && out.sortState.attrs ? sort.attrs.length - out.sortState.attrs.length : 0) +
    (sort && sort.extra && out.sortState && out.sortState.extra === undefined ? 1 : 0) +
    (sort && !out.sortState ? 1 : 0);
  notes.add('Filter: an attribute or an extension of another namespace (or a sort state outside the sheet) is left out', dropped);
  return out;
}

// --------------------------------------------------------------- page setup

/** The margins applications give a new sheet; they stand in for one the model states no length for. */
const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
/** Page breaks in one direction that applications take. */
const MAX_PAGE_BREAKS = 1026;

const PAGE_ORIENTATIONS = ['default', 'portrait', 'landscape'];
const PAGE_ORDERS = ['downThenOver', 'overThenDown'];
const PAGE_COMMENTS = ['none', 'asDisplayed', 'atEnd'];
const PAGE_ERRORS = ['displayed', 'blank', 'dash', 'NA'];
const PAPER_MEASURE = /^[0-9]{1,9}(\.[0-9]{1,9})?(mm|cm|in|pt|pc|pi)$/;

const isUint32 = (value: unknown): value is number => isIndex(value, 0x100000000);

const BAD_PAGE_VALUE = 'Page setup: a value a file cannot state (a number out of range, a name that is none) is left out';

/** The attributes of an element, each from a field of the model that is checked by `valid`; a field that fails is left out and noted. */
function pageAttrs(source: object, fields: [name: string, valid: (value: unknown) => boolean][], notes: Notes): string {
  let out = '';
  for (const [name, valid] of fields) {
    const value = (source as Record<string, unknown>)[name];
    if (value === undefined) continue;
    if (!valid(value)) notes.add(BAD_PAGE_VALUE);
    else out += ` ${name}="${typeof value === 'boolean' ? (value ? '1' : '0') : escapeAttr(String(value))}"`;
  }
  return out;
}

const isFlag = (value: unknown) => typeof value === 'boolean';
const oneOf = (names: string[]) => (value: unknown) => typeof value === 'string' && names.indexOf(value) >= 0;

/** <rowBreaks> or <colBreaks>: the breaks inside the sheet, one per row or column, in order. `name` may carry a prefix ("x:rowBreaks"). */
export function pageBreaksXml(name: string, breaks: PageBreak[] | undefined, rows: boolean, notes: { add(note: string): void } = new Notes()): string {
  if (!Array.isArray(breaks) || breaks.length === 0) return '';
  const limit = rows ? MAX_ROWS : MAX_COLS;
  const crossLimit = rows ? MAX_COLS : MAX_ROWS;
  const seen = new Set<number>();
  const kept: PageBreak[] = [];
  for (const item of breaks) {
    if (!item || !isIndex(item.at, limit) || item.at < 1 || !isIndex(item.min, crossLimit) || !isIndex(item.max, crossLimit) || item.min > item.max || seen.has(item.at)) {
      notes.add('Page setup: a page break outside the sheet, or a second one at the same place, is left out');
      continue;
    }
    seen.add(item.at);
    kept.push(item);
  }
  if (kept.length === 0) return '';
  if (kept.length > MAX_PAGE_BREAKS) notes.add(`Page setup: a sheet has more than ${MAX_PAGE_BREAKS} page breaks in one direction, which applications may refuse`);
  kept.sort((a, b) => a.at - b.at);
  const prefix = name.slice(0, name.lastIndexOf(':') + 1);
  let manual = 0;
  let inner = '';
  for (const item of kept) {
    if (item.manual) manual++;
    // 0 is the default of min and max: it is not written.
    let attrs = ` id="${item.at}"${item.min !== 0 ? ` min="${item.min}"` : ''}${item.max !== 0 ? ` max="${item.max}"` : ''}${item.manual ? ' man="1"' : ''}${item.pivot ? ' pt="1"' : ''}`;
    // What else the break said in the file goes back with it.
    for (const pair of Array.isArray(item.kept) ? item.kept : []) {
      if (!pair || typeof pair[0] !== 'string' || typeof pair[1] !== 'string' || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(pair[0]) || /^(?:id|min|max|man|pt|xmlns)$/.test(pair[0])) continue;
      attrs += ` ${pair[0]}="${escapeAttr(pair[1])}"`;
    }
    inner += `<${prefix}brk${attrs}/>`;
  }
  return `<${name} count="${kept.length}" manualBreakCount="${manual}">${inner}</${name}>`;
}

interface PageSetupXml {
  /** <pageSetUpPr>, a child of <sheetPr>; '' when the model states nothing for it. */
  properties: string;
  /** <printOptions>, <pageMargins>, <pageSetup> and <headerFooter>, in the order of the schema. */
  page: string;
  rowBreaks: string;
  colBreaks: string;
}

/** The elements that say how a sheet is printed. What the model states and a file cannot hold is noted. */
function pageSetupXml(setup: SheetPageSetup, notes: Notes): PageSetupXml {
  const properties = pageAttrs(setup, [['autoPageBreaks', isFlag], ['fitToPage', isFlag]], notes);
  let page = '';

  if (setup.print) {
    page += `<printOptions${pageAttrs(setup.print, [['horizontalCentered', isFlag], ['verticalCentered', isFlag], ['headings', isFlag], ['gridLines', isFlag], ['gridLinesSet', isFlag]], notes)}/>`;
  }

  if (setup.margins) {
    // The element has to state all six.
    page += '<pageMargins';
    for (const name of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
      let value = setup.margins[name];
      if (!(typeof value === 'number' && value >= 0 && value < Infinity)) {
        value = DEFAULT_MARGINS[name];
        notes.add('Page setup: a margin that is no length is written as the default one');
      }
      page += ` ${name}="${value}"`;
    }
    page += '/>';
  }

  if (setup.page) {
    const isMeasure = (value: unknown) => typeof value === 'string' && PAPER_MEASURE.test(value);
    // "Use the first page number" without a number the file can state would number the pages from 1 by a
    // setting nobody made: the two go together.
    const settings = setup.page.useFirstPageNumber && !isUint32(setup.page.firstPageNumber) ? { ...setup.page, useFirstPageNumber: undefined } : setup.page;
    page += `<pageSetup${pageAttrs(
      settings,
      [
        ['paperSize', isUint32],
        ['paperHeight', isMeasure],
        ['paperWidth', isMeasure],
        ['scale', (value) => isUint32(value) && value >= 10 && value <= 400],
        ['firstPageNumber', isUint32],
        ['fitToWidth', isUint32],
        ['fitToHeight', isUint32],
        ['pageOrder', oneOf(PAGE_ORDERS)],
        ['orientation', oneOf(PAGE_ORIENTATIONS)],
        ['usePrinterDefaults', isFlag],
        ['blackAndWhite', isFlag],
        ['draft', isFlag],
        ['cellComments', oneOf(PAGE_COMMENTS)],
        ['useFirstPageNumber', isFlag],
        ['errors', oneOf(PAGE_ERRORS)],
        ['horizontalDpi', isUint32],
        ['verticalDpi', isUint32],
        ['copies', isUint32],
      ],
      notes,
    )}/>`;
  }

  const texts = setup.headerFooter;
  if (texts) {
    let inner = '';
    for (const name of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) {
      const text = texts[name];
      if (text === undefined) continue;
      if (typeof text !== 'string') notes.add(BAD_PAGE_VALUE);
      // An empty text is none. What XML cannot carry is written as an escape, as in cell text.
      // Spaces at either end are part of the text.
      else if (text !== '') inner += `<${name}${/^\s|\s$/.test(text) ? ' xml:space="preserve"' : ''}>${escapeXstring(text)}</${name}>`;
    }
    const attrs = pageAttrs(texts, [['differentOddEven', isFlag], ['differentFirst', isFlag], ['scaleWithDoc', isFlag], ['alignWithMargins', isFlag]], notes);
    page += inner === '' ? `<headerFooter${attrs}/>` : `<headerFooter${attrs}>${inner}</headerFooter>`;
  }

  return {
    properties: properties === '' ? '' : `<pageSetUpPr${properties}/>`,
    page,
    rowBreaks: pageBreaksXml('rowBreaks', setup.rowBreaks, true, notes),
    colBreaks: pageBreaksXml('colBreaks', setup.colBreaks, false, notes),
  };
}

// ----------------------------------------------------------------- drawings

const emu = (value: number) => Math.round(Math.min(Math.max(value, 0), MAX_EMU));

/** How far a marker is into its cell, in pixels. */
function markerOffset(marker: DrawingMarker, metrics: SheetMetrics): { x: number; y: number } {
  const share = (value: number) => (value > 0 ? (value < 1 ? value : 1) : 0);
  const px = (value: number) => (value > 0 && value < Infinity ? value / EMU_PER_PIXEL : 0);
  const width = metrics.cols.size(marker.col);
  const height = metrics.rows.size(marker.row);
  return {
    x: marker.colShare !== undefined ? share(marker.colShare) * width : Math.min(px(marker.colOff), width),
    y: marker.rowShare !== undefined ? share(marker.rowShare) * height : Math.min(px(marker.rowOff), height),
  };
}

const validMarker = (marker: DrawingMarker | undefined): marker is DrawingMarker => !!marker && isIndex(marker.col, MAX_COLS) && isIndex(marker.row, MAX_ROWS);

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Media of the new package: one part per picture of the model, whichever drawings show it. */
interface MediaStore {
  /** Key of the model -> name of the part ("xl/media/image1.png"); undefined for a picture whose bytes could not be had. */
  parts: Map<string, string | undefined>;
}

interface DrawingContext {
  rels: Rels;
  media: MediaStore;
  /** Relationship of each picture part this drawing shows. */
  embeds: Map<string, string>;
  notes: Notes;
  /** Ids of the objects of one drawing part, which have to differ. */
  nextId: number;
  /** The chart parts of the new package, in the order they are numbered (xl/charts/chart1.xml, ...): a chart that is written adds its part. */
  charts: string[];
  workbook: Workbook;
}

function transformXml(object: DrawingObject, box: Box, group: boolean): string {
  let out = '<a:xfrm';
  if (typeof object.rotation === 'number' && isFinite(object.rotation) && object.rotation !== 0) out += ` rot="${Math.round((((object.rotation % 360) + 360) % 360) * 60000)}"`;
  if (object.flipH) out += ' flipH="1"';
  if (object.flipV) out += ' flipV="1"';
  const place = `<a:off x="${emu(box.x * EMU_PER_PIXEL)}" y="${emu(box.y * EMU_PER_PIXEL)}"/><a:ext cx="${emu(box.width * EMU_PER_PIXEL)}" cy="${emu(box.height * EMU_PER_PIXEL)}"/>`;
  // The members of a group are placed in the coordinates of the sheet, so the space of the group is its own box.
  return `${out}>${place}${group ? place.replace('<a:off', '<a:chOff').replace('<a:ext', '<a:chExt') : ''}</a:xfrm>`;
}

function namesXml(object: DrawingObject, id: number, fallback: string): string {
  let out = `<xdr:cNvPr id="${id}" name="${escapeAttr(object.name !== undefined && object.name !== '' ? object.name : `${fallback} ${id - 1}`)}"`;
  if (object.description !== undefined && object.description !== '') out += ` descr="${escapeAttr(object.description)}"`;
  if (object.hidden) out += ' hidden="1"';
  return `${out}/>`;
}

/** One object of a drawing in the box it takes up on the sheet (pixels); '' for what cannot be written. */
function objectXml(context: DrawingContext, object: DrawingObject, box: Box, depth: number): string {
  if (object.type === 'placeholder') {
    context.notes.add(`Drawing: a ${object.what} is not in the model and is left out`);
    return '';
  }
  if (object.type === 'chart') {
    // A chart of the kinds the commands make is written from its model, with the cells it shows cached as they are;
    // what its part held beside the model (formatting, extensions) is not in a file that is built anew.
    if (!chartWritable(object.chart)) {
      // A chart a reader marked as only shown is left out whatever its type: the reason is where it came from.
      context.notes.add(object.chart && object.chart.shownOnly === true ? 'Drawing: a chart read from an xls file is only shown and is left out' : 'Drawing: a chart of a type that is not written from the model is left out');
      return '';
    }
    if ((object.part !== undefined && object.known !== 'all') || object.source !== undefined) context.notes.add('Drawing: a chart is written from the model; formatting its part held beside it is not kept');
    context.charts.push(chartSpaceXml(object.chart, chartCellsOf(context.workbook, object.chart.live === true)));
    const rel = context.rels.add(`${NS_REL}/chart`, `../charts/chart${context.charts.length}.xml`);
    const frameId = ++context.nextId;
    return (
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${frameId}" name="${escapeAttr(object.name !== undefined && object.name !== '' ? object.name : `Chart ${context.charts.length}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rel}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
    );
  }
  const id = ++context.nextId;
  if (object.type === 'picture') {
    let blip = '<a:blip/>';
    const part = object.media === undefined ? undefined : context.media.parts.get(object.media);
    if (part !== undefined) {
      let embed = context.embeds.get(part);
      if (embed === undefined) context.embeds.set(part, (embed = context.rels.add(`${NS_REL}/image`, `../media/${part.slice(part.lastIndexOf('/') + 1)}`)));
      blip = `<a:blip r:embed="${embed}"/>`;
    }
    else if (object.media !== undefined) context.notes.add('Drawing: a picture whose bytes cannot be read (a metafile, or one missing from the file) is written as an empty frame');
    else if (object.external) context.notes.add('Drawing: the address of a linked picture is not in the model; it is written as an empty frame');
    let crop = '';
    if (object.crop) {
      const side = (value: number) => Math.round(Math.min(Math.max(value, 0), 1) * 100000);
      crop = `<a:srcRect l="${side(object.crop.left)}" t="${side(object.crop.top)}" r="${side(object.crop.right)}" b="${side(object.crop.bottom)}"/>`;
    }
    return (
      `<xdr:pic><xdr:nvPicPr>${namesXml(object, id, 'Picture')}<xdr:cNvPicPr/></xdr:nvPicPr>` +
      `<xdr:blipFill>${blip}${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr>${transformXml(object, box, false)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
    );
  }
  if (object.type === 'shape') {
    // One writer of shapes for every file (see shape-write.ts). A connector has no text of its own; one that the model gives a text is written as a shape of the same geometry.
    const emuBox = { x: box.x * EMU_PER_PIXEL, y: box.y * EMU_PER_PIXEL, width: box.width * EMU_PER_PIXEL, height: box.height * EMU_PER_PIXEL };
    const text = object.text && object.text.paragraphs.length > 0 ? object.text : undefined;
    // What a file cannot hold is written as U+FFFD and counted, as for every other text of the workbook; a name too long is cut and said.
    return shapeElementXml(text === object.text ? object : { ...object, text }, id, emuBox, `${writtenAsConnector({ ...object, text }) ? 'Connector' : 'Shape'} ${id - 1}`, {
      clean: counted,
      onNameCut: () => context.notes.add('Drawing: the name of a shape longer than 255 characters was cut'),
    });
  }
  let members = '';
  if (depth < MAX_GROUP_DEPTH) {
    for (const member of object.children) {
      const share = member.box;
      if (!share) {
        context.notes.add('Drawing: a member of a group without a place in the group is left out');
        continue;
      }
      members += objectXml(context, member, { x: box.x + share.x * box.width, y: box.y + share.y * box.height, width: share.width * box.width, height: share.height * box.height }, depth + 1);
    }
  } else if (object.children.length > 0) context.notes.add('Drawing: the members of a group nested too deeply are left out');
  return `<xdr:grpSp><xdr:nvGrpSpPr>${namesXml(object, id, 'Group')}<xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>${transformXml(object, box, true)}</xdr:grpSpPr>${members}</xdr:grpSp>`;
}

function markerXml(name: string, marker: DrawingMarker, metrics: SheetMetrics): string {
  const offset = markerOffset(marker, metrics);
  return `<xdr:${name}><xdr:col>${marker.col}</xdr:col><xdr:colOff>${emu(offset.x * EMU_PER_PIXEL)}</xdr:colOff><xdr:row>${marker.row}</xdr:row><xdr:rowOff>${emu(offset.y * EMU_PER_PIXEL)}</xdr:rowOff></xdr:${name}>`;
}

/** The anchor element around `inner`, and the box it stands for; undefined for an anchor that names no place on the sheet. */
function anchorXml(anchor: DrawingAnchor, metrics: SheetMetrics): { open: string; close: string; box: Box } | undefined {
  const size = (value: number | undefined) => (typeof value === 'number' && value >= 0 && value < Infinity ? value : 0);
  const start = (marker: DrawingMarker) => {
    const offset = markerOffset(marker, metrics);
    return { x: metrics.cols.offset(marker.col) + offset.x, y: metrics.rows.offset(marker.row) + offset.y };
  };
  if (anchor.kind === 'absolute') {
    const box = { x: size(anchor.x) / EMU_PER_PIXEL, y: size(anchor.y) / EMU_PER_PIXEL, width: size(anchor.width) / EMU_PER_PIXEL, height: size(anchor.height) / EMU_PER_PIXEL };
    return { open: `<xdr:absoluteAnchor><xdr:pos x="${emu(size(anchor.x))}" y="${emu(size(anchor.y))}"/><xdr:ext cx="${emu(size(anchor.width))}" cy="${emu(size(anchor.height))}"/>`, close: '</xdr:absoluteAnchor>', box };
  }
  if (!validMarker(anchor.from)) return undefined;
  const from = start(anchor.from);
  if (anchor.kind === 'oneCell') {
    const box = { ...from, width: size(anchor.width) / EMU_PER_PIXEL, height: size(anchor.height) / EMU_PER_PIXEL };
    return { open: `<xdr:oneCellAnchor>${markerXml('from', anchor.from, metrics)}<xdr:ext cx="${emu(size(anchor.width))}" cy="${emu(size(anchor.height))}"/>`, close: '</xdr:oneCellAnchor>', box };
  }
  if (!validMarker(anchor.to)) return undefined;
  const to = start(anchor.to);
  const editAs = anchor.editAs === 'oneCell' || anchor.editAs === 'absolute' ? ` editAs="${anchor.editAs}"` : '';
  return {
    open: `<xdr:twoCellAnchor${editAs}>${markerXml('from', anchor.from, metrics)}${markerXml('to', anchor.to, metrics)}`,
    close: '</xdr:twoCellAnchor>',
    box: { ...from, width: Math.max(to.x - from.x, 0), height: Math.max(to.y - from.y, 0) },
  };
}

/** The drawing part of a sheet, or '' when nothing of its drawings can be written. */
function drawingXml(sheet: Sheet, metrics: SheetMetrics, context: DrawingContext): string {
  let anchors = '';
  for (const drawing of sheet.drawings || []) {
    // The model keeps anchors as the file has them; rows and columns moved since move them here.
    const anchor = anchorXml(drawingAnchorNow(drawing, sheet.structuralEdits), metrics);
    if (!anchor) {
      context.notes.add('Drawing: an object whose anchor names no cell of the sheet is left out');
      continue;
    }
    const object = objectXml(context, drawing.object, anchor.box, 0);
    if (object !== '') anchors += `${anchor.open}${object}<xdr:clientData/>${anchor.close}`;
  }
  return anchors === '' ? '' : `${XML_DECLARATION}<xdr:wsDr xmlns:xdr="${NS_SHEET_DRAWING}" xmlns:a="${NS_DRAWING}" xmlns:r="${NS_REL}">${anchors}</xdr:wsDr>`;
}

/** What kind of picture the bytes are, from how they start. */
function mediaKind(bytes: Uint8Array): { extension: string; type: string } | undefined {
  const starts = (...signature: number[]) => signature.every((byte, i) => bytes[i] === byte);
  if (starts(0x89, 0x50, 0x4e, 0x47)) return { extension: 'png', type: 'image/png' };
  if (starts(0xff, 0xd8, 0xff)) return { extension: 'jpeg', type: 'image/jpeg' };
  if (starts(0x47, 0x49, 0x46, 0x38)) return { extension: 'gif', type: 'image/gif' };
  if (starts(0x42, 0x4d)) return { extension: 'bmp', type: 'image/bmp' };
  if (starts(0x49, 0x49, 0x2a, 0x00) || starts(0x4d, 0x4d, 0x00, 0x2a)) return { extension: 'tiff', type: 'image/tiff' };
  if (starts(0x01, 0x00, 0x00, 0x00) && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46) return { extension: 'emf', type: 'image/x-emf' };
  if (starts(0xd7, 0xcd, 0xc6, 0x9a)) return { extension: 'wmf', type: 'image/x-wmf' };
  return undefined;
}

function collectMedia(objects: DrawingObject[], keys: string[], seen: Set<string>, depth: number): void {
  for (const object of objects) {
    if (object.type === 'picture' && object.media !== undefined && !seen.has(object.media)) {
      seen.add(object.media);
      keys.push(object.media);
    } else if (object.type === 'group' && depth < MAX_GROUP_DEPTH) collectMedia(object.children, keys, seen, depth + 1);
  }
}

// ----------------------------------------------------------------- comments

interface WrittenComments {
  comments: string;
  vml: string;
}

/** The comments part of a sheet and the legacy drawing that holds their boxes; undefined for a sheet without comments. */
function commentsXml(sheet: Sheet, metrics: () => SheetMetrics, blocks: { next: number }, notes: Notes): WrittenComments | undefined {
  const list: SheetComment[] = [];
  const seen = new Set<number>();
  for (const stated of sheet.comments || []) {
    // Like anchors, comments stay where the file has them in the model: they follow the rows and columns
    // moved since, and one whose cell was deleted went with it.
    let comment = stated;
    // A note added since stands where it was put: only the edits made after it move it.
    const edits = commentEdits(stated, sheet.structuralEdits);
    if (edits && edits.length > 0 && isIndex(stated.row, MAX_ROWS) && isIndex(stated.col, MAX_COLS)) {
      const cell = adjustCommentCell(stated.row, stated.col, edits);
      if (!cell) continue;
      comment = { ...stated, row: cell.row, col: cell.col };
      if (stated.box) comment.box = adjustDrawingAnchor(stated.box, edits);
    }
    if (!isIndex(comment.row, MAX_ROWS) || !isIndex(comment.col, MAX_COLS)) {
      notes.add('Comment: a comment outside the sheet is left out');
      continue;
    }
    const key = comment.row * MAX_COLS + comment.col;
    if (seen.has(key)) {
      notes.add('Comment: a second comment on the same cell is left out');
      continue;
    }
    seen.add(key);
    list.push(comment);
  }
  if (list.length === 0) return undefined;

  const authors: string[] = [];
  const authorIds = new Map<string, number>();
  const authorId = (name: string): number => {
    let id = authorIds.get(name);
    if (id === undefined) {
      id = authors.length;
      authors.push(name);
      authorIds.set(name, id);
    }
    return id;
  };
  let items = '';
  let shapes = '';
  const used: number[] = [];
  list.forEach((comment, index) => {
    if (comment.threaded) notes.add('Comment: a conversation is written as a note with its first message; replies are left out', 1 + (comment.replies ? comment.replies.length : 0));
    const text = typeof comment.text === 'string' ? comment.text : '';
    items += `<comment ref="${cellName(comment.row, comment.col)}" authorId="${authorId(comment.author || '')}"><text>${richTextXml(text, comment.runs)}</text></comment>`;

    // The box: where the model says it is, else beside the cell, where applications put a new one.
    let numbers = defaultNoteAnchor(comment.row, comment.col);
    const box = comment.box;
    if (box && box.kind === 'twoCell' && validMarker(box.from) && validMarker(box.to)) {
      const from = markerOffset(box.from, metrics());
      const to = markerOffset(box.to, metrics());
      numbers = [box.from.col, Math.round(from.x), box.from.row, Math.round(from.y), box.to.col, Math.round(to.x), box.to.row, Math.round(to.y)];
    }
    // Shape ids come in blocks; each block a drawing uses is named in its <o:idmap>.
    // Block n stands for the ids n*1024 .. n*1024+1023, of which the first is not used.
    if (index % (VML_BLOCK - 1) === 0) used.push(blocks.next++);
    const id = used[used.length - 1] * VML_BLOCK + (index % (VML_BLOCK - 1)) + 1;
    shapes += noteShapeXml(id, comment.row, comment.col, numbers, !!comment.visible);
  });

  let comments = `${XML_DECLARATION}<comments xmlns="${NS_MAIN}"><authors>`;
  for (const author of authors) comments += `<author>${escapeXstring(author)}</author>`;
  comments += `</authors><commentList>${items}</commentList></comments>`;
  const vml =
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${used.join(',')}"/></o:shapelayout>` +
    `${noteShapeTypeXml()}${shapes}</xml>`;
  return { comments, vml };
}

// ------------------------------------------------------------------- sheets

/** Cells of one link that are remembered as linked; a range larger than this is told from others by its first cells. */
const LINK_CELLS = 4096;

/**
 * The column ranges as a file states them: in order, none overlapping an earlier one, with a
 * width a file can hold and a format the workbook has.
 */
function portableColumns(columns: ColumnRange[], formatCount: number, badFormat: string, notes: Notes): ColumnRange[] {
  const out: ColumnRange[] = [];
  let next = 0;
  for (const range of columns) {
    let min = range.min;
    const max = Math.min(range.max, MAX_COLS - 1);
    const valid = isIndex(min, MAX_COLS) && isIndex(max, MAX_COLS) && min <= max;
    if (valid && min < next) min = next;
    if (!valid || min !== range.min) notes.add('Column: a range that overlaps an earlier one, or is not one, is cut to what is left of it');
    if (!valid || min > max) continue;
    const copy: ColumnRange = { ...range, min, max };
    if (copy.width !== undefined && !(typeof copy.width === 'number' && copy.width >= 0 && copy.width <= 255)) {
      if (typeof copy.width === 'number' && copy.width > 255 && copy.width < Infinity) copy.width = 255;
      else delete copy.width;
      notes.add('Column: a width a file cannot state is left out or cut to 255');
    }
    if (copy.style !== undefined && !isIndex(copy.style, formatCount)) {
      copy.style = 0;
      notes.add(badFormat);
    }
    if (!isIndex(copy.outlineLevel, 8)) copy.outlineLevel = 0;
    out.push(copy);
    next = max + 1;
  }
  return out;
}

/** Attributes of <row> that the writer of rows states from the fields of the model. */
const ROW_FIELDS: { [name: string]: true } = { r: true, spans: true, customFormat: true, ht: true, customHeight: true, hidden: true, outlineLevel: true, collapsed: true };

/**
 * The attributes a row keeps beside its fields, as far as they can be written: each once, under a
 * name that is one, and none that the writer states itself or that belongs to a namespace the
 * new part does not declare. The list itself when all of it can be written.
 */
function portableRowAttrs(row: Row, formatCount: number, badFormat: string, notes: Notes): Row['attrs'] {
  const attrs = row.attrs;
  if (!attrs) return attrs;
  const seen = new Set<string>();
  const kept = attrs.filter((pair) => {
    const name = pair[0];
    const value = pair[1];
    if (typeof value === 'string') counted(value);
    if (name === 's' && typeof value === 'string' && !seen.has(name)) {
      // The format of a row that does not apply it (no customFormat): kept by the reader, an index all the same.
      seen.add(name);
      if (row.style !== undefined) return true;
      if (/^\d{1,9}$/.test(value) && Number(value) < formatCount) return true;
      notes.add(badFormat);
      return false;
    }
    if (typeof name !== 'string' || typeof value !== 'string' || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(name) || name === 'xmlns' || ROW_FIELDS[name] === true || seen.has(name)) {
      notes.add('Row: an attribute that cannot be written as it is (one the writer states itself, one of another namespace, or one whose name is not a name) is left out');
      return false;
    }
    seen.add(name);
    return true;
  });
  return kept.length === attrs.length ? attrs : kept;
}

interface SheetParts {
  xml: string;
  /** A formula of the sheet has no stored result. */
  pending?: boolean;
  rels: Rels;
  drawing?: { xml: string; rels: Rels };
  comments?: WrittenComments;
}

interface SheetNumbers {
  /** The chart parts written so far, all sheets together: see DrawingContext.charts. */
  charts: string[];
  drawing: number;
  comments: number;
  vmlBlocks: { next: number };
  /** A dynamic array formula was written: the package gets the cell metadata that marks it (`cm="1"`). */
  dynamic?: boolean;
}

function sheetXml(workbook: Workbook, sheet: Sheet, strings: StringTable, media: MediaStore, numbers: SheetNumbers, notes: Notes, selected: boolean, dxfOf: (rule: ConditionalRule) => number | undefined): SheetParts {
  const rels = new Rels();
  const out: SheetParts = { xml: '', rels };
  let metrics: SheetMetrics | undefined;
  const metricsOf = () => metrics || (metrics = createSheetMetrics(sheet));

  const parts: string[] = [`${XML_DECLARATION}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`];
  const pageSetup = sheet.pageSetup ? pageSetupXml(sheet.pageSetup, notes) : undefined;
  const properties = (sheet.tabColor ? colorXml('tabColor', sheet.tabColor) : '') + (pageSetup ? pageSetup.properties : '');
  if (properties !== '') parts.push(`<sheetPr>${properties}</sheetPr>`);
  // Filled in once the rows are written: the range of what was written.
  const dimensionAt = parts.length;
  parts.push('');
  parts.push(`<sheetViews>${sheetViewXml(sheet, selected)}</sheetViews>`);
  parts.push(sheetFormatXml(sheet));
  const formatCount = Math.max(workbook.styles.cellXfs.length, 1);
  const badFormat = 'Format: a cell, row or column refers to a format the workbook does not have and is written with the default one';
  const columns = portableColumns(sheet.columns, formatCount, badFormat, notes);
  parts.push(colsXml({ ...sheet, columns }, '', []));

  // Rows in the order of the sheet, whatever order the model got them in.
  const indexes: number[] = [];
  let outside = 0;
  for (const index of sheet.rows.keys()) {
    if (isIndex(index, MAX_ROWS)) indexes.push(index);
    else outside++;
  }
  if (outside > 0) notes.add('Sheet: a row outside the sheet is left out', outside);
  indexes.sort((a, b) => a - b);
  const masters = sharedMastersOf(sheet, indexes);
  const writer: CellWriter = {
    prefix: '',
    shared: workbook.sharedStrings,
    styledColumns: columns.filter((range) => !!range.style),
    intern: (text, runs) => strings.intern(text, runs),
    dynamicCm: () => {
      numbers.dynamic = true;
      return 1;
    },
  };
  const used = { top: -1, bottom: -1, left: MAX_COLS, right: -1 };
  const dataAt = parts.length;
  parts.push('<sheetData>');
  for (const index of indexes) {
    let row = sheet.rows.get(index) as Row;
    // Only the cells of the sheet are looked at: a row that is sparse far beyond it is not walked.
    let source = row.cells;
    if (source.length > MAX_COLS) {
      source = source.slice(0, MAX_COLS);
      notes.add('Sheet: cells outside the sheet are left out');
    }
    let cells: Cell[] | undefined = source === row.cells ? undefined : source;
    source.forEach((cell, col) => {
      if (used.top < 0) used.top = index;
      used.bottom = index;
      if (col < used.left) used.left = col;
      if (col > used.right) used.right = col;
      let writable = cell;
      if (!isIndex(cell.style, formatCount)) {
        writable = { ...writable, style: 0 };
        notes.add(badFormat);
      }
      if (cell.kind === 'number' && !(typeof cell.value === 'number' && isFinite(cell.value))) notes.add('Cell: a number that is not finite is written as #NUM!');
      else if (cell.kind === 'error' || cell.kind === 'date') counted(String(cell.value));
      else if (cell.kind === 'string' && !cell.formula && cell.si !== undefined) strings.uses++;
      if (cell.formula || cell.cm !== undefined || cell.vm !== undefined) writable = writableCell(writable, index, col, masters, notes);
      if (writable.formula && writable.kind === 'blank') out.pending = true;
      if (writable === cell) return;
      // slice() keeps the holes of a sparse row.
      if (!cells) cells = source.slice();
      cells[col] = writable;
    });
    const attrs = portableRowAttrs(row, formatCount, badFormat, notes);
    let style = row.style;
    if (style !== undefined && !isIndex(style, formatCount)) {
      style = 0;
      notes.add(badFormat);
    }
    if (cells || attrs !== row.attrs || style !== row.style) row = { ...row, cells: cells || row.cells, attrs, style };
    const text = rowXml(writer, index, row);
    if (text !== '') parts.push(text);
  }
  if (parts.length === dataAt + 1) parts[dataAt] = '<sheetData/>';
  else parts.push('</sheetData>');
  parts[dimensionAt] = `<dimension ref="${used.top < 0 ? 'A1' : rangeName(used)}"/>`;
  // Behind the cells, where the schema has it.
  parts.push(protectionXml(sheet));

  // Conditional formats and validations are kept as the file has them, sheet names included: the
  // rows and columns moved since (in this sheet, and in the sheets their formulas name) are matched
  // with the names the file has, as a save matches them, and then the sheets get the names they
  // have now in one pass (a sheet that was deleted leaves an error).
  // UNSURE: a 3D reference in such content may come out differently than the same reference in a
  // cell when a sheet at its end was deleted and rows moved in its neighbour afterwards: cells are
  // adjusted one operation after the other, this content by the state of the file at once. Excel
  // allows no 3D reference in a conditional format or a validation, so it takes a hand-made file.
  const inFile = workbook.sheets.some(isFileSheet) || (!!workbook.removedSheetNames && workbook.removedSheetNames.length > 0);
  /** The name content of the file has for a sheet; undefined for a sheet nothing of the file can name. */
  const fileNameOf = (of: Sheet): string | undefined => {
    const name = fileSheetName(of);
    // A workbook none of whose sheets was read from a file (a model put together by hand) is taken as it is.
    return name === undefined && !inFile ? of.name : name;
  };
  const editsOf = (of: Sheet): StructuralEdit[] => {
    const name = fileNameOf(of);
    return name === undefined ? [] : (of.structuralEdits || []).map((edit) => ({ ...edit, sheetName: name }));
  };
  const ownFileName = fileNameOf(sheet);
  // References without a sheet name are the sheet's own, whatever it is called.
  const ownName = ownFileName === undefined ? NOT_IN_FILE : ownFileName;
  const own: StructuralEdit[] = (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: ownName }));
  let foreign: StructuralEdit[] = [];
  for (const other of workbook.sheets) if (other !== sheet) foreign = foreign.concat(editsOf(other));
  const nameMap = sheetNameMap(workbook);
  let rename: { map: SheetNameMap; names: string[] } | undefined;
  if (nameMap) {
    const names = (workbook.removedSheetNames || []).slice();
    for (const other of workbook.sheets) if (isFileSheet(other) && other.nameInFile !== undefined && other.nameInFile !== other.name) names.push(other.nameInFile);
    rename = { map: nameMap, names };
  }

  if (sheet.autoFilter && validRange(sheet.autoFilter.range)) parts.push(autoFilterXml({ ...sheet, autoFilter: portableFilter(sheet.autoFilter, notes) }, ''));
  const merges = sheet.merges.filter(validRange);
  if (merges.length !== sheet.merges.length) notes.add('Sheet: a merged range outside the sheet is left out', sheet.merges.length - merges.length);
  parts.push(mergeCellsXml({ ...sheet, merges }, ''));
  for (const format of conditionalFormatsXml(sheet, dxfOf, notes, (element) => adjustSheetElement(element, ownName, own, foreign, rename))) parts.push(format);
  parts.push(dataValidationsXml(sheet, (element) => adjustSheetElement(element, ownName, own, foreign, rename), notes));

  let links = '';
  /** Cells that have a link; the cells of a very large range are noted as far as LINK_CELLS goes. */
  const linked = new Set<number>();
  for (const link of sheet.hyperlinks) {
    const target = typeof link.target === 'string' && link.target !== '' ? link.target : undefined;
    const location = typeof link.location === 'string' && link.location !== '' ? link.location : undefined;
    if (!validRange(link.range) || (target === undefined && location === undefined)) {
      notes.add('Hyperlink: a link without a target (or outside the sheet) is left out');
      continue;
    }
    const keys: number[] = [];
    let taken = false;
    cover: for (let r = link.range.top; r <= link.range.bottom; r++) {
      for (let c = link.range.left; c <= link.range.right; c++) {
        if (keys.length >= LINK_CELLS) break cover;
        const key = r * MAX_COLS + c;
        if (linked.has(key)) {
          taken = true;
          break cover;
        }
        keys.push(key);
      }
    }
    if (taken) {
      notes.add('Hyperlink: a second link on the same cell is left out');
      continue;
    }
    for (const key of keys) linked.add(key);
    links += `<hyperlink ref="${rangeName(link.range)}"`;
    // A control character is no part of an address: it is written percent-encoded. Nothing else of the target is touched.
    if (target !== undefined) links += ` r:id="${rels.add(`${NS_REL}/hyperlink`, target.replace(/[\u0000-\u001f\u007f]/g, (ch) => `%${(ch.charCodeAt(0) < 16 ? '0' : '') + ch.charCodeAt(0).toString(16).toUpperCase()}`), true)}"`;
    if (location !== undefined) links += ` location="${escapeAttr(location)}"`;
    if (link.display !== undefined) links += ` display="${escapeAttr(link.display)}"`;
    if (link.tooltip !== undefined) links += ` tooltip="${escapeAttr(link.tooltip)}"`;
    links += '/>';
  }
  if (links !== '') parts.push(`<hyperlinks>${links}</hyperlinks>`);

  if (pageSetup) {
    parts.push(pageSetup.page);
    // Like conditional formats, the breaks are kept as the file has them and follow the rows and columns moved since.
    // Breaks set since the file was read follow only the edits made after them (SheetPageSetup.breaksSince).
    const since = sheet.pageSetup && typeof sheet.pageSetup.breaksSince === 'number' && sheet.pageSetup.breaksSince > 0 ? Math.floor(sheet.pageSetup.breaksSince) : 0;
    const moves = since > 0 ? own.slice(Math.min(since, own.length)) : own;
    if (pageSetup.rowBreaks !== '') parts.push(adjustSheetElement(pageSetup.rowBreaks, ownName, moves, foreign));
    if (pageSetup.colBreaks !== '') parts.push(adjustSheetElement(pageSetup.colBreaks, ownName, moves, foreign));
  }

  if (sheet.drawings && sheet.drawings.length > 0) {
    const context: DrawingContext = { rels: new Rels(), media, embeds: new Map(), notes, nextId: 1, charts: numbers.charts, workbook };
    const xml = drawingXml(sheet, metricsOf(), context);
    if (xml !== '') {
      out.drawing = { xml, rels: context.rels };
      parts.push(`<drawing r:id="${rels.add(`${NS_REL}/drawing`, `../drawings/drawing${++numbers.drawing}.xml`)}"/>`);
    }
  }
  const comments = commentsXml(sheet, metricsOf, numbers.vmlBlocks, notes);
  if (comments) {
    out.comments = comments;
    numbers.comments++;
    parts.push(`<legacyDrawing r:id="${rels.add(`${NS_REL}/vmlDrawing`, `../drawings/vmlDrawing${numbers.comments}.vml`)}"/>`);
    rels.add(`${NS_REL}/comments`, `../comments${numbers.comments}.xml`);
  }

  // What stands around the sheet in the file it was read from and is not in the model.
  for (const kind of sheet.anchored || []) {
    if (kind !== 'drawing' && kind !== 'comments' && kind !== 'vmlDrawing' && kind !== 'threadedComment' && kind !== 'image') notes.add(`Sheet: ${quote(kind)} content of the source file is not in the model and is left out`);
  }
  for (const warning of sheet.warnings || []) notes.add(`Sheet ${quote(sheet.name)} was read with a warning: ${warning}`);

  parts.push('</worksheet>');
  out.xml = parts.join('');
  return out;
}

// ----------------------------------------------------------------- workbook

/** Sheet ids of the file: the ones of the model where they can be used, new ones elsewhere. */
function sheetIdsOf(sheets: Sheet[]): number[] {
  const taken = new Set<number>();
  const ids = sheets.map((sheet) => {
    const id = sheet.sheetId;
    if (typeof id !== 'number' || id < 1 || id > 65534 || Math.floor(id) !== id || taken.has(id)) return 0;
    taken.add(id);
    return id;
  });
  let next = 1;
  return ids.map((id) => {
    if (id !== 0) return id;
    while (taken.has(next)) next++;
    taken.add(next);
    return next;
  });
}

/**
 * Writes a workbook model as a new xlsx file: every sheet with its cells,
 * formulas and their stored results, formats, merged cells, sizes, panes,
 * links, comments, drawings and page setup, the defined names and the styles. Nothing of
 * the file the model was read from is used, so a model read from an xls file
 * (or made in code) can be written. What the model holds and the file cannot
 * carry is listed in `options.report`.
 */
export async function buildXlsx(workbook: Workbook, options: BuildXlsxOptions = {}): Promise<Uint8Array> {
  // The pictures are read first (which takes moments); then the workbook is read in one go, so that the
  // file is the workbook of one moment. A picture the workbook shows by then and that was not read is read, and the reading starts over.
  const pictures = new Map<string, Uint8Array | undefined>();
  for (;;) {
    const keys: string[] = [];
    const seen = new Set<string>();
    for (const sheet of workbook.sheets) if (sheet.drawings) collectMedia(sheet.drawings.map((drawing) => drawing.object), keys, seen, 0);
    for (const key of keys) {
      if (pictures.has(key)) continue;
      let bytes: Uint8Array | undefined;
      try {
        // Pictures added since the workbook was read are in the model; the others are read from the file.
        bytes = await pictureSource(workbook).readPart(key);
      } catch (error) {
        bytes = undefined;
      }
      pictures.set(key, bytes);
    }
    const zip = readModel(workbook, () => builtXlsx(workbook, options, pictures));
    if (zip) return zip.write();
  }
}

/** The package of a workbook, from the workbook as it is and the pictures that were read; undefined when the workbook shows a picture that was not. */
function builtXlsx(workbook: Workbook, options: BuildXlsxOptions, pictures: Map<string, Uint8Array | undefined>): ZipArchive | undefined {
  const notes = new Notes();
  const strings = new StringTable(workbook.sharedStrings);

  let sheets = workbook.sheets;
  if (sheets.length === 0) {
    // A workbook has at least one sheet.
    notes.add('Workbook: it has no sheets; an empty one is written');
    sheets = [
      {
        name: 'Sheet1',
        sheetId: 1,
        state: 'visible',
        kind: 'worksheet',
        part: '',
        rows: new Map(),
        columns: [],
        merges: [],
        hyperlinks: [],
        rowCount: 0,
        colCount: 0,
        view: { showGridLines: true, showRowColHeaders: true, showZeros: true, rightToLeft: false, tabSelected: true, zoomScale: 100, frozenRows: 0, frozenCols: 0 },
      },
    ];
  }

  // Pictures first: every drawing that shows one refers to the same part.
  const media: MediaStore = { parts: new Map() };
  const mediaData: [name: string, data: Uint8Array][] = [];
  const mediaTypes = new Map<string, string>();
  const keys: string[] = [];
  const seenKeys = new Set<string>();
  for (const sheet of sheets) if (sheet.drawings) collectMedia(sheet.drawings.map((drawing) => drawing.object), keys, seenKeys, 0);
  for (const key of keys) if (!pictures.has(key)) return undefined;
  for (const key of keys) {
    const bytes = pictures.get(key);
    const kind = bytes && bytes.length > 0 ? mediaKind(bytes) : undefined;
    if (!bytes || !kind) {
      media.parts.set(key, undefined);
      continue;
    }
    const name = `xl/media/image${mediaData.length + 1}.${kind.extension}`;
    media.parts.set(key, name);
    mediaData.push([name, bytes]);
    mediaTypes.set(kind.extension, kind.type);
  }

  // The sheet the file opens on, which is also the one sheet its views mark as selected: one that is shown.
  let active = isIndex(workbook.activeSheet, sheets.length) ? workbook.activeSheet : 0;
  const shown = sheets.findIndex((sheet) => sheet.state !== 'hidden' && sheet.state !== 'veryHidden');
  const activeHidden = shown >= 0 && (sheets[active].state === 'hidden' || sheets[active].state === 'veryHidden');
  if (activeHidden) active = shown;

  // From here to the end nothing is awaited: the count is this build's alone.
  replacedCharacters = 0;
  const numbers: SheetNumbers = { charts: [], drawing: 0, comments: 0, vmlBlocks: { next: 1 } };
  // The formats conditional rules state themselves become records of the style sheet.
  const dxfs = planDxfs(workbook);
  const written: { name: string; parts: SheetParts; drawing: number; comments: number }[] = [];
  let pendingFormulas = false;
  sheets.forEach((sheet, index) => {
    if (sheet.kind !== 'worksheet') {
      // A chart sheet has no cells and a macro sheet is never read: both are written as worksheets, so
      // that every sheet keeps its name and its position (formulas and names refer to them).
      notes.add(sheet.kind === 'dialogsheet' ? 'Sheet: a dialog sheet is written as a worksheet' : `Sheet: the content of a ${sheet.kind === 'chartsheet' ? 'chart sheet' : 'macro sheet'} is not in the model; it is written as an empty worksheet`);
    }
    const before = { drawing: numbers.drawing, comments: numbers.comments };
    const parts = sheetXml(workbook, sheet, strings, media, numbers, notes, index === active, dxfs.indexOf);
    if (parts.pending) pendingFormulas = true;
    written.push({ name: `xl/worksheets/sheet${index + 1}.xml`, parts, drawing: numbers.drawing > before.drawing ? numbers.drawing : 0, comments: numbers.comments > before.comments ? numbers.comments : 0 });
  });

  // The workbook part: children in the order CT_Workbook gives them.
  const bookRels = new Rels();
  const ids = sheetIdsOf(sheets);
  let book = `${XML_DECLARATION}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><workbookPr${workbook.date1904 ? ' date1904="1"' : ''}/>`;
  // How the workbook is protected, as it was read (an xls file's PROTECT and WINDOWPROTECT of its globals, say).
  if (workbook.protection && workbook.protection.attrs.length > 0) book += `<workbookProtection${keptAttrs(workbook.protection.attrs)}/>`;
  if (activeHidden) notes.add('Workbook: the active sheet is hidden; the first visible one is made the active sheet');
  book += `<bookViews><workbookView${active > 0 ? ` activeTab="${active}"` : ''}/></bookViews><sheets>`;
  sheets.forEach((sheet, index) => {
    const state = sheet.state === 'hidden' || sheet.state === 'veryHidden' ? ` state="${sheet.state}"` : '';
    book += `<sheet name="${escapeAttr(sheet.name)}" sheetId="${ids[index]}"${state} r:id="${bookRels.add(`${NS_REL}/worksheet`, `worksheets/sheet${index + 1}.xml`)}"/>`;
  });
  book += '</sheets>';
  // Names are written as the model has them: formulas and defined names refer to them. One an application refuses is said.
  const sheetNames = new Set<string>();
  for (const sheet of sheets) {
    const key = sheet.name.toLowerCase();
    if (sheet.name === '' || sheet.name.length > 31 || /[\[\]:*?/\\]/.test(sheet.name) || sheetNames.has(key)) notes.add('Workbook: a sheet has a name applications refuse (empty, longer than 31 characters, one of []:*?/\\ in it, or used twice)');
    sheetNames.add(key);
  }
  if (shown < 0) notes.add('Workbook: no sheet is visible; applications refuse such a file');
  let names = '';
  const scopes = new Set<string>();
  for (const name of workbook.definedNames) {
    if (typeof name.name !== 'string' || name.name === '' || typeof name.value !== 'string' || name.value === '') {
      notes.add('Defined name: a name without a name or without a value is left out');
      continue;
    }
    if (name.localSheet !== undefined && !isIndex(name.localSheet, sheets.length)) {
      notes.add('Defined name: a name of a sheet the workbook does not have is left out');
      continue;
    }
    // Names are told apart without regard to case, each in its scope.
    const scope = `${name.localSheet === undefined ? '' : name.localSheet}!${name.name.toLowerCase()}`;
    if (scopes.has(scope)) {
      notes.add('Defined name: a second name with the same name and scope is left out');
      continue;
    }
    scopes.add(scope);
    if (refersToWorkbook(name.value)) notes.add('Defined name: a name refers to another workbook; the link to that workbook is not in the model and is not written');
    names += `<definedName name="${escapeAttr(name.name)}"${name.localSheet !== undefined ? ` localSheetId="${name.localSheet}"` : ''}${name.hidden ? ' hidden="1"' : ''}>${escapeText(name.value)}</definedName>`;
  }
  if (names !== '') book += `<definedNames>${names}</definedNames>`;
  // A formula without a stored result gets one when the file is opened; results the model has stay as they are.
  if (pendingFormulas) book += '<calcPr fullCalcOnLoad="1"/>';
  book += '</workbook>';

  const overrides: [part: string, type: string][] = [['xl/workbook.xml', `${TYPE_PREFIX}spreadsheetml.sheet.main+xml`]];
  const files: [name: string, data: Uint8Array | string][] = [];
  for (const sheet of written) {
    files.push([sheet.name, sheet.parts.xml]);
    overrides.push([sheet.name, `${TYPE_PREFIX}spreadsheetml.worksheet+xml`]);
    if (sheet.parts.rels.size > 0) files.push([relsNameOf(sheet.name), sheet.parts.rels.text()]);
    if (sheet.parts.drawing) {
      const name = `xl/drawings/drawing${sheet.drawing}.xml`;
      files.push([name, sheet.parts.drawing.xml]);
      overrides.push([name, `${TYPE_PREFIX}drawing+xml`]);
      if (sheet.parts.drawing.rels.size > 0) files.push([relsNameOf(name), sheet.parts.drawing.rels.text()]);
    }
    if (sheet.parts.comments) {
      const name = `xl/comments${sheet.comments}.xml`;
      files.push([name, sheet.parts.comments.comments]);
      overrides.push([name, `${TYPE_PREFIX}spreadsheetml.comments+xml`]);
      files.push([`xl/drawings/vmlDrawing${sheet.comments}.vml`, sheet.parts.comments.vml]);
    }
  }
  for (const [name, data] of mediaData) files.push([name, data]);
  numbers.charts.forEach((xml, i) => {
    files.push([`xl/charts/chart${i + 1}.xml`, xml]);
    overrides.push([`xl/charts/chart${i + 1}.xml`, CHART_CONTENT_TYPE]);
  });

  if (workbook.styles.themeColors.length > 0) {
    files.push(['xl/theme/theme1.xml', themeXml(workbook.styles, notes)]);
    overrides.push(['xl/theme/theme1.xml', `${TYPE_PREFIX}theme+xml`]);
    bookRels.add(`${NS_REL}/theme`, 'theme/theme1.xml');
  }
  files.push(['xl/styles.xml', stylesXml(workbook.styles, notes, dxfs.added)]);
  overrides.push(['xl/styles.xml', `${TYPE_PREFIX}spreadsheetml.styles+xml`]);
  bookRels.add(`${NS_REL}/styles`, 'styles.xml');
  if (numbers.dynamic) {
    files.push(['xl/metadata.xml', newMetadataXml()]);
    overrides.push(['xl/metadata.xml', METADATA_CONTENT_TYPE]);
    bookRels.add(`${NS_REL}/${METADATA_RELATIONSHIP}`, 'metadata.xml');
  }
  if (strings.items.length > 0) {
    files.push(['xl/sharedStrings.xml', strings.xml()]);
    overrides.push(['xl/sharedStrings.xml', `${TYPE_PREFIX}spreadsheetml.sharedStrings+xml`]);
    bookRels.add(`${NS_REL}/sharedStrings`, 'sharedStrings.xml');
  }

  let types = `${XML_DECLARATION}<Types xmlns="${NS_CONTENT_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`;
  if (written.some((sheet) => sheet.parts.comments)) types += `<Default Extension="vml" ContentType="${TYPE_PREFIX}vmlDrawing"/>`;
  for (const [extension, type] of mediaTypes) types += `<Default Extension="${extension}" ContentType="${type}"/>`;
  for (const [part, type] of overrides) types += `<Override PartName="/${part}" ContentType="${type}"/>`;
  types += '</Types>';

  const zip = new ZipArchive();
  const partNames: string[] = [];
  const put = (name: string, data: Uint8Array | string) => {
    zip.set(name, typeof data === 'string' ? encoder.encode(data) : data);
    partNames.push(`/${name}`);
  };
  put('[Content_Types].xml', types);
  put('_rels/.rels', `${XML_DECLARATION}<Relationships xmlns="${NS_PACKAGE_RELS}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  put('xl/workbook.xml', book);
  put('xl/_rels/workbook.xml.rels', bookRels.text());
  for (const [name, data] of files) put(name, data);

  notes.add('Text: a character XML cannot carry (in a formula, an error value, a name or another attribute) is written as U+FFFD', replacedCharacters);
  // Pivot tables (made here in a workbook that is in no xlsx file) are not written by this writer: their cells are.
  notes.add('Pivot table: a pivot table is not written; the cells it shows are', workbook.pivots ? workbook.pivots.tables.length : 0);
  if (options.report) {
    options.report.parts.push(...partNames);
    options.report.warnings.push(...notes.list());
  }
  return zip;
}
