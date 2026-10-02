// Reader for xl/styles.xml and the theme colour scheme.
//
// Elements are matched by local name, so a prefixed document (<x:styleSheet>)
// reads the same as the usual default-namespace one. The style part itself
// stays in the package and is written back from there; the model built here
// holds what the engine draws with.

import { childElements, localName } from '../../xml';
import type { XmlDocument, XmlElement } from '../../xml';
import { emptyStyleSheet } from '../../model/sheet/types';
import type {
  Alignment,
  Border,
  BorderEdge,
  CellFormat,
  Color,
  DifferentialFormat,
  Fill,
  Font,
  NamedCellStyle,
  Protection,
  StyleSheet,
  Underline,
} from '../../model/sheet/types';

/** Upper bound for every style table. Excel itself stops at about 64,000 cell formats. */
export const MAX_STYLE_RECORDS = 65536;

const UNDERLINES: readonly string[] = ['none', 'single', 'double', 'singleAccounting', 'doubleAccounting'];
const HORIZONTALS: readonly string[] = [
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
];
const VERTICALS: readonly string[] = ['top', 'center', 'bottom', 'justify', 'distributed'];

/** Theme colour names in the order spreadsheet theme indexes use: the two light/dark pairs are swapped. */
const THEME_ORDER: readonly string[] = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

const HEX6_OR_8 = /^[0-9A-F]{6}([0-9A-F]{2})?$/;

function firstChild(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const node of parent.children) {
    if (node.type === 'element' && localName(node.name) === local) return node;
  }
  return undefined;
}

/** Child records of a table element, refusing tables too large to be genuine. */
function records(table: XmlElement | undefined, local: string): XmlElement[] {
  if (!table) return [];
  const list = childElements(table, local);
  if (list.length > MAX_STYLE_RECORDS) {
    throw new Error(
      `Invalid styles: <${localName(table.name)}> has ${list.length} records, the limit is ${MAX_STYLE_RECORDS}`,
    );
  }
  return list;
}

/** Non-negative integer attribute; undefined when missing or malformed. */
function intAttr(el: XmlElement, name: string): number | undefined {
  const text = el.attrs.get(name);
  if (text === undefined || !/^\s*\d{1,15}\s*$/.test(text)) return undefined;
  return Number(text);
}

function numAttr(el: XmlElement, name: string): number | undefined {
  const text = el.attrs.get(name);
  // xsd:double as written in files; Number() alone would also take "0x10". The length is
  // checked first: on a long run of digits the expression backtracks for minutes.
  if (text === undefined || text.length > 64 || !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(text)) return undefined;
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

/** xsd:boolean attribute; undefined when missing or not a boolean. */
function boolAttr(el: XmlElement, name: string): boolean | undefined {
  const text = el.attrs.get(name);
  if (text === undefined) return undefined;
  const value = text.trim();
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return undefined;
}

/** <b/>, <i/>, <strike/>: present means on unless val says otherwise. */
const boolElement = (el: XmlElement) => boolAttr(el, 'val') ?? true;

/** An index into a table of `length` entries; missing or out of range reads as 0. */
function indexAttr(el: XmlElement, name: string, length: number): number {
  const value = intAttr(el, name);
  return value !== undefined && value < length ? value : 0;
}

/** Reads a CT_Color element (<color>, <fgColor>, <bgColor>, <tabColor>, ...). */
export function parseColor(el: XmlElement | undefined): Color | undefined {
  if (!el) return undefined;
  const color: Color = {};
  let any = false;
  const rgb = el.attrs.get('rgb')?.trim().toUpperCase();
  if (rgb !== undefined && HEX6_OR_8.test(rgb)) {
    color.rgb = rgb;
    any = true;
  }
  const theme = intAttr(el, 'theme');
  if (theme !== undefined) {
    color.theme = theme;
    any = true;
  }
  const indexed = intAttr(el, 'indexed');
  if (indexed !== undefined) {
    color.indexed = indexed;
    any = true;
  }
  const auto = boolAttr(el, 'auto');
  if (auto !== undefined) {
    color.auto = auto;
    any = true;
  }
  const tint = numAttr(el, 'tint');
  if (tint !== undefined) {
    color.tint = Math.max(-1, Math.min(1, tint));
    any = true;
  }
  return any ? color : undefined;
}

/**
 * Reads a <font> element. Rich text run properties (<rPr>) have the same
 * children except that the name is in <rFont>, which is accepted here too.
 */
export function parseFont(el: XmlElement): Font {
  const font: Font = { bold: false, italic: false, strike: false, underline: 'none' };
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    const val = child.attrs.get('val');
    switch (localName(child.name)) {
      case 'b':
        font.bold = boolElement(child);
        break;
      case 'i':
        font.italic = boolElement(child);
        break;
      case 'strike':
        font.strike = boolElement(child);
        break;
      case 'u':
        font.underline = val !== undefined && UNDERLINES.includes(val) ? (val as Underline) : 'single';
        break;
      case 'sz': {
        const size = numAttr(child, 'val');
        if (size !== undefined && size > 0) font.size = size;
        break;
      }
      case 'color': {
        const color = parseColor(child);
        if (color) font.color = color;
        break;
      }
      case 'name':
      case 'rFont':
        if (val !== undefined) font.name = val;
        break;
      case 'family': {
        const family = intAttr(child, 'val');
        if (family !== undefined) font.family = family;
        break;
      }
      case 'charset': {
        const charset = intAttr(child, 'val');
        if (charset !== undefined) font.charset = charset;
        break;
      }
      case 'scheme':
        if (val === 'major' || val === 'minor' || val === 'none') font.scheme = val;
        break;
      case 'vertAlign':
        if (val === 'superscript' || val === 'subscript') font.vertAlign = val;
        break;
    }
  }
  return font;
}

function parseFill(el: XmlElement): Fill {
  const gradient = firstChild(el, 'gradientFill');
  const pattern = firstChild(el, 'patternFill');
  if (gradient && !pattern) return { type: 'gradient', raw: gradient };
  const fill: Fill = { type: 'pattern', pattern: pattern?.attrs.get('patternType') || 'none' };
  const fg = parseColor(firstChild(pattern, 'fgColor'));
  const bg = parseColor(firstChild(pattern, 'bgColor'));
  if (fg) fill.fg = fg;
  if (bg) fill.bg = bg;
  return fill;
}

function parseEdge(el: XmlElement): BorderEdge | undefined {
  const style = el.attrs.get('style');
  if (!style || style === 'none') return undefined;
  const edge: BorderEdge = { style };
  const color = parseColor(firstChild(el, 'color'));
  if (color) edge.color = color;
  return edge;
}

function parseBorder(el: XmlElement): Border {
  const border: Border = {
    diagonalUp: boolAttr(el, 'diagonalUp') ?? false,
    diagonalDown: boolAttr(el, 'diagonalDown') ?? false,
  };
  for (const child of el.children) {
    if (child.type !== 'element') continue;
    let side = localName(child.name);
    // Older writers use the logical names of the first edition.
    if (side === 'start') side = 'left';
    else if (side === 'end') side = 'right';
    if (side !== 'left' && side !== 'right' && side !== 'top' && side !== 'bottom' && side !== 'diagonal') continue;
    const edge = parseEdge(child);
    if (edge && !border[side]) border[side] = edge;
  }
  return border;
}

function parseAlignment(el: XmlElement): Alignment {
  const alignment: Alignment = {};
  const horizontal = el.attrs.get('horizontal');
  if (horizontal !== undefined && HORIZONTALS.includes(horizontal)) {
    alignment.horizontal = horizontal as Alignment['horizontal'];
  }
  const vertical = el.attrs.get('vertical');
  if (vertical !== undefined && VERTICALS.includes(vertical)) alignment.vertical = vertical as Alignment['vertical'];
  const wrapText = boolAttr(el, 'wrapText');
  if (wrapText !== undefined) alignment.wrapText = wrapText;
  const shrinkToFit = boolAttr(el, 'shrinkToFit');
  if (shrinkToFit !== undefined) alignment.shrinkToFit = shrinkToFit;
  const textRotation = intAttr(el, 'textRotation');
  if (textRotation !== undefined && (textRotation <= 180 || textRotation === 255)) alignment.textRotation = textRotation;
  const indent = intAttr(el, 'indent');
  if (indent !== undefined) alignment.indent = indent;
  const readingOrder = intAttr(el, 'readingOrder');
  if (readingOrder !== undefined) alignment.readingOrder = readingOrder;
  return alignment;
}

function parseProtection(el: XmlElement): Protection {
  const protection: Protection = {};
  const locked = boolAttr(el, 'locked');
  if (locked !== undefined) protection.locked = locked;
  const hidden = boolAttr(el, 'hidden');
  if (hidden !== undefined) protection.hidden = hidden;
  return protection;
}

/**
 * Reads a <dxf> (also an <x14:dxf> of a worksheet's extension list). Only what the
 * element states ends up in the result: a font without <b> says nothing about bold.
 */
export function parseDifferentialFormat(el: XmlElement): DifferentialFormat {
  const format: DifferentialFormat = {};
  const fontEl = firstChild(el, 'font');
  if (fontEl) {
    const font: NonNullable<DifferentialFormat['font']> = {};
    for (const child of fontEl.children) {
      if (child.type !== 'element') continue;
      switch (localName(child.name)) {
        case 'b':
          font.bold = boolElement(child);
          break;
        case 'i':
          font.italic = boolElement(child);
          break;
        case 'strike':
          font.strike = boolElement(child);
          break;
        case 'u': {
          const val = child.attrs.get('val');
          font.underline = val !== undefined && UNDERLINES.includes(val) ? (val as Underline) : 'single';
          break;
        }
        case 'color': {
          const color = parseColor(child);
          if (color) font.color = color;
          break;
        }
      }
    }
    format.font = font;
  }
  const fillEl = firstChild(el, 'fill');
  if (fillEl) {
    const fill = parseFill(fillEl);
    // In a differential format a pattern fill that names no pattern is solid.
    const pattern = firstChild(fillEl, 'patternFill');
    if (fill.type === 'pattern' && pattern && !pattern.attrs.get('patternType')) fill.pattern = 'solid';
    format.fill = fill;
  }
  const borderEl = firstChild(el, 'border');
  if (borderEl) {
    const border = parseBorder(borderEl);
    const edges: NonNullable<DifferentialFormat['border']> = {};
    if (border.left) edges.left = border.left;
    if (border.right) edges.right = border.right;
    if (border.top) edges.top = border.top;
    if (border.bottom) edges.bottom = border.bottom;
    format.border = edges;
  }
  const numFmt = firstChild(el, 'numFmt');
  if (numFmt) {
    const id = intAttr(numFmt, 'numFmtId');
    const code = numFmt.attrs.get('formatCode');
    if (code !== undefined) format.numFmt = { id: id === undefined ? -1 : id, code };
  }
  return format;
}

interface TableSizes {
  fonts: number;
  fills: number;
  borders: number;
  /** Size of cellStyleXfs when reading cellXfs; undefined when reading cellStyleXfs itself. */
  styleXfs?: number;
}

function parseXf(el: XmlElement, sizes: TableSizes): CellFormat {
  const format: CellFormat = {
    // A number format id is a key, not a table index, so it is not clamped.
    numFmtId: intAttr(el, 'numFmtId') ?? 0,
    fontId: indexAttr(el, 'fontId', sizes.fonts),
    fillId: indexAttr(el, 'fillId', sizes.fills),
    borderId: indexAttr(el, 'borderId', sizes.borders),
    quotePrefix: boolAttr(el, 'quotePrefix') ?? false,
  };
  if (sizes.styleXfs !== undefined) format.xfId = indexAttr(el, 'xfId', sizes.styleXfs);
  const alignment = firstChild(el, 'alignment');
  if (alignment) format.alignment = parseAlignment(alignment);
  const protection = firstChild(el, 'protection');
  if (protection) format.protection = parseProtection(protection);
  return format;
}

/**
 * Theme colours ("RRGGBB") in the order spreadsheet theme indexes use.
 *
 * The list has twelve entries when the theme has a colour scheme and is empty
 * otherwise. A scheme colour that is missing or not an RGB/system colour is
 * an empty string, which resolveColor() treats as unresolvable.
 */
export function parseThemeColors(theme: XmlDocument): string[] {
  const scheme = firstChild(firstChild(theme.root, 'themeElements'), 'clrScheme');
  if (!scheme) return [];
  return THEME_ORDER.map((name) => {
    const slot = firstChild(scheme, name);
    const srgb = firstChild(slot, 'srgbClr');
    const sys = firstChild(slot, 'sysClr');
    let value = srgb ? srgb.attrs.get('val') : sys?.attrs.get('lastClr');
    if (value === undefined && sys) {
      // lastClr is optional; the two system colours themes actually use have obvious values.
      const system = sys.attrs.get('val');
      value = system === 'windowText' ? '000000' : system === 'window' ? 'FFFFFF' : undefined;
    }
    value = value?.trim().toUpperCase();
    return value !== undefined && /^[0-9A-F]{6}$/.test(value) ? value : '';
  });
}

/** Reads xl/styles.xml; `theme` is the parsed theme part, when the package has one. Either may be missing. */
export function parseStyles(styles: XmlDocument | undefined, theme: XmlDocument | undefined): StyleSheet {
  const sheet = emptyStyleSheet();
  if (theme) sheet.themeColors = parseThemeColors(theme);
  const root = styles?.root;
  if (!root) return sheet;

  for (const numFmt of records(firstChild(root, 'numFmts'), 'numFmt')) {
    const id = intAttr(numFmt, 'numFmtId');
    const code = numFmt.attrs.get('formatCode');
    if (id !== undefined && code !== undefined) sheet.numFmts.set(id, code);
  }

  // Each table keeps the default entry of emptyStyleSheet() when the file has none.
  const fonts = records(firstChild(root, 'fonts'), 'font').map(parseFont);
  if (fonts.length) sheet.fonts = fonts;
  const fills = records(firstChild(root, 'fills'), 'fill').map(parseFill);
  if (fills.length) sheet.fills = fills;
  const borders = records(firstChild(root, 'borders'), 'border').map(parseBorder);
  if (borders.length) sheet.borders = borders;

  const sizes: TableSizes = { fonts: sheet.fonts.length, fills: sheet.fills.length, borders: sheet.borders.length };
  const styleXfs = records(firstChild(root, 'cellStyleXfs'), 'xf').map((xf) => parseXf(xf, sizes));
  if (styleXfs.length) sheet.cellStyleXfs = styleXfs;
  const cellSizes: TableSizes = { ...sizes, styleXfs: sheet.cellStyleXfs.length };
  const cellXfs = records(firstChild(root, 'cellXfs'), 'xf').map((xf) => parseXf(xf, cellSizes));
  if (cellXfs.length) sheet.cellXfs = cellXfs;

  sheet.cellStyles = records(firstChild(root, 'cellStyles'), 'cellStyle').map((el) => {
    const style: NamedCellStyle = {
      name: el.attrs.get('name') ?? '',
      xfId: indexAttr(el, 'xfId', sheet.cellStyleXfs.length),
    };
    const builtinId = intAttr(el, 'builtinId');
    if (builtinId !== undefined) style.builtinId = builtinId;
    return style;
  });

  sheet.dxfs = records(firstChild(root, 'dxfs'), 'dxf');
  if (sheet.dxfs.length > 0) sheet.dxfFormats = sheet.dxfs.map(parseDifferentialFormat);

  const palette = records(firstChild(firstChild(root, 'colors'), 'indexedColors'), 'rgbColor');
  if (palette.length) {
    sheet.indexedColors = palette.map((el) => {
      const rgb = el.attrs.get('rgb')?.trim().toUpperCase();
      // Positions matter, so a malformed entry still takes its slot.
      return rgb !== undefined && HEX6_OR_8.test(rgb) ? rgb.slice(-6) : '000000';
    });
  }

  return sheet;
}
