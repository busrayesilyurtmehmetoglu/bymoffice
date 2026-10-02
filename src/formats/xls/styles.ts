// FONT, FORMAT, XF, XFEXT, STYLE, PALETTE and THEME records -> the style sheet
// of the model, laid out the way the xlsx reader lays it out: cell formats and
// style formats in separate lists, fills and borders in tables of their own.

import { inflateRaw } from '../../deflate';
import { DEFAULT_INDEXED_COLORS } from '../../model/sheet/color';
import { emptyStyleSheet } from '../../model/sheet/types';
import type { Alignment, Border, BorderEdge, CellFormat, Color, Fill, Font, NamedCellStyle, StyleSheet, Underline } from '../../model/sheet/types';
import { parseXml } from '../../xml';
import { parseThemeColors } from '../xlsx/styles';
import type { RecordReader } from './records';
import { readText16, readText8 } from './strings';

/** Upper bound for every style table, as in the xlsx reader. */
const MAX_STYLE_RECORDS = 65536;
/** A theme is a few kilobytes of XML; this bounds what a hostile archive can expand to. */
const MAX_THEME_BYTES = 4 * 1024 * 1024;

const HORIZONTALS: readonly NonNullable<Alignment['horizontal']>[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const VERTICALS: readonly NonNullable<Alignment['vertical']>[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
const BORDER_STYLES: readonly string[] = [
  'none',
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];
const PATTERNS: readonly string[] = [
  'none',
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
];

/** Names of the built-in cell styles a STYLE record refers to by number. */
const BUILTIN_STYLES: readonly string[] = ['Normal', 'RowLevel_', 'ColLevel_', 'Comma', 'Currency', 'Percent', 'Comma [0]', 'Currency [0]', 'Hyperlink', 'Followed Hyperlink'];

/** The colours of the Office theme, in the order theme indexes use; for files that name theme colours without carrying a theme. */
const OFFICE_THEME: readonly string[] = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];

/** Colour index that stands for "automatic" in a FONT record. */
const FONT_COLOR_AUTO = 0x7fff;

interface RawXf {
  font: number;
  format: number;
  style: boolean;
  parent: number;
  locked: boolean;
  hidden: boolean;
  quotePrefix: boolean;
  horizontal: number;
  vertical: number;
  wrap: boolean;
  rotation: number;
  indent: number;
  shrink: boolean;
  readingOrder: number;
  /** Line styles and colour indexes: left, right, top, bottom, diagonal. */
  lines: number[];
  lineColors: number[];
  diagonalDown: boolean;
  diagonalUp: boolean;
  pattern: number;
  fg: number;
  bg: number;
}

/** Colours an XFEXT record states more exactly than the palette index of its XF does. */
interface XfExtension {
  fg?: Color;
  bg?: Color;
  /** left, right, top, bottom, diagonal. */
  lines: (Color | undefined)[];
  text?: Color;
  indent?: number;
}

const hex2 = (value: number) => (value < 16 ? '0' : '') + value.toString(16).toUpperCase();

/** Finds a member of a ZIP archive by the end of its name and returns its content; undefined when there is none. */
function unzipMember(zip: Uint8Array, suffix: RegExp): Uint8Array | undefined {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  // The end-of-central-directory record is at the end, before an optional comment.
  let eocd = -1;
  for (let i = zip.length - 22; i >= 0 && i >= zip.length - 22 - 0xffff; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return undefined;
  const count = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true);
  for (let n = 0; n < count; n++) {
    if (p + 46 > zip.length || view.getUint32(p, true) !== 0x02014b50) return undefined;
    const method = view.getUint16(p + 10, true);
    const compressedSize = view.getUint32(p + 20, true);
    const size = view.getUint32(p + 24, true);
    const nameLength = view.getUint16(p + 28, true);
    const local = view.getUint32(p + 42, true);
    if (p + 46 + nameLength > zip.length) return undefined;
    let name = '';
    for (let i = 0; i < nameLength; i++) name += String.fromCharCode(zip[p + 46 + i]);
    if (suffix.test(name)) {
      if (local + 30 > zip.length || view.getUint32(local, true) !== 0x04034b50) return undefined;
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      if (start + compressedSize > zip.length || size > MAX_THEME_BYTES) return undefined;
      const data = zip.subarray(start, start + compressedSize);
      if (method === 0) return data;
      if (method === 8) return inflateRaw(data, MAX_THEME_BYTES);
      return undefined;
    }
    p += 46 + nameLength + view.getUint16(p + 30, true) + view.getUint16(p + 32, true);
  }
  return undefined;
}

/** Collects the style records of the workbook globals and builds the style sheet from them. */
export class XlsStyles {
  private readonly fonts: Font[] = [];
  private readonly formats = new Map<number, string>();
  private readonly xfs: RawXf[] = [];
  private readonly extensions = new Map<number, XfExtension>();
  private readonly styleRecords: { xf: number; name: string; builtinId?: number }[] = [];
  private palette: string[] | undefined;
  private themeColors: string[] | undefined;
  private usesTheme = false;

  constructor(
    private readonly biff8: boolean,
    private readonly codepage: () => number,
  ) {}

  /** FONT. */
  addFont(r: RecordReader): void {
    if (this.fonts.length >= MAX_STYLE_RECORDS) throw new Error(`Invalid xls: more than ${MAX_STYLE_RECORDS} FONT records`);
    const height = r.u16();
    const flags = r.u16();
    const colorIndex = r.u16();
    const weight = r.u16();
    const script = r.u16();
    const underline = r.u8();
    const family = r.u8();
    const charset = r.u8();
    r.skip(1);
    const font: Font = {
      name: readText8(r, this.biff8, this.codepage()),
      bold: weight >= 600,
      italic: (flags & 0x02) !== 0,
      strike: (flags & 0x08) !== 0,
      underline: ({ 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' } as Record<number, Underline>)[underline] || 'none',
    };
    if (height > 0) font.size = height / 20;
    if (colorIndex !== FONT_COLOR_AUTO) font.color = { indexed: colorIndex };
    if (script === 1) font.vertAlign = 'superscript';
    else if (script === 2) font.vertAlign = 'subscript';
    if (family !== 0) font.family = family;
    if (charset !== 0) font.charset = charset;
    this.fonts.push(font);
  }

  /** FORMAT: a number format code and the id XF records refer to it by. */
  addFormat(r: RecordReader): void {
    if (this.formats.size >= MAX_STYLE_RECORDS) throw new Error(`Invalid xls: more than ${MAX_STYLE_RECORDS} FORMAT records`);
    const id = r.u16();
    const code = this.biff8 ? readText16(r, true, 0).text : readText8(r, false, this.codepage());
    this.formats.set(id, code);
  }

  /** XF. */
  addXf(r: RecordReader): void {
    if (this.xfs.length >= MAX_STYLE_RECORDS) throw new Error(`Invalid xls: more than ${MAX_STYLE_RECORDS} XF records`);
    const font = r.u16();
    const format = r.u16();
    const type = r.u16();
    const align = r.u8();
    const xf: RawXf = {
      font,
      format,
      locked: (type & 0x01) !== 0,
      hidden: (type & 0x02) !== 0,
      style: (type & 0x04) !== 0,
      quotePrefix: (type & 0x08) !== 0,
      parent: type >> 4,
      horizontal: align & 0x07,
      wrap: (align & 0x08) !== 0,
      vertical: (align >> 4) & 0x07,
      rotation: 0,
      indent: 0,
      shrink: false,
      readingOrder: 0,
      lines: [0, 0, 0, 0, 0],
      lineColors: [0, 0, 0, 0, 0],
      diagonalDown: false,
      diagonalUp: false,
      pattern: 0,
      fg: 64,
      bg: 65,
    };
    if (this.biff8) {
      xf.rotation = r.u8();
      const indent = r.u8();
      xf.indent = indent & 0x0f;
      xf.shrink = (indent & 0x10) !== 0;
      xf.readingOrder = indent >> 6;
      r.skip(1);
      const border1 = r.u32();
      const border2 = r.u32();
      const colors = r.u16();
      xf.lines = [border1 & 0x0f, (border1 >> 4) & 0x0f, (border1 >> 8) & 0x0f, (border1 >> 12) & 0x0f, (border2 >> 21) & 0x0f];
      xf.lineColors = [(border1 >> 16) & 0x7f, (border1 >> 23) & 0x7f, border2 & 0x7f, (border2 >> 7) & 0x7f, (border2 >> 14) & 0x7f];
      xf.diagonalDown = (border1 & 0x40000000) !== 0;
      xf.diagonalUp = (border1 & 0x80000000) !== 0;
      xf.pattern = border2 >>> 26;
      xf.fg = colors & 0x7f;
      xf.bg = (colors >> 7) & 0x7f;
    } else {
      // BIFF5: orientation instead of an angle, and three bits per line style.
      const orientation = r.u8() & 0x03;
      xf.rotation = orientation === 1 ? 255 : orientation === 2 ? 90 : orientation === 3 ? 180 : 0;
      const fill = r.u32();
      const border = r.u32();
      xf.fg = fill & 0x7f;
      xf.bg = (fill >> 7) & 0x7f;
      xf.pattern = (fill >> 16) & 0x3f;
      xf.lines = [(border >> 3) & 0x07, (border >> 6) & 0x07, border & 0x07, (fill >> 22) & 0x07, 0];
      xf.lineColors = [(border >> 16) & 0x7f, (border >> 23) & 0x7f, (border >> 9) & 0x7f, fill >>> 25, 0];
      // Without a vertical alignment for "justify" and "distributed", 3 is the last value.
      if (xf.vertical > 3) xf.vertical = 2;
    }
    this.xfs.push(xf);
  }

  /** A colour of an XFEXT record (FullColorExt). */
  private readFullColor(r: RecordReader): Color | undefined {
    const type = r.u16();
    const tintValue = r.i16();
    const value = r.u32();
    r.skip(8);
    let color: Color | undefined;
    if (type === 1) color = { indexed: value & 0xffff };
    else if (type === 2) color = { rgb: 'FF' + hex2(value & 0xff) + hex2((value >> 8) & 0xff) + hex2((value >> 16) & 0xff) };
    else if (type === 3) {
      color = { theme: value & 0xffff };
      this.usesTheme = true;
    }
    if (color && tintValue !== 0) color.tint = Math.max(-1, Math.min(1, tintValue / 32767));
    return color;
  }

  /** XFEXT: colours of an XF that the palette cannot state (written by Excel 2007 and later). */
  addXfExt(r: RecordReader): void {
    if (r.remaining < 20) return;
    r.skip(14);
    const index = r.u16();
    r.skip(2);
    const count = r.u16();
    // A second record for the same XF adds to the first.
    const extension: XfExtension = this.extensions.get(index) || { lines: [] };
    for (let i = 0; i < count && !r.done; i++) {
      const type = r.u16();
      const size = r.u16();
      const body = size - 4;
      // A property that says it is longer than what is left ends the list; what was read stays.
      if (size < 4 || body > r.remaining) break;
      if ((type === 4 || type === 5 || (type >= 7 && type <= 11) || type === 13) && body >= 16) {
        const color = this.readFullColor(r);
        if (body > 16) r.skipUpTo(body - 16);
        if (!color) continue;
        if (type === 4) extension.fg = color;
        else if (type === 5) extension.bg = color;
        else if (type === 13) extension.text = color;
        // Top, bottom, left, right, diagonal in the file; left, right, top, bottom, diagonal here.
        else extension.lines[[2, 3, 0, 1, 4][type - 7]] = color;
      } else if (type === 15 && body >= 2) {
        extension.indent = r.u16();
        if (body > 2) r.skipUpTo(body - 2);
      } else r.skipUpTo(body);
    }
    this.extensions.set(index, extension);
  }

  /** STYLE: names a style XF. */
  addStyle(r: RecordReader): void {
    if (this.styleRecords.length >= MAX_STYLE_RECORDS) return;
    const field = r.u16();
    const xf = field & 0x0fff;
    if (field & 0x8000) {
      const builtinId = r.u8();
      const level = r.u8();
      const base = BUILTIN_STYLES[builtinId];
      if (base === undefined) return;
      this.styleRecords.push({ xf, builtinId, name: builtinId === 1 || builtinId === 2 ? base + (level + 1) : base });
    } else {
      this.styleRecords.push({ xf, name: this.biff8 ? readText16(r, true, 0).text : readText8(r, false, this.codepage()) });
    }
  }

  /** PALETTE: the colours of indexes 8 and up. */
  addPalette(r: RecordReader): void {
    const count = Math.min(r.u16(), 56);
    const palette = DEFAULT_INDEXED_COLORS.slice(0, 64);
    for (let i = 0; i < count && r.remaining >= 4; i++) {
      const red = r.u8();
      const green = r.u8();
      const blue = r.u8();
      r.skip(1);
      palette[8 + i] = hex2(red) + hex2(green) + hex2(blue);
    }
    this.palette = palette;
  }

  /** THEME: a version number and, for a custom theme, a ZIP archive with the theme part in it. */
  addTheme(r: RecordReader): void {
    r.skip(12);
    const version = r.u32();
    if (version !== 0 || r.remaining === 0) {
      this.themeColors = OFFICE_THEME.slice();
      return;
    }
    if (r.remaining > MAX_THEME_BYTES) return;
    try {
      const part = unzipMember(r.rest(), /theme\/theme\d*\.xml$/i);
      const colors = part ? parseThemeColors(parseXml(part)) : [];
      if (colors.length > 0) this.themeColors = colors;
    } catch {
      // A theme that cannot be read leaves the colours of the Office theme, as a missing one does.
    }
  }

  /** The model's font for a font index of the file. The index 4 is never used: indexes above it are one too high. */
  fontIndex(index: number): number {
    const at = index > 4 ? index - 1 : index;
    return at < this.fonts.length && index !== 4 ? at : 0;
  }

  /**
   * The style sheet, and for every XF index of the file the index of its cell format in
   * `cellXfs`. `xfMap` has one entry more than there are XF records: the last one is what an
   * index outside the table reads as.
   */
  build(): { styles: StyleSheet; xfMap: number[] } {
    const sheet = emptyStyleSheet();
    if (this.fonts.length > 0) sheet.fonts = this.fonts;
    else if (this.xfs.length > 0) sheet.fonts = [{ name: 'Arial', size: 10, bold: false, italic: false, strike: false, underline: 'none' }];
    sheet.numFmts = this.formats;
    if (this.palette) sheet.indexedColors = this.palette;
    if (this.themeColors) sheet.themeColors = this.themeColors;
    else if (this.usesTheme) sheet.themeColors = OFFICE_THEME.slice();
    if (this.xfs.length === 0) return { styles: sheet, xfMap: [0] };

    // The two tables every xlsx style sheet starts with.
    const fills: Fill[] = [{ type: 'pattern', pattern: 'none' }, { type: 'pattern', pattern: 'gray125' }];
    const fillIds = new Map<string, number>([[JSON.stringify(fills[0]), 0], [JSON.stringify(fills[1]), 1]]);
    const borders: Border[] = [{ diagonalUp: false, diagonalDown: false }];
    const borderIds = new Map<string, number>([[JSON.stringify(borders[0]), 0]]);
    const tintedFonts = new Map<string, number>();

    const convert = (xf: RawXf, index: number): CellFormat => {
      const extension = this.extensions.get(index);
      let fontId = this.fontIndex(xf.font);
      if (extension && extension.text) {
        // The colour belongs to this XF only; the font it shares with others gets a copy.
        const key = fontId + '|' + JSON.stringify(extension.text);
        let copy = tintedFonts.get(key);
        if (copy === undefined) {
          copy = sheet.fonts.length;
          sheet.fonts.push({ ...sheet.fonts[fontId], color: extension.text });
          tintedFonts.set(key, copy);
        }
        fontId = copy;
      }

      let fillId = 0;
      if (xf.pattern !== 0) {
        const fill: Fill = { type: 'pattern', pattern: PATTERNS[xf.pattern] || 'solid', fg: (extension && extension.fg) || { indexed: xf.fg }, bg: (extension && extension.bg) || { indexed: xf.bg } };
        const key = JSON.stringify(fill);
        const known = fillIds.get(key);
        if (known !== undefined) fillId = known;
        else {
          fillId = fills.length;
          fills.push(fill);
          fillIds.set(key, fillId);
        }
      }

      const edge = (side: number): BorderEdge | undefined => {
        const style = BORDER_STYLES[xf.lines[side]];
        if (style === undefined || style === 'none') return undefined;
        return { style, color: (extension && extension.lines[side]) || { indexed: xf.lineColors[side] } };
      };
      const border: Border = { diagonalUp: xf.diagonalUp, diagonalDown: xf.diagonalDown };
      const left = edge(0);
      const right = edge(1);
      const top = edge(2);
      const bottom = edge(3);
      const diagonal = xf.diagonalUp || xf.diagonalDown ? edge(4) : undefined;
      if (left) border.left = left;
      if (right) border.right = right;
      if (top) border.top = top;
      if (bottom) border.bottom = bottom;
      if (diagonal) border.diagonal = diagonal;
      const borderKey = JSON.stringify(border);
      let borderId = borderIds.get(borderKey);
      if (borderId === undefined) {
        borderId = borders.length;
        borders.push(border);
        borderIds.set(borderKey, borderId);
      }

      const format: CellFormat = { numFmtId: xf.format, fontId, fillId, borderId, quotePrefix: xf.quotePrefix };
      const alignment: Alignment = {};
      if (xf.horizontal !== 0) alignment.horizontal = HORIZONTALS[xf.horizontal];
      // Bottom is what a cell has when nothing is said.
      if (xf.vertical !== 2 && VERTICALS[xf.vertical]) alignment.vertical = VERTICALS[xf.vertical];
      if (xf.wrap) alignment.wrapText = true;
      if (xf.shrink) alignment.shrinkToFit = true;
      if (xf.rotation !== 0 && (xf.rotation <= 180 || xf.rotation === 255)) alignment.textRotation = xf.rotation;
      const indent = extension && extension.indent !== undefined ? extension.indent : xf.indent;
      if (indent !== 0) alignment.indent = indent;
      if (xf.readingOrder !== 0) alignment.readingOrder = xf.readingOrder;
      if (Object.keys(alignment).length > 0) format.alignment = alignment;
      // A cell is locked and its formula shown unless the format says otherwise.
      if (!xf.locked || xf.hidden) {
        format.protection = {};
        if (!xf.locked) format.protection.locked = false;
        if (xf.hidden) format.protection.hidden = true;
      }
      return format;
    };

    // Style XFs and cell XFs share one table in the file; the model has one list for each.
    const styleIds: number[] = [];
    const cellStyleXfs: CellFormat[] = [];
    const cellXfs: CellFormat[] = [];
    const xfMap: number[] = [];
    this.xfs.forEach((xf, index) => {
      if (xf.style) {
        styleIds[index] = cellStyleXfs.length;
        cellStyleXfs.push(convert(xf, index));
      }
    });
    if (cellStyleXfs.length === 0) cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false });
    this.xfs.forEach((xf, index) => {
      if (xf.style) return;
      xfMap[index] = cellXfs.length;
      cellXfs.push({ ...convert(xf, index), xfId: styleIds[xf.parent] || 0 });
    });
    // A cell that names a style XF as its format gets a cell format with the same properties.
    this.xfs.forEach((xf, index) => {
      if (!xf.style) return;
      xfMap[index] = cellXfs.length;
      cellXfs.push({ ...cellStyleXfs[styleIds[index]], xfId: styleIds[index] });
    });
    // An index outside the table reads as the default cell format: the first one.
    xfMap.push(0);

    sheet.fills = fills;
    sheet.borders = borders;
    sheet.cellStyleXfs = cellStyleXfs;
    sheet.cellXfs = cellXfs;
    const named: NamedCellStyle[] = [];
    for (const record of this.styleRecords) {
      const xfId = styleIds[record.xf];
      if (xfId === undefined) continue;
      const style: NamedCellStyle = { name: record.name, xfId };
      if (record.builtinId !== undefined) style.builtinId = record.builtinId;
      named.push(style);
    }
    sheet.cellStyles = named;
    return { styles: sheet, xfMap };
  }
}
