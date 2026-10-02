// The font table (SttbfFfn, MS-DOC 2.9.286) and the style sheet (STSH,
// MS-DOC 2.9.271) -> the fonts and styles of the document model.
//
// A style of the file is known by its index (istd). The model knows styles by
// an id, as a docx file does; the ids made here are the ones Word itself
// gives the built-in styles ("Normal", "Heading1", ...) and the name without
// its spaces for the others, so that code written for docx finds what it
// expects.

import {
  emptyDocStyles,
  type DocStyle,
  type DocStyles,
  type FontInfo,
  type RunProps,
  type StyleType,
} from '../../model/doc/types';
import { u16, u8, utf16 } from './bytes';
import { charProps, paraProps, type PropContext, type ToggleName } from './props';
import { readTap } from './tables';

export interface FontTable {
  /** Font names by index (ftc). */
  names: string[];
  fonts: Map<string, FontInfo>;
}

/** More fonts than this in one document is not a genuine file. */
const MAX_FONTS = 10000;
const FAMILIES = ['auto', 'roman', 'swiss', 'modern', 'script', 'decorative'];
const PITCHES = ['default', 'fixed', 'variable'];
const HEX = '0123456789ABCDEF';

export function readFonts(table: Uint8Array, fc: number, lcb: number): FontTable {
  const out: FontTable = { names: [], fonts: new Map() };
  const end = fc + lcb;
  if (lcb < 4 || fc < 0 || end > table.length) return out;
  const count = Math.min(u16(table, fc), MAX_FONTS);
  let p = fc + 4;
  for (let i = 0; i < count && p < end; i++) {
    // An FFN: its size less one, then 39 bytes of facts, then the name and an alternative name, each ended by a zero.
    const size = u8(table, p);
    const start = p + 1;
    const stop = Math.min(end, start + size);
    p = start + size;
    if (stop - start < 41) {
      out.names.push('');
      continue;
    }
    const chars = Math.floor((stop - start - 39) / 2);
    const text = utf16(table, start + 39, chars);
    const zero = text.indexOf('\u0000');
    const name = zero < 0 ? text : text.slice(0, zero);
    out.names.push(name);
    if (!name || out.fonts.has(name)) continue;
    const info: FontInfo = { name };
    const altAt = u8(table, start + 4);
    if (altAt > 0 && altAt < text.length) {
      const rest = text.slice(altAt);
      const stopAt = rest.indexOf('\u0000');
      const alt = stopAt < 0 ? rest : rest.slice(0, stopAt);
      if (alt) info.altName = alt;
    }
    const flags = u8(table, start);
    const family = FAMILIES[(flags >> 4) & 7];
    if (family !== undefined) info.family = family;
    const pitch = PITCHES[flags & 3];
    if (pitch !== undefined) info.pitch = pitch;
    const charset = u8(table, start + 3);
    info.charset = HEX.charAt(charset >> 4) + HEX.charAt(charset & 15);
    let panose = '';
    for (let k = 0; k < 10; k++) panose += HEX.charAt(table[start + 5 + k] >> 4) + HEX.charAt(table[start + 5 + k] & 15);
    info.panose = panose;
    out.fonts.set(name, info);
  }
  return out;
}

/** Ids of the built-in styles Word gives an id that is not just the name without spaces, by "sti". */
const BUILT_IN_IDS: Record<number, string> = {
  0: 'Normal',
  1: 'Heading1',
  2: 'Heading2',
  3: 'Heading3',
  4: 'Heading4',
  5: 'Heading5',
  6: 'Heading6',
  7: 'Heading7',
  8: 'Heading8',
  9: 'Heading9',
  19: 'TOC1',
  20: 'TOC2',
  21: 'TOC3',
  22: 'TOC4',
  23: 'TOC5',
  24: 'TOC6',
  25: 'TOC7',
  26: 'TOC8',
  27: 'TOC9',
  29: 'FootnoteText',
  30: 'CommentText',
  31: 'Header',
  32: 'Footer',
  34: 'Caption',
  38: 'FootnoteReference',
  39: 'CommentReference',
  41: 'PageNumber',
  42: 'EndnoteReference',
  43: 'EndnoteText',
  62: 'Title',
  65: 'DefaultParagraphFont',
  66: 'BodyText',
  74: 'Subtitle',
  85: 'Hyperlink',
  86: 'FollowedHyperlink',
  87: 'Strong',
  88: 'Emphasis',
  105: 'TableNormal',
  107: 'NoList',
  153: 'BalloonText',
  154: 'TableGrid',
};

const STI_USER = 0x0ffe;
const ISTD_NIL = 0x0fff;
const STYLE_TYPES: (StyleType | undefined)[] = [undefined, 'paragraph', 'character', 'table', 'numbering'];
/** The default style of each type, by "sti": Normal, Default Paragraph Font, Normal Table, No List. */
const DEFAULT_STI: Record<StyleType, number> = { paragraph: 0, character: 65, table: 105, numbering: 107 };

interface RawStyle {
  istd: number;
  style: DocStyle;
  base: number;
  next: number;
  link: number;
  /** [start, end) of each UPX in the table stream. */
  upx: number[];
  invert?: ToggleName[];
}

export interface StyleSheet {
  styles: DocStyles;
  /** Style id by index in the file. */
  ids: (string | undefined)[];
  /** Index in the file by style id. */
  byId: Map<string, number>;
  /** Fonts the file names as its defaults: for Latin text, East Asian text, other text and complex scripts. */
  defaultFonts: [ascii: number, eastAsia: number, other: number, complex: number | undefined] | undefined;
  /**
   * Character styles that state a toggle as "the opposite of what is below" (operand 0x81), by style id. In the
   * model such a toggle is `true`, which is what a docx file means by it; every other toggle of a character
   * style is absolute in this format (see DocReader.runProps).
   */
  flips?: Map<string, ToggleName[]>;
}

function makeId(sti: number, name: string, istd: number, used: Set<string>): string {
  let id = sti !== STI_USER ? BUILT_IN_IDS[sti] : undefined;
  if (id === undefined) {
    id = '';
    for (let i = 0; i < name.length && id.length < 64; i++) {
      const c = name.charCodeAt(i);
      if ((c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c > 127) id += name.charAt(i);
    }
    if (!id) id = 'Style' + istd;
  }
  // Two styles never share an id, whatever their names are.
  if (used.has(id)) id = id + '_' + istd;
  used.add(id);
  return id;
}

/**
 * Reads the names and relations of the styles. The formatting is read in a
 * second step (finishStyles), because it refers to styles and fonts by index
 * and the property reader needs the ids made here.
 */
export function readStyleSheet(table: Uint8Array, fc: number, lcb: number): { sheet: StyleSheet; raw: RawStyle[] } {
  const sheet: StyleSheet = { styles: emptyDocStyles(), ids: [], byId: new Map(), defaultFonts: undefined };
  const raw: RawStyle[] = [];
  const end = fc + lcb;
  if (lcb < 4 || fc < 0 || end > table.length) return { sheet, raw };
  const headerSize = u16(table, fc);
  const header = fc + 2;
  if (header + headerSize > end || headerSize < 4) return { sheet, raw };
  // An istd has twelve bits, the last of which means "none".
  const count = Math.min(u16(table, header), ISTD_NIL);
  const baseSize = u16(table, header + 2);
  if (headerSize >= 18) {
    sheet.defaultFonts = [u16(table, header + 12), u16(table, header + 14), u16(table, header + 16), headerSize >= 20 ? u16(table, header + 18) : undefined];
  }
  const used = new Set<string>();
  let p = header + headerSize;
  for (let istd = 0; istd < count && p + 2 <= end; istd++) {
    const size = u16(table, p);
    const base = p + 2;
    p = base + size;
    sheet.ids.push(undefined);
    if (size === 0) continue;
    if (p > end || size < baseSize + 2 || baseSize < 10) continue;
    const word0 = u16(table, base);
    const word1 = u16(table, base + 2);
    const word2 = u16(table, base + 4);
    const flags = u16(table, base + 8);
    const type = STYLE_TYPES[word1 & 15];
    if (type === undefined) continue;
    const sti = word0 & 0x0fff;
    const nameLength = u16(table, base + baseSize);
    const nameAt = base + baseSize + 2;
    if (nameAt + nameLength * 2 > p) continue;
    const name = utf16(table, nameAt, nameLength);
    const id = makeId(sti, name, istd, used);
    const style: DocStyle = { id, type, name, isDefault: false };
    if (sti === STI_USER) style.custom = true;
    if (flags & 0x0002) style.hidden = true;
    if (flags & 0x0100) style.semiHidden = true;
    if (flags & 0x0200) style.locked = true;
    if (flags & 0x0800) style.unhideWhenUsed = true;
    if (flags & 0x1000) style.quickFormat = true;
    let link = ISTD_NIL;
    if (baseSize >= 18) {
      link = u16(table, base + 10) & 0x0fff;
      style.uiPriority = u16(table, base + 16) >> 4;
    }
    if (sti === DEFAULT_STI[type] && sheet.styles.defaultIds[type] === undefined) {
      style.isDefault = true;
      sheet.styles.defaultIds[type] = id;
    }
    // The property lists: each starts at an even offset from the start of the style.
    const upx: number[] = [];
    let q = nameAt + nameLength * 2 + 2;
    const lists = word2 & 15;
    for (let k = 0; k < lists; k++) {
      if ((q - base) & 1) q++;
      if (q + 2 > p) break;
      const length = u16(table, q);
      q += 2;
      if (q + length > p) break;
      upx.push(q, q + length);
      q += length;
    }
    sheet.ids[istd] = id;
    sheet.byId.set(id, istd);
    sheet.styles.styles.set(id, style);
    raw.push({ istd, style, base: word1 >> 4, next: word2 >> 4, link, upx });
  }
  // The style at index 0 is the paragraph style everything falls back to, whatever it calls itself.
  if (sheet.styles.defaultIds.paragraph === undefined && sheet.ids[0] !== undefined) {
    const first = sheet.styles.styles.get(sheet.ids[0])!;
    if (first.type === 'paragraph') {
      first.isDefault = true;
      sheet.styles.defaultIds.paragraph = first.id;
    }
  }
  return { sheet, raw };
}

/** Reads the formatting of the styles and what Word assumes for text no style says anything about. */
export function finishStyles(ctx: PropContext, table: Uint8Array, sheet: StyleSheet, raw: RawStyle[], fontNames: string[]): void {
  const styles = sheet.styles.styles;
  const byIstd = new Map<number, RawStyle>();
  for (const entry of raw) byIstd.set(entry.istd, entry);
  const idOf = (istd: number): string | undefined => (istd === ISTD_NIL ? undefined : sheet.ids[istd]);

  for (const entry of raw) {
    const { style, upx } = entry;
    const base = idOf(entry.base);
    // A style is based on a style of its own type; anything else would end the chain when it is followed.
    if (base !== undefined && base !== style.id && styles.get(base)!.type === style.type) style.basedOn = base;
    const next = idOf(entry.next);
    if (next !== undefined && next !== style.id) style.next = next;
    const link = idOf(entry.link);
    if (link !== undefined && link !== style.id) style.link = link;

    const list = (index: number): [Uint8Array, number, number] | undefined => (index * 2 + 1 < upx.length ? [table, upx[index * 2], upx[index * 2 + 1]] : undefined);
    let para: [Uint8Array, number, number] | undefined;
    let run: [Uint8Array, number, number] | undefined;
    let tap: [Uint8Array, number, number] | undefined;
    if (style.type === 'paragraph') {
      para = list(0);
      run = list(1);
    } else if (style.type === 'character') run = list(0);
    else if (style.type === 'table') {
      tap = list(0);
      para = list(1);
      run = list(2);
    } else para = list(0);

    if (para && para[2] - para[1] >= 2) {
      // The list starts with the index of the style itself.
      const info = paraProps(ctx, entry.istd, [[para[0], para[1] + 2, para[2]]], true);
      if (Object.keys(info.props).length) style.para = info.props;
    }
    if (run && run[2] > run[1]) {
      const info = charProps(ctx, [run]);
      // A style does not name another style for its text.
      delete info.props.style;
      if (info.invert && style.type === 'character') {
        // The opposite of what the paragraph has: what `true` says in a character style of the model.
        for (const name of info.invert) info.props[name] = true;
        (sheet.flips || (sheet.flips = new Map())).set(style.id, info.invert);
      } else if (info.invert) entry.invert = info.invert;
      if (Object.keys(info.props).length) style.run = info.props;
    }
    if (tap && tap[2] > tap[1]) {
      const sprms: { op: number; b: Uint8Array; at: number; length: number }[] = [];
      const info = paraProps(ctx, entry.istd, [tap], true);
      for (const sprm of info.table) sprms.push(sprm);
      const row = readTap(ctx, sprms);
      if (Object.keys(row.table).length) style.table = row.table;
    }
  }

  // "The opposite of the style it is based on": worked out from the base, which may itself depend on its base.
  const state = new Map<RawStyle, number>();
  const stated = (entry: RawStyle | undefined, name: ToggleName, depth: number): boolean => {
    // The value a style chain ends up with for a toggle; false where nothing states it.
    for (let hops = 0; entry && hops < 64; hops++) {
      settle(entry, depth);
      const value = entry.style.run ? entry.style.run[name] : undefined;
      if (value !== undefined) return value;
      const base = entry.style.basedOn === undefined ? undefined : sheet.byId.get(entry.style.basedOn);
      entry = base === undefined ? undefined : byIstd.get(base);
    }
    return false;
  };
  const settle = (entry: RawStyle, depth: number): void => {
    if (!entry.invert || state.get(entry) !== undefined) return;
    // In progress: a loop of styles based on each other; the toggle stays unstated.
    state.set(entry, 1);
    if (depth < 64) {
      const baseIstd = entry.style.basedOn === undefined ? undefined : sheet.byId.get(entry.style.basedOn);
      const base = baseIstd === undefined ? undefined : byIstd.get(baseIstd);
      const run: RunProps = entry.style.run || (entry.style.run = {});
      for (const name of entry.invert) run[name] = !stated(base, name, depth + 1);
    }
    state.set(entry, 2);
  };
  for (const entry of raw) settle(entry, 0);

  // What a docx file would say in its document defaults: 10 point text in the default fonts of the file.
  const run: RunProps = { size: 20, sizeCs: 20 };
  if (sheet.defaultFonts) {
    const [ascii, eastAsia, other, complex] = sheet.defaultFonts;
    const fonts: NonNullable<RunProps['fonts']> = {};
    if (fontNames[ascii]) fonts.ascii = fontNames[ascii];
    if (fontNames[other]) fonts.hAnsi = fontNames[other];
    if (fontNames[eastAsia]) fonts.eastAsia = fontNames[eastAsia];
    if (complex !== undefined && fontNames[complex]) fonts.cs = fontNames[complex];
    if (Object.keys(fonts).length) run.fonts = fonts;
  }
  sheet.styles.defaults = { para: {}, run };
}
