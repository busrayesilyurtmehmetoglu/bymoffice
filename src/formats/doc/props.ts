// Sprm lists -> the property bags of the document model.
//
// The model states properties the way WordprocessingML does, and so do the
// functions here: a sprm becomes the property the same setting has in a docx
// file, in the same unit (twips, half-points, eighths of a point).
//
// A sprm that is not applied is counted in `PropContext.unknown` under
// "sprm/0x<code>", so that a reader of the model can tell what was passed
// over. Nothing is lost by that: the file itself is kept and written back as
// it is.

import type {
  Border,
  FrameProps,
  HeaderFooterKind,
  Indent,
  NoteProps,
  PageBorders,
  ParaBorders,
  ParaProps,
  ParaSpacing,
  RevisionInfo,
  RunFonts,
  RunProps,
  SectionProps,
  TabStop,
} from '../../model/doc/types';
import { i16, i32, u16, u32, u8 } from './bytes';
import { brc, brc80, colorRef, dttm, eachSprm, icoColor, icoName, languageTag, numberFormat, shd, shd80, sprmKind } from './sprm';

export interface PropContext {
  /** Name of the font with the index in the font table of the file. */
  fontName(ftc: number): string | undefined;
  /** Id of the style with the index in the style sheet of the file. */
  styleId(istd: number): string | undefined;
  /** Names of the authors of tracked changes (SttbfRMark). */
  authors: string[];
  /** The Data stream: where a paragraph keeps a property list that is too long for its page. */
  data: Uint8Array;
  /** Count of sprms that were passed over, by "sprm/0x<code>". */
  unknown: Map<string, number>;
  /** Things the reader got past. */
  warnings: string[];
}

const HEX = '0123456789ABCDEF';

export function countUnknown(ctx: PropContext, op: number): void {
  const key = 'sprm/0x' + HEX.charAt(op >> 12) + HEX.charAt((op >> 8) & 15) + HEX.charAt((op >> 4) & 15) + HEX.charAt(op & 15);
  ctx.unknown.set(key, (ctx.unknown.get(key) || 0) + 1);
}

/** Sprms that say nothing a reader shows or keeps apart: revision save ids, font hints, layout caches. */
const IGNORED: Record<number, true> = {
  0x6815: true, // sprmCRsidProp
  0x6816: true, // sprmCRsidText
  0x6817: true, // sprmCRsidRMDel
  0x286f: true, // sprmCIdctHint
  0x2a86: true, // sprmCNeedFontFixup
  0x6467: true, // sprmPRsid
  0x703a: true, // sprmSRsid
  0x7479: true, // sprmTRsid
  0x2443: true, // sprmPFNumRMIns
  0x2447: true, // sprmPFUsePgsuSettings
  0x0868: true, // sprmCFUsePgsuSettings
  0x2430: true, // sprmPFLocked
  0x5007: true, // sprmSDmBinFirst: printer trays
  0x5008: true, // sprmSDmBinOther
  0x3006: true, // sprmSFProtected
  0x0802: true, // sprmCFFldVanish
  0x6887: true, // sprmCPbiIBullet: picture bullets are not read; the text of the level shows
  0x4888: true, // sprmCPbiGrf
};

// ----------------------------------------------------------------- characters

/** Toggle sprms: the operand is 0, 1, 0x80 (as the style has it) or 0x81 (the opposite of the style). */
const TOGGLES: Record<number, 'bold' | 'italic' | 'strike' | 'outline' | 'shadow' | 'smallCaps' | 'caps' | 'vanish' | 'imprint' | 'emboss' | 'boldCs' | 'italicCs'> = {
  0x0835: 'bold',
  0x0836: 'italic',
  0x0837: 'strike',
  0x0838: 'outline',
  0x0839: 'shadow',
  0x083a: 'smallCaps',
  0x083b: 'caps',
  0x083c: 'vanish',
  0x0854: 'imprint',
  0x0858: 'emboss',
  0x085c: 'boldCs',
  0x085d: 'italicCs',
};

export type ToggleName = (typeof TOGGLES)[number];

const FLAGS: Record<number, 'doubleStrike' | 'noProof' | 'webHidden' | 'specVanish' | 'rtl' | 'complexScript'> = {
  0x2a53: 'doubleStrike',
  0x0875: 'noProof',
  0x0811: 'webHidden',
  0x0818: 'specVanish',
  0x085a: 'rtl',
  0x0882: 'complexScript',
};

/** Underline styles by Kul (MS-DOC 2.9.129), as the names of ST_Underline. */
const UNDERLINES: Record<number, string> = {
  0x00: 'none',
  0x01: 'single',
  0x02: 'words',
  0x03: 'double',
  0x04: 'dotted',
  0x06: 'thick',
  0x07: 'dash',
  0x09: 'dotDash',
  0x0a: 'dotDotDash',
  0x0b: 'wave',
  0x14: 'dottedHeavy',
  0x17: 'dashedHeavy',
  0x19: 'dashDotHeavy',
  0x1a: 'dashDotDotHeavy',
  0x1b: 'wavyHeavy',
  0x27: 'dashLong',
  0x2b: 'wavyDouble',
  0x37: 'dashLongHeavy',
};

const EFFECTS = ['none', 'lights', 'blinkBackground', 'sparkle', 'antsBlack', 'antsRed', 'shimmer'];
const EMPHASIS = ['none', 'dot', 'comma', 'circle', 'underDot'];

/** What a list of character sprms says: formatting for the model and what the reader itself needs. */
export interface CharProps {
  props: RunProps;
  /** fSpec: the characters of the run are not text but stand for something (a field mark, a picture, a note). */
  special: boolean;
  /** sprmCPicLocation: where the picture (or field data) of a 0x01 character is in the Data stream. */
  picture?: number;
  /** The 0x01 character stands for binary data of a field, or an embedded object, not for a picture. */
  data: boolean;
  ole2: boolean;
  object: boolean;
  /** sprmCSymbol: font and character of a 0x28 character. */
  symbol?: { font?: string; char: string };
  /** The run is a tracked deletion / insertion. */
  deleted?: RevisionInfo;
  inserted?: RevisionInfo;
  /** Toggles stated as "the opposite of the style": the value depends on the styles of the run. */
  invert?: ToggleName[];
}

interface CharBuild extends CharProps {
  del: boolean;
  ins: boolean;
  delAuthor?: number;
  delDate?: number;
  insAuthor?: number;
  insDate?: number;
}

const fonts = (props: RunProps): RunFonts => props.fonts || (props.fonts = {});

function applyChar(ctx: PropContext, state: CharBuild, b: Uint8Array, op: number, at: number, length: number): void {
  const props = state.props;
  const toggle = TOGGLES[op];
  if (toggle !== undefined) {
    const value = u8(b, at);
    if (state.invert) state.invert = state.invert.filter((name) => name !== toggle);
    if (value === 0x81) {
      delete props[toggle];
      (state.invert || (state.invert = [])).push(toggle);
    } else if (value === 0x80) delete props[toggle];
    else props[toggle] = (value & 1) !== 0;
    return;
  }
  const flag = FLAGS[op];
  if (flag !== undefined) {
    props[flag] = (u8(b, at) & 1) !== 0;
    return;
  }
  switch (op) {
    case 0x4a30: {
      // sprmCIstd
      const id = ctx.styleId(u16(b, at));
      if (id !== undefined) props.style = id;
      return;
    }
    case 0x4a4f: {
      const name = ctx.fontName(u16(b, at));
      if (name !== undefined) fonts(props).ascii = name;
      return;
    }
    case 0x4a50: {
      const name = ctx.fontName(u16(b, at));
      if (name !== undefined) fonts(props).eastAsia = name;
      return;
    }
    case 0x4a51: {
      const name = ctx.fontName(u16(b, at));
      if (name !== undefined) fonts(props).hAnsi = name;
      return;
    }
    case 0x4a5e: {
      const name = ctx.fontName(u16(b, at));
      if (name !== undefined) fonts(props).cs = name;
      return;
    }
    case 0x4a3d: {
      // sprmCFtcDefault, of documents converted from Word 6: one font for all text.
      const name = ctx.fontName(u16(b, at));
      if (name !== undefined) {
        const f = fonts(props);
        f.ascii = name;
        f.hAnsi = name;
      }
      return;
    }
    case 0x4a43:
      props.size = u16(b, at);
      return;
    case 0x4a61:
      props.sizeCs = u16(b, at);
      return;
    case 0x2a42: {
      const color = icoColor(u8(b, at));
      if (color) props.color = color;
      return;
    }
    case 0x6870:
      props.color = colorRef(b, at);
      return;
    case 0x2a3e: {
      const style = UNDERLINES[u8(b, at)];
      props.underline = { ...props.underline, style: style === undefined ? 'single' : style };
      return;
    }
    case 0x6877:
      props.underline = { ...props.underline, color: colorRef(b, at) };
      return;
    case 0x2a0c: {
      const name = icoName(u8(b, at));
      if (name !== undefined) props.highlight = name;
      return;
    }
    case 0x2a48: {
      const value = u8(b, at);
      props.vertAlign = value === 1 ? 'superscript' : value === 2 ? 'subscript' : 'baseline';
      return;
    }
    case 0x4845:
      props.position = i16(b, at);
      return;
    case 0x8840:
      props.spacing = i16(b, at);
      return;
    case 0x4852: {
      const value = u16(b, at);
      if (value >= 1 && value <= 600) props.scale = value;
      return;
    }
    case 0x484b:
      props.kerning = u16(b, at);
      return;
    case 0x4866: {
      const shading = shd80(b, at);
      if (shading) props.shading = shading;
      return;
    }
    case 0xca71: {
      const shading = length >= 10 ? shd(b, at) : undefined;
      if (shading) props.shading = shading;
      return;
    }
    case 0x6865: {
      const border = brc80(b, at);
      if (border) props.border = border;
      return;
    }
    case 0xca72: {
      const border = length >= 8 ? brc(b, at) : undefined;
      if (border) props.border = border;
      return;
    }
    case 0x486d:
    case 0x4873: {
      const tag = languageTag(u16(b, at));
      if (tag !== undefined) props.lang = { ...props.lang, val: tag };
      return;
    }
    case 0x486e:
    case 0x4874: {
      const tag = languageTag(u16(b, at));
      if (tag !== undefined) props.lang = { ...props.lang, eastAsia: tag };
      return;
    }
    case 0x485f: {
      const tag = languageTag(u16(b, at));
      if (tag !== undefined) props.lang = { ...props.lang, bidi: tag };
      return;
    }
    case 0x2859: {
      const name = EFFECTS[u8(b, at)];
      if (name !== undefined) props.effect = name;
      return;
    }
    case 0x2a34: {
      const name = EMPHASIS[u8(b, at)];
      if (name !== undefined) props.emphasis = name;
      return;
    }
    case 0x2a33: {
      // sprmCPlain: back to what the style says.
      const style = props.style;
      for (const key of Object.keys(props)) delete (props as Record<string, unknown>)[key];
      if (style !== undefined) props.style = style;
      delete state.invert;
      return;
    }
    case 0x0855:
      state.special = (u8(b, at) & 1) !== 0;
      return;
    case 0x6a03:
      state.picture = u32(b, at);
      return;
    case 0x0806:
      state.data = (u8(b, at) & 1) !== 0;
      return;
    case 0x080a:
      state.ole2 = (u8(b, at) & 1) !== 0;
      return;
    case 0x0856:
      state.object = (u8(b, at) & 1) !== 0;
      return;
    case 0x6a09: {
      // sprmCSymbol: the font and the character a 0x28 character stands for.
      const symbol: { font?: string; char: string } = { char: String.fromCharCode(u16(b, at + 2)) };
      const font = ctx.fontName(u16(b, at));
      if (font !== undefined) symbol.font = font;
      state.symbol = symbol;
      state.special = true;
      return;
    }
    case 0x0800:
      state.del = (u8(b, at) & 1) !== 0;
      return;
    case 0x0801:
      state.ins = (u8(b, at) & 1) !== 0;
      return;
    case 0x4804:
      state.insAuthor = u16(b, at);
      return;
    case 0x6805:
      state.insDate = u32(b, at);
      return;
    case 0x4863:
      state.delAuthor = u16(b, at);
      return;
    case 0x6864:
      state.delDate = u32(b, at);
      return;
  }
  if (!IGNORED[op]) countUnknown(ctx, op);
}

function revision(ctx: PropContext, author: number | undefined, date: number | undefined): RevisionInfo {
  const info: RevisionInfo = {};
  if (author !== undefined && ctx.authors[author] !== undefined) info.author = ctx.authors[author];
  const text = date === undefined ? undefined : dttm(date);
  if (text !== undefined) info.date = text;
  return info;
}

/**
 * Reads lists of sprms, one after the other, for what they say about
 * characters; each is [buffer, start, end]. Sprms of another kind are passed
 * over: the modifier of a piece holds both kinds in one list.
 */
export function charProps(ctx: PropContext, lists: [Uint8Array, number, number][]): CharProps {
  const state: CharBuild = { props: {}, special: false, data: false, ole2: false, object: false, del: false, ins: false };
  for (const [b, start, end] of lists) {
    const whole = eachSprm(b, start, end, (op, at, length) => {
      if (sprmKind(op) === 2) applyChar(ctx, state, b, op, at, length);
    });
    if (!whole) countUnknown(ctx, 0);
  }
  const out: CharProps = { props: state.props, special: state.special, data: state.data, ole2: state.ole2, object: state.object };
  if (state.picture !== undefined) out.picture = state.picture;
  if (state.symbol) out.symbol = state.symbol;
  if (state.del) out.deleted = revision(ctx, state.delAuthor, state.delDate);
  if (state.ins) out.inserted = revision(ctx, state.insAuthor, state.insDate);
  if (state.invert && state.invert.length) out.invert = state.invert;
  return out;
}

// ----------------------------------------------------------------- paragraphs

const PARA_FLAGS: Record<number, keyof ParaProps> = {
  0x2405: 'keepLines',
  0x2406: 'keepNext',
  0x2407: 'pageBreakBefore',
  0x2431: 'widowControl',
  0x240c: 'suppressLineNumbers',
  0x242a: 'suppressAutoHyphens',
  0x2441: 'bidi',
  0x246d: 'contextualSpacing',
  0x2470: 'mirrorIndents',
  0x2433: 'kinsoku',
  0x2434: 'wordWrap',
  0x2435: 'overflowPunct',
  0x2436: 'topLinePunct',
  0x2437: 'autoSpaceDE',
  0x2438: 'autoSpaceDN',
  0x2448: 'adjustRightInd',
};

const ALIGNMENTS = ['left', 'center', 'right', 'both', 'distribute', 'mediumKashida', 'numTab', 'highKashida', 'lowKashida', 'thaiDistribute'];
const FONT_ALIGNMENTS = ['top', 'center', 'baseline', 'bottom', 'auto'];
const TAB_ALIGNMENTS = ['left', 'center', 'right', 'decimal', 'bar'];
const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];

/** A table sprm found in the properties of a paragraph: the row mark carries the properties of its row. */
export interface TableSprm {
  op: number;
  b: Uint8Array;
  at: number;
  length: number;
}

/** What the properties of a paragraph say: formatting for the model and where the paragraph is in a table. */
export interface ParaInfo {
  props: ParaProps;
  /** Index of the paragraph style in the style sheet. */
  istd: number;
  inTable: boolean;
  /** The paragraph is the mark that ends a table row. */
  rowEnd: boolean;
  /** Depth of tables around the paragraph: 0 outside of tables. */
  depth: number;
  /** In a nested table: the paragraph ends a cell / is the mark that ends a row. */
  innerCell: boolean;
  innerRowEnd: boolean;
  table: TableSprm[];
  /**
   * The alignment is a side of the page (sprmPJc80 alone) and the list does not say which way the
   * paragraph runs: in a right-to-left paragraph, which only its style can make it, left and right
   * change places.
   */
  physicalAlign?: boolean;
}

interface ParaBuild extends ParaInfo {
  /** sprmPJc80: the alignment as a side of the page. */
  jc80?: number;
  /** sprmPJc: the alignment as the model has it, "left" being the side a paragraph starts at. */
  jc?: number;
  itap: number;
  /** Property lists in the Data stream that were followed for this paragraph. */
  followed: number;
  /** Frame position bits (sprmPPc), kept until the frame is put together. */
  frame?: FrameProps;
}

const indentOf = (props: ParaProps): Indent => props.indent || (props.indent = {});
const spacingOf = (props: ParaProps): ParaSpacing => props.spacing || (props.spacing = {});
const bordersOf = (props: ParaProps): ParaBorders => props.borders || (props.borders = {});
const frameOf = (state: ParaBuild): FrameProps => state.frame || (state.frame = {});

const PARA_BORDERS_80: Record<number, keyof ParaBorders> = { 0x6424: 'top', 0x6425: 'left', 0x6426: 'bottom', 0x6427: 'right', 0x6428: 'between', 0x6629: 'bar' };
const PARA_BORDERS: Record<number, keyof ParaBorders> = { 0xc64e: 'top', 0xc64f: 'left', 0xc650: 'bottom', 0xc651: 'right', 0xc652: 'between', 0xc653: 'bar' };

const FRAME_ANCHORS_V = ['margin', 'page', 'text'];
const FRAME_ANCHORS_H = ['text', 'margin', 'page'];
const FRAME_WRAPS = ['auto', 'notBeside', 'around', 'none', 'tight', 'through'];

/** Tabs taken from each of the two lists of one sprm; Word itself allows 64 tabs in a paragraph. */
const MAX_TABS = 64;

function addTabs(props: ParaProps, b: Uint8Array, at: number, length: number, withClose: boolean): void {
  const end = at + length;
  const tabs: TabStop[] = props.tabs ? props.tabs.slice() : [];
  const put = (tab: TabStop) => {
    // A later statement about a position replaces an earlier one.
    for (let i = 0; i < tabs.length; i++) {
      if (tabs[i].pos === tab.pos) {
        tabs[i] = tab;
        return;
      }
    }
    tabs.push(tab);
  };
  let p = at;
  // The counts are followed as written, so that what comes after them is read from the right place; only the
  // number of tabs taken from each list is limited.
  const deleted = u8(b, p);
  p += 1;
  if (p + deleted * (withClose ? 4 : 2) > end) return;
  for (let i = 0; i < deleted && i < MAX_TABS; i++) put({ pos: i16(b, p + i * 2), align: 'clear' });
  p += deleted * (withClose ? 4 : 2);
  if (p >= end) {
    if (tabs.length) props.tabs = tabs;
    return;
  }
  const added = u8(b, p);
  p += 1;
  if (p + added * 3 > end) return;
  for (let i = 0; i < added && i < MAX_TABS; i++) {
    const descriptor = u8(b, p + added * 2 + i);
    const tab: TabStop = { pos: i16(b, p + i * 2), align: TAB_ALIGNMENTS[descriptor & 7] || 'left' };
    const leader = TAB_LEADERS[(descriptor >> 3) & 7];
    if (leader !== undefined && leader !== 'none') tab.leader = leader;
    put(tab);
  }
  if (tabs.length) props.tabs = tabs;
}

/** Nesting of "huge" property lists that is followed; they do not nest in a genuine file. */
const MAX_HUGE = 1;
/** Lists in the Data stream one paragraph may point to: one for itself and one for its table row. */
const MAX_FOLLOWED = 2;
/** Table sprms kept for one row; a genuine row has a few dozen. */
const MAX_TABLE_SPRMS = 2048;

function applyPara(ctx: PropContext, state: ParaBuild, b: Uint8Array, op: number, at: number, length: number, depth: number, inStyle: boolean): void {
  const props = state.props;
  const flag = PARA_FLAGS[op];
  if (flag !== undefined) {
    (props as Record<string, unknown>)[flag] = (u8(b, at) & 1) !== 0;
    return;
  }
  const border80 = PARA_BORDERS_80[op];
  if (border80 !== undefined) {
    const border = brc80(b, at);
    if (border) bordersOf(props)[border80] = border;
    return;
  }
  const borderNew = PARA_BORDERS[op];
  if (borderNew !== undefined) {
    const border = length >= 8 ? brc(b, at) : undefined;
    if (border) bordersOf(props)[borderNew] = border;
    return;
  }
  switch (op) {
    case 0x4600: {
      // sprmPIstd
      state.istd = u16(b, at);
      return;
    }
    case 0x2403:
      state.jc80 = u8(b, at);
      return;
    case 0x2461:
      state.jc = u8(b, at);
      return;
    case 0x840f:
    case 0x845e:
      indentOf(props).left = i16(b, at);
      return;
    case 0x840e:
    case 0x845d:
      indentOf(props).right = i16(b, at);
      return;
    case 0x8411:
    case 0x8460: {
      const value = i16(b, at);
      const indent = indentOf(props);
      delete indent.firstLine;
      delete indent.hanging;
      if (value < 0) indent.hanging = -value;
      else indent.firstLine = value;
      return;
    }
    case 0xa413:
      spacingOf(props).before = u16(b, at);
      return;
    case 0xa414:
      spacingOf(props).after = u16(b, at);
      return;
    case 0x245b:
      spacingOf(props).beforeAuto = (u8(b, at) & 1) !== 0;
      return;
    case 0x245c:
      spacingOf(props).afterAuto = (u8(b, at) & 1) !== 0;
      return;
    case 0x6412: {
      // LSPD: a multiple of a line (240ths), or a height in twips that is exact when negative.
      const line = i16(b, at);
      const spacing = spacingOf(props);
      if (u16(b, at + 2) === 1) {
        spacing.line = Math.abs(line);
        spacing.lineRule = 'auto';
      } else if (line < 0) {
        spacing.line = -line;
        spacing.lineRule = 'exact';
      } else {
        spacing.line = line;
        spacing.lineRule = 'atLeast';
      }
      return;
    }
    case 0xc60d:
      addTabs(props, b, at, length, false);
      return;
    case 0xc615:
      addTabs(props, b, at, length, true);
      return;
    case 0x260a: {
      const level = u8(b, at);
      if (level <= 8) props.numbering = { ...props.numbering, level };
      return;
    }
    case 0x460b: {
      // sprmPIlfo: 1-based index of the list; 0 takes the numbering of the style away.
      const ilfo = u16(b, at);
      props.numbering = { ...props.numbering, numId: ilfo >= 0x7ff ? 0 : ilfo };
      return;
    }
    case 0x2640: {
      const level = u8(b, at);
      if (level <= 9) props.outlineLevel = level;
      return;
    }
    case 0x4439: {
      const name = FONT_ALIGNMENTS[u16(b, at)];
      if (name !== undefined) props.textAlignment = name;
      return;
    }
    case 0x442d: {
      const shading = shd80(b, at);
      if (shading) props.shading = shading;
      return;
    }
    case 0xc64d: {
      const shading = length >= 10 ? shd(b, at) : undefined;
      if (shading) props.shading = shading;
      return;
    }
    case 0x2416:
      state.inTable = (u8(b, at) & 1) !== 0;
      return;
    case 0x2417:
      state.rowEnd = (u8(b, at) & 1) !== 0;
      return;
    case 0x6649:
      state.itap = u32(b, at);
      return;
    case 0x664a:
      state.itap = Math.max(0, state.itap + i32(b, at));
      return;
    case 0x244b:
      state.innerCell = (u8(b, at) & 1) !== 0;
      return;
    case 0x244c:
      state.innerRowEnd = (u8(b, at) & 1) !== 0;
      return;
    case 0x6645:
    case 0x6646:
    case 0x646b: {
      // sprmPHugePapx: the real list is in the Data stream, a 16-bit length and then the sprms; it replaces this
      // one. sprmPTableProps: the same for the properties of a table row, which add to what the list says.
      // A list in the Data stream does not point on to another, and a paragraph points to few: otherwise a small
      // file could make the reader walk the same long lists over and over.
      if (depth >= MAX_HUGE || inStyle || state.followed >= MAX_FOLLOWED) return;
      state.followed++;
      const offset = u32(b, at);
      const size = u16(ctx.data, offset);
      if (offset + 2 + size > ctx.data.length) {
        ctx.warnings.push('The properties of a paragraph point outside the Data stream');
        return;
      }
      if (op !== 0x646b) {
        state.props = {};
        state.table = [];
        delete state.frame;
        delete state.jc80;
        delete state.jc;
      }
      paraSprms(ctx, state, ctx.data, offset + 2, offset + 2 + size, depth + 1, inStyle);
      return;
    }
    // The frame a paragraph sits in.
    case 0x8418: {
      const value = i16(b, at);
      const frame = frameOf(state);
      delete frame.x;
      delete frame.xAlign;
      // 0 is "left"; the negative values name an alignment.
      if (value === -4) frame.xAlign = 'center';
      else if (value === -8) frame.xAlign = 'right';
      else if (value === -12) frame.xAlign = 'inside';
      else if (value === -16) frame.xAlign = 'outside';
      else frame.x = value;
      return;
    }
    case 0x8419: {
      const value = i16(b, at);
      const frame = frameOf(state);
      delete frame.y;
      delete frame.yAlign;
      if (value === -4) frame.yAlign = 'top';
      else if (value === -8) frame.yAlign = 'center';
      else if (value === -12) frame.yAlign = 'bottom';
      else if (value === -16) frame.yAlign = 'inside';
      else if (value === -20) frame.yAlign = 'outside';
      else frame.y = value;
      return;
    }
    case 0x841a:
      frameOf(state).width = Math.max(0, i16(b, at));
      return;
    case 0x261b: {
      const value = u8(b, at);
      const frame = frameOf(state);
      frame.vAnchor = FRAME_ANCHORS_V[(value >> 4) & 3] || 'text';
      frame.hAnchor = FRAME_ANCHORS_H[(value >> 6) & 3] || 'text';
      return;
    }
    case 0x2423: {
      const name = FRAME_WRAPS[u8(b, at)];
      if (name !== undefined) frameOf(state).wrap = name;
      return;
    }
    case 0x442b: {
      // The high bit says the height is a minimum.
      const value = u16(b, at);
      const frame = frameOf(state);
      frame.height = value & 0x7fff;
      frame.heightRule = value === 0 ? 'auto' : value & 0x8000 ? 'atLeast' : 'exact';
      return;
    }
    case 0x442c: {
      const value = u16(b, at);
      const frame = frameOf(state);
      const kind = value & 7;
      frame.dropCap = kind === 1 ? 'drop' : kind === 2 ? 'margin' : 'none';
      frame.lines = Math.max(1, (value >> 3) & 0x1f);
      return;
    }
    case 0x842e:
      frameOf(state).vSpace = Math.max(0, i16(b, at));
      return;
    case 0x842f:
      frameOf(state).hSpace = Math.max(0, i16(b, at));
      return;
  }
  if (sprmKind(op) === 5) {
    // Table sprms belong to the row the paragraph ends; the table reader takes them from here.
    if (state.table.length < MAX_TABLE_SPRMS) state.table.push({ op, b, at, length });
    return;
  }
  if (!IGNORED[op]) countUnknown(ctx, op);
}

function paraSprms(ctx: PropContext, state: ParaBuild, b: Uint8Array, start: number, end: number, depth: number, inStyle: boolean): void {
  const whole = eachSprm(b, start, end, (op, at, length) => {
    const kind = sprmKind(op);
    if (kind === 1 || kind === 5) applyPara(ctx, state, b, op, at, length, depth, inStyle);
  });
  if (!whole) countUnknown(ctx, 0);
}

/** Deepest nesting of tables that is read as such; a paragraph that claims more sits at this depth. */
export const MAX_TABLE_DEPTH = 32;

/** Reads lists of sprms, one after the other, for what they say about a paragraph of the style `istd` (and its table row). */
export function paraProps(ctx: PropContext, istd: number, lists: [Uint8Array, number, number][], inStyle = false): ParaInfo {
  const state: ParaBuild = { props: {}, istd, inTable: false, rowEnd: false, depth: 0, innerCell: false, innerRowEnd: false, table: [], itap: 0, followed: 0 };
  for (const [b, start, end] of lists) paraSprms(ctx, state, b, start, end, 0, inStyle);
  if (state.frame) state.props.frame = state.frame;
  // Word 2000 and later state the alignment twice: as a side of the page (sprmPJc80) and as the model has it
  // (sprmPJc), where "left" is the side the paragraph starts at. The second is taken as it is. The first alone
  // (Word 97) names a side: in a right-to-left paragraph the sides change places.
  let physicalAlign = false;
  const logical = state.jc === undefined ? undefined : ALIGNMENTS[state.jc];
  if (logical !== undefined) state.props.align = logical;
  else {
    const side = state.jc80 === undefined ? undefined : ALIGNMENTS[state.jc80];
    if (side !== undefined) {
      state.props.align = state.props.bidi === true ? mirrored(side) : side;
      physicalAlign = state.props.bidi === undefined && (side === 'left' || side === 'right');
    }
  }
  // Word 97 knows no depth: a paragraph is in a table or not.
  let depth = state.itap;
  if (state.inTable && depth === 0) depth = 1;
  if (!state.inTable && !inStyle) depth = 0;
  if (depth > MAX_TABLE_DEPTH) depth = MAX_TABLE_DEPTH;
  const info: ParaInfo = {
    props: state.props,
    istd: state.istd,
    inTable: state.inTable,
    rowEnd: state.rowEnd,
    depth,
    innerCell: state.innerCell,
    innerRowEnd: state.innerRowEnd,
    table: state.table,
  };
  if (physicalAlign) info.physicalAlign = true;
  return info;
}

/** The other side of the page for "left" and "right"; every other alignment is the same either way. */
export const mirrored = (align: string): string => (align === 'left' ? 'right' : align === 'right' ? 'left' : align);

// ------------------------------------------------------------------- sections

const SECTION_TYPES: NonNullable<SectionProps['type']>[] = ['continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'];
const VERTICAL: NonNullable<SectionProps['vAlign']>[] = ['top', 'center', 'both', 'bottom'];
const LINE_RESTART = ['newPage', 'newSection', 'continuous'];
const NOTE_RESTART = ['continuous', 'eachSect', 'eachPage'];
const NOTE_POSITION = ['sectEnd', 'pageBottom', 'beneathText', 'docEnd'];
const CHAPTER_SEPARATORS = ['hyphen', 'period', 'colon', 'emDash', 'enDash'];
const GRID_TYPES = ['default', 'linesAndChars', 'lines', 'snapToChars'];

const PAGE_BORDERS_80: Record<number, 'top' | 'left' | 'bottom' | 'right'> = { 0x702b: 'top', 0x702c: 'left', 0x702d: 'bottom', 0x702e: 'right' };
const PAGE_BORDERS: Record<number, 'top' | 'left' | 'bottom' | 'right'> = { 0xd234: 'top', 0xd235: 'left', 0xd236: 'bottom', 0xd237: 'right' };

/** Most columns a section may state widths for. */
const MAX_COLUMNS = 45;

/**
 * Section properties from the sprms of a SEPX. What Word assumes for a
 * section that says nothing (US Letter, margins of an inch and of an inch and
 * a quarter) is stated in the result, as a docx file states it.
 */
export function sectionProps(ctx: PropContext, b: Uint8Array, start: number, end: number): SectionProps {
  const page = { width: 12240, height: 15840 } as NonNullable<SectionProps['page']>;
  const margins: NonNullable<SectionProps['margins']> = { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 };
  const out: SectionProps = { page, margins };
  let columnCount = 1;
  let columnSpace = 720;
  let evenly = true;
  let separator = false;
  const widths: number[] = [];
  const spaces: number[] = [];
  let restart = false;
  let pageStart: number | undefined;
  let pageFormat: string | undefined;
  let chapterStyle = 0;
  let chapterSeparator: string | undefined;
  const line: NonNullable<SectionProps['lineNumbering']> = {};
  let lineMode: number | undefined;
  const footnotes: NoteProps = {};
  const endnotes: NoteProps = {};
  let grid: NonNullable<SectionProps['docGrid']> | undefined;
  let borders: PageBorders | undefined;

  const whole = eachSprm(b, start, end, (op, at, length) => {
    const edge80 = PAGE_BORDERS_80[op];
    if (edge80 !== undefined) {
      const border = brc80(b, at);
      if (border) (borders || (borders = {}))[edge80] = border as Border;
      return;
    }
    const edge = PAGE_BORDERS[op];
    if (edge !== undefined) {
      const border = length >= 8 ? brc(b, at) : undefined;
      if (border) (borders || (borders = {}))[edge] = border as Border;
      return;
    }
    switch (op) {
      case 0x3009: {
        const type = SECTION_TYPES[u8(b, at)];
        if (type !== undefined) out.type = type;
        return;
      }
      case 0x300a:
        out.titlePage = (u8(b, at) & 1) !== 0;
        return;
      case 0x500b:
        columnCount = Math.min(MAX_COLUMNS, u16(b, at) + 1);
        return;
      case 0x900c:
        columnSpace = Math.max(0, i16(b, at));
        return;
      case 0x3005:
        evenly = (u8(b, at) & 1) !== 0;
        return;
      case 0x3019:
        separator = (u8(b, at) & 1) !== 0;
        return;
      case 0xf203:
      case 0xf204: {
        // Width of, or space after, the column with the index in the first byte.
        const index = u8(b, at);
        if (index < MAX_COLUMNS) (op === 0xf203 ? widths : spaces)[index] = u16(b, at + 1);
        return;
      }
      case 0x300e:
        pageFormat = numberFormat(u8(b, at));
        return;
      case 0x3001:
        // Heading level whose number goes in front of the page number; 0 for none.
        chapterStyle = u8(b, at);
        return;
      case 0x3000:
        chapterSeparator = CHAPTER_SEPARATORS[u8(b, at)];
        return;
      case 0x3011:
        restart = (u8(b, at) & 1) !== 0;
        return;
      case 0x501c:
        pageStart = u16(b, at);
        return;
      case 0x7044:
        pageStart = u32(b, at);
        return;
      case 0x3013:
        lineMode = u8(b, at);
        return;
      case 0x5015:
        line.countBy = u16(b, at);
        return;
      case 0x9016:
        line.distance = Math.max(0, i16(b, at));
        return;
      case 0x501b:
        line.start = u16(b, at) + 1;
        return;
      case 0xb017:
        margins.header = u16(b, at);
        return;
      case 0xb018:
        margins.footer = u16(b, at);
        return;
      case 0x301a: {
        const value = VERTICAL[u8(b, at)];
        if (value !== undefined) out.vAlign = value;
        return;
      }
      case 0x301d: {
        const value = u8(b, at);
        page.orientation = value === 2 ? 'landscape' : 'portrait';
        return;
      }
      case 0xb01f:
        page.width = u16(b, at);
        return;
      case 0xb020:
        page.height = u16(b, at);
        return;
      case 0xb021:
        margins.left = u16(b, at);
        return;
      case 0xb022:
        margins.right = u16(b, at);
        return;
      case 0x9023:
        margins.top = i16(b, at);
        return;
      case 0x9024:
        margins.bottom = i16(b, at);
        return;
      case 0xb025:
        margins.gutter = u16(b, at);
        return;
      case 0x5026:
        page.code = u16(b, at);
        return;
      case 0x3228:
        out.bidi = (u8(b, at) & 1) !== 0;
        return;
      case 0x322a:
        out.rtlGutter = (u8(b, at) & 1) !== 0;
        return;
      case 0x522f: {
        // sprmSPgbProp: which pages, in front or behind, measured from the page or the text.
        const value = u16(b, at);
        const target = borders || (borders = {});
        target.display = ['allPages', 'firstPage', 'notFirstPage'][value & 7] || 'allPages';
        target.zOrder = (value >> 3) & 3 ? 'back' : 'front';
        target.offsetFrom = (value >> 5) & 7 ? 'page' : 'text';
        return;
      }
      case 0x9031:
        (grid || (grid = {})).linePitch = i16(b, at);
        return;
      case 0x7030:
        (grid || (grid = {})).charSpace = i32(b, at);
        return;
      case 0x5032: {
        const type = GRID_TYPES[u16(b, at)];
        if (type !== undefined) (grid || (grid = {})).type = type;
        return;
      }
      case 0x303b: {
        const value = NOTE_POSITION[u8(b, at)];
        if (value !== undefined) footnotes.position = value;
        return;
      }
      case 0x303c: {
        const value = NOTE_RESTART[u8(b, at)];
        if (value !== undefined) footnotes.restart = value;
        return;
      }
      case 0x303e: {
        const value = NOTE_RESTART[u8(b, at)];
        if (value !== undefined) endnotes.restart = value;
        return;
      }
      case 0x503f:
        footnotes.start = u16(b, at);
        return;
      case 0x5040:
        footnotes.format = numberFormat(u16(b, at));
        return;
      case 0x5041:
        endnotes.start = u16(b, at);
        return;
      case 0x5042:
        endnotes.format = numberFormat(u16(b, at));
        return;
    }
    if (IGNORED[op]) return;
    countUnknown(ctx, op);
  });
  if (!whole) countUnknown(ctx, 0);

  if (columnCount > 1 || widths.length) {
    const columns: NonNullable<SectionProps['columns']> = { count: columnCount, space: columnSpace };
    if (separator) columns.separator = true;
    if (!evenly && widths.length) {
      columns.equalWidth = false;
      columns.columns = [];
      for (let i = 0; i < columnCount; i++) {
        const column: { width?: number; space?: number } = {};
        if (widths[i] !== undefined) column.width = widths[i];
        if (spaces[i] !== undefined) column.space = spaces[i];
        columns.columns.push(column);
      }
    }
    out.columns = columns;
  }
  const chapter = chapterStyle >= 1 && chapterStyle <= 9;
  if (pageFormat !== undefined || (restart && pageStart !== undefined) || chapter) {
    const numbering: NonNullable<SectionProps['pageNumbering']> = {};
    if (pageFormat !== undefined) numbering.format = pageFormat;
    if (restart && pageStart !== undefined) numbering.start = pageStart;
    if (chapter) {
      numbering.chapStyle = chapterStyle;
      if (chapterSeparator !== undefined) numbering.chapSep = chapterSeparator;
    }
    out.pageNumbering = numbering;
  }
  if (line.countBy !== undefined && line.countBy > 0) {
    if (lineMode !== undefined && LINE_RESTART[lineMode] !== undefined) line.restart = LINE_RESTART[lineMode];
    out.lineNumbering = line;
  }
  if (grid) out.docGrid = grid;
  if (borders) out.pageBorders = borders;
  if (Object.keys(footnotes).length) out.footnotes = footnotes;
  if (Object.keys(endnotes).length) out.endnotes = endnotes;
  return out;
}

/** Keys of the header and footer stories of a section, in the order the file lists them. */
export const HEADER_ORDER: { list: 'headers' | 'footers'; kind: HeaderFooterKind }[] = [
  { list: 'headers', kind: 'even' },
  { list: 'headers', kind: 'default' },
  { list: 'footers', kind: 'even' },
  { list: 'footers', kind: 'default' },
  { list: 'headers', kind: 'first' },
  { list: 'footers', kind: 'first' },
];
