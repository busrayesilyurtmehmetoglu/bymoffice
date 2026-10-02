// The formatting state of an RTF reader and how it becomes the properties of
// the document model.
//
// RTF states formatting absolutely: "\plain" and "\pard" set everything back
// to fixed values, and what follows says what differs from those. The style
// of a paragraph ("\s1") is a name tag; the file repeats the formatting of
// the style on every paragraph that has it. The model works the other way
// round: a paragraph names its style and states only what differs from it.
//
// So the state is first written out in full (fullRunProps, fullParaProps:
// every property with a value, "not bold" included) and then compared with
// what the styles of the paragraph give by themselves (diffRunProps,
// diffParaProps). What is left over is the direct formatting, and the
// resolvers of the docx reader arrive at exactly the state of the file.

import type { Border, DocColor, FrameProps, Language, ParaBorders, ParaProps, RunFonts, RunProps, Shading, TabStop } from '../../model/doc/types';
import { languageTag } from '../doc/sprm';

/** Character formatting, as the control words state it. */
export interface Chp {
  /** Index into the font table; -1 is the default font of the document. */
  font: number;
  /** Fonts for text of the low and high half of the code page, for double-byte text and for complex scripts; -1 follows `font`. */
  fontLo: number;
  fontHi: number;
  fontDb: number;
  fontCs: number;
  /** Half-points; `sizeCs` 0 follows `size`. */
  size: number;
  sizeCs: number;
  bold: boolean;
  italic: boolean;
  boldCs: boolean;
  italicCs: boolean;
  caps: boolean;
  smallCaps: boolean;
  strike: boolean;
  doubleStrike: boolean;
  outline: boolean;
  shadow: boolean;
  emboss: boolean;
  imprint: boolean;
  hidden: boolean;
  noProof: boolean;
  /** As in ST_Underline; "none" for no underline. */
  underline: string;
  /** Indexes into the colour table; 0 is automatic. */
  underlineColor: number;
  color: number;
  background: number;
  patternColor: number;
  /** Hundredths of a percent of the pattern colour over the background. */
  shading: number;
  highlight: number;
  /** 0 baseline, 1 superscript, 2 subscript. */
  vertical: number;
  /** Raised (positive) or lowered, half-points. */
  position: number;
  /** Twips between characters. */
  spacing: number;
  scale: number;
  kerning: number;
  lang: number;
  langFe: number;
  langCs: number;
  /** Number of the character style; -1 for none. */
  style: number;
  border: Border | undefined;
  deleted: boolean;
  inserted: boolean;
  /** Run properties made from this state, by the paragraph style and direction they were made for. */
  cache: Map<string, RunProps> | undefined;
}

export function defaultChp(): Chp {
  return {
    font: -1,
    fontLo: -1,
    fontHi: -1,
    fontDb: -1,
    fontCs: -1,
    size: 24,
    sizeCs: 0,
    bold: false,
    italic: false,
    boldCs: false,
    italicCs: false,
    caps: false,
    smallCaps: false,
    strike: false,
    doubleStrike: false,
    outline: false,
    shadow: false,
    emboss: false,
    imprint: false,
    hidden: false,
    noProof: false,
    underline: 'none',
    underlineColor: 0,
    color: 0,
    background: 0,
    patternColor: 0,
    shading: 0,
    highlight: 0,
    vertical: 0,
    position: 0,
    spacing: 0,
    scale: 100,
    kerning: 0,
    lang: 0,
    langFe: 0,
    langCs: 0,
    style: -1,
    border: undefined,
    deleted: false,
    inserted: false,
    cache: undefined,
  };
}

export function cloneChp(chp: Chp): Chp {
  const copy = { ...chp };
  copy.cache = undefined;
  return copy;
}

/** Paragraph formatting, as the control words state it. */
export interface Pap {
  /** Number of the paragraph style; 0 is the style every document has. */
  style: number;
  align: string;
  left: number;
  right: number;
  first: number;
  before: number;
  after: number;
  beforeAuto: boolean;
  afterAuto: boolean;
  /** "\sl": 0 for single spacing; negative for an exact height. */
  line: number;
  /** "\slmult1": `line` counts 240ths of a line. */
  lineMultiple: boolean;
  tabs: TabStop[] | undefined;
  keep: boolean;
  keepNext: boolean;
  pageBreakBefore: boolean;
  /** -1 as the document says, 0 off, 1 on. */
  widowControl: number;
  noLineNumbers: boolean;
  noHyphens: boolean;
  contextual: boolean;
  rtl: boolean;
  borders: ParaBorders | undefined;
  /** Hundredths of a percent; -1 when no shading is stated. */
  shading: number;
  patternColor: number;
  background: number;
  /** 0..8; -1 for body text. */
  outline: number;
  inTable: boolean;
  /** Depth of the table the paragraph is in, as "\itap" says; 0 when it does not say. */
  depth: number;
  /** "\ls": number of the list override; 0 for none. */
  list: number;
  level: number;
  frame: FrameProps | undefined;
  /** Paragraph properties made from this state. */
  cache: ParaProps | undefined;
}

export function defaultPap(): Pap {
  return {
    style: 0,
    align: 'left',
    left: 0,
    right: 0,
    first: 0,
    before: 0,
    after: 0,
    beforeAuto: false,
    afterAuto: false,
    line: 0,
    lineMultiple: false,
    tabs: undefined,
    keep: false,
    keepNext: false,
    pageBreakBefore: false,
    widowControl: -1,
    noLineNumbers: false,
    noHyphens: false,
    contextual: false,
    rtl: false,
    borders: undefined,
    shading: -1,
    patternColor: 0,
    background: 0,
    outline: -1,
    inTable: false,
    depth: 0,
    list: 0,
    level: 0,
    frame: undefined,
    cache: undefined,
  };
}

export function clonePap(pap: Pap): Pap {
  const copy = { ...pap };
  copy.cache = undefined;
  return copy;
}

/** What the conversion needs to know about the document. */
export interface FormatContext {
  /** Name of the font with the index; of the default font for -1 and for an index the table lacks. */
  fontName(index: number): string | undefined;
  /** Colour with the index of the colour table; undefined for the automatic colour. */
  color(index: number): DocColor | undefined;
  /** Id of the paragraph style with the number; undefined for the default style and for a number no style has. */
  paraStyleId(number: number): string | undefined;
  charStyleId(number: number): string | undefined;
  /** "\widowctrl": paragraphs that do not say have widow control. */
  widowControl: boolean;
}

/**
 * A table that names from the file are looked up in: an object without the
 * members every object has, so that a control word or a style called
 * "constructor" finds nothing.
 */
export function lookup<T>(entries: { [name: string]: T }): { [name: string]: T } {
  return Object.assign(Object.create(null), entries);
}

export const AUTO_COLOR: DocColor = Object.freeze({ auto: true });
export const NIL_BORDER: Border = Object.freeze({ style: 'nil' });
const CLEAR_SHADING: Shading = Object.freeze({ pattern: 'clear', fill: AUTO_COLOR });

/** The sixteen colours a highlight can have, by "RRGGBB". */
const HIGHLIGHTS: { [rgb: string]: string } = lookup({
  '000000': 'black',
  '0000FF': 'blue',
  '00FFFF': 'cyan',
  '00FF00': 'green',
  FF00FF: 'magenta',
  FF0000: 'red',
  FFFF00: 'yellow',
  FFFFFF: 'white',
  '000080': 'darkBlue',
  '008080': 'darkCyan',
  '008000': 'darkGreen',
  '800080': 'darkMagenta',
  '800000': 'darkRed',
  '808000': 'darkYellow',
  '808080': 'darkGray',
  C0C0C0: 'lightGray',
});

const PERCENTS = [5, 10, 12, 15, 20, 25, 30, 35, 37, 40, 45, 50, 55, 60, 62, 65, 70, 75, 80, 85, 87, 90, 95];

/** The pattern name of ST_Shd for a share of the pattern colour, in hundredths of a percent. */
export function shadingPattern(share: number): string {
  if (share <= 0) return 'clear';
  if (share >= 10000) return 'solid';
  const percent = share / 100;
  let best = PERCENTS[0];
  for (const candidate of PERCENTS) if (Math.abs(candidate - percent) < Math.abs(best - percent)) best = candidate;
  return 'pct' + best;
}

export function makeShading(ctx: FormatContext, share: number, pattern: number, background: number): Shading | undefined {
  if (share <= 0 && background === 0 && pattern === 0) return undefined;
  const shading: Shading = { pattern: shadingPattern(share) };
  const fill = ctx.color(background);
  shading.fill = fill || AUTO_COLOR;
  const color = ctx.color(pattern);
  if (color) shading.color = color;
  else if (share > 0) shading.color = AUTO_COLOR;
  return shading;
}

const VERTICAL: NonNullable<RunProps['vertAlign']>[] = ['baseline', 'superscript', 'subscript'];

/** Every run property the reader knows, with a value: the state written out in full. */
export function fullRunProps(ctx: FormatContext, chp: Chp, rtl: boolean): RunProps {
  const props: RunProps = {
    bold: chp.bold,
    italic: chp.italic,
    boldCs: chp.boldCs,
    italicCs: chp.italicCs,
    caps: chp.caps,
    smallCaps: chp.smallCaps,
    strike: chp.strike,
    doubleStrike: chp.doubleStrike,
    outline: chp.outline,
    shadow: chp.shadow,
    emboss: chp.emboss,
    imprint: chp.imprint,
    vanish: chp.hidden,
    noProof: chp.noProof,
    rtl,
    size: chp.size,
    sizeCs: chp.sizeCs || chp.size,
    vertAlign: VERTICAL[chp.vertical] || 'baseline',
    position: chp.position,
    spacing: chp.spacing,
    scale: chp.scale,
    kerning: chp.kerning,
    color: ctx.color(chp.color) || AUTO_COLOR,
  };
  const underline: NonNullable<RunProps['underline']> = { style: chp.underline };
  const underlineColor = ctx.color(chp.underlineColor);
  if (underlineColor && chp.underline !== 'none') underline.color = underlineColor;
  props.underline = underline;

  const fonts: RunFonts = {};
  const base = ctx.fontName(chp.font);
  const ascii = chp.fontLo >= 0 ? ctx.fontName(chp.fontLo) : base;
  const hAnsi = chp.fontHi >= 0 ? ctx.fontName(chp.fontHi) : base;
  const eastAsia = chp.fontDb >= 0 ? ctx.fontName(chp.fontDb) : base;
  const cs = chp.fontCs >= 0 ? ctx.fontName(chp.fontCs) : base;
  if (ascii !== undefined) fonts.ascii = ascii;
  if (hAnsi !== undefined) fonts.hAnsi = hAnsi;
  if (eastAsia !== undefined) fonts.eastAsia = eastAsia;
  if (cs !== undefined) fonts.cs = cs;
  props.fonts = fonts;

  // A highlight is one of sixteen colours; any other background colour is shading.
  const highlight = ctx.color(chp.highlight);
  const name = highlight && highlight.rgb !== undefined ? HIGHLIGHTS[highlight.rgb] : undefined;
  props.highlight = name === undefined ? 'none' : name;
  const background = name === undefined && highlight ? chp.highlight : chp.background;
  const shading = makeShading(ctx, chp.shading, chp.patternColor, background);
  if (shading) props.shading = shading;
  if (chp.border) props.border = chp.border;

  const lang: Language = {};
  const val = languageTag(chp.lang);
  if (val !== undefined) lang.val = val;
  const eastAsiaLang = languageTag(chp.langFe);
  if (eastAsiaLang !== undefined) lang.eastAsia = eastAsiaLang;
  const bidi = languageTag(chp.langCs);
  if (bidi !== undefined) lang.bidi = bidi;
  props.lang = lang;

  const style = chp.style < 0 ? undefined : ctx.charStyleId(chp.style);
  if (style !== undefined) props.style = style;
  return props;
}

const RUN_FLAGS = [
  'bold',
  'italic',
  'boldCs',
  'italicCs',
  'caps',
  'smallCaps',
  'strike',
  'doubleStrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'vanish',
  'noProof',
  'rtl',
] as const;

const FONT_SLOTS = ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const;
const LANG_SLOTS = ['val', 'eastAsia', 'bidi'] as const;

const colorKey = (color: DocColor | undefined): string => (color && color.rgb !== undefined ? color.rgb : 'auto');
const same = (a: unknown, b: unknown): boolean => a === b || JSON.stringify(a) === JSON.stringify(b);

/**
 * The properties of `full` (made by fullRunProps) that differ from `base`,
 * the properties in effect without them. A property the base has and the
 * state does not is stated as "off".
 */
export function diffRunProps(full: RunProps, base: RunProps): RunProps {
  const out: RunProps = {};
  for (const key of RUN_FLAGS) {
    const value = full[key] === true;
    if (value !== (base[key] === true)) out[key] = value;
  }
  if (full.size !== base.size) out.size = full.size;
  if (full.sizeCs !== (base.sizeCs === undefined ? base.size : base.sizeCs)) out.sizeCs = full.sizeCs;
  if (full.vertAlign !== (base.vertAlign || 'baseline')) out.vertAlign = full.vertAlign;
  if (full.position !== (base.position || 0)) out.position = full.position;
  if (full.spacing !== (base.spacing || 0)) out.spacing = full.spacing;
  if (full.scale !== (base.scale === undefined ? 100 : base.scale)) out.scale = full.scale;
  if (full.kerning !== (base.kerning || 0)) out.kerning = full.kerning;
  if (colorKey(full.color) !== colorKey(base.color)) out.color = full.color;
  if (full.highlight !== (base.highlight || 'none')) out.highlight = full.highlight;

  const underline = full.underline!;
  const baseUnderline = base.underline;
  const baseStyle = (baseUnderline && baseUnderline.style) || 'none';
  if (underline.style !== baseStyle || colorKey(underline.color) !== colorKey(baseUnderline && baseUnderline.color)) out.underline = underline;

  if (full.fonts) {
    let fonts: RunFonts | undefined;
    for (const slot of FONT_SLOTS) {
      const name = full.fonts[slot];
      if (name !== undefined && name !== (base.fonts && base.fonts[slot])) (fonts || (fonts = {}))[slot] = name;
    }
    if (fonts) out.fonts = fonts;
  }
  if (full.lang) {
    let lang: Language | undefined;
    for (const slot of LANG_SLOTS) {
      const tag = full.lang[slot];
      if (tag !== undefined && tag !== (base.lang && base.lang[slot])) (lang || (lang = {}))[slot] = tag;
    }
    if (lang) out.lang = lang;
  }
  if (!same(full.shading, base.shading)) out.shading = full.shading || CLEAR_SHADING;
  if (!same(full.border, base.border)) out.border = full.border || NIL_BORDER;
  if (full.style !== undefined) out.style = full.style;
  return out;
}

const PARA_FLAGS = ['keepNext', 'keepLines', 'pageBreakBefore', 'contextualSpacing', 'suppressLineNumbers', 'suppressAutoHyphens', 'bidi'] as const;
const BORDER_SIDES = ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const;

/** Every paragraph property the reader knows, with a value. */
export function fullParaProps(ctx: FormatContext, pap: Pap): ParaProps {
  const props: ParaProps = {
    align: pap.align,
    keepNext: pap.keepNext,
    keepLines: pap.keep,
    pageBreakBefore: pap.pageBreakBefore,
    widowControl: pap.widowControl < 0 ? ctx.widowControl : pap.widowControl === 1,
    contextualSpacing: pap.contextual,
    suppressLineNumbers: pap.noLineNumbers,
    suppressAutoHyphens: pap.noHyphens,
    bidi: pap.rtl,
  };
  props.indent = pap.first < 0 ? { left: pap.left, right: pap.right, hanging: -pap.first } : { left: pap.left, right: pap.right, firstLine: pap.first };
  const spacing: NonNullable<ParaProps['spacing']> = { before: pap.before, after: pap.after, beforeAuto: pap.beforeAuto, afterAuto: pap.afterAuto };
  if (pap.line === 0) {
    spacing.line = 240;
    spacing.lineRule = 'auto';
  } else if (pap.line < 0) {
    spacing.line = -pap.line;
    spacing.lineRule = 'exact';
  } else {
    spacing.line = pap.line;
    spacing.lineRule = pap.lineMultiple ? 'auto' : 'atLeast';
  }
  props.spacing = spacing;
  if (pap.tabs && pap.tabs.length) props.tabs = pap.tabs;
  if (pap.borders) props.borders = pap.borders;
  if (pap.shading >= 0 || pap.background !== 0 || pap.patternColor !== 0) {
    const shading = makeShading(ctx, Math.max(0, pap.shading), pap.patternColor, pap.background);
    if (shading) props.shading = shading;
  }
  if (pap.outline >= 0 && pap.outline <= 8) props.outlineLevel = pap.outline;
  if (pap.list > 0) props.numbering = { numId: pap.list, level: pap.level };
  if (pap.frame) props.frame = pap.frame;
  const style = ctx.paraStyleId(pap.style);
  if (style !== undefined) props.style = style;
  return props;
}

/** The properties of `full` (made by fullParaProps) that differ from `base`; see diffRunProps. */
export function diffParaProps(full: ParaProps, base: ParaProps): ParaProps {
  const out: ParaProps = {};
  if (full.align !== (base.align || 'left')) out.align = full.align;
  for (const key of PARA_FLAGS) {
    const value = full[key] === true;
    if (value !== (base[key] === true)) out[key] = value;
  }
  // Widow control is on where nothing is said.
  if (full.widowControl !== (base.widowControl !== false)) out.widowControl = full.widowControl;

  const indent = full.indent!;
  const baseIndent = base.indent || {};
  const first = (indent.firstLine || 0) - (indent.hanging || 0);
  const baseFirst = (baseIndent.firstLine || 0) - (baseIndent.hanging || 0);
  let stated: NonNullable<ParaProps['indent']> | undefined;
  if (indent.left !== (baseIndent.left || 0)) (stated || (stated = {})).left = indent.left;
  if (indent.right !== (baseIndent.right || 0)) (stated || (stated = {})).right = indent.right;
  if (first !== baseFirst) {
    if (first < 0) (stated || (stated = {})).hanging = -first;
    else (stated || (stated = {})).firstLine = first;
  }
  if (stated) out.indent = stated;

  const spacing = full.spacing!;
  const baseSpacing = base.spacing || {};
  let space: NonNullable<ParaProps['spacing']> | undefined;
  if (spacing.before !== (baseSpacing.before || 0)) (space || (space = {})).before = spacing.before;
  if (spacing.after !== (baseSpacing.after || 0)) (space || (space = {})).after = spacing.after;
  if (spacing.beforeAuto !== (baseSpacing.beforeAuto === true)) (space || (space = {})).beforeAuto = spacing.beforeAuto;
  if (spacing.afterAuto !== (baseSpacing.afterAuto === true)) (space || (space = {})).afterAuto = spacing.afterAuto;
  const baseLine = baseSpacing.line === undefined ? 240 : baseSpacing.line;
  if (spacing.line !== baseLine || spacing.lineRule !== (baseSpacing.lineRule || 'auto')) {
    space = space || {};
    space.line = spacing.line;
    space.lineRule = spacing.lineRule;
  }
  if (space) out.spacing = space;

  if (!same(full.tabs || [], base.tabs || [])) {
    // The tab stops of the style that the paragraph does not have are taken out.
    const tabs: TabStop[] = [];
    const kept = new Set<number>();
    if (full.tabs) for (const tab of full.tabs) kept.add(tab.pos);
    if (base.tabs) for (const tab of base.tabs) if (!kept.has(tab.pos)) tabs.push({ pos: tab.pos, align: 'clear' });
    if (full.tabs) for (const tab of full.tabs) tabs.push(tab);
    out.tabs = tabs;
  }
  if (!same(full.borders, base.borders)) {
    const borders: ParaBorders = {};
    for (const side of BORDER_SIDES) {
      const border = full.borders && full.borders[side];
      const below = base.borders && base.borders[side];
      if (border) borders[side] = border;
      else if (below && below.style !== 'nil' && below.style !== 'none') borders[side] = NIL_BORDER;
    }
    out.borders = borders;
  }
  if (!same(full.shading, base.shading)) out.shading = full.shading || CLEAR_SHADING;
  const outline = full.outlineLevel === undefined ? 9 : full.outlineLevel;
  if (outline !== (base.outlineLevel === undefined ? 9 : base.outlineLevel)) out.outlineLevel = outline;
  if (full.numbering) out.numbering = full.numbering;
  else if (base.numbering && base.numbering.numId) out.numbering = { numId: 0 };
  if (full.frame) out.frame = full.frame;
  if (full.style !== undefined) out.style = full.style;
  return out;
}

/** Border styles by control word, as the names of ST_Border. */
export const BORDER_STYLES: { [word: string]: string } = lookup({
  brdrs: 'single',
  brdrth: 'thick',
  brdrsh: 'single',
  brdrdb: 'double',
  brdrdot: 'dotted',
  brdrdash: 'dashed',
  brdrhair: 'single',
  brdrdashsm: 'dashSmallGap',
  brdrdashd: 'dotDash',
  brdrdashdd: 'dotDotDash',
  brdrdashdot: 'dotDash',
  brdrdashdotdot: 'dotDotDash',
  brdrinset: 'inset',
  brdroutset: 'outset',
  brdrtriple: 'triple',
  brdrtnthsg: 'thinThickSmallGap',
  brdrthtnsg: 'thickThinSmallGap',
  brdrtnthtnsg: 'thinThickThinSmallGap',
  brdrtnthmg: 'thinThickMediumGap',
  brdrthtnmg: 'thickThinMediumGap',
  brdrtnthtnmg: 'thinThickThinMediumGap',
  brdrtnthlg: 'thinThickLargeGap',
  brdrthtnlg: 'thickThinLargeGap',
  brdrtnthtnlg: 'thinThickThinLargeGap',
  brdrwavy: 'wave',
  brdrwavydb: 'doubleWave',
  brdrdashdotstr: 'dashDotStroked',
  brdremboss: 'threeDEmboss',
  brdrengrave: 'threeDEngrave',
  brdrnone: 'none',
  brdrnil: 'nil',
});

/** Underline styles by control word, as the names of ST_Underline. */
export const UNDERLINES: { [word: string]: string } = lookup({
  ul: 'single',
  uld: 'dotted',
  uldash: 'dash',
  uldashd: 'dotDash',
  uldashdd: 'dotDotDash',
  uldb: 'double',
  ulhwave: 'wavyHeavy',
  ulldash: 'dashLong',
  ulth: 'thick',
  ulthd: 'dottedHeavy',
  ulthdash: 'dashedHeavy',
  ulthdashd: 'dashDotHeavy',
  ulthdashdd: 'dashDotDotHeavy',
  ulthldash: 'dashLongHeavy',
  ululdbwave: 'wavyDouble',
  ulw: 'words',
  ulwave: 'wave',
});

/** Largest length in twips the reader takes from a file; larger ones are cut to it. */
export const MAX_TWIPS = 1000000;
export const twipsOf = (value: number): number => (value > MAX_TWIPS ? MAX_TWIPS : value < -MAX_TWIPS ? -MAX_TWIPS : value);

/** A border with one property changed by a control word; `border` is not touched. */
export function borderWith(ctx: FormatContext, border: Border | undefined, word: string, param: number): Border | undefined {
  const style = BORDER_STYLES[word];
  if (style !== undefined) {
    if (style === 'nil' || style === 'none') return { style };
    const next: Border = { ...border, style };
    if (word === 'brdrsh') next.shadow = true;
    return next;
  }
  const next: Border = { style: 'single', ...border };
  switch (word) {
    case 'brdrw':
      // Twips -> eighths of a point; no line is thicker than 75 twips in Word.
      next.size = Math.round(Math.max(0, Math.min(param, 255)) * 0.4);
      return next;
    case 'brsp':
      next.space = Math.round(Math.max(0, Math.min(param, 31 * 20)) / 20);
      return next;
    case 'brdrcf': {
      const color = ctx.color(param);
      next.color = color || AUTO_COLOR;
      return next;
    }
    case 'brdrframe':
      next.frame = true;
      return next;
  }
  return border;
}
