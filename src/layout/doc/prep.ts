// Prepares a paragraph for line breaking: effective formatting in px and the
// content as a flat list of measured "atoms" (words with their trailing
// spaces, tabs, breaks, inline boxes, fields). Nothing here depends on the
// width the paragraph is broken to, so the result is kept per paragraph and
// used again whenever the paragraph is laid out.

import { abstractOf, formatCounter, levelOf } from '../../formats/docx/numbering';
import { resolveNumberRunProps, styleChain } from '../../formats/docx/resolve';
import {
  EMPTY_RUN_PROPS,
  findWrapper,
  inFieldCode,
  safeHyperlinkTarget,
  type Inline,
  type ParaProps,
  type Paragraph,
  type Run,
  type RunItem,
  type Wrapper,
} from '../../model/doc/types';
import { raisedStyle, type Band, type BorderLine, type CellWhere, type FieldValues, type FloatSpec, type LayoutContext, type RunStyle, type StateEvent, type StoryRef } from './context';
import { drawingContent, legacyContent, makeFloat } from './drawing';
import { isLiveField, mapSymbols, mayHoldSymbols, parseFieldCode, symbolFont, type LiveField } from './symbols';
import type { LayoutFont, PageItem } from './types';
import { MAX_LAYOUT_DEPTH, twips } from './units';

export interface TextAtom {
  t: 'text';
  /** A word followed by the spaces after it. */
  text: string;
  /** Length of the word. */
  word: number;
  /** Width of the word and of the spaces. */
  width: number;
  trail: number;
  font: LayoutFont;
  style: RunStyle;
  /** A line may break before the atom. */
  bb: boolean;
  /** A soft hyphen follows: a line may break after the atom, and then shows a hyphen. */
  shy: boolean;
  /** Offset in paragraphText(); -1 for text the engine makes up. */
  offset: number;
  link?: string;
  /** The atom is the mark of this footnote. */
  note?: number;
}

export interface TabAtom {
  t: 'tab';
  font: LayoutFont;
  style: RunStyle;
  offset: number;
  link?: string;
  /** <w:ptab>: aligned to the text area instead of a tab stop. */
  absolute?: { align: string; leader?: string };
}

export interface BreakAtom {
  t: 'break';
  kind: 'line' | 'page' | 'column';
  font: LayoutFont;
  style: RunStyle;
}

export interface BoxAtom {
  t: 'box';
  width: number;
  height: number;
  /** Relative to the top left corner of the box. */
  items: PageItem[];
  font: LayoutFont;
  style: RunStyle;
  bb: boolean;
  offset: number;
}

export interface FieldAtom {
  t: 'field';
  field: LiveField;
  format?: string;
  /** Result in the file, shown while the value is not known. */
  cached: string;
  font: LayoutFont;
  style: RunStyle;
  bb: boolean;
  offset: number;
  link?: string;
}

export type Atom = TextAtom | TabAtom | BreakAtom | BoxAtom | FieldAtom;

export interface PrepTab {
  pos: number;
  align: string;
  leader?: string;
}

export interface PrepLabel {
  text: string;
  font: LayoutFont;
  style: RunStyle;
  width: number;
  suffix: 'tab' | 'space' | 'nothing';
  /** "left", "center" or "right". */
  align: string;
}

export interface ParaPrep {
  para: Paragraph;
  story: string;
  path: number[];
  /** Effective properties. */
  props: ParaProps;
  atoms: Atom[];
  /** Formatting of the paragraph mark: the font of an empty line. */
  mark: RunStyle;
  /** Indents, px: of the text, of the right edge, and of the first line relative to `left` (negative when hanging). */
  left: number;
  right: number;
  first: number;
  tabs: PrepTab[];
  before: number;
  after: number;
  /** The spacing is "automatic": it merges with that of the neighbours, as margins do in HTML. */
  beforeAuto: boolean;
  afterAuto: boolean;
  /** Line spacing: multiple of the natural height ("auto") or px. */
  line: number;
  rule: 'auto' | 'exact' | 'atLeast';
  align: 'left' | 'center' | 'right' | 'both';
  label?: PrepLabel;
  floats?: FloatSpec[];
  borders?: { top?: BorderLine; left?: BorderLine; bottom?: BorderLine; right?: BorderLine; between?: BorderLine };
  shade?: string;
  keepNext: boolean;
  keepLines: boolean;
  widow: boolean;
  pageBreakBefore: boolean;
  contextual: boolean;
  /** A field atom is among the atoms: the lines depend on the page. */
  live: boolean;
  /** Lines laid out before, by width and surroundings. */
  cache?: { key: string; bands: Band[] };
  /** Content widths for automatic table layout, computed on demand. */
  widths?: { min: number; max: number };
  /** Length of paragraphText() of the paragraph. */
  length: number;
  /** Drawings in the paragraph: what they show carries paths that begin with the path of the paragraph. */
  drawings?: number;
  /** What preparing the paragraph took from the state of the document (see StateEvent); missing when nothing. */
  events?: StateEvent[];
}

// Classes of characters for line breaking; a practical subset of UAX #14.
const NONE = 0; // start of the paragraph, or after a break
const SP = 1; // spaces
const HY = 2; // hyphens and dashes after a letter or digit
const ID = 3; // ideographs, kana, hangul: break on both sides
const CL = 4; // East Asian closing punctuation: no break before
const OP = 5; // East Asian opening punctuation: no break after
const ZW = 6; // zero width space: break after
const GL = 7; // non-breaking characters
const AL = 8; // everything else
const NU = 9; // digits
const LOW = 10; // second half of a surrogate pair

const CLOSING = '、。，．）」』】〉》〕！？：；ー々・ぁぃぅぇぉっゃゅょァィゥェォッャュョ';
const OPENING = '「『【〈《〔（';

function classOf(code: number): number {
  if (code < 0x80) {
    if (code === 0x20 || code < 0x20) return SP;
    if (code === 0x2d) return HY;
    if (code >= 0x30 && code <= 0x39) return NU;
    return AL;
  }
  if (code < 0x2000) return code === 0xa0 ? GL : code === 0x1680 ? SP : AL;
  if (code < 0x2e80) {
    if (code <= 0x200a) return code === 0x2007 ? GL : SP;
    if (code === 0x200b) return ZW;
    if (code === 0x2010 || code === 0x2012 || code === 0x2013 || code === 0x2014) return HY;
    if (code === 0x2011 || code === 0x202f || code === 0x2060) return GL;
    if (code === 0x205f) return SP;
    return AL;
  }
  if (code === 0x3000) return SP;
  if (code <= 0x9fff || (code >= 0xac00 && code <= 0xd7a3) || (code >= 0xf900 && code <= 0xfaff) || (code >= 0xff00 && code <= 0xffef)) {
    const char = String.fromCharCode(code);
    if (CLOSING.indexOf(char) >= 0) return CL;
    if (OPENING.indexOf(char) >= 0) return OP;
    // Half width forms at the end of the block are letters like any other.
    if (code >= 0xff61 && code <= 0xffdc) return AL;
    return ID;
  }
  // Ideographs beyond the basic plane.
  if (code >= 0xd840 && code <= 0xd87f) return ID;
  if (code >= 0xdc00 && code <= 0xdfff) return LOW;
  if (code === 0xfeff) return GL;
  return AL;
}

/** May a line break between a character of class `prev` and one of class `cur`. */
function breakBetween(prev: number, cur: number): boolean {
  if (cur === SP || cur === LOW || cur === GL || cur === CL) return false;
  if (prev === SP || prev === ZW) return true;
  if (prev === NONE || prev === GL || prev === OP || prev === LOW) return false;
  if (prev === HY) return cur === AL || cur === ID || cur === OP;
  if (prev === ID || prev === CL) return true;
  return cur === ID || cur === OP;
}

/**
 * Capitals of text set in capitals. A character whose capital is longer than
 * it ("ß" -> "SS", "ﬁ" -> "FI") stays as it is: what follows it would
 * otherwise show at offsets of the paragraph that are not its own.
 */
function upper(text: string, dotted: boolean): string {
  if (dotted) text = text.replace(/i/g, 'İ');
  const capitals = text.toUpperCase();
  // No capital is shorter than its letter: the same length means every letter has a capital of its own length.
  if (capitals.length === text.length) return capitals;
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const pair = code >= 0xd800 && code <= 0xdbff && i + 1 < text.length;
    const char = pair ? text.slice(i, i + 2) : text.charAt(i);
    if (pair) i++;
    const capital = char.toUpperCase();
    out += capital.length === char.length ? capital : char;
  }
  return out;
}

const isLower = (char: string) => char.toUpperCase() !== char;

/** Collects the atoms of one paragraph. */
class Builder {
  atoms: Atom[] = [];
  /** Class of the last character, for break opportunities across runs. */
  prev = NONE;
  /** A break is allowed before whatever comes next (after a soft hyphen). */
  pending = false;
  live = false;

  constructor(private readonly ctx: LayoutContext) {}

  private push(text: string, font: LayoutFont, style: RunStyle, bb: boolean, offset: number, link: string | undefined): TextAtom {
    let word = text.length;
    while (word > 0 && classOf(text.charCodeAt(word - 1)) === SP) word--;
    const ctx = this.ctx;
    let width = word === 0 ? 0 : ctx.width(font, word === text.length ? text : text.slice(0, word));
    let trail = 0;
    if (word < text.length) trail = ctx.width(font, ' ') * (text.length - word);
    if (style.scale !== 1) {
      width *= style.scale;
      trail *= style.scale;
    }
    if (style.spacing !== 0) {
      width = Math.max(0, width + style.spacing * word);
      trail = Math.max(0, trail + style.spacing * (text.length - word));
    }
    const atom: TextAtom = { t: 'text', text, word, width, trail, font, style, bb, shy: false, offset };
    if (link !== undefined) atom.link = link;
    this.atoms.push(atom);
    return atom;
  }

  /** Cuts a stretch of text in one font into atoms. */
  private segment(text: string, font: LayoutFont, style: RunStyle, offset: number, link: string | undefined): void {
    let start = 0;
    let bb = false;
    let prev = this.prev;
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (code === 0x09) {
        if (i > start) this.push(text.slice(start, i), font, style, bb, offset < 0 ? -1 : offset + start, link);
        const tab: TabAtom = { t: 'tab', font, style, offset: offset < 0 ? -1 : offset + i };
        if (link !== undefined) tab.link = link;
        this.atoms.push(tab);
        start = i + 1;
        bb = false;
        prev = SP;
        this.pending = false;
        continue;
      }
      if (code === 0xad) {
        // A soft hyphen is not shown; it marks a place where the word may be broken.
        if (i > start) this.push(text.slice(start, i), font, style, bb, offset < 0 ? -1 : offset + start, link).shy = true;
        else this.softHyphen();
        start = i + 1;
        bb = true;
        this.pending = false;
        continue;
      }
      let cls = classOf(code);
      // A hyphen only breaks a word it is inside of: "-5" and " - " stay whole.
      if (cls === HY && prev !== AL && prev !== NU && prev !== ID && prev !== CL) cls = AL;
      const allowed = breakBetween(prev, cls) || this.pending;
      this.pending = false;
      if (allowed) {
        if (i > start) {
          this.push(text.slice(start, i), font, style, bb, offset < 0 ? -1 : offset + start, link);
          start = i;
        }
        bb = true;
      }
      prev = cls;
    }
    if (text.length > start) this.push(text.slice(start), font, style, bb, offset < 0 ? -1 : offset + start, link);
    else if (bb) this.pending = true;
    this.prev = prev;
  }

  /** Adds text of a run. `offset` is where it starts in paragraphText(), -1 for text without a source. */
  text(raw: string, style: RunStyle, offset: number, link: string | undefined): void {
    if (raw === '' || style.hidden) return;
    const ctx = this.ctx;
    let text = raw;
    let font = style.font;
    let mapped = false;
    if (mayHoldSymbols(style.symbol, text)) {
      const result = mapSymbols(style.symbol, text);
      if (result.changed) {
        text = result.text;
        mapped = true;
        font = symbolTextFont(ctx, font);
        const flags = result.mapped;
        if (flags && font !== style.font) {
          // Glyphs the tables do not know stay in the font of the document, the others are ordinary text.
          let start = 0;
          for (let i = 1; i <= text.length; i++) {
            if (i < text.length && flags[i] === flags[start]) continue;
            this.segment(text.slice(start, i), flags[start] ? font : style.font, style, offset < 0 ? -1 : offset + start, link);
            start = i;
          }
          return;
        }
      }
    }
    if (style.caps) text = upper(text, style.dotted);
    // Characters that need another font: East Asian ones, and the small capitals.
    let plain = mapped || (!style.small && !style.eastAsia);
    if (!plain && !style.small) {
      plain = true;
      for (let i = 0; i < text.length; i++) {
        if (text.charCodeAt(i) >= 0x2e80) {
          plain = false;
          break;
        }
      }
    }
    if (plain) {
      this.segment(text, font, style, offset, link);
      return;
    }
    let start = 0;
    let current: LayoutFont | undefined;
    let piece = '';
    for (let i = 0; i < text.length; i++) {
      let char = text.charAt(i);
      const code = text.charCodeAt(i);
      let wanted = font;
      if (style.small && isLower(char)) {
        wanted = style.small;
        const capital = upper(char, style.dotted);
        // A capital of another length ("ß") would shift the offsets of what follows.
        if (capital.length === 1) char = capital;
      } else if (style.eastAsia && code >= 0x2e80 && classOf(code) !== AL) wanted = style.eastAsia;
      else if (code === 0x20 && current) wanted = current;
      if (wanted !== current && current) {
        this.segment(piece, current, style, offset < 0 ? -1 : offset + start, link);
        piece = '';
        start = i;
      }
      current = wanted;
      piece += char;
    }
    if (current && piece) this.segment(piece, current, style, offset < 0 ? -1 : offset + start, link);
  }

  softHyphen(): void {
    const last = this.atoms[this.atoms.length - 1];
    if (last && last.t === 'text' && last.word === last.text.length && last.word > 0) last.shy = true;
    this.pending = true;
  }

  tab(style: RunStyle, offset: number, link: string | undefined, absolute?: TabAtom['absolute']): void {
    if (style.hidden) return;
    const atom: TabAtom = { t: 'tab', font: style.font, style, offset };
    if (link !== undefined) atom.link = link;
    if (absolute) atom.absolute = absolute;
    this.atoms.push(atom);
    this.prev = SP;
    this.pending = false;
  }

  lineBreak(kind: 'line' | 'page' | 'column', style: RunStyle): void {
    if (style.hidden) return;
    this.atoms.push({ t: 'break', kind, font: style.font, style });
    this.prev = NONE;
    this.pending = false;
  }

  box(width: number, height: number, items: PageItem[], style: RunStyle, offset: number): void {
    if (style.hidden) return;
    const bb = this.prev !== NONE && this.prev !== GL;
    this.atoms.push({ t: 'box', width, height, items, font: style.font, style, bb, offset });
    this.prev = ID;
    this.pending = false;
  }

  field(field: LiveField, format: string | undefined, cached: string, style: RunStyle, offset: number, link: string | undefined): void {
    if (style.hidden) return;
    const bb = breakBetween(this.prev, NU) || this.pending;
    const atom: FieldAtom = { t: 'field', field, cached, font: style.font, style, bb, offset };
    if (format !== undefined) atom.format = format;
    if (link !== undefined) atom.link = link;
    this.atoms.push(atom);
    this.prev = NU;
    this.pending = false;
    this.live = true;
  }

  /** The mark of a note: text that stays with what is before it. */
  mark(text: string, style: RunStyle, note: number | undefined): void {
    if (text === '' || style.hidden) return;
    const atom = this.push(text, style.font, style, false, -1, undefined);
    if (note !== undefined) atom.note = note;
    this.prev = AL;
    this.pending = false;
  }
}

/** The font mapped symbol characters are shown in: they are ordinary Unicode now. */
function symbolTextFont(ctx: LayoutContext, font: LayoutFont): LayoutFont {
  if (symbolFont(font.family) === 0) return font;
  return ctx.font(ctx.options.fallbackFamily || 'Arial', font.size, font.bold, false);
}

/** Target of the hyperlink an inline is in, already checked. */
function wrapperLink(wrap: Wrapper | undefined): string | undefined {
  for (let w = wrap; w; w = w.parent) {
    if (w.kind !== 'hyperlink') continue;
    let target = w.target;
    if (w.anchor !== undefined) target = (target === undefined ? '' : target) + '#' + w.anchor;
    return safeHyperlinkTarget(target);
  }
  return undefined;
}

function hiddenByRevision(wrap: Wrapper | undefined): boolean {
  for (let w = wrap; w; w = w.parent) if (w.kind === 'del' || w.kind === 'moveFrom') return true;
  return false;
}

/** Colours for the authors of tracked changes: dark enough to read as text on white. */
const REVISION_COLORS = ['#C0392B', '#1F6FB2', '#1E8449', '#8E44AD', '#B9770E', '#117A8B', '#AD1457', '#5D6D7E'];
/** No more of a name than this decides its colour. */
const MAX_AUTHOR_HASHED = 256;

/** The colour the tracked changes of an author are shown in: one of a few, by a hash of the name. */
export function revisionColor(author: string | undefined): string {
  let hash = 2166136261;
  if (author !== undefined) {
    for (let i = 0; i < author.length && i < MAX_AUTHOR_HASHED; i++) hash = Math.imul(hash ^ author.charCodeAt(i), 16777619);
  }
  return REVISION_COLORS[(hash >>> 0) % REVISION_COLORS.length];
}

/** The tracked change that decides how an inline shows: a deletion or move away anywhere around it, else the innermost insertion or move in. */
function revisionOf(wrap: Wrapper | undefined): { gone: boolean; author: string | undefined } | undefined {
  let shown: { gone: boolean; author: string | undefined } | undefined;
  for (let w = wrap; w; w = w.parent) {
    if (w.kind === 'del' || w.kind === 'moveFrom') return { gone: true, author: w.author };
    if (!shown && (w.kind === 'ins' || w.kind === 'moveTo')) shown = { gone: false, author: w.author };
  }
  return shown;
}

const REVISION_STYLES = new WeakMap<RunStyle, Map<string, RunStyle>>();

/** A run style as a tracked change shows it: struck through when it was deleted, underlined when it was inserted, in the colour of the author. */
function revisionStyle(style: RunStyle, gone: boolean, author: string | undefined): RunStyle {
  const color = revisionColor(author);
  const key = (gone ? 'd' : 'i') + color;
  let known = REVISION_STYLES.get(style);
  if (!known) REVISION_STYLES.set(style, (known = new Map()));
  let out = known.get(key);
  if (!out) {
    out = gone ? { ...style, color, strike: style.strike === 'double' ? 'double' : true, deleted: true } : { ...style, color, underline: { color, style: 'single' } };
    if (gone && out.underline) out.underline = { color, style: out.underline.style };
    known.set(key, out);
  }
  return out;
}

/**
 * Shows the text of a run that a tracked change deleted. It has no place in
 * the text (no offsets), so nothing of it can be selected or changed; fields,
 * drawings and note marks in it are not shown.
 */
function walkDeleted(walk: Walk, run: Run, style: RunStyle): void {
  if (run.fieldCode === true || walk.skip || walk.simple) return;
  const builder = walk.builder;
  for (let i = 0; i < run.content.length; i++) {
    if (inFieldCode(run, i)) continue;
    const item = run.content[i];
    if (item.type === 'text' || item.type === 'deletedText') builder.text(item.text, style, -1, undefined);
    else if (item.type === 'tab') builder.tab(style, -1, undefined);
    else if (item.type === 'noBreakHyphen') builder.text('\u2011', style, -1, undefined);
  }
}

interface Walk {
  builder: Builder;
  /** Index in the children of the paragraph of the inline being walked. */
  inline: number;
  /** Offset in paragraphText(). */
  offset: number;
  /** Open complex fields: the link of a HYPERLINK field, undefined for the others. */
  fields: (string | undefined)[];
  /** Depth of fields inside a live field whose cached result is being passed over; 0 when not. */
  skip: number;
  /** The simple field wrapper the last inline was in, whether its content is being passed over, and its link. */
  simpleSeen: Wrapper | undefined;
  simple: Wrapper | undefined;
  simpleLink: string | undefined;
  drawings: number;
  floats: FloatSpec[];
}

interface PrepArgs {
  ctx: LayoutContext;
  para: Paragraph;
  story: StoryRef;
  path: number[];
  where: CellWhere | undefined;
  depth: number;
  fields: FieldValues | undefined;
  /** Formatting of the paragraph mark, which text without a run (a formula) is shown in. */
  markStyle: RunStyle;
}

function noteMark(args: PrepArgs, walk: Walk, kind: 'footnote' | 'endnote', id: number, custom: boolean, style: RunStyle): void {
  const { ctx } = args;
  const collection = kind === 'footnote' ? ctx.doc.footnotes : ctx.doc.endnotes;
  const note = collection && collection.items.get(id);
  // Only the body numbers notes; a reference elsewhere, or to a note that is not there, shows nothing.
  if (!note || note.kind !== undefined || !args.story.main) return;
  const numbers = kind === 'footnote' ? ctx.footnoteNumbers : ctx.endnoteNumbers;
  let number = numbers.get(id);
  if (number === undefined) {
    number = numbers.size + 1;
    numbers.set(id, number);
    if (kind === 'endnote') ctx.endnoteOrder.push(id);
  }
  ctx.noteEvent(kind === 'footnote' ? 'f' : 'e', id);
  const raised = raisedStyle(ctx, style);
  if (custom) {
    // The mark is text of the run itself; the note still belongs to this place.
    if (kind === 'footnote') {
      const atom = walk.builder.atoms[walk.builder.atoms.length - 1];
      const carrier: Atom = atom && atom.t === 'text' && atom.note === undefined ? atom : pushEmpty(walk.builder, raised);
      (carrier as { note?: number }).note = id;
    }
    return;
  }
  walk.builder.mark(noteLabel(ctx, kind, number), raised, kind === 'footnote' ? id : undefined);
}

function pushEmpty(builder: Builder, style: RunStyle): Atom {
  const atom: Atom = { t: 'text', text: '', word: 0, width: 0, trail: 0, font: style.font, style, bb: false, shy: false, offset: -1 };
  builder.atoms.push(atom);
  return atom;
}

/** Text of the number of a note. */
export function noteLabel(ctx: LayoutContext, kind: 'footnote' | 'endnote', number: number): string {
  const section = ctx.doc.body.section;
  const props = section && (kind === 'footnote' ? section.footnotes : section.endnotes);
  const format = props && props.format !== undefined ? props.format : kind === 'footnote' ? 'decimal' : 'lowerRoman';
  const start = props && props.start !== undefined ? props.start : 1;
  return formatCounter(number + start - 1, format);
}

function walkItems(args: PrepArgs, walk: Walk, run: Run | undefined, items: RunItem[], style: RunStyle, link: string | undefined): void {
  const { ctx } = args;
  const builder = walk.builder;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (item.type === 'fieldChar') {
      if (item.charType === 'begin') {
        if (walk.skip) {
          walk.skip++;
          continue;
        }
        if (run && inFieldCode(run, i)) {
          // A field in the code of another field is part of that code: it shows nothing and computes nothing.
          walk.fields.push(undefined);
          continue;
        }
        const code = item.field ? parseFieldCode(item.field.instruction) : undefined;
        if (code && item.field && isLiveField(code.name) && !item.field.unclosed) {
          // The value takes the place of the result in the file and its formatting with it.
          const lender = run ? fieldFormatRun(args.para.children, walk.inline, i) : undefined;
          const shown = lender && lender !== run ? ctx.runStyle(args.para.props, lender.props, args.where) : style;
          builder.field(code.name, code.format, item.field.result, shown, walk.offset, link);
          walk.skip = 1;
          continue;
        }
        walk.fields.push(code && code.name === 'HYPERLINK' ? safeHyperlinkTarget(code.target) : undefined);
      } else if (item.charType === 'end') {
        if (walk.skip) walk.skip--;
        else walk.fields.pop();
      }
      continue;
    }
    if (run && inFieldCode(run, i)) continue;
    // What paragraphText() counts is counted here too, shown or not.
    const shown = !walk.skip && !walk.simple;
    switch (item.type) {
      case 'text':
        if (shown) builder.text(item.text, style, walk.offset, link);
        walk.offset += item.text.length;
        break;
      case 'tab':
        if (shown) builder.tab(style, walk.offset, link);
        walk.offset += 1;
        break;
      case 'break':
        if (shown) builder.lineBreak(item.kind, style);
        walk.offset += 1;
        break;
      case 'noBreakHyphen':
        if (shown) builder.text('‑', style, walk.offset, link);
        walk.offset += 1;
        break;
      case 'softHyphen':
        if (shown) builder.softHyphen();
        break;
      case 'symbol': {
        if (shown) {
          const kind = symbolFont(item.font);
          const mapped = mapSymbols(kind, item.char);
          if (mapped.changed || item.font === undefined) builder.text(mapped.text, mapped.changed ? { ...style, symbol: 0 } : style, walk.offset, link);
          else {
            // Not in the tables: the character in the font it names.
            const font = ctx.font(item.font, style.font.size, style.font.bold, style.font.italic);
            builder.text(item.char, { ...style, font, symbol: 0, eastAsia: undefined, small: undefined, caps: false, smallCaps: false }, walk.offset, link);
          }
        }
        walk.offset += item.char.length;
        break;
      }
      case 'alternate':
        walkItems(args, walk, undefined, item.items, style, link);
        break;
      case 'positionalTab':
        if (shown) builder.tab(style, -1, link, { align: item.alignment || 'left', leader: item.leader });
        break;
      case 'pageNumber':
        if (shown) builder.field('PAGE', undefined, '', style, -1, link);
        break;
      case 'footnoteReference':
        if (shown) noteMark(args, walk, 'footnote', item.id, item.customMark === true, style);
        break;
      case 'endnoteReference':
        if (shown) noteMark(args, walk, 'endnote', item.id, item.customMark === true, style);
        break;
      case 'footnoteRef':
      case 'endnoteRef':
        if (shown && args.story.noteLabel !== undefined) builder.mark(args.story.noteLabel, raisedStyle(ctx, style), undefined);
        break;
      case 'separator':
      case 'continuationSeparator': {
        if (!shown) break;
        // Word draws a line of two inches for the separator, over the whole column for the continuation.
        const width = item.type === 'separator' ? 192 : 384;
        const ascent = ctx.metrics(style.font).ascent;
        const y = ascent * 0.6;
        builder.box(width, ascent, [{ type: 'line', x1: 0, y1: y, x2: width, y2: y, color: '#000000', width: 0.75, style: 'single' }], style, -1);
        break;
      }
      case 'drawing': {
        if (!shown || args.depth >= MAX_LAYOUT_DEPTH) break;
        const content = drawingContent(ctx, item, args.story, args.path.concat([-1, walk.drawings++]), args.depth, args.fields);
        if (item.placement === 'anchor' && item.anchor) walk.floats.push(makeFloat(content, item.anchor, 0, args.where !== undefined));
        else builder.box(content.width, content.height, content.items, style, walk.offset);
        break;
      }
      case 'pict':
      case 'object': {
        if (!shown || args.depth >= MAX_LAYOUT_DEPTH) break;
        const content = legacyContent(ctx, item, args.story, args.path.concat([-1, walk.drawings++]), args.depth, args.fields);
        if (content.anchor) walk.floats.push(makeFloat(content, content.anchor, 0, args.where !== undefined));
        else if (content.width > 0 || content.items.length) builder.box(content.width, content.height, content.items, style, walk.offset);
        break;
      }
    }
  }
}

/** Items looked at to find the run a field takes its formatting from. */
const MAX_FIELD_SCAN = 4096;

/**
 * The run a live field is shown in the formatting of. Word keeps the
 * formatting of the result when it updates a field, so that is the first run
 * of the result in the file; a field without a result takes the formatting of
 * its code. Undefined when the paragraph has neither: the mark that begins the
 * field decides then. `inline` and `item` say where that mark is.
 */
function fieldFormatRun(children: Inline[], inline: number, item: number): Run | undefined {
  let depth = 1;
  let inResult = false;
  let code: Run | undefined;
  let scanned = 0;
  for (let i = inline; i < children.length; i++) {
    const child = children[i];
    if (child.type !== 'run') continue;
    if (child.wrap && hiddenByRevision(child.wrap)) continue;
    const content = child.content;
    for (let k = i === inline ? item + 1 : 0; k < content.length; k++) {
      if (++scanned > MAX_FIELD_SCAN) return code;
      const entry = content[k];
      if (entry.type === 'fieldChar') {
        if (entry.charType === 'begin') depth++;
        else if (entry.charType === 'end') {
          if (--depth === 0) return code;
        } else if (depth === 1) inResult = true;
        continue;
      }
      const isCode = entry.type === 'instrText' || inFieldCode(child, k);
      // What a field in the result shows is part of the result; its code is not.
      if (inResult && !isCode) return child;
      if (!inResult && depth === 1 && isCode && !code) code = child;
    }
  }
  return code;
}

function walkInlines(args: PrepArgs, walk: Walk, children: Inline[]): void {
  const { ctx, para, where } = args;
  for (let at = 0; at < children.length; at++) {
    const inline = children[at];
    walk.inline = at;
    if (inline.type !== 'run' && inline.type !== 'math') continue;
    const wrap = inline.wrap;
    const revision = wrap ? revisionOf(wrap) : undefined;
    if (revision && revision.gone) {
      // Shown, struck through, but not part of the text: it takes no offsets.
      if (inline.type === 'math') {
        if (!walk.skip && !walk.simple) walk.builder.text(inline.text, revisionStyle(args.markStyle, true, revision.author), -1, undefined);
      } else walkDeleted(walk, inline, revisionStyle(ctx.runStyle(para.props, inline.props, where), true, revision.author));
      continue;
    }
    let link = wrap ? wrapperLink(wrap) : undefined;
    const simple = wrap ? findWrapper(inline, 'simpleField') : undefined;
    if (simple !== walk.simpleSeen) {
      walk.simpleSeen = simple;
      walk.simple = undefined;
      walk.simpleLink = undefined;
      if (simple) {
        const code = parseFieldCode(simple.instruction);
        if (isLiveField(code.name) && inline.type === 'run') {
          // The wrapped runs are the cached result: one field atom stands for all of them.
          walk.builder.field(code.name, code.format, '', ctx.runStyle(para.props, inline.props, where), walk.offset, link);
          walk.simple = simple;
        } else if (code.name === 'HYPERLINK') walk.simpleLink = safeHyperlinkTarget(code.target);
      }
    }
    if (link === undefined) link = walk.simpleLink;
    if (link === undefined) {
      for (let k = walk.fields.length - 1; k >= 0 && link === undefined; k--) link = walk.fields[k];
    }
    if (inline.type === 'math') {
      if (!walk.skip && !walk.simple) walk.builder.text(inline.text, revision ? revisionStyle(args.markStyle, false, revision.author) : args.markStyle, walk.offset, link);
      walk.offset += inline.text.length;
      continue;
    }
    if (inline.fieldCode === true) {
      // Nothing of the run shows, but the field characters in it still open and close fields.
      let any = false;
      for (const item of inline.content) if (item.type === 'fieldChar') any = true;
      if (!any) continue;
    }
    const style = ctx.runStyle(para.props, inline.props, where);
    walkItems(args, walk, inline, inline.content, revision ? revisionStyle(style, false, revision.author) : style, link);
  }
}

const ALIGN: Record<string, ParaPrep['align']> = { center: 'center', right: 'right', end: 'right', both: 'both', distribute: 'both' };

/** The level a numbered paragraph uses when its properties name only the numbering. */
function styleLevel(ctx: LayoutContext, props: ParaProps, numId: number): number {
  const abstract = abstractOf(ctx.doc.numbering, ctx.doc.styles, numId);
  if (!abstract) return 0;
  const chain = styleChain(ctx.doc.styles, props.style, 'paragraph');
  let level = 0;
  for (const candidate of abstract.levels.values()) {
    if (candidate.paraStyle !== undefined && chain.some((style) => style.id === candidate.paraStyle)) level = candidate.level;
  }
  return level;
}

/** 14 points, what "automatic" spacing before and after a paragraph comes to. */
const AUTO_SPACING = 280;

/** Is the space before/after of these (effective) properties automatic. A compatibility setting turns that off. */
export function isAutoSpace(ctx: LayoutContext, props: ParaProps, side: 'before' | 'after'): boolean {
  const spacing = props.spacing;
  if (!spacing || !(side === 'before' ? spacing.beforeAuto : spacing.afterAuto)) return false;
  return !ctx.doc.settings.compat.get('doNotUseHTMLParagraphAutoSpacing');
}

/** Space before/after of (effective) paragraph properties, px, before the neighbours are looked at. */
export function paragraphSpace(ctx: LayoutContext, props: ParaProps, side: 'before' | 'after'): number {
  const spacing = props.spacing || {};
  if (isAutoSpace(ctx, props, side)) return twips(AUTO_SPACING);
  return Math.max(0, twips(side === 'before' ? spacing.before : spacing.after));
}

export function prepareParagraph(
  ctx: LayoutContext,
  para: Paragraph,
  story: StoryRef,
  path: number[],
  where: CellWhere | undefined,
  depth: number,
  fields: FieldValues | undefined,
): ParaPrep {
  const props = ctx.paraProps(para.props, where);
  const mark = ctx.runStyle(para.props, para.props.markProps || EMPTY_RUN_PROPS, where);
  const indent = props.indent || {};
  const spacing = props.spacing || {};
  const rule = spacing.lineRule === 'exact' || spacing.lineRule === 'atLeast' ? spacing.lineRule : 'auto';
  let line = spacing.line === undefined ? (rule === 'auto' ? 240 : 0) : spacing.line;
  if (rule === 'auto') line = line > 0 ? Math.min(line / 240, 132) : 1;
  else line = Math.max(0, twips(line));

  const prep: ParaPrep = {
    para,
    story: story.id,
    path,
    props,
    atoms: [],
    mark,
    left: twips(indent.left),
    right: twips(indent.right),
    first: indent.hanging !== undefined ? -twips(indent.hanging) : twips(indent.firstLine),
    tabs: [],
    before: paragraphSpace(ctx, props, 'before'),
    after: paragraphSpace(ctx, props, 'after'),
    beforeAuto: isAutoSpace(ctx, props, 'before'),
    afterAuto: isAutoSpace(ctx, props, 'after'),
    line,
    rule,
    align: (props.align !== undefined && ALIGN[props.align]) || 'left',
    keepNext: props.keepNext === true,
    keepLines: props.keepLines === true,
    widow: props.widowControl !== false,
    pageBreakBefore: props.pageBreakBefore === true,
    contextual: props.contextualSpacing === true,
    live: false,
    length: 0,
  };
  if (props.tabs) {
    for (const tab of props.tabs) {
      if (tab.align === 'clear' || tab.align === 'bar') continue;
      const stop: PrepTab = { pos: twips(tab.pos), align: tab.align === 'end' ? 'right' : tab.align === 'start' || tab.align === 'num' ? 'left' : tab.align };
      if (tab.leader !== undefined && tab.leader !== 'none') stop.leader = tab.leader;
      prep.tabs.push(stop);
    }
  }
  if (props.borders) {
    const b = props.borders;
    const borders = { top: ctx.border(b.top), left: ctx.border(b.left), bottom: ctx.border(b.bottom), right: ctx.border(b.right), between: ctx.border(b.between) };
    if (borders.top || borders.left || borders.bottom || borders.right || borders.between) prep.borders = borders;
  }
  const shade = ctx.shade(props.shading);
  if (shade !== undefined) prep.shade = shade;

  const numbering = props.numbering;
  if (numbering && numbering.numId !== undefined && numbering.numId !== 0) {
    const level = numbering.level !== undefined ? numbering.level : styleLevel(ctx, props, numbering.numId);
    // The number is written in the language of its text: that of the formatting the number has, else of the document.
    const definition = levelOf(ctx.doc.numbering, ctx.doc.styles, numbering.numId, level >= 0 && level <= 8 ? level : 0);
    const numberProps = definition ? resolveNumberRunProps(ctx.doc, para.props, definition, where ? where.context : undefined) : undefined;
    let lang = numberProps && numberProps.lang ? numberProps.lang.val : undefined;
    if (lang === undefined && ctx.doc.settings.themeFontLang) lang = ctx.doc.settings.themeFontLang.val;
    const label = ctx.nextNumber(story.main, numbering.numId, level, lang);
    if (label) {
      const style = ctx.makeStyle(numberProps && label.level === definition ? numberProps : resolveNumberRunProps(ctx.doc, para.props, label.level, where ? where.context : undefined));
      let text = label.text;
      let font = style.font;
      if (mayHoldSymbols(style.symbol, text)) {
        const mapped = mapSymbols(style.symbol, text);
        if (mapped.changed) {
          text = mapped.text;
          font = symbolTextFont(ctx, font);
        }
      }
      if (style.caps) text = upper(text, style.dotted);
      const align = label.level.align === 'center' ? 'center' : label.level.align === 'right' || label.level.align === 'end' ? 'right' : 'left';
      if (!style.hidden) prep.label = { text, font, style, width: ctx.width(font, text), suffix: label.level.suffix || 'tab', align };
    }
  }

  const builder = new Builder(ctx);
  const walk: Walk = { builder, inline: 0, offset: 0, fields: [], skip: 0, simpleSeen: undefined, simple: undefined, simpleLink: undefined, drawings: 0, floats: [] };
  walkInlines({ ctx, para, story, path, where, depth, fields, markStyle: mark }, walk, para.children);
  prep.atoms = builder.atoms;
  prep.live = builder.live;
  prep.length = walk.offset;
  if (walk.floats.length) prep.floats = walk.floats;
  if (walk.drawings) prep.drawings = walk.drawings;
  return prep;
}

/**
 * A paragraph of which nothing shows, its mark included (all of it is hidden
 * text): it takes no room at all, as if it were not there.
 */
export function isHiddenParagraph(prep: ParaPrep): boolean {
  return prep.mark.hidden && prep.atoms.length === 0 && !prep.label && !prep.floats;
}

/** Narrowest and widest the content of the paragraph can be set, px, for automatic table layout. */
export function contentWidths(prep: ParaPrep): { min: number; max: number } {
  if (prep.widths) return prep.widths;
  let min = 0;
  let max = 0;
  let line = prep.label ? prep.label.width : 0;
  let group = 0;
  for (const atom of prep.atoms) {
    if (atom.t === 'break') {
      if (line > max) max = line;
      line = 0;
      group = 0;
      continue;
    }
    const width = atom.t === 'text' || atom.t === 'box' ? atom.width : 0;
    const breakable = atom.t === 'tab' || ((atom.t === 'text' || atom.t === 'box' || atom.t === 'field') && atom.bb);
    if (breakable) group = 0;
    group += width;
    if (group > min) min = group;
    line += width + (atom.t === 'text' ? atom.trail : 0);
  }
  if (line > max) max = line;
  const indents = prep.left + prep.right + Math.max(0, prep.first);
  prep.widths = { min: min + indents, max: max + indents };
  return prep.widths;
}

/** Space a paragraph border takes beside the text, px. */
export const borderRoom = (line: BorderLine | undefined): number => (line ? line.space + line.width : 0);
