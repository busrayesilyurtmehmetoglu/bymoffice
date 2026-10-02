// The style sheet of an RTF file -> the styles of the document model.
//
// A style of the file is known by a number ("\s1", "\cs10", "\ts11"); the
// model knows styles by an id, as a docx file does. The ids made here are
// the ones Word itself gives the built-in styles ("Normal", "Heading1", ...)
// and the name without its spaces for the others, so that code written for
// docx finds what it expects.
//
// A style of the file states its formatting in full, what it has from the
// style it is based on included. The model states only what a style adds to
// its base, so each style is compared with what its base gives (see
// state.ts).
//
// Table styles are kept by name only. A row of an RTF table states its
// borders, padding and shading itself; laying a table style under that would
// add what the file does not say.

import {
  defaultSettings,
  emptyDocStyles,
  emptyTheme,
  type DocModel,
  type DocStyle,
  type DocStyles,
  type ParaProps,
  type RunProps,
  type StyleType,
} from '../../model/doc/types';
import { resolveParaProps, resolveRunProps } from '../docx/resolve';
import { defaultChp, diffParaProps, diffRunProps, fullParaProps, fullRunProps, lookup, type Chp, type FormatContext, type Pap } from './state';

/** More styles than this in one document is not a genuine file. */
export const MAX_STYLES = 20000;

/** A style while its group is read. */
export interface RawStyle {
  number: number;
  type: StyleType;
  name: string;
  basedOn: number | undefined;
  next: number | undefined;
  link: number | undefined;
  hidden: boolean;
  semiHidden: boolean;
  unhideWhenUsed: boolean;
  locked: boolean;
  quickFormat: boolean;
  priority: number | undefined;
  chp: Chp | undefined;
  pap: Pap | undefined;
  rtl: boolean;
}

export function newRawStyle(): RawStyle {
  return {
    number: 0,
    type: 'paragraph',
    name: '',
    basedOn: undefined,
    next: undefined,
    link: undefined,
    hidden: false,
    semiHidden: false,
    unhideWhenUsed: false,
    locked: false,
    quickFormat: false,
    priority: undefined,
    chp: undefined,
    pap: undefined,
    rtl: false,
  };
}

/** A control word that describes a style itself; false when it is not one. */
export function styleWord(style: RawStyle, word: string, param: number): boolean {
  switch (word) {
    case 's':
      style.type = 'paragraph';
      style.number = param;
      return true;
    case 'cs':
      style.type = 'character';
      style.number = param;
      return true;
    case 'ts':
      style.type = 'table';
      style.number = param;
      return true;
    case 'ds':
      // A section style: the model has no such thing.
      style.type = 'numbering';
      style.number = param;
      return true;
    case 'sbasedon':
      style.basedOn = param;
      return true;
    case 'snext':
      style.next = param;
      return true;
    case 'slink':
      style.link = param;
      return true;
    case 'shidden':
      style.hidden = true;
      return true;
    case 'ssemihidden':
      style.semiHidden = true;
      return true;
    case 'sunhideused':
      style.unhideWhenUsed = true;
      return true;
    case 'slocked':
      style.locked = true;
      return true;
    case 'sqformat':
      style.quickFormat = true;
      return true;
    case 'spriority':
      style.priority = param;
      return true;
    case 'additive':
    case 'sautoupd':
    case 'spersonal':
    case 'scompose':
    case 'sreply':
    case 'styrsid':
      return true;
  }
  return false;
}

/** Ids of built-in styles that are not just the name with its words capitalised and joined. */
const BUILT_IN_IDS: { [name: string]: string } = lookup({
  'normal table': 'TableNormal',
  'table normal': 'TableNormal',
  'annotation text': 'CommentText',
  'annotation reference': 'CommentReference',
  'annotation subject': 'CommentSubject',
  'default paragraph font': 'DefaultParagraphFont',
});

function makeId(name: string, number: number, used: Set<string>): string {
  const lower = name.toLowerCase();
  let id = BUILT_IN_IDS[lower];
  if (id === undefined) {
    // The built-in styles are written in lower case ("heading 1", "footnote text"); their ids capitalise each word.
    const builtIn = name === lower;
    id = '';
    let start = true;
    for (let i = 0; i < name.length && id.length < 64; i++) {
      const c = name.charCodeAt(i);
      const letter = (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c > 127;
      if (letter || (c >= 48 && c <= 57)) {
        id += start && builtIn && letter ? name.charAt(i).toUpperCase() : name.charAt(i);
        start = false;
      } else start = true;
    }
    if (lower.slice(0, 4) === 'toc ') id = 'TOC' + id.slice(3);
    if (!id) id = 'Style' + number;
  }
  // Two styles never share an id, whatever their names are.
  if (used.has(id)) id = id + '_' + number;
  used.add(id);
  return id;
}

export interface StyleSheet {
  styles: DocStyles;
  /** Style ids by the number of the file, for the paragraph and the character styles. */
  paraIds: Map<number, string>;
  charIds: Map<number, string>;
}

/** A model that holds nothing but styles: what the resolvers need to say what a style gives. */
export function styleModel(styles: DocStyles): DocModel {
  return {
    body: { part: '', rootName: '', rootAttrs: [], blocks: [], relationships: new Map(), source: '' },
    styles,
    numbering: { abstracts: new Map(), nums: new Map() },
    theme: emptyTheme(),
    settings: defaultSettings(),
    fonts: new Map(),
    headers: new Map(),
    footers: new Map(),
    strict: false,
    warnings: [],
    unknown: new Map(),
  };
}

/**
 * Makes the styles of the model. `ctx` gives font names and colours; it must
 * not name styles (a style does not carry the name of another one in its
 * formatting).
 */
export function buildStyles(raw: RawStyle[], ctx: FormatContext, language: string | undefined): StyleSheet {
  const styles = emptyDocStyles();
  const sheet: StyleSheet = { styles, paraIds: new Map(), charIds: new Map() };

  // What the file means where it says nothing: the state after "\plain".
  const plain = fullRunProps(ctx, defaultChp(), false);
  const run: RunProps = { size: plain.size, sizeCs: plain.sizeCs };
  if (plain.fonts && Object.keys(plain.fonts).length) run.fonts = plain.fonts;
  if (language !== undefined) run.lang = { val: language };
  styles.defaults = { para: {}, run };

  const used = new Set<string>();
  const byNumber = new Map<number, { raw: RawStyle; style: DocStyle }>();
  const entries: { raw: RawStyle; style: DocStyle }[] = [];
  for (const item of raw) {
    // The first style with a number counts; a section style has no place in the model.
    if (item.type === 'numbering' || byNumber.has(item.number) || entries.length >= MAX_STYLES) continue;
    const name = item.name || (item.type === 'paragraph' && item.number === 0 ? 'Normal' : 'Style' + item.number);
    const style: DocStyle = { id: makeId(name, item.number, used), type: item.type, name, isDefault: false };
    if (item.hidden) style.hidden = true;
    if (item.semiHidden) style.semiHidden = true;
    if (item.unhideWhenUsed) style.unhideWhenUsed = true;
    if (item.locked) style.locked = true;
    if (item.quickFormat) style.quickFormat = true;
    if (item.priority !== undefined) style.uiPriority = item.priority;
    const entry = { raw: item, style };
    byNumber.set(item.number, entry);
    entries.push(entry);
    styles.styles.set(style.id, style);
    if (item.type === 'paragraph') sheet.paraIds.set(item.number, style.id);
    else if (item.type === 'character') sheet.charIds.set(item.number, style.id);
  }

  for (const { raw: item, style } of entries) {
    const base = item.basedOn === undefined ? undefined : byNumber.get(item.basedOn);
    // A style is based on a style of its own type; anything else would end the chain when it is followed.
    if (base && base.style !== style && base.style.type === style.type) style.basedOn = base.style.id;
    const next = item.next === undefined ? undefined : byNumber.get(item.next);
    if (next && next.style !== style) style.next = next.style.id;
    const link = item.link === undefined ? undefined : byNumber.get(item.link);
    if (link && link.style !== style) style.link = link.style.id;
    // The paragraph style with the number 0 is the one every paragraph has that names none.
    if (style.type === 'paragraph' && item.number === 0) {
      style.isDefault = true;
      styles.defaultIds.paragraph = style.id;
    } else if (style.type === 'character' && style.id === 'DefaultParagraphFont' && styles.defaultIds.character === undefined) {
      style.isDefault = true;
      styles.defaultIds.character = style.id;
    }
  }

  // Character styles add to the text they are put on: what differs from the plain state is what they state.
  for (const { raw: item, style } of entries) {
    if (style.type !== 'character' || !item.chp) continue;
    const props = diffRunProps(fullRunProps(ctx, item.chp, false), plain);
    delete props.style;
    if (Object.keys(props).length) style.run = props;
  }

  // Paragraph styles, each after the style it is based on.
  const model = styleModel(styles);
  const bare = styleModel({ defaults: styles.defaults, styles: new Map(), defaultIds: {} });
  const baseRun = resolveRunProps(bare, {}, {});
  const basePara = resolveParaProps(bare, {});
  const state = new Map<DocStyle, number>();
  const byId = new Map<string, { raw: RawStyle; style: DocStyle }>();
  for (const entry of entries) byId.set(entry.style.id, entry);
  const settle = (entry: { raw: RawStyle; style: DocStyle }, depth: number): void => {
    const { raw: item, style } = entry;
    if (state.get(style) !== undefined) return;
    state.set(style, 1);
    let below: ParaProps = basePara;
    let belowRun: RunProps = baseRun;
    const base = style.basedOn === undefined ? undefined : byId.get(style.basedOn);
    if (base && state.get(base.style) === 1) {
      // A loop of styles based on each other: this one stands alone.
      delete style.basedOn;
    } else if (base && depth < 64) {
      settle(base, depth + 1);
      below = resolveParaProps(model, { style: base.style.id });
      belowRun = resolveRunProps(model, { style: base.style.id }, {});
    } else if (base) delete style.basedOn;
    if (item.pap) {
      const props = diffParaProps(fullParaProps(ctx, item.pap), below);
      delete props.style;
      if (Object.keys(props).length) style.para = props;
    }
    if (item.chp) {
      const props = diffRunProps(fullRunProps(ctx, item.chp, item.rtl), belowRun);
      delete props.style;
      if (Object.keys(props).length) style.run = props;
    }
    state.set(style, 2);
  };
  for (const entry of entries) if (entry.style.type === 'paragraph') settle(entry, 0);
  return sheet;
}
