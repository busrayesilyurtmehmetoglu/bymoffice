// Formatting commands: character formatting over a range, paragraph
// formatting, lists. Property bags are shared between runs and paragraphs and
// are never changed: a change makes a new bag.

import { levelOf } from '../../formats/docx/numbering';
import { resolveColor, resolveParaProps, resolveRunProps, type TableContext } from '../../formats/docx/resolve';
import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  type AbstractNum,
  type DocColor,
  type DocModel,
  type Indent,
  type NumLevel,
  type ParaProps,
  type ParaSpacing,
  type Paragraph,
  type RunProps,
} from '../../model/doc/types';
import { buildEdit, tableContextOf, type Draft, type DocEdit } from './draft';
import { boundary, flatten, inheritedProps, markRunProps, paragraphWith, rebuild, sameInlines, sameProps, type Piece } from './pieces';
import { asRange, containerBlocks, inTextBox, isCollapsed, orderRange, paragraphLength, paragraphsBetween, resolvePosition, type DocPosition, type DocRange } from './position';

const bad = (message: string): never => {
  throw new RangeError(message);
};

/**
 * A change of character formatting. A value sets the property on the text
 * itself, `false` turns it off there even when a style turns it on, and
 * `null` removes it from the text so that the style decides again.
 */
export interface RunFormatPatch {
  bold?: boolean | null;
  italic?: boolean | null;
  /** true is a single line; a text names the kind of line as in ST_Underline ("double", "dotted", ...). */
  underline?: boolean | string | null;
  strike?: boolean | null;
  /** Font name, for all four script slots. */
  font?: string | null;
  /** Half-points: 24 is 12 pt. */
  size?: number | null;
  /** "RRGGBB", "auto", or a colour of the model (a theme colour). */
  color?: string | DocColor | null;
  /** Highlight colour name ("yellow", "green", ...; "none" hides one a style brings). */
  highlight?: string | null;
  vertAlign?: 'baseline' | 'superscript' | 'subscript' | null;
  /** Character style id. */
  style?: string | null;
}

/** Run properties with a patch laid over them: the same object when nothing changes. */
export function patchRunProps(props: RunProps, patch: RunFormatPatch): RunProps {
  const out: RunProps = { ...props };
  const set = <K extends keyof RunProps>(key: K, value: RunProps[K] | null | undefined) => {
    if (value === undefined) return;
    if (value === null) delete out[key];
    else out[key] = value;
  };
  set('bold', patch.bold);
  set('boldCs', patch.bold);
  set('italic', patch.italic);
  set('italicCs', patch.italic);
  set('strike', patch.strike);
  if (patch.strike === true) delete out.doubleStrike;
  if (patch.underline !== undefined) {
    const u = patch.underline;
    if (u !== null && typeof u !== 'boolean' && (typeof u !== 'string' || !/^[A-Za-z]{1,32}$/.test(u))) bad('The kind of underline is not valid');
    set('underline', u === null ? null : { style: u === true ? 'single' : u === false ? 'none' : u });
  }
  if (patch.font !== undefined) {
    const font = patch.font;
    if (font !== null && (typeof font !== 'string' || font === '' || font.length > 255)) bad('The font name is not valid');
    if (font === null) delete out.fonts;
    else {
      out.fonts = { ascii: font, hAnsi: font, eastAsia: font, cs: font };
      if (props.fonts && props.fonts.hint !== undefined) out.fonts.hint = props.fonts.hint;
    }
  }
  if (patch.size !== undefined) {
    const size = patch.size;
    if (size !== null && !(typeof size === 'number' && size >= 2 && size <= 3276 && Math.floor(size) === size)) bad('The font size is not valid');
    set('size', size);
    set('sizeCs', size);
  }
  if (patch.color !== undefined) {
    const color = patch.color;
    if (typeof color === 'string') {
      if (color === 'auto') out.color = { auto: true };
      else if (/^[0-9A-Fa-f]{6}$/.test(color)) out.color = { rgb: color.toUpperCase() };
      else bad('The colour is not valid');
    } else set('color', color);
  }
  if (patch.highlight !== undefined) {
    if (patch.highlight !== null && !/^[A-Za-z]{1,32}$/.test(String(patch.highlight))) bad('The highlight colour is not valid');
    set('highlight', patch.highlight);
  }
  if (patch.vertAlign !== undefined) {
    const v = patch.vertAlign;
    if (v !== null && v !== 'baseline' && v !== 'superscript' && v !== 'subscript') bad('The vertical alignment is not valid');
    set('vertAlign', v);
  }
  if (patch.style !== undefined) {
    if (patch.style !== null && (typeof patch.style !== 'string' || patch.style === '')) bad('The style id is not valid');
    set('style', patch.style);
  }
  if (sameProps(out, props)) return props;
  return Object.keys(out).length ? out : EMPTY_RUN_PROPS;
}

/** One paragraph of a range, with the part of it the range covers. */
interface Covered {
  path: number[];
  paragraph: Paragraph;
  from: number;
  to: number;
  /** The range goes on after the paragraph: its mark is covered too. */
  mark: boolean;
}

/** The paragraphs a range touches, in document order. */
function coveredParagraphs(doc: DocModel, range: DocRange): { story: string; list: Covered[] } {
  const { start, end } = orderRange(range);
  resolvePosition(doc, start);
  resolvePosition(doc, end);
  const story = start.story;
  const paths = paragraphsBetween(doc, story, start.path, end.path);
  const list: Covered[] = [];
  for (let i = 0; i < paths.length; i++) {
    const path = paths[i];
    const blocks = containerBlocks(doc, story, path.slice(0, -1));
    const paragraph = blocks[path[path.length - 1]] as Paragraph;
    const last = i === paths.length - 1;
    list.push({ path, paragraph, from: i === 0 ? start.offset : 0, to: last ? end.offset : paragraphLength(paragraph), mark: !last });
  }
  return { story, list };
}

const editable = (range: DocRange): void => {
  if (inTextBox(range.anchor.path) || inTextBox(range.focus.path)) bad('The content of a text box cannot be edited');
};

/** May this piece take character formatting: shown content of a run, not the marks of a field. */
const formattable = (piece: Piece): boolean => piece.item !== undefined && !piece.code && !piece.frozen && piece.item.type !== 'fieldChar';

function mapRunProps(draft: Draft, range: DocRange, change: (props: RunProps, paragraph: Paragraph, where: TableContext | undefined) => RunProps): void {
  editable(range);
  const { story, list } = coveredParagraphs(draft.doc, range);
  for (const part of list) {
    const paragraph = part.paragraph;
    const where = tableContextOf(draft.doc, story, part.path);
    const pieces = flatten(paragraph);
    const from = boundary(pieces, part.from, 'start');
    const to = boundary(pieces, part.to, 'caret');
    // Runs share their properties: each bag is changed once.
    const memo = new Map<RunProps, RunProps>();
    const changed = (props: RunProps): RunProps => {
      let next = memo.get(props);
      if (!next) memo.set(props, (next = change(props, paragraph, where)));
      return next;
    };
    for (let q = from; q < to; q++) {
      const piece = pieces[q];
      if (!formattable(piece)) continue;
      const next = changed(piece.props);
      if (next !== piece.props) pieces[q] = { ...piece, props: next };
    }
    let props = paragraph.props;
    if (part.mark) {
      // The paragraph mark is in the range: what is typed at the end of the paragraph is formatted alike.
      const mark = props.markProps || EMPTY_RUN_PROPS;
      const next = change(mark, paragraph, where);
      if (next !== mark) {
        props = { ...props, markProps: next };
        if (next === EMPTY_RUN_PROPS) delete props.markProps;
      }
    }
    const children = rebuild(pieces);
    if (props === paragraph.props && sameInlines(children, paragraph.children)) continue;
    draft.setParagraph(story, part.path, paragraphWith(paragraph, children, props));
  }
}

/** Sets character formatting on the text of a range. A collapsed range changes nothing: the caller keeps the format for what is typed next (see insertText). */
export function setRunFormat(doc: DocModel, range: DocRange, patch: RunFormatPatch): DocEdit {
  return buildEdit(doc, (draft) => {
    patchRunProps(EMPTY_RUN_PROPS, patch);
    if (!isCollapsed(range)) mapRunProps(draft, range, (props) => patchRunProps(props, patch));
    else resolvePosition(doc, range.anchor);
    return range;
  });
}

/** Formatting in effect for the text of a range. A property is missing when the range mixes values; `mixed` names those. */
export interface RangeFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font of Latin text. */
  font?: string;
  /** Half-points. */
  size?: number;
  /** "RRGGBB", or "auto". */
  color?: string;
  /** Highlight colour name; "none" without one. */
  highlight?: string;
  vertAlign?: 'baseline' | 'superscript' | 'subscript';
  /** Character style id; "" without one. */
  style?: string;
  /** Of the paragraphs of the range: alignment ("left", "center", "right", "both", ...). */
  align?: string;
  /** Paragraph style id as the paragraphs state it; "" for the default style. */
  paraStyle?: string;
  list?: 'bullet' | 'number' | 'none';
  /** List level 0..8; missing outside lists. */
  level?: number;
  mixed: string[];
}

type FormatValues = Omit<RangeFormat, 'mixed'>;

function runValues(doc: DocModel, effective: RunProps, direct: RunProps): FormatValues {
  const color = effective.color && !effective.color.auto ? resolveColor(doc, effective.color) : undefined;
  return {
    bold: effective.bold === true,
    italic: effective.italic === true,
    underline: effective.underline !== undefined && effective.underline.style !== undefined && effective.underline.style !== 'none',
    strike: effective.strike === true,
    font: effective.fonts ? effective.fonts.ascii : undefined,
    size: effective.size,
    color: color === undefined ? 'auto' : color.replace('#', '').toUpperCase(),
    highlight: effective.highlight === undefined ? 'none' : effective.highlight,
    vertAlign: effective.vertAlign === undefined ? 'baseline' : effective.vertAlign,
    style: direct.style === undefined ? '' : direct.style,
  };
}

/** What kind of list (effective) paragraph properties put a paragraph in. */
function listKind(doc: DocModel, effective: ParaProps): 'bullet' | 'number' | 'none' {
  const numbering = effective.numbering;
  if (!numbering || numbering.numId === undefined || numbering.numId === 0) return 'none';
  const level = levelOf(doc.numbering, doc.styles, numbering.numId, numbering.level === undefined ? 0 : numbering.level);
  return level && level.format === 'bullet' ? 'bullet' : 'number';
}

function paraValues(doc: DocModel, paragraph: Paragraph, where: TableContext | undefined): FormatValues {
  const effective = resolveParaProps(doc, paragraph.props, where);
  const list = listKind(doc, effective);
  const out: FormatValues = { align: effective.align === undefined ? 'left' : effective.align, paraStyle: paragraph.props.style === undefined ? '' : paragraph.props.style, list };
  if (list !== 'none') out.level = effective.numbering && effective.numbering.level !== undefined ? effective.numbering.level : 0;
  return out;
}

/**
 * The formatting in effect over a range, styles included: what a toolbar
 * shows. For a collapsed range it is the formatting typed text would take.
 */
export function rangeFormat(doc: DocModel, at: DocRange | DocPosition): RangeFormat {
  const range = asRange(at);
  const { story, list } = coveredParagraphs(doc, range);
  const out: Record<string, unknown> = {};
  const mixed: string[] = [];
  const seen = new Set<string>();
  const merge = (values: FormatValues, keys: string[]) => {
    for (const key of keys) {
      if (mixed.indexOf(key) >= 0) continue;
      const value = (values as Record<string, unknown>)[key];
      if (!seen.has(key)) {
        seen.add(key);
        if (value !== undefined) out[key] = value;
      } else if (out[key] !== value) {
        delete out[key];
        mixed.push(key);
      }
    }
  };
  const RUN_KEYS = ['bold', 'italic', 'underline', 'strike', 'font', 'size', 'color', 'highlight', 'vertAlign', 'style'];
  const PARA_KEYS = ['align', 'paraStyle', 'list', 'level'];
  const single = isCollapsed(range);
  let anyText = false;
  for (const part of list) {
    const paragraph = part.paragraph;
    const where = tableContextOf(doc, story, part.path);
    merge(paraValues(doc, paragraph, where), PARA_KEYS);
    const pieces = flatten(paragraph);
    const memo = new Map<RunProps, FormatValues>();
    const valuesOf = (props: RunProps): FormatValues => {
      let values = memo.get(props);
      if (!values) memo.set(props, (values = runValues(doc, resolveRunProps(doc, paragraph.props, props, where), props)));
      return values;
    };
    if (single) {
      merge(valuesOf(inheritedProps(pieces, boundary(pieces, part.from, 'caret'), paragraph.props.markProps)), RUN_KEYS);
      anyText = true;
      continue;
    }
    const from = boundary(pieces, part.from, 'start');
    const to = boundary(pieces, part.to, 'caret');
    for (let q = from; q < to; q++) {
      const piece = pieces[q];
      if (!piece.item || piece.code || piece.len === 0) continue;
      merge(valuesOf(piece.props), RUN_KEYS);
      anyText = true;
    }
  }
  if (!anyText && list.length) {
    // Nothing but paragraph marks in the range: the formatting of the first of them.
    const paragraph = list[0].paragraph;
    merge(runValues(doc, resolveRunProps(doc, paragraph.props, markRunProps(paragraph.props.markProps), tableContextOf(doc, story, list[0].path)), markRunProps(paragraph.props.markProps)), RUN_KEYS);
  }
  (out as { mixed: string[] }).mixed = mixed;
  return out as unknown as RangeFormat;
}

export type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

/**
 * Turns a property on for the whole range, or off when all of the range has
 * it already. Turning off removes the property from the text; where a style
 * still turns it on, it is switched off on the text itself.
 */
export function toggleRunFormat(doc: DocModel, range: DocRange, key: ToggleKey): DocEdit {
  if (key !== 'bold' && key !== 'italic' && key !== 'underline' && key !== 'strike') bad('The property cannot be toggled');
  const on = rangeFormat(doc, range)[key] === true;
  if (!on) return setRunFormat(doc, range, { [key]: true });
  return buildEdit(doc, (draft) => {
    if (isCollapsed(range)) return range;
    mapRunProps(draft, range, (props, paragraph, where) => {
      const removed = patchRunProps(props, { [key]: null });
      const still = runValues(doc, resolveRunProps(doc, paragraph.props, removed, where), removed)[key] === true;
      return still ? patchRunProps(props, { [key]: false }) : removed;
    });
    return range;
  });
}

// ------------------------------------------------------------------ paragraphs

type Nullable<T> = { [K in keyof T]?: T[K] | null };

/** A change of paragraph formatting; `null` removes a property so that the style decides again. */
export interface ParaFormatPatch {
  align?: string | null;
  /** Laid over the indents the paragraph states; null removes all of them. */
  indent?: Nullable<Indent> | null;
  spacing?: Nullable<ParaSpacing> | null;
  /** Paragraph style id. */
  style?: string | null;
  /** numId 0 takes the paragraph out of a list its style puts it in. */
  numbering?: { numId: number; level?: number } | null;
  keepNext?: boolean | null;
  keepLines?: boolean | null;
  pageBreakBefore?: boolean | null;
  widowControl?: boolean | null;
  contextualSpacing?: boolean | null;
  outlineLevel?: number | null;
}

const PARA_PATCH_KEYS = ['align', 'style', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing', 'outlineLevel'] as const;

function patchNested<T extends object>(base: T | undefined, patch: Nullable<T> | null): T | undefined {
  if (patch === null) return undefined;
  const out = { ...base } as Record<string, unknown>;
  for (const key of Object.keys(patch)) {
    const value = (patch as Record<string, unknown>)[key];
    if (value === undefined) continue;
    if (value === null) delete out[key];
    else if (typeof value === 'number' ? value !== value || Math.abs(value) > 1000000 : typeof value !== 'boolean' && typeof value !== 'string') bad('The paragraph format is not valid');
    else out[key] = value;
  }
  return Object.keys(out).length ? (out as T) : undefined;
}

/** Paragraph properties with a patch laid over them: the same object when nothing changes. */
export function patchParaProps(props: ParaProps, patch: ParaFormatPatch): ParaProps {
  const out: ParaProps = { ...props };
  const target = out as Record<string, unknown>;
  for (const key of PARA_PATCH_KEYS) {
    const value = patch[key];
    if (value === undefined) continue;
    if (value === null) delete target[key];
    else target[key] = value;
  }
  if (typeof patch.align === 'string' && !/^[A-Za-z]{1,32}$/.test(patch.align)) bad('The alignment is not valid');
  if (typeof patch.style === 'string' && patch.style === '') bad('The style id is not valid');
  if (typeof patch.outlineLevel === 'number' && !(patch.outlineLevel >= 0 && patch.outlineLevel <= 9 && Math.floor(patch.outlineLevel) === patch.outlineLevel)) bad('The outline level is not valid');
  if (patch.indent !== undefined) {
    const indent = patchNested(props.indent, patch.indent);
    if (indent) {
      // A first line indent and a hanging indent replace each other.
      if (patch.indent && typeof patch.indent.firstLine === 'number' && patch.indent.hanging === undefined) delete indent.hanging;
      if (patch.indent && typeof patch.indent.hanging === 'number' && patch.indent.firstLine === undefined) delete indent.firstLine;
      out.indent = indent;
    } else delete out.indent;
  }
  if (patch.spacing !== undefined) {
    const spacing = patchNested(props.spacing, patch.spacing);
    if (spacing) out.spacing = spacing;
    else delete out.spacing;
  }
  if (patch.numbering !== undefined) {
    const n = patch.numbering;
    if (n === null) delete out.numbering;
    else {
      const whole = (value: number, max: number) => typeof value === 'number' && value >= 0 && value <= max && Math.floor(value) === value;
      if (!whole(n.numId, 0x7fffffff) || (n.level !== undefined && !whole(n.level, 8))) bad('The numbering is not valid');
      out.numbering = n.level === undefined ? { numId: n.numId } : { numId: n.numId, level: n.level };
    }
  }
  if (sameProps(out, props)) return props;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

function mapParaProps(draft: Draft, range: DocRange, change: (props: ParaProps, paragraph: Paragraph, where: TableContext | undefined, path: number[]) => ParaProps): void {
  editable(range);
  const { story, list } = coveredParagraphs(draft.doc, range);
  for (const part of list) {
    const next = change(part.paragraph.props, part.paragraph, tableContextOf(draft.doc, story, part.path), part.path);
    if (next !== part.paragraph.props) draft.setParagraph(story, part.path, paragraphWith(part.paragraph, part.paragraph.children, next));
  }
}

/** Sets paragraph formatting on every paragraph a range touches. */
export function setParaFormat(doc: DocModel, at: DocRange | DocPosition, patch: ParaFormatPatch): DocEdit {
  const range = asRange(at);
  return buildEdit(doc, (draft) => {
    patchParaProps(EMPTY_PARA_PROPS, patch);
    mapParaProps(draft, range, (props) => patchParaProps(props, patch));
    return range;
  });
}

/** Paragraph properties that take a paragraph out of its list: the numbering removed, or switched off when a style brings it. */
export function withoutNumbering(doc: DocModel, props: ParaProps, where: TableContext | undefined): ParaProps {
  const out: ParaProps = { ...props };
  delete out.numbering;
  const fromStyle = resolveParaProps(doc, out, where).numbering;
  if (fromStyle && fromStyle.numId !== undefined && fromStyle.numId !== 0) out.numbering = { numId: 0 };
  if (sameProps(out, props)) return props;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

/** Is a paragraph with these properties an item of a list. */
export function inList(doc: DocModel, props: ParaProps, where: TableContext | undefined): boolean {
  return listKind(doc, resolveParaProps(doc, props, where)) !== 'none';
}

const BULLETS: [text: string, font: string][] = [
  ['', 'Symbol'],
  ['o', 'Courier New'],
  ['', 'Wingdings'],
];
const NUMBER_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'];

/** The nine levels of a list as Word makes them: half an inch deeper per level, the number hanging a quarter inch. */
function listLevels(kind: 'bullet' | 'number'): Map<number, NumLevel> {
  const levels = new Map<number, NumLevel>();
  for (let i = 0; i < 9; i++) {
    const level: NumLevel = { level: i, start: 1, align: 'left', para: { indent: { left: 720 * (i + 1), hanging: 360 } } };
    if (kind === 'bullet') {
      const [text, font] = BULLETS[i % 3];
      level.format = 'bullet';
      level.text = text;
      level.run = { fonts: { ascii: font, hAnsi: font, hint: 'default' } };
      if (font === 'Courier New') level.run.fonts!.cs = font;
    } else {
      level.format = NUMBER_FORMATS[i % 3];
      level.text = `%${i + 1}.`;
      if (i % 3 === 2) {
        level.align = 'right';
        level.para = { indent: { left: 720 * (i + 1), hanging: 180 } };
      }
    }
    levels.set(i, level);
  }
  return levels;
}

function nextId(map: Map<number, unknown>, least: number): number {
  let id = least;
  map.forEach((_value, key) => {
    if (key >= id) id = key + 1;
  });
  if (id > 0x7ffffffe) bad('The document has no room for another list');
  return id;
}

/** The nearest paragraph before (step -1) or after (1) the block at `index` of the list of blocks at `container`; only blocks without content are passed over. */
function siblingParagraph(doc: DocModel, story: string, path: number[], step: 1 | -1): { paragraph: Paragraph; path: number[] } | undefined {
  const container = path.slice(0, -1);
  const blocks = containerBlocks(doc, story, container);
  for (let i = path[path.length - 1] + step, tries = 0; i >= 0 && i < blocks.length && tries < 64; i += step, tries++) {
    const block = blocks[i];
    if (block.type === 'paragraph') return { paragraph: block, path: container.concat([i]) };
    if (block.type === 'table') return undefined;
  }
  return undefined;
}

/**
 * Makes the paragraphs of a range items of a bulleted or numbered list, or
 * takes them out of it when all of them are items of such a list already.
 * The list of a neighbouring paragraph of the same kind is continued;
 * otherwise the document gets a new list definition.
 */
export function toggleList(doc: DocModel, at: DocRange | DocPosition, kind: 'bullet' | 'number'): DocEdit {
  const range = asRange(at);
  if (kind !== 'bullet' && kind !== 'number') bad('The kind of list is not valid');
  return buildEdit(doc, (draft) => {
    editable(range);
    const { story, list } = coveredParagraphs(doc, range);
    const kinds = list.map((part) => listKind(doc, resolveParaProps(doc, part.paragraph.props, tableContextOf(doc, story, part.path))));
    if (kinds.every((value) => value === kind)) {
      mapParaProps(draft, range, (props, _paragraph, where) => withoutNumbering(doc, props, where));
      return range;
    }
    let numId: number | undefined;
    for (const step of [-1, 1] as const) {
      const beside = numId === undefined ? siblingParagraph(doc, story, step === -1 ? list[0].path : list[list.length - 1].path, step) : undefined;
      if (!beside) continue;
      const effective = resolveParaProps(doc, beside.paragraph.props, tableContextOf(doc, story, beside.path));
      if (listKind(doc, effective) === kind) numId = (effective.numbering as { numId: number }).numId;
    }
    if (numId === undefined) {
      const abstract: AbstractNum = { id: nextId(doc.numbering.abstracts, 0), levels: listLevels(kind), multiLevelType: 'hybridMultilevel' };
      numId = nextId(doc.numbering.nums, 1);
      draft.apply({ type: 'setNumbering', abstracts: [[abstract.id, abstract]], nums: [[numId, { id: numId, abstractId: abstract.id, overrides: new Map() }]] });
    }
    const id = numId;
    mapParaProps(draft, range, (props, _paragraph, where) => {
      const effective = resolveParaProps(doc, props, where);
      // An item of another list keeps its depth.
      const level = listKind(doc, effective) !== 'none' && effective.numbering && effective.numbering.level !== undefined ? effective.numbering.level : 0;
      return patchParaProps(props, { numbering: { numId: id, level } });
    });
    return range;
  });
}

/** One step of indent outside lists, twips: half an inch. */
const INDENT_STEP = 720;

/**
 * Moves the list items of a range `delta` levels deeper (levels 0 to 8);
 * paragraphs outside lists are indented by half an inch per step instead.
 */
export function indentList(doc: DocModel, at: DocRange | DocPosition, delta: number): DocEdit {
  const range = asRange(at);
  if (typeof delta !== 'number' || Math.floor(delta) !== delta || Math.abs(delta) > 8) bad('The change of level is not valid');
  return buildEdit(doc, (draft) => {
    mapParaProps(draft, range, (props, _paragraph, where) => {
      const effective = resolveParaProps(doc, props, where);
      if (listKind(doc, effective) !== 'none') {
        const numbering = effective.numbering as { numId: number; level?: number };
        const level = Math.max(0, Math.min(8, (numbering.level === undefined ? 0 : numbering.level) + delta));
        return patchParaProps(props, { numbering: { numId: numbering.numId, level } });
      }
      const left = Math.max(0, ((effective.indent && effective.indent.left) || 0) + delta * INDENT_STEP);
      // Back at the indent the style gives, the paragraph states none of its own.
      const bare = patchParaProps(props, { indent: { left: null } });
      const inherited = resolveParaProps(doc, bare, where).indent;
      return left === ((inherited && inherited.left) || 0) ? bare : patchParaProps(props, { indent: { left } });
    });
    return range;
  });
}
