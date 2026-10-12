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
import { buildEdit, tableContextOf, type Draft, type DocEdit, type DocEditOptions } from './draft';
import { boundary, flatten, inheritedProps, markRunProps, paragraphWith, rebuild, sameInlines, sameProps, type Piece } from './pieces';
import { asRange, containerBlocks, inTextBox, isCollapsed, orderRange, paragraphLength, paragraphsBetween, resolvePosition, storyOf, type DocPosition, type DocRange } from './position';
import { inSessionInsertion, isOwnParagraph, trackedParaProps, trackedRunProps } from './track';

// Patches come from the caller of the editor, who may hand over anything: every field is checked before it reaches
// the model, where the writer would put it into the file.

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

const RUN_PATCH_KEYS = new Set<string>(['bold', 'italic', 'underline', 'strike', 'font', 'size', 'color', 'highlight', 'vertAlign', 'style']);

/** ST_Underline. */
const UNDERLINES = new Set<string>([
  'single',
  'words',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
  'none',
]);

/** ST_HighlightColor. */
const HIGHLIGHTS = new Set<string>([
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'none',
]);

const COLOR_KEYS = new Set<string>(['rgb', 'auto', 'theme', 'tint', 'shade']);

/** A patch is an object of known properties. */
function checkPatch(patch: unknown, keys: Set<string>, what: string): void {
  if (patch === null || typeof patch !== 'object' || Array.isArray(patch)) bad(`A change of ${what} formatting must be an object`);
  for (const key of Object.keys(patch as object)) if (!keys.has(key)) bad(`"${key}" is not a property of ${what} formatting that can be set`);
}

function checkFlag(value: unknown, name: string): void {
  if (value !== undefined && value !== null && typeof value !== 'boolean') bad(`"${name}" must be true, false or null`);
}

const isWhole = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && value >= min && value <= max && Math.floor(value) === value;

/** A colour of a patch as a colour of the model: a copy, checked field by field. */
function patchColor(color: unknown): DocColor {
  if (typeof color === 'string') {
    if (color === 'auto') return { auto: true };
    if (/^[0-9A-Fa-f]{6}$/.test(color)) return { rgb: color.toUpperCase() };
    return bad('The colour is not valid');
  }
  if (color === null || typeof color !== 'object' || Array.isArray(color)) return bad('The colour is not valid');
  const given = color as Record<string, unknown>;
  const out: DocColor = {};
  for (const key of Object.keys(given)) {
    const value = given[key];
    if (value === undefined) continue;
    if (!COLOR_KEYS.has(key)) return bad('The colour is not valid');
    if (key === 'rgb') {
      if (typeof value !== 'string' || !/^[0-9A-Fa-f]{6}$/.test(value)) return bad('The colour is not valid');
      out.rgb = value.toUpperCase();
    } else if (key === 'auto') {
      if (typeof value !== 'boolean') return bad('The colour is not valid');
      if (value) out.auto = true;
    } else if (key === 'theme') {
      if (typeof value !== 'string' || !/^[A-Za-z0-9]{1,32}$/.test(value)) return bad('The colour is not valid');
      out.theme = value;
    } else {
      if (!isWhole(value, 0, 255)) return bad('The colour is not valid');
      out[key as 'tint' | 'shade'] = value;
    }
  }
  if (out.rgb === undefined && out.auto === undefined && out.theme === undefined) return bad('The colour is not valid');
  return out;
}

/**
 * Run properties with a patch laid over them: the same object when nothing
 * changes. Every field of the patch is checked; what is not a valid value,
 * and a property the patch cannot set, throws RangeError. With `doc`, a
 * character style has to be one of the document.
 */
export function patchRunProps(props: RunProps, patch: RunFormatPatch, doc?: DocModel): RunProps {
  checkPatch(patch, RUN_PATCH_KEYS, 'character');
  checkFlag(patch.bold, 'bold');
  checkFlag(patch.italic, 'italic');
  checkFlag(patch.strike, 'strike');
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
    if (u !== null && typeof u !== 'boolean' && (typeof u !== 'string' || !UNDERLINES.has(u))) bad('The kind of underline is not valid');
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
    // Half-points: 1 pt to 1638 pt, the sizes Word takes.
    if (size !== null && !isWhole(size, 2, 3276)) bad('The font size is not valid');
    set('size', size);
    set('sizeCs', size);
  }
  if (patch.color !== undefined) {
    if (patch.color === null) delete out.color;
    else out.color = patchColor(patch.color);
  }
  if (patch.highlight !== undefined) {
    if (patch.highlight !== null && (typeof patch.highlight !== 'string' || !HIGHLIGHTS.has(patch.highlight))) bad('The highlight colour is not valid');
    set('highlight', patch.highlight);
  }
  if (patch.vertAlign !== undefined) {
    const v = patch.vertAlign;
    if (v !== null && v !== 'baseline' && v !== 'superscript' && v !== 'subscript') bad('The vertical alignment is not valid');
    set('vertAlign', v);
  }
  if (patch.style !== undefined) {
    if (patch.style !== null) checkStyle(doc, patch.style, 'character');
    set('style', patch.style);
  }
  if (sameProps(out, props)) return props;
  return Object.keys(out).length ? out : EMPTY_RUN_PROPS;
}

/** A style id of a patch: a text, and with `doc` the id of a style of the right type. */
function checkStyle(doc: DocModel | undefined, id: unknown, type: 'character' | 'paragraph'): void {
  if (typeof id !== 'string' || id === '' || id.length > 255) bad('The style id is not valid');
  if (!doc) return;
  const style = doc.styles.styles.get(id as string);
  if (!style || style.type !== type) bad(`The document has no ${type} style with this id`);
}

/** One paragraph of a range, with the part of it the range covers. */
interface Covered {
  path: number[];
  paragraph: Paragraph;
  from: number;
  to: number;
  /** DocPosition.before of the ends of the range, where they are in this paragraph. */
  fromBefore: boolean;
  toBefore: boolean;
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
    const paragraph = paragraphAt(doc, story, path);
    const last = i === paths.length - 1;
    list.push({
      path,
      paragraph,
      from: i === 0 ? start.offset : 0,
      to: last ? end.offset : paragraphLength(paragraph),
      fromBefore: i === 0 && start.before === true,
      toBefore: last && end.before === true,
      mark: !last,
    });
  }
  return { story, list };
}

/** The paragraph at a path as the document has it now. */
function paragraphAt(doc: DocModel, story: string, path: number[]): Paragraph {
  const blocks = containerBlocks(doc, story, path.slice(0, -1));
  return blocks[path[path.length - 1]] as Paragraph;
}

const editable = (range: DocRange): void => {
  if (inTextBox(range.anchor.path) || inTextBox(range.focus.path)) bad('The content of a text box cannot be edited');
};

/** May this piece take character formatting: shown content of a run, not the marks of a field. */
const formattable = (piece: Piece): boolean => piece.item !== undefined && !piece.code && !piece.frozen && piece.item.type !== 'fieldChar';

/** Applies a change to the formatting of a paragraph mark; returns the properties of the paragraph, the same object when nothing changes. */
function withMark(props: ParaProps, next: RunProps): ParaProps {
  const out: ParaProps = { ...props, markProps: next };
  if (next === EMPTY_RUN_PROPS || Object.keys(next).length === 0) delete out.markProps;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

/**
 * Changes the character formatting of a range. Where the draft tracks, the change is a tracked one: see
 * trackedRunProps for what gets a record of the formatting it had.
 */
function mapRunProps(draft: Draft, range: DocRange, change: (props: RunProps, paragraph: Paragraph, where: TableContext | undefined) => RunProps): void {
  editable(range);
  const track = draft.track;
  const host = track ? storyOf(draft.doc, range.anchor.story).host : undefined;
  /** The properties of a paragraph mark after the change. */
  const markChanged = (paragraph: Paragraph, mark: RunProps, next: RunProps): RunProps => (track && host ? trackedRunProps(track, host, mark, next, isOwnParagraph(paragraph, track)) : next);
  if (isCollapsed(range)) {
    // A caret. In an empty paragraph the formatting goes onto the paragraph mark, which text typed there takes
    // over (what Word does); in a paragraph with text there is nothing to format: the caller keeps the format
    // for what is typed next (see insertText).
    const at = resolvePosition(draft.doc, range.anchor);
    if (at.length !== 0) return;
    const paragraph = at.paragraph;
    const mark = paragraph.props.markProps || EMPTY_RUN_PROPS;
    const next = markChanged(paragraph, mark, change(mark, paragraph, tableContextOf(draft.doc, range.anchor.story, range.anchor.path)));
    if (next !== mark) draft.setParagraph(range.anchor.story, range.anchor.path, paragraphWith(paragraph, paragraph.children, withMark(paragraph.props, next)));
    return;
  }
  const { story, list } = coveredParagraphs(draft.doc, range);
  for (const part of list) {
    // As the draft has it now: an earlier step of this command may have put another object there (see Draft.replace).
    const paragraph = paragraphAt(draft.doc, story, part.path);
    const where = tableContextOf(draft.doc, story, part.path);
    const pieces = flatten(paragraph);
    const from = boundary(pieces, part.from, 'start', part.fromBefore);
    const to = boundary(pieces, part.to, 'caret', part.toBefore);
    // Runs share their properties: each bag is changed once.
    const memo = new Map<RunProps, RunProps>();
    const changed = (props: RunProps): RunProps => {
      let next = memo.get(props);
      if (!next) memo.set(props, (next = change(props, paragraph, where)));
      return next;
    };
    // Tracked, the runs of one paragraph that share their properties share the record too, and the draft gives every
    // run an id of its own (see distinctRevisions). It does so within one paragraph only, so a record is never shared
    // between paragraphs: each makes its own. Text in an insertion of the author gets none (see trackedRunProps).
    const recorded = new Map<RunProps, RunProps>();
    const mine = new Map<RunProps, RunProps>();
    for (let q = from; q < to; q++) {
      const piece = pieces[q];
      if (!formattable(piece)) continue;
      let next = changed(piece.props);
      if (track && host && next !== piece.props) {
        const own = inSessionInsertion(piece.wrap, track);
        const made = own ? mine : recorded;
        const known = made.get(piece.props);
        if (known) next = known;
        else made.set(piece.props, (next = trackedRunProps(track, host, piece.props, next, own)));
      }
      if (next !== piece.props) pieces[q] = { ...piece, props: next };
    }
    let props = paragraph.props;
    if (part.mark) {
      // The paragraph mark is in the range: what is typed at the end of the paragraph is formatted alike.
      const mark = props.markProps || EMPTY_RUN_PROPS;
      const next = markChanged(paragraph, mark, change(mark, paragraph, where));
      if (next !== mark) props = withMark(props, next);
    }
    const children = rebuild(pieces);
    if (props === paragraph.props && sameInlines(children, paragraph.children)) continue;
    draft.setParagraph(story, part.path, paragraphWith(paragraph, children, props));
  }
}

/**
 * Sets character formatting on the text of a range. At a caret (a collapsed
 * range) in an empty paragraph the formatting is set on the paragraph mark,
 * so that text typed there takes it; at a caret in a paragraph with text
 * nothing changes: the caller keeps the format for what is typed next (see
 * insertText). An end of the range with `before` set leaves out the picture
 * at its offset, a start with it takes that picture in.
 */
export function setRunFormat(doc: DocModel, range: DocRange, patch: RunFormatPatch, options?: DocEditOptions): DocEdit {
  const at = asRange(range);
  return buildEdit(
    doc,
    (draft) => {
      patchRunProps(EMPTY_RUN_PROPS, patch, doc);
      // The change does not depend on the paragraph: each bag of properties is patched once for the whole range.
      const memo = new Map<RunProps, RunProps>();
      mapRunProps(draft, at, (props) => {
        let next = memo.get(props);
        if (!next) memo.set(props, (next = patchRunProps(props, patch, doc)));
        return next;
      });
      return at;
    },
    options,
    true,
  );
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
      merge(valuesOf(inheritedProps(pieces, boundary(pieces, part.from, 'caret', part.fromBefore), paragraph.props.markProps)), RUN_KEYS);
      anyText = true;
      continue;
    }
    const from = boundary(pieces, part.from, 'start', part.fromBefore);
    const to = boundary(pieces, part.to, 'caret', part.toBefore);
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
 * still turns it on, it is switched off on the text itself. At a caret see
 * setRunFormat.
 */
export function toggleRunFormat(doc: DocModel, range: DocRange, key: ToggleKey, options?: DocEditOptions): DocEdit {
  if (key !== 'bold' && key !== 'italic' && key !== 'underline' && key !== 'strike') bad('The property cannot be toggled');
  const at = asRange(range);
  const on = rangeFormat(doc, at)[key] === true;
  if (!on) return setRunFormat(doc, at, { [key]: true }, options);
  return buildEdit(
    doc,
    (draft) => {
      mapRunProps(draft, at, (props, paragraph, where) => {
        const removed = patchRunProps(props, { [key]: null });
        const still = runValues(doc, resolveRunProps(doc, paragraph.props, removed, where), removed)[key] === true;
        return still ? patchRunProps(props, { [key]: false }) : removed;
      });
      return at;
    },
    options,
    true,
  );
}

// ------------------------------------------------------------------ paragraphs

type Nullable<T> = { [K in keyof T]?: T[K] | null };

/** A change of paragraph formatting; `null` removes a property so that the style decides again. */
export interface ParaFormatPatch {
  /** "left", "center", "right", "both" (justified), "start", "end", "distribute", ... as in ST_Jc. */
  align?: string | null;
  /** Laid over the indents the paragraph states, twips; null removes all of them. */
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

const PARA_FLAGS = ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing'] as const;
const PARA_PATCH_KEYS = new Set<string>(['align', 'indent', 'spacing', 'style', 'numbering', 'outlineLevel', ...PARA_FLAGS]);

/** ST_Jc. */
const ALIGNMENTS = new Set<string>(['left', 'center', 'right', 'both', 'start', 'end', 'distribute', 'mediumKashida', 'highKashida', 'lowKashida', 'thaiDistribute', 'numTab']);

/** The widest page of Word, twips: no indent or spacing is larger. */
const MAX_PARA_TWIPS = 31680;

/** Fields of an indent a patch can set, each with the least value it takes. */
const INDENT_FIELDS: Record<string, number> = {
  left: -MAX_PARA_TWIPS,
  right: -MAX_PARA_TWIPS,
  firstLine: 0,
  hanging: 0,
  leftChars: -MAX_PARA_TWIPS,
  rightChars: -MAX_PARA_TWIPS,
  firstLineChars: 0,
  hangingChars: 0,
};

const SPACING_FIELDS: Record<string, number> = { before: 0, after: 0, line: 0, beforeLines: 0, afterLines: 0 };

const own = (record: object, key: string): boolean => Object.prototype.hasOwnProperty.call(record, key);

/** A group of properties with a patch laid over it; undefined when nothing is left. */
function patchNested<T extends object>(base: T | undefined, patch: unknown, check: (key: string, value: unknown) => boolean): T | undefined {
  if (patch === null) return undefined;
  if (typeof patch !== 'object' || Array.isArray(patch)) return bad('The paragraph format is not valid');
  const out = { ...base } as Record<string, unknown>;
  for (const key of Object.keys(patch as object)) {
    const value = (patch as Record<string, unknown>)[key];
    if (!check(key, value)) return bad('The paragraph format is not valid');
    if (value === undefined) continue;
    if (value === null) delete out[key];
    else out[key] = value;
  }
  return Object.keys(out).length ? (out as T) : undefined;
}

const indentField = (key: string, value: unknown): boolean => own(INDENT_FIELDS, key) && (value === undefined || value === null || isWhole(value, INDENT_FIELDS[key], MAX_PARA_TWIPS));

function spacingField(key: string, value: unknown): boolean {
  if (value === undefined || value === null) return own(SPACING_FIELDS, key) || key === 'lineRule' || key === 'beforeAuto' || key === 'afterAuto';
  if (key === 'lineRule') return value === 'auto' || value === 'exact' || value === 'atLeast';
  if (key === 'beforeAuto' || key === 'afterAuto') return typeof value === 'boolean';
  return own(SPACING_FIELDS, key) && isWhole(value, SPACING_FIELDS[key], MAX_PARA_TWIPS);
}

/**
 * Paragraph properties with a patch laid over them: the same object when
 * nothing changes. Every field of the patch is checked; what is not a valid
 * value, and a property the patch cannot set, throws RangeError. With `doc`,
 * a style has to be a paragraph style of the document and a list one of its
 * lists.
 */
export function patchParaProps(props: ParaProps, patch: ParaFormatPatch, doc?: DocModel): ParaProps {
  checkPatch(patch, PARA_PATCH_KEYS, 'paragraph');
  const out: ParaProps = { ...props };
  for (const key of PARA_FLAGS) {
    const value = patch[key];
    checkFlag(value, key);
    if (value === undefined) continue;
    if (value === null) delete out[key];
    else out[key] = value;
  }
  if (patch.align !== undefined) {
    if (patch.align === null) delete out.align;
    else if (typeof patch.align !== 'string' || !ALIGNMENTS.has(patch.align)) bad('The alignment is not valid');
    else out.align = patch.align;
  }
  if (patch.style !== undefined) {
    if (patch.style === null) delete out.style;
    else {
      checkStyle(doc, patch.style, 'paragraph');
      out.style = patch.style;
    }
  }
  if (patch.outlineLevel !== undefined) {
    if (patch.outlineLevel === null) delete out.outlineLevel;
    else if (!isWhole(patch.outlineLevel, 0, 9)) bad('The outline level is not valid');
    else out.outlineLevel = patch.outlineLevel;
  }
  if (patch.indent !== undefined) {
    const indent = patchNested(props.indent, patch.indent, indentField);
    if (indent) {
      // A first line indent and a hanging indent replace each other.
      if (patch.indent && typeof patch.indent.firstLine === 'number' && patch.indent.hanging === undefined) delete indent.hanging;
      if (patch.indent && typeof patch.indent.hanging === 'number' && patch.indent.firstLine === undefined) delete indent.firstLine;
      out.indent = indent;
    } else delete out.indent;
  }
  if (patch.spacing !== undefined) {
    const spacing = patchNested(props.spacing, patch.spacing, spacingField);
    if (spacing) out.spacing = spacing;
    else delete out.spacing;
  }
  if (patch.numbering !== undefined) {
    const n = patch.numbering as unknown;
    if (n === null) delete out.numbering;
    else {
      if (typeof n !== 'object' || Array.isArray(n)) return bad('The numbering is not valid');
      const given = n as { numId: unknown; level?: unknown };
      for (const key of Object.keys(given)) if (key !== 'numId' && key !== 'level') return bad('The numbering is not valid');
      if (!isWhole(given.numId, 0, 0x7fffffff) || (given.level !== undefined && !isWhole(given.level, 0, 8))) return bad('The numbering is not valid');
      if (doc && given.numId !== 0 && !doc.numbering.nums.has(given.numId)) return bad('The document has no list with this id');
      out.numbering = given.level === undefined ? { numId: given.numId } : { numId: given.numId, level: given.level };
    }
  }
  if (sameProps(out, props)) return props;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

/** Changes the formatting of the paragraphs of a range. Where the draft tracks, each change is a tracked one: see trackedParaProps. */
function mapParaProps(draft: Draft, range: DocRange, change: (props: ParaProps, paragraph: Paragraph, where: TableContext | undefined, path: number[]) => ParaProps): void {
  editable(range);
  const { story, list } = coveredParagraphs(draft.doc, range);
  const track = draft.track;
  const host = track ? storyOf(draft.doc, story).host : undefined;
  for (const part of list) {
    // As the draft has it now: an earlier step of this command may have put another object there (see Draft.replace).
    const paragraph = paragraphAt(draft.doc, story, part.path);
    let next = change(paragraph.props, paragraph, tableContextOf(draft.doc, story, part.path), part.path);
    if (track && host) next = trackedParaProps(track, host, paragraph.props, next, isOwnParagraph(paragraph, track));
    if (next !== paragraph.props) draft.setParagraph(story, part.path, paragraphWith(paragraph, paragraph.children, next));
  }
}

/** Sets paragraph formatting on every paragraph a range touches. */
export function setParaFormat(doc: DocModel, at: DocRange | DocPosition, patch: ParaFormatPatch, options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  return buildEdit(
    doc,
    (draft) => {
      patchParaProps(EMPTY_PARA_PROPS, patch, doc);
      mapParaProps(draft, range, (props) => patchParaProps(props, patch, doc));
      return range;
    },
    options,
    true,
  );
}

/** Paragraph properties that take a paragraph out of its list: the numbering removed, or switched off when a style brings it. */
export function withoutNumbering(doc: DocModel, props: ParaProps, where: TableContext | undefined): ParaProps {
  const out: ParaProps = { ...props };
  delete out.numbering;
  const fromStyle = resolveParaProps(doc, out, where).numbering;
  if (fromStyle && fromStyle.numId !== undefined && fromStyle.numId !== 0) out.numbering = { numId: 0 };
  else if (out.extras) {
    // The <w:numPr> the model kept for what it could not read of it (a recorded change of the number) would be
    // written back as it is, list and all, once no field stands for it: it goes with the numbering.
    const extras = out.extras.filter((extra) => extra.index < 0 || extra.el.name.slice(extra.el.name.indexOf(':') + 1) !== 'numPr');
    if (extras.length !== out.extras.length) {
      if (extras.length) out.extras = extras;
      else delete out.extras;
    }
  }
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
export function toggleList(doc: DocModel, at: DocRange | DocPosition, kind: 'bullet' | 'number', options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  if (kind !== 'bullet' && kind !== 'number') bad('The kind of list is not valid');
  return buildEdit(
    doc,
    (draft) => {
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
    },
    options,
    true,
  );
}

/** One step of indent outside lists, twips: half an inch. */
const INDENT_STEP = 720;

/**
 * Moves the list items of a range `delta` levels deeper (levels 0 to 8);
 * paragraphs outside lists are indented by half an inch per step instead.
 */
export function indentList(doc: DocModel, at: DocRange | DocPosition, delta: number, options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  if (typeof delta !== 'number' || Math.floor(delta) !== delta || Math.abs(delta) > 8) bad('The change of level is not valid');
  return buildEdit(
    doc,
    (draft) => {
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
    },
    options,
    true,
  );
}
