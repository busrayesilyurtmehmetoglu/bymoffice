// A change of a document as it travels between those who edit it together: read from what
// arrived (nothing that arrives is trusted), written for the relay, and run on a document.
//
// A change travels as the command that made it and the place it was made at. Every
// participant runs the command on its own document, which is the document its author ran it
// on (the relay sees to that), so that all end up with the same. After it ran, what it did to
// the places around it is read off the document itself (marksAfter): text came in here, text
// went there. Where that cannot be read off with certainty the change is only known as "these
// blocks were replaced" or "something in this story", and gives way to, or makes give way,
// whatever else is about the same (see rebase.ts).

import { EMPTY_PARA_PROPS, EMPTY_RUN_PROPS, paragraphText, type AttrList, type Block, type Comment, type DocColor, type DocModel, type DocRelationship, type ExtraXml, type Note, type StoryCollection, type Inline, type Language, type Paragraph, type ParaProps, type Run, type RunFonts, type RunItem, type RunProps, type TabStop } from '../../model/doc/types';
import type { DocEdit } from '../../ops/doc/draft';
import { indentList, patchParaProps, patchRunProps, rangeFormat, setParaFormat, setRunFormat, toggleList, toggleRunFormat, type ParaFormatPatch, type RunFormatPatch, type ToggleKey } from '../../ops/doc/format';
import type { DocOp } from '../../ops/doc/ops';
import { containerBlocks, eachGraphic, storyIds, storyOf, textBoxRootLength, type DocPosition, type DocRange } from '../../ops/doc/position';
import { editInTextBox, readInTextBox } from '../../ops/doc/text-box';
import { applyDocOp } from '../../ops/doc/ops';
import { deleteColumn, deleteRow, insertColumn, insertHyperlink, insertRow, insertTable } from '../../ops/doc/structure';
import { deleteBackward, deleteForward, deleteRange, insertPageBreak, insertText, splitParagraph } from '../../ops/doc/text';
import { WireError } from '../types';
import type { XmlElement } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { commentsForWriting } from '../../formats/docx/comments';
import { bodyText, collectionText, createBodyContext, storyText } from '../../formats/docx/write-package';
import { treeText, writeBlocks, type BodyContext } from '../../formats/docx/write-body';
import { plainCopy, stable, xmlText } from '../wire';
import { ALL_DRAWINGS, MAX_BOX_DEPTH, type DocWire, type Mark, type PutInline, type PutItem, type PutParagraph, type WirePos } from './rebase';
import { MAX_STEP_COMMENTS, commentEdit, readComment, writeComment } from './comments';

/** Most commands of one step. */
export const MAX_STEP_WIRES = 64;
/** Deepest path of a place: tables in tables, three numbers a level. */
const MAX_PATH = 96;
const MAX_INDEX = 100000000;
/** Most characters of a text that is typed or pasted in one command (a step is bounded in bytes as well). */
const MAX_TEXT = 800000;
/** Most paragraph ends one command may bring, and most paragraphs one `put` may. */
const MAX_TEXT_PARAGRAPHS = 5000;
const MAX_PUT_PARAGRAPHS = 2000;
/** The largest table a command makes. */
const MAX_NEW_ROWS = 100;
const MAX_NEW_COLUMNS = 63;
/** Most blocks one list of a document may hold, and most table cells one story may, while it is edited together: a command that would go beyond is refused, by everybody alike. */
export const MAX_LIST_BLOCKS = 200000;
export const MAX_STORY_CELLS = 100000;
/** Most characters of a story while a document is edited together. */
export const MAX_STORY_CHARS_LIMIT = 20000000;

const fail = (what: string): never => {
  throw new WireError(`Invalid ${what}`);
};

const isInt = (value: unknown, low: number, high: number): value is number => typeof value === 'number' && Math.floor(value) === value && value >= low && value <= high;

export function readPos(value: unknown): WirePos {
  if (!Array.isArray(value) || (value.length !== 3 && value.length !== 4)) return fail('place');
  const [s, p, o, b] = value as unknown[];
  if (typeof s !== 'string' || s.length === 0 || s.length > 120 || !xmlText(s) || !/^(body|(header|footer|footnote|endnote):[^\s]{1,110})$/.test(s)) return fail('story');
  if (!Array.isArray(p) || p.length === 0 || p.length > MAX_PATH) return fail('path');
  const path: number[] = [];
  let boxes = 0;
  let last = -3;
  for (let i = 0; i < p.length; i++) {
    const index = p[i];
    // A text box: the paragraph that anchors it (in the story, or in the box before), -1, its drawing and its number
    // there, and a path in it (see boxBase).
    if (index === -1) {
      if (i < last + 4 || i + 3 >= p.length || ++boxes > MAX_BOX_DEPTH || !isInt(p[i + 1], 0, MAX_INDEX) || !isInt(p[i + 2], 0, MAX_INDEX)) return fail('path');
      last = i;
      path.push(-1);
      continue;
    }
    path.push(isInt(index, 0, MAX_INDEX) ? index : fail('path'));
  }
  if (!isInt(o, 0, MAX_INDEX)) return fail('offset');
  if (value.length === 4 && b !== 1) return fail('place');
  return value.length === 4 ? { s, p: path, o, b: true } : { s, p: path, o };
}

const writePos = (pos: WirePos): unknown[] => (pos.b ? [pos.s, pos.p, pos.o, 1] : [pos.s, pos.p, pos.o]);

function readPatch(value: unknown, what: string): Record<string, unknown> {
  const copied = plainCopy(value, what);
  if (copied === null || typeof copied !== 'object' || Array.isArray(copied)) return fail(what);
  return copied as Record<string, unknown>;
}

const TOGGLES: readonly ToggleKey[] = ['bold', 'italic', 'underline', 'strike'];

/** What a named command takes after its place, how it is run, and what is known of what it does. */
interface Command {
  /** Reads what the command was called with. */
  args(a: unknown[]): unknown[];
  /** `range`: the command takes a range; otherwise a place. */
  range?: true;
  run(doc: DocModel, at: DocPosition, range: DocRange, a: unknown[]): DocEdit;
  /** `text`: it puts text in or takes text away at its place; `touch`: it changes paragraphs without moving text; `table`: it changes the table its place is in. */
  does: 'text' | 'touch' | 'table';
  /** False: not run in a text box (the editor refuses it there, and so does everybody; see commandInTextBox). */
  box?: false;
}

const one = <T>(a: unknown[], allowed: readonly T[]): T[] => (a.length === 1 && allowed.indexOf(a[0] as T) >= 0 ? [a[0] as T] : fail('command'));
const none = (a: unknown[]): unknown[] => (a.length === 0 ? [] : fail('command'));

export const COMMANDS: { [name: string]: Command } = {
  toggleList: { args: (a) => one(a, ['bullet', 'number']), range: true, run: (doc, _at, range, a) => toggleList(doc, range, a[0] as 'bullet' | 'number'), does: 'touch' },
  indentList: { args: (a) => (a.length === 1 && isInt(a[0], -8, 8) ? [a[0]] : fail('command')), range: true, run: (doc, _at, range, a) => indentList(doc, range, a[0] as number), does: 'touch' },
  insertPageBreak: { args: none, run: (doc, at) => insertPageBreak(doc, at), does: 'text', box: false },
  insertTable: { args: (a) => (a.length === 2 && isInt(a[0], 1, MAX_NEW_ROWS) && isInt(a[1], 1, MAX_NEW_COLUMNS) ? [a[0], a[1]] : fail('command')), run: (doc, at, _range, a) => insertTable(doc, at, a[0] as number, a[1] as number), does: 'text' },
  insertRow: { args: (a) => one(a, ['above', 'below']), run: (doc, at, _range, a) => insertRow(doc, at, a[0] as 'above' | 'below'), does: 'table' },
  deleteRow: { args: none, run: (doc, at) => deleteRow(doc, at), does: 'table' },
  insertColumn: { args: (a) => one(a, ['left', 'right']), run: (doc, at, _range, a) => insertColumn(doc, at, a[0] as 'left' | 'right'), does: 'table' },
  deleteColumn: { args: none, run: (doc, at) => deleteColumn(doc, at), does: 'table' },
  deleteBackward: { args: (a) => one(a, ['char', 'word']), run: (doc, at, _range, a) => deleteBackward(doc, at, a[0] as 'char' | 'word'), does: 'text' },
  deleteForward: { args: (a) => one(a, ['char', 'word']), run: (doc, at, _range, a) => deleteForward(doc, at, a[0] as 'char' | 'word'), does: 'text' },
  insertHyperlink: {
    // The address and, where its author gave it one, the id of the relationship of the link: everybody makes the
    // same, whatever relationships each holds besides. Without one, it is made of the command itself.
    args: (a) =>
      (a.length === 1 || (a.length === 2 && typeof a[1] === 'string' && /^rId[A-Za-z0-9]{1,32}$/.test(a[1]) && !/[0-9]{10}/.test(a[1]))) && typeof a[0] === 'string' && a[0].length <= 2083 && xmlText(a[0]) ? a.slice() : fail('command'),
    range: true,
    run: (doc, at, range, a) => insertHyperlink(doc, range, a[0] as string, { relationshipId: typeof a[1] === 'string' ? a[1] : 'rIdL' + hashOf(stable([at, range.focus, a[0]])) }),
    does: 'text',
    box: false,
  },
};

/**
 * Whether a named command runs in a text box: the one rule of the editor (which refuses the others there before they
 * run) and of everybody in a room (who refuse a step that brings one). A page break and a link are not put into boxes.
 */
export function commandInTextBox(name: string): boolean {
  const command = Object.prototype.hasOwnProperty.call(COMMANDS, name) ? COMMANDS[name] : undefined;
  return command !== undefined && command.box !== false;
}

/**
 * Every kind of command of a document that travels: with COMMANDS, what the fingerprint of the
 * wire is taken from (see PROTOCOL_VERSION). A kind of DocWire that is left out here is an error
 * of the type checker.
 */
export const DOC_WIRE_KINDS: { readonly [K in DocWire['t']]: true } = { text: true, del: true, split: true, run: true, unrun: true, para: true, cmd: true, put: true, cmt: true };

function readWire(value: unknown): DocWire {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return fail('operation');
  const from = value as Record<string, unknown>;
  const has = (key: string): boolean => Object.prototype.hasOwnProperty.call(from, key) && from[key] !== undefined;
  switch (from.t) {
    case 'text': {
      // A tab and a line break (a vertical tab: Shift+Enter) are characters the command takes; everything else is what a file can hold as text.
      if (typeof from.x !== 'string' || from.x.length > MAX_TEXT || !xmlText(from.x.replace(/\u000b/g, ' '))) return fail('text');
      if (from.x.split('\n').length - 1 > MAX_TEXT_PARAGRAPHS) return fail('text');
      const out: DocWire = { t: 'text', at: readPos(from.at), x: from.x };
      if (has('to')) out.to = readPos(from.to);
      if (has('f')) out.f = readPatch(from.f, 'format') as RunFormatPatch;
      return out;
    }
    case 'del':
      return { t: 'del', at: readPos(from.at), to: readPos(from.to) };
    case 'split':
      return { t: 'split', at: readPos(from.at) };
    case 'run':
      return { t: 'run', at: readPos(from.at), to: readPos(from.to), f: readPatch(from.f, 'format') as RunFormatPatch };
    case 'unrun':
      return { t: 'unrun', at: readPos(from.at), to: readPos(from.to), k: TOGGLES.indexOf(from.k as ToggleKey) >= 0 ? (from.k as ToggleKey) : fail('format') };
    case 'para':
      return { t: 'para', at: readPos(from.at), to: readPos(from.to), f: readPatch(from.f, 'format') as ParaFormatPatch };
    case 'cmd': {
      const name = typeof from.name === 'string' && Object.prototype.hasOwnProperty.call(COMMANDS, from.name) ? from.name : fail('command');
      const command = COMMANDS[name];
      if (!Array.isArray(from.a) || from.a.length > 4) return fail('command');
      const out: DocWire = { t: 'cmd', name, at: readPos(from.at), a: command.args(plainCopy(from.a, 'command') as unknown[]) };
      if (command.range) out.to = readPos(from.to);
      else if (has('to')) return fail('command');
      return out;
    }
    case 'put': {
      const at = readPos(from.at);
      if (at.o !== 0 || at.b) return fail('place');
      if (!Array.isArray(from.x) || from.x.length === 0 || from.x.length > MAX_PUT_PARAGRAPHS) return fail('paragraphs');
      if (typeof from.h !== 'string' || !/^[0-9a-f]{8}$/.test(from.h)) return fail('paragraphs');
      const x = from.x.map(readPutParagraph);
      return { t: 'put', at, n: isInt(from.n, 1, MAX_PUT_PARAGRAPHS) ? from.n : fail('paragraphs'), x, h: from.h };
    }
    case 'cmt':
      return readComment(from);
  }
  return fail('operation');
}

// ------------------------------------------------------- paragraphs that travel

// A `put` carries paragraphs: what was there before a step, for the step to be taken back.
// Only plain paragraphs travel - runs of text, tabs and line breaks with the formatting the
// model has a field for - and every field is read by its kind. Whatever else a paragraph
// holds (a picture, a field, a link, a bookmark, a tracked change, XML the model only keeps)
// makes it one that does not travel: a step that would have to bring such a paragraph back
// cannot be undone while the document is edited together.

/** The kinds of marks of proofing a paragraph may hold. */
const PROOF_KINDS = ['spellStart', 'spellEnd', 'gramStart', 'gramEnd'];

/** A language (ST_Lang): a tag as Word writes them ("tr-TR", "en-US", "x-none"), or a language id of four hex digits. */
const LANGUAGE = /^(?:[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,3}|[xXiI](?:-[A-Za-z0-9]{1,8}){1,3}|[0-9A-Fa-f]{4})$/;

/**
 * What a run is formatted with, as it travels: a change of formatting (RunFormatPatch, read by
 * patchRunProps - the rule every change of formatting is read by) and, where the file states
 * them apart, the slots of the scripts that a change of formatting sets together: the bold,
 * italic and size of complex scripts, the four fonts, and the languages. Each of those is read
 * by the same rule as its plain counterpart; the languages are tags.
 */
interface PutRunFormat {
  f?: RunFormatPatch;
  cs?: { bold?: boolean; italic?: boolean; size?: number };
  fonts?: { ascii?: string; hAnsi?: string; eastAsia?: string; cs?: string; hint?: string };
  lang?: { val?: string; eastAsia?: string; bidi?: string };
}

const RUN_FORMAT_KEYS = ['f', 'cs', 'fonts', 'lang'];

/** Run properties of what travelled: RangeError for any value a change of formatting would not take. With `doc`, a style has to be one of the document. */
function runPropsOf(format: PutRunFormat, doc?: DocModel): RunProps {
  const base = patchRunProps(EMPTY_RUN_PROPS, format.f || {}, doc);
  // A change of formatting sets the slot of complex scripts with the plain one; what travels states it apart.
  const out: RunProps = { ...base };
  delete out.boldCs;
  delete out.italicCs;
  delete out.sizeCs;
  if (format.cs) {
    const cs = patchRunProps(EMPTY_RUN_PROPS, { bold: format.cs.bold, italic: format.cs.italic, size: format.cs.size });
    if (cs.bold !== undefined) out.boldCs = cs.bold;
    if (cs.italic !== undefined) out.italicCs = cs.italic;
    if (cs.size !== undefined) out.sizeCs = cs.size;
  }
  if (format.fonts) {
    const fonts: RunFonts = {};
    for (const slot of ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const) {
      const name = format.fonts[slot];
      if (name !== undefined) fonts[slot] = (patchRunProps(EMPTY_RUN_PROPS, { font: name }).fonts as RunFonts).ascii;
    }
    const hint = format.fonts.hint;
    if (hint !== undefined) {
      if (hint !== 'default' && hint !== 'eastAsia' && hint !== 'cs') throw new RangeError('The hint of the fonts is not valid');
      fonts.hint = hint;
    }
    if (Object.keys(fonts).length > 0) out.fonts = fonts;
    else delete out.fonts;
  }
  if (format.lang) {
    const lang: Language = {};
    for (const key of ['val', 'eastAsia', 'bidi'] as const) {
      const tag = format.lang[key];
      if (tag === undefined) continue;
      if (typeof tag !== 'string' || !LANGUAGE.test(tag)) throw new RangeError('The language is not valid');
      lang[key] = tag;
    }
    if (Object.keys(lang).length > 0) out.lang = lang;
  }
  return Object.keys(out).length > 0 ? out : EMPTY_RUN_PROPS;
}

/** What travels of run properties; undefined when they hold what no change of formatting can set (see putOf: the round trip is checked). */
function runFormatOf(props: RunProps): PutRunFormat | undefined {
  const f: RunFormatPatch = {};
  const out: PutRunFormat = {};
  const all = props as Record<string, unknown>;
  for (const key of Object.keys(all)) {
    const value = all[key];
    if (value === undefined) continue;
    switch (key) {
      case 'bold':
      case 'italic':
      case 'strike':
      case 'size':
      case 'highlight':
      case 'vertAlign':
      case 'style':
        (f as Record<string, unknown>)[key] = value;
        break;
      case 'boldCs':
      case 'italicCs':
      case 'sizeCs':
        out.cs = out.cs || {};
        (out.cs as Record<string, unknown>)[key.slice(0, -2)] = value;
        break;
      case 'underline': {
        const line = value as { style?: string };
        if (Object.keys(line).some((k) => k !== 'style' && (line as Record<string, unknown>)[k] !== undefined) || typeof line.style !== 'string') return undefined;
        f.underline = line.style;
        break;
      }
      case 'color':
        f.color = { ...(value as DocColor) };
        break;
      case 'fonts': {
        const fonts = value as RunFonts;
        const slots = { ascii: fonts.ascii, hAnsi: fonts.hAnsi, eastAsia: fonts.eastAsia, cs: fonts.cs, hint: fonts.hint };
        for (const k of Object.keys(fonts)) if ((fonts as Record<string, unknown>)[k] !== undefined && !(k in slots)) return undefined;
        out.fonts = {};
        for (const k of Object.keys(slots) as (keyof typeof slots)[]) if (slots[k] !== undefined) out.fonts[k] = slots[k];
        break;
      }
      case 'lang':
        out.lang = { ...(value as Language) };
        break;
      default:
        return undefined;
    }
  }
  if (Object.keys(f).length > 0) out.f = f;
  return out;
}

/** Paragraph properties of what travelled, read by patchParaProps, with the mark of the paragraph read as a run. */
function paraPropsOf(put: PutParagraph, doc?: DocModel): ParaProps {
  const props = patchParaProps(EMPTY_PARA_PROPS, (put.p || {}) as ParaFormatPatch, doc);
  if (!put.m) return props;
  const mark = runPropsOf(put.m, doc);
  if (mark === EMPTY_RUN_PROPS) return props;
  return { ...props, markProps: mark };
}

function paraFormatOf(props: ParaProps): { p?: ParaFormatPatch; m?: PutRunFormat } | undefined {
  const p: Record<string, unknown> = {};
  const out: { p?: ParaFormatPatch; m?: PutRunFormat } = {};
  const all = props as Record<string, unknown>;
  for (const key of Object.keys(all)) {
    const value = all[key];
    if (value === undefined) continue;
    if (key === 'markProps') {
      const mark = runFormatOf(value as RunProps);
      if (!mark) return undefined;
      out.m = mark;
    } else if (key === 'tabs') {
      // The stops the paragraph states, as a list: only when everybody reads them as they are.
      const tabs = (value as TabStop[]).map((tab) => ({ ...tab }));
      try {
        patchParaProps(EMPTY_PARA_PROPS, { tabs });
      } catch {
        return undefined;
      }
      p.tabs = tabs;
    } else if (PARA_KEYS.indexOf(key) >= 0) p[key] = typeof value === 'object' && value !== null ? { ...(value as object) } : value;
    else return undefined;
  }
  if (Object.keys(p).length > 0) out.p = p as ParaFormatPatch;
  return out;
}

/** The properties of a paragraph a change of paragraph formatting sets (ParaFormatPatch). */
const PARA_KEYS = ['align', 'indent', 'spacing', 'style', 'numbering', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'contextualSpacing', 'outlineLevel'];

/** A run format read from what arrived: plain data of the known members; the values are read when the paragraph is made (paragraphOf). */
function readRunFormat(value: unknown): PutRunFormat | undefined {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const from = value as Record<string, unknown>;
  for (const key of Object.keys(from)) if (from[key] !== undefined && RUN_FORMAT_KEYS.indexOf(key) < 0) return undefined;
  let copy: PutRunFormat;
  try {
    copy = plainCopy(from, 'paragraph') as PutRunFormat;
  } catch {
    return undefined;
  }
  try {
    runPropsOf(copy);
  } catch (error) {
    if (error instanceof RangeError) return undefined;
    throw error;
  }
  return copy;
}

/** The attributes of a paragraph or a run that travel: the marks Word leaves of the session that wrote them, and the identity of a paragraph. */
const KEPT_ATTRS: { [name: string]: true } = { 'w:rsidR': true, 'w:rsidRPr': true, 'w:rsidP': true, 'w:rsidRDefault': true, 'w:rsidDel': true, 'w:rsidTr': true, 'w14:paraId': true, 'w14:textId': true };

/** What a run holds beside text, by the number it travels as. */
const ITEMS: readonly RunItem[] = [{ type: 'tab' }, { type: 'break', kind: 'line' }, { type: 'break', kind: 'page' }, { type: 'break', kind: 'column' }, { type: 'noBreakHyphen' }, { type: 'softHyphen' }, { type: 'lastRenderedPageBreak' }];

/**
 * A value of an attribute that travels: an rsid is eight hex digits (ST_LongHexNumber); the id of
 * a paragraph or of its text is eight hex digits below 80000000 (as the schema of Word 2010 says).
 */
function validAttr(name: string, value: string): boolean {
  if (!/^[0-9A-Fa-f]{8}$/.test(value)) return false;
  if (name === 'w14:paraId' || name === 'w14:textId') return parseInt(value, 16) < 0x80000000;
  return true;
}

function attrsOf(value: unknown): [string, string][] | undefined {
  if (!Array.isArray(value) || value.length > 16) return undefined;
  const out: [string, string][] = [];
  for (const pair of value) {
    if (!Array.isArray(pair) || pair.length !== 2 || typeof pair[0] !== 'string' || KEPT_ATTRS[pair[0]] !== true || typeof pair[1] !== 'string' || !/^[0-9A-Fa-f]{1,8}$/.test(pair[1])) return undefined;
    if (out.some((seen) => seen[0] === pair[0])) return undefined;
    out.push([pair[0], pair[1]]);
  }
  return out;
}

/** The number of a comment as a mark names it: digits as a file writes them, 0 to 2^31-1. */
const commentNumber = (value: unknown): value is string => typeof value === 'string' && /^(0|[1-9][0-9]{0,9})$/.test(value) && Number(value) <= 2147483647;

function inlineOf(value: unknown): PutInline | undefined {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const from = value as Record<string, unknown>;
  const keys = Object.keys(from).filter((key) => from[key] !== undefined);
  if (keys.length === 1 && keys[0] === 'e') return typeof from.e === 'string' && PROOF_KINDS.indexOf(from.e) >= 0 ? { e: from.e } : undefined;
  if (keys.length === 1 && keys[0] === 'b') {
    // The start or the end of a bookmark: its number, and for a start its name.
    // (A start has a name - CT_Bookmark says so; an id is a number written as numbers are, 0 to 2^31-1.)
    const mark = from.b;
    if (!Array.isArray(mark) || (mark.length !== 2 && mark.length !== 3) || (mark[0] !== 0 && mark[0] !== 1) || typeof mark[1] !== 'string' || !/^(0|[1-9][0-9]{0,9})$/.test(mark[1]) || Number(mark[1]) > 2147483647) return undefined;
    if (mark.length === 2) return mark[0] === 1 ? { b: [1, mark[1]] } : undefined;
    return typeof mark[2] === 'string' && /^[A-Za-z0-9_.-]{1,40}$/.test(mark[2]) ? { b: [mark[0], mark[1], mark[2]] } : undefined;
  }
  if (keys.length === 1 && keys[0] === 'k') {
    // The start or the end of the range of a comment: its number, as a file writes it.
    const mark = from.k;
    return Array.isArray(mark) && mark.length === 2 && (mark[0] === 0 || mark[0] === 1) && commentNumber(mark[1]) ? { k: [mark[0], mark[1]] } : undefined;
  }
  for (const key of keys) if (key !== 'p' && key !== 'a' && key !== 'c') return undefined;
  if (!Array.isArray(from.c) || from.c.length > 4096) return undefined;
  const c: PutItem[] = [];
  const text = (item: unknown): item is string => typeof item === 'string' && item.length <= MAX_TEXT && xmlText(item) && !/[\t\n\r]/.test(item);
  for (const item of from.c) {
    if (isInt(item, 0, ITEMS.length - 1)) c.push(item);
    else if (text(item)) c.push(item);
    else if (Array.isArray(item) && item.length === 2 && text(item[0]) && text(item[1])) c.push([item[0], item[1]]);
    else if (item !== null && typeof item === 'object' && !Array.isArray(item) && Object.keys(item).length === 1 && isInt((item as { r?: unknown }).r, 0, 2147483647)) c.push({ r: (item as { r: number }).r });
    else return undefined;
  }
  const out: PutInline = { c };
  if (from.p !== undefined) {
    const format = readRunFormat(from.p);
    if (format === undefined) return undefined;
    out.p = format as Record<string, unknown>;
  }
  if (from.a !== undefined) {
    const attrs = attrsOf(from.a);
    if (!attrs) return undefined;
    out.a = attrs;
  }
  return out;
}

function readPutParagraph(value: unknown): PutParagraph {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return fail('paragraph');
  const from = value as Record<string, unknown>;
  for (const key of Object.keys(from)) if (from[key] !== undefined && key !== 'p' && key !== 'm' && key !== 'a' && key !== 'r') return fail('paragraph');
  if (!Array.isArray(from.r) || from.r.length > 20000) return fail('paragraph');
  const out: PutParagraph = { r: [] };
  for (const item of from.r) {
    const inline = inlineOf(item);
    if (!inline) return fail('paragraph');
    out.r.push(inline);
  }
  if (from.p !== undefined) {
    let copy: unknown;
    try {
      copy = plainCopy(from.p, 'paragraph');
      // A paragraph that is put states its own stops as a list (see paraFormatOf()); a change of stops is no paragraph.
      const tabs = (copy as { tabs?: unknown }).tabs;
      if (tabs !== undefined && !Array.isArray(tabs)) return fail('paragraph');
      patchParaProps(EMPTY_PARA_PROPS, copy as ParaFormatPatch);
    } catch (error) {
      if (error instanceof RangeError || error instanceof WireError) return fail('paragraph');
      throw error;
    }
    out.p = copy as Record<string, unknown>;
  }
  if (from.m !== undefined) {
    const mark = readRunFormat(from.m);
    if (mark === undefined) return fail('paragraph');
    out.m = mark as Record<string, unknown>;
  }

  if (from.a !== undefined) {
    const attrs = attrsOf(from.a);
    if (!attrs) return fail('paragraph');
    out.a = attrs;
  }
  return out;
}

const hasKeys = (value: object): boolean => Object.keys(value).length > 0;

/** A paragraph as it travels; undefined for one that holds something that does not travel. */
export function putOf(paragraph: Paragraph): PutParagraph | undefined {
  for (const key of Object.keys(paragraph)) {
    if (key !== 'type' && key !== 'props' && key !== 'children' && key !== 'attrs' && key !== 'srcStart' && key !== 'srcEnd' && (paragraph as unknown as Record<string, unknown>)[key] !== undefined) return undefined;
  }
  const out: PutParagraph = { r: [] };
  if (hasKeys(paragraph.props)) {
    const format = paraFormatOf(paragraph.props);
    if (!format) return undefined;
    if (format.p) out.p = format.p as Record<string, unknown>;
    if (format.m) out.m = format.m as Record<string, unknown>;
  }
  if (paragraph.attrs) {
    const attrs = attrsOf(paragraph.attrs);
    if (!attrs) return undefined;
    out.a = attrs;
  }
  for (const inline of paragraph.children) {
    const all = inline as unknown as Record<string, unknown>;
    if (inline.type === 'proofError') {
      for (const key of Object.keys(all)) if (key !== 'type' && key !== 'kind' && all[key] !== undefined) return undefined;
      if (PROOF_KINDS.indexOf(inline.kind) < 0) return undefined;
      out.r.push({ e: inline.kind });
      continue;
    }
    if (inline.type === 'bookmarkStart' || inline.type === 'bookmarkEnd') {
      for (const key of Object.keys(all)) if (key !== 'type' && key !== 'id' && key !== 'name' && all[key] !== undefined) return undefined;
      out.r.push(inline.name !== undefined ? { b: [inline.type === 'bookmarkStart' ? 0 : 1, inline.id, inline.name] } : { b: [inline.type === 'bookmarkStart' ? 0 : 1, inline.id] });
      continue;
    }
    if (inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') {
      for (const key of Object.keys(all)) if (key !== 'type' && key !== 'id' && all[key] !== undefined) return undefined;
      if (!commentNumber(inline.id)) return undefined;
      out.r.push({ k: [inline.type === 'commentRangeStart' ? 0 : 1, inline.id] });
      continue;
    }
    if (inline.type !== 'run') return undefined;
    for (const key of Object.keys(all)) if (key !== 'type' && key !== 'props' && key !== 'content' && key !== 'attrs' && all[key] !== undefined) return undefined;
    const c: PutItem[] = [];
    for (const item of inline.content) {
      const fields = item as unknown as Record<string, unknown>;
      if (item.type === 'commentReference') {
        for (const key of Object.keys(fields)) if (key !== 'type' && key !== 'id' && fields[key] !== undefined) return undefined;
        c.push({ r: item.id });
      } else if (item.type === 'text') {
        for (const key of Object.keys(fields)) if (key !== 'type' && key !== 'text' && key !== 'source' && fields[key] !== undefined) return undefined;
        c.push(item.source !== undefined ? [item.text, item.source] : item.text);
      } else {
        const code = ITEMS.findIndex((known) => known.type === item.type && (known as { kind?: string }).kind === (item as { kind?: string }).kind);
        if (code < 0 || Object.keys(fields).filter((key) => fields[key] !== undefined).length !== Object.keys(ITEMS[code]).length) return undefined;
        c.push(code);
      }
    }
    const run: PutInline = { c };
    if (hasKeys(inline.props)) {
      const format = runFormatOf(inline.props);
      if (!format) return undefined;
      run.p = format as Record<string, unknown>;
    }
    if (inline.attrs) {
      const attrs = attrsOf(inline.attrs);
      if (!attrs) return undefined;
      run.a = attrs;
    }
    out.r.push(run);
  }
  // What was read is what will be read by the others: a paragraph that is not what it travels as does not travel.
  try {
    const read = readPutParagraph(JSON.parse(JSON.stringify(out)));
    return samePlain(paragraphOf(read), paragraph) ? read : undefined;
  } catch {
    return undefined;
  }
}

/**
 * The paragraph a travelling paragraph stands for. Its properties are made by the rules of a
 * change of formatting; with `doc`, the styles and lists they name have to be the document's.
 * Throws RangeError for what those rules do not take.
 */
export function paragraphOf(put: PutParagraph, doc?: DocModel): Paragraph {
  const children: Inline[] = [];
  for (const inline of put.r) {
    if ('e' in inline) {
      children.push({ type: 'proofError', kind: inline.e });
      continue;
    }
    if ('b' in inline) {
      const mark = inline.b;
      children.push(mark.length === 3 ? { type: mark[0] === 0 ? 'bookmarkStart' : 'bookmarkEnd', id: mark[1], name: mark[2] } : { type: mark[0] === 0 ? 'bookmarkStart' : 'bookmarkEnd', id: mark[1] });
      continue;
    }
    if ('k' in inline) {
      children.push({ type: inline.k[0] === 0 ? 'commentRangeStart' : 'commentRangeEnd', id: inline.k[1] });
      continue;
    }
    const content: RunItem[] = inline.c.map((item): RunItem =>
      typeof item === 'number' ? ({ ...ITEMS[item] } as RunItem) : typeof item === 'string' ? { type: 'text', text: item } : Array.isArray(item) ? { type: 'text', text: item[0], source: item[1] } : { type: 'commentReference', id: item.r },
    );
    const run: Run = { type: 'run', props: inline.p ? runPropsOf(inline.p as PutRunFormat, doc) : EMPTY_RUN_PROPS, content };
    if (inline.a) run.attrs = inline.a.map((pair): [string, string] => [pair[0], pair[1]]);
    children.push(run);
  }
  const paragraph: Paragraph = { type: 'paragraph', props: put.p || put.m ? paraPropsOf(put, doc) : EMPTY_PARA_PROPS, children };
  if (put.a) paragraph.attrs = put.a.map((pair): [string, string] => [pair[0], pair[1]]);
  return paragraph;
}

/** Whether two values hold the same, wherever in a file they were read from. */
function samePlain(a: unknown, b: unknown, depth = 0): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null || depth > 32) return false;
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!samePlain(a[i], b[i], depth + 1)) return false;
    return true;
  }
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = (of: Record<string, unknown>): string[] => Object.keys(of).filter((key) => of[key] !== undefined && key !== 'srcStart' && key !== 'srcEnd');
  const mine = keys(left);
  if (mine.length !== keys(right).length) return false;
  for (const key of mine) if (!samePlain(left[key], right[key], depth + 1)) return false;
  return true;
}

/** A paragraph that travels, or a table of nothing but such paragraphs: what a `put` may take out of a document. */
function isPlain(block: Block, depth = 0): boolean {
  if (block.type === 'paragraph') return putOf(block) !== undefined;
  if (block.type !== 'table' || depth > 8 || block.wrap !== undefined) return false;
  for (const row of block.rows) for (const cell of row.cells) for (const inner of cell.blocks) if (!isPlain(inner, depth + 1)) return false;
  return true;
}

/**
 * What a `put` takes out, in eight hex digits: the text and the formatting of the blocks (as
 * they travel; not the attributes, nor how a text was spelled in its file - two participants
 * may hold the same text read from a file and made by editing). A put that would take out
 * other blocks than the ones its author took out is refused.
 */
export function putDigest(blocks: readonly Block[]): string {
  const parts: unknown[] = [];
  const add = (block: Block, depth: number): void => {
    if (block.type === 'paragraph') {
      const put = putOf(block);
      if (!put) return void parts.push(['?', paragraphText(block)]);
      parts.push([put.p || null, put.m || null, put.r.map((inline) => ('c' in inline ? [inline.p || null, inline.c.map((item) => (Array.isArray(item) ? item[0] : item))] : inline))]);
    } else if (block.type === 'table' && depth < 8) {
      parts.push('table');
      for (const row of block.rows) for (const cell of row.cells) for (const inner of cell.blocks) add(inner, depth + 1);
      parts.push('end');
    } else parts.push(block.type);
  };
  for (const block of blocks) add(block, 0);
  return hashOf(stable(parts));
}

/** FNV-1a of a text, in eight hex digits. */
export function hashOf(text: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return ('0000000' + hash.toString(16)).slice(-8);
}

/** Most characters of the blocks a block mark says the text of. */
const MAX_CARET_TEXT = 1 << 20;

const plainOf = (block: Block): string => (block.type === 'paragraph' ? paragraphText(block) : '');

/**
 * One thing that names itself in a document: a paragraph (by its id) or a bookmark (its start,
 * with its name, or its end); or the number of a comment that a comment or a mark of one carries
 * (`cmt`: only told to those who ask for it, see namesOfDocument).
 */
type Named = { kind: 'para'; id: string } | { kind: 'start'; id: string; name: string | undefined } | { kind: 'end'; id: string } | { kind: 'cmt'; id: string };

/** The elements whose `id` is the number of a comment. */
const COMMENT_ELEMENTS: { [local: string]: true } = { comment: true, commentRangeStart: true, commentRangeEnd: true, commentReference: true };

/** Where a thing that names itself stands: outside what a put replaces, in what it replaces, or in what it brings (see namesOfDocument). */
export type NamedWhere = 'outside' | 'replaced' | 'put';

/**
 * Most XML tokens read in one part for what it names: above what the largest document that is
 * edited together writes there - MAX_STORY_CHARS characters (each an element at most), MAX_LIST_BLOCKS
 * blocks, MAX_STORY_CELLS cells, with what stands around them - so that no such document is ever
 * cut short. A part that takes more is one that cannot be read whole.
 */
export const NAME_TOKENS_PER_PART = 4 * (MAX_STORY_CHARS_LIMIT + MAX_LIST_BLOCKS + 4 * MAX_STORY_CELLS) + 1000000;
/** How deep XML is read for what it names: well beyond what a file is read with (64 levels), so that nothing read is cut short. */
const NAME_DEPTH = 512;

/** A canonical id of a bookmark: the number it is, without leading zeros ("09" is 9); an id that is no number as it is. */
const bookmarkKey = (id: string): string => (/^[0-9]+$/.test(id) ? id.replace(/^0+(?=[0-9])/, '') : id);

/** A counter of what is read, shared by the parts of one reading of a document. */
interface Reading {
  tokens: number;
}

/**
 * What XML text names, in document order: the id of a paragraph (w14:paraId of any element), the
 * starts of bookmarks (with their names) and their ends. Of an mc:AlternateContent both branches are
 * read and what they name is counted once (the branches stand for each other). Throws RangeError
 * where the text cannot be read whole (it is not XML, too many tokens, too deep).
 */
function namesInXml(text: string, reading: Reading): Named[] {
  const out: Named[] = [];
  let sc: XmlScanner;
  try {
    sc = new XmlScanner(`<x>${text}</x>`);
  } catch {
    throw new RangeError('What the document is written as cannot be read for what it names');
  }
  let depth = 0;
  /** Depth of the outermost mc:AlternateContent read now, and what it named so far. */
  let alternate = -1;
  let seen = new Set<string>();
  const add = (thing: Named): void => {
    if (alternate >= 0) {
      const key = thing.kind === 'para' ? `p ${thing.id}` : thing.kind === 'end' ? `e ${bookmarkKey(thing.id)}` : thing.kind === 'cmt' ? `c ${thing.id}` : `s ${bookmarkKey(thing.id)} ${thing.name}`;
      if (seen.has(key)) return;
      seen.add(key);
    }
    out.push(thing);
  };
  try {
    for (;;) {
      if (!sc.next()) break;
      if (++reading.tokens > NAME_TOKENS_PER_PART) throw new RangeError('The document is too large to be read whole for what it names');
      if (sc.kind === 'close') {
        depth--;
        if (alternate >= 0 && depth <= alternate) {
          alternate = -1;
          seen = new Set<string>();
        }
        continue;
      }
      if (sc.kind !== 'open') continue;
      const name = sc.name.slice(sc.name.indexOf(':') + 1);
      if (name === 'AlternateContent' && alternate < 0) alternate = depth;
      let id: string | undefined;
      let label: string | undefined;
      for (const [key, value] of sc.attrs()) {
        const local = key.slice(key.indexOf(':') + 1);
        if (local === 'paraId') add({ kind: 'para', id: value.toUpperCase() });
        else if (local === 'id') id = value;
        else if (local === 'name') label = value;
      }
      if (name === 'bookmarkStart' || name === 'bookmarkEnd') {
        if (id === undefined) throw new RangeError('A bookmark of the document has no id');
        add(name === 'bookmarkStart' ? { kind: 'start', id, name: label } : { kind: 'end', id });
      } else if (id !== undefined && COMMENT_ELEMENTS[name] === true) add({ kind: 'cmt', id });
      if (sc.selfClosing) {
        if (alternate >= 0 && depth === alternate) {
          alternate = -1;
          seen = new Set<string>();
        }
        continue;
      }
      if (++depth > NAME_DEPTH) throw new RangeError('The document is too deep to be read whole for what it names');
    }
  } catch (error) {
    if (error instanceof RangeError) throw error;
    throw new RangeError('What the document is written as cannot be read for what it names');
  }
  return out;
}

/** What a part writes around its blocks, kept with what it depends on (see namesOfDocument). */
const shellNames = new WeakMap<object, { parts: unknown[]; names: Named[] }>();

/** What a part that is not written names: its source, kept while the source is the same. */
const partNames = new WeakMap<object, { source: string; names: Named[] }>();

/** What an mc:AlternateContent the model keeps as it was read names: both branches, once (the text is the whole element). */
function namesOfAlternate(text: string, reading: Reading): Named[] {
  return namesInXml(text, reading);
}

/** What a block names as it is written, kept while the block object lives (blocks are not changed, but made anew). */
const blockNames = new WeakMap<Block, Named[]>();
/** The same for wrappers around blocks (their own elements: properties, kept children). */
const wrapperNames = new WeakMap<object, Named[]>();

/** True when [start, end) of a source can be an element the writer copies (as write-package's spanOk). */
function spanOk(source: string, start: number | undefined, end: number | undefined): boolean {
  return start !== undefined && end !== undefined && start >= 0 && end > start + 1 && end <= source.length && source.charCodeAt(start) === 60 && source.charCodeAt(end - 1) === 62;
}

/**
 * Every paragraph id and every bookmark of a document, in document order, counted in what the
 * document is written as - "what the file will hold": each block as the docx writer writes it (or
 * the bytes it copies of an untouched block), what stands around the blocks of a part as the writer
 * writes it (the root, the last section, kept children, the background), headers, footers, notes,
 * comments. The tests read the saved file the same way (test/tools/package_check.py, named()).
 * `skip`: the blocks a put replaces; `put`: what it brings, said where the first of them stands.
 * `comment`: told the number of every comment and every mark of a comment the file will hold
 * (the comments themselves, their ranges, their references), as the file writes it. Throws
 * RangeError where the document cannot be read whole: what cannot be seen is not taken to be free.
 */
export function namesOfDocument(doc: DocModel, skip: readonly Block[], see: (thing: Named, where: NamedWhere) => void, put?: Named[], target?: string, comment?: (id: string) => void): void {
  const visit = (thing: Named, where: NamedWhere): void => {
    if (thing.kind !== 'cmt') see(thing, where);
    else if (comment) comment(thing.id);
  };
  const fail = (error: unknown): never => {
    if (error instanceof RangeError) throw error;
    throw new RangeError('The document cannot be written to be read for what it names');
  };
  const contextOf = (story: { part: string | null; rootAttrs: AttrList; relationships: Map<string, DocRelationship> }): BodyContext => {
    try {
      return createBodyContext(story.part || '/word/document.xml', story.rootAttrs, story.relationships);
    } catch (error) {
      return fail(error);
    }
  };
  const blocks = (ctx: BodyContext, list: readonly Block[], source: string, reading: Reading): void => {
    const wrappers = new Set<object>();
    for (const block of list) {
      if (put && skip.length > 0 && block === skip[0]) for (const thing of put) visit(thing, 'put');
      const where: NamedWhere = skip.indexOf(block) >= 0 ? 'replaced' : 'outside';
      // What an mc:AlternateContent around blocks holds is written as it was read, both branches (see the wrapper below).
      let alternate = false;
      for (let w = block.wrap, steps = 0; w && steps < 256; w = w.parent, steps++) if (w.kind === 'alternate') alternate = true;
      let found = alternate ? [] : blockNames.get(block);
      if (!found) {
        let text: string;
        if (spanOk(source, block.srcStart, block.srcEnd)) text = source.slice(block.srcStart as number, block.srcEnd as number);
        else {
          const out: string[] = [];
          try {
            writeBlocks(ctx, [block.wrap ? ({ ...block, wrap: undefined } as Block) : block], out, 0);
          } catch (error) {
            fail(error);
          }
          text = out.join('');
        }
        found = namesInXml(text, reading);
        blockNames.set(block, found);
      } else reading.tokens++;
      for (const thing of found) visit(thing, where);
      // The wrappers around blocks: their own elements, once each.
      for (let w = block.wrap, steps = 0; w && steps < 256; w = w.parent, steps++) {
        if (wrappers.has(w)) continue;
        wrappers.add(w);
        let own = wrapperNames.get(w);
        if (!own) {
          const out: string[] = [];
          const record = w as unknown as Record<string, unknown>;
          try {
            for (const key of ['pr', 'endPr', 'xml']) {
              const el = record[key] as XmlElement | undefined;
              if (el && el.type === 'element') out.push(treeText(ctx, el));
            }
            for (const extra of (record.extras as ExtraXml[] | undefined) || []) out.push(treeText(ctx, extra.el));
          } catch (error) {
            fail(error);
          }
          own = w.kind === 'alternate' ? namesOfAlternate(out.join(''), reading) : namesInXml(out.join(''), reading);
          wrapperNames.set(w, own);
        }
        for (const thing of own) visit(thing, where);
      }
    }
  };
  /** What a part writes around its blocks: the part written as the writer writes it, with no blocks. */
  const shell = (text: () => string, reading: Reading, owner?: object, parts?: unknown[]): void => {
    // What a part writes around its blocks depends on these (its root, its last section, what it keeps): kept while they are the same.
    const kept = owner ? shellNames.get(owner) : undefined;
    let names: Named[];
    if (kept && parts && kept.parts.length === parts.length && kept.parts.every((part, i) => part === parts[i])) names = kept.names;
    else {
      let written: string;
      try {
        written = text();
      } catch (error) {
        return fail(error);
      }
      names = namesInXml(written.replace(/^<\?xml[^>]*\?>/, ''), reading);
      if (owner && parts) shellNames.set(owner, { parts, names });
    }
    for (const thing of names) visit(thing, 'outside');
  };
  /** A part that is not written (not edited, and no put is meant for it) stays as it was read: what its source names. */
  const asRead = (story: { source: string; modified?: boolean }, isTarget: boolean): boolean => {
    if (isTarget || story.modified || story.source === '') return false;
    let found = partNames.get(story);
    if (!found || found.source !== story.source) {
      found = { source: story.source, names: namesInXml(story.source.replace(/^\s*<\?xml[^>]*\?>/, ''), { tokens: 0 }) };
      partNames.set(story, found);
    }
    for (const thing of found.names) visit(thing, 'outside');
    return true;
  };
  // The body.
  if (!asRead(doc.body, target === 'body')) {
    const reading: Reading = { tokens: 0 };
    const ctx = contextOf(doc.body);
    blocks(ctx, doc.body.blocks, doc.body.source, reading);
    const body = doc.body;
    shell(() => bodyText(ctx, { ...body, blocks: [] }, { regenerate: true }), reading, body, [body.rootName, body.rootAttrs, body.prolog, body.epilog, body.section, body.sectionIndex, body.extras, body.misc]);
  }
  // Headers and footers.
  for (const [stories, kind] of [
    [doc.headers, 'header'],
    [doc.footers, 'footer'],
  ] as const) {
    stories.forEach((story, id) => {
      if (asRead(story, target === `${kind}:${id}`)) return;
      const reading: Reading = { tokens: 0 };
      const ctx = contextOf(story);
      blocks(ctx, story.blocks, story.source, reading);
      shell(() => storyText(ctx, { ...story, blocks: [] }, { regenerate: true, root: { name: ctx.w + (kind === 'header' ? 'hdr' : 'ftr'), attrs: story.rootAttrs } }), reading, story, [story.rootName, story.rootAttrs, story.prolog, story.epilog]);
    });
  }
  // Notes and comments: each one's blocks, and the part around them (the items with no blocks).
  const collections: [StoryCollection<Note> | StoryCollection<Comment> | undefined, 'footnote' | 'endnote' | 'comment'][] = [
    [doc.footnotes, 'footnote'],
    [doc.endnotes, 'endnote'],
    [doc.comments, 'comment'],
  ];
  for (const [collection, local] of collections) {
    if (!collection) continue;
    if (asRead(collection, target !== undefined && target.indexOf(`${local}:`) === 0)) continue;
    const reading: Reading = { tokens: 0 };
    const ctx = contextOf(collection);
    // A comment is written with the paragraph id it is known by on its last paragraph (see commentsForWriting).
    const items = local === 'comment' ? commentsAsWritten(collection as StoryCollection<Comment>) : (collection.items as Map<number, Note | Comment>);
    items.forEach((item) => blocks(ctx, item.blocks, collection.source, reading));
    const empty = new Map<number, Note | Comment>();
    items.forEach((item, id) => empty.set(id, { ...item, blocks: [] }));
    shell(() => collectionText(ctx, { ...collection, items: empty } as StoryCollection<Note>, local, { regenerate: true, root: { name: ctx.w + local + 's', attrs: collection.rootAttrs } }), reading);
  }
}

/** Each comment as the writer writes it, kept while the comment object lives (comments are not changed, but made anew). */
const writtenComments = new WeakMap<Comment, Comment>();

function commentsAsWritten(collection: StoryCollection<Comment>): Map<number, Comment> {
  const out = new Map<number, Comment>();
  collection.items.forEach((comment, id) => {
    let written = writtenComments.get(comment);
    if (!written) {
      const one = commentsForWriting({ ...collection, items: new Map([[id, comment]]) });
      written = one.items.get(id) as Comment;
      writtenComments.set(comment, written);
    }
    out.set(id, written);
  });
  return out;
}

/** What a put brings that names itself, and the attributes it brings (name=value). */
function namedIn(x: readonly PutParagraph[]): { things: Named[]; attrs: [string, string][] } {
  const things: Named[] = [];
  const attrs: [string, string][] = [];
  for (const item of x) {
    for (const [name, value] of item.a || []) {
      attrs.push([name, value]);
      if (name === 'w14:paraId') things.push({ kind: 'para', id: value.toUpperCase() });
    }
    for (const inline of item.r) {
      if ('c' in inline) for (const [name, value] of inline.a || []) attrs.push([name, value]);
      if (!('b' in inline)) continue;
      if (inline.b[0] === 0) things.push({ kind: 'start', id: inline.b[1], name: inline.b.length === 3 ? inline.b[2] : undefined });
      else things.push({ kind: 'end', id: inline.b[1] });
    }
  }
  return { things, attrs };
}

/** A count by key. */
class Tally {
  private readonly counts = new Map<string, number>();
  add(key: string): void {
    this.counts.set(key, (this.counts.get(key) || 0) + 1);
  }
  of(key: string): number {
    return this.counts.get(key) || 0;
  }
}

/**
 * Throws RangeError where what a put brings would leave the document naming one thing twice, or
 * holding values no file holds - more than it did before the put. Each paragraph id, bookmark
 * start (by id), bookmark end (by id) and bookmark name (whatever its case) stands at most once in
 * the document after the put (or no more often than it stood before, where the file had it twice
 * already); a bookmark the put touches has exactly one start and one end after it (or the very
 * starts and ends it had before); an rsid is eight hex digits and an id of a paragraph or of its
 * text eight hex digits below 80000000, but for a value the blocks the put replaces carry already
 * (a file that holds it is not refused the undoing of its own text). The same function decides
 * for the author and for everybody else.
 */
function checkNames(doc: DocModel, wire: Extract<DocWire, { t: 'put' }>, list: number[], row: number): void {
  const replaced = containerBlocks(doc, wire.at.s, list).slice(row, row + wire.n);
  const names = new PutNames(wire.x);
  namesOfDocument(doc, replaced, (thing, where) => names.see(thing, where !== 'put', where !== 'replaced', where === 'replaced'), names.brought.things, wire.at.s);
  names.judge(replaced);
}

/**
 * checkNames() for a put in a text box: the paragraphs of the box it replaces are in a drawing of
 * a paragraph, which is one block of its story; the document is read as it is and as it would be
 * after the operations of the put (`ops`, of `story`), and what each holds is what stands before
 * and after it. The same rule.
 */
function checkNamesAcross(doc: DocModel, ops: readonly DocOp[], story: string, x: readonly PutParagraph[], replaced: readonly Block[]): void {
  const after = withOps(doc, ops, story);
  const names = new PutNames(x);
  // The bookmarks of the replaced paragraphs: these are plain paragraphs, which hold their bookmarks as marks of the model.
  const walk = (blocks: readonly Block[], depth: number): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        for (const inline of block.children) if (inline.type === 'bookmarkStart' || inline.type === 'bookmarkEnd') names.held.add(bookmarkKey(inline.id));
      } else if (block.type === 'table' && depth < 32) for (const r of block.rows) for (const cell of r.cells) walk(cell.blocks, depth + 1);
    }
  };
  walk(replaced, 0);
  namesOfDocument(doc, [], (thing) => names.see(thing, true, false, false), undefined, story);
  namesOfDocument(after, [], (thing) => names.see(thing, false, true, false), undefined, story);
  names.judge(replaced);
}

/** The document as it would be after operations of one story, the document itself unchanged: its story is a copy the operations change. */
function withOps(doc: DocModel, ops: readonly DocOp[], story: string): DocModel {
  const out: DocModel = { ...doc };
  const [kind, id] = story.indexOf(':') < 0 ? [story, ''] : [story.slice(0, story.indexOf(':')), story.slice(story.indexOf(':') + 1)];
  if (kind === 'body') out.body = { ...doc.body, blocks: doc.body.blocks.slice() };
  else if (kind === 'header' || kind === 'footer') {
    const stories = new Map(kind === 'header' ? doc.headers : doc.footers);
    const one = stories.get(id);
    if (one) stories.set(id, { ...one, blocks: one.blocks.slice() });
    if (kind === 'header') out.headers = stories;
    else out.footers = stories;
  } else if (kind === 'footnote' || kind === 'endnote') {
    const collection = kind === 'footnote' ? doc.footnotes : doc.endnotes;
    if (collection) {
      const copy = { ...collection, items: new Map(collection.items) };
      if (kind === 'footnote') out.footnotes = copy;
      else out.endnotes = copy;
    }
  } else throw new RangeError('The story cannot be read for what it names');
  for (const op of ops) {
    if (op.type !== 'replaceBlocks' && op.type !== 'setNote') throw new RangeError('The change cannot be read for what it names');
    applyDocOp(out, op);
  }
  return out;
}

/** What is counted for a put: each thing that names itself before and after it, the order of the bookmarks, the bookmarks it touches. */
class PutNames {
  readonly before = new Tally();
  readonly after = new Tally();
  readonly brought: { things: Named[]; attrs: [string, string][] };
  /** The bookmarks in document order, before the put and after it. */
  readonly orderBefore: string[] = [];
  readonly orderAfter: string[] = [];
  /** The bookmarks the replaced blocks hold. */
  readonly held = new Set<string>();

  constructor(x: readonly PutParagraph[]) {
    this.brought = namedIn(x);
  }

  static keys(thing: Named): string[] {
    return thing.kind === 'para' ? [`p:${thing.id}`] : thing.kind === 'end' ? [`e:${bookmarkKey(thing.id)}`] : thing.kind === 'cmt' ? [] : thing.name !== undefined ? [`s:${bookmarkKey(thing.id)}`, `n:${thing.name.toLowerCase()}`] : [`s:${bookmarkKey(thing.id)}`];
  }

  /** A thing that stands in the document before the put (`before`), after it (`after`), in what it replaces (`replaced`). */
  see(thing: Named, before: boolean, after: boolean, replaced: boolean): void {
    if (before) for (const key of PutNames.keys(thing)) this.before.add(key);
    if (after) for (const key of PutNames.keys(thing)) this.after.add(key);
    if (thing.kind !== 'para') {
      const key = `${thing.kind === 'start' ? 's' : 'e'}:${bookmarkKey(thing.id)}`;
      if (before) this.orderBefore.push(key);
      if (after) this.orderAfter.push(key);
    }
    if (replaced && thing.kind !== 'para') this.held.add(bookmarkKey(thing.id));
  }

  /** Throws RangeError where the put names a thing more often than it stood, brings an id no file holds, or leaves a bookmark that does not start once and end once. */
  judge(replaced: readonly Block[]): void {
    const carried = new Set<string>();
    // The attributes of runs the replaced blocks carry, too (the walker gives those of paragraphs).
    const runs = (blocks: readonly Block[], depth: number): void => {
      if (depth > 32) return;
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          for (const [name, value] of block.attrs || []) carried.add(`${name}=${value}`);
          for (const inline of block.children) if (inline.type === 'run') for (const [name, value] of inline.attrs || []) carried.add(`${name}=${value}`);
        } else if (block.type === 'table') for (const r of block.rows) for (const cell of r.cells) runs(cell.blocks, depth + 1);
      }
    };
    runs(replaced, 0);
    for (const [name, value] of this.brought.attrs) if (!carried.has(`${name}=${value}`) && !validAttr(name, value)) throw new RangeError('A paragraph or a run carries an id that is no value of a file');
    // What was there twice already may stay so; nothing comes to be there twice.
    const check = (key: string): void => {
      if (this.after.of(key) > Math.max(1, this.before.of(key))) throw new RangeError(key.charAt(0) === 'p' ? 'The document has a paragraph of this id' : key.charAt(0) === 'n' ? 'The document has a bookmark of this name' : 'The document has a bookmark of this number');
    };
    for (const thing of this.brought.things) for (const key of PutNames.keys(thing)) check(key);
    // The bookmarks the put touches - those it brings, and those the replaced blocks held - start once and end once
    // after it, the start first; or stand as they stood before it.
    const touched = new Set<string>(this.held);
    for (const thing of this.brought.things) if (thing.kind !== 'para') touched.add(bookmarkKey(thing.id));
    const shape = (order: string[], id: string): string => order.filter((key) => key === `s:${id}` || key === `e:${id}`).join(' ');
    for (const id of touched) {
      const now = shape(this.orderAfter, id);
      if (now !== `s:${id} e:${id}` && now !== shape(this.orderBefore, id)) throw new RangeError('A bookmark would not start once and end once after its start');
    }
  }
}

/** Most characters of a paragraph, and of a story, while a document is edited together. */
const MAX_PARAGRAPH_CHARS = 1000000;
const MAX_STORY_CHARS = MAX_STORY_CHARS_LIMIT;

function putLength(item: PutParagraph): number {
  let length = 0;
  for (const inline of item.r) if ('c' in inline) for (const piece of inline.c) length += typeof piece === 'number' ? 1 : typeof piece === 'string' ? piece.length : Array.isArray(piece) ? piece[0].length : 0;
  return length;
}

/** The characters of the text of blocks, counted until there are more than `most`. */
function storyChars(blocks: readonly Block[], most: number, depth = 0): number {
  let count = 0;
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const inline of block.children) {
        if (inline.type !== 'run') continue;
        for (const item of inline.content) count += item.type === 'text' ? item.text.length : 1;
      }
    } else if (block.type === 'table' && depth < 32) {
      for (const row of block.rows) for (const cell of row.cells) count += storyChars(cell.blocks, most - count, depth + 1);
    }
    if (count > most) return count;
  }
  return count;
}

/** The cells of the tables of a list of blocks, those of its text boxes too, taken from `budget` until there are none left. */
function countCells(blocks: readonly Block[], budget: { left: number }, depth = 0, boxes = 0): void {
  for (const block of blocks) {
    if (budget.left < 0) return;
    if (block.type === 'paragraph' && boxes < MAX_BOX_DEPTH) {
      eachGraphic(block, (item) => {
        for (const box of item.textBoxes || []) countCells(box, budget, depth, boxes + 1);
        return budget.left < 0;
      });
    }
    if (block.type !== 'table' || depth > 32) continue;
    for (const row of block.rows) {
      budget.left -= row.cells.length;
      if (budget.left < 0) return;
      for (const cell of row.cells) countCells(cell.blocks, budget, depth + 1, boxes);
    }
  }
}

/** Refuses a command that would make a list or the tables of a story larger than a document is taken to be: everybody finds the same, on the same document. */
function withinBudget(doc: DocModel, wire: Exclude<DocWire, { t: 'cmt' }>): void {
  const grows = wire.t === 'put' ? wire.x.length : wire.t === 'text' ? wire.x.split('\n').length - 1 : wire.t === 'split' ? 1 : wire.t === 'cmd' && wire.name === 'insertTable' ? 2 : wire.t === 'cmd' && wire.name === 'insertPageBreak' ? 1 : 0;
  if (grows > 0) {
    let length = 0;
    try {
      length = containerBlocks(doc, wire.at.s, wire.at.p.slice(0, -1)).length;
    } catch {
      return;
    }
    if (length + grows > MAX_LIST_BLOCKS) throw new RangeError('The document is too large to be edited together');
  }
  const adds = wire.t === 'text' ? wire.x.length : wire.t === 'put' ? wire.x.reduce((sum, item) => sum + putLength(item), 0) : 0;
  if (adds > 0) {
    // A paragraph of more than MAX_PARAGRAPH_CHARS characters, a story of more than MAX_STORY_CHARS.
    let story: Block[];
    try {
      story = storyOf(doc, wire.at.s).blocks;
    } catch {
      return;
    }
    if (wire.t === 'text') {
      let here = 0;
      try {
        const block = containerBlocks(doc, wire.at.s, wire.at.p.slice(0, -1))[wire.at.p[wire.at.p.length - 1]];
        if (block && block.type === 'paragraph') here = paragraphText(block).length;
      } catch {
        here = 0;
      }
      const longest = wire.x.split('\n').reduce((most, line) => Math.max(most, line.length), 0);
      if (here + longest > MAX_PARAGRAPH_CHARS) throw new RangeError('A paragraph would be too long to be edited together');
    } else if (wire.t === 'put') for (const item of wire.x) if (putLength(item) > MAX_PARAGRAPH_CHARS) throw new RangeError('A paragraph would be too long to be edited together');
    // The text of its text boxes is text of the story as well.
    const most = MAX_STORY_CHARS - adds;
    const chars = storyChars(story, most);
    if (chars > most || chars + boxChars(story, most - chars) > most) throw new RangeError('The document would be too long to be edited together');
  }
  if (wire.t === 'cmd' && (wire.name === 'insertTable' || wire.name === 'insertRow' || wire.name === 'insertColumn')) {
    const budget = { left: MAX_STORY_CELLS - (wire.name === 'insertTable' ? (wire.a[0] as number) * (wire.a[1] as number) : 0) };
    try {
      countCells(storyOf(doc, wire.at.s).blocks, budget);
    } catch {
      return;
    }
    // A row or a column more in a table that is at the limit: refused as well (how many cells it brings is not asked).
    if (budget.left < (wire.name === 'insertTable' ? 0 : MAX_NEW_COLUMNS)) throw new RangeError('The tables of the document are too large to be edited together');
  }
}

/** The operations that put paragraphs in the place of plain paragraphs. Throws RangeError where the document does not have such paragraphs there, or does not know what the new ones name. */
function putEdit(doc: DocModel, wire: Extract<DocWire, { t: 'put' }>, names = true): DocEdit {
  const c = wire.at.p.slice(0, -1);
  const row = wire.at.p[wire.at.p.length - 1];
  const story = storyOf(doc, wire.at.s);
  const blocks = containerBlocks(doc, wire.at.s, c);
  if (row + wire.n > blocks.length) throw new RangeError('The paragraphs to replace are not in the document');
  // Only plain paragraphs, and tables of them, are taken out: whatever else stands there (a paragraph with a picture or a note) is not this command's to remove.
  for (let i = row; i < row + wire.n; i++) if (!isPlain(blocks[i])) throw new RangeError('What stands there cannot be replaced');
  const declared = (prefix: string): boolean => story.host.rootAttrs.some((pair) => pair[0] === `xmlns:${prefix}`);
  const insert: Block[] = [];
  // What is replaced is what the author replaced: should it be anything else (a step read wrongly through what others did),
  // nothing is written over it.
  if (putDigest(blocks.slice(row, row + wire.n)) !== wire.h) throw new RangeError('The paragraphs to replace are not the ones the step was made for');
  // What names one thing names it once in the document (in a text box: in the whole document, see runInBox).
  if (names) checkNames(doc, wire, c, row);
  // The marks of comments a put brings are the very marks of the blocks it replaces, in the same order: no comment gets a mark it
  // did not have, none loses one, none comes to stand before another it stood behind (a range that ends before it begins, two
  // ranges that cross, two comments that swap their numbers). Which paragraph a mark stands in is not asked: undoing an Enter
  // or a join brings the same marks back in other paragraphs.
  const marks = (list: readonly Block[]): string => {
    const out: string[] = [];
    const walk = (blocks: readonly Block[], depth: number): void => {
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          for (const inline of block.children) {
            if (inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') out.push(`${inline.type} ${inline.id}`);
            else if (inline.type === 'run') for (const item of inline.content) if (item.type === 'commentReference') out.push(`commentReference ${item.id}`);
          }
        } else if (block.type === 'table' && depth < 32) for (const r of block.rows) for (const cell of r.cells) walk(cell.blocks, depth + 1);
      }
    };
    walk(list, 0);
    return out.join('|');
  };
  const bringsMarks = (wire.x as PutParagraph[]).map((item) => paragraphOf(item));
  if (marks(bringsMarks) !== marks(blocks.slice(row, row + wire.n))) throw new RangeError('The paragraphs to put back hold other marks of comments than the ones they replace');
  for (const item of wire.x) {
    const paragraph = paragraphOf(item, doc);
    const attrs: AttrList[] = [paragraph.attrs || []];
    for (const inline of paragraph.children) if (inline.type === 'run' && inline.attrs) attrs.push(inline.attrs);
    // An attribute of a namespace the part does not declare would make the part one no reader takes.
    for (const list of attrs) for (const pair of list) if (pair[0].indexOf('w:') !== 0 && !declared(pair[0].slice(0, pair[0].indexOf(':')))) throw new RangeError('The paragraph carries an attribute this document has no name for');
    insert.push(paragraph);
  }
  const at: DocPosition = { story: wire.at.s, path: wire.at.p.slice(), offset: 0 };
  const op: DocOp = { type: 'replaceBlocks', story: wire.at.s, path: c, at: row, remove: wire.n, insert };
  return { ops: [op], selection: { anchor: at, focus: at } };
}

/** The commands of a step, read from what arrived. Throws WireError for what is no step of a document. */
export function decodeStep(raw: unknown): DocWire[] {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_STEP_WIRES) return fail('step');
  const wires = raw.map(readWire);
  // What the puts of a step take out and put back, together: a step of more is no step.
  // (And few puts: each one reads the whole document for what it names. A composition is undone by one: see joinedPuts.)
  let paragraphs = 0;
  let puts = 0;
  // And few commands of comments: each reads the whole document as well.
  if (wires.filter((wire) => wire.t === 'cmt').length > MAX_STEP_COMMENTS) return fail('step');
  for (const wire of wires) {
    if (wire.t !== 'put') continue;
    puts++;
    paragraphs += wire.n + wire.x.length;
  }
  if (paragraphs > MAX_STEP_PUT_PARAGRAPHS || puts > MAX_STEP_PUTS) return fail('step');
  return wires;
}

/** Most paragraphs the puts of a step take out and put back together, and most puts in a step. */
export const MAX_STEP_PUT_PARAGRAPHS = 4000;
export const MAX_STEP_PUTS = 16;

/** The commands of a step as they go to the relay. */
export function encodeStep(wires: readonly DocWire[]): unknown[] {
  return wires.map((wire) => {
    if (wire.t === 'cmt') return writeComment(wire);
    const out: Record<string, unknown> = { ...wire, at: writePos(wire.at) };
    if ('to' in wire && wire.to !== undefined) out.to = writePos(wire.to);
    return out;
  });
}

export const toPosition = (pos: WirePos): DocPosition => (pos.b ? { story: pos.s, path: pos.p.slice(), offset: pos.o, before: true } : { story: pos.s, path: pos.p.slice(), offset: pos.o });
export const toWirePos = (pos: DocPosition): WirePos => (pos.before ? { s: pos.story, p: pos.path.slice(), o: pos.offset, b: true } : { s: pos.story, p: pos.path.slice(), o: pos.offset });

// ---------------------------------------------------------------- text boxes
//
// A command in a text box travels like any other: its places are places of the box
// ([..., paragraph, -1, drawing, box, ...path in the box], as positions of the model name them).
// Everybody runs it the way the editor of one runs a command there (editInTextBox): on a
// stand-in document whose body is the box, with the places lowered into it, and what it changed
// goes back into the box as one replacement of the paragraph that anchors it. So the box gets
// what the body would get, by the same rules, and where the box may not be changed (its other
// branch could not take it, its shape or a control locks its text) everybody refuses it alike:
// that is read off the document. What it did is read off the box the same way and said in the
// places of the document (the list of blocks of the box is the paths that lead into it); what
// moves the box itself - text that went from the paragraph that anchors it, a paragraph end that
// came into it - makes what is about the box give way (see boxShaken in rebase.ts).

/** The path of the list of blocks of the innermost text box a command is in (`[..., paragraph, -1, drawing, box]`); undefined outside text boxes. Throws RangeError for ends in different boxes. */
export function boxBase(wire: Exclude<DocWire, { t: 'cmt' }>): number[] | undefined {
  const root = textBoxRootLength(wire.at.p);
  const to = 'to' in wire && wire.to !== undefined ? wire.to : undefined;
  if (to && (to.s !== wire.at.s || textBoxRootLength(to.p) !== root || !wire.at.p.slice(0, root).every((value, i) => to.p[i] === value))) throw new RangeError('The positions are in different text boxes');
  return root === 0 ? undefined : wire.at.p.slice(0, root);
}

const moved = (pos: WirePos, s: string, p: number[]): WirePos => (pos.b ? { s, p, o: pos.o, b: true } : { s, p, o: pos.o });

/** A command of a text box as it runs on the stand-in whose body is the box. */
function lowerWire<W extends Exclude<DocWire, { t: 'cmt' }>>(wire: W, base: number[]): W {
  const out = { ...wire, at: moved(wire.at, 'body', wire.at.p.slice(base.length)) };
  if ('to' in wire && wire.to !== undefined) (out as { to?: WirePos }).to = moved(wire.to, 'body', wire.to.p.slice(base.length));
  return out;
}

/** A command made on the stand-in of a text box, in the places of the document; undefined for one that names no place. */
export function liftWire(wire: DocWire, story: string, base: number[]): DocWire | undefined {
  if (wire.t === 'cmt') return undefined;
  const out = { ...wire, at: moved(wire.at, story, base.concat(wire.at.p)) };
  if ('to' in wire && wire.to !== undefined) (out as { to?: WirePos }).to = moved(wire.to, story, base.concat(wire.to.p));
  return out;
}

/** The paragraph that anchors the outermost text box of a path, as a block of its list. */
function outerAnchor(base: number[]): { c: number[]; row: number } {
  const at = base.indexOf(-1);
  return { c: base.slice(0, at - 1), row: base[at - 1] };
}

/** What a command did in a text box (marks of the stand-in), in the places of the document. Something unknown in the box: the paragraph that anchors it was replaced. */
function liftMarks(marks: readonly Mark[], story: string, base: number[]): Mark[] {
  const anchor = outerAnchor(base);
  return marks.map((mark): Mark => (mark.k === 'opq' ? { k: 'blk', s: story, c: anchor.c, at: anchor.row, n: 1, m: 1 } : { ...mark, s: story, c: base.concat(mark.c) }));
}

const boxRange = (wire: Exclude<DocWire, { t: 'cmt' }>): DocRange => {
  const at = toPosition(wire.at);
  return { anchor: at, focus: 'to' in wire && wire.to !== undefined ? toPosition(wire.to) : at };
};

/** The operations of the stand-in a command of a text box came to, by the operations of the document it came to (for undo: inverseOf). */
const BOX_OPS = new WeakMap<DocOp[], DocOp[]>();

/** Runs a command of a text box (see the head of this section). */
function runInBox(doc: DocModel, wire: Exclude<DocWire, { t: 'cmt' }>, base: number[]): DocEdit {
  if (wire.t === 'cmd' && !commandInTextBox(wire.name)) throw new RangeError('This command is not run in a text box');
  const lowered = lowerWire(wire, base);
  let inner: DocEdit | undefined;
  let replaced: Block[] = [];
  const edit = editInTextBox(doc, boxRange(wire), (inside) => {
    if (lowered.t === 'put') {
      replaced = containerBlocks(inside, 'body', lowered.at.p.slice(0, -1)).slice(lowered.at.p[lowered.at.p.length - 1], lowered.at.p[lowered.at.p.length - 1] + lowered.n);
      inner = putEdit(inside, lowered, false);
    } else inner = runCommand(inside, lowered);
    return inner;
  });
  if (!inner) throw new RangeError('The command did not run in the text box');
  // What the paragraphs a put brings into a box name is counted in the whole document, as everywhere else.
  if (wire.t === 'put' && edit.ops.length > 0) checkNamesAcross(doc, edit.ops, wire.at.s, wire.x, replaced);
  BOX_OPS.set(edit.ops, inner.ops);
  return edit;
}

/** The characters of the text boxes of a list of blocks (in its paragraphs and tables, and in boxes in boxes), counted until there are more than `most`. */
function boxChars(blocks: readonly Block[], most: number, depth = 0, boxes = 0): number {
  let count = 0;
  for (const block of blocks) {
    if (count > most) return count;
    if (block.type === 'table' && depth < 32) {
      for (const row of block.rows) for (const cell of row.cells) count += boxChars(cell.blocks, most - count, depth + 1, boxes);
    } else if (block.type === 'paragraph' && boxes < MAX_BOX_DEPTH) {
      eachGraphic(block, (item) => {
        for (const box of item.textBoxes || []) {
          count += storyChars(box, most - count) + boxChars(box, most - count, depth, boxes + 1);
          if (count > most) return true;
        }
        return false;
      });
    }
  }
  return count;
}

/**
 * How many of the first drawings of a paragraph are, after a change, the very same drawings at the same numbers;
 * ALL_DRAWINGS where it holds the very drawings it held, and no other (none of its boxes moved).
 */
function keptGraphics(was: Paragraph, now: Paragraph): number {
  const before: unknown[] = [];
  eachGraphic(was, (item) => void before.push(item));
  let kept = 0;
  let count = 0;
  let same = true;
  eachGraphic(now, (item, index) => {
    count++;
    if (same && index < before.length && before[index] === item) kept++;
    else same = false;
    return false;
  });
  return kept === before.length && count === before.length ? ALL_DRAWINGS : kept;
}

/** What a command would do to the document: its operations. Throws RangeError for a command the document does not take. */
export function runWire(doc: DocModel, wire: DocWire): DocEdit {
  if (wire.t === 'cmt') return commentEdit(doc, wire);
  withinBudget(doc, wire);
  const base = boxBase(wire);
  if (base) return runInBox(doc, wire, base);
  if (wire.t === 'put') return putEdit(doc, wire);
  return runCommand(doc, wire);
}

/** A command that is no put, on the document (or the stand-in of a text box) its places are in. */
function runCommand(doc: DocModel, wire: Exclude<DocWire, { t: 'cmt' | 'put' }>): DocEdit {
  const at = toPosition(wire.at);
  const range: DocRange = { anchor: at, focus: 'to' in wire && wire.to !== undefined ? toPosition(wire.to) : at };
  switch (wire.t) {
    case 'text':
      return insertText(doc, wire.to ? range : at, wire.x, wire.f);
    case 'del':
      return deleteRange(doc, range);
    case 'split':
      return splitParagraph(doc, at);
    case 'run':
      return setRunFormat(doc, range, wire.f);
    case 'unrun':
      // What its author turned off was on in all of the range: where it no longer is, turning it "off" would turn it on.
      if (rangeFormat(doc, range)[wire.k] !== true) throw new RangeError('The formatting to turn off is no longer on in all of the range');
      return toggleRunFormat(doc, range, wire.k);
    case 'para':
      return setParaFormat(doc, range, wire.f);
    case 'cmd': {
      const command = COMMANDS[wire.name];
      if (!command) throw new RangeError('Unknown command');
      return command.run(doc, at, range, wire.a);
    }
  }
}

/** What is kept of the document before a command runs, to read off afterwards what the command did. */
export interface Before {
  s: string;
  /** The list of blocks both ends of the command are in (or, for ends in different lists, the nearest list that holds both). */
  c: number[];
  r1: number;
  o1: number;
  r2: number;
  o2: number;
  count: number;
  /** Both ends are paragraphs of the list, and so is everything between them: `first` and `last` are their paragraphs. */
  text: boolean;
  first?: Paragraph;
  last?: Paragraph;
  prev?: Block;
  next?: Block;
  /** The blocks r1..r2 as they were. */
  old: Block[];
  does: 'text' | 'touch' | 'table';
  /** A named command: where it leaves the text as it was, it changed how the text looks (a link over a selection). */
  named: boolean;
  /**
   * A command in a text box: the rest is about the stand-in whose body is the box (`base`: the path of the box in
   * story `s`, see boxBase); `marks`: what it did there, once marksAfter() read it.
   */
  box?: { s: string; base: number[]; marks?: Mark[] };
}

/** Looks at the document before a command is applied. Undefined for a command about a place the document does not have (applying it will say so). */
export function lookBefore(doc: DocModel, wire: DocWire): Before | undefined {
  // A command of comments says what it did itself (see commentMarks).
  if (wire.t === 'cmt') return undefined;
  let base: number[] | undefined;
  try {
    base = boxBase(wire);
  } catch {
    return undefined;
  }
  if (base) {
    // In a text box: the box, as the body of its stand-in.
    const inner = base;
    let found: Before | undefined;
    try {
      found = readInTextBox(doc, boxRange(wire), (inside) => lookBefore(inside, lowerWire(wire, inner)));
    } catch {
      return undefined;
    }
    return found ? { ...found, box: { s: wire.at.s, base: inner } } : undefined;
  }
  let a = wire.at;
  let b = 'to' in wire && wire.to !== undefined ? wire.to : wire.at;
  // The blocks a `put` replaces: the last of them is its other end.
  if (wire.t === 'put') b = { s: a.s, p: a.p.slice(0, -1).concat([a.p[a.p.length - 1] + wire.n - 1]), o: 0 };
  if (a.s !== b.s) return undefined;
  let does: Before['does'] = wire.t === 'text' || wire.t === 'del' || wire.t === 'split' ? 'text' : wire.t === 'cmd' ? COMMANDS[wire.name].does : wire.t === 'put' ? 'table' : 'touch';
  // The list both ends are in: paths name a block, then (for a table) a row, a cell and a block again.
  let depth = 0;
  const least = Math.min(a.p.length, b.p.length) - 1;
  while (depth + 3 <= least && a.p[depth] === b.p[depth] && a.p[depth + 1] === b.p[depth + 1] && a.p[depth + 2] === b.p[depth + 2]) depth += 3;
  if (wire.t === 'put') depth = a.p.length - 1;
  else if (does === 'table') {
    // The table the place is in: the block of the list one level up.
    if (a.p.length < 4) return undefined;
    depth = a.p.length - 4;
    b = a;
  }
  if (a.p[depth] > b.p[depth] || (a.p[depth] === b.p[depth] && a.p.length === depth + 1 && b.p.length === depth + 1 && a.o > b.o)) [a, b] = [b, a];
  const c = a.p.slice(0, depth);
  let blocks: Block[];
  try {
    blocks = containerBlocks(doc, a.s, c);
  } catch {
    return undefined;
  }
  const r1 = a.p[depth];
  const r2 = b.p[depth];
  if (r2 >= blocks.length) return undefined;
  const direct = a.p.length === depth + 1 && b.p.length === depth + 1;
  let text = direct && does === 'text';
  for (let i = r1; text && i <= r2; i++) if (blocks[i].type !== 'paragraph') text = false;
  if (does === 'text' && !text) does = 'table';
  const out: Before = { s: a.s, c, r1, o1: direct ? a.o : 0, r2, o2: direct ? b.o : 0, count: blocks.length, text, does, named: wire.t === 'cmd', old: blocks.slice(r1, r2 + 1) };
  if (text) {
    out.first = blocks[r1] as Paragraph;
    out.last = blocks[r2] as Paragraph;
  }
  if (r1 > 0) out.prev = blocks[r1 - 1];
  if (r2 + 1 < blocks.length) out.next = blocks[r2 + 1];
  return out;
}

/**
 * What a command did to the places around it, read off the document after it was applied.
 * Text that came in or went is said place by place only when the document shows exactly
 * that: the text before the place and the text after the range are what they were, and
 * everything between is paragraphs. Otherwise the blocks the command was about are said to
 * have been replaced; and when not even that can be seen, the story to have changed.
 */
export function marksAfter(doc: DocModel, before: Before | undefined, story: string): Mark[] {
  const unknown: Mark[] = [{ k: 'opq', s: story }];
  if (!before) return unknown;
  const box = before.box;
  if (box) {
    // Read off the box as it is now (the command moved no box), in the places of the document.
    const inner: Before = { ...before };
    delete inner.box;
    let marks: Mark[];
    try {
      const place: DocRange = { anchor: { story: box.s, path: box.base.concat([0]), offset: 0 }, focus: { story: box.s, path: box.base.concat([0]), offset: 0 } };
      marks = readInTextBox(doc, place, (inside) => marksAfter(inside, inner, 'body'));
    } catch {
      marks = [{ k: 'opq', s: 'body' }];
    }
    box.marks = marks;
    return liftMarks(marks, box.s, box.base);
  }
  const { s, c, r1, o1, r2, o2 } = before;
  let blocks: Block[];
  try {
    blocks = containerBlocks(doc, s, c);
  } catch {
    return unknown;
  }
  const rows = blocks.length - before.count + (r2 - r1);
  if (rows < 0) return unknown;
  // What stood beside the blocks stands beside what they became.
  if ((r1 > 0 ? blocks[r1 - 1] : undefined) !== before.prev || (r1 + rows + 1 < blocks.length ? blocks[r1 + rows + 1] : undefined) !== before.next) return unknown;
  const replaced: Mark[] = [{ k: 'blk', s, c, at: r1, n: r2 - r1 + 1, m: rows + 1 }];
  // One paragraph that became one: what it kept of its drawings (their boxes are where they were; see boxShaken).
  const one = before.old[0];
  if (r1 === r2 && rows === 0 && one && one.type === 'paragraph' && blocks[r1] && blocks[r1].type === 'paragraph') (replaced[0] as Extract<Mark, { k: 'blk' }>).gk = keptGraphics(one, blocks[r1] as Paragraph);
  // The text of the blocks before and after, for the carets in them (not for a large text: they then stay where they were in their block).
  const was = before.old.map(plainOf);
  const now = blocks.slice(r1, r1 + rows + 1).map(plainOf);
  if (was.concat(now).reduce((sum, text) => sum + text.length, 0) <= MAX_CARET_TEXT) {
    (replaced[0] as Extract<Mark, { k: 'blk' }>).was = was;
    (replaced[0] as Extract<Mark, { k: 'blk' }>).now = now;
  }
  if (before.does === 'touch') {
    if (rows !== r2 - r1) return replaced;
    // What each paragraph keeps of its drawings (formatting keeps them all; see boxShaken).
    const gks: number[] = [];
    for (let i = 0; i <= r2 - r1; i++) {
      const was = before.old[i];
      const now = blocks[r1 + i];
      gks.push(was && now && was.type === 'paragraph' && now.type === 'paragraph' ? keptGraphics(was, now) : 0);
    }
    return [{ k: 'tch', s, c, r1, r2, gks }];
  }
  if (before.does === 'table' || !before.text) return replaced;
  for (let i = r1; i <= r1 + rows; i++) if (blocks[i].type !== 'paragraph') return replaced;
  const head = paragraphText(before.first as Paragraph).slice(0, o1);
  const tail = paragraphText(before.last as Paragraph).slice(o2);
  const start = paragraphText(blocks[r1] as Paragraph);
  const end = paragraphText(blocks[r1 + rows] as Paragraph);
  // The drawings that are still the first ones of the paragraph where the command began: their boxes did not move (see
  // boxShaken). A drawing has no width: one that went (deleted, or written over) shows in no text.
  const kept = keptGraphics(before.first as Paragraph, blocks[r1] as Paragraph);
  // The very same text in the same paragraph: the command changed how it looks, and every place in it is where it was.
  if (before.named && rows === 0 && r1 === r2 && start === paragraphText(before.first as Paragraph)) return [{ k: 'tch', s, c, r1, r2, gks: [kept] }];
  const last = end.length - tail.length - (rows === 0 ? o1 : 0);
  if (last < 0 || start.slice(0, o1) !== head || end.slice(end.length - tail.length) !== tail || (rows > 0 && start.length < o1)) return replaced;
  const out: Mark[] = [];
  if (r1 !== r2 || o1 !== o2) out.push({ k: 'del', s, c, r1, o1, r2, o2, gk: kept });
  if (rows > 0 || last > 0) out.push({ k: 'ins', s, c, row: r1, off: o1, rows, last, gk: kept });
  // Nothing came and nothing went in the text: what the command changed is how the paragraph looks, or what has no width in it.
  return out.length > 0 ? out : [{ k: 'tch', s, c, r1, r2, gks: [kept] }];
}

/**
 * The commands that take a command back, for the one who made it: what goes into its list of
 * what can be undone. `ops`: the operations the command came to; `marks`: what it did, as
 * marksAfter() read it. Null for a command that cannot be taken back by commands: it changed
 * more than the blocks of one list (a list that was made, a link that got its address), or
 * what it replaced does not travel (see putOf).
 */
export function inverseOf(doc: DocModel, wire: DocWire, before: Before | undefined, marks: readonly Mark[], ops: readonly DocOp[]): DocWire[] | null {
  if (ops.length === 0) return [];
  if (!before || wire.t === 'cmt') return null;
  const box = before.box;
  if (box) {
    // What takes it back in the box, made on its stand-in as it is now, in the places of the document.
    const done = BOX_OPS.get(ops as DocOp[]);
    const did = box.marks;
    if (!done || !did) return null;
    const inner: Before = { ...before };
    delete inner.box;
    let back: DocWire[] | null;
    try {
      const place: DocRange = { anchor: { story: box.s, path: box.base.concat([0]), offset: 0 }, focus: { story: box.s, path: box.base.concat([0]), offset: 0 } };
      back = readInTextBox(doc, place, (inside) => inverseOf(inside, lowerWire(wire, box.base), inner, did, done));
    } catch {
      return null;
    }
    if (!back) return null;
    const out: DocWire[] = [];
    for (const one of back) {
      const lifted = liftWire(one, box.s, box.base);
      if (!lifted) return null;
      out.push(lifted);
    }
    return out;
  }
  for (const op of ops) if (op.type !== 'replaceBlocks' || op.story !== before.s) return null;
  const { s, c, r1, r2 } = before;
  const first = marks[0];
  // Text that came in and nothing else: it goes again, wherever it has moved to and whatever stands in it by then.
  if (marks.length === 1 && first.k === 'ins') {
    const at: WirePos = { s, p: c.concat([first.row]), o: first.off };
    const to: WirePos = first.rows === 0 ? { s, p: c.concat([first.row]), o: first.off + first.last } : { s, p: c.concat([first.row + first.rows]), o: first.last };
    return [{ t: 'del', at, to }];
  }
  for (const mark of marks) if (mark.k === 'opq') return null;
  let blocks: Block[];
  try {
    blocks = containerBlocks(doc, s, c);
  } catch {
    return null;
  }
  // The blocks that stand there now, for the blocks that stood there.
  const n = r2 - r1 + 1 + blocks.length - before.count;
  if (n < 1) return null;
  for (let i = r1; i < r1 + n; i++) if (!blocks[i] || !isPlain(blocks[i])) return null;
  const x: PutParagraph[] = [];
  for (const block of before.old) {
    const put = block.type === 'paragraph' ? putOf(block) : undefined;
    if (!put) return null;
    x.push(put);
  }
  if (x.length === 0 || x.length > MAX_PUT_PARAGRAPHS || n > MAX_PUT_PARAGRAPHS) return null;
  const put: DocWire = { t: 'put', at: { s, p: c.concat([r1]), o: 0 }, n, x, h: putDigest(blocks.slice(r1, r1 + n)) };
  // As the others will read it: what they would not take (two paragraphs that bring one bookmark twice, say) is no undoing.
  try {
    decodeStep(JSON.parse(JSON.stringify(encodeStep([put]))));
  } catch (error) {
    if (error instanceof WireError) return null;
    throw error;
  }
  return [put];
}
