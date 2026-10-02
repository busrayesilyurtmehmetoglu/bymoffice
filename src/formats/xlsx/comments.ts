// Reads the comments of a sheet into the read-only view the grid shows them
// from: the legacy comments part (xl/commentsN.xml), the boxes of its legacy
// drawing (whether a box is always shown, and where), and threaded comments
// with the people who wrote them.
//
// The parts stay in the package and are written back from there.

import type { OpcPackage, Relationship } from '../../opc';
import { MAX_COLS, MAX_ROWS, parseCellRef, type DrawingAnchor, type RichRun, type SheetComment } from '../../model/sheet/types';
import { localName, parseXml, type XmlElement } from '../../xml';
import { decodeXstring } from './shared-strings';
import { parseFont } from './styles';

/** Comments read from one sheet; a sheet with more shows the first ones. */
export const MAX_SHEET_COMMENTS = 100000;
/** Longest text kept of one comment (the longest text a cell holds). */
const MAX_COMMENT_TEXT = 32767;
const MAX_COMMENT_RUNS = 1024;
const MAX_REPLIES = 256;
/** Parts larger than this are not read; the sheet is shown without its comments. */
const MAX_PART_BYTES = 64 * 1024 * 1024;
const EMU_PER_PIXEL = 9525;

function child(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const node of parent.children) if (node.type === 'element' && localName(node.name) === local) return node;
  return undefined;
}

function children(parent: XmlElement | undefined, local: string): XmlElement[] {
  const out: XmlElement[] = [];
  if (parent) for (const node of parent.children) if (node.type === 'element' && localName(node.name) === local) out.push(node);
  return out;
}

function textOf(el: XmlElement | undefined): string {
  if (!el) return '';
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

const keyOf = (row: number, col: number): number => row * MAX_COLS + col;

/** The text of a <text> element: plain (<t>) or in runs (<r>); phonetic text is left out. */
function readText(el: XmlElement | undefined): { text: string; runs?: RichRun[] } {
  let text = '';
  const runs: RichRun[] = [];
  let formatted = false;
  if (el) {
    for (const node of el.children) {
      if (node.type !== 'element' || text.length >= MAX_COMMENT_TEXT) continue;
      const local = localName(node.name);
      if (local === 't') {
        const part = decodeXstring(textOf(node)).slice(0, MAX_COMMENT_TEXT - text.length);
        text += part;
        if (part !== '' && runs.length < MAX_COMMENT_RUNS) runs.push({ text: part });
      } else if (local === 'r') {
        const part = decodeXstring(textOf(child(node, 't'))).slice(0, MAX_COMMENT_TEXT - text.length);
        if (part === '') continue;
        text += part;
        const props = child(node, 'rPr');
        if (runs.length < MAX_COMMENT_RUNS) {
          if (props) {
            runs.push({ text: part, font: parseFont(props) });
            formatted = true;
          } else runs.push({ text: part });
        } else runs[runs.length - 1].text += part;
      }
    }
  }
  return formatted && runs.length > 1 ? { text, runs } : { text };
}

/** The comments of a legacy comments part, in the order of the part; a cell stated twice keeps its first comment. */
export function parseComments(root: XmlElement, warnings: string[]): SheetComment[] {
  const authors = children(child(root, 'authors'), 'author').map((author) => decodeXstring(textOf(author)));
  const out: SheetComment[] = [];
  const seen = new Set<number>();
  let skipped = 0;
  for (const list of children(root, 'commentList')) {
    for (const el of children(list, 'comment')) {
      if (out.length >= MAX_SHEET_COMMENTS) {
        warnings.push(`Comments: only the first ${out.length} are shown`);
        return out;
      }
      const ref = el.attrs.get('ref');
      const cell = ref === undefined ? undefined : parseCellRef(ref);
      if (!cell || seen.has(keyOf(cell.row, cell.col))) {
        skipped++;
        continue;
      }
      seen.add(keyOf(cell.row, cell.col));
      const comment: SheetComment = { row: cell.row, col: cell.col, ...readText(child(el, 'text')) };
      const authorId = el.attrs.get('authorId');
      const author = authorId !== undefined && /^\d{1,9}$/.test(authorId) ? authors[Number(authorId)] : undefined;
      if (author !== undefined && author !== '') comment.author = author.slice(0, 255);
      out.push(comment);
    }
  }
  if (skipped > 0) warnings.push(`Comments: ${skipped} comment(s) could not be read and are not shown`);
  return out;
}

/** Names of the people of threaded comments, by id. */
export function parsePersons(root: XmlElement): Map<string, string> {
  const out = new Map<string, string>();
  for (const person of children(root, 'person')) {
    const id = person.attrs.get('id');
    const name = person.attrs.get('displayName');
    if (id !== undefined && name !== undefined && out.size < 100000) out.set(id, name.slice(0, 255));
  }
  return out;
}

/**
 * Puts the threaded comments of a sheet over its legacy ones: a cell with a
 * conversation shows its first message and the replies, by the people who
 * wrote them, in place of the note older applications are given for it.
 */
export function applyThreadedComments(comments: SheetComment[], root: XmlElement, persons: Map<string, string>): SheetComment[] {
  const byCell = new Map<number, SheetComment>();
  for (const comment of comments) byCell.set(keyOf(comment.row, comment.col), comment);
  const byId = new Map<string, SheetComment>();
  const out = comments.slice();
  for (const el of children(root, 'threadedComment')) {
    const text = decodeXstring(textOf(child(el, 'text'))).slice(0, MAX_COMMENT_TEXT);
    const personId = el.attrs.get('personId');
    const author = personId !== undefined ? persons.get(personId) : undefined;
    const parentId = el.attrs.get('parentId');
    if (parentId !== undefined) {
      const parent = byId.get(parentId);
      if (parent && parent.replies && parent.replies.length < MAX_REPLIES) parent.replies.push(author !== undefined ? { author, text } : { text });
      continue;
    }
    const ref = el.attrs.get('ref');
    const cell = ref === undefined ? undefined : parseCellRef(ref);
    const id = el.attrs.get('id');
    if (!cell) continue;
    const key = keyOf(cell.row, cell.col);
    const legacy = byCell.get(key);
    // A second conversation on the same cell is not something applications write; the first one stays.
    if (legacy && legacy.threaded) continue;
    const comment: SheetComment = { row: cell.row, col: cell.col, text, threaded: true, replies: [] };
    if (author !== undefined) comment.author = author;
    if (legacy) {
      if (legacy.visible) comment.visible = true;
      if (legacy.box) comment.box = legacy.box;
      out[out.indexOf(legacy)] = comment;
    } else {
      if (out.length >= MAX_SHEET_COMMENTS) continue;
      out.push(comment);
    }
    byCell.set(key, comment);
    if (id !== undefined) byId.set(id, comment);
  }
  for (const comment of out) if (comment.replies && comment.replies.length === 0) delete comment.replies;
  return out;
}

/** What the legacy drawing says about the box of a comment. */
export interface NoteBox {
  row: number;
  col: number;
  visible: boolean;
  box?: DrawingAnchor;
}

function uint(text: string | undefined): number {
  if (text === undefined) return -1;
  const trimmed = text.trim();
  return /^\d{1,9}$/.test(trimmed) ? Number(trimmed) : -1;
}

/**
 * <x:Anchor>: left column, offset, top row, offset, right column, offset,
 * bottom row, offset; the offsets are pixels.
 */
function boxOf(text: string | undefined): DrawingAnchor | undefined {
  if (text === undefined) return undefined;
  const numbers = text.split(',').map((part) => uint(part));
  if (numbers.length !== 8 || numbers.some((n) => n < 0)) return undefined;
  if (numbers[0] >= MAX_COLS || numbers[4] >= MAX_COLS || numbers[2] >= MAX_ROWS || numbers[6] >= MAX_ROWS) return undefined;
  return {
    kind: 'twoCell',
    from: { col: numbers[0], colOff: numbers[1] * EMU_PER_PIXEL, row: numbers[2], rowOff: numbers[3] * EMU_PER_PIXEL },
    to: { col: numbers[4], colOff: numbers[5] * EMU_PER_PIXEL, row: numbers[6], rowOff: numbers[7] * EMU_PER_PIXEL },
  };
}

function noteOf(row: number, col: number, anchor: string | undefined, visible: boolean): NoteBox | undefined {
  if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return undefined;
  const note: NoteBox = { row, col, visible };
  const box = boxOf(anchor);
  if (box) note.box = box;
  return note;
}

/** A box is always shown when its client data says so; the "visibility" of the shape's style follows it in every file Excel writes. */
function isVisible(value: string | undefined): boolean {
  if (value === undefined) return false;
  const text = value.trim().toLowerCase();
  return text === '' || text === 'true' || text === 't' || text === '1';
}

/**
 * The comment boxes of a legacy drawing part. VML is older than the rule
 * that it has to be well-formed, so a part the XML reader refuses is read as
 * text, one shape at a time, the way the writer moves these boxes.
 */
export function parseVmlNotes(bytes: Uint8Array): NoteBox[] {
  const out: NoteBox[] = [];
  let root: XmlElement | undefined;
  try {
    root = parseXml(bytes).root;
  } catch (error) {
    root = undefined;
  }
  if (root) {
    for (const shape of children(root, 'shape')) {
      if (out.length >= MAX_SHEET_COMMENTS) break;
      const data = child(shape, 'ClientData');
      if (!data || data.attrs.get('ObjectType') !== 'Note') continue;
      const visible = child(data, 'Visible');
      const note = noteOf(uint(textOf(child(data, 'Row'))), uint(textOf(child(data, 'Column'))), child(data, 'Anchor') ? textOf(child(data, 'Anchor')) : undefined, visible ? isVisible(textOf(visible)) : false);
      if (note) out.push(note);
    }
    return out;
  }
  const text = new TextDecoder().decode(bytes);
  let i = 0;
  while (out.length < MAX_SHEET_COMMENTS) {
    const start = text.indexOf('<v:shape', i);
    if (start < 0) break;
    const after = text.charCodeAt(start + 8);
    // "<v:shapetype" and the like are other elements.
    if (!(after === 32 || after === 62 || after === 9 || after === 10 || after === 13)) {
      i = start + 8;
      continue;
    }
    const close = text.indexOf('</v:shape>', start);
    if (close < 0) break;
    i = close + 10;
    const block = text.slice(start, i);
    if (!/<x:ClientData[^>]*ObjectType\s*=\s*["']Note["']/.test(block)) continue;
    const row = /<x:Row>\s*(\d{1,7})\s*<\/x:Row>/.exec(block);
    const col = /<x:Column>\s*(\d{1,5})\s*<\/x:Column>/.exec(block);
    if (!row || !col) continue;
    const anchor = /<x:Anchor>([^<]*)<\/x:Anchor>/.exec(block);
    const visible = /<x:Visible\s*(?:\/>|>([^<]*)<\/x:Visible>)/.exec(block);
    const note = noteOf(Number(row[1]), Number(col[1]), anchor ? anchor[1] : undefined, visible ? isVisible(visible[1] || '') : false);
    if (note) out.push(note);
  }
  return out;
}

/** Sets what the legacy drawing says about their boxes on the comments of a sheet. */
export function applyNoteBoxes(comments: SheetComment[], notes: NoteBox[]): void {
  if (notes.length === 0) return;
  const byCell = new Map<number, NoteBox>();
  for (const note of notes) if (!byCell.has(keyOf(note.row, note.col))) byCell.set(keyOf(note.row, note.col), note);
  for (const comment of comments) {
    const note = byCell.get(keyOf(comment.row, comment.col));
    if (!note) continue;
    if (note.visible) comment.visible = true;
    if (note.box) comment.box = note.box;
  }
}

const kindOf = (rel: Relationship): string => rel.type.slice(rel.type.lastIndexOf('/') + 1);

async function partBytes(pkg: OpcPackage, rel: Relationship): Promise<Uint8Array | undefined> {
  if (rel.external || !rel.part || !pkg.hasPart(rel.part)) return undefined;
  const bytes = await pkg.readPart(rel.part);
  if (bytes.length > MAX_PART_BYTES) throw new Error('the part is too large');
  return bytes;
}

/**
 * The comments of a sheet from the parts its relationships name; undefined for
 * a sheet without any. Never throws for what the parts hold: what cannot be
 * read is left out with a note in `warnings`.
 */
export async function readSheetComments(pkg: OpcPackage, sheetRels: Relationship[], persons: Map<string, string>, warnings: string[]): Promise<SheetComment[] | undefined> {
  let comments: SheetComment[] = [];
  for (const rel of sheetRels) {
    if (kindOf(rel) !== 'comments') continue;
    try {
      const bytes = await partBytes(pkg, rel);
      if (bytes) comments = comments.concat(parseComments(parseXml(bytes).root, warnings));
    } catch (error) {
      warnings.push(`Comments could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  for (const rel of sheetRels) {
    if (kindOf(rel) !== 'vmlDrawing' || comments.length === 0) continue;
    try {
      const bytes = await partBytes(pkg, rel);
      if (bytes) applyNoteBoxes(comments, parseVmlNotes(bytes));
    } catch (error) {
      // Without the boxes the comments are still shown, at their cells.
    }
  }
  for (const rel of sheetRels) {
    if (kindOf(rel) !== 'threadedComment') continue;
    try {
      const bytes = await partBytes(pkg, rel);
      if (bytes) comments = applyThreadedComments(comments, parseXml(bytes).root, persons);
    } catch (error) {
      warnings.push(`Threaded comments could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  return comments.length > 0 ? comments : undefined;
}
