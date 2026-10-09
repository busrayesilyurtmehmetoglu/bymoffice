// Operations on a document model. Every edit is one or more of these, and
// applying one returns the operation that undoes it.
//
// The contract with the writers (Story.modified): blocks, rows, cells and
// notes are never changed in place. An operation puts new objects where the
// old ones were, all the way up to the story, and a new block of the story
// carries no source span, so the writer writes it from the model and copies
// its untouched neighbours from the source. The inverse puts the old objects
// back by reference, source spans and all, and restores `modified`, so that
// "edit, undo, save" writes the bytes that were read.

import type { AbstractNum, Block, Body, DocMedia, DocModel, DocRelationship, DocStyles, FontInfo, Note, NumInstance, SectionProps, SourceSpan, Story, StoryCollection, Table, TableCell, TableRow } from '../../model/doc/types';
import { storyOf, type StoryHandle } from './position';

/** `modified` flag to set after an operation: true when missing, null to remove the flag. */
type Flag = boolean | null;

export type DocOp =
  | {
      /**
       * Replaces `remove` blocks at index `at` of a list of blocks by the
       * blocks of `insert`. `path` names the list: [] is the story itself,
       * [table, row, cell] the content of a cell (and so on for nested tables).
       */
      type: 'replaceBlocks';
      story: string;
      path: number[];
      at: number;
      remove: number;
      insert: Block[];
      modified?: Flag;
      /** Body only: the value Body.sectionIndex gets (null: none), instead of the one that follows from the change. */
      sectionIndex?: number | null;
    }
  /** Puts a note object in the place of the one with its id: what undoes a change of the blocks of a note. */
  | { type: 'setNote'; story: string; note: Note; modified?: Flag }
  /**
   * Takes a note ("footnote:<id>", "endnote:<id>") out of its collection: what
   * follows when the last reference to it is deleted, so that no note is left
   * that nothing refers to.
   */
  | { type: 'removeNote'; story: string; modified?: Flag }
  /** Puts back the whole table of notes of a kind, in its order: what undoes removeNote. */
  | { type: 'setNotes'; kind: 'footnote' | 'endnote'; items: Map<number, Note>; modified?: Flag }
  /** Adds, replaces or (null) removes numbering definitions and instances. */
  | { type: 'setNumbering'; abstracts?: [id: number, value: AbstractNum | null][]; nums?: [id: number, value: NumInstance | null][]; modified?: Flag }
  /** Adds, replaces or (null) removes one relationship of the part a story is in. */
  | { type: 'setRelationship'; story: string; id: string; relationship: DocRelationship | null }
  /** Puts back the whole relationship table of the part a story is in: what undoes setRelationship. */
  | { type: 'setRelationships'; story: string; relationships: Map<string, DocRelationship> }
  /** Adds, replaces or (null) removes a picture of DocModel.media; `dropMap` also removes the map itself once it is empty. */
  | { type: 'setMedia'; name: string; media: DocMedia | null; dropMap?: boolean }
  /**
   * Puts another table of styles in the place of DocModel.styles: how styles
   * are added (a paste from another document brings its own). The table is
   * never changed in place, so that what depends on the styles (the layout)
   * sees a new object.
   */
  | { type: 'setStyles'; styles: DocStyles }
  /**
   * Puts other properties in the place of those of the last section of the
   * body (Body.section). Without `span` the section has no place in the
   * source any more and is written from the model; the inverse carries the
   * span the section had, so that undoing gives the text of the file back.
   */
  | { type: 'setBodySection'; section: SectionProps; span?: SourceSpan; modified?: Flag }
  /** Turns the "track changes" setting of the document (DocSettings.trackRevisions) on or off. */
  | { type: 'setTrackRevisions'; on: boolean }
  /** Adds, replaces or (null) removes entries of the font table, by font name. The table itself stays the object it is. */
  | { type: 'setFonts'; fonts: [name: string, value: FontInfo | null][] };

const bad = (message: string): never => {
  throw new RangeError(message);
};

function setFlag(target: { modified?: boolean }, flag: Flag | undefined): Flag {
  const before: Flag = target.modified === undefined ? null : target.modified;
  if (flag === null) delete target.modified;
  else target.modified = flag === undefined ? true : flag;
  return before;
}

export function checkSplice(blocks: Block[], at: number, remove: number, insert: Block[]): void {
  const whole = (value: number) => typeof value === 'number' && value >= 0 && Math.floor(value) === value;
  if (!whole(at) || !whole(remove) || at + remove > blocks.length) bad('The blocks to replace are outside their list');
  if (!Array.isArray(insert)) bad('The blocks to insert must be a list');
  for (const block of insert) if (!block || typeof block.type !== 'string') bad('Only blocks can be inserted');
}

/** A copy of a list of blocks with `remove` of them at `at` replaced by `insert`. */
export function splicedBlocks(blocks: Block[], at: number, remove: number, insert: Block[]): Block[] {
  checkSplice(blocks, at, remove, insert);
  const out = blocks.slice(0, at);
  for (let i = 0; i < insert.length; i++) out.push(insert[i]);
  for (let i = at + remove; i < blocks.length; i++) out.push(blocks[i]);
  return out;
}

/** Most blocks handed to Array.splice at once; beyond that the list is built anew. */
const SPLICE_LIMIT = 4096;

/**
 * A copy of a table in which the blocks of the cell at `path` ([row, cell],
 * then [table, row, cell] for each nested table) are replaced: every object
 * on the way is new, everything beside it is shared. The copy has no source
 * span.
 */
export function tableWithBlocks(table: Table, path: number[], at: number, remove: number, insert: Block[], depth = 0): Table {
  if (depth > 512) bad('The path is too deep');
  const row: TableRow | undefined = table.rows[path[0]];
  const cell: TableCell | undefined = row && row.cells[path[1]];
  if (!cell || path.length < 2) return bad('The path does not lead to a table cell');
  let blocks: Block[];
  if (path.length === 2) blocks = splicedBlocks(cell.blocks, at, remove, insert);
  else {
    const inner = cell.blocks[path[2]];
    if (!inner || inner.type !== 'table') return bad(path[3] === -1 ? 'The content of a text box cannot be edited' : 'The path does not lead to a table cell');
    blocks = cell.blocks.slice();
    blocks[path[2]] = tableWithBlocks(inner, path.slice(3), at, remove, insert, depth + 1);
  }
  const cells = row.cells.slice();
  cells[path[1]] = { ...cell, blocks };
  const rows = table.rows.slice();
  rows[path[0]] = { ...row, cells };
  const out: Table = { ...table, rows };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

function replaceTop(doc: DocModel, handle: StoryHandle, op: Extract<DocOp, { type: 'replaceBlocks' }>, at: number, remove: number, insert: Block[]): DocOp {
  if (handle.note) {
    // A note has a source span of its own, so it is replaced like a block of a story.
    const note: Note = { ...handle.note, blocks: splicedBlocks(handle.blocks, at, remove, insert) };
    delete note.srcStart;
    delete note.srcEnd;
    return applyDocOp(doc, { type: 'setNote', story: op.story, note, modified: op.modified });
  }
  const story = handle.host as Story;
  const blocks = story.blocks;
  checkSplice(blocks, at, remove, insert);
  const removed = blocks.slice(at, at + remove);
  if (remove === 1 && insert.length === 1) blocks[at] = insert[0];
  else if (insert.length <= SPLICE_LIMIT) blocks.splice(at, remove, ...insert);
  else story.blocks = splicedBlocks(blocks, at, remove, insert);
  const inverse: DocOp = { type: 'replaceBlocks', story: op.story, path: [], at, remove: insert.length, insert: removed, modified: setFlag(story, op.modified) };
  if (handle.kind === 'body') {
    const body = story as Body;
    const before = body.sectionIndex;
    if (op.sectionIndex !== undefined) {
      if (op.sectionIndex === null) delete body.sectionIndex;
      else body.sectionIndex = op.sectionIndex;
    } else if (before !== undefined && at < before) {
      body.sectionIndex = at + remove <= before ? before + insert.length - remove : at + insert.length;
    }
    if (body.sectionIndex !== before) inverse.sectionIndex = before === undefined ? null : before;
  }
  return inverse;
}

/** Applies one operation to the document and returns the operation that undoes it. Throws RangeError for an operation that does not fit the document, which is then unchanged. */
export function applyDocOp(doc: DocModel, op: DocOp): DocOp {
  switch (op.type) {
    case 'replaceBlocks': {
      const handle = storyOf(doc, op.story);
      if (!Array.isArray(op.path)) return bad('An operation needs a path');
      if (op.path.length === 0) return replaceTop(doc, handle, op, op.at, op.remove, op.insert);
      const top = handle.blocks[op.path[0]];
      if (!top || top.type !== 'table') return bad(op.path[1] === -1 ? 'The content of a text box cannot be edited' : 'The path does not lead to a table cell');
      return replaceTop(doc, handle, op, op.path[0], 1, [tableWithBlocks(top, op.path.slice(1), op.at, op.remove, op.insert)]);
    }
    case 'setNote': {
      const handle = storyOf(doc, op.story);
      if (!handle.note || handle.noteId === undefined || !op.note || op.note.id !== handle.noteId) return bad('The note does not belong to the story');
      const collection = handle.host as { items: Map<number, Note>; modified?: boolean };
      // Setting an existing key keeps the order of the notes.
      collection.items.set(handle.noteId, op.note);
      return { type: 'setNote', story: op.story, note: handle.note, modified: setFlag(collection, op.modified) };
    }
    case 'removeNote': {
      const handle = storyOf(doc, op.story);
      if (!handle.note || handle.noteId === undefined || (handle.kind !== 'footnote' && handle.kind !== 'endnote')) return bad('The story is not a note');
      const collection = handle.host as StoryCollection<Note>;
      const before = collection.items;
      // A new table: the old one goes back as it is, with the note at its place among the others.
      const next = new Map(before);
      next.delete(handle.noteId);
      collection.items = next;
      return { type: 'setNotes', kind: handle.kind, items: before, modified: setFlag(collection, op.modified) };
    }
    case 'setNotes': {
      const collection = op.kind === 'footnote' ? doc.footnotes : op.kind === 'endnote' ? doc.endnotes : undefined;
      if (!collection || !(op.items instanceof Map)) return bad('The document has no such notes');
      const before = collection.items;
      collection.items = op.items;
      return { type: 'setNotes', kind: op.kind, items: before, modified: setFlag(collection, op.modified) };
    }
    case 'setNumbering': {
      const numbering = doc.numbering;
      const inverse: Extract<DocOp, { type: 'setNumbering' }> = { type: 'setNumbering' };
      const put = <T>(map: Map<number, T>, entries: [number, T | null][] | undefined): [number, T | null][] | undefined => {
        if (!entries) return undefined;
        const before: [number, T | null][] = [];
        for (const [id] of entries) if (typeof id !== 'number' || Math.floor(id) !== id) bad('A numbering id must be a whole number');
        // Undone last first, so that an id given twice ends up as it began.
        for (const [id, value] of entries) {
          const old = map.get(id);
          before.unshift([id, old === undefined ? null : old]);
          if (value === null) map.delete(id);
          else map.set(id, value);
        }
        return before;
      };
      const abstracts = put(numbering.abstracts, op.abstracts);
      const nums = put(numbering.nums, op.nums);
      if (abstracts) inverse.abstracts = abstracts;
      if (nums) inverse.nums = nums;
      inverse.modified = setFlag(numbering, op.modified);
      return inverse;
    }
    case 'setRelationship': {
      const host = storyOf(doc, op.story).host;
      if (typeof op.id !== 'string' || op.id === '') return bad('A relationship needs an id');
      const before = host.relationships;
      // A copy: the readers of the binary formats share one empty table between all their stories.
      const next = new Map(before);
      if (op.relationship === null) next.delete(op.id);
      else next.set(op.id, op.relationship);
      host.relationships = next;
      return { type: 'setRelationships', story: op.story, relationships: before };
    }
    case 'setRelationships': {
      const host = storyOf(doc, op.story).host;
      const before = host.relationships;
      host.relationships = op.relationships;
      return { type: 'setRelationships', story: op.story, relationships: before };
    }
    case 'setStyles': {
      const next = op.styles;
      if (!next || typeof next !== 'object' || !(next.styles instanceof Map) || !next.defaults || !next.defaultIds) return bad('A table of styles is needed');
      const before = doc.styles;
      doc.styles = next;
      return { type: 'setStyles', styles: before };
    }
    case 'setBodySection': {
      const body = doc.body;
      if (!op.section || typeof op.section !== 'object' || !body.section) return bad('The body has no section to replace');
      const inverse: DocOp = { type: 'setBodySection', section: body.section };
      if (body.sectionSpan) inverse.span = body.sectionSpan;
      body.section = op.section;
      if (op.span) body.sectionSpan = op.span;
      else delete body.sectionSpan;
      inverse.modified = setFlag(body, op.modified);
      return inverse;
    }
    case 'setTrackRevisions': {
      if (typeof op.on !== 'boolean' || !doc.settings) return bad('Tracking is either on or off');
      const before = doc.settings.trackRevisions === true;
      doc.settings.trackRevisions = op.on;
      return { type: 'setTrackRevisions', on: before };
    }
    case 'setFonts': {
      const fonts = doc.fonts;
      if (!Array.isArray(op.fonts) || !(fonts instanceof Map)) return bad('Entries of the font table are needed');
      for (const entry of op.fonts) {
        if (!Array.isArray(entry) || typeof entry[0] !== 'string' || entry[0] === '') return bad('A font needs a name');
        const value = entry[1];
        if (value === null) continue;
        if (!value || typeof value !== 'object' || value.name !== entry[0]) return bad('An entry of the font table carries the name it is kept under');
        for (const field of ['altName', 'panose', 'charset', 'family', 'pitch'] as const) {
          if (value[field] !== undefined && typeof value[field] !== 'string') return bad('What a font table says of a font is text');
        }
        if (value.embedded !== undefined && !(value.embedded instanceof Map)) return bad('The embedded files of a font are kept by face');
      }
      const before: [string, FontInfo | null][] = [];
      // Undone last first, so that a name given twice ends up as it began.
      for (const [name, value] of op.fonts) {
        const old = fonts.get(name);
        before.unshift([name, old === undefined ? null : old]);
        if (value === null) fonts.delete(name);
        else fonts.set(name, value);
      }
      return { type: 'setFonts', fonts: before };
    }
    case 'setMedia': {
      if (typeof op.name !== 'string' || op.name === '') return bad('A picture needs a name');
      const had = doc.media !== undefined;
      const media = doc.media || (doc.media = new Map<string, DocMedia>());
      const old = media.get(op.name);
      if (op.media === null) media.delete(op.name);
      else media.set(op.name, op.media);
      if (op.dropMap && media.size === 0) delete doc.media;
      const inverse: DocOp = { type: 'setMedia', name: op.name, media: old === undefined ? null : old };
      if (!had) inverse.dropMap = true;
      return inverse;
    }
  }
  return bad('Unknown operation');
}

/**
 * Applies operations in order and returns their inverses, last first. If one
 * of them throws, the ones already applied are undone and the error is passed
 * on: the document is changed by all of them or by none.
 */
export function applyDocOps(doc: DocModel, ops: DocOp[]): DocOp[] {
  const inverse: DocOp[] = [];
  try {
    for (const op of ops) inverse.push(applyDocOp(doc, op));
  } catch (error) {
    for (let i = inverse.length - 1; i >= 0; i--) applyDocOp(doc, inverse[i]);
    throw error;
  }
  return inverse.reverse();
}
