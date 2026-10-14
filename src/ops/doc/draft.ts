// How commands are built. A command works on a draft: every change it makes
// is applied to the document at once, so that the next step of the command
// sees it, and recorded. When the command is done the draft takes everything
// back and hands out the recorded operations; the caller applies them through
// a DocHistory. A command that throws leaves the document as it was.

import { editsOf, restoreEdits } from '../../model/edits';
import type { Block, DocModel, Paragraph, RunItem, Story, Table, TableCell, TableRow, Wrapper } from '../../model/doc/types';
import { applyDocOp, checkSplice, splicedBlocks, type DocOp } from './ops';
import { READ_ONLY } from './pieces';
import { collapsed, inTextBox, storyOf, type DocPosition, type DocRange } from './position';
import { TRACK_UNSUPPORTED, distinctRevisions, trackingFor, type TrackOptions, type Tracking } from './track';

export { tableContextOf } from './position';

/** What a command returns: the operations to apply and the selection they leave. */
export interface DocEdit {
  ops: DocOp[];
  selection: DocRange;
  /** Set for typing: steps with the same key become one undo step (see DocHistory.apply). */
  mergeKey?: string;
  /** The step ends what is being merged: the character typed was a space. */
  seal?: boolean;
}

/**
 * How a command tracks its changes. A document that asks for it
 * (DocSettings.trackRevisions, or a protection that allows only tracked
 * changes) is always edited with tracked changes; `trackChanges` does the
 * same for any other document. A command that cannot make its change as a
 * tracked one throws RangeError(TRACK_UNSUPPORTED) and changes nothing.
 */
export type DocEditOptions = TrackOptions;

const bad = (message: string): never => {
  throw new RangeError(message);
};

function rootWrapper(wrap: Wrapper | undefined): Wrapper | undefined {
  let root = wrap;
  while (root && root.parent) root = root.parent;
  return root;
}

function inAlternate(wrap: Wrapper | undefined): boolean {
  for (let w = wrap; w; w = w.parent) if (w.kind === 'alternate') return true;
  return false;
}

/** Most blocks handed to Array.splice at once; beyond that the list is built anew. */
const SPLICE_LIMIT = 4096;

/**
 * How many copies drafts have made of tables, rows and cells since the module
 * was loaded. For tests: a command over a large table must copy the table
 * once, not once per paragraph.
 */
export const draftCounters = { tables: 0, rows: 0, cells: 0 };

/** Adds the notes the items of a run refer to, as the ids of their stories ("footnote:2"). */
function noteRefsOfItems(items: RunItem[], out: Set<string>, depth: number): void {
  for (const item of items) {
    if (item.type === 'footnoteReference') out.add('footnote:' + item.id);
    else if (item.type === 'endnoteReference') out.add('endnote:' + item.id);
    else if (depth < 64) {
      if (item.type === 'alternate') noteRefsOfItems(item.items, out, depth + 1);
      else if ((item.type === 'drawing' || item.type === 'pict' || item.type === 'object') && item.textBoxes) {
        for (const box of item.textBoxes) noteRefs(box, undefined, undefined, out, depth + 1);
      }
    }
  }
}

/**
 * Adds the notes that blocks refer to. Blocks, rows and cells in `skip` are
 * passed over; the ones visited are added to `seen`.
 */
function noteRefs(blocks: Block[], skip: Set<object> | undefined, seen: Set<object> | undefined, out: Set<string>, depth = 0): void {
  if (depth > 256) return;
  for (const block of blocks) {
    if (skip && skip.has(block)) continue;
    if (seen) seen.add(block);
    if (block.type === 'paragraph') {
      for (const inline of block.children) if (inline.type === 'run') noteRefsOfItems(inline.content, out, depth);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        if (skip && skip.has(row)) continue;
        if (seen) seen.add(row);
        for (const cell of row.cells) {
          if (skip && skip.has(cell)) continue;
          if (seen) seen.add(cell);
          noteRefs(cell.blocks, skip, seen, out, depth + 1);
        }
      }
    }
  }
}

export class Draft {
  readonly ops: DocOp[] = [];
  private readonly inverses: DocOp[] = [];
  /** Block-level wrappers this draft has copied (to drop their source span), by the wrapper they were. */
  private readonly copies = new Map<Wrapper, Wrapper>();
  /**
   * Tables, rows and cells this draft made and put into the document. Only
   * its own operations refer to them, so a further change inside the same
   * table is made in them directly instead of copying the table again.
   */
  private readonly owned = new WeakSet<object>();
  /** Notes a reference to which was removed: they go when nothing refers to them at the end (see finish). */
  private readonly unreferenced = new Set<string>();
  private readonly hasNotes: boolean;

  /**
   * @param track Set when the changes of the command are tracked: what it puts in and takes out are revisions.
   */
  constructor(
    readonly doc: DocModel,
    readonly track?: Tracking,
  ) {
    let notes = false;
    for (const collection of [doc.footnotes, doc.endnotes]) {
      if (!collection || notes) continue;
      for (const note of collection.items.values()) {
        if (note.kind === undefined) {
          notes = true;
          break;
        }
      }
    }
    this.hasNotes = notes;
  }

  apply(op: DocOp): void {
    this.inverses.push(applyDocOp(this.doc, op));
    this.ops.push(op);
  }

  /** Takes back everything the draft did. */
  rollback(): void {
    for (let i = this.inverses.length - 1; i >= 0; i--) applyDocOp(this.doc, this.inverses[i]);
    this.inverses.length = 0;
  }

  /**
   * Replaces blocks of the list at `path` of a story. What goes into the
   * story is always a replacement of blocks of the story itself: the first
   * change inside a table replaces the table by a copy (one operation), and
   * the changes that follow in the same table are made in that copy, so a
   * command over every cell of a table costs one copy of its list of rows,
   * one of each row and cell it touches, and one operation.
   *
   * Blocks in a block-level wrapper (a content control around paragraphs)
   * have no source span of their own; the wrapper has it. When such a block
   * changes, the blocks of the wrapper all get a copy of it without the
   * span, so that the writer does not copy the old text of the wrapper. The
   * draft remembers the copy: blocks that come later and still point to the
   * wrapper as it was are bound to the same copy, so one command leaves one
   * wrapper however many of its blocks it changes.
   */
  replace(story: string, path: number[], at: number, remove: number, incoming: Block[]): void {
    if (inTextBox(path)) bad('The content of a text box cannot be edited');
    // No two elements of what goes in share the id of a tracked change (see distinctRevisions).
    const top = storyOf(this.doc, story).blocks;
    const insert = Array.isArray(incoming) ? distinctRevisions(this.doc, incoming, this.standing(top, path, at, remove, incoming)) : incoming;
    if (path.length === 0) {
      for (let i = at; i < at + remove && i < top.length; i++) if (inAlternate(top[i].wrap)) bad(READ_ONLY);
      checkSplice(top, at, remove, insert);
      this.trackNotes(top, at, remove, insert);
      this.replaceTop(story, top, at, remove, insert);
      return;
    }
    // Everything the path goes through must be content the writer writes from the model.
    let blocks = top;
    for (let i = 0; i < path.length; i += 3) {
      const table = blocks[path[i]];
      if (!table || table.type !== 'table') return bad('The path does not lead to a table cell');
      const row = table.rows[path[i + 1]];
      const cell = row && row.cells[path[i + 2]];
      if (!cell) return bad('The path does not lead to a table cell');
      if (inAlternate(table.wrap) || inAlternate(row.wrap) || inAlternate(cell.wrap)) bad(READ_ONLY);
      blocks = cell.blocks;
    }
    for (let i = at; i < at + remove && i < blocks.length; i++) if (inAlternate(blocks[i].wrap)) bad(READ_ONLY);
    checkSplice(blocks, at, remove, insert);
    this.trackNotes(blocks, at, remove, insert);

    let table = top[path[0]] as Table;
    if (!this.owned.has(table)) {
      this.replaceTop(story, top, path[0], 1, [this.tableCopy(table)]);
      // The block that went in may be another object than the copy: one bound to the copy of a wrapper.
      table = storyOf(this.doc, story).blocks[path[0]] as Table;
      this.owned.add(table);
    }
    let cell = this.ownCell(table, path[1], path[2]);
    for (let i = 3; i < path.length; i += 3) {
      let inner = cell.blocks[path[i]] as Table;
      if (!this.owned.has(inner)) {
        inner = this.tableCopy(inner);
        cell.blocks[path[i]] = inner;
        this.owned.add(inner);
      }
      cell = this.ownCell(inner, path[i + 1], path[i + 2]);
    }
    if (remove === 1 && insert.length === 1) cell.blocks[at] = insert[0];
    else if (insert.length <= SPLICE_LIMIT) cell.blocks.splice(at, remove, ...insert);
    else cell.blocks = splicedBlocks(cell.blocks, at, remove, insert);
  }

  /**
   * The blocks among `incoming` that are in the list they go into already, as they are: a command that replaces a
   * stretch of blocks hands back the ones it did not change. Undefined when there is none.
   */
  private standing(top: Block[], path: number[], at: number, remove: number, incoming: Block[]): Set<Block> | undefined {
    if (!(remove > 0) || !(at >= 0)) return undefined;
    let blocks: Block[] | undefined = top;
    for (let i = 0; i < path.length && blocks; i += 3) {
      const table: Block | undefined = blocks[path[i]];
      const row: TableRow | undefined = table && table.type === 'table' ? table.rows[path[i + 1]] : undefined;
      const cell: TableCell | undefined = row ? row.cells[path[i + 2]] : undefined;
      blocks = cell ? cell.blocks : undefined;
    }
    if (!blocks) return undefined;
    const old = new Set<Block>();
    for (let i = at; i < at + remove && i < blocks.length; i++) old.add(blocks[i]);
    let out: Set<Block> | undefined;
    for (const block of incoming) if (old.has(block)) (out || (out = new Set<Block>())).add(block);
    return out;
  }

  /** A copy of a table with a list of rows of its own, and no source span. */
  private tableCopy(table: Table): Table {
    const copy: Table = { ...table, rows: table.rows.slice() };
    delete copy.srcStart;
    delete copy.srcEnd;
    draftCounters.tables++;
    return copy;
  }

  /** The cell of a table of this draft as a cell of this draft: the row and the cell are copied the first time. */
  private ownCell(table: Table, r: number, c: number): TableCell {
    let row = table.rows[r];
    if (!this.owned.has(row)) {
      row = { ...row, cells: row.cells.slice() };
      table.rows[r] = row;
      this.owned.add(row);
      draftCounters.rows++;
    }
    let cell = row.cells[c];
    if (!this.owned.has(cell)) {
      cell = { ...cell, blocks: cell.blocks.slice() };
      row.cells[c] = cell;
      this.owned.add(cell);
      draftCounters.cells++;
    }
    return cell;
  }

  /** Notes the removed blocks refer to and the inserted ones do not: candidates for removal. */
  private trackNotes(blocks: Block[], at: number, remove: number, insert: Block[]): void {
    if (!this.hasNotes || remove === 0) return;
    const kept = new Set<object>();
    const still = new Set<string>();
    noteRefs(insert, undefined, kept, still);
    const gone = new Set<string>();
    noteRefs(blocks.slice(at, at + remove), kept, undefined, gone);
    gone.forEach((ref) => (still.has(ref) ? 0 : this.unreferenced.add(ref)));
  }

  private replaceTop(story: string, top: Block[], at: number, remove: number, incoming: Block[]): void {
    // Blocks that still point to a wrapper this draft has copied follow the copy.
    let insert = incoming;
    if (this.copies.size) {
      insert = incoming.map((block) => {
        const copy = block.wrap && this.copies.get(block.wrap);
        return copy ? ({ ...block, wrap: copy } as Block) : block;
      });
    }
    // Wrappers with a source span among what goes and what comes.
    let spanned: Set<Wrapper> | undefined;
    const note = (block: Block | undefined) => {
      const root = block && rootWrapper(block.wrap);
      if (root && root.srcStart !== undefined) (spanned || (spanned = new Set())).add(root);
    };
    for (let i = at; i < at + remove && i < top.length; i++) note(top[i]);
    for (const block of insert) note(block);
    if (!spanned) {
      this.apply({ type: 'replaceBlocks', story, path: [], at, remove, insert });
      return;
    }
    const roots = spanned;
    const inGroup = (block: Block | undefined) => {
      const root = block && rootWrapper(block.wrap);
      return root !== undefined && roots.has(root);
    };
    let from = at;
    let to = at + remove;
    while (from > 0 && inGroup(top[from - 1])) from--;
    while (to < top.length && inGroup(top[to])) to++;
    const copies = this.copies;
    const copyOf = (wrap: Wrapper): Wrapper => {
      let copy = copies.get(wrap);
      if (!copy) {
        copy = { ...wrap };
        delete copy.srcStart;
        delete copy.srcEnd;
        if (wrap.parent) copy.parent = copyOf(wrap.parent);
        copies.set(wrap, copy);
      }
      return copy;
    };
    const moved = (block: Block): Block => (inGroup(block) ? ({ ...block, wrap: copyOf(block.wrap as Wrapper) } as Block) : block);
    const blocks: Block[] = [];
    for (let i = from; i < at; i++) blocks.push(moved(top[i]));
    for (const block of insert) blocks.push(moved(block));
    for (let i = at + remove; i < to; i++) blocks.push(moved(top[i]));
    this.apply({ type: 'replaceBlocks', story, path: [], at: from, remove: to - from, insert: blocks });
  }

  /** Puts another paragraph where the one at `path` is. */
  setParagraph(story: string, path: number[], paragraph: Paragraph): void {
    this.replace(story, path.slice(0, -1), path[path.length - 1], 1, [paragraph]);
  }

  /**
   * Ends the command: a note whose last reference the command removed leaves
   * its collection, in the same undo step, so that the writer does not write
   * a note nothing refers to. The separators of the notes area stay.
   */
  finish(): void {
    if (this.unreferenced.size === 0) return;
    const doc = this.doc;
    const live = new Set<string>();
    const stories = new Set<Story>([doc.body]);
    doc.headers.forEach((story) => stories.add(story));
    doc.footers.forEach((story) => stories.add(story));
    stories.forEach((story) => noteRefs(story.blocks, undefined, undefined, live));
    for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) {
      if (collection) collection.items.forEach((item: { blocks: Block[] }) => noteRefs(item.blocks, undefined, undefined, live));
    }
    this.unreferenced.forEach((ref) => {
      if (live.has(ref)) return;
      const collection = ref.charAt(0) === 'f' ? doc.footnotes : doc.endnotes;
      const id = ref.slice(ref.indexOf(':') + 1);
      // An id no story id can name (storyOf) stays: such a note cannot be edited either.
      const note = collection && /^-?\d{1,9}$/.test(id) ? collection.items.get(Number(id)) : undefined;
      if (note && note.kind === undefined) this.apply({ type: 'removeNote', story: ref });
    });
    this.unreferenced.clear();
  }
}

/**
 * Runs a command on a draft and returns its operations; the document is as
 * before when this returns or throws. `tracks` says that the command makes
 * its changes as revisions when the draft tracks (Draft.track); a command
 * that does not say so is refused where changes are tracked, so that nothing
 * is changed without a trace.
 */
export function buildEdit(doc: DocModel, command: (draft: Draft) => DocRange | DocPosition | Omit<DocEdit, 'ops'>, options?: DocEditOptions, tracks = false): DocEdit {
  const tracking = trackingFor(doc, options);
  if (tracking && !tracks) bad(TRACK_UNSUPPORTED);
  const draft = new Draft(doc, tracking);
  // The command is tried on the document and taken back before anything else runs: the document is what it was,
  // and is counted as unchanged (a file that is being written from it is not written again for a look at a command).
  const counted = editsOf(doc);
  try {
    const result = command(draft);
    draft.finish();
    if ('selection' in result) return { ...result, ops: draft.ops };
    return { ops: draft.ops, selection: 'anchor' in result ? result : collapsed(result) };
  } finally {
    draft.rollback();
    restoreEdits(doc, counted);
  }
}
