// How commands are built. A command works on a draft: every change it makes
// is applied to the document at once, so that the next step of the command
// sees it, and recorded. When the command is done the draft takes everything
// back and hands out the recorded operations; the caller applies them through
// a DocHistory. A command that throws leaves the document as it was.

import type { TableContext } from '../../formats/docx/resolve';
import type { Block, DocModel, Paragraph, Table, Wrapper } from '../../model/doc/types';
import { applyDocOp, tableWithBlocks, type DocOp } from './ops';
import { collapsed, containerBlocks, inTextBox, storyOf, type DocPosition, type DocRange } from './position';

/** What a command returns: the operations to apply and the selection they leave. */
export interface DocEdit {
  ops: DocOp[];
  selection: DocRange;
  /** Set for typing: steps with the same key become one undo step (see DocHistory.apply). */
  mergeKey?: string;
  /** The step ends what is being merged: the character typed was a space. */
  seal?: boolean;
}

export interface DocEditOptions {
  /**
   * Reserved. Documents that track changes (DocSettings.trackRevisions) are
   * edited directly for now: no revision marks are made, whatever this says.
   */
  trackChanges?: boolean;
}

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

const READ_ONLY = 'This part of the document cannot be edited';

export class Draft {
  readonly ops: DocOp[] = [];
  private readonly inverses: DocOp[] = [];

  constructor(readonly doc: DocModel) {}

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
   * story is always a replacement of blocks of the story itself: a change
   * inside a table replaces the table by a copy.
   *
   * Blocks in a block-level wrapper (a content control around paragraphs)
   * have no source span of their own; the wrapper has it. When such a block
   * changes, the blocks of the wrapper all get a copy of it without the
   * span, so that the writer does not copy the old text of the wrapper.
   */
  replace(story: string, path: number[], at: number, remove: number, insert: Block[]): void {
    if (inTextBox(path)) bad('The content of a text box cannot be edited');
    const handle = storyOf(this.doc, story);
    const top = handle.blocks;
    if (path.length > 0) {
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
      const table = top[path[0]] as Table;
      this.replaceTop(story, top, path[0], 1, [tableWithBlocks(table, path.slice(1), at, remove, insert)]);
      return;
    }
    for (let i = at; i < at + remove && i < top.length; i++) if (inAlternate(top[i].wrap)) bad(READ_ONLY);
    this.replaceTop(story, top, at, remove, insert);
  }

  private replaceTop(story: string, top: Block[], at: number, remove: number, insert: Block[]): void {
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
    const copies = new Map<Wrapper, Wrapper>();
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
}

/** Runs a command on a draft and returns its operations; the document is as before when this returns or throws. */
export function buildEdit(doc: DocModel, command: (draft: Draft) => DocRange | DocPosition | Omit<DocEdit, 'ops'>): DocEdit {
  const draft = new Draft(doc);
  try {
    const result = command(draft);
    if ('selection' in result) return { ...result, ops: draft.ops };
    return { ops: draft.ops, selection: 'anchor' in result ? result : collapsed(result) };
  } finally {
    draft.rollback();
  }
}

/** Where a paragraph sits when its path goes through a table: what the style resolution needs. */
export function tableContextOf(doc: DocModel, story: string, path: number[]): TableContext | undefined {
  if (path.length < 4 || path[path.length - 4] === -1) return undefined;
  const tablePath = path.slice(0, -3);
  const blocks = containerBlocks(doc, story, tablePath.slice(0, -1));
  const table = blocks[tablePath[tablePath.length - 1]];
  if (!table || table.type !== 'table') return undefined;
  return { table, row: path[path.length - 3], cell: path[path.length - 2] };
}
