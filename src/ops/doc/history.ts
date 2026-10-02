// Undo and redo for documents: the counterpart of History in ../ops.ts, which
// is tied to workbooks. The interface is the same, plus merging of typing
// into one step.

import type { Block, DocModel, Table } from '../../model/doc/types';
import { applyDocOps, type DocOp } from './ops';
import { storyOf, type DocRange } from './position';

/**
 * A step keeps the blocks it replaced. Blocks share what did not change, but
 * a step that replaced many of them (a large paste, a deleted chapter) holds
 * on to all of those, so steps are dropped by what they hold as well as by
 * their number. The unit is roughly one paragraph kept alive by the step.
 */
const MAX_UNDO_WEIGHT = 400000;

/** Most cells of one table that are looked at to weigh a step; beyond that the rest counts one each. */
const MAX_WEIGHED_CELLS = 200000;

function blocksWeight(blocks: Block[], budget: { left: number }, depth: number): number {
  let weight = 0;
  for (const block of blocks) weight += block.type === 'table' ? tableWeight(block, undefined, budget, depth + 1) : 1;
  return weight;
}

/**
 * What a step holds by keeping `old` when the document now has `now` in its
 * place: the rows and cells the two do not share. A table is copied row by
 * row and cell by cell, so a change in one cell of a large table holds one
 * row and one cell, not the table.
 */
function tableWeight(old: Table, now: Table | undefined, budget: { left: number }, depth: number): number {
  // The list of rows itself.
  let weight = 1 + (old.rows.length >> 6);
  if (depth > 64) return weight;
  let live: Set<object> | undefined;
  if (now && now.rows.length !== old.rows.length) live = new Set<object>(now.rows);
  for (let r = 0; r < old.rows.length; r++) {
    const row = old.rows[r];
    const other = now ? now.rows[r] : undefined;
    if (now && (live ? live.has(row) : other === row)) continue;
    weight++;
    const aligned = !live && other && other.cells.length === row.cells.length ? other : undefined;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      const twin = aligned ? aligned.cells[c] : undefined;
      if (twin === cell) continue;
      if (budget.left-- <= 0) {
        weight++;
        continue;
      }
      if (!twin || twin.blocks === cell.blocks) {
        weight += twin ? 1 : blocksWeight(cell.blocks, budget, depth);
        continue;
      }
      // Blocks of the cell the document no longer has.
      const blocks = cell.blocks;
      const others = twin.blocks;
      for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (others[i] === block) continue;
        const beside = others[i];
        weight += block.type === 'table' ? tableWeight(block, beside && beside.type === 'table' ? beside : undefined, budget, depth + 1) : 1;
      }
    }
  }
  return weight;
}

/**
 * What the operations of a step (the inverses of what was just applied) keep
 * alive. Called right after they were made: the document holds the new
 * blocks where the operations hold the old ones.
 */
function weightOf(doc: DocModel, ops: DocOp[]): number {
  let weight = 0;
  for (const op of ops) {
    weight++;
    if (op.type === 'setNote') weight += op.note.blocks.length;
    if (op.type !== 'replaceBlocks') continue;
    const budget = { left: MAX_WEIGHED_CELLS };
    let now: Block[] | undefined;
    if (op.path.length === 0 && op.insert.length === op.remove) {
      try {
        now = storyOf(doc, op.story).blocks;
      } catch (error) {
        now = undefined;
      }
    }
    for (let i = 0; i < op.insert.length; i++) {
      const block = op.insert[i];
      const beside = now ? now[op.at + i] : undefined;
      weight += block.type === 'table' ? tableWeight(block, beside && beside.type === 'table' ? beside : undefined, budget, 0) : 1;
    }
  }
  return weight;
}

interface Step<M> {
  ops: DocOp[];
  before: M;
  after: M;
  /** Steps with the same key that follow each other become one. */
  key?: string;
  /** What the step holds (see weightOf), as it was when the step was made. */
  weight: number;
}

/** What a command hands back: see DocEdit in ./draft.ts. */
interface EditLike<M> {
  ops: DocOp[];
  selection: M;
  mergeKey?: string;
  seal?: boolean;
}

/** One block of a story replaced by one block: the shape of the operation that typing a character comes to. */
const slot = (op: DocOp): string | undefined =>
  op.type === 'replaceBlocks' && op.path.length === 0 && op.remove === 1 && op.insert.length === 1 ? op.story + '\u0001' + op.at : op.type === 'setNote' ? op.story : undefined;

/** Undo and redo over applyDocOp(). `M` is what the caller wants back with each step: the selection before and after it. */
export class DocHistory<M = DocRange | undefined> {
  private undoStack: Step<M>[] = [];
  private redoStack: Step<M>[] = [];
  private weight = 0;
  private sealed = true;

  constructor(
    readonly doc: DocModel,
    private readonly limit = 200,
  ) {}

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Applies operations as one undoable step. If one of them throws, the ones
   * already applied are undone and the error is passed on. With a `mergeKey`
   * the step joins the one before it when that has the same key and seal()
   * was not called in between: typing a word is one step.
   */
  apply(ops: DocOp[], before: M, after: M, mergeKey?: string): DocOp[] {
    if (ops.length === 0) return [];
    const inverse = applyDocOps(this.doc, ops);
    const last = this.undoStack[this.undoStack.length - 1];
    this.redoStack = [];
    if (last && mergeKey !== undefined && last.key === mergeKey && !this.sealed) {
      // Undoing the older change of the same block also undoes the newer one, which so need not be kept.
      const same = inverse.length === 1 && last.ops.length === 1 && slot(inverse[0]) !== undefined && slot(inverse[0]) === slot(last.ops[0]);
      if (!same) {
        const added = weightOf(this.doc, inverse);
        this.weight += added;
        last.weight += added;
        last.ops = inverse.concat(last.ops);
      }
      last.after = after;
      return inverse;
    }
    const weight = weightOf(this.doc, inverse);
    this.undoStack.push({ ops: inverse, before, after, key: mergeKey, weight });
    this.sealed = false;
    this.weight += weight;
    this.trim();
    return inverse;
  }

  /** The oldest steps go when there are too many of them or they hold too much; the newest always stays. */
  private trim(): void {
    while (this.undoStack.length > 1 && (this.undoStack.length > this.limit || this.weight > MAX_UNDO_WEIGHT)) {
      this.weight -= (this.undoStack.shift() as Step<M>).weight;
    }
  }

  /** Applies what a command returned: its operations, with `before` and the selection the command leaves. */
  applyEdit(edit: EditLike<M>, before: M): DocOp[] {
    const inverse = this.apply(edit.ops, before, edit.selection, edit.mergeKey);
    if (edit.seal) this.seal();
    return inverse;
  }

  /** Ends the step that is being merged into: call when the caret moves for another reason than typing. */
  seal(): void {
    this.sealed = true;
  }

  /** Undoes the last step and returns what was passed as `before` for it. */
  undo(): { meta: M } | undefined {
    const step = this.undoStack.pop();
    if (!step) return undefined;
    this.sealed = true;
    this.weight -= step.weight;
    this.redoStack.push({ ops: applyDocOps(this.doc, step.ops), before: step.before, after: step.after, weight: 0 });
    return { meta: step.before };
  }

  redo(): { meta: M } | undefined {
    const step = this.redoStack.pop();
    if (!step) return undefined;
    this.sealed = true;
    const inverse = applyDocOps(this.doc, step.ops);
    const weight = weightOf(this.doc, inverse);
    this.weight += weight;
    this.undoStack.push({ ops: inverse, before: step.before, after: step.after, weight });
    return { meta: step.after };
  }

  /** Forgets every step, e.g. after the document was saved and loaded again. */
  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
    this.weight = 0;
    this.sealed = true;
  }
}
