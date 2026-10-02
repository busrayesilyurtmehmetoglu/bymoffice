// Undo and redo for documents: the counterpart of History in ../ops.ts, which
// is tied to workbooks. The interface is the same, plus merging of typing
// into one step.

import type { DocModel } from '../../model/doc/types';
import { applyDocOps, type DocOp } from './ops';
import type { DocRange } from './position';

/**
 * A step keeps the blocks it replaced. Blocks share what did not change, but
 * a step that replaced many of them (a large paste, a deleted chapter) holds
 * on to all of those, so steps are dropped by what they hold as well as by
 * their number.
 */
const MAX_UNDO_WEIGHT = 400000;

function weightOf(ops: DocOp[]): number {
  let weight = 0;
  for (const op of ops) weight += op.type === 'replaceBlocks' ? 1 + op.insert.length : 1;
  return weight;
}

interface Step<M> {
  ops: DocOp[];
  before: M;
  after: M;
  /** Steps with the same key that follow each other become one. */
  key?: string;
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
        this.weight += weightOf(inverse);
        last.ops = inverse.concat(last.ops);
      }
      last.after = after;
      return inverse;
    }
    this.undoStack.push({ ops: inverse, before, after, key: mergeKey });
    this.sealed = false;
    this.weight += weightOf(inverse);
    // The oldest steps go when there are too many of them or they hold too much; the newest always stays.
    while (this.undoStack.length > 1 && (this.undoStack.length > this.limit || this.weight > MAX_UNDO_WEIGHT)) {
      this.weight -= weightOf((this.undoStack.shift() as Step<M>).ops);
    }
    return inverse;
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
    this.weight -= weightOf(step.ops);
    this.redoStack.push({ ops: applyDocOps(this.doc, step.ops), before: step.before, after: step.after });
    return { meta: step.before };
  }

  redo(): { meta: M } | undefined {
    const step = this.redoStack.pop();
    if (!step) return undefined;
    this.sealed = true;
    const inverse = applyDocOps(this.doc, step.ops);
    this.weight += weightOf(inverse);
    this.undoStack.push({ ops: inverse, before: step.before, after: step.after });
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
