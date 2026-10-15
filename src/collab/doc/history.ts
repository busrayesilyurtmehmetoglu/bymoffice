// The history of a document that is edited together: what a view of a document uses in the
// place of DocHistory (../../ops/doc/history.ts) while it is in a room. It has the same
// methods; behind them every change is a step of the session, and undo and redo are steps as
// well (see session.ts).

import type { DocOp } from '../../ops/doc/ops';
import { CollabUnsupportedError } from '../types';
import type { DocCollabSession } from './session';
import { wiresOf } from './trace';

/** What a view asks of the history of its document, whichever it is. */
export interface DocHistoryLike<M> {
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  apply(ops: DocOp[], before: M, after: M, mergeKey?: string): DocOp[];
  applyEdit(edit: { ops: DocOp[]; selection: M; mergeKey?: string; seal?: boolean }, before: M): DocOp[];
  seal(): void;
  undo(): { meta: M } | undefined;
  redo(): { meta: M } | undefined;
  clear(): void;
}

export { holdsRevisions } from './session';

/** Said for a change that does not travel: a picture, a paste with its formatting, an accepted revision. */
export const NOT_SHARED = 'This change cannot be made while the document is edited together';

export class CollabDocHistory<M> implements DocHistoryLike<M> {
  constructor(private readonly session: DocCollabSession) {}

  get canUndo(): boolean {
    return this.session.canUndo;
  }

  get canRedo(): boolean {
    return this.session.canRedo;
  }

  /**
   * Makes the change the operations stand for as a step of the session: the command that
   * returned them is run again by the session, here as by everybody else. Throws
   * CollabUnsupportedError for operations no shared command returned; nothing was changed then.
   */
  apply(ops: DocOp[], before: M, after: M, mergeKey?: string): DocOp[] {
    if (ops.length === 0) return [];
    const wires = wiresOf(ops);
    if (!wires) throw new CollabUnsupportedError('other', NOT_SHARED);
    // Letters that are typed one after the other wait for the relay as one step.
    this.session.submit(wires, { before, after, key: mergeKey, join: mergeKey !== undefined });
    return [];
  }

  applyEdit(edit: { ops: DocOp[]; selection: M; mergeKey?: string; seal?: boolean }, before: M): DocOp[] {
    const inverse = this.apply(edit.ops, before, edit.selection, edit.mergeKey);
    if (edit.seal) this.seal();
    return inverse;
  }

  seal(): void {
    this.session.seal();
  }

  /** Undefined also while the undoing waits for changes that are on their way: the view is told when it was done (DocChange.local). */
  undo(): { meta: M } | undefined {
    const done = this.session.undo();
    return done ? { meta: done.meta as M } : undefined;
  }

  redo(): { meta: M } | undefined {
    const done = this.session.redo();
    return done ? { meta: done.meta as M } : undefined;
  }

  clear(): void {
    // What can be undone is the session's to keep: it is forgotten when the file is loaded again.
  }
}
