// One participant's side of a workbook that is edited together.
//
// The relay puts the steps of all participants into one order. The workbook here is always
// that order applied to the file, with the steps of this participant that the relay has not
// placed yet (the pending ones) applied on top. A step of someone else is placed under the
// pending ones: they are taken back, the step is applied, and they are applied again, each read
// through what came before it (rebase.ts). A pending step that cannot be applied any more is
// dropped and said. Only one step is on its way at a time, and the relay takes it only when it
// was made on the latest order - so every step is applied by everybody on the very workbook its
// author applied it to, and applying it is all it takes to end up the same.

import type { Workbook } from '../model/sheet/types';
import { applyOp, type Op } from '../ops/ops';
import { MAX_STEP_BYTES, utf8Length } from './protocol';
import { carryOf, inverseShift, mapPosition, transformStep, type Carry } from './rebase';
import { RoomSession, type CollabNotice, type CollabNoticeCode, type RoomHost, type RoomOptions, type RoomPeer } from './room';
import { CollabUnsupportedError, WireError, type Rec, type Shift, type WireOp } from './types';
import { MAX_STEP_OPS, StyleTableFull, applyWire, decodeStep, encodeOp, encodeStep, shiftOf, stable, type Applied } from './wire';

export { PEER_COLORS, PROTOCOL_VERSION, type CollabNotice, type CollabNoticeCode, type CollabState, type CollabTransport, type RoomStamp } from './room';

/** Where a participant is: the sheet, the active cell and the selected range. */
export interface CollabPresence {
  sheet: number;
  row: number;
  col: number;
  range: { top: number; left: number; bottom: number; right: number };
}

export type CollabPeer = RoomPeer<CollabPresence>;

/** What a change of the workbook that the host did not make itself was: for bringing a view up to date. */
export interface SessionChange {
  /** The cells whose content changed; missing when cells moved or too many changed, which leaves every formula to be calculated again. */
  cells?: { sheet: number; row: number; col: number }[];
  /** Sizes, hidden rows or merged cells changed. */
  layout: boolean;
  /** The list of sheets, or a name in it, changed. */
  sheets: boolean;
  /** What moved, in order: the positions the host keeps (a selection, a cell that is being typed in) are read through these. */
  shifts: Shift[];
  /** The rules or the charts of a sheet changed. */
  lists?: boolean;
  /** The change is this participant's own (an undo that waited for its turn), not someone else's. */
  local?: boolean;
}

/** What the session needs of whoever holds the workbook. */
export interface SessionHost extends RoomHost {
  /** The workbook as it is now (another one after reload()). */
  workbook(): Workbook;
  /** The workbook was changed by something the host did not call for: a step of someone else, a step of its own that was dropped. */
  changed(change: SessionChange): void;
}

export type SessionOptions = RoomOptions<CollabPresence>;

/** What submit(), undo() and redo() answer with. */
export interface Submitted {
  change: SessionChange;
  /** What was left out of the step because it does not travel. */
  skipped: number;
  /** The operations of the model that take the step back exactly, last first: what a view reads the changed cells from. */
  inverse: Op[];
}

interface StackEntry {
  /** The operations that take the step back, in order; null when it cannot be taken back. */
  recs: Rec[] | null;
  lossy: boolean;
  /** Someone else changed or removed what it would bring back. */
  gone?: boolean;
  meta: unknown;
  /** The step moved rows, columns or sheets: what lies under it on the stack is about the workbook without it. */
  shifty: boolean;
  /** A step that takes it back (or makes it again) is on its way: it leaves the stack when that step has its place. */
  busy?: boolean;
}

interface Step {
  id: string;
  recs: Rec[];
  /** The operations of the model that take the step back exactly, in the order they are applied in. */
  inverse: Op[];
  /** What applying it changed. */
  change: SessionChange;
  /** A change, or the undoing or the making again of one. */
  kind: 'edit' | 'undo' | 'redo';
  /** The operations that take it back as a step of their own, to be applied right after it; null when it cannot be taken back. */
  undo: Rec[] | null;
  lossy: boolean;
  /** Something else was changed here after this undoing: what it took back is not to be made again. */
  final?: boolean;
  /** What the caller gave with the step: handed back with its undoing. */
  meta?: unknown;
  /** For an undoing or a making again: what it takes from its stack once it has its place. */
  from?: StackEntry;
}

/** How long an undo waits for the steps of its participant to be placed. */
const UNDO_WAIT_MS = 3000;

function emptyChange(): SessionChange {
  return { cells: [], layout: false, sheets: false, shifts: [] };
}

function addChange(change: SessionChange, applied: Applied): void {
  if (applied.layout) change.layout = true;
  if (applied.sheets) change.sheets = true;
  if (applied.lists) change.lists = true;
  if (!applied.cells) change.cells = undefined;
  else if (change.cells) for (const cell of applied.cells) change.cells.push(cell);
}

function readPresence(value: unknown): CollabPresence | undefined {
  if (!value || typeof value !== 'object') return undefined;
  const from = value as { s?: unknown; a?: unknown; r?: unknown };
  const whole = (n: unknown): n is number => typeof n === 'number' && Math.floor(n) === n && n >= 0 && n < 1048576;
  if (!whole(from.s) || !Array.isArray(from.a) || !Array.isArray(from.r) || from.a.length !== 2 || from.r.length !== 4) return undefined;
  if (!from.a.every(whole) || !from.r.every(whole)) return undefined;
  const [row, col] = from.a as number[];
  const [top, left, bottom, right] = from.r as number[];
  if (col >= 16384 || right >= 16384 || bottom < top || right < left) return undefined;
  return { sheet: from.s, row, col, range: { top, left, bottom, right } };
}

export class CollabSession extends RoomSession<Step, CollabPresence> {
  private readonly host: SessionHost;
  private readonly options: SessionOptions;
  /** Undo and redo that wait for the steps on their way. */
  private wanted: ('undo' | 'redo')[] = [];
  private wantedTimer: (() => void) | undefined;
  private undoStack: StackEntry[] = [];
  private redoStack: StackEntry[] = [];

  constructor(host: SessionHost, options: SessionOptions) {
    super(host, options);
    this.host = host;
    this.options = options;
    this.start();
  }

  get peers(): CollabPeer[] {
    return super.peers;
  }

  get user(): CollabPeer {
    return super.user;
  }


  get canUndo(): boolean {
    return this.undoStack.some((entry) => !entry.busy) || this.pendingSteps.some((step) => step.kind !== 'undo');
  }

  get canRedo(): boolean {
    return this.redoStack.some((entry) => !entry.busy) || this.pendingSteps.some((step) => step.kind === 'undo' && !step.final);
  }

  // ------------------------------------------------------------------ editing

  /**
   * Applies operations of the model as one step of this participant and queues it for the
   * others. Throws what the workbook throws for an operation it does not take, and
   * CollabUnsupportedError for what cannot be shared; nothing was changed then.
   */
  submit(ops: readonly Op[], meta?: unknown): Submitted {
    if (this.readonlyNow) throw new CollabUnsupportedError('readOnly', 'The relay takes no changes from this participant');
    // While the file is loaded again the workbook is about to be another one: what is changed in it now would be lost here and sent to the others.
    if (this.reloading) throw new CollabUnsupportedError('other', 'The file is being loaded again');
    if (ops.length > MAX_STEP_OPS) throw new CollabUnsupportedError('size', 'The change is too large to be shared in one step');
    const workbook = this.host.workbook();
    const recs: Rec[] = [];
    const inverse: Op[] = [];
    const undo: Rec[][] = [];
    const change = emptyChange();
    let lossy = false;
    let undoable = true;
    let skipped = 0;
    let bytes = 0;
    try {
      for (const op of ops) {
        const encoded = encodeOp(workbook, op);
        if (encoded.skipped) skipped++;
        if (!encoded.op) continue;
        // What is applied here is what the others read: the operation goes through its own text first.
        const text = JSON.stringify(encodeStep([encoded.op], workbook.styles));
        bytes += utf8Length(text);
        if (bytes > MAX_STEP_BYTES) throw new CollabUnsupportedError('size', 'The change is too large to be shared in one step');
        let wire: WireOp;
        try {
          wire = decodeStep(JSON.parse(text), workbook.styles)[0];
        } catch (error) {
          // What the others would refuse is not done here either (a text with a character no file holds, say).
          if (error instanceof WireError) throw new CollabUnsupportedError('other', 'The change holds something that cannot be shared');
          throw error;
        }
        const applied = applyWire(workbook, wire);
        recs.push(applied.wide ? { w: wire, nm: applied.nm, wide: true } : { w: wire, nm: applied.nm });
        for (let i = applied.inverse.length - 1; i >= 0; i--) inverse.unshift(applied.inverse[i]);
        if (applied.undo) undo.unshift(applied.undo);
        else undoable = false;
        if (applied.lossy) lossy = true;
        addChange(change, applied);
        const shift = shiftOf(wire, applied.nm, applied.wide);
        if (shift) change.shifts.push(shift);
      }
    } catch (error) {
      for (const op of inverse) applyOp(workbook, op);
      throw error;
    }
    if (recs.length === 0) return { change, skipped, inverse: [] };
    this.redoStack = [];
    for (const step of this.pendingSteps) if (step.kind === 'undo') step.final = true;
    // An undo that waited for its turn does not get it: what it was to take back is no longer the last thing done.
    if (this.wanted.length > 0) {
      this.wanted = [];
      this.host.notice({ code: 'undo-wait' });
    }
    this.placed({ id: '', recs, inverse, change, kind: 'edit', undo: undoable ? ([] as Rec[]).concat(...undo) : null, lossy, meta });
    return { change, skipped, inverse };
  }

  /**
   * Undoes the last step of this participant. What is undone is always a step the relay has
   * placed: while steps of this participant are still on their way, the undoing waits for them
   * (undefined is answered now, and SessionHost.changed() is told when it was done; when they do
   * not arrive in a few seconds it is said and nothing is undone). Undefined also when there is
   * nothing to undo or the step cannot be undone, which is said.
   */
  undo(): (Submitted & { meta: unknown }) | undefined {
    return this.ask('undo');
  }

  redo(): (Submitted & { meta: unknown }) | undefined {
    return this.ask('redo');
  }

  private ask(kind: 'undo' | 'redo'): (Submitted & { meta: unknown }) | undefined {
    if (this.readonlyNow || this.reloading) return undefined;
    if (this.pendingSteps.length === 0 || this.current === 'closed' || this.current === 'detached') return this.travel(kind);
    this.wanted.push(kind);
    if (!this.wantedTimer) {
      this.wantedTimer = this.timer(() => {
        this.wantedTimer = undefined;
        if (this.wanted.length === 0) return;
        this.wanted = [];
        this.host.notice({ code: 'undo-wait' });
      }, UNDO_WAIT_MS);
    }
    return undefined;
  }

  protected idle(): void {
    if (this.wanted.length > 0) this.runWanted();
  }

  protected forget(all: boolean): void {
    this.wanted = [];
    if (!all) return;
    this.undoStack = [];
    this.redoStack = [];
  }

  /** The undoing that waited for the steps of this participant to be placed: done now, one at a time. */
  private runWanted(): void {
    if (this.wanted.length === 0 || this.pendingSteps.length > 0) return;
    const kind = this.wanted.shift() as 'undo' | 'redo';
    if (this.wanted.length === 0 && this.wantedTimer) {
      this.wantedTimer();
      this.wantedTimer = undefined;
    }
    const done = this.travel(kind);
    if (done) this.host.changed({ ...done.change, local: true });
    else this.runWanted();
  }

  private travel(kind: 'undo' | 'redo'): (Submitted & { meta: unknown }) | undefined {
    const stack = kind === 'undo' ? this.undoStack : this.redoStack;
    const entry = stack[stack.length - 1];
    // One at a time: the step that takes the one above back has no place yet.
    if (!entry || entry.busy) return undefined;
    if (!entry.recs || entry.gone || entry.recs.length === 0) {
      // The step stays behind: the one before it is next.
      this.abandon(stack, entry);
      this.host.notice({ code: entry.recs ? 'undo-gone' : 'undo-none' });
      return undefined;
    }
    const workbook = this.host.workbook();
    let done: ReturnType<CollabSession['applyRecs']>;
    try {
      done = this.applyRecs(workbook, entry.recs);
    } catch (error) {
      if (!(error instanceof RangeError)) throw error;
      this.abandon(stack, entry);
      this.host.notice({ code: 'undo-gone' });
      return undefined;
    }
    if (entry.lossy) this.host.notice({ code: 'undo-lossy' });
    entry.busy = true;
    this.placed({ id: '', recs: done.recs, inverse: done.inverse, change: done.change, kind, undo: done.undo, lossy: done.lossy, meta: entry.meta, from: entry });
    return { change: done.change, skipped: 0, inverse: done.inverse, meta: entry.meta };
  }

  /**
   * A step that will never be taken back leaves its stack. What lies under it was written for the
   * workbook without it: when it moved rows, columns or sheets, none of that is known to be right
   * any more, and it goes as well (which is said when it is asked for).
   */
  private abandon(stack: StackEntry[], entry: StackEntry): void {
    const at = stack.indexOf(entry);
    if (at < 0) return;
    stack.splice(at, 1);
    if (entry.shifty) for (let i = 0; i < at; i++) stack[i].gone = true;
  }

  /** A step that was applied here is queued for the relay; in a session that tells nobody it is a step like one the relay placed. */
  private placed(step: Step): void {
    const carry: Carry[] = [];
    for (const rec of step.recs) {
      const sh = shiftOf(rec.w, rec.nm, rec.wide);
      if (sh) carry.push({ sh });
    }
    if (carry.length > 0) this.movePresences(carry);
    if (this.paused > 0) this.touchedWhilePaused = true;
    if (this.current === 'closed' || this.current === 'detached') return this.settle(step);
    this.queue(step);
  }


  /**
   * A step of this participant has its place in the order (or, in a session that tells nobody, was
   * made): what undoes the steps before it is read through it, and what undoes it is kept. The
   * stacks are always about the workbook as the relay has it, without the steps that still wait.
   */
  protected settle(step: Step): void {
    // What it took back, or made again, is done with.
    if (step.from) {
      const source = step.kind === 'undo' ? this.undoStack : this.redoStack;
      const at = source.indexOf(step.from);
      if (at >= 0) source.splice(at, 1);
      step.from = undefined;
    }
    if (step.kind === 'undo' && step.final) return;
    const stack = step.kind === 'undo' ? this.redoStack : this.undoStack;
    const carry: Carry[] = [];
    for (const rec of step.recs) {
      const sh = shiftOf(rec.w, rec.nm, rec.wide);
      if (sh) carry.push({ sh });
    }
    let shifty = carry.length > 0;
    if (shifty && step.undo === null) {
      // A step that cannot be taken back stays for good: what lies under it is read through it, once.
      if (stack === this.undoStack) this.throughStack(stack, carry, false);
      else for (const entry of stack) entry.gone = true;
      shifty = false;
    }
    // Everything else on the stacks stays as it was written: a step is taken back only after every step
    // that was made after it, on the workbook it was written for.
    stack.push({ recs: step.undo, lossy: step.lossy, meta: step.meta, shifty });
    const limit = this.options.undoLimit === undefined ? 200 : this.options.undoLimit;
    while (stack.length > limit) stack.shift();
  }

  /**
   * Reads a stack through what was applied to the workbook the stack is about (`carry`, as it
   * is read on that workbook: the one the topmost entry was written for). Each entry is about the
   * workbook without the steps above it, so the carry reaches it read through what takes those
   * steps back. An entry that gives way is gone; when its step moved rows, columns or sheets,
   * so is everything under it.
   */
  private throughStack(stack: StackEntry[], carry: readonly Carry[], undo: boolean): void {
    // An undo is whole or it is nothing: what someone else changed since is not taken back in part. Nor is one
    // that would bring back, or take away, another run of rows or columns than it was made for.
    const shape = (recs: readonly Rec[]): string => recs.map((rec) => (rec.w.t === 'ins' || rec.w.t === 'del' ? `${rec.w.t}${rec.w.x}${rec.w.n}` : rec.w.t === 'cells' ? `c${rec.w.c.length}` : rec.w.t)).join();
    let over: readonly Carry[] = carry;
    for (let i = stack.length - 1; i >= 0; i--) {
      const entry = stack[i];
      if (entry.recs && !entry.gone) {
        const moved = transformStep(entry.recs, over, undo);
        if (moved.recs && !moved.partial && shape(moved.recs) === shape(entry.recs)) {
          entry.recs = moved.recs;
          over = moved.carry;
          continue;
        }
        entry.gone = true;
      }
      if (!entry.shifty) continue;
      for (let k = 0; k < i; k++) stack[k].gone = true;
      return;
    }
  }

  /** Applies wire operations one after the other; when one throws, the ones before it are taken back and the error is passed on. */
  private applyRecs(workbook: Workbook, recs: readonly Rec[]): { recs: Rec[]; inverse: Op[]; undo: Rec[] | null; lossy: boolean; change: SessionChange } {
    const out: Rec[] = [];
    const inverse: Op[] = [];
    const undo: Rec[][] = [];
    const change = emptyChange();
    let undoable = true;
    let lossy = false;
    try {
      for (const rec of recs) {
        const applied = applyWire(workbook, rec.w);
        const nm = applied.nm !== undefined ? applied.nm : rec.nm;
        const made: Rec = applied.wide ? { w: rec.w, nm, wide: true } : { w: rec.w, nm };
        if (rec.fold !== undefined) made.fold = rec.fold;
        if (rec.cut) made.cut = true;
        out.push(made);
        for (let i = applied.inverse.length - 1; i >= 0; i--) inverse.unshift(applied.inverse[i]);
        if (applied.undo) undo.unshift(applied.undo);
        else undoable = false;
        if (applied.lossy) lossy = true;
        addChange(change, applied);
        const shift = shiftOf(rec.w, nm, applied.wide);
        if (shift) change.shifts.push(shift);
      }
    } catch (error) {
      for (const op of inverse) applyOp(workbook, op);
      throw error;
    }
    return { recs: out, inverse, undo: undoable ? ([] as Rec[]).concat(...undo) : null, lossy, change };
  }

  protected encodeFirst(): unknown[] {
    return encodeStep(this.pendingSteps[0].recs.map((rec) => rec.w), this.host.workbook().styles);
  }

  protected sameAsFirst(ops: unknown[]): boolean {
    const step = this.pendingSteps[0];
    // As values, not as text: a relay may write the same operations with their members in another order.
    return stable(encodeStep(step.recs.map((rec) => rec.w), this.host.workbook().styles)) === stable(ops);
  }

  /** Takes every pending step back and forgets it. */
  protected dropAll(code: CollabNoticeCode): void {
    const workbook = this.host.workbook();
    const count = this.pendingSteps.length;
    const change = emptyChange();
    change.cells = undefined;
    change.layout = true;
    change.sheets = true;
    for (let i = this.pendingSteps.length - 1; i >= 0; i--) {
      const step = this.pendingSteps[i];
      for (const op of step.inverse) applyOp(workbook, op);
      // An undoing that was taken back took nothing back.
      if (step.from) step.from.busy = false;
    }
    this.pendingSteps = [];
    this.flying = undefined;
    if (count > 0) {
      this.host.notice({ code, count });
      this.host.changed(change);
    }
    this.checkSettled();
  }

  /** A step of someone else: placed under the pending ones. False when the session was closed over it. */
  protected remote(raw: unknown[]): boolean {
    const workbook = this.host.workbook();
    let wires: WireOp[];
    try {
      wires = decodeStep(raw, workbook.styles);
    } catch (error) {
      // No room for its formats here, which says nothing of the others: the step cannot be passed over (they
      // apply it), and cannot be applied. The workbook leaves the room as it is.
      if (error instanceof StyleTableFull) throw error;
      if (!(error instanceof WireError) && !(error instanceof RangeError)) throw error;
      // No step of a workbook. Everybody reads it the same way and passes it over; the room goes on.
      this.badStep();
      return true;
    }
    const pending = this.pendingSteps;
    for (let i = pending.length - 1; i >= 0; i--) for (const op of pending[i].inverse) applyOp(workbook, op);
    const change = emptyChange();
    // Everything that is done to the workbook from here on is what the view has to follow.
    const merge = (part: SessionChange): void => {
      if (part.layout) change.layout = true;
      if (part.sheets) change.sheets = true;
      if (part.lists) change.lists = true;
      if (!part.cells) change.cells = undefined;
      else if (change.cells) for (const cell of part.cells) change.cells.push(cell);
    };
    for (const step of pending) merge(step.change);
    let theirs: Rec[] = [];
    try {
      const done = this.applyRecs(workbook, wires.map((w) => ({ w })));
      theirs = done.recs;
      merge(done.change);
    } catch (error) {
      // A step this workbook does not take: nothing of it was applied, here or anywhere (everybody applies
      // it to the same workbook). It is passed over, and the pending steps go back on as they were.
      this.badStep();
    }

    let carry: Carry[] = theirs.map(carryOf);
    // What undoes the steps the relay has placed is about the workbook without the waiting steps: it is read
    // through what arrived, as it arrived, and gives way wherever that changed the same thing.
    this.throughStack(this.undoStack, carry, true);
    this.throughStack(this.redoStack, carry, true);
    this.movePresences(carry);
    const kept: Step[] = [];
    let dropped = 0;
    let partial = 0;
    let undone = 0;
    let rest = false;
    for (const step of pending) {
      // An undoing that waits gives way like any undoing: where someone else changed the same thing it is not applied at all.
      const undoing = step.kind !== 'edit';
      const moved = rest ? undefined : transformStep(step.recs, carry, undoing);
      let done: ReturnType<CollabSession['applyRecs']> | undefined;
      if (moved && moved.recs && !(undoing && moved.partial)) {
        try {
          done = this.applyRecs(workbook, moved.recs);
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
      if (moved && done) {
        step.recs = done.recs;
        step.inverse = done.inverse;
        step.undo = done.undo;
        // Rows or columns that came or went while rows or columns of the same sheet came or went beside them: what
        // taking them back would have to bring back is not known to come out right. Such a step is not one to undo.
        for (const rec of done.recs) {
          if (rec.w.t !== 'ins' && rec.w.t !== 'del') continue;
          const sheet = rec.w.s;
          if (carry.some((item) => item.sh !== undefined && item.sh.k === 'rc' && item.sh.s === sheet)) step.undo = null;
        }
        step.lossy = done.lossy;
        step.change = done.change;
        merge(done.change);
        if (moved.partial) partial++;
        carry = moved.carry;
        kept.push(step);
        continue;
      }
      if (moved && moved.noop) this.stats.quiet++;
      else {
        this.stats.dropped++;
        if (undoing) undone++;
        else dropped++;
      }
      if (this.flying === step.id) this.flying = undefined;
      // What it was to take back stays as it is, for good.
      if (step.from) this.abandon(step.kind === 'undo' ? this.undoStack : this.redoStack, step.from);
      // Where the steps that were made after it belong cannot be told any more: they go with it.
      if (moved && moved.poison) rest = true;
      if (rest) continue;
      // What was made after it was made on a workbook that had it.
      const back: Carry[] = [];
      for (let i = step.recs.length - 1; i >= 0; i--) {
        const sh = shiftOf(step.recs[i].w, step.recs[i].nm, step.recs[i].wide);
        if (sh) back.push({ sh: inverseShift(sh) });
      }
      carry = back.concat(carry);
    }
    this.pendingSteps = kept;
    for (const item of carry) if (item.sh) change.shifts.push(item.sh);
    if (dropped > 0) this.host.notice({ code: 'dropped', count: dropped });
    if (undone > 0) this.host.notice({ code: 'undo-gone', count: undone });
    if (partial > 0) this.host.notice({ code: 'partial', count: partial });
    this.host.changed(change);
    return true;
  }


  protected readPresence(value: unknown): CollabPresence | undefined {
    return readPresence(value);
  }

  protected writePresence(p: CollabPresence | undefined): unknown {
    return p ? { s: p.sheet, a: [p.row, p.col], r: [p.range.top, p.range.left, p.range.bottom, p.range.right] } : null;
  }

  /** The selections of the others follow rows, columns and sheets that moved, until they say where they are again. */
  private movePresences(carry: readonly Carry[]): void {
    let moved = false;
    this.others.forEach((peer) => {
      const presence = readPresence(peer.presence);
      if (!presence) return;
      let active: { s: number; row: number; col: number } | undefined = { s: presence.sheet, row: presence.row, col: presence.col };
      let first: { s: number; row: number; col: number } | undefined = { s: presence.sheet, row: presence.range.top, col: presence.range.left };
      let last: { s: number; row: number; col: number } | undefined = { s: presence.sheet, row: presence.range.bottom, col: presence.range.right };
      for (const item of carry) {
        if (!item.sh || item.sh.k === 'perm' || item.sh.k === 'sName') continue;
        if (active) active = mapPosition(active, item.sh);
        if (first) first = mapPosition(first, item.sh);
        if (last) last = mapPosition(last, item.sh);
      }
      const next = active && first && last && last.row >= first.row && last.col >= first.col ? { s: active.s, a: [active.row, active.col], r: [first.row, first.col, last.row, last.col] } : null;
      if (JSON.stringify(next) !== JSON.stringify(peer.presence)) {
        peer.presence = next;
        moved = true;
      }
    });
    if (moved) this.refreshPeers();
  }

  /** The workbook that is edited: what a file of the room is written from. */
  get model(): Workbook {
    return this.host.workbook();
  }
}
