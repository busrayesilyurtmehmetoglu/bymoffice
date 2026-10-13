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
import { MAX_PRESENCE_BYTES, MAX_STEP_BYTES, cleanUser, parseServerMessage, utf8Length, type ClientMessage, type CollabUser, type PeerMessage } from './protocol';
import { carryOf, inverseShift, mapPosition, transformStep, type Carry } from './rebase';
import { CollabUnsupportedError, WireError, type Rec, type Shift, type WireOp } from './types';
import { MAX_STEP_OPS, StyleTableFull, applyWire, decodeStep, encodeOp, encodeStep, shiftOf, stable, type Applied } from './wire';

/** What carries the messages: a socket, a channel of the application, anything that keeps their order. */
export interface CollabTransport {
  send(text: string): void;
  /** Called once, with what every message that arrives is given to. */
  onMessage(listener: (text: string) => void): void;
  /** Called once, when the transport tells whether it is connected; a transport without it is taken to be open, and to stay so. */
  onState?(listener: (state: 'open' | 'closed') => void): void;
  close(): void;
}

/**
 * `detached`: the room is at another version of the file than this participant's workbook, which
 * holds changes no file has; it stays as it is, tells nobody what is changed in it, and joins
 * again once it was saved (rejoin()) or given up for the file as it is stored (reload()).
 */
export type CollabState = 'connecting' | 'online' | 'offline' | 'detached' | 'closed';

/** The version of the contract this side speaks: a room takes participants of one version only. */
export const PROTOCOL_VERSION = 2;

/** Where a participant is: the sheet, the active cell and the selected range. */
export interface CollabPresence {
  sheet: number;
  row: number;
  col: number;
  range: { top: number; left: number; bottom: number; right: number };
}

export interface CollabPeer {
  id: string;
  /** The name the participant gave; undefined for one who gave none (shown as a guest with `guest` as the number). */
  name?: string;
  /** "#RRGGBB": the colour given, or one of the palette. */
  color: string;
  guest: number;
  readOnly: boolean;
  presence?: CollabPresence;
}

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

export type CollabNoticeCode =
  /** A change of this participant could not be applied after what someone else did, and was dropped. */
  | 'dropped'
  /** Part of a change went with rows, columns or a sheet someone else deleted. */
  | 'partial'
  /** What was to be undone was changed or removed by someone else since. */
  | 'undo-gone'
  /** The step cannot be undone while the workbook is edited together (a deleted sheet, a rule, a chart). */
  | 'undo-none'
  /** Undone, but not everything came back (notes, links or rules of deleted rows). */
  | 'undo-lossy'
  /** Rules or drawings that came with a paste or a cut were left out: they do not travel. */
  | 'skipped'
  /** Someone sent something that is no change of this workbook; the session was closed. */
  | 'peer'
  /** The file was changed elsewhere and could not be loaded again; the session was closed. */
  | 'stale'
  /** The file was loaded again; changes that were not sent yet were lost. */
  | 'reloaded'
  /** The relay refused this participant or closed the room. */
  | 'refused'
  /** Many changes wait to be sent. */
  | 'backlog'
  /** A step arrived that is no step of this workbook; everybody skipped it. */
  | 'bad-step'
  /** The relay takes no changes from this participant: the view is read from here on. */
  | 'read-only'
  /** The room is at another version of the file, and this workbook holds changes no file has: it stays, and is to be saved. */
  | 'detached'
  /** The room cannot be joined yet: someone in it has to save first. It is tried again. */
  | 'waiting'
  /** Undo was asked for while changes of this participant were still on their way, and they did not arrive in time. */
  | 'undo-wait'
  /** A step has been on its way for a long time: the connection is made anew. */
  | 'slow';

export interface CollabNotice {
  code: CollabNoticeCode;
  /** How many changes it is about, where that says something. */
  count?: number;
}

/** What the session needs of whoever holds the workbook. */
export interface SessionHost {
  /** The workbook as it is now (another one after reload()). */
  workbook(): Workbook;
  /** The workbook was changed by something the host did not call for: a step of someone else, a step of its own that was dropped. */
  changed(change: SessionChange): void;
  notice(notice: CollabNotice): void;
  /**
   * Loads the file again (the room is at a version this participant's file is not), after
   * which workbook() answers with it; resolves with the versions of the file that was loaded.
   * Without it the session closes when that happens.
   */
  reload?(): Promise<string[] | { versions: string[]; stamp?: RoomStamp }>;
  /** Whether the workbook holds changes that no file holds: such a workbook is never given up for the file without being asked. */
  unsaved?(): boolean;
}

/** The room, as its relay names it, and the step of its order a file holds: what a file that was saved in a room says of itself. */
export interface RoomStamp {
  ep: string;
  rev: number;
}

export interface SessionOptions {
  transport: CollabTransport;
  /** What the file that was read says of the room and the step it was saved at, when it says anything. */
  stamp?: RoomStamp;
  /** What the file this participant read is known by (an ETag, a hash of its bytes): see docs/collab-sozlesmesi.md. */
  versions: string[];
  user?: CollabUser;
  room?: string;
  token?: string;
  /** Only follows: sends no change. */
  readOnly?: boolean;
  onState?(state: CollabState): void;
  onPeers?(peers: CollabPeer[]): void;
  /** Someone saved the file: its versions, and whether this participant's workbook is what was saved. */
  onSaved?(saved: { versions: string[]; rev?: number; clean: boolean }): void;
  /** The relay made this participant one that only reads. */
  onReadOnly?(): void;
  /** Someone cannot join because no file holds the steps of the room: a participant that can save should. */
  onSync?(): void;
  /** Most steps kept for undo. Default 200. */
  undoLimit?: number;
  /** Runs a callback later and returns what calls it off. Default: setTimeout. */
  timer?(callback: () => void, ms: number): () => void;
}

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

/** Colours participants are told apart by when they give none. */
export const PEER_COLORS = ['#D93025', '#1A73E8', '#188038', '#E37400', '#9334E6', '#00838F', '#C2185B', '#5F6368'];

/** Most pending steps before it is said that many changes wait. */
const BACKLOG_NOTICE = 500;
/** How often a selection is sent at most. */
const PRESENCE_MS = 100;
/** How often in a row the file is loaded again before the session gives up. */
const MAX_RELOADS = 6;
/** How long what the others say of where they are is gathered before it is passed on. */
const PEERS_MS = 50;
/** How long an undo waits for the steps of its participant to be placed. */
const UNDO_WAIT_MS = 3000;
/** How long a step may be on its way before the connection is taken to be stuck. */
const FLY_MS = 30000;

function hashOf(text: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) h = Math.imul(h ^ text.charCodeAt(i), 0x01000193);
  return h >>> 0;
}

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

function cleanStamp(stamp: RoomStamp | undefined): RoomStamp | undefined {
  if (!stamp || typeof stamp.ep !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(stamp.ep)) return undefined;
  return Number.isSafeInteger(stamp.rev) && stamp.rev >= 0 ? { ep: stamp.ep, rev: stamp.rev } : undefined;
}

export class CollabSession {
  private readonly host: SessionHost;
  private readonly options: SessionOptions;
  private readonly transport: CollabTransport;
  /** What the steps of this session are told apart by, whatever the relay calls the connection. */
  private readonly key: string;
  private seq = 0;
  private current: CollabState = 'connecting';
  /** How many steps of the order the workbook holds under the pending ones. */
  private confirmed = 0;
  /** The order was joined at least once: joining again asks for what came since. */
  private joined = false;
  /** The order was asked for again after a step was missing, and the answer is not there yet. */
  private resyncing = false;
  private badStepSaid = false;
  /** What the steps of this session are named with, for everybody to see; the key is told to the relay alone. */
  private readonly tag = `${hashOf(String(Math.random())).toString(36)}${hashOf(String(Date.now())).toString(36)}`;
  /** Undo and redo that wait for the steps on their way. */
  private wanted: ('undo' | 'redo')[] = [];
  private wantedTimer: (() => void) | undefined;
  private flyTimer: (() => void) | undefined;
  private slowSaid = false;
  /** The step the workbook was at when it joined with its file. */
  private startedAt = 0;
  private lastSaved: ClientMessage | undefined;
  /** The relay counts this connection as one of a room: it takes no second "join" from it. */
  private inRoom = false;
  private waitingSaid = false;
  private peersTimer: (() => void) | undefined;
  private peersStale = false;
  /** The others moved, and whoever listens was not told yet. */
  private peersDue = false;
  /** Something was changed here while what arrives was held back (see pause()). */
  private touchedWhilePaused = false;
  /** The step the relay is at as far as this side knows: nothing is sent while the workbook is behind it. */
  private target = 0;
  private pendingSteps: Step[] = [];
  /** The room this side is in, as the relay names it; and what the file says it was saved at, until it was used or found wrong. */
  private epoch: string | undefined;
  private stamp: RoomStamp | undefined;
  /** The id of the step that is on its way. */
  private flying: string | undefined;
  private undoStack: StackEntry[] = [];
  private redoStack: StackEntry[] = [];
  private versions: string[];
  private me = '';
  private others: Map<string, PeerMessage> = new Map();
  private peerList: CollabPeer[] = [];
  private readonlyNow: boolean;
  private open = false;
  private reloads = 0;
  private reloading = false;
  private backlogSaid = false;
  private presenceNow: CollabPresence | undefined;
  private presenceSent = '';
  private presenceTimer: (() => void) | undefined;
  private settled: (() => void)[] = [];
  private paused = 0;
  private inbox: string[] = [];
  /** How many steps this participant queued, and what became of them: placed by the relay, dropped and said, or dropped because what they wanted was so already. */
  readonly stats = { made: 0, placed: 0, dropped: 0, quiet: 0 };

  constructor(host: SessionHost, options: SessionOptions) {
    this.host = host;
    this.options = options;
    this.stamp = cleanStamp(options.stamp);
    this.transport = options.transport;
    this.versions = options.versions.slice(0, 8);
    this.readonlyNow = options.readOnly === true;
    this.key = `${hashOf(String(Math.random())).toString(36)}${hashOf(String(Date.now() + Math.random())).toString(36)}`;
    this.transport.onMessage((text) => this.receive(text));
    if (this.transport.onState) {
      this.transport.onState((state) => {
        if (this.current === 'closed') return;
        if (state === 'open') {
          this.open = true;
          this.inRoom = false;
          if (this.current !== 'detached') this.join();
        } else {
          this.open = false;
          this.inRoom = false;
          this.flying = undefined;
          // A workbook that is detached stays so, whatever the connection does.
          if (this.current !== 'detached') this.setState('offline');
        }
      });
    } else {
      this.open = true;
      this.join();
    }
  }

  get state(): CollabState {
    return this.current;
  }

  /** How many steps of the common order the workbook holds. */
  get rev(): number {
    return this.confirmed;
  }

  /** How many steps of this participant the relay has not placed yet. */
  get pending(): number {
    return this.pendingSteps.length;
  }

  get peers(): CollabPeer[] {
    if (this.peersStale) this.buildPeers();
    return this.peerList.slice();
  }

  /** This participant as the others see it; `id` is empty until the relay named the connection. */
  get user(): CollabPeer {
    return this.lookOf(this.me, cleanUser(this.options.user), this.readonlyNow);
  }

  get readOnly(): boolean {
    return this.readonlyNow;
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
    step.id = `${this.tag}:${++this.seq}`;
    this.pendingSteps.push(step);
    this.stats.made++;
    if (this.pendingSteps.length > BACKLOG_NOTICE && !this.backlogSaid) {
      this.backlogSaid = true;
      this.host.notice({ code: 'backlog', count: this.pendingSteps.length });
    }
    this.pump();
  }

  /**
   * A step of this participant has its place in the order (or, in a session that tells nobody, was
   * made): what undoes the steps before it is read through it, and what undoes it is kept. The
   * stacks are always about the workbook as the relay has it, without the steps that still wait.
   */
  private settle(step: Step): void {
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

  // ---------------------------------------------------------------- the relay

  private send(message: ClientMessage): void {
    if (!this.open || this.current === 'closed') return;
    try {
      this.transport.send(JSON.stringify(message));
    } catch {
      // A transport that cannot send says so by its state; what was not sent is sent when it is open again.
    }
  }

  /** Makes a new connection where the transport can make one; false where it cannot. */
  private reconnect(): boolean {
    const transport = this.transport as CollabTransport & { suspend?(): void; resume?(): void };
    if (typeof transport.suspend !== 'function' || typeof transport.resume !== 'function') return false;
    transport.suspend();
    transport.resume();
    return true;
  }

  private join(): void {
    if (this.reloading) return;
    // A connection that is in a room joins no second time: the room is joined over a new one, or, where the
    // transport cannot make one (a transport of the application), left first.
    if (this.inRoom) {
      if (this.reconnect()) return;
      this.inRoom = false;
      this.send({ t: 'leave' });
    }
    const message: ClientMessage = { t: 'join', v: this.versions, pv: PROTOCOL_VERSION, key: this.key };
    if (this.joined && this.epoch !== undefined) {
      // Back in the room this side was in: from the step it is at. In another room that step means nothing (the relay sees to it).
      message.since = this.confirmed;
      message.ep = this.epoch;
    } else if (!this.joined && this.stamp) {
      // With a file that was saved in a room: from the step that was written into it.
      message.since = this.stamp.rev;
      message.ep = this.stamp.ep;
    }
    const user = cleanUser(this.options.user);
    if (user.name !== undefined || user.color !== undefined) message.user = user;
    if (typeof this.options.token === 'string' && this.options.token !== '') message.token = this.options.token;
    if (this.options.readOnly === true) message.ro = true;
    if (typeof this.options.room === 'string' && this.options.room !== '') message.room = this.options.room;
    this.send(message);
  }

  private pump(): void {
    if (this.current !== 'online' || this.flying !== undefined || this.pendingSteps.length === 0 || this.confirmed < this.target || this.readonlyNow) {
      this.checkSettled();
      return;
    }
    const step = this.pendingSteps[0];
    this.flying = step.id;
    this.send({ t: 'op', base: this.confirmed, id: step.id, ops: encodeStep(step.recs.map((rec) => rec.w), this.host.workbook().styles) });
    // A step that is neither placed nor refused for a long time: the connection is not what it seems.
    if (this.flyTimer) this.flyTimer();
    const id = step.id;
    this.flyTimer = this.timer(() => {
      this.flyTimer = undefined;
      if (this.flying !== id || this.current !== 'online') return;
      if (!this.slowSaid) {
        this.slowSaid = true;
        this.host.notice({ code: 'slow' });
      }
      this.flying = undefined;
      if (!this.reconnect()) this.pump();
    }, FLY_MS);
  }

  private checkSettled(): void {
    if (this.pendingSteps.length === 0 && this.wanted.length > 0 && this.current === 'online') this.runWanted();
    if (this.pendingSteps.length > 0 || this.settled.length === 0) return;
    const waiting = this.settled;
    this.settled = [];
    for (const done of waiting) done();
  }

  /** Resolves once no step of this participant waits for the relay: true; or after `ms`: false. */
  whenSettled(ms: number): Promise<boolean> {
    if (this.pendingSteps.length === 0) return Promise.resolve(true);
    return new Promise<boolean>((resolve) => {
      let over = false;
      const stop = this.timer(() => {
        if (over) return;
        over = true;
        resolve(false);
      }, ms);
      this.settled.push(() => {
        if (over) return;
        over = true;
        stop();
        resolve(true);
      });
    });
  }

  private timer(callback: () => void, ms: number): () => void {
    if (this.options.timer) return this.options.timer(callback, ms);
    const handle = setTimeout(callback, ms);
    return () => clearTimeout(handle);
  }

  private setState(state: CollabState): void {
    if (this.current === state) return;
    this.current = state;
    if (this.options.onState) this.options.onState(state);
  }

  private fatal(code: CollabNoticeCode): void {
    if (this.current === 'closed') return;
    this.host.notice({ code });
    this.close();
  }

  private close(): void {
    if (this.current === 'closed') return;
    this.setState('closed');
    this.open = false;
    this.flying = undefined;
    if (this.presenceTimer) this.presenceTimer();
    this.presenceTimer = undefined;
    // What waits is this participant's alone from here on: steps like any that were placed.
    for (const step of this.pendingSteps) this.settle(step);
    this.pendingSteps = [];
    this.others.clear();
    this.refreshPeers();
    this.checkSettled();
    try {
      this.transport.close();
    } catch {
      // Nothing more to do with a transport that cannot even be closed.
    }
  }

  /** Leaves the room. The workbook stays as it is; what is changed from here on is told to nobody. */
  leave(): void {
    this.close();
  }

  /**
   * Holds what arrives until resume(): while a file is written from the workbook, nothing that
   * arrives may change it half way through.
   */
  pause(): void {
    if (this.paused++ === 0) this.touchedWhilePaused = false;
  }

  /** Lets what was held back in. Answers whether the workbook was changed here meanwhile: a file written in that time is of no one step. */
  resume(): boolean {
    if (this.paused === 0) return false;
    const touched = this.touchedWhilePaused;
    if (--this.paused > 0) return touched;
    const held = this.inbox;
    this.inbox = [];
    for (const text of held) this.receive(text);
    return touched;
  }

  private receive(text: string): void {
    if (this.current === 'closed' || this.current === 'detached') return;
    if (this.paused > 0) {
      this.inbox.push(text);
      return;
    }
    let message;
    try {
      message = parseServerMessage(text);
    } catch {
      return this.fatal('peer');
    }
    try {
      this.handle(message);
    } catch (error) {
      // Nothing that arrives may throw out of here. The workbook is kept as it is, apart from the room.
      const now = this.current as CollabState;
      if (now !== 'closed' && now !== 'detached') this.detach();
    }
  }

  private handle(message: ReturnType<typeof parseServerMessage>): void {
    switch (message.t) {
      case 'hello': {
        // The log from before where this side is brings steps it has, which are passed over; one that starts
        // behind it leaves steps out.
        this.inRoom = true;
        // A room that is behind this side is another room (the relay was started anew, or the room was let go of).
        if (this.joined && (message.ep !== this.epoch || message.from > this.confirmed || message.rev < this.confirmed)) return this.restart();
        if (!this.joined) {
          const stamp = this.stamp;
          if (stamp && message.ep === stamp.ep) {
            // The step the file names, in the room it names. The room has to have heard of this very file as well: one
            // that was changed elsewhere since carries the stamp still, and is another file.
            if (message.from !== stamp.rev || message.v.length === 0) {
              this.stamp = undefined;
              return this.join();
            }
          } else if (message.from !== 0) {
            // Without the word of the file there is only the start of the room to join at.
            return this.fatal('peer');
          }
          this.epoch = message.ep;
          this.confirmed = message.from;
          this.startedAt = message.from;
        }
        this.joined = true;
        this.resyncing = false;
        this.reloads = 0;
        this.waitingSaid = false;
        this.target = message.rev;
        this.me = message.you;
        if (message.ro && !this.readonlyNow) {
          this.readonlyNow = true;
          // Asked for by nobody here: whoever edits is told, and what waits is taken back.
          if (this.options.readOnly !== true) {
            this.dropAll('refused');
            this.host.notice({ code: 'read-only' });
            if (this.options.onReadOnly) this.options.onReadOnly();
          }
        }
        this.others.clear();
        for (const peer of message.peers) this.others.set(peer.id, peer);
        this.refreshPeers();
        this.setState('online');
        // The selection is told anew: the relay does not remember it across connections.
        this.presenceSent = '';
        this.flushPresence();
        this.pump();
        return;
      }
      case 'op':
        if (!this.joined) return;
        if (message.rev <= this.confirmed) return;
        // A step is missing: the order is asked for again from where it is known.
        if (message.rev !== this.confirmed + 1) return this.resync();
        // This participant's own step, as the relay placed it: by its name and by what it says. A step that
        // only has the name (anybody can read the names) is someone else's.
        if (this.pendingSteps.length > 0 && this.pendingSteps[0].id === message.id && this.sameAsFirst(message.ops)) {
          const mine = this.pendingSteps.shift() as Step;
          this.stats.placed++;
          this.confirmed = message.rev;
          this.flying = undefined;
          this.slowSaid = false;
          this.settle(mine);
          if (this.pendingSteps.length < BACKLOG_NOTICE) this.backlogSaid = false;
        } else if (!this.remote(message.ops)) return;
        else this.confirmed = message.rev;
        if (this.target < this.confirmed) this.target = this.confirmed;
        this.pump();
        return;
      case 'nack':
        if (this.flying !== message.id) return;
        this.flying = undefined;
        if (message.reason === 'read-only') {
          this.readonlyNow = true;
          this.dropAll('refused');
          this.host.notice({ code: 'read-only' });
          if (this.options.onReadOnly) this.options.onReadOnly();
          return;
        }
        if (message.reason === 'rate') {
          // Too many steps in a short time: the step waits a moment.
          this.timer(() => this.pump(), 1000);
          return;
        }
        if (message.rev > this.target) this.target = message.rev;
        this.pump();
        return;
      case 'reset':
        this.inRoom = false;
        return this.restart();
      case 'presence': {
        const peer = this.others.get(message.from);
        if (!peer) return;
        peer.presence = message.data;
        // However often the others say where they are, it is passed on a few times a second.
        this.peersStale = true;
        this.peersDue = true;
        if (!this.peersTimer) {
          this.peersTimer = this.timer(() => {
            this.peersTimer = undefined;
            if (this.peersDue) this.refreshPeers();
          }, PEERS_MS);
        }
        return;
      }
      case 'peer':
        this.others.set(message.id, { id: message.id, user: message.user, ro: message.ro });
        this.refreshPeers();
        // Who comes in has not heard what this side last saved: the file it read may be an older one than the one that is stored.
        if (this.lastSaved) this.send(this.lastSaved);
        return;
      case 'left':
        if (this.others.delete(message.id)) this.refreshPeers();
        return;
      case 'saved': {
        const clean = message.rev !== undefined && message.rev === this.confirmed && this.pendingSteps.length === 0;
        if (this.options.onSaved) this.options.onSaved({ versions: message.v, rev: message.rev, clean });
        return;
      }
      case 'sync':
        if (this.options.onSync && !this.readonlyNow) this.options.onSync();
        return;
      case 'pong':
        return;
      case 'error':
        // This side's own new connection took the place of this one: nothing to say.
        if (message.code === 'replaced') return;
        return this.fatal('refused');
    }
  }

  /**
   * A step is missing (a transport that does not keep the order): the order is asked for again
   * from the last step that is known, once, over a new connection where the transport can make
   * one. The file is not loaded again for this.
   */
  private resync(): void {
    if (this.resyncing) return;
    this.resyncing = true;
    this.flying = undefined;
    if (!this.reconnect()) this.join();
  }

  private sameAsFirst(ops: unknown[]): boolean {
    const step = this.pendingSteps[0];
    // As values, not as text: a relay may write the same operations with their members in another order.
    return stable(encodeStep(step.recs.map((rec) => rec.w), this.host.workbook().styles)) === stable(ops);
  }

  /** Takes every pending step back and forgets it. */
  private dropAll(code: CollabNoticeCode): void {
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

  private badStep(): void {
    if (this.badStepSaid) return;
    this.badStepSaid = true;
    this.host.notice({ code: 'bad-step' });
  }

  /**
   * The room is at a version of the file this participant does not have. A workbook that holds
   * nothing but its file is given up for the file as it is stored now, and the room is joined
   * anew. One that holds changes no file has is never given up without being asked: it stays as
   * it is, detached from the room (see CollabState).
   */
  private restart(): void {
    if (this.reloading) return;
    let ahead = this.pendingSteps.length > 0;
    try {
      // A host that does not say whether its workbook is saved is taken to have saved nothing since it joined.
      if (!ahead) ahead = this.host.unsaved ? this.host.unsaved() === true : this.confirmed > this.startedAt;
    } catch {
      ahead = true;
    }
    if (ahead) return this.detach();
    this.loadAgain();
  }

  private detach(): void {
    // What waited is in the workbook, and stays there; nobody will be told of it.
    for (const step of this.pendingSteps) this.settle(step);
    this.pendingSteps = [];
    this.wanted = [];
    this.lastSaved = undefined;
    this.flying = undefined;
    this.joined = false;
    this.epoch = undefined;
    this.stamp = undefined;
    this.resyncing = false;
    this.confirmed = 0;
    this.target = 0;
    this.others.clear();
    this.refreshPeers();
    this.setState('detached');
    // Out of the room, where the transport can leave it: nothing of it is of any use here.
    const transport = this.transport as CollabTransport & { suspend?(): void };
    if (typeof transport.suspend === 'function') transport.suspend();
    this.host.notice({ code: 'detached' });
    this.checkSettled();
  }

  /**
   * Joins the room again with the workbook as it is: after it was saved, under what the saved file is
   * known by. For a detached session; does nothing in any other state.
   */
  rejoin(versions: string[]): void {
    if (this.current !== 'detached') return;
    this.stamp = undefined;
    this.epoch = undefined;
    this.versions = versions.filter((version) => typeof version === 'string' && version !== '' && version.length <= 200).slice(0, 8);
    this.reloads = 0;
    this.setState('connecting');
    const transport = this.transport as CollabTransport & { resume?(): void };
    if (this.open) this.join();
    else if (typeof transport.resume === 'function') transport.resume();
    else this.setState('offline');
  }

  /**
   * Gives the workbook up for the file as it is stored and joins the room with that: what a detached
   * session does when its user says so. Changes that are in no file are lost.
   */
  reload(): void {
    if (this.current === 'closed' || this.reloading) return;
    this.reloads = 0;
    const detached = this.current === 'detached';
    this.loadAgain();
    const transport = this.transport as CollabTransport & { resume?(): void };
    if (detached && !this.open && typeof transport.resume === 'function') transport.resume();
  }

  private loadAgain(): void {
    const reload = this.host.reload;
    if (!reload || ++this.reloads > MAX_RELOADS) return this.fatal('stale');
    const lost = this.pendingSteps.length;
    this.reloading = true;
    this.wanted = [];
    this.lastSaved = undefined;
    this.pendingSteps = [];
    this.flying = undefined;
    this.undoStack = [];
    this.redoStack = [];
    this.joined = false;
    this.epoch = undefined;
    this.stamp = undefined;
    this.resyncing = false;
    this.confirmed = 0;
    this.target = 0;
    this.setState('connecting');
    this.checkSettled();
    // The file that is stored may be one the room cannot place yet (it was saved, and the room was not told
    // which step it holds): the relay asks those in the room to save, and this side tries again a little later.
    const again = this.reloads > 1;
    if (again && !this.waitingSaid) {
      this.waitingSaid = true;
      this.host.notice({ code: 'waiting' });
    }
    const load = (): void => {
      if (this.current === 'closed') return;
      let loading: Promise<string[] | { versions: string[]; stamp?: RoomStamp }>;
      try {
        loading = Promise.resolve(reload.call(this.host));
      } catch {
        this.reloading = false;
        return this.fatal('stale');
      }
      loading.then(
        (loaded) => {
          this.reloading = false;
          if (this.current === 'closed') return;
          const versions = Array.isArray(loaded) ? loaded : loaded.versions;
          this.stamp = Array.isArray(loaded) ? undefined : cleanStamp(loaded.stamp);
          this.versions = versions.slice(0, 8);
          if (lost > 0) this.host.notice({ code: 'reloaded', count: lost });
          if (this.open) this.join();
          else this.setState('offline');
        },
        () => {
          this.reloading = false;
          this.fatal('stale');
        },
      );
    };
    if (again) this.timer(load, Math.min(500 * Math.pow(2, this.reloads - 2), 8000));
    else load();
  }

  /** A step of someone else: placed under the pending ones. False when the session was closed over it. */
  private remote(raw: unknown[]): boolean {
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

  // ----------------------------------------------------------------- presence

  private lookOf(id: string, user: CollabUser, readOnly: boolean, presence?: CollabPresence): CollabPeer {
    const hash = hashOf(id);
    const peer: CollabPeer = { id, color: user.color !== undefined ? user.color : PEER_COLORS[hash % PEER_COLORS.length], guest: (hash % 99) + 1, readOnly };
    if (user.name !== undefined) peer.name = user.name;
    if (presence) peer.presence = presence;
    return peer;
  }

  private buildPeers(): void {
    this.peersStale = false;
    const list: CollabPeer[] = [];
    this.others.forEach((peer) => list.push(this.lookOf(peer.id, peer.user, peer.ro, readPresence(peer.presence))));
    this.peerList = list;
  }

  /** Whether this participant is the one that saves by itself what the others changed: the one of those who write whose name of the relay comes first. */
  get leads(): boolean {
    if (this.me === '' || this.readonlyNow || this.current !== 'online') return false;
    let first = true;
    this.others.forEach((peer) => {
      if (!peer.ro && peer.id < this.me) first = false;
    });
    return first;
  }

  private refreshPeers(): void {
    this.peersStale = false;
    this.peersDue = false;
    const list: CollabPeer[] = [];
    this.others.forEach((peer) => list.push(this.lookOf(peer.id, peer.user, peer.ro, readPresence(peer.presence))));
    this.peerList = list;
    if (this.options.onPeers) this.options.onPeers(list.slice());
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

  /** Tells the others where this participant is; at most ten times a second. */
  setPresence(presence: CollabPresence | undefined): void {
    this.presenceNow = presence;
    if (this.presenceTimer || this.current === 'closed') return;
    this.presenceTimer = this.timer(() => {
      this.presenceTimer = undefined;
      this.flushPresence();
    }, PRESENCE_MS);
  }

  private flushPresence(): void {
    const p = this.presenceNow;
    const data = p ? { s: p.sheet, a: [p.row, p.col], r: [p.range.top, p.range.left, p.range.bottom, p.range.right] } : null;
    const text = JSON.stringify(data);
    if (text === this.presenceSent || this.current !== 'online' || text.length > MAX_PRESENCE_BYTES) return;
    this.presenceSent = text;
    this.send({ t: 'presence', data });
  }

  // ------------------------------------------------------------------- saving

  /**
   * The step of the common order the workbook is at when nothing of this participant waits:
   * a file written now is the file of that step. Undefined while something waits.
   */
  /**
   * What a file written now is to say of itself: the room, and the step of its order the workbook is at.
   * Undefined while something of this participant waits, or the workbook is in no room.
   */
  get stampNow(): RoomStamp | undefined {
    return this.current === 'online' && this.pendingSteps.length === 0 && this.epoch !== undefined ? { ep: this.epoch, rev: this.confirmed } : undefined;
  }

  get savableRev(): number | undefined {
    return this.current === 'online' && this.pendingSteps.length === 0 ? this.confirmed : undefined;
  }

  /** Tells the others that the file was saved: what it is known by now and, when known, the step it holds. */
  announceSaved(versions: string[], rev: number | undefined): void {
    const v = versions.filter((version) => typeof version === 'string' && version !== '' && version.length <= 200).slice(0, 8);
    if (v.length === 0 || this.current !== 'online') return;
    const message: ClientMessage = { t: 'saved', v };
    if (rev !== undefined) message.rev = rev;
    this.lastSaved = message;
    this.send(message);
  }
}
