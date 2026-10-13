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
import { MAX_STEP_OPS, applyWire, decodeStep, encodeOp, encodeStep, shiftOf, type Applied } from './wire';

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
export const PROTOCOL_VERSION = 1;

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
  | 'waiting';

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
  reload?(): Promise<string[]>;
  /** Whether the workbook holds changes that no file holds: such a workbook is never given up for the file without being asked. */
  unsaved?(): boolean;
}

export interface SessionOptions {
  transport: CollabTransport;
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
}

interface Step {
  id: string;
  recs: Rec[];
  /** The operations of the model that take the step back exactly, in the order they are applied in. */
  inverse: Op[];
  /** What undoes or redoes the step, on whichever stack it stands. */
  entry?: StackEntry;
  /** What applying it changed. */
  change: SessionChange;
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
  /** The step the workbook was at when it joined with its file. */
  private startedAt = 0;
  private lastSaved: ClientMessage | undefined;
  private keyTries = 0;
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
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  // ------------------------------------------------------------------ editing

  /**
   * Applies operations of the model as one step of this participant and queues it for the
   * others. Throws what the workbook throws for an operation it does not take, and
   * CollabUnsupportedError for what cannot be shared; nothing was changed then.
   */
  submit(ops: readonly Op[], meta?: unknown): Submitted {
    if (this.readonlyNow) throw new CollabUnsupportedError('readOnly', 'The relay takes no changes from this participant');
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
    const entry: StackEntry = { recs: undoable ? ([] as Rec[]).concat(...undo) : null, lossy, meta };
    this.placed(recs, inverse, entry, 'edit', change);
    return { change, skipped, inverse };
  }

  /** Undoes the last step of this participant that can be undone; undefined when nothing was done (which was said). */
  undo(): (Submitted & { meta: unknown }) | undefined {
    return this.travel(this.undoStack, 'undo');
  }

  redo(): (Submitted & { meta: unknown }) | undefined {
    return this.travel(this.redoStack, 'redo');
  }

  /** What the step that undo() would take back was given as `meta`. */
  get undoMeta(): unknown {
    const top = this.undoStack[this.undoStack.length - 1];
    return top ? top.meta : undefined;
  }

  private travel(stack: StackEntry[], kind: 'undo' | 'redo'): (Submitted & { meta: unknown }) | undefined {
    const entry = stack.pop();
    if (!entry) return undefined;
    if (!entry.recs || entry.gone || entry.recs.length === 0) {
      // The step stays behind: the one before it is next.
      this.host.notice({ code: entry.recs ? 'undo-gone' : 'undo-none' });
      return undefined;
    }
    const workbook = this.host.workbook();
    let done: ReturnType<CollabSession['applyRecs']>;
    try {
      done = this.applyRecs(workbook, entry.recs);
    } catch (error) {
      if (!(error instanceof RangeError)) throw error;
      this.host.notice({ code: 'undo-gone' });
      return undefined;
    }
    if (entry.lossy) this.host.notice({ code: 'undo-lossy' });
    const back: StackEntry = { recs: done.undo, lossy: done.lossy, meta: entry.meta };
    this.placed(done.recs, done.inverse, back, kind, done.change);
    return { change: done.change, skipped: 0, inverse: done.inverse, meta: entry.meta };
  }

  /** A step that was applied here: the stacks are read through it, its own entry is put on its stack, and it is queued. */
  private placed(recs: Rec[], inverse: Op[], entry: StackEntry, kind: 'edit' | 'undo' | 'redo', change: SessionChange): void {
    const carry: Carry[] = [];
    for (const rec of recs) {
      const sh = shiftOf(rec.w, rec.nm, rec.wide);
      if (sh) carry.push({ sh });
    }
    if (carry.length > 0) {
      this.throughStack(this.undoStack, carry, false);
      this.throughStack(this.redoStack, carry, false);
      this.movePresences(carry);
    }
    if (kind === 'edit') this.redoStack = [];
    const stack = kind === 'undo' ? this.redoStack : this.undoStack;
    stack.push(entry);
    const limit = this.options.undoLimit === undefined ? 200 : this.options.undoLimit;
    while (stack.length > limit) stack.shift();
    // A session that is over applies what it is given and tells nobody.
    if (this.paused > 0) this.touchedWhilePaused = true;
    if (this.current === 'closed' || this.current === 'detached' || this.readonlyNow) return;
    this.pendingSteps.push({ id: `${this.key}:${++this.seq}`, recs, inverse, entry, change });
    this.stats.made++;
    if (this.pendingSteps.length > BACKLOG_NOTICE && !this.backlogSaid) {
      this.backlogSaid = true;
      this.host.notice({ code: 'backlog', count: this.pendingSteps.length });
    }
    this.pump();
  }

  private throughStack(stack: StackEntry[], carry: readonly Carry[], undo: boolean, skip?: Set<StackEntry>): void {
    for (const entry of stack) {
      if (!entry.recs || entry.gone || (skip && skip.has(entry))) continue;
      const moved = transformStep(entry.recs, carry, undo);
      if (moved.recs) entry.recs = moved.recs;
      else entry.gone = true;
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
        out.push(applied.wide ? { w: rec.w, nm, wide: true } : { w: rec.w, nm });
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
    // A connection that is in a room joins no second time: the room is joined over a new one.
    if (this.inRoom && this.reconnect()) return;
    const message: ClientMessage = { t: 'join', v: this.versions, pv: PROTOCOL_VERSION, key: this.key };
    if (this.joined) message.since = this.confirmed;
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
  }

  private checkSettled(): void {
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
    // What waits is this participant's alone from here on.
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
    switch (message.t) {
      case 'hello': {
        // The log from before where this side is brings steps it has, which are passed over; one that starts
        // behind it leaves steps out.
        this.inRoom = true;
        // A room that is behind this side is another room (the relay was started anew, or the room was let go of).
        if (this.joined && (message.from > this.confirmed || message.rev < this.confirmed)) return this.restart();
        if (!this.joined) {
          this.confirmed = message.from;
          this.startedAt = message.from;
        }
        this.joined = true;
        this.resyncing = false;
        this.keyTries = 0;
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
          this.pendingSteps.shift();
          this.stats.placed++;
          this.confirmed = message.rev;
          this.flying = undefined;
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
        // The relay does not take this participant's names for steps: nothing it makes will ever be placed.
        if (message.reason === 'id') return this.fatal('refused');
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
        // The relay still holds the connection this one takes the place of (it has not noticed yet that it is gone): tried again a little later.
        if (message.code === 'key-taken' && this.keyTries++ < 5) {
          this.timer(() => {
            if (this.current !== 'closed' && this.current !== 'detached') this.reconnect();
          }, 1500);
          return;
        }
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
    return JSON.stringify(encodeStep(step.recs.map((rec) => rec.w), this.host.workbook().styles)) === JSON.stringify(ops);
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
      this.forget(step.entry);
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

  private forget(entry: StackEntry | undefined): void {
    if (!entry) return;
    for (const stack of [this.undoStack, this.redoStack]) {
      const at = stack.indexOf(entry);
      if (at >= 0) stack.splice(at, 1);
    }
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
    this.pendingSteps = [];
    this.flying = undefined;
    this.joined = false;
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
    this.pendingSteps = [];
    this.flying = undefined;
    this.undoStack = [];
    this.redoStack = [];
    this.joined = false;
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
      let loading: Promise<string[]>;
      try {
        loading = Promise.resolve(reload.call(this.host));
      } catch {
        this.reloading = false;
        return this.fatal('stale');
      }
      loading.then(
        (versions) => {
          this.reloading = false;
          if (this.current === 'closed') return;
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
    const kept: Step[] = [];
    const fresh = new Set<StackEntry>();
    let dropped = 0;
    let partial = 0;
    for (const step of pending) {
      const moved = transformStep(step.recs, carry);
      let done: ReturnType<CollabSession['applyRecs']> | undefined;
      if (moved.recs) {
        try {
          done = this.applyRecs(workbook, moved.recs);
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
      if (done) {
        step.recs = done.recs;
        step.inverse = done.inverse;
        if (step.entry) {
          step.entry.recs = done.undo;
          step.entry.lossy = done.lossy;
          fresh.add(step.entry);
        }
        step.change = done.change;
        merge(done.change);
        if (moved.partial) partial++;
        carry = moved.carry;
        kept.push(step);
      } else {
        if (!moved.noop) dropped++;
        else this.stats.quiet++;
        this.forget(step.entry);
        if (this.flying === step.id) this.flying = undefined;
        // What was made after it was made on a workbook that had it.
        const back: Carry[] = [];
        for (let i = step.recs.length - 1; i >= 0; i--) {
          const sh = shiftOf(step.recs[i].w, step.recs[i].nm, step.recs[i].wide);
          if (sh) back.push({ sh: inverseShift(sh) });
        }
        carry = back.concat(carry);
      }
    }
    this.pendingSteps = kept;
    this.throughStack(this.undoStack, carry, true, fresh);
    this.throughStack(this.redoStack, carry, true, fresh);
    this.movePresences(carry);
    for (const item of carry) if (item.sh) change.shifts.push(item.sh);
    this.stats.dropped += dropped;
    if (dropped > 0) this.host.notice({ code: 'dropped', count: dropped });
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
