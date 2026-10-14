// One participant's side of a document that is edited together.
//
// The connection to the room is the one of a workbook (../room.ts). What is different is what
// a step is: the commands that were run (./wire.ts), which every participant runs on its own
// document. A step of someone else is placed under the steps of this participant that wait:
// they are taken back, the step is applied, and they are applied again, each read through
// what came before it (./rebase.ts). A waiting step that does not go together with what
// arrived is dropped, and so is every step that waits behind it (they were made on a document
// that had it); that is said.
//
// Undo is a step like any other. For each step it makes, this participant keeps the commands
// that take it back: "delete what was typed", or "put these paragraphs where those stand"
// with the paragraphs as they were (see inverseOf in wire.ts). To undo is to run those
// commands as a new step, which travels like any step: whoever receives it needs nothing but
// its document. The commands that are kept are read through every step of someone else that
// gets its place after them (never through this participant's own later steps: those are
// undone first, last in, first out); where someone else changed the same paragraphs since,
// the step cannot be undone any more, and that is said when it is asked for.

import type { DocModel } from '../../model/doc/types';
import { applyDocOp, type DocOp } from '../../ops/doc/ops';
import { listRevisions } from '../../ops/doc/revisions';
import { isTracking } from '../../ops/doc/track';
import { MAX_STEP_BYTES, utf8Length } from '../protocol';
import { RoomSession, type CollabNoticeCode, type RoomHost, type RoomOptions, type RoomPeer } from '../room';
import { CollabUnsupportedError, WireError } from '../types';
import { stable } from '../wire';
import { arrived, mapPos, orderOf, rebaseWire, type DocWire, type Mark, type WirePos } from './rebase';
import { MAX_STEP_WIRES, decodeStep, encodeStep, inverseOf, lookBefore, marksAfter, runWire } from './wire';

/** Where a participant is in a document: the story, and the two ends of its selection (the same place for a caret). */
export interface DocPresence {
  story: string;
  anchor: { path: number[]; offset: number };
  focus: { path: number[]; offset: number };
}

export type DocPeer = RoomPeer<DocPresence>;

/** What a change of the document that the host did not make itself was: for bringing a view up to date. */
export interface DocChange {
  /** What moved, in order: the places the host keeps (a caret, a selection) are read through these. */
  marks: Mark[];
  /** Steps of this participant were taken out of the document: a place the host keeps may be no place any more. */
  reset?: boolean;
  /** The change is this participant's own (an undo that waited for its turn), not someone else's; `meta` is what was given with the step it undid. */
  local?: boolean;
  meta?: unknown;
}

export interface DocSessionHost extends RoomHost {
  /** The document as it is now (another one after reload()). */
  document(): DocModel;
  /** The document was changed by something the host did not call for: a step of someone else, a step of its own that was dropped. */
  changed(change: DocChange): void;
}

export type DocSessionOptions = RoomOptions<DocPresence>;

/** What takes one step back: part of what Undo takes back at once. */
interface Part {
  /** The commands, in the order they are run in; null for a step that cannot be taken back by commands. */
  wires: DocWire[] | null;
  /** The step moved text or blocks: what lies under it is about the document without it. */
  moves: boolean;
  /** The step, while it waits for its place: its commands are made anew whenever it is applied again. */
  step?: Step;
}

/** What Undo (or Redo) takes back at once: the steps of a word that was typed, or one step of another kind. */
interface Entry {
  parts: Part[];
  key?: string;
  before: unknown;
  after: unknown;
  /** Someone else changed what it would bring back. */
  gone?: boolean;
}

interface Step {
  id: string;
  wires: DocWire[];
  /** What each of the commands did when it was applied last. */
  effects: Mark[][];
  /** The operations of the model that take the step back exactly, in the order they are applied in: for this side alone, while the step waits. */
  inverse: DocOp[];
  kind: 'edit' | 'undo' | 'redo';
  part: Part;
  entry: Entry;
  /** For an undoing or a making again: what it took from its stack, to be put back when the step is taken back unsent. */
  from?: Entry;
  /** Typing that more typing may be joined to while it waits (see submit()). */
  joins?: boolean;
}

interface Applied {
  inverse: DocOp[];
  effects: Mark[][];
  /** The commands that take it back; null when there are none. Only made for a step of this participant. */
  back: DocWire[] | null;
}

/** How long an undo waits for the steps of its participant to be placed. */
const UNDO_WAIT_MS = 3000;
/** Most parts of one entry: a run of typing longer than this is undone in pieces. */
const MAX_PARTS = 1000;
/** Most characters of typing that wait as one step. */
const MAX_JOINED = 256;

const flat = (effects: Mark[][]): Mark[] => ([] as Mark[]).concat(...effects);

const samePos = (a: WirePos, b: WirePos): boolean => orderOf(a, b) === 0;

/**
 * Puts that follow each other over the same blocks, one put in their place: the first takes out what
 * stands there, the last says what is to stand there in the end (an entry made of many steps over one
 * paragraph - a composition - is taken back by one put). Only where each put replaces exactly what
 * the one before it put there; anything else stays as it is.
 */
function joinedPuts(wires: DocWire[]): DocWire[] {
  const out: DocWire[] = [];
  for (const wire of wires) {
    const last = out[out.length - 1];
    if (wire.t === 'put' && last && last.t === 'put' && samePlace(last.at, wire.at) && wire.n === last.x.length) {
      out[out.length - 1] = { t: 'put', at: last.at, n: last.n, x: wire.x, h: last.h };
      continue;
    }
    out.push(wire);
  }
  return out;
}

const samePlace = (a: WirePos, b: WirePos): boolean => a.s === b.s && a.o === b.o && !a.b && !b.b && a.p.length === b.p.length && a.p.every((value, i) => value === b.p[i]);

/** Deletions that follow each other and lie side by side are one deletion: typing that is taken back letter by letter, as one range. */
function joinedBack(wires: readonly DocWire[]): DocWire[] {
  const out: DocWire[] = [];
  for (const wire of wires) {
    const last = out[out.length - 1];
    if (wire.t === 'del' && last && last.t === 'del' && orderOf(wire.at, wire.to) === -1 && orderOf(last.at, last.to) === -1) {
      // The one that runs second ends where the first began, or begins there.
      if (samePos(wire.to, last.at)) {
        out[out.length - 1] = { t: 'del', at: wire.at, to: last.to };
        continue;
      }
    }
    out.push(wire);
  }
  return out;
}

/**
 * Whether a document tracks its changes or holds tracked changes: such a document is not
 * edited together. A tracked change carries who made it and when and a number of its own, and
 * a change that is cut in two gets a new number from a counter of this participant: it is not
 * the same change for two participants.
 */
export function holdsRevisions(doc: DocModel): boolean {
  if (isTracking(doc)) return true;
  try {
    const found = listRevisions(doc, { limit: 1 });
    return found.list.length > 0 || (found as { more?: unknown }).more === true;
  } catch {
    return true;
  }
}

/** What a command that is kept for undo will do to the places around it, as far as its words say it. */
function effectOf(wire: DocWire): Mark[] {
  if (wire.t === 'put') return [{ k: 'blk', s: wire.at.s, c: wire.at.p.slice(0, -1), at: wire.at.p[wire.at.p.length - 1], n: wire.n, m: wire.x.length }];
  if (wire.t === 'del') {
    const order = orderOf(wire.at, wire.to);
    if (order === undefined || order === 0) return [];
    const [a, b] = order === -1 ? [wire.at, wire.to] : [wire.to, wire.at];
    const depth = a.p.length - 1;
    return [{ k: 'del', s: a.s, c: a.p.slice(0, depth), r1: a.p[depth], o1: a.o, r2: b.p[depth], o2: b.o }];
  }
  return [];
}

function readPlace(value: unknown): { path: number[]; offset: number } | undefined {
  if (!Array.isArray(value) || value.length !== 2) return undefined;
  const [path, offset] = value as unknown[];
  if (!Array.isArray(path) || path.length === 0 || path.length > 96 || typeof offset !== 'number' || Math.floor(offset) !== offset || offset < 0 || offset > 100000000) return undefined;
  for (const index of path) if (typeof index !== 'number' || Math.floor(index) !== index || index < 0 || index > 100000000) return undefined;
  return { path: path.slice() as number[], offset };
}

export class DocCollabSession extends RoomSession<Step, DocPresence> {
  private readonly host: DocSessionHost;
  private readonly options: DocSessionOptions;
  /** Undo and redo that wait for the steps on their way. */
  private wanted: ('undo' | 'redo')[] = [];
  private wantedTimer: (() => void) | undefined;
  private undoStack: Entry[] = [];
  private redoStack: Entry[] = [];
  /** The entry that is being merged into was ended (see DocHistory.seal). */
  private sealed = true;

  constructor(host: DocSessionHost, options: DocSessionOptions) {
    // Before anything is said to the room: see holdsRevisions().
    if (holdsRevisions(host.document())) throw new CollabUnsupportedError('other', 'A document that tracks its changes or holds tracked changes is not edited together');
    super(host, options);
    this.host = host;
    this.options = options;
    this.start();
  }

  get peers(): DocPeer[] {
    return super.peers;
  }

  get user(): DocPeer {
    return super.user;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  // ----------------------------------------------------------------- applying

  /** Takes back what was applied, as it was applied: nothing was changed in between. */
  private takeBack(inverse: readonly DocOp[]): void {
    const doc = this.host.document();
    for (const op of inverse) applyDocOp(doc, op);
  }

  /**
   * Runs the commands of a step one after the other. When one throws, the ones before it are
   * taken back and the error is passed on. `own`: the commands that take the step back are
   * made as well.
   */
  private applyStep(wires: readonly DocWire[], own: boolean): Applied {
    const doc = this.host.document();
    let inverse: DocOp[] = [];
    const effects: Mark[][] = [];
    let back: DocWire[] | null = own ? [] : null;
    try {
      for (const wire of wires) {
        const before = lookBefore(doc, wire);
        const edit = runWire(doc, wire);
        const done: DocOp[] = [];
        try {
          for (const op of edit.ops) done.unshift(applyDocOp(doc, op));
        } catch (error) {
          for (const op of done) applyDocOp(doc, op);
          throw error;
        }
        inverse = done.concat(inverse);
        const marks = edit.ops.length > 0 ? marksAfter(doc, before, wire.at.s) : [];
        effects.push(marks);
        if (back) {
          const undo = inverseOf(doc, wire, before, marks, edit.ops);
          // What takes the later command back runs first.
          back = undo ? undo.concat(back) : null;
        }
      }
    } catch (error) {
      this.takeBack(inverse);
      throw error;
    }
    return { inverse, effects, back };
  }

  /** The commands of a step as everybody will read them. Throws CollabUnsupportedError for what cannot travel. */
  private travelled(wires: readonly DocWire[]): DocWire[] {
    if (wires.length === 0 || wires.length > MAX_STEP_WIRES) throw new CollabUnsupportedError('size', 'The change is too large to be shared in one step');
    // What is applied here is what the others read: the step goes through its own text first.
    const text = JSON.stringify(encodeStep(wires));
    if (utf8Length(text) > MAX_STEP_BYTES) throw new CollabUnsupportedError('size', 'The change is too large to be shared in one step');
    try {
      return decodeStep(JSON.parse(text));
    } catch (error) {
      // What the others would refuse is not done here either (a text with a character no file holds, say).
      if (error instanceof WireError) throw new CollabUnsupportedError('other', 'The change holds something that cannot be shared');
      throw error;
    }
  }

  // ------------------------------------------------------------------ editing

  /**
   * Runs commands on the document as one step of this participant and queues it for the
   * others. Throws what the document throws for a command it does not take, and
   * CollabUnsupportedError for what cannot be shared; nothing was changed then. `meta`: what is
   * handed back when the step is undone (`before`) or made again (`after`); steps with the same
   * `key` that follow each other are undone together (typing a word), until seal() is called.
   * With `join`, a letter that is typed right behind letters that still wait for the relay goes
   * into their step (`joined` in the answer): however fast someone types, few steps wait.
   */
  submit(wires: readonly DocWire[], meta?: { before?: unknown; after?: unknown; key?: string; join?: boolean }): { marks: Mark[]; joined?: true } {
    if (this.readonlyNow) throw new CollabUnsupportedError('readOnly', 'The relay takes no changes from this participant');
    // While the file is loaded again the document is about to be another one: what is changed in it now would be lost here and sent to the others.
    if (this.reloading) throw new CollabUnsupportedError('other', 'The file is being loaded again');
    const read = this.travelled(wires);
    if (read.some((wire) => wire.t === 'put')) throw new CollabUnsupportedError('other', 'Paragraphs are put back by undo()');
    const done = this.applyStep(read, true);
    const marks = flat(done.effects);
    if (done.inverse.length === 0) return { marks };
    this.redoStack = [];
    // An undo that waited for its turn does not get it: what it was to take back is no longer the last thing done.
    if (this.wanted.length > 0) {
      this.wanted = [];
      this.host.notice({ code: 'undo-wait' });
    }
    const key = meta ? meta.key : undefined;
    const last = this.undoStack[this.undoStack.length - 1];
    const merges = last !== undefined && key !== undefined && last.key === key && !this.sealed && last.parts.length < MAX_PARTS;
    const typing = read.length === 1 && read[0].t === 'text' && read[0].to === undefined && read[0].x.indexOf('\n') < 0;
    if (merges && typing && meta && meta.key !== undefined && meta.join !== false && this.joinTyping(read[0] as Extract<DocWire, { t: 'text' }>, done, last as Entry)) {
      (last as Entry).after = meta.after;
      return { marks, joined: true };
    }
    const part: Part = { wires: done.back, moves: marks.some((mark) => mark.k !== 'tch') };
    let entry: Entry;
    if (merges) {
      entry = last as Entry;
      entry.parts.push(part);
      entry.after = meta ? meta.after : undefined;
    } else {
      entry = { parts: [part], key, before: meta ? meta.before : undefined, after: meta ? meta.after : undefined };
      this.undoStack.push(entry);
      this.sealed = false;
      const limit = this.options.undoLimit === undefined ? 200 : this.options.undoLimit;
      while (this.undoStack.length > Math.max(1, limit)) this.undoStack.shift();
    }
    const step: Step = { id: '', wires: read, effects: done.effects, inverse: done.inverse, kind: 'edit', part, entry };
    if (typing && meta && meta.key !== undefined && meta.join !== false) step.joins = true;
    part.step = step;
    this.placed(step, marks);
    return { marks, joined: undefined };
  }

  /**
   * A letter that was typed right behind the letters of the last waiting step goes into that
   * step. Both are taken back and the longer text is typed as one command - which is what the
   * others will run - so that this side holds what they will hold. False when the two do not
   * make one text (the step is on its way, or is about another place); nothing was changed then.
   */
  private joinTyping(wire: Extract<DocWire, { t: 'text' }>, done: Applied, entry: Entry): boolean {
    const prev = this.pendingSteps[this.pendingSteps.length - 1];
    if (!prev || !prev.joins || prev.kind !== 'edit' || prev.entry !== entry || prev.id === this.flying || this.current === 'closed' || this.current === 'detached') return false;
    const first = prev.wires[0] as Extract<DocWire, { t: 'text' }>;
    if (first.at.s !== wire.at.s || first.at.b || wire.at.b || first.at.p.join() !== wire.at.p.join() || first.at.o + first.x.length !== wire.at.o) return false;
    if (stable(first.f === undefined ? null : first.f) !== stable(wire.f === undefined ? null : wire.f) || first.x.length + wire.x.length > MAX_JOINED) return false;
    const both: DocWire = { ...first, x: first.x + wire.x };
    this.takeBack(done.inverse);
    this.takeBack(prev.inverse);
    let again: Applied;
    try {
      again = this.applyStep(this.travelled([both]), true);
    } catch (error) {
      // As they were, one after the other.
      const one = this.applyStep(prev.wires, true);
      prev.inverse = one.inverse;
      prev.effects = one.effects;
      prev.part.wires = one.back;
      const two = this.applyStep([wire], true);
      done.inverse = two.inverse;
      done.effects = two.effects;
      done.back = two.back;
      if (error instanceof RangeError) return false;
      throw error;
    }
    prev.wires = [both];
    prev.inverse = again.inverse;
    prev.effects = again.effects;
    prev.part.wires = again.back;
    if (this.paused > 0) this.touchedWhilePaused = true;
    this.movePresences(flat(done.effects));
    return true;
  }

  /** Ends the entry that is being merged into: what is typed next is undone by itself. */
  seal(): void {
    this.sealed = true;
  }

  /**
   * Undoes the last step of this participant. What is undone is always a step the relay has
   * placed: while steps of this participant are still on their way, the undoing waits for them
   * (undefined is answered now, and DocSessionHost.changed() is told when it was done; when
   * they do not arrive in a few seconds it is said and nothing is undone). Undefined also when
   * there is nothing to undo or the step cannot be undone, which is said.
   */
  undo(): { marks: Mark[]; meta: unknown } | undefined {
    return this.ask('undo');
  }

  redo(): { marks: Mark[]; meta: unknown } | undefined {
    return this.ask('redo');
  }

  private ask(kind: 'undo' | 'redo'): { marks: Mark[]; meta: unknown } | undefined {
    if (this.readonlyNow || this.reloading) return undefined;
    // As the history of one user: typing that follows an undoing is a step of its own, and one that follows a key that undid nothing is not.
    if ((kind === 'undo' ? this.undoStack : this.redoStack).length > 0) this.sealed = true;
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
    // The document is about to be another one.
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
    if (done) this.host.changed({ marks: done.marks, local: true, meta: done.meta });
    else this.runWanted();
  }

  private travel(kind: 'undo' | 'redo'): { marks: Mark[]; meta: unknown } | undefined {
    const stack = kind === 'undo' ? this.undoStack : this.redoStack;
    const entry = stack[stack.length - 1];
    if (!entry) return undefined;
    /** The entry stays behind: the one before it is next. */
    const given = (code: CollabNoticeCode): undefined => {
      this.abandon(stack, entry);
      this.host.notice({ code });
      return undefined;
    };
    if (entry.gone) return given('undo-gone');
    let wires: DocWire[] = [];
    // The step that was made last is taken back first.
    for (let i = entry.parts.length - 1; i >= 0; i--) {
      const part = entry.parts[i];
      if (part.wires === null) return given('undo-none');
      wires = wires.concat(part.wires);
    }
    wires = joinedPuts(joinedBack(wires));
    // Nothing is left of it (someone else deleted what was typed): there is nothing to take back, and nothing under it is in the way.
    if (wires.length === 0) {
      stack.pop();
      this.host.notice({ code: 'undo-gone' });
      return undefined;
    }
    if (wires.length > MAX_STEP_WIRES) return given('undo-none');
    let read: DocWire[];
    let done: Applied;
    try {
      read = this.travelled(wires);
      done = this.applyStep(read, true);
    } catch (error) {
      if (!(error instanceof RangeError)) throw error;
      // A step that cannot be shared (too large) cannot be undone; one the document does not take, someone else
      // changed what it would bring back.
      return given(error instanceof CollabUnsupportedError ? 'undo-none' : 'undo-gone');
    }
    stack.pop();
    const marks = flat(done.effects);
    // What takes the undoing back is there to be made again at once, as in the history of one user: what is done
    // next lies above it.
    const part: Part = { wires: done.back, moves: marks.some((mark) => mark.k !== 'tch') };
    const made: Entry = { parts: [part], before: entry.before, after: entry.after };
    (kind === 'undo' ? this.redoStack : this.undoStack).push(made);
    const step: Step = { id: '', wires: read, effects: done.effects, inverse: done.inverse, kind, part, entry: made, from: entry };
    part.step = step;
    this.placed(step, marks);
    return { marks, meta: kind === 'undo' ? entry.before : entry.after };
  }

  /**
   * An entry that will never be taken back leaves its stack. What lies under it was kept for
   * the document without its step: when that step moved text or blocks, none of it is known to
   * be right any more, and it is gone as well (which is said when it is asked for).
   */
  private abandon(stack: Entry[], entry: Entry): void {
    const at = stack.indexOf(entry);
    if (at < 0) return;
    stack.splice(at, 1);
    if (entry.parts.some((part) => part.moves)) for (let i = 0; i < at; i++) stack[i].gone = true;
  }

  /** A step that was applied here is queued for the relay; in a session that tells nobody it is a step like one the relay placed. */
  private placed(step: Step, marks: readonly Mark[]): void {
    this.movePresences(marks);
    if (this.paused > 0) this.touchedWhilePaused = true;
    if (this.current === 'closed' || this.current === 'detached') return this.settle(step);
    this.queue(step);
  }

  /**
   * A step of this participant has its place in the order (or, in a session that tells nobody,
   * was made): what takes it back is kept as it is from here on, and read through what the
   * others do.
   */
  protected settle(step: Step): void {
    const part = step.part;
    part.step = undefined;
    const entry = step.entry;
    const at = entry.parts.indexOf(part);
    if (part.wires === null) {
      // A step that cannot be taken back stays for good: what lies under it is read through it, once, like through a step of someone else.
      const marks = arrived(flat(step.effects));
      if (part.moves && this.undoStack.indexOf(entry) >= 0) this.throughStack(this.undoStack, marks, this.undoStack.indexOf(entry), at);
      // What could be made again under an undoing that cannot be: it was to be made on the document without it.
      else if (part.moves) for (let i = this.redoStack.indexOf(entry) - 1; i >= 0; i--) this.redoStack[i].gone = true;
      part.moves = false;
      return;
    }
    // Typing that is taken back letter by letter is kept as one range.
    const older = at > 0 ? entry.parts[at - 1] : undefined;
    if (older && !older.step && older.wires) {
      const both = joinedBack(part.wires.concat(older.wires));
      if (both.length < part.wires.length + older.wires.length) {
        older.wires = both;
        older.moves = older.moves || part.moves;
        entry.parts.splice(at, 1);
      }
    }
  }

  /**
   * Reads what is kept for undo through what was applied to the document (`carry`, as it is
   * read on the document as it is without the steps that wait). Each entry is about the
   * document without the steps above it, so the carry reaches it read through what takes
   * those steps back. An entry that does not go together with what was applied is gone; when
   * its step moved text, so is everything under it. `from`, `below`: only what lies under that
   * part of that entry is read.
   */
  private throughStack(stack: Entry[], carry: readonly Mark[], from = stack.length, below = 0): void {
    let over: Mark[] = carry.slice();
    let dead = false;
    for (let i = Math.min(from, stack.length - 1); i >= 0; i--) {
      const entry = stack[i];
      if (dead) {
        entry.gone = true;
        continue;
      }
      if (entry.gone) {
        if (entry.parts.some((part) => part.moves)) dead = true;
        continue;
      }
      for (let j = (i === from ? below : entry.parts.length) - 1; j >= 0 && !entry.gone; j--) {
        const part = entry.parts[j];
        // A step that waits is made anew, with what takes it back, when it is applied again.
        if (part.step || part.wires === null) continue;
        const next: DocWire[] = [];
        for (const wire of part.wires) {
          const moved = rebaseWire(wire, effectOf(wire), over);
          if (moved.drop) {
            entry.gone = true;
            break;
          }
          for (const one of moved.wires) next.push(one);
          over = moved.carry;
        }
        if (!entry.gone) part.wires = next;
      }
      if (entry.gone && entry.parts.some((part) => part.moves)) dead = true;
    }
  }

  /** A waiting step that is given up: what would have undone it is let go of. */
  private unlink(step: Step): void {
    if (this.flying === step.id) this.flying = undefined;
    const entry = step.entry;
    const at = entry.parts.indexOf(step.part);
    if (at >= 0) entry.parts.splice(at, 1);
    if (entry.parts.length === 0) {
      for (const stack of [this.undoStack, this.redoStack]) {
        const place = stack.indexOf(entry);
        if (place >= 0) stack.splice(place, 1);
      }
    }
  }

  // ---------------------------------------------------------------- the relay

  protected encodeFirst(): unknown[] {
    return encodeStep(this.pendingSteps[0].wires);
  }

  protected sameAsFirst(ops: unknown[]): boolean {
    // As values, not as text: a relay may write the same operations with their members in another order.
    return stable(encodeStep(this.pendingSteps[0].wires)) === stable(ops);
  }

  protected dropAll(code: CollabNoticeCode): void {
    const count = this.pendingSteps.length;
    for (let i = this.pendingSteps.length - 1; i >= 0; i--) {
      const step = this.pendingSteps[i];
      this.takeBack(step.inverse);
      this.unlink(step);
      // An undoing that was taken back took nothing back: what it was about is there to be undone again.
      if (step.from) (step.kind === 'undo' ? this.undoStack : this.redoStack).push(step.from);
    }
    this.pendingSteps = [];
    this.flying = undefined;
    if (count > 0) {
      this.host.notice({ code, count });
      this.host.changed({ marks: [], reset: true });
    }
    this.checkSettled();
  }

  protected remote(raw: unknown[]): boolean {
    let wires: DocWire[] | undefined;
    try {
      wires = decodeStep(raw);
    } catch (error) {
      if (!(error instanceof WireError) && !(error instanceof RangeError)) throw error;
    }
    if (!wires) {
      // No step of a document. Everybody reads it the same way and passes it over; the room goes on.
      this.badStep();
      return true;
    }
    const pending = this.pendingSteps;
    for (let i = pending.length - 1; i >= 0; i--) this.takeBack(pending[i].inverse);
    let carry: Mark[] = [];
    try {
      carry = arrived(flat(this.applyStep(wires, false).effects));
    } catch (error) {
      // The waiting steps go back on as they were: the document did not take what arrived.
      for (const step of pending) {
        const again = this.applyStep(step.wires, true);
        step.inverse = again.inverse;
        step.effects = again.effects;
        step.part.wires = again.back;
      }
      if (!(error instanceof RangeError)) throw error;
      // A step this document does not take: nothing of it was applied, here or anywhere (everybody applies
      // it to the same document). It is passed over.
      this.badStep();
      return true;
    }
    // What is kept for undo is about the document with the undoing (or redoing) that waits at the head of the queue -
    // its entry left the stack when it was asked for - and without the changes that wait behind it. What arrived is
    // read into the stacks as it reads after that undoing (see below).
    let stacked: Mark[] | undefined;
    let leading = true;
    // The stack the step at the head of the queue was made into (redo for an undoing, undo for a redoing) holds, under
    // that step, entries about the document BEFORE it: those are read with what arrived as it arrived.
    const travelling = pending.length > 0 && pending[0].kind !== 'edit' ? pending[0].kind : undefined;
    const before = carry;
    const kept: Step[] = [];
    let dropped = 0;
    let undoing = 0;
    let cut = 0;
    let rest = false;
    for (const step of pending) {
      if (leading && step.kind === 'edit') {
        leading = false;
        stacked = carry;
      }
      let next: DocWire[] = [];
      let over = carry;
      let failed = rest;
      let partly = false;
      for (let i = 0; !failed && i < step.wires.length; i++) {
        const moved = rebaseWire(step.wires[i], step.effects[i] || [], over);
        if (moved.drop) failed = true;
        else {
          next = next.concat(moved.wires);
          over = moved.carry;
          if (moved.cut) partly = true;
        }
      }
      if (!failed && next.length === 0) {
        // Nothing is left to do: what it was about went with what someone else deleted. That is said: should the
        // other one think again (undo), the text comes back without this change - a deletion of it did not happen here.
        this.stats.quiet++;
        cut++;
        this.unlink(step);
        carry = over;
        continue;
      }
      if (!failed && next.length <= MAX_STEP_WIRES) {
        try {
          const done = this.applyStep(next, true);
          step.wires = next;
          step.effects = done.effects;
          step.inverse = done.inverse;
          step.part.wires = done.back;
          // A step that was cut in pieces is no typing to join more typing to.
          if (next.length !== 1 || next[0].t !== 'text' || next[0].to !== undefined) step.joins = undefined;
          kept.push(step);
          carry = over;
          if (partly) cut++;
          continue;
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
      this.stats.dropped++;
      if (step.kind === 'edit') dropped++;
      else undoing++;
      this.unlink(step);
      // What an undoing that is given up was to take back stays in the document for good: what lies under it in its stack was kept for the document without it.
      if (step.from && step.from.parts.some((part) => part.moves)) for (const entry of step.kind === 'undo' ? this.undoStack : this.redoStack) entry.gone = true;
      // What was made after it was made on a document that had it: where that belongs now is not known, and it goes
      // with it. Unless the step moved nothing (it changed how text looks): then every place is where it was.
      if (flat(step.effects).some((mark) => mark.k !== 'tch')) rest = true;
    }
    this.pendingSteps = kept;
    // (An undoing that was dropped took what lies under it with it, when it moved anything; when it moved nothing, every place is where it was.)
    if (stacked === undefined) stacked = carry;
    this.throughStack(this.undoStack, travelling === 'redo' ? before : stacked);
    this.throughStack(this.redoStack, travelling === 'undo' ? before : stacked);
    this.movePresences(carry);
    if (dropped > 0) this.host.notice({ code: 'dropped', count: dropped });
    if (undoing > 0) this.host.notice({ code: 'undo-gone', count: undoing });
    if (cut > 0) this.host.notice({ code: 'partial', count: cut });
    this.host.changed(dropped + undoing > 0 ? { marks: carry, reset: true } : { marks: carry });
    return true;
  }

  // ----------------------------------------------------------------- presence

  protected readPresence(value: unknown): DocPresence | undefined {
    if (!value || typeof value !== 'object') return undefined;
    const from = value as { st?: unknown; a?: unknown; f?: unknown };
    const anchor = readPlace(from.a);
    const focus = readPlace(from.f);
    if (typeof from.st !== 'string' || from.st.length === 0 || from.st.length > 120 || !anchor || !focus) return undefined;
    return { story: from.st, anchor, focus };
  }

  protected writePresence(presence: DocPresence | undefined): unknown {
    return presence ? { st: presence.story, a: [presence.anchor.path, presence.anchor.offset], f: [presence.focus.path, presence.focus.offset] } : null;
  }

  /** The carets of the others follow text that came or went before them, until they say where they are again. */
  private movePresences(marks: readonly Mark[]): void {
    if (marks.length === 0) return;
    let moved = false;
    this.others.forEach((peer) => {
      const presence = this.readPresence(peer.presence);
      if (!presence) return;
      let anchor: WirePos | undefined = { s: presence.story, p: presence.anchor.path, o: presence.anchor.offset };
      let focus: WirePos | undefined = { s: presence.story, p: presence.focus.path, o: presence.focus.offset };
      for (const mark of marks) {
        if (!anchor || !focus) break;
        if (mark.k === 'opq' && (mark.s === '*' || mark.s === presence.story)) {
          anchor = undefined;
          break;
        }
        // Text that someone else put in at a caret stands before it: the caret is where its owner goes on typing.
        const a = mapPos(anchor, mark, true);
        const f = mapPos(focus, mark, true);
        anchor = a.gone && mark.k === 'blk' ? undefined : a.pos;
        focus = f.gone && mark.k === 'blk' ? undefined : f.pos;
      }
      const next = anchor && focus ? { st: presence.story, a: [anchor.p, anchor.o], f: [focus.p, focus.o] } : null;
      if (JSON.stringify(next) !== JSON.stringify(peer.presence)) {
        peer.presence = next;
        moved = true;
      }
    });
    if (moved) this.refreshPeers();
  }

  /** The document that is edited: what a file of the room is written from. */
  get model(): DocModel {
    return this.host.document();
  }
}
