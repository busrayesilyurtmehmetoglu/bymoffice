// One participant's connection to a room: what is the same whatever is edited together (a
// workbook, a document). The relay puts the steps of all participants into one order; this side
// joins, follows that order, sends its own steps one at a time and learns their place, falls
// back to the file when the room is at another version of it, and tells the others where it is.
// What a step is, how it is applied and how the waiting ones are read through one that arrived
// is the business of the session that extends this (session.ts for workbooks, doc/session.ts
// for documents).

import { MAX_PRESENCE_BYTES, cleanUser, parseServerMessage, type ClientMessage, type CollabUser, type PeerMessage } from './protocol';

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

/**
 * The version of the contract this side speaks: a room takes participants of one version only.
 * It goes up whenever a kind of step or a field of one is added to what goes over the wire (of
 * workbooks or of documents), or a step that is there comes to mean something else: a side that
 * does not know the new step would pass it over, and the room would drift apart. The fingerprint
 * of the wire in test/collab-protocol-version.test.ts says when. It goes up too when what a
 * formula calculates to changes (a function added, or one that gives something else): every
 * participant calculates the formulas of the room itself, and one that does not know a function
 * (#NAME?) or calculates it the old way drifts apart from the others with the same steps. The
 * fingerprint does not see that; it is raised by hand.
 * 3: panes of a sheet (`pane`), tab stops of paragraphs (`tabs`).
 * 4: dynamic array formulas (`d` of a formula).
 * 5: VSTACK, HSTACK, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, EXPAND, TOCOL, TOROW, WRAPROWS, WRAPCOLS and
 *    TEXTSPLIT calculated; TEXTBEFORE and TEXTAFTER take an array of delimiters as one list.
 */
export const PROTOCOL_VERSION = 5;

/** A participant of a room as the others see it; `P` is how it says where it is. */
export interface RoomPeer<P> {
  id: string;
  /** The name the participant gave; undefined for one who gave none (shown as a guest with `guest` as the number). */
  name?: string;
  /** "#RRGGBB": the colour given, or one of the palette. */
  color: string;
  guest: number;
  readOnly: boolean;
  presence?: P;
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
  /** The room is edited with another version of this package (another version of the contract): the page is to be loaded again. */
  | 'version'
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

/** What a session needs of whoever holds what is edited, whatever that is. */
export interface RoomHost {
  notice(notice: CollabNotice): void;
  /**
   * Loads the file again (the room is at a version this participant's file is not), after
   * which the host answers with what it holds then; resolves with the versions of the file
   * that was loaded. Without it the session closes when that happens.
   */
  reload?(): Promise<string[] | { versions: string[]; stamp?: RoomStamp }>;
  /** Whether changes are held that no file holds: such work is never given up for the file without being asked. */
  unsaved?(): boolean;
}

/** The room, as its relay names it, and the step of its order a file holds: what a file that was saved in a room says of itself. */
export interface RoomStamp {
  ep: string;
  rev: number;
}

export interface RoomOptions<P> {
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
  onPeers?(peers: RoomPeer<P>[]): void;
  /** Someone saved the file: its versions, and whether what this participant holds is what was saved. */
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

/** Colours participants are told apart by when they give none. */
export const PEER_COLORS = ['#D93025', '#1A73E8', '#188038', '#E37400', '#9334E6', '#00838F', '#C2185B', '#5F6368'];

/** How long after someone came in the last save is said again (to everybody who came meanwhile, once). */
const RESEND_SAVED_MS = 1500;
/** Most pending steps before it is said that many changes wait. */
export const BACKLOG_NOTICE = 500;
/** How often a selection is sent at most. */
const PRESENCE_MS = 100;
/** How often in a row the file is loaded again before the session gives up. */
const MAX_RELOADS = 6;
/** How long what the others say of where they are is gathered before it is passed on. */
const PEERS_MS = 50;
/** How long a step may be on its way before the connection is taken to be stuck. */
const FLY_MS = 30000;

export function hashOf(text: string): number {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) h = Math.imul(h ^ text.charCodeAt(i), 0x01000193);
  return h >>> 0;
}

function cleanStamp(stamp: RoomStamp | undefined): RoomStamp | undefined {
  if (!stamp || typeof stamp.ep !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(stamp.ep)) return undefined;
  return Number.isSafeInteger(stamp.rev) && stamp.rev >= 0 ? { ep: stamp.ep, rev: stamp.rev } : undefined;
}

/**
 * The connection of one participant. `S` is a step that waits for its place in the order (it
 * carries the name the relay answers with), `P` how a participant says where it is.
 *
 * A session that extends this calls start() at the end of its own constructor: nothing may
 * arrive before all of it is there.
 */
export abstract class RoomSession<S extends { id: string }, P> {
  private readonly roomHost: RoomHost;
  private readonly roomOptions: RoomOptions<P>;
  protected readonly transport: CollabTransport;
  /** What the steps of this session are told apart by, whatever the relay calls the connection. */
  private readonly key: string;
  protected seq = 0;
  protected current: CollabState = 'connecting';
  /** How many steps of the order are held under the pending ones. */
  protected confirmed = 0;
  /** The order was joined at least once: joining again asks for what came since. */
  private joined = false;
  /** The order was asked for again after a step was missing, and the answer is not there yet. */
  private resyncing = false;
  private badStepSaid = false;
  /** What closed the session, when it closed by itself (see reason). */
  private endedBy: CollabNoticeCode | undefined;
  /** What the steps of this session are named with, for everybody to see; the key is told to the relay alone. */
  protected readonly tag = `${hashOf(String(Math.random())).toString(36)}${hashOf(String(Date.now())).toString(36)}`;
  private flyTimer: (() => void) | undefined;
  private slowSaid = false;
  /** The step that was held when the room was joined with the file. */
  protected startedAt = 0;
  private lastSaved: ClientMessage | undefined;
  /** A save was made while the room could not be reached, and nobody was told. */
  private savedUnsaid = false;
  /** The relay counts this connection as one of a room: it takes no second "join" from it. */
  private inRoom = false;
  private waitingSaid = false;
  private peersTimer: (() => void) | undefined;
  private peersStale = false;
  /** The others moved, and whoever listens was not told yet. */
  private peersDue = false;
  /** Something was changed here while what arrives was held back (see pause()). */
  protected touchedWhilePaused = false;
  /** The step the relay is at as far as this side knows: nothing is sent while this side is behind it. */
  private target = 0;
  protected pendingSteps: S[] = [];
  /** settle() is called for a step the relay placed (at `confirmed`), not for one that is kept without the room. */
  /** The room this side is in, as the relay names it; and what the file says it was saved at, until it was used or found wrong. */
  private epoch: string | undefined;
  private stamp: RoomStamp | undefined;
  /** The id of the step that is on its way. */
  protected flying: string | undefined;
  private versions: string[];
  private me = '';
  private resendTimer: (() => void) | undefined;
  private seenAs: CollabUser | undefined;
  protected others: Map<string, PeerMessage> = new Map();
  private peerList: RoomPeer<P>[] = [];
  protected readonlyNow: boolean;
  private open = false;
  private reloads = 0;
  protected reloading = false;
  protected backlogSaid = false;
  private presenceNow: P | undefined;
  private presenceSent = '';
  private presenceTimer: (() => void) | undefined;
  private settled: (() => void)[] = [];
  protected paused = 0;
  private inbox: string[] = [];
  /** How many steps this participant queued, and what became of them: placed by the relay, dropped and said, or dropped because what they wanted was so already. */
  readonly stats = { made: 0, placed: 0, dropped: 0, quiet: 0 };

  constructor(host: RoomHost, options: RoomOptions<P>) {
    this.roomHost = host;
    this.roomOptions = options;
    this.stamp = cleanStamp(options.stamp);
    this.transport = options.transport;
    this.versions = options.versions.slice(0, 8);
    this.readonlyNow = options.readOnly === true;
    this.key = `${hashOf(String(Math.random())).toString(36)}${hashOf(String(Date.now() + Math.random())).toString(36)}`;
  }

  /** Listens to the transport and joins the room. Called once, by the session that extends this, when all of it is there. */
  protected start(): void {
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
          // What is detached stays so, whatever the connection does.
          if (this.current !== 'detached') this.setState('offline');
        }
      });
    } else {
      this.open = true;
      this.join();
    }
  }

  // ------------------------------------------------- what a session says of its steps

  /** The first of the pending steps as it goes to the relay. */
  protected abstract encodeFirst(): unknown[];
  /** Whether what the relay placed under the name of the first pending step is that step. */
  protected abstract sameAsFirst(ops: unknown[]): boolean;
  /** A step of this participant has its place in the order (or, in a session that tells nobody, was made). */
  protected abstract settle(step: S): void;
  /** A step of someone else: placed under the pending ones. False when the session was closed over it. */
  protected abstract remote(raw: unknown[]): boolean;
  /** Takes every pending step back and forgets it. */
  protected abstract dropAll(code: CollabNoticeCode): void;
  /** Nothing of this participant waits for the relay any more, and the room is there: what waited for that has its turn. */
  protected abstract idle(): void;
  /** What waited for its turn (an undo) is given up; with `all`, everything that could be undone is forgotten as well (the file is loaded again). */
  protected abstract forget(all: boolean): void;
  /** Where a participant says it is, read from what arrived; undefined for what says nothing. */
  protected abstract readPresence(value: unknown): P | undefined;
  /** Where this participant is, as it is told to the others. */
  protected abstract writePresence(presence: P | undefined): unknown;

  get state(): CollabState {
    return this.current;
  }

  /**
   * Why a session that is `closed` ended by itself: `version` (the room speaks another version of
   * the contract), `refused`, `peer` or `stale`, from the moment its notice is given (the session
   * closes right after); undefined while it is open, and for one that was left.
   */
  get reason(): CollabNoticeCode | undefined {
    // Set before the notice is given, so that whoever hears the notice can read it.
    return this.endedBy;
  }

  /** How many steps of the common order are held. */
  get rev(): number {
    return this.confirmed;
  }

  /** How many steps of this participant the relay has not placed yet. */
  get pending(): number {
    return this.pendingSteps.length;
  }

  get peers(): RoomPeer<P>[] {
    if (this.peersStale) this.buildPeers();
    return this.peerList.slice();
  }

  /** This participant as the others see it; `id` is empty until the relay named the connection. */
  get user(): RoomPeer<P> {
    return this.lookOf(this.me, this.seenAs || cleanUser(this.roomOptions.user), this.readonlyNow, this.presenceNow);
  }

  get readOnly(): boolean {
    return this.readonlyNow;
  }

  /** A step that was applied here gets its name and waits for the relay. */
  protected queue(step: S): void {
    step.id = `${this.tag}:${++this.seq}`;
    this.pendingSteps.push(step);
    this.stats.made++;
    if (this.pendingSteps.length > BACKLOG_NOTICE && !this.backlogSaid) {
      this.backlogSaid = true;
      this.roomHost.notice({ code: 'backlog', count: this.pendingSteps.length });
    }
    this.pump();
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
    const user = cleanUser(this.roomOptions.user);
    if (user.name !== undefined || user.color !== undefined) message.user = user;
    if (typeof this.roomOptions.token === 'string' && this.roomOptions.token !== '') message.token = this.roomOptions.token;
    if (this.roomOptions.readOnly === true) message.ro = true;
    if (typeof this.roomOptions.room === 'string' && this.roomOptions.room !== '') message.room = this.roomOptions.room;
    this.send(message);
  }

  protected pump(): void {
    if (this.current !== 'online' || this.flying !== undefined || this.pendingSteps.length === 0 || this.confirmed < this.target || this.readonlyNow) {
      this.checkSettled();
      return;
    }
    const step = this.pendingSteps[0];
    this.flying = step.id;
    this.send({ t: 'op', base: this.confirmed, id: step.id, ops: this.encodeFirst() });
    // A step that is neither placed nor refused for a long time: the connection is not what it seems.
    if (this.flyTimer) this.flyTimer();
    const id = step.id;
    const late = (): void => {
      this.flyTimer = undefined;
      if (this.flying !== id || this.current !== 'online') return;
      // While what arrives is held back (a file is written, an input method composes) the answer may well be among it.
      if (this.paused > 0) {
        this.flyTimer = this.timer(late, FLY_MS);
        return;
      }
      if (!this.slowSaid) {
        this.slowSaid = true;
        this.roomHost.notice({ code: 'slow' });
      }
      this.flying = undefined;
      if (!this.reconnect()) this.pump();
    };
    this.flyTimer = this.timer(late, FLY_MS);
  }

  protected checkSettled(): void {
    if (this.pendingSteps.length === 0 && this.current === 'online') this.idle();
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

  protected timer(callback: () => void, ms: number): () => void {
    if (this.roomOptions.timer) return this.roomOptions.timer(callback, ms);
    const handle = setTimeout(callback, ms);
    return () => clearTimeout(handle);
  }

  private setState(state: CollabState): void {
    if (this.current === state) return;
    this.current = state;
    if (this.roomOptions.onState) this.roomOptions.onState(state);
  }

  private fatal(code: CollabNoticeCode): void {
    if (this.current === 'closed') return;
    this.endedBy = code;
    this.roomHost.notice({ code });
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

  /** Leaves the room. What is edited stays as it is; what is changed from here on is told to nobody. */
  leave(): void {
    this.close();
  }

  /**
   * Holds what arrives until resume(): while a file is written, nothing that arrives may
   * change what it is written from half way through.
   */
  pause(): void {
    if (this.paused++ === 0) this.touchedWhilePaused = false;
  }

  /** Lets what was held back in. Answers whether something was changed here meanwhile: a file written in that time is of no one step. */
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
      // Nothing that arrives may throw out of here. What is edited is kept as it is, apart from the room.
      const now = this.current as CollabState;
      if (now !== 'closed' && now !== 'detached') this.detach();
    }
  }

  private handle(message: ReturnType<typeof parseServerMessage>): void {
    switch (message.t) {
      case 'hello': {
        // The log from before where this side is brings steps it has, which are passed over; one that starts
        // behind it leaves steps out.
        // A relay that names a version names that of the room; one of another version is not joined.
        if (message.pv !== undefined && message.pv !== PROTOCOL_VERSION) return this.fatal('version');
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
          this.begun();
        }
        this.joined = true;
        this.resyncing = false;
        // (The count of loads is not started again here: a room that says welcome and sends the file away
        // again has brought nothing forward. A step that is taken does: see below.)
        this.target = message.rev;
        this.me = message.you;
        // As the room has this participant: the relay may have named it itself.
        this.seenAs = message.user;
        if (message.ro && !this.readonlyNow) {
          this.readonlyNow = true;
          // Asked for by nobody here: whoever edits is told, and what waits is taken back.
          if (this.roomOptions.readOnly !== true) {
            this.dropAll('refused');
            this.roomHost.notice({ code: 'read-only' });
            if (this.roomOptions.onReadOnly) this.roomOptions.onReadOnly();
          }
        }
        this.others.clear();
        for (const peer of message.peers) this.others.set(peer.id, peer);
        this.refreshPeers();
        this.setState('online');
        // A save that was made while the room could not be reached is said now.
        if (this.savedUnsaid && this.lastSaved) this.send(this.lastSaved);
        this.savedUnsaid = false;
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
          const mine = this.pendingSteps.shift() as S;
          this.stats.placed++;
          this.confirmed = message.rev;
          this.flying = undefined;
          this.slowSaid = false;
          this.settle(mine);
          if (this.pendingSteps.length < BACKLOG_NOTICE) this.backlogSaid = false;
        } else if (!this.remote(message.ops)) return;
        else this.confirmed = message.rev;
        // The room went on with this side in it.
        this.reloads = 0;
        this.waitingSaid = false;
        if (this.target < this.confirmed) this.target = this.confirmed;
        this.pump();
        return;
      case 'nack':
        if (this.flying !== message.id) return;
        this.flying = undefined;
        if (message.reason === 'read-only') {
          this.readonlyNow = true;
          this.dropAll('refused');
          this.roomHost.notice({ code: 'read-only' });
          if (this.roomOptions.onReadOnly) this.roomOptions.onReadOnly();
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
        // Once for however many come in a moment: each word of a save makes everybody ask the server.
        if (this.lastSaved && !this.resendTimer) {
          this.resendTimer = this.timer(() => {
            this.resendTimer = undefined;
            if (this.lastSaved && this.current === 'online') this.send(this.lastSaved);
          }, RESEND_SAVED_MS);
        }
        return;
      case 'left':
        if (this.others.delete(message.id)) this.refreshPeers();
        return;
      case 'saved': {
        const clean = message.rev !== undefined && message.rev === this.confirmed && this.pendingSteps.length === 0;
        if (this.roomOptions.onSaved) this.roomOptions.onSaved({ versions: message.v, rev: message.rev, clean });
        return;
      }
      case 'sync':
        if (this.roomOptions.onSync && !this.readonlyNow) this.roomOptions.onSync();
        return;
      case 'pong':
        return;
      case 'error':
        // This side's own new connection took the place of this one: nothing to say.
        if (message.code === 'replaced') return;
        // The room is edited by a page of another version: what is wrong is the page, not this participant.
        if (message.code === 'version') return this.fatal('version');
        return this.fatal('refused');
    }
  }

  /** The room was joined with the file, at the step `startedAt`: nothing of the order is held yet. */
  protected begun(): void {
    // Nothing to do for a session that keeps nothing by step.
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

  protected badStep(): void {
    if (this.badStepSaid) return;
    this.badStepSaid = true;
    this.roomHost.notice({ code: 'bad-step' });
  }

  /**
   * The room is at a version of the file this participant does not have. What holds nothing
   * but its file is given up for the file as it is stored now, and the room is joined anew.
   * What holds changes no file has is never given up without being asked: it stays as it is,
   * detached from the room (see CollabState).
   */
  protected restart(): void {
    if (this.reloading) return;
    let ahead = this.pendingSteps.length > 0;
    try {
      // A host that does not say whether its work is saved is taken to have saved nothing since it joined.
      if (!ahead) ahead = this.roomHost.unsaved ? this.roomHost.unsaved() === true : this.confirmed > this.startedAt;
    } catch {
      ahead = true;
    }
    if (ahead) return this.detach();
    this.loadAgain();
  }

  private detach(): void {
    // What waited is in what is edited, and stays there; nobody will be told of it.
    for (const step of this.pendingSteps) this.settle(step);
    this.pendingSteps = [];
    this.forget(false);
    this.lastSaved = undefined;
    this.savedUnsaid = false;
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
    this.roomHost.notice({ code: 'detached' });
    this.checkSettled();
  }

  /**
   * Joins the room again with what is edited as it is: after it was saved, under what the saved file is
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
   * Gives what is edited up for the file as it is stored and joins the room with that: what a detached
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
    const reload = this.roomHost.reload;
    if (!reload) return this.fatal('stale');
    if (++this.reloads > MAX_RELOADS) {
      // Load after load with nothing in between: it is said, and what is edited stays, apart from the room.
      this.roomHost.notice({ code: 'stale' });
      return this.detach();
    }
    const lost = this.pendingSteps.length;
    this.reloading = true;
    this.lastSaved = undefined;
    this.savedUnsaid = false;
    this.pendingSteps = [];
    this.flying = undefined;
    this.forget(true);
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
      this.roomHost.notice({ code: 'waiting' });
    }
    const load = (): void => {
      if (this.current === 'closed') return;
      let loading: Promise<string[] | { versions: string[]; stamp?: RoomStamp }>;
      try {
        loading = Promise.resolve(reload.call(this.roomHost));
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
          if (lost > 0) this.roomHost.notice({ code: 'reloaded', count: lost });
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

  // ----------------------------------------------------------------- presence

  private lookOf(id: string, user: CollabUser, readOnly: boolean, presence?: P): RoomPeer<P> {
    const hash = hashOf(id);
    const peer: RoomPeer<P> = { id, color: user.color !== undefined ? user.color : PEER_COLORS[hash % PEER_COLORS.length], guest: (hash % 99) + 1, readOnly };
    if (user.name !== undefined) peer.name = user.name;
    if (presence) peer.presence = presence;
    return peer;
  }

  private buildPeers(): void {
    this.peersStale = false;
    const list: RoomPeer<P>[] = [];
    this.others.forEach((peer) => list.push(this.lookOf(peer.id, peer.user, peer.ro, this.readPresence(peer.presence))));
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

  protected refreshPeers(): void {
    this.peersStale = false;
    this.peersDue = false;
    const list: RoomPeer<P>[] = [];
    this.others.forEach((peer) => list.push(this.lookOf(peer.id, peer.user, peer.ro, this.readPresence(peer.presence))));
    this.peerList = list;
    if (this.roomOptions.onPeers) this.roomOptions.onPeers(list.slice());
  }

  /** Tells the others where this participant is; at most ten times a second. */
  setPresence(presence: P | undefined): void {
    this.presenceNow = presence;
    if (this.presenceTimer || this.current === 'closed') return;
    this.presenceTimer = this.timer(() => {
      this.presenceTimer = undefined;
      this.flushPresence();
    }, PRESENCE_MS);
  }

  private flushPresence(): void {
    const data = this.writePresence(this.presenceNow);
    const text = JSON.stringify(data);
    if (text === this.presenceSent || this.current !== 'online' || text.length > MAX_PRESENCE_BYTES) return;
    this.presenceSent = text;
    this.send({ t: 'presence', data });
  }

  // ------------------------------------------------------------------- saving

  /**
   * What a file written now is to say of itself: the room, and the step of its order that is held.
   * Undefined while something of this participant waits, or this side is in no room.
   */
  get stampNow(): RoomStamp | undefined {
    return this.current === 'online' && this.pendingSteps.length === 0 && this.epoch !== undefined ? { ep: this.epoch, rev: this.confirmed } : undefined;
  }

  /**
   * The step of the common order that is held when nothing of this participant waits: a file
   * written now is the file of that step. Undefined while something waits.
   */
  get savableRev(): number | undefined {
    return this.current === 'online' && this.pendingSteps.length === 0 ? this.confirmed : undefined;
  }

  /** Tells the others that the file was saved: what it is known by now and, when known, the step it holds. */
  announceSaved(versions: string[], rev: number | undefined): void {
    const v = versions.filter((version) => typeof version === 'string' && version !== '' && version.length <= 200).slice(0, 8);
    if (v.length === 0 || this.current === 'closed' || this.current === 'detached') return;
    const message: ClientMessage = { t: 'saved', v };
    if (rev !== undefined) message.rev = rev;
    this.lastSaved = message;
    // Out of reach of the room: said when it is reached again (the others would meet a conflict with a file nobody told them of).
    if (this.current !== 'online') {
      this.savedUnsaid = true;
      return;
    }
    this.send(message);
  }
}
