// What ties a session to the saving of the file: what someone else says was saved is asked of
// the server before it is believed, and the one participant who saves for the others does so
// when the room asks. Nothing here needs a page: mount() hands in what it has.

import type { CollabState } from './session';

/** What is needed of the save controller. */
export interface SyncController {
  readonly currentETag: string | undefined;
  readonly dirty: boolean;
  /** Whether this side itself changed something that no save of its own holds. */
  readonly ownChanges: boolean;
  adopt(etag: string | undefined, clean: boolean): void;
  hint(etag: string | undefined): void;
  changed(remote?: boolean): void;
  saveQuietly(): Promise<boolean>;
  rearm(): void;
  readonly autoSave: { delayMs: number } | undefined;
  /** Another number after every change. */
  readonly changes: number;
  readonly saving: boolean;
}

/** What is needed of the session. */
export interface SyncSession {
  readonly state: CollabState;
  readonly rev: number;
  readonly pending: number;
  readonly leads: boolean;
}

export interface SaveSyncHost {
  controller: SyncController;
  session(): SyncSession | undefined;
  /** Where the file is saved to with a request; undefined for a function of the application (which checks for itself). */
  target: { url: string; credentials: RequestCredentials; headers?: [string, string][] } | undefined;
  /** How many are in the room, this participant too: the more, the wider apart they stand in for a saver who does not save. */
  people?(): number;
  /** An ETag that `If-Match` can be sent with, or undefined. */
  strong(etag: string | null | undefined): string | undefined;
  readOnly(): boolean;
  fetch?: typeof fetch;
  /** Calls back once after `ms`; answers what cancels it. */
  timer?(run: () => void, ms: number): () => void;
  now?(): number;
  /** How long work waits for the one who saves for all, in place of five delays of auto-save and at least WATCH_MIN_MS. */
  watchMs?: number;
  /** Between 0 and 1: what keeps two participants from standing in at the same moment. */
  random?(): number;
  /**
   * The file as this workbook would be written now, with the step it holds (undefined: the file of no one step).
   * With it a save of someone else can be PROVEN to hold this side's own work: the same step gives the same bytes.
   */
  ownBytes?(): Promise<{ bytes: Uint8Array; rev: number | undefined }>;
  /**
   * Whether two files are the same file whatever they were packed with (two browsers do not compress alike):
   * the same parts, each with the same content. Without it: the same bytes.
   */
  sameFile?(stored: Uint8Array, own: Uint8Array): Promise<boolean>;
}

/** A stored file larger than this is not fetched to prove anything: the work stays to be saved from here. */
export const PROOF_MAX_BYTES = 32 * 1024 * 1024;
/** How long after word of a save the stored file is looked at (word of another save meanwhile starts the wait over). */
export const PROOF_DELAY_MS = 400;
/** For how long what the server said of the name of the file is taken to hold when the same name is said again. */
export const CONFIRMED_FOR_MS = 5000;
/** The widest those who wait are spread, however many they are. */
export const WATCH_SPREAD_MAX_MS = 30000;
/** How far apart, for each further participant, those who wait are spread. */
export const WATCH_SLOT_MS = 3000;
/** The least time this side's own work waits for a save of the room before this side saves it itself. */
export const WATCH_MIN_MS = 15000;

/** How long a save that met a conflict waits for word of the save that made it. */
export const CONFLICT_WAIT_MS = 2500;

export interface SaveSync {
  /** A participant says the file was saved. Resolves when what was said has been looked into. */
  saved(said: { versions: string[]; rev?: number }): Promise<void>;
  /** The room asks for the file to be saved (someone cannot join). */
  sync(): Promise<boolean>;
  /** Whether the server names the file by this ETag now. */
  storedAs(tag: string): Promise<boolean>;
  /** A change arrived from someone else: there is something to save, and the one who saves for all is to save it. */
  remoteChanged(): void;
  /** This side changed something: its work waits for a save from now, and not for ever. */
  localChanged(): void;
  /** Whether this side is the one to save by itself now (SaveHost.autoAllowed). */
  autoAllowed(): boolean;
  /** Why it is not (SaveHost.autoWaiting). */
  waiting(): 'room' | 'offline' | undefined;
  /** A save of this side was taken: its work waits for nothing. */
  ownSaved(): void;
  /** Nothing more is looked at or waited for. */
  stop(): void;
  /** A save here met a conflict (SaveHost.conflictWait): true when a save of someone in the room explains it and was found to be so. */
  conflictWait(): Promise<boolean>;
  /** Who is in the room changed, or the state of the session: the one who saves for all may be this side now. */
  roomChanged(): void;
}

/** The body of an answer, read no further than `limit` bytes: undefined when it is longer (what is left of it is not fetched). */
async function bodyOf(answer: Response, limit: number): Promise<Uint8Array | undefined> {
  const stream = answer.body as (ReadableStream<Uint8Array> & { getReader?: unknown }) | null;
  if (!stream || typeof stream.getReader !== 'function') {
    const all = new Uint8Array(await answer.arrayBuffer());
    return all.length > limit ? undefined : all;
  }
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let length = 0;
  for (;;) {
    const part = await reader.read();
    if (part.done) break;
    length += part.value.length;
    if (length > limit) {
      void reader.cancel().catch(() => undefined);
      return undefined;
    }
    chunks.push(part.value);
  }
  const out = new Uint8Array(length);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

/** The headers of the page for its target, without those that are about a body or a condition. */
function sent(target: { headers?: [string, string][] }): { [name: string]: string } {
  const out: { [name: string]: string } = {};
  for (const [name, value] of target.headers || []) if (!/^(content-|if-)/i.test(name)) out[name] = value;
  return out;
}

export function createSaveSync(host: SaveSyncHost): SaveSync {
  const ask = host.fetch || (typeof fetch !== 'undefined' ? fetch : undefined);

  /** What the server was last found to name the file by, and when: the same word again within a moment asks nothing. */
  let confirmed: { tag: string; at: number } | undefined;
  const storedAs = async (tag: string): Promise<boolean> => {
    if (confirmed && confirmed.tag === tag && nowMs() - confirmed.at < CONFIRMED_FOR_MS) return true;
    const found = await asks(tag);
    confirmed = found ? { tag, at: nowMs() } : undefined;
    return found;
  };
  const nowMs = (): number => (host.now ? host.now() : Date.now());
  /** Tags the server itself was found to name the file otherwise than: a word of one of them is no hint of anything. */
  const refuted = new Set<string>();
  const asks = async (tag: string): Promise<boolean> => {
    const target = host.target;
    if (!target || !ask || host.strong(tag) === undefined) return false;
    // With what the file is asked for when it is saved (a server that knows its users by a header knows them here too).
    const request = (method: string, headers?: { [name: string]: string }): Promise<Response> => ask(target.url, { method, cache: 'no-store', credentials: target.credentials, headers: { ...sent(target), ...(headers || {}) } });
    try {
      let answer = await request('HEAD');
      if (answer.status === 405 || answer.status === 501) {
        // A server without HEAD: the file is asked for only if it is another one, and its body is not read.
        answer = await request('GET', { 'If-None-Match': tag });
        const same = answer.status === 304 || (answer.ok && host.strong(answer.headers.get('ETag')) === tag);
        if (answer.body) void answer.body.cancel().catch(() => undefined);
        return same;
      }
      const named = answer.ok ? host.strong(answer.headers.get('ETag')) : undefined;
      // The server names the file otherwise: what was said is no longer so (or never was), and leads nowhere.
      if (named !== undefined && named !== tag) {
        refuted.add(tag);
        if (refuted.size > 64) refuted.delete(refuted.values().next().value as string);
      }
      return named === tag;
    } catch {
      return false;
    }
  };

  /** Saves of others that are being looked into, and who waits for word of one. */
  let checking = 0;
  let waiters: ((moved: boolean) => void)[] = [];
  const answer = (moved: boolean): void => {
    const all = waiters;
    waiters = [];
    for (const waiter of all) waiter(moved);
  };

  const saved = async (said: { versions: string[]; rev?: number }): Promise<void> => {
    checking++;
    const before = host.controller.currentETag;
    try {
      await look(said);
    } finally {
      checking--;
      // Whoever waited for this hears whether it moved what the file is known by.
      // (A word that could not be confirmed ends nobody's wait: the one that can be may be a moment behind it.)
      if (host.controller.currentETag !== before) answer(true);
    }
  };

  const look = async (said: { versions: string[]; rev?: number }): Promise<void> => {
    let tag: string | undefined;
    for (const version of said.versions) if (typeof version === 'string' && version.slice(0, 2) === 'e:') tag = version.slice(2);
    if (tag === undefined) return;
    const session = host.session();
    if (!session) return;
    // The ETag this side knows the file by already, with changes made since: whatever is said, that file does not hold them.
    const stale = host.controller.currentETag === tag && host.controller.dirty;
    const stored = await storedAs(tag);
    const now = host.session();
    if (!now || now !== session || now.state !== 'online') return;
    // Not found to be so: nothing is taken to be saved. The ETag is only tried, as a condition, if a save here meets a conflict.
    // (Only when the server could not be asked: a name the server itself says is not the file's is no hint of anything.)
    if (!stored) return refuted.has(tag) ? undefined : host.controller.hint(tag);
    // Saved here as well only when this workbook is, right now, the workbook of the step that was saved.
    const atStep = !stale && said.rev !== undefined && now.rev === said.rev && now.pending === 0;
    if (!host.controller.ownChanges) {
      host.controller.adopt(tag, atStep);
      if (atStep) ownSaved();
      return;
    }
    // Work this side did itself is saved when a file holds it, and on nobody's word - whoever says it, and
    // whatever it is in the room. The name of the file is taken; whether the file holds the work is looked at.
    host.controller.adopt(tag, false);
    if (atStep) prove(tag, said.rev as number);
  };

  const later = (run: () => void, ms: number): (() => void) =>
    (host.timer ||
      ((callback, wait) => {
        const id = setTimeout(callback, wait);
        return () => clearTimeout(id);
      }))(run, ms);

  /** The last save that was heard of and is to be looked at: one look, a moment after the last word. */
  let proving: { tag: string; rev: number; stop: () => void } | undefined;
  const prove = (tag: string, rev: number): void => {
    if (!host.ownBytes || !host.target || !ask) return;
    if (proving) proving.stop();
    const mine = { tag, rev, stop: later(() => void check(mine), PROOF_DELAY_MS) };
    proving = mine;
  };
  const check = async (mine: { tag: string; rev: number }): Promise<void> => {
    if (proving !== mine) return;
    proving = undefined;
    const target = host.target;
    const session = host.session();
    const controller = host.controller;
    if (!target || !ask || !host.ownBytes || !session) return;
    const holds = (): boolean => host.session() === session && session.state === 'online' && session.rev === mine.rev && session.pending === 0 && controller.dirty && !controller.saving && controller.currentETag === mine.tag;
    if (!holds()) return;
    const changes = controller.changes;
    try {
      // The file as it is stored, when it is the one that was named and not too large to look at.
      const answer = await ask(target.url, { method: 'GET', cache: 'no-store', credentials: target.credentials, headers: sent(target) });
      const length = Number(answer.headers.get('Content-Length'));
      if (!answer.ok || host.strong(answer.headers.get('ETag')) !== mine.tag || (length === length && length > PROOF_MAX_BYTES)) {
        if (answer.body) void answer.body.cancel().catch(() => undefined);
        return;
      }
      const stored = await bodyOf(answer, PROOF_MAX_BYTES);
      if (!stored || !holds() || controller.changes !== changes) return;
      // The same workbook at the same step is written as the same bytes (the room and the step are among them).
      const own = await host.ownBytes();
      if (own.rev !== mine.rev || !holds() || controller.changes !== changes) return;
      if (host.sameFile) {
        if (!(await host.sameFile(stored, own.bytes)) || !holds() || controller.changes !== changes) return;
      } else {
        if (own.bytes.length !== stored.length) return;
        for (let i = 0; i < stored.length; i++) if (own.bytes[i] !== stored[i]) return;
      }
      controller.adopt(mine.tag, true);
      ownSaved();
    } catch {
      // Not found out: the work stays to be saved.
    }

  };

  const sync = async (): Promise<boolean> => {
    const session = host.session();
    // One participant saves for the others; a workbook that is only read, or has steps on their way, does not.
    if (!session || session.state !== 'online' || !session.leads || session.pending > 0 || host.readOnly()) return false;
    if (!host.controller.dirty) host.controller.changed(true);
    // As a save nobody asked for: what is being typed is not applied for it.
    return host.controller.saveQuietly();
  };

  const clock = (): number => (host.now ? host.now() : Date.now());
  /** Since when this side has work (its own, or of others) that no file is known to hold; and whether it stands in for the saver now. */
  let ownSince: number | undefined;
  let standingIn = false;
  let watch: (() => void) | undefined;
  /** How long this side's own work waits for the one who saves for all: five delays of auto-save, a quarter more or less. */
  let patience: number | undefined;
  const baseMs = (): number => {
    const auto = host.controller.autoSave;
    return host.watchMs !== undefined ? host.watchMs : Math.max(WATCH_MIN_MS, (auto ? auto.delayMs : 0) * 5);
  };
  const watchMs = (): number => {
    if (patience === undefined) {
      const base = baseMs();
      // Wider apart the more there are who may stand in (those who write; who only reads saves nothing): one save
      // takes a moment, and the word of it clears the others before their turn comes. A quarter of the wait at
      // least, for two; half a minute at most, however many there are.
      const people = Math.max(2, host.people ? host.people() : 2);
      const spread = Math.min(WATCH_SPREAD_MAX_MS, Math.max(base * 0.25, (people - 1) * WATCH_SLOT_MS));
      patience = Math.round(base + spread * (host.random ? host.random() : Math.random()));
    }
    return patience;
  };
  /** Looks at whether this side's own work still waits, and for how long it has. */
  const look2 = (): void => {
    if (watch) watch();
    watch = undefined;
    const controller = host.controller;
    if (!controller.dirty) {
      ownSince = undefined;
      standingIn = false;
      return;
    }
    if (ownSince === undefined) {
      ownSince = clock();
      // Another draw for every wait.
      patience = undefined;
    }
    const left = ownSince + watchMs() - clock();
    if (left > 0) {
      watch = later(look2, left);
      return;
    }
    // Nobody saved it in all that time (the one who saves for all has auto-save off, cannot save, or is stuck):
    // this side saves it itself, as any save nobody asked for - conditional on the file it knows.
    const session = host.session();
    if (session && session.state === 'online' && !session.leads && !standingIn) {
      standingIn = true;
      controller.rearm();
    }
  };

  const waiting = (): 'room' | 'offline' | undefined => {
    const session = host.session();
    // On its own (in no room, or apart from it) a workbook is saved as any other.
    if (!session || session.state === 'closed' || session.state === 'detached') return undefined;
    // Who cannot reach the room has steps the room does not have: a file written now would be of no step of it.
    if (session.state !== 'online') return 'offline';
    if (session.leads) return undefined;
    if (standingIn && host.controller.dirty) return undefined;
    return 'room';
  };

  const autoAllowed = (): boolean => waiting() === undefined;

  const localChanged = (): void => {
    if (ownSince === undefined || !watch) look2();
  };

  const remoteChanged = (): void => {
    // Only the one who saves for all starts its delay over a change of someone else.
    const session = host.session();
    const saver = session !== undefined && session.state === 'online' && session.leads && host.controller.autoSave !== undefined;
    host.controller.changed(!saver);
    // Work of others waits here as well: when nobody who can saves it in a while, this side does.
    localChanged();
  };

  const conflictWait = (): Promise<boolean> => {
    const session = host.session();
    if (!session || session.state !== 'online') return Promise.resolve(false);
    return new Promise<boolean>((resolve) => {
      let done = false;
      const stop = (host.timer || ((run, ms) => {
        const id = setTimeout(run, ms);
        return () => clearTimeout(id);
      }))(() => finish(false), CONFLICT_WAIT_MS);
      const finish = (moved: boolean): void => {
        if (done) return;
        done = true;
        stop();
        waiters = waiters.filter((waiter) => waiter !== finish);
        resolve(moved);
      };
      waiters.push(finish);
    });
  };

  let wasOnline = false;
  let awaySince: number | undefined;
  const roomChanged = (): void => {
    const session = host.session();
    const online = session !== undefined && session.state === 'online';
    if (!online && awaySince === undefined) awaySince = clock();
    // Back in the room after a long time away: everybody who waited through it would stand in at once. Each
    // gets a moment of its own within the spread instead - and no more than that: the wait itself is over.
    // After a short break (a connection that comes and goes) the wait goes on where it was; a room that is
    // lost every few seconds does not keep work unsaved for ever.
    if (online && !wasOnline) {
      const away = awaySince === undefined ? 0 : clock() - awaySince;
      if (ownSince !== undefined && away >= baseMs()) {
        patience = undefined;
        ownSince = clock() - baseMs();
        standingIn = false;
        if (watch) watch();
        watch = undefined;
      }
      awaySince = undefined;
    }
    wasOnline = online;
    // Also when nothing is to be sent: why auto-save waits is another thing now.
    host.controller.rearm();
    if (host.controller.dirty) localChanged();
  };

  const ownSaved = (): void => {
    if (watch) watch();
    watch = undefined;
    ownSince = undefined;
    standingIn = false;
    if (host.controller.dirty) look2();
  };

  const stop = (): void => {
    if (watch) watch();
    watch = undefined;
    if (proving) proving.stop();
    proving = undefined;
    answer(false);
  };

  return { saved, sync, storedAs, remoteChanged, localChanged, ownSaved, autoAllowed, waiting, conflictWait, roomChanged, stop };
}
