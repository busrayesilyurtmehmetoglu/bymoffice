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
  readonly autoSave: unknown;
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
  target: { url: string; credentials: RequestCredentials } | undefined;
  /** An ETag that `If-Match` can be sent with, or undefined. */
  strong(etag: string | null | undefined): string | undefined;
  readOnly(): boolean;
  fetch?: typeof fetch;
  /** Calls back once after `ms`; answers what cancels it. */
  timer?(run: () => void, ms: number): () => void;
}

/** How long a save that met a conflict waits for word of the save that made it. */
export const CONFLICT_WAIT_MS = 2500;

export interface SaveSync {
  /** A participant says the file was saved. Resolves when what was said has been looked into. */
  saved(said: { versions: string[]; rev?: number; byLeader?: boolean }): Promise<void>;
  /** The room asks for the file to be saved (someone cannot join). */
  sync(): Promise<boolean>;
  /** Whether the server names the file by this ETag now. */
  storedAs(tag: string): Promise<boolean>;
  /** A change arrived from someone else: there is something to save, and the one who saves for all is to save it. */
  remoteChanged(): void;
  /** Whether this side is the one to save by itself now (SaveHost.autoAllowed). */
  autoAllowed(): boolean;
  /** A save here met a conflict (SaveHost.conflictWait): true when a save of someone in the room explains it and was found to be so. */
  conflictWait(): Promise<boolean>;
  /** Who is in the room changed, or the state of the session: the one who saves for all may be this side now. */
  roomChanged(): void;
}

export function createSaveSync(host: SaveSyncHost): SaveSync {
  const ask = host.fetch || (typeof fetch !== 'undefined' ? fetch : undefined);

  const storedAs = async (tag: string): Promise<boolean> => {
    const target = host.target;
    if (!target || !ask || host.strong(tag) === undefined) return false;
    const request = (method: string, headers?: { [name: string]: string }): Promise<Response> => ask(target.url, { method, cache: 'no-store', credentials: target.credentials, headers });
    try {
      let answer = await request('HEAD');
      if (answer.status === 405 || answer.status === 501) {
        // A server without HEAD: the file is asked for only if it is another one, and its body is not read.
        answer = await request('GET', { 'If-None-Match': tag });
        const same = answer.status === 304 || (answer.ok && host.strong(answer.headers.get('ETag')) === tag);
        if (answer.body) void answer.body.cancel().catch(() => undefined);
        return same;
      }
      return answer.ok && host.strong(answer.headers.get('ETag')) === tag;
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

  const saved = async (said: { versions: string[]; rev?: number; byLeader?: boolean }): Promise<void> => {
    checking++;
    const before = host.controller.currentETag;
    try {
      await look(said);
    } finally {
      checking--;
      // Whoever waited for this hears whether it moved what the file is known by.
      if (host.controller.currentETag !== before) answer(true);
      else if (checking === 0) answer(false);
    }
  };

  const look = async (said: { versions: string[]; rev?: number; byLeader?: boolean }): Promise<void> => {
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
    if (!stored) return host.controller.hint(tag);
    // Saved here as well only when this workbook is, right now, the workbook of the step that was saved.
    // Changes this side made itself are in the file on the word of one participant only: the one who saves for all
    // (it is the one this side leaves its saving to). On anybody else's word they stay to be saved: a workbook wrongly
    // taken to be saved loses work, one wrongly taken to be unsaved is saved once more.
    const clean = !stale && (!host.controller.ownChanges || said.byLeader === true) && said.rev !== undefined && now.rev === said.rev && now.pending === 0;
    host.controller.adopt(tag, clean);
  };

  const sync = async (): Promise<boolean> => {
    const session = host.session();
    // One participant saves for the others; a workbook that is only read, or has steps on their way, does not.
    if (!session || session.state !== 'online' || !session.leads || session.pending > 0 || host.readOnly()) return false;
    if (!host.controller.dirty) host.controller.changed(true);
    // As a save nobody asked for: what is being typed is not applied for it.
    return host.controller.saveQuietly();
  };

  const autoAllowed = (): boolean => {
    const session = host.session();
    // On its own (in no room, or apart from it) a workbook is saved as any other. In a room one participant saves
    // for all; who cannot reach the room does not know who that is, and holds still.
    if (!session || session.state === 'closed' || session.state === 'detached') return true;
    return session.state === 'online' && session.leads;
  };

  const remoteChanged = (): void => {
    // Only the one who saves for all starts its delay over a change of someone else.
    host.controller.changed(!(autoAllowed() && host.controller.autoSave !== undefined));
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

  const roomChanged = (): void => {
    if (autoAllowed()) host.controller.rearm();
  };

  return { saved, sync, storedAs, remoteChanged, autoAllowed, conflictWait, roomChanged };
}
