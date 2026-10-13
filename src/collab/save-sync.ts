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
}

export interface SaveSync {
  /** A participant says the file was saved. Resolves when what was said has been looked into. */
  saved(said: { versions: string[]; rev?: number }): Promise<void>;
  /** The room asks for the file to be saved (someone cannot join). */
  sync(): Promise<boolean>;
  /** Whether the server names the file by this ETag now. */
  storedAs(tag: string): Promise<boolean>;
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

  const saved = async (said: { versions: string[]; rev?: number }): Promise<void> => {
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
    // And never over changes this side made itself: that they are in the file is only what someone says, and a
    // workbook wrongly taken to be saved loses work, while one wrongly taken to be unsaved is saved once more.
    const clean = !stale && !host.controller.ownChanges && said.rev !== undefined && now.rev === said.rev && now.pending === 0;
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

  return { saved, sync, storedAs };
}
