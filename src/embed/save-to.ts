// Saving a file to where it came from: what MountOptions.saveTo stands for,
// and what the bridge page does for a `load` with `save`. The controller knows
// how many changes were made and how many of them the last save holds, sends
// the bytes (an HTTP request, or a function of the application) and says what
// came of it. Nothing here touches the DOM, so that it runs where there is no
// document; the editors only see it as the SaveCommand of their Save button.
import type { SaveCommand } from '../ui/save-command';

/** What a saved file is, for whoever takes its bytes. */
export interface SaveInfo {
  /** The format of the bytes: never one of the formats that cannot be written. */
  format: 'docx' | 'xlsx';
  /** The format the file had when it was opened, when it was converted to be edited. */
  convertedFrom?: 'doc' | 'rtf' | 'xls';
  /** The name of the file, when it is known: the last part of the path it was fetched from, or the name of a File. */
  name?: string;
}

/** An application's own way to save: resolves (or returns) when the bytes are safe, rejects or throws when they are not. */
export type SaveToFunction = (bytes: Uint8Array, info: SaveInfo) => void | Promise<void>;

/** The request a file is saved with. */
export interface SaveToRequest {
  /** Where the bytes are sent. Default: the address the file was fetched from. */
  url?: string;
  /** Default "PUT". */
  method?: 'PUT' | 'POST';
  /** Headers added to the request (a token, say). One that has the name of a header of the engine takes its place. */
  headers?: { [name: string]: string };
  /** Default "same-origin", as the file is fetched: nothing of the user is sent to another origin. */
  credentials?: 'omit' | 'same-origin' | 'include';
  /**
   * How long the request may take, in milliseconds, the upload of the file
   * included. Default 120 000 (two minutes): time enough for a file of some
   * megabytes on a slow connection.
   */
  timeoutMs?: number;
  /**
   * A file that was converted to be edited (doc, RTF, xls) is saved as a docx
   * or an xlsx. Sent to the address the old file was fetched from, that
   * replaces the original with a file of another format, which a server that
   * does not read `X-Bym-Converted-From` cannot know. So it is not done
   * unless the page says here that its server expects it. Default false.
   */
  allowConverted?: boolean;
}

/** What saveNow() can be asked for. */
export interface SaveNowOptions {
  /**
   * Saves over whatever the server holds: the request carries no `If-Match`.
   * For the user who was told that the file on the server is not the one
   * that was opened (SaveError "conflict") and chose to keep their own.
   */
  overwrite?: boolean;
}

/** MountOptions.saveTo: `true` for the address the file was fetched from, an address, a request, or a function. */
export type SaveTo = true | string | SaveToRequest | SaveToFunction;

/** What MountOptions.onSaveState is told. */
export interface SaveState {
  /** There are changes that no save holds. */
  dirty: boolean;
  /** A save is under way. */
  saving: boolean;
  /** When the last save succeeded, in milliseconds since 1970. */
  lastSavedAt?: number;
  /** Why the last save failed, in the language of the view; gone when a save starts or succeeds. */
  error?: string;
}

export type SaveErrorCode =
  /** What is being typed (a note, the name of a sheet, a dialog with changes) could not be applied. */
  | 'pending'
  /** The file could not be written out. */
  | 'write'
  /** HTTP 412: the file on the server is not the one that was opened. */
  | 'conflict'
  /** The server sent the request on to another address (to a login page, when a session has run out). It is not followed. */
  | 'redirected'
  /** The file was converted, and saving it over the original was not allowed (SaveToRequest.allowConverted). */
  | 'converted'
  /** HTTP 401 or 403. */
  | 'denied'
  /** HTTP 423. */
  | 'locked'
  /** No answer in time. */
  | 'timeout'
  /** The request did not get through. */
  | 'network'
  /** Any other status that is not 2xx. */
  | 'http'
  /** The function of the application (or the application behind the bridge) rejected or threw; or the request could not be made at all. */
  | 'refused';

/** What MountOptions.onError is given when a save fails. The message is English and holds nothing of the file. */
export class SaveError extends Error {
  readonly code: SaveErrorCode;
  /** The status of the answer, for the codes that come from one. */
  readonly status?: number;
  /** What the application's function rejected with, for "refused"; what the writer threw, for "write". */
  readonly cause?: unknown;

  constructor(code: SaveErrorCode, message: string, status?: number, cause?: unknown) {
    super(message);
    this.name = 'SaveError';
    this.code = code;
    if (status !== undefined) this.status = status;
    if (cause !== undefined) this.cause = cause;
  }
}

/** The media types of the formats that are written. */
export const SAVE_MEDIA_TYPES: { [format in SaveInfo['format']]: string } = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

/**
 * How long a save request may take unless the page says otherwise. The time
 * covers the upload of the file, not only the wait for the answer.
 */
export const SAVE_TIMEOUT_MS = 120000;
/** How long a function of the application may take to save: long, and an end all the same. */
export const SAVE_FUNCTION_TIMEOUT_MS = 600000;

/** Where a file is saved to, with everything that was left out filled in. */
export type SaveTarget =
  | { kind: 'function'; save: SaveToFunction }
  | {
      kind: 'request';
      /** Absolute, http: or https:. */
      url: string;
      method: 'PUT' | 'POST';
      headers: [name: string, value: string][];
      credentials: 'omit' | 'same-origin' | 'include';
      /** The address is the one the file was fetched from: its ETag is what `If-Match` is sent with. */
      sameAsFile: boolean;
      /**
       * The file was not fetched from an address, and the one it is saved to ends in the extension of a
       * format that is not written (.doc, .rtf, .xls): a converted file sent there would take the place
       * of an original, as at the address of the file itself. Only there when it holds.
       */
      replacesOriginal?: boolean;
      timeoutMs: number;
      /** A converted file may be saved to the address of its original. */
      allowConverted: boolean;
    };

/** RFC 9110 `token`: what the name of a header is made of. */
const HEADER_NAME = /^[!#$%&'*+\-.^_`|~0-9A-Za-z]+$/;
/**
 * What fetch takes as the value of a header: a ByteString (nothing above U+00FF) without control
 * characters but the tab. A line break would start another header; a letter outside Latin-1 makes
 * fetch throw at every save.
 */
const HEADER_VALUE = /^[\u0009\u0020-\u007E\u0080-\u00FF]*$/;

/** An http(s) address as an absolute one without its fragment; undefined for anything else. */
function httpAddress(text: string, base: string | undefined): string | undefined {
  let url: URL;
  try {
    url = base === undefined ? new URL(text) : new URL(text, base);
  } catch {
    return undefined;
  }
  // After parsing: the parser drops the white space and control characters a scheme can hide behind.
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return undefined;
  // fetch refuses an address with a user or a password in it.
  if (url.username !== '' || url.password !== '') return undefined;
  url.hash = '';
  return url.href;
}

/**
 * What tells one http(s) address from another for a server: origin, path and query. The fragment is
 * never sent, and an empty query ("?") is no query. The case of the path and how it is encoded are
 * kept as they are: a server may tell them apart. Undefined for what is not an address.
 */
function addressKey(text: string): string | undefined {
  try {
    const url = new URL(text);
    return url.origin + url.pathname + url.search;
  } catch {
    return undefined;
  }
}

/** Whether the path of an address ends as the name of a file of a format that is not written. */
function endsInOldFormat(address: string): boolean {
  try {
    return /\.(doc|rtf|xls)$/i.test(new URL(address).pathname);
  } catch {
    return false;
  }
}

/**
 * Reads MountOptions.saveTo. `fileUrl` is the address the file was fetched
 * from, when it was; `base` is the address of the page, which a relative
 * address is resolved against. Throws TypeError for what cannot be saved to:
 * an address that is not http: or https:, `true` or a request without `url`
 * for a file that was not fetched from one, a method other than PUT and POST.
 */
export function resolveSaveTarget(saveTo: unknown, fileUrl: string | undefined, base: string | undefined): SaveTarget {
  if (typeof saveTo === 'function') return { kind: 'function', save: saveTo as SaveToFunction };
  const fetched = fileUrl === undefined ? undefined : httpAddress(fileUrl, base);
  let given: string | undefined;
  let request: SaveToRequest = {};
  if (saveTo === true) given = undefined;
  else if (typeof saveTo === 'string') given = saveTo;
  else if (typeof saveTo === 'object' && saveTo !== null && !Array.isArray(saveTo)) {
    request = saveTo as SaveToRequest;
    if (request.url !== undefined && typeof request.url !== 'string') throw new TypeError('"saveTo.url" must be an address');
    given = request.url;
  } else throw new TypeError('"saveTo" must be true, an address, a request ({ url, method, headers, credentials }) or a function');

  let url: string | undefined;
  if (given === undefined) {
    if (fetched === undefined) throw new TypeError('"saveTo" needs an address: the file was not fetched from an http(s) address it could be saved to');
    url = fetched;
  } else {
    // An empty address would resolve to the page itself: that is a value that went missing, not a place to save to.
    url = given.trim() === '' ? undefined : httpAddress(given, base);
    if (url === undefined) throw new TypeError('"saveTo" must be an http(s) address, or one relative to the page');
  }

  const method = request.method === undefined ? 'PUT' : request.method;
  if (method !== 'PUT' && method !== 'POST') throw new TypeError('"saveTo.method" must be "PUT" or "POST"');
  const credentials = request.credentials === undefined ? 'same-origin' : request.credentials;
  if (credentials !== 'omit' && credentials !== 'same-origin' && credentials !== 'include') throw new TypeError('"saveTo.credentials" must be "omit", "same-origin" or "include"');
  const headers: [string, string][] = [];
  if (request.headers !== undefined) {
    if (typeof request.headers !== 'object' || request.headers === null || Array.isArray(request.headers)) throw new TypeError('"saveTo.headers" must be an object of header names and values');
    for (const name of Object.keys(request.headers)) {
      const value = request.headers[name];
      // "__proto__" would be a token, and as the key of an object it is the prototype: it could only be lost.
      if (name === '__proto__' || !HEADER_NAME.test(name) || typeof value !== 'string' || !HEADER_VALUE.test(value)) {
        throw new TypeError('"saveTo.headers" holds a name or a value that a header cannot have (values are Latin-1 text without line breaks)');
      }
      headers.push([name, value]);
    }
  }
  const timeoutMs = request.timeoutMs === undefined ? SAVE_TIMEOUT_MS : request.timeoutMs;
  if (typeof timeoutMs !== 'number' || !(timeoutMs > 0) || timeoutMs === Infinity) throw new TypeError('"saveTo.timeoutMs" must be a positive number of milliseconds');
  if (request.allowConverted !== undefined && typeof request.allowConverted !== 'boolean') throw new TypeError('"saveTo.allowConverted" must be true or false');
  // The same address written another way ("?" at its end, a fragment, the port spelled out) is the same file.
  const sameAsFile = fetched !== undefined && addressKey(url) === addressKey(fetched);
  const target: SaveTarget = { kind: 'request', url, method, headers, credentials, sameAsFile, timeoutMs, allowConverted: request.allowConverted === true };
  // A file given as bytes has no address of its own to compare with: the name of the address is what is left to go by.
  if (fetched === undefined && endsInOldFormat(url)) target.replacesOriginal = true;
  return target;
}

/**
 * An ETag that `If-Match` can be sent with: a strong one. A weak one ("W/…")
 * never matches there (RFC 9110, 13.1.1), so a save with it would always be refused.
 */
export function strongETag(value: string | null | undefined): string | undefined {
  return typeof value === 'string' && /^"[!#-~\u0080-ÿ]*"$/.test(value) ? value : undefined;
}

/** The last part of the path of an address, decoded: the name a file is known by. Undefined when there is none. */
export function nameFromUrl(address: string, base: string | undefined): string | undefined {
  try {
    const url = base === undefined ? new URL(address) : new URL(address, base);
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return undefined;
    const last = url.pathname.split('/').pop() || '';
    const name = decodeURIComponent(last);
    return name === '' ? undefined : name;
  } catch {
    return undefined;
  }
}

/**
 * The ETag a fetched file came with, for the saves to the address it was asked from. None when
 * the answer came from another address (a redirect): that ETag is not of the address saved to.
 */
export function fileETag(response: { redirected?: boolean; headers: { get(name: string): string | null } }): string | null {
  return response.redirected === true ? null : response.headers.get('etag');
}

/** What a saved file is: its format, what it was converted from, and its name with the extension of what it is now. */
export function saveInfoOf(format: string, conversion: { from: 'doc' | 'rtf' | 'xls'; to: string } | undefined, name: string | undefined): SaveInfo {
  const info: SaveInfo = { format: format as SaveInfo['format'] };
  if (conversion) info.convertedFrom = conversion.from;
  // "report.doc" that became a docx is "report.docx"; any other end of a name is kept.
  if (name !== undefined) info.name = conversion ? `${name.replace(/\.(doc|rtf|xls)$/i, '')}.${conversion.to}` : name;
  return info;
}

/** Whether what a document's protection enforces keeps it from being written: everything but "tracked changes only". */
export function locksSaving(protection: string | undefined): boolean {
  return protection !== undefined && protection !== 'trackedChanges';
}

/**
 * Asks before the page is left while `dirty()` holds (the browser's own question). The result
 * takes the listener away again; safe to call twice.
 */
export function watchLeaving(
  win: { addEventListener(type: 'beforeunload', listener: (event: BeforeUnloadEvent) => void): void; removeEventListener(type: 'beforeunload', listener: (event: BeforeUnloadEvent) => void): void },
  dirty: () => boolean,
): () => void {
  const onLeave = (event: BeforeUnloadEvent) => {
    if (!dirty()) return;
    event.preventDefault();
    event.returnValue = '';
  };
  win.addEventListener('beforeunload', onLeave);
  return () => win.removeEventListener('beforeunload', onLeave);
}

/** The words a save is spoken of with: part of the view's strings. */
export interface SaveStrings {
  saved: string;
  saveFailed(detail: string): string;
  saveConflict: string;
  saveRedirected: string;
  saveConvertedNotAllowed(to: string): string;
  saveDenied: string;
  saveLocked: string;
  saveTimedOut: string;
  saveNoConnection: string;
  savePending: string;
}

/** What a controller needs of whoever holds the file. */
export interface SaveHost {
  target: SaveTarget;
  /** Whether the file can be written with its changes: a docx or an xlsx that is not locked. */
  writable(): boolean;
  /**
   * Applies what is being typed, as leaving it would. False when something
   * stays open that could not be applied: nothing is saved then, and whoever
   * answers false has said why.
   */
  commit(): boolean;
  /** The file as it is now. The model is taken before the first await. */
  bytes(): Promise<Uint8Array>;
  info(): SaveInfo;
  strings(): SaveStrings;
  /** Says something to the user. */
  notice?(message: string, kind: 'error' | 'info'): void;
  onState?(state: SaveState): void;
  /** Called with the bytes of every save that succeeded. */
  onSaved?(bytes: Uint8Array): void;
  /** Called with a SaveError for every save that failed. */
  onError?(error: unknown): void;
  /** The ETag the file came with, when it was fetched. */
  etag?: string | null;
  /** Default: the global fetch. */
  fetch?: typeof fetch;
  /** Default: Date.now. */
  now?(): number;
  /** Runs a callback after a delay and returns what calls it off. Default: setTimeout. */
  timer?(callback: () => void, ms: number): () => void;
  /** Instead of the time limit of the target (tests). */
  timeoutMs?: number;
}

const defaultTimer = (callback: () => void, ms: number): (() => void) => {
  const handle = setTimeout(callback, ms);
  return () => clearTimeout(handle);
};

/** What an application's failure says, as one short line of text. */
function detailOf(error: unknown): string {
  let text = '';
  if (error instanceof Error) text = error.message;
  else if (typeof error === 'string') text = error;
    text = text.replace(/[\u0000-\u001F\u007F]+/g, ' ').trim();
  return text.length > 200 ? `${text.slice(0, 200)}…` : text;
}

/** A body that is not read is let go of, so that the connection is free again. */
function drop(response: Response): void {
  try {
    const body = response.body;
    if (body && typeof body.cancel === 'function') void body.cancel().catch(() => undefined);
  } catch {
    // A body that cannot be cancelled is left to the browser.
  }
}

export class SaveController implements SaveCommand {
  private readonly host: SaveHost;
  /** Counts the changes: the same number means the same file. */
  private count = 0;
  /** The count the last save that succeeded holds. */
  private savedCount = 0;
  /** The count the save under way took the file at. */
  private held = 0;
  private busy = false;
  private lastSavedAt: number | undefined;
  private error: string | undefined;
  private etag: string | undefined;
  private running: Promise<boolean> | undefined;
  /** The one save that waits for the one under way: it holds what was changed since that one took the file. */
  private next: Promise<boolean> | undefined;
  /** The save that waits was asked to overwrite. */
  private nextOverwrites = false;
  private abort: (() => void) | undefined;
  /** The host is being told of a change it is in the middle of making (see changed()). */
  private changing = 0;
  /** A call is being told that no save could start: one that comes from inside that telling is not told again. */
  private refusing = false;
  private readonly listeners: (() => void)[] = [];
  private told = '';
  private destroyed = false;

  constructor(host: SaveHost) {
    this.host = host;
    this.etag = strongETag(host.etag);
  }

  get dirty(): boolean {
    return this.count !== this.savedCount;
  }

  get saving(): boolean {
    return this.busy;
  }

  get state(): SaveState {
    const state: SaveState = { dirty: this.dirty, saving: this.busy };
    if (this.lastSavedAt !== undefined) state.lastSavedAt = this.lastSavedAt;
    if (this.error !== undefined) state.error = this.error;
    return state;
  }

  /**
   * Whether saving is held back because the file was converted and would replace its original
   * (SaveToRequest.allowConverted): a save then fails with "converted" before anything is sent.
   */
  get blockedConversion(): boolean {
    const target = this.host.target;
    return target.kind === 'request' && (target.sameAsFile || target.replacesOriginal === true) && !target.allowConverted && this.host.info().convertedFrom !== undefined;
  }

  /** The ETag the next save is sent with, when there is one. */
  get currentETag(): string | undefined {
    return this.etag;
  }

  /** Says that the file was changed (an undo and a redo are changes as well). */
  changed(): void {
    if (this.destroyed) return;
    this.count++;
    // Whoever calls is still making the change (a cell that is being written has its editor open until
    // this returns): a save asked for from inside onState waits until the change is over.
    this.changing++;
    try {
      this.emit();
    } finally {
      this.changing--;
    }
  }

  run(): void {
    void this.saveNow();
  }

  watch(listener: () => void): () => void {
    this.listeners.push(listener);
    return () => {
      const at = this.listeners.indexOf(listener);
      if (at >= 0) this.listeners.splice(at, 1);
    };
  }

  /**
   * Saves the file as it is now. Resolves to true when the target holds it,
   * also when there was nothing to save; to false when it could not be saved,
   * which the notice, the state and onError tell about. Never rejects. A call
   * while a save is under way sends nothing when that save holds everything,
   * and otherwise waits for it and saves once more. With `overwrite` the
   * request carries no `If-Match`; such a call always waits for a save that
   * is under way, and then saves if there is still something to save.
   *
   * Safe to call from inside any callback of the host (onState, notice, onError, onSaved and the
   * rest): the save under way is known before the first of them is called, so that such a call
   * joins it or waits for it and never sends the file beside it. A call from inside the onState
   * of a change that is being made saves as soon as that change is over.
   */
  saveNow(options?: SaveNowOptions): Promise<boolean> {
    const overwrite = options !== undefined && options !== null && options.overwrite === true;
    // From inside the notification of a change: the file is taken once the change is made, not half way through it.
    if (this.changing > 0 && !this.destroyed) return Promise.resolve().then(() => this.saveNow(options));
    if (this.destroyed || !this.host.writable()) return Promise.resolve(false);
    // Asked from inside onState or onError while they hear that no save could start: they have heard.
    if (this.refusing) return Promise.resolve(false);
    let settled = false;
    try {
      settled = this.host.commit();
    } catch {
      settled = false;
    }
    if (!settled) {
      // Why it could not be applied was said where it happened; here only the state and onError learn of it.
      // A save that is under way goes on: when it succeeds, this is no longer what the state says.
      this.refuse(() => {
        this.error = this.host.strings().savePending;
        this.emit();
        this.report(new SaveError('pending', 'What is being typed could not be applied; the file was not saved'));
      });
      return Promise.resolve(false);
    }
    // The last question to the host. From here to the point where the save under way is known nothing
    // of the host is called: a callback that asks for a save cannot find the controller half way.
    const blocked = this.blockedConversion;
    if (this.running) {
      if (this.count === this.held && !overwrite) return this.running;
      if (overwrite) this.nextOverwrites = true;
      if (!this.next) {
        this.next = this.running.then(() => {
          this.next = undefined;
          const over = this.nextOverwrites;
          this.nextOverwrites = false;
          return this.saveNow(over ? { overwrite: true } : undefined);
        });
      }
      return this.next;
    }
    if (!this.dirty) return Promise.resolve(true);
    if (blocked) {
      this.refuse(() => void this.failed(new SaveError('converted', 'The file was converted; saving it over the original needs saveTo.allowConverted')));
      return Promise.resolve(false);
    }
    let settle: (ok: boolean) => void = () => undefined;
    const running = new Promise<boolean>((resolve) => {
      settle = resolve;
    });
    // Before anything is said: onState hears "saving" next, and may ask for a save from there.
    this.running = running;
    this.held = this.count;
    this.busy = true;
    this.error = undefined;
    const end = (ok: boolean): void => {
      if (this.running === running) this.running = undefined;
      settle(ok);
    };
    this.start(overwrite).then(end, () => {
      // A callback of the host threw where nothing catches it: the save has an end all the same.
      this.busy = false;
      try {
        this.emit();
      } catch {
        // As above.
      }
      end(false);
    });
    return running;
  }

  /** Tells of a save that could not start; a call of saveNow() from inside the telling answers false and tells nothing. */
  private refuse(tell: () => void): void {
    this.refusing = true;
    try {
      tell();
    } finally {
      this.refusing = false;
    }
  }

  /** Gives up the request under way and stops speaking. Safe to call twice. */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.listeners.length = 0;
    const abort = this.abort;
    this.abort = undefined;
    if (abort) abort();
  }

  private async start(overwrite: boolean): Promise<boolean> {
    const host = this.host;
    // saveNow() has marked the save as under way; this is the first the host hears of it.
    this.emit();
    let bytes: Uint8Array;
    try {
      // Before the file is taken: what is changed from here on is not in these bytes. (Taken again
      // here, after onState was told: what a callback changed there is in them.)
      this.held = this.count;
      bytes = await host.bytes();
    } catch (error) {
      if (this.destroyed) return false;
      return this.failed(new SaveError('write', 'The file could not be written', undefined, error));
    }
    if (this.destroyed) return false;
    try {
      const target = host.target;
      const info = host.info();
      if (target.kind === 'function') await this.call(target.save, bytes, info);
      else await this.request(target, bytes, info, overwrite);
    } catch (error) {
      if (this.destroyed) return false;
      if (error instanceof SaveError) return this.failed(error);
      const detail = detailOf(error);
      return this.failed(new SaveError('refused', detail === '' ? 'The file was not taken' : detail, undefined, error), detail);
    }
    if (this.destroyed) return false;
    this.savedCount = this.held;
    this.busy = false;
    // Also what was said meanwhile of something typed that could not be applied: the state does not
    // say "saved" and "could not be saved" at once.
    this.error = undefined;
    this.lastSavedAt = host.now ? host.now() : Date.now();
    this.emit();
    // With changes made meanwhile there is still something to save: "saved" would say otherwise.
    if (!this.dirty) this.say(host.strings().saved, 'info');
    if (host.onSaved) {
      try {
        host.onSaved(bytes);
      } catch (error) {
        this.report(error);
      }
    }
    return true;
  }

  /** Lets the application's function save, and gives up waiting for it when it takes too long. */
  private call(save: SaveToFunction, bytes: Uint8Array, info: SaveInfo): Promise<void> {
    const host = this.host;
    const timeout = host.timeoutMs === undefined ? SAVE_FUNCTION_TIMEOUT_MS : host.timeoutMs;
    return new Promise<void>((resolve, reject) => {
      let over = false;
      const stop = (host.timer || defaultTimer)(() => {
        if (over) return;
        over = true;
        reject(new SaveError('timeout', `The application did not finish saving within ${Math.round(timeout / 1000)} s`));
      }, timeout);
      const end = (error?: { reason: unknown }) => {
        if (over) return;
        over = true;
        stop();
        if (this.abort === give) this.abort = undefined;
        if (error) reject(error.reason);
        else resolve();
      };
      // The view is destroyed: the function is no longer waited for, and its time limit is called off.
      const give = () => end({ reason: new SaveError('network', 'The save was given up') });
      this.abort = give;
      let job: void | Promise<void>;
      try {
        job = save(bytes, info);
      } catch (error) {
        return end({ reason: error });
      }
      Promise.resolve(job).then(
        () => end(),
        (reason) => end({ reason }),
      );
    });
  }

  /** `detail`: what the application said of its failure, one line; shown with "refused" only. */
  private failed(error: SaveError, detail = ''): boolean {
    const strings = this.host.strings();
    let text: string;
    switch (error.code) {
      case 'conflict':
        text = strings.saveConflict;
        break;
      case 'denied':
        text = strings.saveDenied;
        break;
      case 'locked':
        text = strings.saveLocked;
        break;
      case 'timeout':
        text = strings.saveTimedOut;
        break;
      case 'network':
        text = strings.saveNoConnection;
        break;
      case 'http':
        text = strings.saveFailed(`HTTP ${error.status}`);
        break;
      case 'redirected':
        text = strings.saveRedirected;
        break;
      case 'converted':
        text = strings.saveConvertedNotAllowed(this.host.info().format);
        break;
      case 'refused':
        text = strings.saveFailed(detail);
        break;
      default:
        text = strings.saveFailed('');
    }
    this.busy = false;
    this.error = text;
    this.emit();
    this.say(text, 'error');
    this.report(error);
    return false;
  }

  /** Sends the bytes; resolves for a 2xx answer and rejects with a SaveError for anything else. */
  private request(target: Extract<SaveTarget, { kind: 'request' }>, bytes: Uint8Array, info: SaveInfo, overwrite: boolean): Promise<void> {
    const host = this.host;
    const send = host.fetch || (typeof fetch === 'function' ? fetch : undefined);
    if (!send) return Promise.reject(new SaveError('network', 'There is no fetch to send the file with'));
    const headers: { [name: string]: string } = {};
    const set = (name: string, value: string) => {
      const lower = name.toLowerCase();
      for (const other of Object.keys(headers)) if (other.toLowerCase() === lower) delete headers[other];
      headers[name] = value;
    };
    set('Content-Type', SAVE_MEDIA_TYPES[info.format]);
    set('X-Bym-Format', info.format);
    if (info.convertedFrom !== undefined) set('X-Bym-Converted-From', info.convertedFrom);
    // Only where the file came from: the ETag says nothing about what is at another address.
    // Not when the user chose to save over what the server holds.
    if (target.sameAsFile && this.etag !== undefined && !overwrite) set('If-Match', this.etag);
    for (let i = 0; i < target.headers.length; i++) {
      // An overwrite carries no If-Match at all, not the page's own either.
      if (overwrite && target.headers[i][0].toLowerCase() === 'if-match') continue;
      set(target.headers[i][0], target.headers[i][1]);
    }

    const controller = typeof AbortController === 'function' ? new AbortController() : undefined;
    // The view itself: only its part of the buffer is sent.
    // A redirect is never followed: a server whose session ran out answers a save with one to its login
    // page, and following it would end in "200" with nothing saved, and could carry the file and the
    // headers of the page to another origin. A browser hands back an opaque answer (status 0), Node the 3xx itself.
    const init: RequestInit = { method: target.method, headers, body: bytes as unknown as BodyInit, credentials: target.credentials, redirect: 'manual' };
    if (controller) init.signal = controller.signal;
    const timer = host.timer || defaultTimer;
    const timeout = host.timeoutMs === undefined ? target.timeoutMs : host.timeoutMs;

    return new Promise<void>((resolve, reject) => {
      let over = false;
      let stop: (() => void) | undefined;
      const end = (error?: SaveError) => {
        if (over) return;
        over = true;
        if (stop) stop();
        if (this.abort === give) this.abort = undefined;
        if (error) reject(error);
        else resolve();
      };
      // Without AbortController the request runs on, and its answer is ignored.
      const give = () => {
        if (over) return;
        if (controller) controller.abort();
        end(new SaveError('network', 'The save was given up'));
      };
      this.abort = give;
      stop = timer(() => {
        if (over) return;
        if (controller) controller.abort();
        end(new SaveError('timeout', `No answer within ${Math.round(timeout / 1000)} s`));
      }, timeout);
      let answer: Promise<Response>;
      try {
        answer = send(target.url, init);
      } catch (error) {
        // fetch throws at once for a request it will not make (a header it does not take): no server was asked.
        return end(new SaveError('refused', `The request could not be made: ${detailOf(error)}`, undefined, error));
      }
      answer.then(
        (response) => {
          // The body is never read: an answer of any size costs no memory.
          drop(response);
          if (over) return;
          const status = response.status;
          // A fetch that does not know `redirect` (one made of XMLHttpRequest, in an old WebView) follows by itself and
          // says nothing of it: the address its answer came from is then all that tells a login page from the file.
          const from = typeof response.url === 'string' && response.url !== '' ? addressKey(response.url) : undefined;
          const elsewhere = from !== undefined && from !== addressKey(target.url);
          if (status === 0 || (status >= 300 && status < 400) || response.type === 'opaqueredirect' || response.redirected === true || elsewhere) {
            return end(new SaveError('redirected', 'The server redirected the save to another address; it was not followed', status === 0 ? undefined : status));
          }
          if (status >= 200 && status < 300) {
            // What the server holds now. Without a new ETag the old one names a file that is gone:
            // sending it again would have every later save refused.
            if (target.sameAsFile) this.etag = strongETag(response.headers && typeof response.headers.get === 'function' ? response.headers.get('etag') : undefined);
            return end();
          }
          if (status === 412) return end(new SaveError('conflict', 'The file on the server is not the one that was opened (HTTP 412)', status));
          if (status === 401 || status === 403) return end(new SaveError('denied', `Saving is not allowed (HTTP ${status})`, status));
          if (status === 423) return end(new SaveError('locked', 'The file is locked (HTTP 423)', status));
          end(new SaveError('http', `The server answered HTTP ${status}`, status));
        },
        // What a browser says of a request that failed is not for the user, and differs from one to the next.
        () => end(new SaveError('network', 'The request did not get through')),
      );
    });
  }

  private say(message: string, kind: 'error' | 'info'): void {
    if (this.destroyed || !this.host.notice) return;
    try {
      this.host.notice(message, kind);
    } catch {
      // A page that fails at showing a notice does not undo the save.
    }
  }

  private report(error: unknown): void {
    if (this.destroyed || !this.host.onError) return;
    try {
      this.host.onError(error);
    } catch {
      // As above.
    }
  }

  private emit(): void {
    if (this.destroyed) return;
    const state = this.state;
    const key = `${state.dirty} ${state.saving} ${state.lastSavedAt} ${state.error}`;
    if (key === this.told) return;
    this.told = key;
    if (this.host.onState) {
      try {
        this.host.onState(state);
      } catch {
        // As above.
      }
    }
    const listeners = this.listeners.slice();
    for (let i = 0; i < listeners.length; i++) listeners[i]();
  }
}
