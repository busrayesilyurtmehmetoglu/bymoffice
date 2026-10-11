// The WebView bridge: lets a native application drive the editor with JSON
// messages. All messages travel as strings in both directions, so that every
// platform (Flutter, Android, iOS, Windows, React Native, an iframe) uses the
// same contract. Nothing here touches the DOM when the module is imported;
// the window and the container are handed in.
import { createBlank } from '../formats/blank';
import { needsConversion, openConverted } from '../formats/convert';
import { documentPictures } from '../model/doc/pictures';
import type { OfficeFile } from '../index';
import { enforcedProtection, hasPermissionRanges, tracksChanges, type DocModel } from '../model/doc/types';
import { resolveColor } from '../model/sheet/color';
import { cellName, rangeName, type Workbook } from '../model/sheet/types';
import { registerFont, sanitizeFontFamily, settleFonts, type FontFailure, type FontRegistration } from '../render/fonts';
import type { DocumentEditorOptions } from '../ui/doc-editor';
import type { DocumentViewOptions } from '../ui/doc-view';
import type { SpreadsheetEditorOptions } from '../ui/editor';
import type { SheetPrintOptions } from '../ui/sheet-print';
import type { WorkbookView, WorkbookViewOptions } from '../ui/sheet-view';
import { resolveUiStrings } from '../ui/strings';
import { SaveController, SaveError, type SaveInfo } from './save-to';

export type BridgeErrorCode =
  | 'unsupported'
  | 'corrupt'
  | 'load-failed'
  | 'save-failed'
  | 'bad-message'
  | 'not-loaded'
  | 'print-unavailable'
  | 'print-failed'
  | 'read-only'
  | 'no-sheet'
  | 'sheet-refused'
  | 'no-document';

/** What the page shows around the editor: nothing open, a file on its way, a file open, a file that failed. */
export type BridgeState = 'empty' | 'loading' | 'open' | 'error';

/**
 * What the bridge needs of a document view. Only destroy() is required: a
 * view without the editing members shows the document and nothing more, and
 * the commands that need them are confirmed without an effect.
 */
export interface BridgeDocumentView {
  destroy(): void;
  setLocale?(locale: string): void;
  /** Turns editing off or on. */
  setReadOnly?(readOnly: boolean): void;
  undo?(): unknown;
  redo?(): unknown;
  readonly canUndo?: boolean;
  readonly canRedo?: boolean;
  /** Ends what an input method is composing, so that it is part of the saved file. */
  commitEditing?(): void;
  /** Counts the changes made in the view: the same number means the same document. */
  readonly changeCount?: number;
  /** Says that the document as it is now has been saved. */
  markSaved?(): void;
  /**
   * Prints the document through the window's print dialog; resolves when the
   * dialog was closed. An error with `code: 'unavailable'` says that the window cannot print.
   */
  print?(options?: { dpi?: number }): Promise<unknown>;
  /** The tracked changes of the document: their number is what the application is told. */
  revisions?(): unknown[];
  /** Accept or reject every tracked change; a view that cannot (a protection) changes nothing. */
  acceptAllRevisions?(): unknown;
  rejectAllRevisions?(): unknown;
  /** Turns the tracking of changes on or off; a view that cannot changes nothing. */
  setTrackChanges?(on: boolean): unknown;
  /** Whether the changes made in the view are tracked. */
  readonly trackingChanges?: boolean;
  /** Says something to the user: that the document was saved, or why it was not. */
  showNotice?(message: string, kind?: 'error' | 'info'): void;
}

export interface BridgeHost {
  /** The element the editor is shown in. */
  container: HTMLElement;
  /** The engine's open(). */
  open(bytes: Uint8Array): Promise<OfficeFile>;
  /** Shows a workbook in the container: `new WorkbookView(container, workbook, options)`. */
  createView(container: HTMLElement, workbook: Workbook, options: WorkbookViewOptions): WorkbookView;
  /**
   * Shows a document in the container: `new DocumentView(container, doc, options)`.
   * Without it documents are refused as unsupported.
   */
  createDocumentView?(container: HTMLElement, doc: DocModel, options: DocumentViewOptions): BridgeDocumentView;
  /**
   * Shows a workbook with the editor's shell around the grid (toolbar, formula
   * bar, status bar), for a `load` with `toolbar: true`, and returns the grid.
   * Without it such a load is shown with createView().
   */
  createEditor?(container: HTMLElement, workbook: Workbook, options: SpreadsheetEditorOptions): WorkbookView;
  /**
   * Shows a document with the formatting toolbar above the pages, for a `load`
   * with `toolbar: true`. Without it such a load is shown with createDocumentView().
   */
  createDocumentEditor?(container: HTMLElement, doc: DocModel, options: DocumentEditorOptions): BridgeDocumentView;
  /**
   * Hands one message to the application. Returns false (or throws) when there
   * is no channel to send it on yet; the message is then kept and sent later.
   */
  post(json: string): boolean | void;
  /** Sent in `ready`; the package version, which is also the version of the contract. */
  version: string;
  /** The page's own address. A `load` with a `url` is resolved against it; without it only blob: and data: URLs load. */
  baseUrl?: string;
  /** False refuses every `load` with a `url` (a framed page). Default true. */
  allowUrl?: boolean;
  /**
   * Fetches the file of a `load` with a `url`; must give up when the signal
   * aborts and fail when the file is larger than `maxBytes`. Without it such
   * a load fails with `load-failed`.
   */
  fetchBytes?(url: string, limits: { signal?: AbortSignal; maxBytes: number }): Promise<Uint8Array>;
  /** How long a fetch may take. Default 30 s. */
  fetchTimeoutMs?: number;
  /** Largest file a fetch may return. Default 256 MB. */
  maxFetchBytes?: number;
  /** Characters of undelivered messages kept before `saved` messages are given up. Default 64 M. */
  maxPendingChars?: number;
  /** How long the page waits for the `saveResult` of a save the user started. Default 60 s. */
  saveResultTimeoutMs?: number;
  /**
   * Registers a font of a `load` with `fonts`; resolves when the font can be
   * used. Default: the engine's registerFont().
   */
  registerFont?(font: FontRegistration): Promise<unknown>;
  /** Called when what the page should show around the editor changes. `code` comes with 'error'. */
  onState?(state: BridgeState, code?: BridgeErrorCode): void;
  /** Called with the language tag of a successful `load` or `setLocale`. */
  onLocale?(locale: string): void;
  /** Runs a callback on the next frame; selection messages are sent from it. Default: a 16 ms timer. */
  frame?(callback: () => void): void;
  /** Runs a callback after a delay. Default: setTimeout. */
  timer?(callback: () => void, ms: number): void;
}

export interface Bridge {
  /**
   * Takes one message from the application: the JSON text, or the object it
   * stands for. Resolves when the message has been dealt with; never rejects.
   */
  receive(message: unknown): Promise<void>;
  /** Announces the page with `ready`. Called by receive() too; a second call does nothing. */
  start(): void;
  /** Closes the file, gives up a load that is under way and stops answering. */
  destroy(): void;
}

const ENCODE = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
/** Input bytes turned into text per step: a multiple of 3, and few enough characters for one fromCharCode call. */
const ENCODE_STEP = 3 * 4096;

const SKIP = 64;
const PAD = 65;
const INVALID = 255;
/** Character code -> 6-bit value, SKIP for white space, PAD for "=", INVALID for the rest. */
const DECODE = /* @__PURE__ */ (() => {
  const table = new Uint8Array(128).fill(INVALID);
  for (let i = 0; i < 64; i++) table[ENCODE.charCodeAt(i)] = i;
  // The URL-safe alphabet, which some platform encoders produce.
  table[0x2d] = 62;
  table[0x5f] = 63;
  // Line breaks and blanks: several platform encoders wrap their output.
  table[0x09] = table[0x0a] = table[0x0d] = table[0x20] = SKIP;
  table[0x3d] = PAD;
  return table;
})();

/** Bytes as base64 text. Works in steps, so the size of the input is only limited by memory. */
export function base64Encode(bytes: Uint8Array): string {
  const parts: string[] = [];
  const codes = new Uint8Array((ENCODE_STEP / 3) * 4);
  for (let start = 0; start < bytes.length; start += ENCODE_STEP) {
    const end = Math.min(start + ENCODE_STEP, bytes.length);
    let n = 0;
    let i = start;
    for (; i + 2 < end; i += 3) {
      const triple = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
      codes[n++] = ENCODE.charCodeAt(triple >> 18);
      codes[n++] = ENCODE.charCodeAt((triple >> 12) & 63);
      codes[n++] = ENCODE.charCodeAt((triple >> 6) & 63);
      codes[n++] = ENCODE.charCodeAt(triple & 63);
    }
    // Only the last step can end with one or two bytes left over.
    if (i < end) {
      const triple = (bytes[i] << 16) | (i + 1 < end ? bytes[i + 1] << 8 : 0);
      codes[n++] = ENCODE.charCodeAt(triple >> 18);
      codes[n++] = ENCODE.charCodeAt((triple >> 12) & 63);
      codes[n++] = i + 1 < end ? ENCODE.charCodeAt((triple >> 6) & 63) : 0x3d;
      codes[n++] = 0x3d;
    }
    parts.push(String.fromCharCode.apply(null, (n === codes.length ? codes : codes.subarray(0, n)) as unknown as number[]));
  }
  return parts.join('');
}

/**
 * Base64 text as bytes. Accepts the standard and the URL-safe alphabet, white
 * space between characters and missing padding; throws an Error for anything else.
 */
export function base64Decode(text: string): Uint8Array {
  const length = text.length;
  let end = length;
  while (end > 0 && text.charCodeAt(end - 1) === 0x3d) end--;
  // Exact when the text has no white space, which is the usual case; otherwise an upper bound.
  const out = new Uint8Array(Math.floor((end * 3) / 4));
  let n = 0;
  let acc = 0;
  let bits = 0;
  let count = 0;
  let pads = 0;
  for (let i = 0; i < length; i++) {
    const code = text.charCodeAt(i);
    const value = code < 128 ? DECODE[code] : INVALID;
    if (value === SKIP) continue;
    if (value === PAD) {
      pads++;
      continue;
    }
    if (value === INVALID || pads > 0) throw new Error(`Invalid base64: unexpected character at offset ${i}`);
    acc = ((acc << 6) | value) & 0xfff;
    bits += 6;
    count++;
    if (bits >= 8) {
      bits -= 8;
      out[n++] = (acc >> bits) & 0xff;
    }
  }
  const rest = count % 4;
  if (rest === 1) throw new Error('Invalid base64: the text is cut short');
  if (pads > 0 && (pads > 2 || (rest + pads) % 4 !== 0)) throw new Error('Invalid base64: wrong padding');
  return n === out.length ? out : out.slice(0, n);
}

/**
 * The address a `load` may fetch, resolved against the page's own address, or
 * undefined when it may not: only blob: and data: URLs and http(s) addresses
 * of the page's own origin, without user name or password. A page that was
 * not loaded over http(s) (file:, an application's asset scheme) therefore
 * gets blob: and data: only. White space, control characters and backslashes
 * are refused outright, because URL parsers drop or rewrite them.
 */
export function resolveLoadUrl(url: string, base?: string): string | undefined {
  if (url.length === 0) return undefined;
  for (let i = 0; i < url.length; i++) {
    const code = url.charCodeAt(i);
    if (code <= 0x20 || code === 0x7f || code === 0x5c) return undefined;
  }
  let target: URL;
  let page: URL | undefined;
  try {
    page = base === undefined ? undefined : new URL(base);
    target = page ? new URL(url, page.href) : new URL(url);
  } catch {
    return undefined;
  }
  if (target.protocol === 'blob:' || target.protocol === 'data:') return target.href;
  if (target.protocol !== 'http:' && target.protocol !== 'https:') return undefined;
  if (target.username !== '' || target.password !== '') return undefined;
  if (!page || (page.protocol !== 'http:' && page.protocol !== 'https:') || page.origin !== target.origin) return undefined;
  return target.href;
}

/** Whether a `load` may fetch this address; see resolveLoadUrl(). */
export function isAllowedUrl(url: string, base?: string): boolean {
  return resolveLoadUrl(url, base) !== undefined;
}

/** A message that does not follow the contract. */
class BadMessage extends Error {}

type Message = { [field: string]: unknown };

function field<T>(message: Message, name: string, type: 'string' | 'boolean' | 'number', required: boolean): T | undefined {
  const value = message[name];
  if (value === undefined || value === null) {
    if (required) throw new BadMessage(`"${name}" is missing`);
    return undefined;
  }
  if (typeof value !== type) throw new BadMessage(`"${name}" must be a ${type}`);
  return value as T;
}

/** The message as an object: parsed when it is text. Throws BadMessage for anything that is not a JSON object. */
function parseMessage(input: unknown): Message {
  let message = input;
  if (typeof message === 'string') {
    try {
      message = JSON.parse(message);
    } catch {
      throw new BadMessage('The message is not valid JSON');
    }
  }
  if (typeof message !== 'object' || message === null || Array.isArray(message)) throw new BadMessage('The message must be a JSON object');
  return message as Message;
}

const errorText = (error: unknown) => (error instanceof Error && error.message ? error.message : String(error));

/** How often and for how long a channel that is not there is looked for: every 50 ms for 10 s after the last delivery. */
const RETRY_MS = 50;
const MAX_RETRIES = 200;

const DEFAULT_FETCH_TIMEOUT_MS = 30000;
const DEFAULT_MAX_FETCH_BYTES = 256 * 1024 * 1024;
const DEFAULT_MAX_PENDING_CHARS = 64 * 1024 * 1024;
/** An application stores the file somewhere, which can take a while on a slow connection. */
const DEFAULT_SAVE_RESULT_TIMEOUT_MS = 60000;
/** What an application says of a save that failed is shown to the user: one line of it. */
const MAX_RESULT_MESSAGE_CHARS = 200;

/** Messages that say how things are now: of those waiting for a channel only the newest of each kind is kept. */
const isStatus = (type: string) => type === 'changed' || type === 'selection' || type === 'sheet' || type === 'sheets' || type === 'saveState';

/** Longest name of a sheet a message may carry. What a sheet may be called is the view's to say; this only bounds what is looked at. */
const MAX_NAME_CHARS = 255;

/** A field that counts sheets: a whole number from 0 that a workbook could have as an index. */
function sheetIndexField(message: Message, name: string, required: boolean): number | undefined {
  const value = field<number>(message, name, 'number', required);
  if (value !== undefined && !(Number.isInteger(value) && value >= 0 && value <= 0x7fffffff)) throw new BadMessage(`"${name}" must be a sheet index`);
  return value;
}

function sheetNameField(message: Message, required: boolean): string | undefined {
  const value = field<string>(message, 'name', 'string', required);
  if (value !== undefined && value.length > MAX_NAME_CHARS) throw new BadMessage('"name" is too long');
  return value;
}

/** Fonts one `load` may bring. */
const MAX_FONTS = 16;
/** Base64 characters of one font: about 24 MB of font file. */
const MAX_FONT_CHARS = 32 * 1024 * 1024;
/** Largest font file fetched from an address. */
const MAX_FONT_BYTES = 24 * 1024 * 1024;
const MAX_FONT_ALIASES = 32;
/** How long a `load` waits for its fonts before the file is shown without them. */
const FONT_WAIT_MS = 3000;

/**
 * The `fonts` of a `load` as registrations: each `{ family, data | url,
 * weight?, style?, as?, prefer? }`, with the font file as base64 in `data` or
 * at an address a `load` may fetch from. Throws BadMessage for anything else.
 */
function fontsField(message: Message, host: BridgeHost): FontRegistration[] {
  const value = message.fonts;
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) throw new BadMessage('"fonts" must be a list');
  if (value.length > MAX_FONTS) throw new BadMessage(`"fonts" names more than ${MAX_FONTS} fonts`);
  const out: FontRegistration[] = [];
  /** Family, weight and style of the fonts so far, as the registry tells faces apart. */
  const faces: string[] = [];
  for (let i = 0; i < value.length; i++) {
    const where = `fonts[${i}]`;
    const entry = value[i] as Message;
    if (typeof entry !== 'object' || entry === null || Array.isArray(entry)) throw new BadMessage(`${where} must be an object`);
    const read = <T>(name: string, type: 'string' | 'boolean', required: boolean): T | undefined => {
      try {
        return field<T>(entry, name, type, required);
      } catch (error) {
        throw new BadMessage(`${where}: ${errorText(error)}`);
      }
    };
    const family = read<string>('family', 'string', true) as string;
    if (family.length > MAX_NAME_CHARS) throw new BadMessage(`${where}: "family" is too long`);
    if (sanitizeFontFamily(family) === '') throw new BadMessage(`${where}: "family" must be a font name of letters, digits, spaces or hyphens`);
    const data = read<string>('data', 'string', false);
    const url = read<string>('url', 'string', false);
    if ((data === undefined ? 0 : 1) + (url === undefined ? 0 : 1) !== 1) throw new BadMessage(`${where} needs exactly one of "data" and "url"`);
    let source: string | Uint8Array;
    if (data !== undefined) {
      if (data.length > MAX_FONT_CHARS) throw new BadMessage(`${where}: "data" is too long`);
      try {
        source = base64Decode(data);
      } catch (error) {
        throw new BadMessage(`${where}: ${errorText(error)}`);
      }
      if (source.length === 0) throw new BadMessage(`${where}: "data" is empty`);
    } else {
      // The same addresses as for the file itself: nothing is fetched from another origin.
      if (host.allowUrl === false) throw new BadMessage(`${where}: "url" is not accepted by a framed page; send the font as "data"`);
      const address = resolveLoadUrl(url as string, host.baseUrl);
      if (address === undefined) throw new BadMessage(`${where}: "url" must be a blob: or data: URL or an address of the page's own origin`);
      source = address;
    }
    const font: FontRegistration = { family, source };
    const weight = entry.weight;
    if (weight !== undefined && weight !== null) {
      if (weight !== 'normal' && weight !== 'bold' && !(typeof weight === 'number' && weight >= 1 && weight <= 1000)) {
        throw new BadMessage(`${where}: "weight" must be "normal", "bold" or a number from 1 to 1000`);
      }
      font.weight = weight;
    }
    const style = entry.style;
    if (style !== undefined && style !== null) {
      if (style !== 'normal' && style !== 'italic') throw new BadMessage(`${where}: "style" must be "normal" or "italic"`);
      font.style = style;
    }
    const as = entry.as;
    if (as !== undefined && as !== null) {
      const names = typeof as === 'string' ? [as] : as;
      const named = Array.isArray(names) && names.every((name) => typeof name === 'string' && name.length <= MAX_NAME_CHARS);
      if (!named) throw new BadMessage(`${where}: "as" must be a font name or a list of font names`);
      if ((names as string[]).length > MAX_FONT_ALIASES) throw new BadMessage(`${where}: "as" names more than ${MAX_FONT_ALIASES} fonts`);
      font.as = (names as string[]).slice();
    }
    const prefer = read<boolean>('prefer', 'boolean', false);
    if (prefer !== undefined) font.prefer = prefer;
    // One face twice: which of the two files would be the face depends on which loads last.
    const face = sanitizeFontFamily(family).toLowerCase() + '|' + (font.weight === undefined || font.weight === 'normal' ? 400 : font.weight === 'bold' ? 700 : Math.round(font.weight * 100) / 100) + '|' + (font.style || 'normal');
    const earlier = faces.indexOf(face);
    if (earlier >= 0) throw new BadMessage(`${where} is the same face as fonts[${earlier}] (family, weight and style)`);
    faces.push(face);
    out.push(font);
  }
  return out;
}

/** The sheets of a workbook as `loaded` and `sheets` list them. */
const sheetList = (workbook: Workbook): Message[] =>
  workbook.sheets.map((entry) => {
    const out: Message = { name: entry.name, state: entry.state, kind: entry.kind };
    // The colour of the tab, for a sheet that has one: a CSS colour ("#RRGGBB").
    const color = entry.tabColor ? resolveColor(entry.tabColor, workbook.styles) : undefined;
    if (color) out.tabColor = color;
    return out;
  });

interface Outgoing {
  type: string;
  id: string | undefined;
  json: string;
  /** Sheets of a waiting `changed`, which a later one takes over. */
  sheets?: number[];
  /** Which of the files that were shown the message speaks of (see fileSerial). */
  file: number;
  /** A `saved` that comes from the user: the save it hands over, which waits for its `saveResult`. */
  saveId?: string;
}

/** What a command on sheets is refused with when the view was asked and did nothing, for a reason it did not give. */
const NOT_DONE = 'The view did not change its sheets';

/** What a command is refused with while something in the view is still open that could not be applied or closed. */
const STILL_EDITING = 'A note or a sheet name that is being typed cannot be applied, or a dialog of rules is open with changes: apply or close it first';

/** The file that is shown: a workbook in `view`, or a document in `pages`. */
/**
 * Applies what is being typed in a workbook view (a cell, a note, the name of a sheet) and closes
 * a dialog of rules in which nothing was changed; false when a note or a name cannot be applied
 * and stays in its field, or a dialog of rules is open with changes. A cell that cannot be
 * written throws.
 */
function commitPendingOf(view: WorkbookView): boolean {
  if (typeof view.commitPending === 'function') return view.commitPending();
  // A view of before that method: only its cell editor is known.
  if (view.isEditing) view.commitEditing();
  return true;
}

/** What a document view says of its tracked changes; undefined for a view that knows nothing of them. */
function revisionState(pages: BridgeDocumentView): { count: number; tracking: boolean } | undefined {
  if (typeof pages.revisions !== 'function') return undefined;
  let count = 0;
  try {
    const list = pages.revisions();
    count = Array.isArray(list) ? list.length : 0;
  } catch {
    count = 0;
  }
  return { count, tracking: pages.trackingChanges === true };
}

interface Shown {
  file: OfficeFile;
  view: WorkbookView | undefined;
  pages: BridgeDocumentView | undefined;
  name: string | undefined;
  /** The file cannot be written with changes (doc, RTF, xls as they are): it stays read-only whatever is asked. */
  locked: boolean;
  /** The format the file had before it was converted to be edited. */
  convertedFrom: string | undefined;
  /** A document: it is shown read-only now (it was loaded so, or `setReadOnly` said so). */
  readOnly?: boolean;
  /** A document: what the application was last told about its tracked changes, as "count tracking". */
  revisionsSent?: string;
  /** A `load` with `save`: what the Save button and Ctrl+S of the page run. */
  saver?: SaveController;
  /** The language the file was loaded with or `setLocale` chose: what a save is spoken of in. */
  locale?: string;
}

/** The message handling of the bridge page, with everything it needs from the page handed in. */
export function createBridge(host: BridgeHost): Bridge {
  const timer = host.timer || ((callback: () => void, ms: number) => void setTimeout(callback, ms));
  const frame = host.frame || ((callback: () => void) => timer(callback, 16));
  const fetchTimeout = host.fetchTimeoutMs === undefined ? DEFAULT_FETCH_TIMEOUT_MS : host.fetchTimeoutMs;
  const maxFetchBytes = host.maxFetchBytes === undefined ? DEFAULT_MAX_FETCH_BYTES : host.maxFetchBytes;
  const maxPendingChars = host.maxPendingChars === undefined ? DEFAULT_MAX_PENDING_CHARS : host.maxPendingChars;
  const saveResultTimeout = host.saveResultTimeoutMs === undefined ? DEFAULT_SAVE_RESULT_TIMEOUT_MS : host.saveResultTimeoutMs;

  let shown: Shown | undefined;
  /** How many sheets the workbook that is shown has now; undefined while none is shown. */
  let sheetCount: (() => number) | undefined;
  let destroyed = false;
  /**
   * Counts loads and closes. A load that finds another number than its own
   * when it comes back from waiting has been overtaken and gives up.
   */
  let generation = 0;
  /** Ends the fetch that is under way, if any. */
  let cancelFetch: ((reason: string) => void) | undefined;
  /** Messages are handled one after the other: a `save` sent right after a `load` waits for it. */
  let chain: Promise<void> = Promise.resolve();

  const pending: Outgoing[] = [];
  let pendingChars = 0;
  let started = false;
  let retrying = false;
  let retries = 0;

  /** Counts the files that were shown: a status message speaks of one of them and is never joined with one of another. */
  let fileSerial = 0;
  let selectionDue: { sheet: number; range: string; active: string } | undefined;
  let selectionScheduled = false;
  let selectionSent = '';
  /** The list of sheets as the application was last told (in `loaded` or `sheets`), as JSON; '' while no workbook is shown. */
  let sheetsSent = '';
  /** The save of the user that the application has been handed and has not answered with `saveResult` yet. */
  let awaited: { saveId: string; resolve(): void; reject(error: unknown): void } | undefined;
  /** Counts the saves of the user: each `saved` that comes from one names itself, and its `saveResult` names it back. */
  let userSaves = 0;
  /** The application was told the state of the save of a file (`saveState`) that has been closed since, and not that it is over. */
  let saveStateLeft = false;

  const deliver = (json: string): boolean => {
    try {
      return host.post(json) !== false;
    } catch {
      return false;
    }
  };

  let flushing = false;
  const flush = (): boolean => {
    // An application that answers from inside its postMessage comes back here through receive(), while the
    // message it answers is still the first of the queue: the pass that is under way goes on with the rest.
    if (flushing) return true;
    flushing = true;
    try {
      while (pending.length > 0) {
        const entry = pending[0];
        if (!deliver(entry.json)) return false;
        // What was sent goes by itself: the application may have changed the queue while it was called.
        const at = pending.indexOf(entry);
        if (at >= 0) {
          pending.splice(at, 1);
          pendingChars -= entry.json.length;
        }
        // A channel that works earns the full time again when it goes away.
        retries = 0;
      }
      return true;
    } finally {
      flushing = false;
    }
  };

  const retry = () => {
    retrying = false;
    if (destroyed || flush()) return;
    if (retries++ < MAX_RETRIES) {
      retrying = true;
      timer(retry, RETRY_MS);
    }
  };

  const errorMessage = (code: BridgeErrorCode, message: string, id?: string): Message => {
    const out: Message = { type: 'error' };
    if (id !== undefined) out.id = id;
    out.code = code;
    out.message = message;
    return out;
  };

  const send = (message: Message) => {
    if (destroyed) return;
    const type = message.type as string;
    if (type === 'sheets') {
      // The list that waits is brought up to date where it stands: what follows it (a `sheet`, a
      // `changed`, the answer to a command) speaks of sheets by index, and must not arrive before it.
      for (let i = pending.length - 1; i >= 0; i--) {
        if (pending[i].type !== type || pending[i].file !== fileSerial) continue;
        const json = JSON.stringify(message);
        pendingChars += json.length - pending[i].json.length;
        pending[i] = { ...pending[i], json };
        flush();
        return;
      }
    } else if (isStatus(type)) {
      for (let i = pending.length - 1; i >= 0; i--) {
        if (pending[i].type !== type || pending[i].file !== fileSerial) continue;
        const older = pending.splice(i, 1)[0];
        pendingChars -= older.json.length;
        // One `changed` stands for all that were waiting: it names every sheet they named.
        if (type === 'changed' && older.sheets) {
          const sheets = (message.sheets as number[]).slice();
          // A sheet may have been deleted since: an index of before that no sheet has now names nothing.
          const count = sheetCount ? sheetCount() : Infinity;
          for (const index of older.sheets) if (index < count && sheets.indexOf(index) < 0) sheets.push(index);
          message = { ...message, sheets: sheets.sort((a, b) => a - b) };
        }
      }
    }
    const entry: Outgoing = { type, id: message.id as string | undefined, json: JSON.stringify(message), file: fileSerial };
    if (type === 'changed') entry.sheets = message.sheets as number[];
    if (type === 'saved' && typeof message.saveId === 'string') entry.saveId = message.saveId;
    pending.push(entry);
    pendingChars += entry.json.length;
    if (flush()) return;
    // Nobody is listening. Answers are never dropped, but a saved file is big:
    // past the limit the oldest ones become errors the application can act on.
    for (let i = 0; i < pending.length && pendingChars > maxPendingChars; i++) {
      const given = pending[i];
      if (given.type !== 'saved') continue;
      const error = errorMessage('save-failed', 'The saved file was given up: the application did not take it', given.id);
      // A save of the user answers no request: the error names it as its `saved` would have.
      if (given.saveId !== undefined) error.saveId = given.saveId;
      const json = JSON.stringify(error);
      pendingChars += json.length - given.json.length;
      pending[i] = { type: 'error', id: given.id, json, file: given.file };
      // No `saveResult` can come for a file the application never gets: the save has failed now, not when its time is up.
      const waiting = awaited;
      if (given.saveId !== undefined && waiting && waiting.saveId === given.saveId) {
        awaited = undefined;
        waiting.reject(new SaveError('refused', 'The saved file was given up: the application did not take it'));
      }
    }
    if (!retrying && retries < MAX_RETRIES) {
      retrying = true;
      timer(retry, RETRY_MS);
    }
  };

  const fail = (code: BridgeErrorCode, message: string, id?: string) => send(errorMessage(code, message, id));

  const done = (command: string, id: string | undefined) => {
    if (id !== undefined) send({ type: 'done', id, command });
  };

  // The page's own callbacks must not be able to leave the bridge half way through a step.
  const setState = (state: BridgeState, code?: BridgeErrorCode) => {
    try {
      if (host.onState) host.onState(state, code);
    } catch {
      // Nothing to do: what the page shows around the editor is not the bridge's state.
    }
  };

  const setLocale = (locale: string | undefined) => {
    try {
      if (locale !== undefined && host.onLocale) host.onLocale(locale);
    } catch {
      // As above.
    }
  };

  const closeShown = () => {
    selectionDue = undefined;
    selectionSent = '';
    sheetsSent = '';
    if (!shown) return;
    fileSerial++;
    const view = shown.view || shown.pages;
    // A save that waits for its result is of a file that is gone: nothing is said about it any more.
    // What the application was last told of it (unsaved, saving) is ended by whoever closed it: see settleSaveState().
    if (shown.saver) {
      shown.saver.destroy();
      saveStateLeft = true;
    }
    awaited = undefined;
    shown = undefined;
    sheetCount = undefined;
    if (view) view.destroy();
  };

  /**
   * Tells the application where the save stands after the file that is shown was replaced or closed: a
   * file with a Save button starts as saved and not saving, and the state of a file that is gone is ended
   * with the same words. Without it the application would hold "unsaved" or "saving" of the file before
   * for a file that is neither. Nothing is said when no file with a Save button is or was there.
   */
  const settleSaveState = () => {
    const saver = shown && shown.saver;
    if (!saver && !saveStateLeft) return;
    saveStateLeft = false;
    send({ type: 'saveState', dirty: saver ? saver.dirty : false, saving: saver ? saver.saving : false });
  };

  /** Gives up whatever load is under way or waiting. */
  const supersede = (reason: string): number => {
    generation++;
    if (cancelFetch) cancelFetch(reason);
    return generation;
  };

  const flushSelection = () => {
    selectionScheduled = false;
    const due = selectionDue;
    selectionDue = undefined;
    if (!due || !shown) return;
    const key = `${due.sheet}!${due.range}!${due.active}`;
    if (key === selectionSent) return;
    selectionSent = key;
    send({ type: 'selection', sheet: due.sheet, range: due.range, active: due.active });
  };

  /** Fetches with a time limit and a way to be cancelled; the size limit is the host's and is checked again here. */
  const fetchFile = (url: string): Promise<Uint8Array> =>
    new Promise<Uint8Array>((resolve, reject) => {
      const fetchBytes = host.fetchBytes;
      if (!fetchBytes) return reject(new Error('loading from a URL is not available here'));
      const controller = typeof AbortController === 'function' ? new AbortController() : undefined;
      let over = false;
      const cancel = (reason: string) => {
        if (over) return;
        over = true;
        if (cancelFetch === cancel) cancelFetch = undefined;
        // Without AbortController the request runs on, and its result is ignored.
        if (controller) controller.abort();
        reject(new Error(reason));
      };
      cancelFetch = cancel;
      timer(() => cancel(`no answer within ${Math.round(fetchTimeout / 1000)} s`), fetchTimeout);
      let request: Promise<Uint8Array>;
      try {
        request = fetchBytes(url, { signal: controller ? controller.signal : undefined, maxBytes: maxFetchBytes });
      } catch (error) {
        request = Promise.reject(error);
      }
      request.then(
        (bytes) => {
          if (over) return;
          over = true;
          if (cancelFetch === cancel) cancelFetch = undefined;
          if (bytes.length > maxFetchBytes) reject(new Error(`the file is larger than ${maxFetchBytes} bytes`));
          else resolve(bytes);
        },
        (error) => {
          if (over) return;
          over = true;
          if (cancelFetch === cancel) cancelFetch = undefined;
          reject(error);
        },
      );
    });

  /**
   * Registers a font of a `load`. A font at an address is fetched here, as
   * the file is (the page's policy allows no other way, and the size is
   * bounded), and registered as bytes.
   */
  const registerLoadFont = (font: FontRegistration): Promise<unknown> => {
    const register = host.registerFont ? (entry: FontRegistration) => (host.registerFont as NonNullable<BridgeHost['registerFont']>)(entry) : registerFont;
    const address = font.source;
    if (typeof address !== 'string') return register(font);
    return new Promise<Uint8Array>((resolve, reject) => {
      const fetchBytes = host.fetchBytes;
      if (!fetchBytes) return reject(new Error('loading from a URL is not available here'));
      const controller = typeof AbortController === 'function' ? new AbortController() : undefined;
      let over = false;
      timer(() => {
        if (over) return;
        over = true;
        if (controller) controller.abort();
        reject(new Error(`no answer within ${Math.round(fetchTimeout / 1000)} s`));
      }, fetchTimeout);
      let request: Promise<Uint8Array>;
      try {
        request = fetchBytes(address, { signal: controller ? controller.signal : undefined, maxBytes: MAX_FONT_BYTES });
      } catch (error) {
        request = Promise.reject(error);
      }
      request.then(
        (bytes) => {
          if (over) return;
          over = true;
          if (bytes.length > MAX_FONT_BYTES) reject(new Error(`the font file is larger than ${MAX_FONT_BYTES} bytes`));
          else if (bytes.length === 0) reject(new Error('the font file is empty'));
          else resolve(bytes);
        },
        (error) => {
          if (over) return;
          over = true;
          reject(error);
        },
      );
    }).then((bytes) => register(Object.assign({}, font, { source: bytes })));
  };

  const load = async (message: Message, id: string | undefined, mine: number) => {
    const data = field<string>(message, 'data', 'string', false);
    const url = field<string>(message, 'url', 'string', false);
    const name = field<string>(message, 'name', 'string', false);
    const locale = field<string>(message, 'locale', 'string', false);
    const readOnly = field<boolean>(message, 'readOnly', 'boolean', false);
    const sheet = field<number>(message, 'sheet', 'number', false);
    const blank = field<string>(message, 'blank', 'string', false);
    const convert = field<boolean>(message, 'convert', 'boolean', false);
    const toolbar = field<boolean>(message, 'toolbar', 'boolean', false);
    const author = field<string>(message, 'author', 'string', false);
    const trackChanges = field<boolean>(message, 'trackChanges', 'boolean', false);
    const userSaves = field<boolean>(message, 'save', 'boolean', false);
    if (author !== undefined && author.length > 255) throw new BadMessage('"author" is too long');
    if ((data === undefined ? 0 : 1) + (url === undefined ? 0 : 1) + (blank === undefined ? 0 : 1) !== 1) {
      throw new BadMessage('"load" needs exactly one of "data", "url" and "blank"');
    }
    if (blank !== undefined && blank !== 'docx' && blank !== 'xlsx') throw new BadMessage('"blank" must be "docx" or "xlsx"');
    if (sheet !== undefined && (!Number.isInteger(sheet) || sheet < 0)) throw new BadMessage('"sheet" must be a sheet index');
    const fonts = fontsField(message, host);
    let address: string | undefined;
    if (url !== undefined) {
      if (host.allowUrl === false) throw new BadMessage('"url" is not accepted by a framed page; send the file as "data"');
      address = resolveLoadUrl(url, host.baseUrl);
      if (address === undefined) throw new BadMessage('"url" must be a blob: or data: URL or an address of the page\'s own origin');
    }
    let bytes: Uint8Array | undefined;
    if (data !== undefined) {
      try {
        bytes = base64Decode(data);
      } catch (error) {
        throw new BadMessage(errorText(error));
      }
    }

    // A load that was overtaken (by a later load, a close or destroy) answers its id and changes nothing.
    const overtaken = () => {
      if (mine === generation && !destroyed) return false;
      if (id !== undefined) fail('load-failed', 'The load was cancelled', id);
      return true;
    };
    // The file that is open stays open, and stays on screen, until the new one has been read.
    const refuse = (code: BridgeErrorCode, text: string) => {
      if (!shown) setState(code === 'bad-message' ? 'empty' : 'error', code === 'bad-message' ? undefined : code);
      fail(code, text, id);
    };

    if (overtaken()) return;
    if (!shown) setState('loading');
    // The fonts load while the file is fetched and read; the file is shown without those that fail or are late.
    const fontsSettled: Promise<FontFailure[]> = fonts.length > 0 ? settleFonts(fonts, FONT_WAIT_MS, registerLoadFont, timer) : Promise.resolve([]);
    if (blank !== undefined) {
      try {
        bytes = await createBlank(blank as 'docx' | 'xlsx', locale === undefined ? {} : { locale });
      } catch (error) {
        if (overtaken()) return;
        return refuse('load-failed', `Could not make the new file: ${errorText(error)}`);
      }
      if (overtaken()) return;
    }
    if (bytes === undefined) {
      try {
        bytes = await fetchFile(address as string);
      } catch (error) {
        if (overtaken()) return;
        return refuse('load-failed', `Could not load the file: ${errorText(error)}`);
      }
      if (overtaken()) return;
    }

    let file: OfficeFile;
    let workbook: Workbook | undefined;
    let doc: DocModel | undefined;
    let convertedFrom: string | undefined;
    let warnings: string[] | undefined;
    try {
      file = await host.open(bytes);
      if (overtaken()) return;
      const showable = (): void => {
        if (file.kind === 'document' && (!file.document || !host.createDocumentView)) {
          throw Object.assign(new Error('This kind of document cannot be shown'), { name: 'UnsupportedFileError' });
        }
      };
      showable();
      if (convert === true && readOnly !== true && needsConversion(file.format)) {
        // The file is written out in the format that can be saved, and that file is what is shown.
        const opened = await openConverted(file, host.open);
        if (overtaken()) return;
        if (opened.conversion) {
          file = opened.file;
          convertedFrom = opened.conversion.from;
          warnings = opened.conversion.warnings;
          showable();
        } else if (opened.failed !== undefined) {
          // Shown as it is, read-only: better than refusing a file that can be read.
          warnings = [`The file could not be converted and is shown read-only: ${opened.failed}`];
        }
      }
      if (file.kind === 'document') {
        doc = await (file.document as NonNullable<OfficeFile['document']>)();
      } else {
        workbook = await file.workbook();
        if (workbook.sheets.length === 0) throw new Error('The workbook has no sheets');
      }
    } catch (error) {
      if (overtaken()) return;
      return refuse(error instanceof Error && error.name === 'UnsupportedFileError' ? 'unsupported' : 'corrupt', errorText(error));
    }
    if (overtaken()) return;
    const fontFailures = await fontsSettled;
    if (overtaken()) return;
    for (let i = 0; i < fontFailures.length; i++) {
      const failure = fontFailures[i];
      const text = failure.pending ? `The font "${failure.family}" is still loading; the file is shown without it until it arrives` : `The font "${failure.family}" could not be loaded: ${failure.reason}`;
      (warnings || (warnings = [])).push(text);
    }
    // What the document enforces: a protection for comments or forms allows no editing here, as the read-only one.
    const protection = doc ? enforcedProtection(doc) : undefined;
    // Only docx and xlsx are written back with their changes; the other formats stay read-only until they are converted.
    const locked = needsConversion(file.format) || (protection !== undefined && protection !== 'trackedChanges');

    const loaded = (activeSheet: number, sheets: unknown[], document: boolean) => {
      const out: Message = { type: 'loaded' };
      if (id !== undefined) out.id = id;
      out.format = file.format;
      if (document) out.kind = 'document';
      out.sheets = sheets;
      out.activeSheet = activeSheet;
      // The file is shown read-only whatever was asked for, and `setReadOnly` will not change that.
      if (locked) out.readOnly = true;
      // The page has a Save button and takes Ctrl+S: the application will be handed what the user saves.
      if (shown && shown.saver) out.save = true;
      if (convertedFrom !== undefined) out.convertedFrom = convertedFrom;
      if (warnings !== undefined) out.warnings = warnings;
      if (doc) {
        // Why the document is read-only, or that its changes are tracked.
        if (protection !== undefined) out.protection = protection;
        if (protection !== undefined && hasPermissionRanges(doc)) out.protectionExceptions = true;
        if (!locked && (trackChanges === true || tracksChanges(doc))) out.trackChanges = true;
        // How many tracked changes it has, when it has any.
        const state = shown && shown.pages ? revisionState(shown.pages) : undefined;
        if (state && state.count > 0) out.revisions = state.count;
        if (shown && state) shown.revisionsSent = `${state.count} ${state.tracking}`;
      }
      send(out);
      // Right after it: where the save of this file stands, in place of what was said of the file before.
      settleSaveState();
    };

    if (doc) {
      let pages: BridgeDocumentView | undefined;
      const documentOptions: DocumentEditorOptions = {
        package: documentPictures(doc as DocModel, file.package),
        readOnly: locked || readOnly === true,
        // A document has no sheets: the message says that it changed and what can be undone.
        onChange: () => {
          // Nothing is said for a view that is not there yet or that was closed.
          if (!pages || !shown || shown.pages !== pages) return;
          if (shown.saver) shown.saver.changed();
          // The tracked changes first, when their number or the tracking is not what was told last.
          const state = revisionState(pages);
          if (state) {
            const key = `${state.count} ${state.tracking}`;
            if (key !== shown.revisionsSent) {
              shown.revisionsSent = key;
              send({ type: 'revisions', count: state.count, trackChanges: state.tracking });
            }
          }
          send({ type: 'changed', sheets: [], canUndo: pages.canUndo === true, canRedo: pages.canRedo === true });
        },
        // The page must not be replaced by what a link points to: the application decides what to do with it.
        onLinkClick: (target) => {
          send({ type: 'link', url: target });
          return false;
        },
      };
      if (locale !== undefined) documentOptions.locale = locale;
      if (author !== undefined) documentOptions.author = author;
      if (trackChanges === true) documentOptions.trackChanges = true;
      const saver = userSaves === true && !locked ? userSaver() : undefined;
      if (saver) documentOptions.saveCommand = saver;
      const create = (toolbar === true && host.createDocumentEditor) || (host.createDocumentView as NonNullable<BridgeHost['createDocumentView']>);
      closeShown();
      try {
        pages = create.call(host, host.container, doc, documentOptions);
      } catch (error) {
        if (saver) saver.destroy();
        setState('error', 'corrupt');
        fail('corrupt', errorText(error), id);
        // The file before it was closed for this one, and nothing is shown now.
        return settleSaveState();
      }
      shown = { file, view: undefined, pages, name, locked, convertedFrom, readOnly: locked || readOnly === true, saver, locale };
      setLocale(locale);
      setState('open');
      loaded(0, [], true);
      return;
    }
    if (!workbook) return;
    const book = workbook;
    if (sheet !== undefined && sheet >= workbook.sheets.length) return refuse('bad-message', `No sheet with index ${sheet}`);

    let view: WorkbookView | undefined;
    /**
     * Tells the application the list of sheets when it is not the one it was told last: a
     * sheet was inserted, deleted or renamed, by a command, by the user or by undoing that.
     * Sent before the message that made the bridge look, so that what that one says about
     * a sheet (an index, a name) is said of a list the application knows.
     */
    const tellSheets = () => {
      if (!view || !shown || shown.view !== view) return;
      const sheets = sheetList(book);
      const json = JSON.stringify(sheets);
      if (json === sheetsSent) return;
      sheetsSent = json;
      // A selection that waits to be told was noted with the index its sheet had in the list before:
      // it is told of the sheet that is shown now, by the index that sheet has now.
      if (selectionDue) {
        const selection = view.selection;
        selectionDue = { sheet: view.activeSheet, range: rangeName(selection.range), active: cellName(selection.active.row, selection.active.col) };
      }
      send({ type: 'sheets', sheets, activeSheet: view.activeSheet });
    };
    const options: SpreadsheetEditorOptions = {
      onChange: (change) => {
        if (!view) return;
        if (shown && shown.view === view && shown.saver) shown.saver.changed();
        tellSheets();
        send({ type: 'changed', sheets: change.sheets, canUndo: view.history.canUndo, canRedo: view.history.canRedo });
      },
      // At most one selection message per frame, and only the last selection of that frame.
      onSelectionChange: (selection, sheetIndex) => {
        selectionDue = { sheet: sheetIndex, range: rangeName(selection.range), active: cellName(selection.active.row, selection.active.col) };
        if (selectionScheduled) return;
        selectionScheduled = true;
        frame(flushSelection);
      },
      onSheetChange: (index) => {
        tellSheets();
        const entry = book.sheets[index];
        send({ type: 'sheet', index, name: entry ? entry.name : '' });
      },
    };
    if (locale !== undefined) options.locale = locale;
    if (readOnly !== undefined) options.readOnly = readOnly;
    if (locked) options.readOnly = true;
    if (sheet !== undefined) options.sheet = sheet;
    // Who the notes added to cells are by.
    if (author !== undefined) options.author = author;

    const saver = userSaves === true && !locked ? userSaver() : undefined;
    if (saver) options.saveCommand = saver;
    // A load that says `toolbar: false` wants the workbook without the bars it is worked on with: the grid, its tabs
    // and the status bar are what is left (a viewer). Without the field a workbook is shown as it always was.
    if (toolbar === false) {
      options.toolbar = false;
      options.formulaBar = false;
    }
    const createSheet = (toolbar === true && host.createEditor) || host.createView;
    closeShown();
    try {
      view = createSheet.call(host, host.container, workbook, options);
    } catch (error) {
      if (saver) saver.destroy();
      setState('error', 'corrupt');
      fail('corrupt', errorText(error), id);
      // The file before it was closed for this one, and nothing is shown now.
      return settleSaveState();
    }
    shown = { file, view, pages: undefined, name, locked, convertedFrom, saver, locale };
    sheetCount = () => (Array.isArray(book.sheets) ? book.sheets.length : Infinity);
    setLocale(locale);
    setState('open');
    const list = sheetList(workbook);
    sheetsSent = JSON.stringify(list);
    loaded(view.activeSheet, list, false);
  };

  /**
   * Accepts or rejects every tracked change of the document that is shown, or turns its
   * tracking on or off. What results is told by `revisions` and `changed` when anything
   * changed; the answer says what holds afterwards, also when the view changed nothing
   * (a protection that is enforced, tracking that cannot be turned off).
   */
  const reviseDocument = (type: 'acceptAllRevisions' | 'rejectAllRevisions' | 'setTrackChanges', message: Message, id: string | undefined) => {
    const on = type === 'setTrackChanges' ? field<boolean>(message, 'on', 'boolean', true) : undefined;
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    if (!pages) return fail('no-document', 'The open file is not a document', id);
    if (shown.locked || shown.readOnly === true) return fail('read-only', 'The document is read-only', id);
    if (type === 'setTrackChanges') {
      if (pages.setTrackChanges) pages.setTrackChanges(on === true);
    } else if (type === 'acceptAllRevisions') {
      if (pages.acceptAllRevisions) pages.acceptAllRevisions();
    } else if (pages.rejectAllRevisions) pages.rejectAllRevisions();
    if (id === undefined) return;
    const out: Message = { type: 'done', id, command: type };
    const state = revisionState(pages);
    if (state) {
      out.revisions = state.count;
      out.trackChanges = state.tracking;
    }
    send(out);
  };

  /**
   * Inserts, renames or deletes a sheet through the view, as one undo step. The list that
   * results is told by `sheets` (see tellSheets() in load()), whether or not the command
   * has an id; the answer says that the command was done, or why it was not.
   */
  const changeSheets = (type: 'insertSheet' | 'renameSheet' | 'deleteSheet', message: Message, id: string | undefined) => {
    // Everything that comes from outside is looked at before anything is asked of the view.
    const at = type === 'insertSheet' ? sheetIndexField(message, 'at', false) : undefined;
    const index = type === 'insertSheet' ? undefined : sheetIndexField(message, 'index', true);
    const name = type === 'deleteSheet' ? undefined : sheetNameField(message, type === 'renameSheet');
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    if (!view) return fail('no-sheet', 'The open file is a document: it has no sheets', id);
    if (shown.locked || view.readOnly === true) return fail('read-only', 'The file is read-only', id);
    if (typeof view.insertSheet !== 'function' || typeof view.renameSheet !== 'function' || typeof view.deleteSheet !== 'function') {
      return fail('sheet-refused', 'This view cannot change its sheets', id);
    }
    const count = sheetCount ? sheetCount() : 0;
    if (index !== undefined && index >= count) return fail('no-sheet', `No sheet with index ${index}`, id);
    if (at !== undefined && at > count) return fail('no-sheet', `A sheet cannot be inserted at index ${at}: there are ${count} sheets`, id);
    const answer: Message = { type: 'done', id, command: type };
    try {
      // What is being typed is applied first, as the view itself does it. What cannot be applied (and a dialog of
      // rules with changes) stays open, the view does nothing to its sheets then, and "done" would be a lie.
      if (!commitPendingOf(view)) return fail('sheet-refused', STILL_EDITING, id);
      if (type === 'insertSheet') {
        const placed = view.insertSheet(at, name);
        // The view was asked and did nothing: it is not read-only (that was looked at), so it refused.
        if (placed < 0) return fail('sheet-refused', NOT_DONE, id);
        const entry = view.workbook && Array.isArray(view.workbook.sheets) ? view.workbook.sheets[placed] : undefined;
        answer.index = placed;
        if (entry) answer.name = entry.name;
      } else if (type === 'renameSheet') {
        // False from here on means the sheet has that name already: nothing to do is done.
        view.renameSheet(index as number, name as string);
      } else if (view.deleteSheet(index as number) === false) return fail('sheet-refused', NOT_DONE, id);
    } catch (error) {
      // The reason in the words the view has for it, in its language; the code is the same for every refusal.
      let text: string;
      try {
        text = typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
      } catch {
        text = errorText(error);
      }
      return fail('sheet-refused', text, id);
    }
    if (id !== undefined) send(answer);
  };

  /**
   * Moves, hides or shows, copies a sheet or colours its tab through the view, as one undo
   * step. Like changeSheets(): the list that results is told by `sheets`, and the answer says
   * that the command was done, or why it was not.
   */
  const arrangeSheets = (type: 'moveSheet' | 'setSheetState' | 'copySheet' | 'setTabColor', message: Message, id: string | undefined) => {
    // Everything that comes from outside is looked at before anything is asked of the view.
    const index = sheetIndexField(message, type === 'moveSheet' ? 'from' : 'index', true) as number;
    const to = type === 'moveSheet' ? (sheetIndexField(message, 'to', true) as number) : undefined;
    const at = type === 'copySheet' ? sheetIndexField(message, 'at', false) : undefined;
    const name = type === 'copySheet' ? sheetNameField(message, false) : undefined;
    const state = type === 'setSheetState' ? field<string>(message, 'state', 'string', true) : undefined;
    if (state !== undefined && state !== 'visible' && state !== 'hidden' && state !== 'veryHidden') throw new BadMessage('"state" must be "visible", "hidden" or "veryHidden"');
    // No colour, or null, takes the colour away.
    const color = type === 'setTabColor' && message.color !== null ? field<string>(message, 'color', 'string', false) : undefined;
    if (color !== undefined && !/^#[0-9a-fA-F]{6}$/.test(color)) throw new BadMessage('"color" must be written #RRGGBB');
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    if (!view) return fail('no-sheet', 'The open file is a document: it has no sheets', id);
    if (shown.locked || view.readOnly === true) return fail('read-only', 'The file is read-only', id);
    if (typeof view.moveSheet !== 'function' || typeof view.setSheetState !== 'function' || typeof view.copySheet !== 'function' || typeof view.setTabColor !== 'function') {
      return fail('sheet-refused', 'This view cannot change its sheets', id);
    }
    const count = sheetCount ? sheetCount() : 0;
    if (index >= count) return fail('no-sheet', `No sheet with index ${index}`, id);
    if (to !== undefined && to >= count) return fail('no-sheet', `No sheet with index ${to}`, id);
    if (at !== undefined && at > count) return fail('no-sheet', `A sheet cannot be inserted at index ${at}: there are ${count} sheets`, id);
    const answer: Message = { type: 'done', id, command: type };
    try {
      // As in changeSheets(): nothing is said to be done that the view did not do.
      if (!commitPendingOf(view)) return fail('sheet-refused', STILL_EDITING, id);
      // False from these three means the sheet is where, how or of the colour that was asked for already.
      if (type === 'moveSheet') view.moveSheet(index, to as number);
      else if (type === 'setSheetState') view.setSheetState(index, state as 'visible' | 'hidden' | 'veryHidden');
      else if (type === 'setTabColor') view.setTabColor(index, color);
      else {
        const options: { at?: number; name?: string } = {};
        if (at !== undefined) options.at = at;
        if (name !== undefined) options.name = name;
        const placed = view.copySheet(index, options);
        if (placed < 0) return fail('sheet-refused', NOT_DONE, id);
        const entry = view.workbook && Array.isArray(view.workbook.sheets) ? view.workbook.sheets[placed] : undefined;
        answer.index = placed;
        if (entry) answer.name = entry.name;
      }
    } catch (error) {
      let text: string;
      try {
        text = typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
      } catch {
        text = errorText(error);
      }
      return fail('sheet-refused', text, id);
    }
    // A command that changed nothing a message tells of (a colour) still has its list looked at.
    if (id !== undefined) send(answer);
  };

  /** The file that is shown as bytes, with what is being typed in it. */
  const writeShown = async (current: Shown): Promise<Uint8Array> => {
    // What is being typed (in a cell, in the field of a note, in a tab) is written first, as leaving it would do.
    // A note or a name that cannot be applied stays in its field, and nothing is saved around it.
    // Nor is anything saved under a dialog of rules that is open with changes: those are the user's to confirm.
    if (current.view && !commitPendingOf(current.view)) throw new Error(STILL_EDITING);
    const pages = current.pages;
    // A composition still under way belongs to the document that is written.
    if (pages && pages.commitEditing) pages.commitEditing();
    const changes = pages ? pages.changeCount : undefined;
    const bytes = await current.file.save();
    // A change made while the file was written is not in these bytes.
    if (pages && pages.markSaved && pages.changeCount === changes) pages.markSaved();
    return bytes;
  };

  /**
   * What the Save button and Ctrl+S of the page run for a `load` with `save`: the application is
   * handed the file in a `saved` message that says it comes from the user, and answers with
   * `saveResult` when it has stored it. Until then the page shows the save as under way; without
   * an answer in time it has failed.
   */
  function userSaver(): SaveController {
    const mine: SaveController = new SaveController({
      target: {
        kind: 'function',
        save: (bytes: Uint8Array, info: SaveInfo) =>
          new Promise<void>((resolve, reject) => {
            // Never the id of a save before it, of this file or of another: an answer that comes late names a save that is gone.
            const saveId = `u${++userSaves}`;
            const waiting = { saveId, resolve, reject };
            const out: Message = { type: 'saved', by: 'user', saveId, data: base64Encode(bytes), format: info.format };
            if (info.name !== undefined) out.name = info.name;
            if (info.convertedFrom !== undefined) out.convertedFrom = info.convertedFrom;
            awaited = waiting;
            send(out);
            timer(() => {
              if (awaited !== waiting) return;
              awaited = undefined;
              reject(new SaveError('timeout', `The application did not answer the save within ${Math.round(saveResultTimeout / 1000)} s`));
            }, saveResultTimeout);
          }),
      },
      writable: () => shown !== undefined && shown.saver === mine && !shown.locked,
      commit: () => {
        const current = shown;
        if (!current || current.saver !== mine) return false;
        if (!current.view) {
          // A composition still under way is a change of its own: it is ended before the changes are counted.
          if (current.pages && current.pages.commitEditing) current.pages.commitEditing();
          return true;
        }
        try {
          return commitPendingOf(current.view);
        } catch (error) {
          // A cell whose text cannot be written: the editor stays open, and the user is told why.
          const view = current.view;
          const text = typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
          if (typeof view.showNotice === 'function') view.showNotice(text, 'error');
          return false;
        }
      },
      bytes: () => {
        const current = shown;
        if (!current || current.saver !== mine) return Promise.reject(new Error('No file is open'));
        return writeShown(current);
      },
      info: () => {
        const current = shown;
        const out: SaveInfo = { format: (current ? current.file.format : 'docx') as SaveInfo['format'] };
        if (current && current.name !== undefined) out.name = current.name;
        if (current && current.convertedFrom !== undefined) out.convertedFrom = current.convertedFrom as SaveInfo['convertedFrom'];
        return out;
      },
      strings: () => resolveUiStrings(shown ? shown.locale : undefined),
      // The clock of the page, as everything else here uses it.
      timer: (callback, ms) => {
        let off = false;
        timer(() => {
          if (!off) callback();
        }, ms);
        return () => {
          off = true;
        };
      },
      notice: (text, kind) => {
        const current = shown;
        if (!current || current.saver !== mine) return;
        const view = current.view || current.pages;
        if (view && typeof view.showNotice === 'function') view.showNotice(text, kind);
      },
      // What the page knows and the application cannot count for itself: whether the file as it is now is the file
      // the application was last handed and said it stored. (A change made while the bytes were written is in no
      // `saved` message; a history that was cut can be undone to its end and still not be the file that was opened.)
      onState: (state) => {
        const current = shown;
        if (!current || current.saver !== mine) return;
        send({ type: 'saveState', dirty: state.dirty, saving: state.saving });
      },
    });
    return mine;
  }

  /** The application's answer to a `saved` message that came from the user. */
  const saveResult = (message: Message, id: string | undefined) => {
    const saveId = field<string>(message, 'saveId', 'string', true);
    const ok = field<boolean>(message, 'ok', 'boolean', true);
    const text = field<string>(message, 'message', 'string', false);
    const waiting = awaited;
    // Only the save that waits now is answered, and only by its own id: a result that comes after the page
    // stopped waiting, twice, or for a file that was closed must not be taken for the result of the next save.
    if (!waiting || waiting.saveId !== saveId) return fail('bad-message', 'No save with this "saveId" is waiting for a result', id);
    awaited = undefined;
    if (ok === true) waiting.resolve();
    else waiting.reject(new Error(text === undefined ? '' : text.slice(0, MAX_RESULT_MESSAGE_CHARS)));
    done('saveResult', id);
  };

  const save = async (id: string | undefined) => {
    const current = shown;
    if (!current) return fail('not-loaded', 'No file is open', id);
    let data: string;
    try {
      data = base64Encode(await writeShown(current));
    } catch (error) {
      return fail('save-failed', errorText(error), id);
    }
    const out: Message = { type: 'saved' };
    if (id !== undefined) out.id = id;
    out.data = data;
    if (current.name !== undefined) out.name = current.name;
    // The application gave another format than it gets back: the message says which.
    if (current.convertedFrom !== undefined) {
      out.format = current.file.format;
      out.convertedFrom = current.convertedFrom;
    }
    send(out);
  };

  /**
   * Starts printing the document or the workbook and answers when the dialog
   * was closed. The messages after it do not wait for that: a dialog can stay
   * open for long.
   */
  const print = (message: Message, id: string | undefined) => {
    const dpi = field<number>(message, 'dpi', 'number', false);
    if (dpi !== undefined && !(dpi > 0 && dpi < Infinity)) throw new BadMessage('"dpi" must be a positive number');
    const sheets = field<string>(message, 'sheets', 'string', false);
    if (sheets !== undefined && sheets !== 'active' && sheets !== 'all') throw new BadMessage('"sheets" must be "active" or "all"');
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    const view = shown.view;
    if (pages ? !pages.print : !view || typeof view.print !== 'function') return fail('print-unavailable', 'This view cannot print', id);
    let job: Promise<unknown>;
    try {
      if (pages && pages.print) job = Promise.resolve(pages.print(dpi === undefined ? undefined : { dpi }));
      else {
        // Of a workbook: the sheet that is shown unless all are asked for; the name of the file is what "&F" in a header stands for.
        const options: SheetPrintOptions = {};
        if (dpi !== undefined) options.dpi = dpi;
        if (sheets !== undefined) options.sheets = sheets as 'active' | 'all';
        if (shown.name !== undefined) options.fileName = shown.name;
        job = Promise.resolve((view as WorkbookView).print(options));
      }
    } catch (error) {
      job = Promise.reject(error);
    }
    job.then(
      () => done('print', id),
      (error) => fail(error && (error as { code?: unknown }).code === 'unavailable' ? 'print-unavailable' : 'print-failed', errorText(error), id),
    );
  };

  /** The id of a message; it is read first so that every later complaint can carry it. */
  const idOf = (fields: Message): string | undefined => {
    if (typeof fields.id === 'string') return fields.id;
    if (fields.id !== undefined && fields.id !== null) throw new BadMessage('"id" must be a string');
    return undefined;
  };

  const handle = async (input: unknown, mine: number) => {
    if (destroyed) return;
    let id: string | undefined;
    try {
      const fields = parseMessage(input);
      id = idOf(fields);
      const type = fields.type;
      if (typeof type !== 'string') throw new BadMessage('"type" is missing');

      if (type === 'load') return await load(fields, id, mine);
      if (type === 'save') return await save(id);
      if (type === 'print') return print(fields, id);
      if (type === 'insertSheet' || type === 'renameSheet' || type === 'deleteSheet') return changeSheets(type, fields, id);
      if (type === 'moveSheet' || type === 'setSheetState' || type === 'copySheet' || type === 'setTabColor') return arrangeSheets(type, fields, id);
      if (type === 'acceptAllRevisions' || type === 'rejectAllRevisions' || type === 'setTrackChanges') return reviseDocument(type, fields, id);
      if (type !== 'setReadOnly' && type !== 'setLocale' && type !== 'undo' && type !== 'redo') {
        throw new BadMessage(`Unknown message type "${type.slice(0, 40)}"`);
      }
      const readOnly = type === 'setReadOnly' ? field<boolean>(fields, 'readOnly', 'boolean', true) : undefined;
      const locale = type === 'setLocale' ? field<string>(fields, 'locale', 'string', true) : undefined;
      if (!shown) return fail('not-loaded', 'No file is open', id);
      const view = shown.view;
      const pages = shown.pages;
      /** Set when the command was heard but could not have its effect: the answer says what holds instead. */
      let forced = false;
      if (type === 'setReadOnly') {
        // A file that cannot be written with changes (doc, RTF, xls) stays read-only whatever is asked.
        forced = shown.locked && readOnly === false;
        const value = shown.locked ? true : (readOnly as boolean);
        if (view) view.setReadOnly(value);
        else if (pages && pages.setReadOnly) pages.setReadOnly(value);
        if (pages) shown.readOnly = value;
      } else if (type === 'setLocale') {
        if (view) view.setLocale(locale as string);
        else if (pages && pages.setLocale) pages.setLocale(locale as string);
        shown.locale = locale;
        setLocale(locale);
      } else if (type === 'undo') {
        if (view) view.undo();
        else if (pages && pages.undo) pages.undo();
      } else if (view) view.redo();
      else if (pages && pages.redo) pages.redo();
      // The command is confirmed as before; a request to make a file editable that cannot be learns that it stays read-only.
      if (forced && id !== undefined) send({ type: 'done', id, command: type, readOnly: true });
      else done(type, id);
    } catch (error) {
      // Whatever went wrong, the page stays alive and the application hears about it.
      fail('bad-message', errorText(error), id);
    }
  };

  /**
   * `hello` and `close` do not wait for their turn: they must work while a load hangs. Nor does
   * `saveResult`: it answers a save of the file that is shown now, not of the one a load that waits will show.
   */
  const handleNow = (fields: Message, type: string) => {
    let id: string | undefined;
    try {
      id = idOf(fields);
      if (type === 'hello') return send({ type: 'ready', version: host.version });
      if (type === 'saveResult') return saveResult(fields, id);
      supersede('the file was closed');
      closeShown();
      setState('empty');
      settleSaveState();
      done('close', id);
    } catch (error) {
      fail('bad-message', errorText(error), id);
    }
  };

  /** Ctrl+S or Cmd+S in the editor of a `load` with `save`: what the Save button does, and not the browser's "save page". */
  const onKeyDown = (event: KeyboardEvent) => {
    if (destroyed || !shown || !shown.saver) return;
    if (!(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey || (event.key !== 's' && event.key !== 'S')) return;
    event.preventDefault();
    shown.saver.run();
  };
  const keyTarget = host.container as { addEventListener?: unknown; removeEventListener?: unknown } | undefined;
  const listensToKeys = !!keyTarget && typeof keyTarget.addEventListener === 'function' && typeof keyTarget.removeEventListener === 'function';
  // While the event goes down: a dialog or a field inside the editor that keeps its keys to itself does not hide this one.
  if (listensToKeys) host.container.addEventListener('keydown', onKeyDown, true);

  const start = () => {
    if (started || destroyed) return;
    started = true;
    send({ type: 'ready', version: host.version });
  };

  return {
    start,
    receive(message: unknown): Promise<void> {
      if (destroyed) return Promise.resolve();
      // A call proves the application is there: `ready` and anything kept get another try.
      const announced = started;
      start();
      flush();
      // A first look, only to see what must not wait; complaints are made in turn, by handle().
      let fields: Message | undefined;
      try {
        fields = parseMessage(message);
      } catch {
        fields = undefined;
      }
      const type = fields && fields.type;
      if (fields && (type === 'hello' || type === 'close' || type === 'saveResult')) {
        // A `hello` that is the first thing to arrive has just been answered by start().
        if (type !== 'hello' || announced) handleNow(fields, type as string);
        return Promise.resolve();
      }
      // A new load overtakes the one before it, whether that one is fetching or still waiting.
      const mine = type === 'load' ? supersede('another file was requested') : generation;
      chain = chain.then(() => handle(fields || message, mine));
      return chain;
    },
    destroy() {
      if (destroyed) return;
      supersede('the bridge was destroyed');
      closeShown();
      if (listensToKeys) host.container.removeEventListener('keydown', onKeyDown, true);
      destroyed = true;
      pending.length = 0;
      pendingChars = 0;
    },
  };
}

export interface InstallBridgeOptions {
  /** The page's window. */
  window: Window;
  container: HTMLElement;
  open: BridgeHost['open'];
  createView: BridgeHost['createView'];
  createDocumentView?: BridgeHost['createDocumentView'];
  createEditor?: BridgeHost['createEditor'];
  createDocumentEditor?: BridgeHost['createDocumentEditor'];
  version: string;
  /**
   * Origins that may embed the page in an iframe and drive it with window
   * messages, separated by white space. Empty or missing: no page may.
   * This comes from the page itself, never from its address.
   */
  parentOrigins?: string;
  onState?: BridgeHost['onState'];
  onLocale?: BridgeHost['onLocale'];
}

/** The ways an application can be listening, as they appear on a window. */
interface ChannelWindow {
  BymOfficeHost?: { postMessage?(json: string): void };
  webkit?: { messageHandlers?: { bymoffice?: { postMessage?(json: string): void } } };
  chrome?: {
    webview?: {
      postMessage?(json: string): void;
      addEventListener?(type: 'message', listener: (event: { data: unknown }) => void): void;
      removeEventListener?(type: 'message', listener: (event: { data: unknown }) => void): void;
    };
  };
  ReactNativeWebView?: { postMessage?(json: string): void };
  flutter_inappwebview?: { callHandler?(name: string, json: string): unknown };
  BymOfficeBridge?: { receive(message: unknown): void; version: string };
}

/** An origin in the form browsers report it ("https://app.example"), or '' when the text is not one. "null" never is. */
export function normalizeOrigin(text: string): string {
  try {
    const url = new URL(text);
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return '';
    return url.origin === 'null' ? '' : url.origin;
  } catch {
    return '';
  }
}

/** Reads a response body, giving up as soon as it is larger than allowed. */
async function readLimited(response: Response, maxBytes: number): Promise<Uint8Array> {
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const tooLarge = () => new Error(`the file is larger than ${maxBytes} bytes`);
  const declared = Number(response.headers.get('content-length'));
  if (declared > maxBytes) throw tooLarge();
  const reader = response.body && typeof response.body.getReader === 'function' ? response.body.getReader() : undefined;
  if (!reader) {
    const whole = new Uint8Array(await response.arrayBuffer());
    if (whole.length > maxBytes) throw tooLarge();
    return whole;
  }
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const step = await reader.read();
    if (step.done) break;
    total += step.value.length;
    if (total > maxBytes) {
      void reader.cancel().catch(() => undefined);
      throw tooLarge();
    }
    chunks.push(step.value);
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (let i = 0; i < chunks.length; i++) {
    out.set(chunks[i], offset);
    offset += chunks[i].length;
  }
  return out;
}

/**
 * Wires a bridge to a page: finds the application's channel for outgoing
 * messages, takes incoming ones through `window.BymOfficeBridge.receive()` and
 * `message` events, and sends `ready`. destroy() of the result undoes all of it.
 *
 * A framed page talks to its parent only when the parent's origin is in
 * `parentOrigins`, and never loads from a URL.
 */
export function installBridge(options: InstallBridgeOptions): Bridge {
  const win = options.window;
  const channels = win as unknown as ChannelWindow;
  const framed = win.parent !== win;

  const allowed: string[] = [];
  const listed = (options.parentOrigins || '').split(/\s+/);
  for (let i = 0; i < listed.length; i++) {
    const origin = normalizeOrigin(listed[i]);
    if (origin && allowed.indexOf(origin) < 0) allowed.push(origin);
  }
  // Where messages for the parent go. Known at once when the browser tells
  // who the parent is or only one origin is allowed; otherwise it is the
  // origin of the first message the parent sends (a `hello`).
  let parentOrigin = '';
  let parentRefused = false;
  if (framed && allowed.length > 0) {
    const ancestors = (win.location as { ancestorOrigins?: { length: number; [index: number]: string } }).ancestorOrigins;
    if (ancestors && ancestors.length > 0) {
      const origin = normalizeOrigin(ancestors[0]);
      if (allowed.indexOf(origin) >= 0) parentOrigin = origin;
      else parentRefused = true;
    } else if (allowed.length === 1) {
      parentOrigin = allowed[0];
    }
  }

  // The first channel that exists gets the message. The objects are looked up
  // on every call because applications may inject them after the page loaded,
  // and their methods are called in place: some of them need their `this`.
  const post = (json: string): boolean => {
    const own = channels.BymOfficeHost;
    if (own && typeof own.postMessage === 'function') return own.postMessage(json), true;
    const handlers = channels.webkit && channels.webkit.messageHandlers;
    if (handlers && handlers.bymoffice && typeof handlers.bymoffice.postMessage === 'function') return handlers.bymoffice.postMessage(json), true;
    // Before WebView2's own channel: on Windows this plugin runs on top of WebView2.
    const inApp = channels.flutter_inappwebview;
    if (inApp && typeof inApp.callHandler === 'function') return inApp.callHandler('bymoffice', json), true;
    const webview = channels.chrome && channels.chrome.webview;
    if (webview && typeof webview.postMessage === 'function') return webview.postMessage(json), true;
    const native = channels.ReactNativeWebView;
    if (native && typeof native.postMessage === 'function') return native.postMessage(json), true;
    if (framed && parentOrigin) return win.parent.postMessage(json, parentOrigin), true;
    return false;
  };

  const fetchBytes =
    typeof win.fetch === 'function'
      ? (url: string, limits: { signal?: AbortSignal; maxBytes: number }) =>
          win.fetch(url, limits.signal ? { signal: limits.signal } : undefined).then((response) => readLimited(response, limits.maxBytes))
      : undefined;

  const bridge = createBridge({
    container: options.container,
    open: options.open,
    createView: options.createView,
    createDocumentView: options.createDocumentView,
    createEditor: options.createEditor,
    createDocumentEditor: options.createDocumentEditor,
    version: options.version,
    post,
    baseUrl: win.location.href,
    // A framed page would fetch with its own origin's cookies for whoever embeds it.
    allowUrl: !framed,
    fetchBytes,
    onState: options.onState,
    onLocale: options.onLocale,
    frame: typeof win.requestAnimationFrame === 'function' ? (callback) => void win.requestAnimationFrame(callback) : undefined,
    timer: (callback, ms) => void win.setTimeout(callback, ms),
  });

  // Nothing is returned to the caller: some WebViews report an error when the
  // evaluated script yields a value they cannot pass on, such as a promise.
  const entry = {
    receive(message: unknown) {
      void bridge.receive(message);
    },
    version: options.version,
  };
  channels.BymOfficeBridge = entry;

  // Only an allowed embedding page may talk to a framed editor; a top-level page takes no window messages at all.
  const onWindowMessage = (event: MessageEvent) => {
    if (!framed || parentRefused || event.source !== win.parent) return;
    const origin = normalizeOrigin(event.origin);
    if (!origin || allowed.indexOf(origin) < 0) return;
    if (parentOrigin && origin !== parentOrigin) return;
    parentOrigin = origin;
    void bridge.receive(event.data);
  };
  win.addEventListener('message', onWindowMessage);
  // WebView2 delivers the application's messages on its own object, not on the window.
  const webview = channels.chrome && channels.chrome.webview;
  const onWebViewMessage = (event: { data: unknown }) => void bridge.receive(event.data);
  if (webview && typeof webview.addEventListener === 'function') webview.addEventListener('message', onWebViewMessage);

  bridge.start();
  return {
    start: bridge.start,
    receive: bridge.receive,
    destroy() {
      bridge.destroy();
      win.removeEventListener('message', onWindowMessage);
      if (webview && typeof webview.removeEventListener === 'function') webview.removeEventListener('message', onWebViewMessage);
      if (channels.BymOfficeBridge === entry) delete channels.BymOfficeBridge;
    },
  };
}
