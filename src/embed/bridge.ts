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
import { pdfFileName, type PdfExportOptions } from '../pdf/export';
import type { PdfLinkReport } from '../pdf/links';
import type { SheetPdfExportOptions, WorkbookView, WorkbookViewOptions } from '../ui/sheet-view';
import { resolveUiStrings } from '../ui/strings';
import { parsePrintArea, parseTitleBand, type PageSetupPatch } from '../ops/page-setup';
import type { SectionPatch } from '../ops/doc/section';
import { askOverwrite, editorRootIn } from '../ui/save-confirm';
import { readPivotCell, readPivotSource } from '../ui/pivot-dialog';
import { PasswordPrompt, type PasswordRequest } from '../ui/password-dialog';
import { readCollabStamp, writeSharedFile, type CollabStamp } from '../formats/xlsx/collab-stamp';
import { listPeople } from '../collab/people';
import { packageTurn } from '../package-turn';
import { CollabSession, contentVersion, isSocketUrl, socketTransport, type CollabNotice, type CollabTransport, type CollabUser } from '../collab/index';
import { DocCollabSession, type DocChange, type DocPeer } from '../collab/doc/session';
import { holdsRevisions } from '../collab/doc/history';
import { SaveController, SaveError, resolveAutoSave, type AutoSaveSettings, type ConflictQuestion, type SaveInfo } from './save-to';

export type BridgeErrorCode =
  | 'unsupported'
  | 'corrupt'
  | 'load-failed'
  | 'password-cancelled'
  | 'password-not-expected'
  | 'save-failed'
  | 'bad-message'
  | 'not-loaded'
  | 'print-unavailable'
  | 'print-failed'
  | 'pdf-unavailable'
  | 'pdf-failed'
  | 'pdf-unsupported'
  | 'pdf-too-large'
  | 'read-only'
  | 'no-sheet'
  | 'sheet-refused'
  | 'locale-refused'
  | 'no-document'
  | 'comment-refused'
  | 'pivot-refused'
  | 'page-setup-refused'
  /** The protection of a sheet, or of the structure of the workbook, refuses the command. */
  | 'protected';

/**
 * What the page shows around the editor: nothing open, a file on its way, a file open, a file that failed;
 * "password": a file on its way waits for its password (the question is asked inside the container).
 */
export type BridgeState = 'empty' | 'loading' | 'open' | 'error' | 'password';

/**
 * The longest password a message may carry, in characters. More than the 255 that a password which is SET may
 * have (MAX_PASSWORD_CHARS): a file is opened with whatever password it has, and saved with it again; the bridge
 * has no message that sets a new one.
 */
const MAX_PASSWORD_FIELD = 1024;

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
  /** The root element of the view: what is asked of the user about the file is asked over the editor it belongs to. */
  readonly element?: HTMLElement;
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
  /**
   * The document as a PDF file. An error with `code: 'unavailable'` says that
   * no PDF can be made here.
   */
  exportPdf?(options?: PdfExportOptions): Promise<Uint8Array>;
  /** The same, as the parts the file was written in: what the bridge asks for when the view has it, so that the file is never held twice. */
  exportPdfParts?(options?: PdfExportOptions): Promise<Uint8Array[]>;
  /** The tracked changes of the document: their number is what the application is told. */
  revisions?(): unknown[];
  /** Accept or reject every tracked change; a view that cannot (a protection) changes nothing. */
  acceptAllRevisions?(): unknown;
  rejectAllRevisions?(): unknown;
  /** Turns the tracking of changes on or off; a view that cannot changes nothing. */
  setTrackChanges?(on: boolean): unknown;
  /** Whether the changes made in the view are tracked. */
  readonly trackingChanges?: boolean;
  /** Sets and clears tab stops of the paragraphs of the selection (see TabStopsPatch); whether anything changed. */
  setTabStops?(patch: unknown): boolean;
  /** The tab stops of the paragraph at the caret, positions in twips. */
  readonly tabStops?: unknown[];
  /** Sets the distance between the automatic tab stops of the document, twips; whether it changed. */
  setDefaultTabStop?(twips: number): boolean;
  /** Changes the paper, orientation or margins of the sections of the selection or of the whole document (see SectionPatch); whether anything changed. */
  setSectionSetup?(patch: unknown, scope?: 'section' | 'document'): boolean;
  /** A section break at the caret; whether one was put in. */
  insertSectionBreak?(kind?: string): boolean;
  /** The page of the section of the caret, in twips, and how many sections there are. */
  readonly sectionSetup?: unknown;
  readonly defaultTabStop?: number;
  /** The comments of the document, answers after the comment they answer: what the application is told. */
  comments?(): BridgeComment[];
  /** Changes whenever what comments() says of the comments changes: while it stays the same the list is not asked for again. */
  readonly commentsRevision?: number;
  /** Whether comments can be added and changed (also true where a protection allows comments and nothing else). */
  readonly canComment?: boolean;
  /** Adds a comment about the selection, or about the range or place given; the id of the comment, or undefined when none was added. */
  addComment?(text: string, at?: unknown): number | undefined;
  editComment?(id: number, text: string): boolean;
  /** Answers a comment; the id of the answer, or undefined. */
  replyToComment?(id: number, text: string): number | undefined;
  resolveComment?(id: number, done?: boolean): boolean;
  removeComment?(id: number): boolean;
  /** Whether the panel that lists the comments beside the pages is shown. */
  readonly commentsPanelOpen?: boolean;
  /** Shows or hides that panel; whether it is shown afterwards. */
  setCommentsPanel?(open: boolean, focus?: boolean): boolean;
  /** Says something to the user: that the document was saved, or why it was not. */
  showNotice?(message: string, kind?: 'error' | 'info'): void;
  /** Whether an input method composes or something is dragged: auto-save waits meanwhile. */
  hasPendingInput?(): boolean;
  /** A document that is edited together: the view is tied to its session, and told what the session says (see DocumentView). */
  attachCollab?(session: DocCollabSession | undefined): void;
  collabChanged?(change: DocChange): void;
  collabNotice?(notice: CollabNotice): void;
  collabPeersChanged?(peers: DocPeer[]): void;
  readonly collabStrings?: { guest(n: number): string };
}

/** A comment as a document view lists it; the bridge passes on the fields the contract names. */
export interface BridgeComment {
  id: number;
  author: string;
  date?: string;
  text: string;
  done: boolean;
  parent?: number;
}

export interface BridgeHost {
  /** The element the editor is shown in. */
  container: HTMLElement;
  /** The engine's open(). */
  open(bytes: Uint8Array, options?: { password?: string; signal?: { readonly aborted: boolean } }): Promise<OfficeFile>;
  /**
   * Makes the question for the password of a password-protected file,
   * asked inside the container next to the `passwordRequired` message: the
   * user's answer and the application's `password` message count alike,
   * whichever comes first. Without it only the application is asked.
   */
  passwordPrompt?(container: HTMLElement, locale: string | undefined, onCancel: () => void): { ask(request: PasswordRequest): Promise<string | null>; close(): void };
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
  /**
   * What carries the messages of a `load` with `collab` to its relay. Default: a WebSocket to the
   * address (socketTransport()).
   */
  createTransport?(url: string): CollabTransport;
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
   * Asks the user whether a save the application answered with `conflict` is to overwrite what
   * the application holds, in the language of the file that is shown, and returns what takes the
   * question away again. Without it nothing is asked: such a save has failed, and the view says so.
   * `element` is the root element of the view that is shown, when the view names one. Nothing is
   * asked about a file that is only read.
   */
  askOverwrite?(question: ConflictQuestion, locale: string | undefined, element: HTMLElement | undefined): (() => void) | void;
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
  /**
   * Runs a callback after a delay. Default: setTimeout. It may return what calls the timer off
   * (as clearTimeout does): the waits that end early are then really ended. Without one the page
   * only lets go of what the callback held, and the timer runs out empty.
   */
  timer?(callback: () => void, ms: number): void | (() => void);
  /** The clock auto-save counts with, in milliseconds. Default: Date.now. */
  now?(): number;
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
  /**
   * Says that the page was hidden (the application went to the background, the WebView is being
   * closed): a file loaded with `save: { auto }` hands what waits to the application at once.
   */
  pageHidden(): void;
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
 * Bytes of a PDF one `pdf` message carries at most (as base64 they are a
 * third more). A message is one string, which the page builds, the channel
 * copies and the application decodes: beyond this the answer is the error
 * `pdf-too-large`, and the application asks for a lower resolution.
 */
export const MAX_BRIDGE_PDF_BYTES = 64 * 1024 * 1024;

/**
 * The parts of a file as the base64 text of the whole file, without joining
 * them into one array first: every part is encoded where three bytes end,
 * and the one or two bytes left over are carried into the next part.
 */
export function base64OfParts(parts: readonly Uint8Array[]): string {
  const pieces: string[] = [];
  const carry = new Uint8Array(3);
  let carried = 0;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!(part instanceof Uint8Array)) throw new TypeError('A part of the file is no array of bytes');
    let at = 0;
    // Fill what was left over up to three bytes.
    while (carried > 0 && carried < 3 && at < part.length) carry[carried++] = part[at++];
    if (carried === 3) {
      pieces.push(base64Encode(carry));
      carried = 0;
    }
    const whole = at + Math.floor((part.length - at) / 3) * 3;
    if (whole > at) pieces.push(base64Encode(part.subarray(at, whole)));
    for (at = whole; at < part.length; at++) carry[carried++] = part[at];
  }
  if (carried > 0) pieces.push(base64Encode(carry.subarray(0, carried)));
  return pieces.join('');
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
const isStatus = (type: string) => type === 'changed' || type === 'selection' || type === 'sheet' || type === 'sheets' || type === 'saveState' || type === 'collabState';

/** Longest name of a sheet a message may carry. What a sheet may be called is the view's to say; this only bounds what is looked at. */
const MAX_NAME_CHARS = 255;

/** A field that counts sheets: a whole number from 0 that a workbook could have as an index. */
function sheetIndexField(message: Message, name: string, required: boolean): number | undefined {
  const value = field<number>(message, name, 'number', required);
  if (value !== undefined && !(Number.isInteger(value) && value >= 0 && value <= 0x7fffffff)) throw new BadMessage(`"${name}" must be a sheet index`);
  return value;
}

/** A list of numbers by name as a message carries it (margins), each a number; undefined when the message has none. */
function numbersField(message: Message, name: string, keys: readonly string[]): { [key: string]: number } | undefined {
  const value = message[name];
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'object' || Array.isArray(value)) throw new BadMessage(`"${name}" must be an object`);
  const out: { [key: string]: number } = {};
  for (const key of keys) {
    const entry = (value as Record<string, unknown>)[key];
    if (entry === undefined || entry === null) continue;
    if (typeof entry !== 'number' || !isFinite(entry)) throw new BadMessage(`"${name}.${key}" must be a number`);
    out[key] = entry;
  }
  return out;
}

/** The change of the page setup of a sheet a `setPageSetup` message asks for (see PageSetupPatch); what the values may be is checked by the view. */
function sheetPageSetupPatch(message: Message): PageSetupPatch {
  const patch: PageSetupPatch = {};
  const paperSize = field<number>(message, 'paperSize', 'number', false);
  if (paperSize !== undefined) patch.paperSize = paperSize;
  const orientation = field<string>(message, 'orientation', 'string', false);
  if (orientation !== undefined) {
    if (orientation !== 'portrait' && orientation !== 'landscape') throw new BadMessage('"orientation" must be "portrait" or "landscape"');
    patch.orientation = orientation;
  }
  const margins = numbersField(message, 'margins', ['left', 'right', 'top', 'bottom', 'header', 'footer']);
  if (margins) patch.margins = margins;
  const scale = field<number>(message, 'scale', 'number', false);
  if (scale !== undefined) patch.scale = scale;
  const fit = numbersField(message, 'fitTo', ['width', 'height']);
  if (fit) {
    if (fit.width === undefined || fit.height === undefined) throw new BadMessage('"fitTo" must have a width and a height');
    patch.fitTo = { width: fit.width, height: fit.height };
  }
  for (const name of ['horizontalCentered', 'verticalCentered', 'gridLines', 'headings'] as const) {
    const value = field<boolean>(message, name, 'boolean', false);
    if (value !== undefined) patch[name] = value;
  }
  // References as text ("A1:D20", "$1:$2"); null or '' clears.
  const text = (name: string): string | null | undefined => {
    if (message[name] === null) return null;
    const value = field<string>(message, name, 'string', false);
    if (value !== undefined && value.length > 4096) throw new BadMessage(`"${name}" is too long`);
    return value;
  };
  const area = text('printArea');
  if (area !== undefined) {
    const areas = area === null ? null : parsePrintArea(area);
    if (areas === undefined) throw new BadMessage('"printArea" must be references such as "A1:D20"');
    patch.printArea = areas;
  }
  for (const [name, axis] of [['titleRows', 'row'], ['titleCols', 'col']] as const) {
    const value = text(name);
    if (value === undefined) continue;
    const band = value === null ? null : parseTitleBand(value, axis);
    if (band === undefined) throw new BadMessage(`"${name}" must be ${axis === 'row' ? 'rows such as "$1:$2"' : 'columns such as "$A:$B"'}`);
    patch[name] = band;
  }
  return patch;
}

/** The change of the page of sections a `setPageSetup` message asks for of a document (see SectionPatch): paper by name, margins in twips. */
function sectionPatch(message: Message): SectionPatch {
  const patch: SectionPatch = {};
  const paper = field<string>(message, 'paper', 'string', false);
  if (paper !== undefined) {
    if (paper.length > 32) throw new BadMessage('"paper" is not a paper');
    patch.paper = paper;
  }
  const orientation = field<string>(message, 'orientation', 'string', false);
  if (orientation !== undefined) {
    if (orientation !== 'portrait' && orientation !== 'landscape') throw new BadMessage('"orientation" must be "portrait" or "landscape"');
    patch.orientation = orientation;
  }
  const margins = numbersField(message, 'margins', ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter']);
  if (margins) patch.margins = margins;
  return patch;
}

/** What a view says of an error, in its language; the error's own text where it says nothing. */
/** The code of a refusal: "protected" for what the protection of a sheet or of the workbook refuses, the command's own code otherwise. */
function refusalCode(error: unknown, code: BridgeErrorCode): BridgeErrorCode {
  return error !== null && typeof error === 'object' && (error as { code?: unknown }).code === 'protected' ? 'protected' : code;
}

function refusalText(view: WorkbookView, error: unknown): string {
  try {
    return typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
  } catch {
    return errorText(error);
  }
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

/** Most comments one message lists; a document with more says so (`more`). */
const MAX_LISTED_COMMENTS = 2000;
/** Most of the text of one comment a message carries, UTF-16 code units; a longer one is cut and says so (`truncated`). */
const MAX_LISTED_TEXT = 2000;
/** Most of the name of an author a message carries: what `load.author` may be. */
const MAX_LISTED_AUTHOR = 255;
/** Most text (names, dates and comments together) one list carries; it ends where more would not fit, and says so (`more`). */
const MAX_LISTED_CHARS = 500000;

/** The first `limit` code units of a text, without half of a surrogate pair at the end. */
function cutText(text: string, limit: number): string {
  if (text.length <= limit) return text;
  const code = text.charCodeAt(limit - 1);
  return text.slice(0, code >= 0xd800 && code <= 0xdbff ? limit - 1 : limit);
}

/** A number for a list of comments as it is told: lists that are told differently have, as good as always, different numbers. */
function listDigest(list: Message[]): string {
  let hash = 0x811c9dc5;
  const mix = (text: string): void => {
    for (let i = 0; i < text.length; i++) hash = Math.imul(hash ^ text.charCodeAt(i), 0x01000193);
    hash = Math.imul(hash ^ 0x1f, 0x01000193);
  };
  for (const entry of list) {
    mix(String(entry.id));
    mix(String(entry.author));
    mix(String(entry.text));
    mix(`${entry.done} ${entry.parent} ${entry.date} ${entry.truncated}`);
  }
  return `${list.length}:${hash >>> 0}`;
}
/** Longest text of a comment a command may bring, UTF-16 code units: what the engine takes. */
const MAX_COMMENT_CHARS = 32767;

/** What a document view says of its comments, as the contract has it; undefined for a view that knows nothing of them. */
function commentState(pages: BridgeDocumentView): { list: Message[]; more: boolean } | undefined {
  if (typeof pages.comments !== 'function') return undefined;
  let all: BridgeComment[] = [];
  try {
    const listed = pages.comments();
    all = Array.isArray(listed) ? listed : [];
  } catch {
    all = [];
  }
  const list: Message[] = [];
  let room = MAX_LISTED_CHARS;
  let more = false;
  for (let i = 0; i < all.length; i++) {
    const comment = all[i];
    if (!comment || typeof comment.id !== 'number') continue;
    if (list.length >= MAX_LISTED_COMMENTS) {
      more = true;
      break;
    }
    // What a file brings is passed on in pieces of a size the application can count on.
    const whole = typeof comment.text === 'string' ? comment.text : '';
    const text = cutText(whole, MAX_LISTED_TEXT);
    const author = cutText(typeof comment.author === 'string' ? comment.author : '', MAX_LISTED_AUTHOR);
    const date = typeof comment.date === 'string' ? cutText(comment.date, 64) : undefined;
    room -= text.length + author.length + (date ? date.length : 0) + 64;
    if (room < 0) {
      more = true;
      break;
    }
    const out: Message = { id: comment.id, author, text, done: comment.done === true };
    if (date !== undefined) out.date = date;
    if (typeof comment.parent === 'number') out.parent = comment.parent;
    if (text.length < whole.length) out.truncated = true;
    list.push(out);
  }
  return { list, more };
}

/** A place of a document in a message: `{ story?, path, offset }`. Throws BadMessage for anything else. */
function positionField(message: Message, name: string): { story: string; path: number[]; offset: number } | undefined {
  const value = message[name];
  if (value === undefined || value === null) return undefined;
  const place = value as { story?: unknown; path?: unknown; offset?: unknown };
  if (typeof value !== 'object' || Array.isArray(value)) throw new BadMessage(`"${name}" must be a position`);
  const story = place.story === undefined ? 'body' : place.story;
  const path = place.path;
  if (typeof story !== 'string' || story.length > 64) throw new BadMessage(`"${name}.story" must be the name of a story`);
  if (!Array.isArray(path) || path.length === 0 || path.length > 256 || !path.every((step) => Number.isInteger(step) && step >= -1 && step <= 0x7fffffff)) throw new BadMessage(`"${name}.path" must be a path`);
  if (!(Number.isInteger(place.offset) && (place.offset as number) >= 0 && (place.offset as number) <= 0x7fffffff)) throw new BadMessage(`"${name}.offset" must be an offset`);
  return { story, path: path.slice() as number[], offset: place.offset as number };
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
  /** A document: a digest of the list of comments the application was last told (see listDigest), not the list. */
  commentsSent?: string;
  /** A document: what the view said its comments were at (BridgeDocumentView.commentsRevision) when they were last looked at. */
  commentsRevision?: number;
  /** A document whose protection allows comments and nothing else. */
  commentsOnly?: boolean;
  /** A `load` with `save`: what the Save button and Ctrl+S of the page run. */
  saver?: SaveController;
  /** A `load` with `collab`: the session the workbook or the document is edited together in. */
  collab?: CollabSession | DocCollabSession;
  /** The file that was written last while it is edited together, and the step it holds (undefined: of no one step). */
  written?: { bytes: Uint8Array; rev: number | undefined };
  /** The language the file was loaded with or `setLocale` chose: what a save is spoken of in. */
  locale?: string;
}

/** The message handling of the bridge page, with everything it needs from the page handed in. */
export function createBridge(host: BridgeHost): Bridge {
  const timer =
    host.timer ||
    ((callback: () => void, ms: number): (() => void) => {
      const handle = setTimeout(callback, ms);
      return () => clearTimeout(handle);
    });
  const frame = host.frame || ((callback: () => void) => void timer(callback, 16));
  /**
   * A timer that can be called off whatever the host's can do: called off, it lets go of its callback
   * (and of everything that callback holds) at once, and ends the host's timer when the host said how.
   */
  const cancellable = (callback: () => void, ms: number): (() => void) => {
    let waiting: (() => void) | undefined = callback;
    const cancel = timer(() => {
      const run = waiting;
      waiting = undefined;
      if (run) run();
    }, ms);
    return () => {
      if (!waiting) return;
      waiting = undefined;
      if (typeof cancel === 'function') cancel();
    };
  };
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
  let awaited: { saveId: string; resolve(): void; reject(error: unknown): void; stop(): void } | undefined;
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
        waiting.stop();
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
    if (shown.collab) shown.collab.leave();
    if (awaited) awaited.stop();
    awaited = undefined;
    // The password of a protected file is not kept longer than the file is shown.
    if (typeof shown.file.close === 'function') shown.file.close();
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
  /** A load that waits for the password of its file: what answers it, and what calls the check of a password off. */
  let passwordAsked: { answer(value: string | null): void; signal: { aborted: boolean }; close(): void } | undefined;

  const supersede = (reason: string): number => {
    generation++;
    if (cancelFetch) cancelFetch(reason);
    // A load that waits for a password, or checks one, is over: nobody is asked any more.
    if (passwordAsked) {
      const asked = passwordAsked;
      passwordAsked = undefined;
      asked.signal.aborted = true;
      asked.close();
      asked.answer(null);
    }
    return generation;
  };

  /** The `password` message: the application's answer to `passwordRequired`. Its value is never sent back, in no message. */
  const passwordGiven = (fields: Message, id: string | undefined) => {
    const value = fields.value;
    if (value !== null && typeof value !== 'string') throw new BadMessage('"value" must be the password as a string, or null to give up');
    if (typeof value === 'string' && value.length > MAX_PASSWORD_FIELD) throw new BadMessage('"value" is too long');
    // Nobody is asking: no file is loading, or the password that was given last is still being checked. The message is
    // not kept for a question that may come: the application answers the `passwordRequired` it gets, each one once.
    if (!passwordAsked) return fail('password-not-expected', 'No password is being asked for: answer a "passwordRequired" message, once', id);
    const asked = passwordAsked;
    // The application answered: the question inside the page is taken away (a wrong password brings it back).
    asked.close();
    asked.answer(value);
    done('password', id);
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

  /**
   * A file that was opened and is not the one that is shown in the end (the load was overtaken, the conversion or
   * the view failed) is closed: the password of a protected file is not kept for a file nobody sees.
   */
  const load = async (message: Message, id: string | undefined, mine: number) => {
    const held: OfficeFile[] = [];
    try {
      return await loadFile(message, id, mine, held);
    } finally {
      for (let i = 0; i < held.length; i++) {
        const file = held[i];
        if ((!shown || shown.file !== file) && typeof file.close === 'function') file.close();
      }
    }
  };

  const loadFile = async (message: Message, id: string | undefined, mine: number, held: OfficeFile[]) => {
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
    const commentsPanel = field<string>(message, 'commentsPanel', 'string', false);
    if (commentsPanel !== undefined && commentsPanel !== 'auto' && commentsPanel !== 'open' && commentsPanel !== 'closed') throw new BadMessage('"commentsPanel" must be "auto", "open" or "closed"');
    const password = field<string>(message, 'password', 'string', false);
    const passwordPrompt = field<boolean>(message, 'passwordPrompt', 'boolean', false);
    if (password !== undefined && password.length > MAX_PASSWORD_FIELD) throw new BadMessage('"password" is too long');
    // `save`: true, or an object that also says whether the page saves by itself ({ auto }).
    const saveField = message.save;
    let userSaves = false;
    let autoSave: AutoSaveSettings | undefined;
    if (typeof saveField === 'object' && saveField !== null && !Array.isArray(saveField)) {
      userSaves = true;
      try {
        autoSave = resolveAutoSave((saveField as { auto?: unknown }).auto, 'save.auto');
      } catch (error) {
        throw new BadMessage(errorText(error));
      }
    } else if (saveField !== undefined && saveField !== null) {
      if (typeof saveField !== 'boolean') throw new BadMessage('"save" must be a boolean or an object ({ auto })');
      userSaves = saveField;
    }
    // `collab`: the address of a relay, or an object with it and who this participant is.
    let collab: { url: string; room?: string; token?: string; user?: CollabUser } | undefined;
    let stamp: CollabStamp | undefined;
    const collabField = message.collab;
    if (typeof collabField === 'string') collab = { url: collabField };
    else if (typeof collabField === 'object' && collabField !== null && !Array.isArray(collabField)) {
      const given = collabField as { url?: unknown; room?: unknown; token?: unknown; user?: unknown };
      if (typeof given.url !== 'string') throw new BadMessage('"collab.url" must be the address of a relay');
      collab = { url: given.url };
      if (given.room !== undefined) {
        if (typeof given.room !== 'string' || given.room.length > 200) throw new BadMessage('"collab.room" must be a short text');
        collab.room = given.room;
      }
      if (given.token !== undefined) {
        if (typeof given.token !== 'string' || given.token.length > 4096) throw new BadMessage('"collab.token" must be a text');
        collab.token = given.token;
      }
      if (typeof given.user === 'object' && given.user !== null) {
        const user = given.user as { name?: unknown; color?: unknown };
        collab.user = {};
        if (typeof user.name === 'string') collab.user.name = user.name;
        if (typeof user.color === 'string') collab.user.color = user.color;
      }
    } else if (collabField !== undefined && collabField !== null) throw new BadMessage('"collab" must be the address of a relay or an object ({ url, room, user, token })');
    if (collab && !isSocketUrl(collab.url)) throw new BadMessage('"collab" must be the address of a relay: "ws://…" or "wss://…"');
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
      // A password-protected file: the password of the `load` is tried, then the application and the user are asked
      // (`passwordRequired`, and the question inside the page) for as long as the password is wrong.
      const signal = { aborted: false };
      let prompt: { ask(request: PasswordRequest): Promise<string | null>; close(): void } | undefined;
      let given = password;
      let attempts = 0;
      try {
        for (;;) {
          try {
            // One derivation for each password: a password that was given is tried alone (see OpenOptions.password).
            file = await host.open(bytes, given === undefined ? { signal } : { password: given, signal });
            held.push(file);
            break;
          } catch (error) {
            const code = error instanceof Error && error.name === 'PasswordError' ? (error as { code?: unknown }).code : undefined;
            if (overtaken()) return;
            // The user gave up in the question of the page while a password was being checked.
            if (code === 'cancelled') return refuse('password-cancelled', 'The file was not opened: no password was given for it');
            if (code !== 'required' && code !== 'wrong') throw error;
            if (code === 'wrong') attempts++;
            const request: PasswordRequest = { attempts, wrong: code === 'wrong' };
            const answer = await new Promise<string | null>((resolve) => {
              let over = false;
              const asked = {
                signal,
                close: () => {
                  if (prompt) prompt.close();
                  prompt = undefined;
                },
                answer: (value: string | null) => {
                  if (over) return;
                  over = true;
                  if (passwordAsked === asked) passwordAsked = undefined;
                  resolve(value);
                },
              };
              passwordAsked = asked;
              const out: Message = { type: 'passwordRequired', wrong: request.wrong, attempts };
              if (id !== undefined) out.id = id;
              const format = (error as { format?: unknown }).format;
              if (typeof format === 'string') out.format = format;
              send(out);
              if (passwordPrompt !== false && host.passwordPrompt) {
                try {
                  if (!prompt) {
                    // Giving up there also ends the check of a password that is under way.
                    prompt = host.passwordPrompt(host.container, locale, () => {
                      signal.aborted = true;
                    });
                  }
                  // The question is over the editor, or over an empty page: the page shows the container for it.
                  if (!shown) setState('password');
                  prompt.ask(request).then(asked.answer, () => asked.answer(null));
                } catch {
                  // No question inside the page: the application is asked alone.
                }
              }
            });
            if (overtaken()) return;
            if (answer === null) return refuse('password-cancelled', 'The file was not opened: no password was given for it');
            if (!shown) setState('loading');
            given = answer;
          }
        }
      } finally {
        if (prompt) prompt.close();
      }
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
          // The converted file took the password of its original over: the original lets go of it.
          if (typeof file.close === 'function') file.close();
          file = opened.file;
          held.push(file);
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
        // What the file says of the room it was saved in, for a document that is to be edited together.
        if (collab && file.package) stamp = await readCollabStamp(file.package);
      } else {
        workbook = await file.workbook();
        // What the file says of the room it was saved in, for a workbook that is to be edited together.
        if (collab && file.package) stamp = await readCollabStamp(file.package);
        if (workbook.sheets.length === 0) throw new Error('The workbook has no sheets');
      }
    } catch (error) {
      if (overtaken()) return;
      // An encryption that is not read is a file that is not supported; an encrypted file that is damaged is a damaged file.
      const unsupported = error instanceof Error && (error.name === 'UnsupportedFileError' || (error.name === 'PasswordError' && (error as { code?: unknown }).code === 'unsupported'));
      return refuse(unsupported ? 'unsupported' : 'corrupt', errorText(error));
    }
    if (overtaken()) return;
    const fontFailures = await fontsSettled;
    if (overtaken()) return;
    for (let i = 0; i < fontFailures.length; i++) {
      const failure = fontFailures[i];
      const text = failure.pending ? `The font "${failure.family}" is still loading; the file is shown without it until it arrives` : `The font "${failure.family}" could not be loaded: ${failure.reason}`;
      (warnings || (warnings = [])).push(text);
    }
    // What the document enforces: a protection for forms allows no editing here, as the read-only one. One for
    // comments allows comments: the text stays as it is (the view sees to that), and the file is written back.
    const protection = doc ? enforcedProtection(doc) : undefined;
    // Only docx and xlsx are written back with their changes; the other formats stay read-only until they are converted.
    const locked = needsConversion(file.format) || (protection !== undefined && protection !== 'trackedChanges' && protection !== 'comments');
    const commentsOnly = !locked && protection === 'comments';

    const loaded = (activeSheet: number, sheets: unknown[], document: boolean) => {
      const out: Message = { type: 'loaded' };
      if (id !== undefined) out.id = id;
      out.format = file.format;
      if (document) out.kind = 'document';
      // The file was opened with a password: what `saved` hands over is encrypted with it again.
      if (file.passwordProtected === true) out.encrypted = true;
      // With the password everyone knows: the file is encrypted and keeps nothing secret.
      if (file.encryption && file.encryption.defaultPassword === true) out.defaultPassword = true;
      // Encrypted without an integrity check: its content could have been changed by someone who does not know the password.
      if (file.warnings && file.warnings.some((warning) => warning.code === 'no-integrity')) out.integrity = false;
      out.sheets = sheets;
      out.activeSheet = activeSheet;
      // The file is shown read-only whatever was asked for, and `setReadOnly` will not change that.
      if (locked) out.readOnly = true;
      // The page has a Save button and takes Ctrl+S: the application will be handed what the user saves.
      if (shown && shown.saver) out.save = true;
      // And saves by itself: such a file arrives in `saved` messages that say `by: "auto"`.
      if (shown && shown.saver && shown.saver.autoSave) out.autoSave = true;
      if (convertedFrom !== undefined) out.convertedFrom = convertedFrom;
      if (warnings !== undefined) out.warnings = warnings;
      if (doc) {
        // Why the document is read-only, or that its changes are tracked.
        if (protection !== undefined) out.protection = protection;
        if (protection !== undefined && hasPermissionRanges(doc)) out.protectionExceptions = true;
        // The text cannot be edited; comments can be added and changed, and the file is saved with them.
        if (commentsOnly) out.commentsOnly = true;
        if (!locked && (trackChanges === true || tracksChanges(doc))) out.trackChanges = true;
        // How many tracked changes it has, when it has any.
        const state = shown && shown.pages ? revisionState(shown.pages) : undefined;
        if (state && state.count > 0) out.revisions = state.count;
        if (shown && state) shown.revisionsSent = `${state.count} ${state.tracking}`;
        // Its comments, when it has any.
        const comments = shown && shown.pages ? commentState(shown.pages) : undefined;
        if (comments && comments.list.length > 0) {
          out.comments = comments.list;
          if (comments.more) out.commentsMore = true;
        }
        if (shown && comments) shown.commentsSent = listDigest(comments.list);
        if (shown && shown.pages && typeof shown.pages.commentsRevision === 'number') shown.commentsRevision = shown.pages.commentsRevision;
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
          // And the comments, when the list is not the one that was told last.
          tellComments(shown);
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
      if (commentsPanel !== undefined) documentOptions.commentsPanel = commentsPanel as 'auto' | 'open' | 'closed';
      const saver = userSaves === true && !locked ? userSaver(autoSave) : undefined;
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
      shown = { file, view: undefined, pages, name, locked, convertedFrom, readOnly: locked || readOnly === true, saver, locale, commentsOnly };
      setLocale(locale);
      setState('open');
      loaded(0, [], true);
      // Only a docx as it was loaded is edited together, and not one that tracks its changes (a tracked change is
      // not the same change for two participants): any other is shown as without the field, and the application is told.
      if (collab) joinDocCollab(shown, pages, doc as DocModel, collab, file.format !== 'docx' || convertedFrom !== undefined || trackChanges === true || holdsRevisions(doc as DocModel) ? undefined : bytes, locked || readOnly === true, stamp);
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
        // What someone else changed is there to be saved, but is theirs to save by itself.
        if (shown && shown.view === view && shown.saver) shown.saver.changed(change.remote === true);
        tellSheets();
        // A view that is edited together answers for the steps of its participant; any other view by its history.
        const shared = typeof view.canUndo === 'boolean';
        send({ type: 'changed', sheets: change.sheets, canUndo: shared ? view.canUndo : view.history.canUndo, canRedo: shared ? view.canRedo : view.history.canRedo });
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
    // What "&F" of a header or footer stands for in what the view prints, previews and makes a PDF of; and the name of that PDF.
    if (name !== undefined) options.fileName = name;

    const saver = userSaves === true && !locked ? userSaver(autoSave) : undefined;
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
    if (collab) joinCollab(shown, view, book, collab, file.format !== 'xlsx' || convertedFrom !== undefined ? undefined : bytes, locked || readOnly === true, stamp);
  };

  /**
   * Makes the workbook that is shown one that is edited together (a `load` with `collab`) and
   * tells the application of the session with `collabState`: at once, and whenever its state or
   * the list of the others changes. `bytes`: the file as it was loaded, undefined for one that
   * cannot be edited together (not an xlsx, or converted to be one here).
   */
  const joinCollab = (current: Shown, view: WorkbookView, book: Workbook, plan: { url: string; room?: string; token?: string; user?: CollabUser }, bytes: Uint8Array | undefined, readOnly: boolean, stamp?: CollabStamp): void => {
    if (!bytes) {
      send({ type: 'collabState', state: 'closed', reason: 'unsupported', rev: 0, pending: 0, peers: [] });
      return;
    }
    let session: CollabSession | undefined;
    const tell = (): void => {
      if (!session || shown !== current) return;
      send({
        type: 'collabState',
        state: session.state,
        // Why a closed session ended by itself ('version': the room is edited with another version of the package).
        ...(session.reason === undefined ? {} : { reason: session.reason }),
        rev: session.rev,
        pending: session.pending,
        // Names are whatever the others said: text, to be shown as text.
        peers: session.peers.map((peer) => ({ id: peer.id, name: peer.name === undefined ? null : peer.name, color: peer.color, readOnly: peer.readOnly })),
        // Everybody in the file, this participant first: as the page lists them (a name for everyone, "Guest N" for who gave none).
        people: listPeople(session.user, session.peers, session.state, view.collabStrings ? view.collabStrings.guest : (n: number) => `Guest ${n}`).list.map((who) => (who.sheet === undefined ? { id: who.id, name: who.name, color: who.color, self: who.self, readOnly: who.readOnly } : { id: who.id, name: who.name, color: who.color, self: who.self, readOnly: who.readOnly, sheet: who.sheet })),
      });
    };
    let transport: CollabTransport;
    try {
      transport = host.createTransport ? host.createTransport(plan.url) : socketTransport(plan.url);
    } catch (error) {
      send({ type: 'collabState', state: 'closed', reason: 'unsupported', rev: 0, pending: 0, peers: [] });
      return;
    }
    session = new CollabSession(
      {
        workbook: () => book,
        changed: (change) => {
          if (shown === current) view.collabChanged(change, change.local !== true);
        },
        notice: (said) => {
          if (shown === current) view.collabNotice(said);
        },
        // With a save of the page: what the application has not stored. Without one the page cannot know, and nothing is taken to be saved.
        unsaved: current.saver ? () => (current.saver as SaveController).dirty : undefined,
      },
      {
        transport,
        versions: [contentVersion(bytes)],
        // What the file says of the room and the step it was saved at: it joins there, whatever anybody says of it.
        stamp,
        user: plan.user,
        room: plan.room,
        token: plan.token,
        readOnly,
        onState: tell,
        // What someone else says was saved is passed on as it was said: whether it is so is the application's to find out.
        onSaved: (saved) => {
          if (shown !== current) return;
          let etag: string | null = null;
          for (const version of saved.versions) if (version.slice(0, 2) === 'e:') etag = version.slice(2);
          send({ type: 'collabSaved', etag, rev: saved.rev === undefined ? null : saved.rev, versions: saved.versions });
        },
        onReadOnly: () => {
          if (shown === current) view.setReadOnly(true);
        },
        onPeers: (peers) => {
          if (shown !== current) return;
          view.collabPeersChanged(peers);
          tell();
        },
      },
    );
    current.collab = session;
    view.attachCollab(session);
    tell();
  };

  /**
   * The same for a document (a docx): the session of a document, and the view of one. `bytes`
   * is undefined for a document that cannot be edited together (not a docx, converted here,
   * or one whose changes are tracked).
   */
  const joinDocCollab = (current: Shown, pages: BridgeDocumentView, doc: DocModel, plan: { url: string; room?: string; token?: string; user?: CollabUser }, bytes: Uint8Array | undefined, readOnly: boolean, stamp?: CollabStamp): void => {
    const none = (): void => send({ type: 'collabState', state: 'closed', reason: 'unsupported', rev: 0, pending: 0, peers: [] });
    if (!bytes || !pages.attachCollab || !pages.collabChanged) return none();
    let session: DocCollabSession | undefined;
    const bare = (peer: DocPeer): { id: string; name?: string; color: string; guest: number; readOnly: boolean } => (peer.name === undefined ? { id: peer.id, color: peer.color, guest: peer.guest, readOnly: peer.readOnly } : { id: peer.id, name: peer.name, color: peer.color, guest: peer.guest, readOnly: peer.readOnly });
    const tell = (): void => {
      if (!session || shown !== current) return;
      send({
        type: 'collabState',
        state: session.state,
        // Why a closed session ended by itself ('version': the room is edited with another version of the package).
        ...(session.reason === undefined ? {} : { reason: session.reason }),
        rev: session.rev,
        pending: session.pending,
        // Names are whatever the others said: text, to be shown as text.
        peers: session.peers.map((peer) => ({ id: peer.id, name: peer.name === undefined ? null : peer.name, color: peer.color, readOnly: peer.readOnly })),
        people: listPeople(bare(session.user), session.peers.map(bare), session.state, pages.collabStrings ? pages.collabStrings.guest : (n: number) => `Guest ${n}`).list.map((who) => ({ id: who.id, name: who.name, color: who.color, self: who.self, readOnly: who.readOnly })),
      });
    };
    let transport: CollabTransport;
    try {
      transport = host.createTransport ? host.createTransport(plan.url) : socketTransport(plan.url);
    } catch (error) {
      return none();
    }
    session = new DocCollabSession(
      {
        document: () => doc,
        changed: (change) => {
          if (shown !== current) return;
          if (pages.collabChanged) pages.collabChanged(change);
          // What someone else changed: the application hears that the document changed, and that there is something to save.
          if (change.local) return;
          if (current.saver) current.saver.changed(true);
          tellComments(current);
          send({ type: 'changed', sheets: [], canUndo: pages.canUndo === true, canRedo: pages.canRedo === true });
        },
        notice: (said) => {
          if (shown === current && pages.collabNotice) pages.collabNotice(said);
        },
        // With a save of the page: what the application has not stored. Without one the page cannot know, and nothing is taken to be saved.
        unsaved: current.saver ? () => (current.saver as SaveController).dirty : undefined,
      },
      {
        transport,
        versions: [contentVersion(bytes)],
        // What the file says of the room and the step it was saved at: it joins there, whatever anybody says of it.
        stamp,
        user: plan.user,
        room: plan.room,
        token: plan.token,
        readOnly,
        onState: tell,
        // What someone else says was saved is passed on as it was said: whether it is so is the application's to find out.
        onSaved: (saved) => {
          if (shown !== current) return;
          let etag: string | null = null;
          for (const version of saved.versions) if (version.slice(0, 2) === 'e:') etag = version.slice(2);
          send({ type: 'collabSaved', etag, rev: saved.rev === undefined ? null : saved.rev, versions: saved.versions });
        },
        onReadOnly: () => {
          if (shown === current && pages.setReadOnly) pages.setReadOnly(true);
        },
        onPeers: (peers) => {
          if (shown !== current) return;
          if (pages.collabPeersChanged) pages.collabPeersChanged(peers);
          tell();
        },
      },
    );
    current.collab = session;
    pages.attachCollab(session);
    tell();
  };

  /** Tells the application the comments of the document, when they are not what it was told last. */
  const tellComments = (of: Shown): void => {
    const pages = of.pages;
    if (!pages) return;
    // A view that says when its comments changed is not asked for the list after every key that is typed.
    const revision = pages.commentsRevision;
    if (typeof revision === 'number') {
      if (revision === of.commentsRevision) return;
      of.commentsRevision = revision;
    }
    const state = commentState(pages);
    if (!state) return;
    const key = listDigest(state.list);
    if (key === of.commentsSent) return;
    of.commentsSent = key;
    const out: Message = { type: 'comments', comments: state.list };
    if (state.more) out.more = true;
    send(out);
  };

  /**
   * Adds, edits, answers, resolves or removes a comment of the document that is shown, as one
   * undo step. The list that results is told by `comments` (before the `changed` of the same
   * change), whether or not the command has an id; the answer says that it was done (with the
   * id of the comment that was made) or why it was not.
   */
  const commentDocument = (type: 'addComment' | 'editComment' | 'replyComment' | 'resolveComment' | 'removeComment', message: Message, id: string | undefined) => {
    // Everything that comes from outside is looked at before anything is asked of the view.
    const text = type === 'addComment' || type === 'editComment' || type === 'replyComment' ? field<string>(message, 'text', 'string', true) : undefined;
    if (text !== undefined && text.length > MAX_COMMENT_CHARS) throw new BadMessage('"text" is too long');
    const commentId = type === 'addComment' ? undefined : field<number>(message, 'commentId', 'number', true);
    if (commentId !== undefined && !(Number.isInteger(commentId) && commentId >= 0 && commentId <= 0x7fffffff)) throw new BadMessage('"commentId" must be the id of a comment');
    const resolved = type === 'resolveComment' ? field<boolean>(message, 'done', 'boolean', false) : undefined;
    const anchor = type === 'addComment' ? positionField(message, 'anchor') : undefined;
    const focus = type === 'addComment' ? positionField(message, 'focus') : undefined;
    if (focus && !anchor) throw new BadMessage('"focus" needs "anchor"');
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    if (!pages) return fail('no-document', 'The open file is not a document', id);
    if (shown.locked || shown.readOnly === true || pages.canComment !== true) return fail('read-only', 'The document takes no comments: it is read-only', id);
    let made: number | undefined;
    let did = false;
    try {
      if (type === 'addComment') {
        made = pages.addComment ? pages.addComment(text as string, anchor ? { anchor, focus: focus || anchor } : undefined) : undefined;
        did = made !== undefined;
      } else if (type === 'replyComment') {
        made = pages.replyToComment ? pages.replyToComment(commentId as number, text as string) : undefined;
        did = made !== undefined;
      } else if (type === 'editComment') did = pages.editComment ? pages.editComment(commentId as number, text as string) === true : false;
      else if (type === 'resolveComment') did = pages.resolveComment ? pages.resolveComment(commentId as number, resolved !== false) === true : false;
      else did = pages.removeComment ? pages.removeComment(commentId as number) === true : false;
    } catch (error) {
      return fail('comment-refused', errorText(error), id);
    }
    // The view said why (onNotice of the page); the application hears that nothing was changed.
    if (!did) return fail('comment-refused', 'The comment was not changed: there is no such comment, no place for it, or nothing to change', id);
    if (id === undefined) return;
    const out: Message = { type: 'done', id, command: type };
    if (made !== undefined) out.commentId = made;
    send(out);
  };

  /**
   * Shows or hides the panel that lists the comments of the document that is shown (`open`;
   * without it the panel is shown when it is hidden and hidden when it is shown). Only looks:
   * a read-only document has the panel too. The answer says whether it is shown afterwards.
   */
  const commentsPanelCommand = (message: Message, id: string | undefined) => {
    const open = field<boolean>(message, 'open', 'boolean', false);
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    if (!pages) return fail('no-document', 'The open file is not a document', id);
    const now = pages.setCommentsPanel ? pages.setCommentsPanel(open === undefined ? pages.commentsPanelOpen !== true : open) === true : false;
    if (id === undefined) return;
    send({ type: 'done', id, command: 'commentsPanel', open: now });
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
    // A protection that allows comments and nothing else: the text, and so its tracked changes, cannot be touched.
    if (shown.commentsOnly) return fail('read-only', 'The protection of the document allows comments only', id);
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
   * Sets or clears tab stops of the paragraphs of the selection of the document that is shown, or
   * the distance between its automatic tab stops, as one undo step. The answer says whether
   * anything changed and what the stops are afterwards; what the view refuses (a change that
   * cannot be tracked, a document that is edited together) it says to the user as well.
   */
  const tabDocument = (type: 'setTabStops' | 'setDefaultTabStop', message: Message, id: string | undefined) => {
    let patch: unknown;
    let twips: number | undefined;
    if (type === 'setTabStops') {
      const limited = (name: string): unknown[] | undefined => {
        const value = message[name];
        if (value === undefined || value === null) return undefined;
        if (!Array.isArray(value) || value.length > 64) throw new BadMessage(`"${name}" must be a list of at most 64 entries`);
        return value;
      };
      const set = limited('set');
      const clear = limited('clear');
      const clearAll = field<boolean>(message, 'clearAll', 'boolean', false);
      // Nothing to set and nothing to clear: the paragraphs' own stops go.
      patch = set === undefined && clear === undefined && clearAll === undefined ? null : { set, clear, clearAll };
      if (patch) for (const key of ['set', 'clear', 'clearAll']) if ((patch as Record<string, unknown>)[key] === undefined) delete (patch as Record<string, unknown>)[key];
    } else twips = field<number>(message, 'twips', 'number', true);
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    if (!pages) return fail('no-document', 'The open file is not a document', id);
    if (shown.locked || shown.readOnly === true || shown.commentsOnly) return fail('read-only', 'The document is read-only', id);
    let changed = false;
    if (type === 'setTabStops') changed = pages.setTabStops ? pages.setTabStops(patch) === true : false;
    else changed = pages.setDefaultTabStop ? pages.setDefaultTabStop(twips as number) === true : false;
    if (id === undefined) return;
    const out: Message = { type: 'done', id, command: type, changed };
    if (pages.tabStops) out.tabStops = pages.tabStops.slice(0, 64);
    if (typeof pages.defaultTabStop === 'number') out.defaultTabStop = pages.defaultTabStop;
    send(out);
  };

  /**
   * The page setup through the view, as one undo step: of a sheet (the active one unless `index`
   * names one) or of the sections of a document. The answer (with an id) says whether anything
   * changed and how the page is set up now.
   */
  const pageSetupCommand = (type: 'setPageSetup' | 'insertPageBreak' | 'removePageBreak' | 'resetPageBreaks' | 'insertSectionBreak', message: Message, id: string | undefined) => {
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    const pages = shown.pages;
    const answer: Message = { type: 'done', id, command: type, changed: false };
    if (view) {
      if (type === 'insertSectionBreak') return fail('no-document', 'The open file is not a document', id);
      // Everything that comes from outside is looked at before anything is asked of the view.
      const index = sheetIndexField(message, 'index', false);
      const patch = type === 'setPageSetup' ? sheetPageSetupPatch(message) : undefined;
      const row = type === 'insertPageBreak' || type === 'removePageBreak' ? sheetIndexField(message, 'row', false) : undefined;
      const col = type === 'insertPageBreak' || type === 'removePageBreak' ? sheetIndexField(message, 'col', false) : undefined;
      if (shown.locked || view.readOnly === true) return fail('read-only', 'The file is read-only', id);
      const count = sheetCount ? sheetCount() : 0;
      if (index !== undefined && index >= count) return fail('no-sheet', `No sheet with index ${index}`, id);
      const sheet = index !== undefined ? index : view.activeSheet;
      // A break of another sheet than the shown one has its place: the active cell is of the shown sheet.
      if ((type === 'insertPageBreak' || type === 'removePageBreak') && sheet !== view.activeSheet && row === undefined && col === undefined) throw new BadMessage('"row" or "col" is needed for a sheet that is not shown');
      try {
        if (!commitPendingOf(view)) return fail('page-setup-refused', STILL_EDITING, id);
        const at = row !== undefined || col !== undefined ? { row, col } : undefined;
        if (type === 'setPageSetup') answer.changed = view.setPageSetup(patch as PageSetupPatch, sheet);
        else if (type === 'insertPageBreak') answer.changed = view.insertPageBreak(at, sheet);
        else if (type === 'removePageBreak') answer.changed = view.removePageBreak(at, sheet);
        else answer.changed = view.resetPageBreaks(sheet);
        answer.pageSetup = view.pageSetup(sheet);
      } catch (error) {
        return fail(refusalCode(error, 'page-setup-refused'), refusalText(view, error), id);
      }
    } else if (pages) {
      if (type === 'insertPageBreak' || type === 'removePageBreak' || type === 'resetPageBreaks') return fail('no-sheet', 'The open file is a document: it has no sheets', id);
      const patch = type === 'setPageSetup' ? sectionPatch(message) : undefined;
      const scope = type === 'setPageSetup' ? field<string>(message, 'scope', 'string', false) : undefined;
      if (scope !== undefined && scope !== 'section' && scope !== 'document') throw new BadMessage('"scope" must be "section" or "document"');
      const kind = type === 'insertSectionBreak' ? field<string>(message, 'kind', 'string', false) : undefined;
      if (kind !== undefined && kind !== 'nextPage' && kind !== 'continuous' && kind !== 'evenPage' && kind !== 'oddPage') throw new BadMessage('"kind" must be "nextPage", "continuous", "evenPage" or "oddPage"');
      if (shown.locked || shown.readOnly === true || shown.commentsOnly) return fail('read-only', 'The document is read-only', id);
      if (typeof pages.setSectionSetup !== 'function' || typeof pages.insertSectionBreak !== 'function') return fail('page-setup-refused', 'This view cannot change the page setup', id);
      // The view says itself why it did not (tracked changes, a document edited together): nothing changed then.
      if (type === 'setPageSetup') answer.changed = pages.setSectionSetup(patch, scope === 'document' ? 'document' : 'section') === true;
      else answer.changed = pages.insertSectionBreak(kind || 'nextPage') === true;
      if (pages.sectionSetup) answer.pageSetup = pages.sectionSetup;
    } else return fail('not-loaded', 'No file is open', id);
    if (id !== undefined) send(answer);
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
      return fail(refusalCode(error, 'sheet-refused'), text, id);
    }
    if (id !== undefined) send(answer);
  };

  /**
   * Refreshes every pivot table of the workbook that can be, as one undo step: the answer is "done" with how many
   * were refreshed and, for each that was not, its name and why (in the language of the view).
   */
  const refreshPivots = (id: string | undefined) => {
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    if (!view) return fail('no-sheet', 'The open file is a document: it has no pivot tables', id);
    if (shown.locked || view.readOnly === true) return fail('read-only', 'The file is read-only', id);
    if (typeof view.refreshAllPivots !== 'function') return fail('pivot-refused', 'This view cannot refresh pivot tables', id);
    const answer: Message = { type: 'done', id, command: 'refreshPivots' };
    try {
      // As in changeSheets(): what is being typed is applied first, and nothing is said to be done that was not.
      if (!commitPendingOf(view)) return fail('pivot-refused', STILL_EDITING, id);
      const result = view.refreshAllPivots();
      answer.refreshed = result.refreshed.length;
      answer.refused = result.refused.map((entry) => ({ name: entry.name, reason: entry.reason }));
    } catch (error) {
      let text: string;
      try {
        text = typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
      } catch {
        text = errorText(error);
      }
      return fail(refusalCode(error, 'pivot-refused'), text, id);
    }
    if (id !== undefined) send(answer);
  };

  /**
   * Makes a pivot table (insertPivot), changes the fields of one (setPivotFields), lists them
   * (pivotFields) or deletes one (deletePivot). What the message holds is copied field by field, as
   * lists and numbers only; the commands check all of it. The answer is "done" with the table and its
   * sheet (insertPivot), whether anything changed (setPivotFields, deletePivot), or the fields
   * (pivotFields).
   */
  const pivotCommand = (type: 'insertPivot' | 'setPivotFields' | 'pivotFields' | 'deletePivot', message: Message, id: string | undefined) => {
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    if (!view) return fail('no-sheet', 'The open file is a document: it has no pivot tables', id);
    if (typeof view.insertPivot !== 'function' || typeof view.setPivotFields !== 'function' || typeof view.pivotFields !== 'function' || (type === 'deletePivot' && typeof view.deletePivot !== 'function')) {
      return fail('pivot-refused', 'This view cannot change pivot tables', id);
    }
    const index = message.index;
    const numberList = (value: unknown): number[] | undefined => (Array.isArray(value) && value.length <= 16384 ? value.map((one) => (typeof one === 'number' ? one : NaN)) : undefined);
    const answer: Message = { type: 'done', id, command: type };
    try {
      if (type === 'pivotFields') {
        if (typeof index !== 'number') return fail('pivot-refused', 'The message names no pivot table (index)', id);
        // A plain copy: what the file names goes out as text.
        answer.fields = JSON.parse(JSON.stringify(view.pivotFields(index)));
        if (id !== undefined) send(answer);
        return;
      }
      if (shown.locked || view.readOnly === true) return fail('read-only', 'The file is read-only', id);
      if (!commitPendingOf(view)) return fail('pivot-refused', STILL_EDITING, id);
      if (type === 'deletePivot') {
        if (typeof index !== 'number') return fail('pivot-refused', 'The message names no pivot table (index)', id);
        answer.changed = view.deletePivot(index);
        if (id !== undefined) send(answer);
        return;
      }
      const lists: Record<string, unknown> = {};
      for (const key of ['rows', 'cols']) if (message[key] !== undefined) lists[key] = numberList(message[key]);
      // A filter is a field, or a field with the item it shows ({ field, item }).
      if (Array.isArray(message.filters) && message.filters.length <= 16384) {
        lists.filters = message.filters.map((one: unknown) => {
          if (typeof one === 'number') return one;
          if (!one || typeof one !== 'object') return NaN;
          const entry = one as Record<string, unknown>;
          const out: Record<string, unknown> = { field: typeof entry.field === 'number' ? entry.field : NaN };
          if (entry.item !== undefined) out.item = typeof entry.item === 'number' ? entry.item : NaN;
          return out;
        });
      } else if (message.filters !== undefined) lists.filters = numberList(message.filters);
      if (Array.isArray(message.values) && message.values.length <= 256) {
        lists.values = message.values.map((one: unknown) => {
          const value = one && typeof one === 'object' ? (one as Record<string, unknown>) : {};
          const out: Record<string, unknown> = { field: typeof value.field === 'number' ? value.field : NaN };
          if (typeof value.subtotal === 'string') out.subtotal = value.subtotal;
          if (typeof value.numFmt === 'string') out.numFmt = value.numFmt;
          if (typeof value.name === 'string' && type === 'setPivotFields') out.name = value.name;
          return out;
        });
      }
      if (type === 'insertPivot') {
        const selection = view.selection;
        const sheet = view.activeSheet;
        const source =
          typeof message.source === 'string' ? readPivotSource(message.source, view.workbook, sheet) : selection ? { sheet, range: selection.range } : undefined;
        if (!source) return fail('pivot-refused', 'The source of the pivot table is not a range or a name of the workbook', id);
        const options: Record<string, unknown> = { ...lists, source };
        if (typeof message.cell === 'string') {
          const cell = readPivotCell(message.cell);
          if (!cell) return fail('pivot-refused', 'The cell of the pivot table is not one', id);
          options.target = { sheet, row: cell.row, col: cell.col };
        }
        const made = view.insertPivot(options as unknown as Parameters<WorkbookView['insertPivot']>[0], message.panel === true);
        if (made < 0) return fail('pivot-refused', NOT_DONE, id);
        answer.index = made;
        answer.sheet = view.activeSheet;
      } else {
        if (typeof index !== 'number') return fail('pivot-refused', 'The message names no pivot table (index)', id);
        const change: Record<string, unknown> = { ...lists };
        if (typeof message.rowGrandTotals === 'boolean') change.rowGrandTotals = message.rowGrandTotals;
        if (typeof message.colGrandTotals === 'boolean') change.colGrandTotals = message.colGrandTotals;
        const pairs = (value: unknown, second: string, kind: 'boolean' | 'string' | 'list'): unknown[] | undefined =>
          Array.isArray(value) && value.length <= 16384
            ? value.map((one: unknown) => {
                const entry = one && typeof one === 'object' ? (one as Record<string, unknown>) : {};
                const other = entry[second];
                return { field: typeof entry.field === 'number' ? entry.field : NaN, [second]: kind === 'list' ? numberList(other) : typeof other === kind ? other : undefined };
              })
            : undefined;
        if (message.subtotals !== undefined) change.subtotals = pairs(message.subtotals, 'on', 'boolean');
        if (message.sort !== undefined) change.sort = pairs(message.sort, 'order', 'string');
        if (message.hidden !== undefined) change.hidden = pairs(message.hidden, 'items', 'list');
        answer.changed = view.setPivotFields(index, change as Parameters<WorkbookView['setPivotFields']>[1]);
      }
    } catch (error) {
      let text: string;
      try {
        text = typeof view.localizeError === 'function' ? view.localizeError(error) : errorText(error);
      } catch {
        text = errorText(error);
      }
      return fail(refusalCode(error, 'pivot-refused'), text, id);
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
      return fail(refusalCode(error, 'sheet-refused'), text, id);
    }
    // A command that changed nothing a message tells of (a colour) still has its list looked at.
    if (id !== undefined) send(answer);
  };

  /**
   * The bytes of the file. For a workbook that is edited together what arrives is held back
   * while the file is written, so that it is the workbook of one moment, and the others are
   * told what these bytes are known by and which step they hold: whoever loads them later
   * joins the room from there.
   */
  const sharedBytes = async (current: Shown): Promise<Uint8Array> => {
    const session = current.collab;
    if (!session) return current.file.save();
    // The stamp and the save are one step under the turn of the package.
    const turn = packageTurn(current.file);
    const done = await writeSharedFile(session, current.file.package, () => turn.save(), turn.locked);
    // Told to the others only once the application says it has stored the file (see userSaver()).
    if (session.state !== 'closed') current.written = done;
    return done.bytes;
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
    const bytes = await sharedBytes(current);
    // A change made while the file was written is not in these bytes.
    if (pages && pages.markSaved && pages.changeCount === changes) pages.markSaved();
    return bytes;
  };

  /** The file that is shown as bytes, as its model is: nothing that is being typed is applied, nothing is closed. */
  const writeShownAsIs = async (current: Shown): Promise<Uint8Array> => {
    const pages = current.pages;
    const changes = pages ? pages.changeCount : undefined;
    const bytes = await sharedBytes(current);
    if (pages && pages.markSaved && pages.changeCount === changes) pages.markSaved();
    return bytes;
  };

  /**
   * What the Save button and Ctrl+S of the page run for a `load` with `save`: the application is
   * handed the file in a `saved` message that says it comes from the user, and answers with
   * `saveResult` when it has stored it. Until then the page shows the save as under way; without
   * an answer in time it has failed. With `auto` the page also hands the file over by itself, in
   * `saved` messages that say `by: "auto"` and are answered the same way.
   */
  function userSaver(auto?: AutoSaveSettings): SaveController {
    const mine: SaveController = new SaveController({
      target: {
        kind: 'function',
        // No closure is made here: what waits for the result (awaitResult) knows the id of the save and
        // never its bytes, so that a file is let go of as soon as its message is.
        save: (bytes: Uint8Array, info: SaveInfo) => {
          // Never the id of a save before it, of this file or of another: an answer that comes late names a save that is gone.
          // One count for both kinds: the letter only says who started the save.
          const saveId = `${info.auto === true ? 'a' : 'u'}${++userSaves}`;
          const out: Message = { type: 'saved', by: info.auto === true ? 'auto' : 'user', saveId, data: base64Encode(bytes), format: info.format };
          if (shown && shown.file.passwordProtected === true) out.encrypted = true;
          // The user was asked and chose to keep their own file: the application writes it whatever it holds.
          if (info.overwrite === true) out.overwrite = true;
          if (info.name !== undefined) out.name = info.name;
          if (info.convertedFrom !== undefined) out.convertedFrom = info.convertedFrom;
          // Before the message goes: an application that answers from inside the call finds the save waiting.
          const result = awaitResult(saveId);
          send(out);
          // The application stored the file: now the others are told what it is known by and which step it holds.
          return result.then(() => {
            const current = shown;
            if (!current || current.saver !== mine || !current.collab || !current.written || current.written.bytes !== bytes) return;
            current.collab.announceSaved([contentVersion(bytes)], current.written.rev);
          });
        },
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
      bytesAsIs: () => {
        const current = shown;
        if (!current || current.saver !== mine) return Promise.reject(new Error('No file is open'));
        return writeShownAsIs(current);
      },
      // What auto-save must not fall into. A view that cannot say is one in which nothing is typed.
      pending: () => {
        const current = shown;
        if (!current || current.saver !== mine) return true;
        const view = current.view || current.pages;
        return !!view && typeof view.hasPendingInput === 'function' && view.hasPendingInput();
      },
      readOnly: () => {
        const current = shown;
        return !current || (current.view ? current.view.readOnly === true : current.readOnly === true);
      },
      autoSave: auto,
      now: host.now,
      info: () => {
        const current = shown;
        const out: SaveInfo = { format: (current ? current.file.format : 'docx') as SaveInfo['format'] };
        if (current && current.name !== undefined) out.name = current.name;
        if (current && current.convertedFrom !== undefined) out.convertedFrom = current.convertedFrom as SaveInfo['convertedFrom'];
        return out;
      },
      strings: () => resolveUiStrings(shown ? shown.locale : undefined),
      // The clock of the page, as everything else here uses it.
      timer: cancellable,
      notice: (text, kind) => {
        const current = shown;
        if (!current || current.saver !== mine) return;
        const view = current.view || current.pages;
        if (view && typeof view.showNotice === 'function') view.showNotice(text, kind);
      },
      // A `saveResult` that said `conflict`: the user decides whether the file goes over what the application holds.
      askOverwrite:
        typeof host.askOverwrite === 'function'
          ? (question) => {
              const current = shown;
              if (!current || current.saver !== mine) return question.answer(false);
              // A file that is only read is not asked about: its conflict was told, and that is all.
              if (current.view ? current.view.readOnly === true : current.readOnly === true) return question.answer(false);
              const element = (current.view || current.pages || {}).element;
              return (host.askOverwrite as NonNullable<BridgeHost['askOverwrite']>)(question, current.locale, element || undefined);
            }
          : undefined,
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

  /**
   * Waits for the `saveResult` of the save that was just handed over: resolves and rejects as the
   * result says, and rejects when none comes in time. Whoever ends the wait calls its timer off.
   */
  function awaitResult(saveId: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const waiting = { saveId, resolve, reject, stop: (): void => undefined };
      awaited = waiting;
      waiting.stop = cancellable(() => {
        if (awaited !== waiting) return;
        awaited = undefined;
        reject(new SaveError('timeout', `The application did not answer the save within ${Math.round(saveResultTimeout / 1000)} s`));
      }, saveResultTimeout);
    });
  }

  /** The application's answer to a `saved` message that came from the user. */
  const saveResult = (message: Message, id: string | undefined) => {
    const saveId = field<string>(message, 'saveId', 'string', true);
    const ok = field<boolean>(message, 'ok', 'boolean', true);
    const text = field<string>(message, 'message', 'string', false);
    const conflict = field<boolean>(message, 'conflict', 'boolean', false);
    const waiting = awaited;
    // Only the save that waits now is answered, and only by its own id: a result that comes after the page
    // stopped waiting, twice, or for a file that was closed must not be taken for the result of the next save.
    if (!waiting || waiting.saveId !== saveId) return fail('bad-message', 'No save with this "saveId" is waiting for a result', id);
    awaited = undefined;
    waiting.stop();
    const said = text === undefined ? '' : text.slice(0, MAX_RESULT_MESSAGE_CHARS);
    if (ok === true) waiting.resolve();
    // The file the application holds is not the one it gave the page: the user is asked, and what the application said is shown with the question.
    else if (conflict === true) waiting.reject(new SaveError('conflict', 'The application holds another file than the one that was opened', undefined, said));
    else waiting.reject(new Error(said));
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
    if (current.file.passwordProtected === true) out.encrypted = true;
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

  /**
   * Opens (or with `open: false` closes) the print preview of the workbook that is shown: a layer
   * over it with the pages print() would print. A document has none: its view shows its pages.
   */
  const printPreviewCommand = (message: Message, id: string | undefined) => {
    const open = field<boolean>(message, 'open', 'boolean', false);
    const sheets = field<string>(message, 'sheets', 'string', false);
    if (sheets !== undefined && sheets !== 'active' && sheets !== 'all') throw new BadMessage('"sheets" must be "active" or "all"');
    const page = field<number>(message, 'page', 'number', false);
    if (page !== undefined && !(Number.isInteger(page) && page >= 1)) throw new BadMessage('"page" must be a whole number of 1 or more');
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const view = shown.view;
    if (shown.pages || !view || typeof (view as WorkbookView).printPreview !== 'function') return fail('no-sheet', 'The open file is a document: its pages are shown already', id);
    const grid = view as WorkbookView;
    let now: boolean;
    if (open === false) {
      grid.closePrintPreview();
      now = false;
    } else {
      const options: { sheets?: 'active' | 'all'; page?: number } = {};
      if (sheets !== undefined) options.sheets = sheets as 'active' | 'all';
      if (page !== undefined) options.page = page;
      now = grid.printPreview(options);
    }
    if (id === undefined) return;
    send({ type: 'done', id, command: 'printPreview', open: now });
  };

  /**
   * Makes a PDF of the document or the workbook and answers with it: `pdf`,
   * with the file as base64, the number of its pages and how many of them
   * were written as text and drawings (`vector`; the others are pictures).
   * As with printing, the messages after it do not wait: the pages of a long
   * file take their time.
   */
  const exportPdf = (message: Message, id: string | undefined) => {
    const dpi = field<number>(message, 'dpi', 'number', false);
    if (dpi !== undefined && !(dpi > 0 && dpi < Infinity)) throw new BadMessage('"dpi" must be a positive number');
    const image = field<string>(message, 'image', 'string', false);
    if (image !== undefined && image !== 'auto' && image !== 'jpeg' && image !== 'lossless') throw new BadMessage('"image" must be "auto", "jpeg" or "lossless"');
    const quality = field<number>(message, 'quality', 'number', false);
    if (quality !== undefined && !(quality > 0 && quality <= 1)) throw new BadMessage('"quality" must be a number above 0 and at most 1');
    const sheets = field<string>(message, 'sheets', 'string', false);
    if (sheets !== undefined && sheets !== 'active' && sheets !== 'all') throw new BadMessage('"sheets" must be "active" or "all"');
    const text = field<string>(message, 'text', 'string', false);
    if (text !== undefined && text !== 'auto' && text !== 'image' && text !== 'vector') throw new BadMessage('"text" must be "auto", "image" or "vector"');
    const title = field<string>(message, 'title', 'string', false);
    const author = field<string>(message, 'author', 'string', false);
    const links = field<boolean>(message, 'links', 'boolean', false);
    const outline = field<boolean>(message, 'outline', 'boolean', false);
    if (!shown) return fail('not-loaded', 'No file is open', id);
    const pages = shown.pages;
    const view = shown.view;
    if (pages ? !pages.exportPdf && !pages.exportPdfParts : !view || typeof view.exportPdf !== 'function') return fail('pdf-unavailable', 'This view cannot make a PDF', id);
    let count = 0;
    let vector = 0;
    let linkReport: PdfLinkReport | undefined;
    const options: SheetPdfExportOptions = {
      onLinks: (report) => {
        linkReport = report;
      },
      onProgress: (_done, total) => {
        count = total;
      },
      onPage: (info) => {
        // Counted anew when the pages are painted again from the start.
        if (info.page === 1) vector = 0;
        if (info.mode === 'vector') vector++;
      },
    };
    if (text !== undefined) options.text = text as 'auto' | 'image' | 'vector';
    if (dpi !== undefined) options.dpi = dpi;
    if (image !== undefined) options.image = image as 'auto' | 'jpeg' | 'lossless';
    if (quality !== undefined) options.quality = quality;
    if (title !== undefined) options.title = title;
    if (author !== undefined) options.author = author;
    if (links !== undefined) options.links = links;
    if (outline !== undefined) options.outline = outline;
    const name = pdfFileName(shown.name, pages ? 'document' : 'workbook');
    // The file in the parts it was written in, where the view hands them out: it is never held as one array here.
    let job: Promise<Uint8Array[]>;
    try {
      if (pages) job = pages.exportPdfParts ? Promise.resolve(pages.exportPdfParts(options)) : Promise.resolve((pages.exportPdf as NonNullable<BridgeDocumentView['exportPdf']>)(options)).then((bytes) => [bytes]);
      else {
        // Of a workbook: the sheet that is shown unless all are asked for; the name of the file is what "&F" in a header stands for.
        if (sheets !== undefined) options.sheets = sheets as 'active' | 'all';
        if (shown.name !== undefined) options.fileName = shown.name;
        const grid = view as WorkbookView;
        job = typeof grid.exportPdfParts === 'function' ? Promise.resolve(grid.exportPdfParts(options)) : Promise.resolve(grid.exportPdf(options)).then((bytes) => [bytes]);
      }
    } catch (error) {
      job = Promise.reject(error);
    }
    const failed = (error: unknown) => {
      const code = error ? (error as { code?: unknown }).code : undefined;
      fail(code === 'unavailable' ? 'pdf-unavailable' : code === 'unsupported' ? 'pdf-unsupported' : 'pdf-failed', errorText(error), id);
    };
    job.then((parts) => {
      // Whatever goes wrong from here on is answered too: the application waits for one of the two.
      try {
        if (!Array.isArray(parts)) throw new TypeError('The view gave no file');
        let size = 0;
        for (let i = 0; i < parts.length; i++) size += parts[i] ? parts[i].length : 0;
        if (size > MAX_BRIDGE_PDF_BYTES) {
          return fail('pdf-too-large', `The PDF has ${size} bytes; a message carries ${MAX_BRIDGE_PDF_BYTES} at most. Ask for a lower "dpi", or for fewer sheets`, id);
        }
        const out: Message = { type: 'pdf' };
        if (id !== undefined) out.id = id;
        out.data = base64OfParts(parts);
        out.name = name;
        out.pages = count;
        out.vector = vector;
        // What became of the hyperlinks and the outline: written, left out and why.
        if (linkReport) out.links = { ...linkReport };
        send(out);
      } catch (error) {
        failed(error);
      }
    }, failed);
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
      if (type === 'exportPdf') return exportPdf(fields, id);
      if (type === 'printPreview') return printPreviewCommand(fields, id);
      if (type === 'insertSheet' || type === 'renameSheet' || type === 'deleteSheet') return changeSheets(type, fields, id);
      if (type === 'moveSheet' || type === 'setSheetState' || type === 'copySheet' || type === 'setTabColor') return arrangeSheets(type, fields, id);
      if (type === 'acceptAllRevisions' || type === 'rejectAllRevisions' || type === 'setTrackChanges') return reviseDocument(type, fields, id);
      if (type === 'refreshPivots') return refreshPivots(id);
      if (type === 'insertPivot' || type === 'setPivotFields' || type === 'pivotFields' || type === 'deletePivot') return pivotCommand(type, fields, id);
      if (type === 'setTabStops' || type === 'setDefaultTabStop') return tabDocument(type, fields, id);
      if (type === 'setPageSetup' || type === 'insertPageBreak' || type === 'removePageBreak' || type === 'resetPageBreaks' || type === 'insertSectionBreak') return pageSetupCommand(type, fields, id);
      if (type === 'addComment' || type === 'editComment' || type === 'replyComment' || type === 'resolveComment' || type === 'removeComment') return commentDocument(type, fields, id);
      if (type === 'commentsPanel') return commentsPanelCommand(fields, id);
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
        // A file that is only read now is not asked about: the question about a conflict is taken away, unanswered.
        if (value && shown.saver) shown.saver.dismissConflict();
        // Editing is on again: what auto-save held back while the file was only read is its to hand over.
        else if (shown.saver) shown.saver.wake();
      } else if (type === 'setLocale') {
        try {
          if (view) view.setLocale(locale as string);
          else if (pages && pages.setLocale) pages.setLocale(locale as string);
        } catch (error) {
          // The view keeps the language it has (a dialog of rules is open with changes, say), and says why in that
          // language: the page around it keeps its own as well.
          if (error instanceof RangeError) return fail('locale-refused', errorText(error), id);
          throw error;
        }
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
   * Nor does `password`: the load it answers is the one that is under way.
   */
  const handleNow = (fields: Message, type: string) => {
    let id: string | undefined;
    try {
      id = idOf(fields);
      if (type === 'hello') return send({ type: 'ready', version: host.version });
      if (type === 'saveResult') return saveResult(fields, id);
      if (type === 'password') return passwordGiven(fields, id);
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
    pageHidden() {
      if (!destroyed && shown && shown.saver) shown.saver.pageHidden();
    },
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
      if (fields && (type === 'hello' || type === 'close' || type === 'saveResult' || type === 'password')) {
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
    // Asked inside the page as well as of the application: an application that does not know the message still gets its file opened.
    passwordPrompt: (container, locale, onCancel) => new PasswordPrompt({ container, strings: resolveUiStrings(locale), onCancel }),
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
    // The question about a conflict is asked over the editor, in the language of the file.
    askOverwrite: (question, locale, element) =>
      askOverwrite({ root: editorRootIn(options.container, element), strings: resolveUiStrings(locale), detail: question.detail, answer: question.answer }),
    frame: typeof win.requestAnimationFrame === 'function' ? (callback) => void win.requestAnimationFrame(callback) : undefined,
    timer: (callback, ms) => {
      const handle = win.setTimeout(callback, ms);
      return () => win.clearTimeout(handle);
    },
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

  // A page that is hidden may never be shown again: what auto-save holds is handed over at once.
  const doc = win.document as Document | undefined;
  const onVisibility = () => {
    if (doc && doc.visibilityState === 'hidden') bridge.pageHidden();
  };
  if (doc && typeof doc.addEventListener === 'function') doc.addEventListener('visibilitychange', onVisibility);

  bridge.start();
  return {
    start: bridge.start,
    receive: bridge.receive,
    pageHidden: bridge.pageHidden,
    destroy() {
      bridge.destroy();
      if (doc && typeof doc.removeEventListener === 'function') doc.removeEventListener('visibilitychange', onVisibility);
      win.removeEventListener('message', onWindowMessage);
      if (webview && typeof webview.removeEventListener === 'function') webview.removeEventListener('message', onWebViewMessage);
      if (channels.BymOfficeBridge === entry) delete channels.BymOfficeBridge;
    },
  };
}
