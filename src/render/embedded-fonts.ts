// The font the package brings: TeX Gyre Bonum, which has the advance widths
// of Bookman Old Style, a family Apple and Android devices do not have.
//
// The four font files are not inside the scripts. They lie next to them
// (dist/fonts/tex-gyre-bonum/) and are fetched the first time a document or a
// spreadsheet draws text in a family they stand in for, on a device that
// cannot show that family, where the application has not seen to it with a
// font of its own. No such text: no request. The page of the WebView bridge
// carries the files inside itself and hands their bytes over through
// `source`.
//
// A script of a page knows its own address and finds the folder by it. An ES
// module does not (import.meta is not part of what the bundle may contain:
// it has to parse as ES2019): the application tells it, with
// configureEmbeddedFonts({ baseUrl }) or the `fontsUrl` of mount().
//
// The font is used with all four of its faces or not at all, and it is in a
// family list only while it is on the page: a document and a spreadsheet go
// by the same rule.
//
// The faces are added to the page outside the registry of the application
// (registeredFonts() lists what the application registered, nothing else),
// and the family comes after everything the application registered for the
// same name: the embedded font is the last resort before a generic family.
//
// Font names of files only ever get here sanitized, and none of them reaches
// an address: the addresses are made of what the application set and of the
// fixed file names below.

import { appFontFor, canLoadFonts, fontAliases, fontsVersion, loadPageFont, notifyFontsChanged, splitFontFamily, type FontFileSource, type FontSource, type PageFont } from './fonts';

/** The family name the embedded font is added to the page with. */
export const EMBEDDED_FONT_FAMILY = 'TeX Gyre Bonum';

/** The folder of the files inside the folder of fonts. */
const FOLDER = 'tex-gyre-bonum/';

/** The files, as GUST e-foundry distributes them: they are shipped unchanged, under their own names. */
export const EMBEDDED_FONT_FACES: readonly { file: string; weight: 400 | 700; style: 'normal' | 'italic' }[] = [
  { file: 'texgyrebonum-regular.otf', weight: 400, style: 'normal' },
  { file: 'texgyrebonum-bold.otf', weight: 700, style: 'normal' },
  { file: 'texgyrebonum-italic.otf', weight: 400, style: 'italic' },
  { file: 'texgyrebonum-bolditalic.otf', weight: 700, style: 'italic' },
];

const DEFAULT_AS = ['Bookman Old Style'];

/** How the embedded font is found and what it is used for; see configureEmbeddedFonts(). */
export interface EmbeddedFontsOptions {
  /**
   * The address of the folder of fonts of the package (its `dist/fonts/`,
   * which holds `tex-gyre-bonum/`), read as an address of the page. `null`
   * turns the embedded font off: nothing is ever fetched, and a font that
   * was loaded is taken off the page. `undefined` goes back to the default,
   * the folder `fonts/` next to the script (a `<script>` of a page; an ES
   * module has no default and must be given the address).
   */
  baseUrl?: string | null;
  /**
   * The font names of files the embedded font stands in for where the device
   * lacks them. Default: `['Bookman Old Style']`. A name with a code page
   * suffix ("Bookman Old Style Tur") is covered by its family.
   */
  as?: string | string[];
  /**
   * Gives a font file itself, by its name ("texgyrebonum-regular.otf"): its
   * bytes, or an address. A file it returns nothing for is fetched from
   * `baseUrl`. For bundlers and for pages that carry the files inside them.
   */
  source?: ((file: string) => FontSource | null | undefined | Promise<FontSource | null | undefined>) | null;
  /** Tells whether the device can show a family; replaces the built-in test where no measurer of a document asks (tests). */
  hasFamily?: ((family: string) => boolean) | null;
}

/** What embeddedFontsState() tells. */
export interface EmbeddedFontsState {
  /** "TeX Gyre Bonum". */
  family: string;
  /** The font names it stands in for. */
  as: string[];
  /** The folder the files are fetched from; null when it is turned off or cannot be worked out. */
  baseUrl: string | null;
  /**
   * "off": turned off. "idle": nothing has asked for it yet. "loading",
   * "loaded" (the four faces are on the page), "failed" (one of them could
   * not be loaded, and none is used; it is tried again when configure()
   * changes something).
   */
  status: 'off' | 'idle' | 'loading' | 'loaded' | 'failed';
}

/** What a view is told when the embedded font was needed and could not be loaded. */
export interface EmbeddedFontNotice {
  /** The font name of the file that asked for it, without its code page suffix. */
  family: string;
  /** Why, in English. */
  reason: string;
}

// ------------------------------------------------------------------ address

/** The address of the script element that is running this code, when it has one. */
function scriptAddress(): string | undefined {
  try {
    // Set while a classic script runs, also one with `async` or `defer` and one added by a script; null in a module
    // and afterwards, which is why it is read once, when this code loads. An inline script has no address.
    const script = typeof document !== 'undefined' ? (document as { currentScript?: { src?: unknown } | null }).currentScript : undefined;
    const src = script ? script.src : undefined;
    return typeof src === 'string' && src !== '' ? src : undefined;
  } catch {
    return undefined;
  }
}

function pageAddress(): string | undefined {
  try {
    const base = typeof document !== 'undefined' ? (document as { baseURI?: unknown }).baseURI : undefined;
    if (typeof base === 'string' && base !== '') return base;
    const href = typeof location !== 'undefined' ? (location as { href?: unknown }).href : undefined;
    return typeof href === 'string' && href !== '' ? href : undefined;
  } catch {
    return undefined;
  }
}

const withSlash = (address: string): string => (address.charAt(address.length - 1) === '/' ? address : address + '/');

/** The scheme of an address, lower case, with its colon; undefined when it has none. */
function schemeOf(address: string | undefined): string | undefined {
  const match = address === undefined ? null : /^([A-Za-z][A-Za-z0-9+.\-]*:)/.exec(address);
  return match ? match[1].toLowerCase() : undefined;
}

/**
 * The folder of fonts, with a slash at its end, or null when there is none:
 * what the application set (`null` is "off"), read as an address of the
 * page; else `fonts/` next to the script. A script without an address of its
 * own (inline, from a blob) has no such folder, and neither has a bundle
 * that names the file of the build machine it was made from.
 */
export function resolveFontsBase(configured: string | null | undefined, script: string | undefined, page: string | undefined): string | null {
  if (configured === null) return null;
  if (typeof configured === 'string') {
    try {
      return withSlash(page === undefined ? new URL(configured).href : new URL(configured, page).href);
    } catch {
      // Not an address on its own and no page to read it against: used as it is.
      return withSlash(configured);
    }
  }
  const scheme = schemeOf(script);
  if (script === undefined || scheme === undefined) return null;
  if (scheme === 'blob:' || scheme === 'data:' || scheme === 'about:' || scheme === 'javascript:') return null;
  // A file of the device is only the script's own address in a page that is a file as well.
  if (scheme === 'file:' && schemeOf(page) !== 'file:') return null;
  try {
    return new URL('fonts/', script).href;
  } catch {
    return null;
  }
}

// -------------------------------------------------------------------- state

/** Read once, while the script runs: document.currentScript is gone afterwards. */
let script = scriptAddress();
/** The page address tests give; the page's own otherwise. */
let pageOverride: { page: string | undefined } | undefined;

let configured: string | null | undefined;
let aliases: string[] = DEFAULT_AS.slice();
let aliasesLower: string[] = aliases.map((name) => name.toLowerCase());
let source: EmbeddedFontsOptions['source'];
let hasFamily: EmbeddedFontsOptions['hasFamily'];

/**
 * "loading" from the moment the files are asked for until all of them have
 * ended, whatever is configured meanwhile: there is one load at a time.
 */
let status: 'idle' | 'loading' | 'loaded' | 'failed' = 'idle';
/** Counts the configurations: a load that fails under another one than it began with is no failure of the new one. */
let generation = 0;
/** The four faces once their files have been read. Kept when the font is turned off, to be put back without a request. */
let faces: PageFont[] | undefined;
/** Whether the faces are on the page. */
let onPage = false;
/** Names found to need no embedded font, with the fontsVersion() that was found under. */
const notNeeded = new Map<string, number>();
const NOT_NEEDED_LIMIT = 256;

let notice: EmbeddedFontNotice | undefined;
let told = false;
const noticeListeners: ((notice: EmbeddedFontNotice) => void)[] = [];

const page = (): string | undefined => (pageOverride ? pageOverride.page : pageAddress());
const base = (): string | null => resolveFontsBase(configured, script, page());

/** Whether the files can be asked for at all: a browser, not turned off, and a place to get them from. */
function reachable(): boolean {
  return configured !== null && canLoadFonts() && (typeof source === 'function' || base() !== null);
}

export function embeddedFontsState(): EmbeddedFontsState {
  return { family: EMBEDDED_FONT_FAMILY, as: aliases.slice(), baseUrl: base(), status: configured === null ? 'off' : status };
}

/**
 * The font names a document can be offered because the embedded font shows
 * them: the names it stands in for, while it is loaded or can be asked for.
 */
export function embeddedFontNames(): string[] {
  if (configured === null || status === 'failed') return [];
  return status === 'loaded' || (canLoadFonts() && faces !== undefined) || reachable() ? aliases.slice() : [];
}

/** The four faces of the embedded font while they are on the page, with where their font files are: what a PDF embeds them from. */
export function embeddedFontFiles(): FontFileSource[] {
  const out: FontFileSource[] = [];
  if (status !== 'loaded' || !onPage || !faces) return out;
  for (let i = 0; i < EMBEDDED_FONT_FACES.length && i < faces.length; i++) {
    const file = faces[i] ? faces[i].file : undefined;
    if (file) out.push({ family: EMBEDDED_FONT_FAMILY, weight: EMBEDDED_FONT_FACES[i].weight, style: EMBEDDED_FONT_FACES[i].style, bytes: file.bytes, url: file.url });
  }
  return out;
}

/** Takes the faces off the page; they are kept. */
function leavePage(): void {
  if (!onPage || !faces) return;
  onPage = false;
  for (let i = 0; i < faces.length; i++) {
    try {
      faces[i].remove();
    } catch {
      // Nothing to take back.
    }
  }
}

/**
 * Sets where the embedded font comes from and what it stands in for, for the
 * whole page. Only the fields that are given change. Views that are open
 * measure their text anew when something changed, and a font that could not
 * be loaded before is tried again.
 */
export function configureEmbeddedFonts(options: EmbeddedFontsOptions): void {
  if (typeof options !== 'object' || options === null) throw new TypeError('The settings of the embedded font must be an object');
  let nextBase = configured;
  let nextAliases = aliases;
  let nextSource = source;
  let nextHas = hasFamily;
  if ('baseUrl' in options) {
    const value = options.baseUrl;
    if (value !== undefined && value !== null && (typeof value !== 'string' || value === '')) throw new TypeError('"baseUrl" must be an address, or null to turn the embedded font off');
    nextBase = value;
  }
  if (options.as !== undefined) nextAliases = fontAliases(options.as);
  if ('source' in options) {
    const value = options.source;
    if (value !== undefined && value !== null && typeof value !== 'function') throw new TypeError('"source" must be a function');
    nextSource = value || undefined;
  }
  if ('hasFamily' in options) {
    const value = options.hasFamily;
    if (value !== undefined && value !== null && typeof value !== 'function') throw new TypeError('"hasFamily" must be a function');
    nextHas = value || undefined;
  }
  const same = nextBase === configured && nextSource === source && nextHas === hasFamily && nextAliases.join('\n') === aliases.join('\n');
  if (same) return;
  configured = nextBase;
  aliases = nextAliases;
  aliasesLower = aliases.map((name) => name.toLowerCase());
  source = nextSource;
  hasFamily = nextHas;
  generation++;
  notNeeded.clear();
  // What failed may work now. What is loading ends as it ends, and is looked at with the settings of that moment.
  if (status === 'failed') status = 'idle';
  // Turned off: the font leaves the page (its files are kept), so that nothing is drawn with it any more.
  if (configured === null && status === 'loaded') {
    leavePage();
    status = 'idle';
  }
  notice = undefined;
  told = false;
  // The family lists of the views name the embedded font or not: they are built anew.
  notifyFontsChanged();
}

/**
 * Calls `listener` when the embedded font was needed and could not be
 * loaded. One listener is told, once: with several views on a page the user
 * is told one time. Returns what takes the listener off.
 */
export function onEmbeddedFontNotice(listener: (notice: EmbeddedFontNotice) => void): () => void {
  noticeListeners.push(listener);
  // The view that asked for the font starts to listen after it measured its text.
  if (notice && !told) later(deliver);
  return () => {
    const at = noticeListeners.indexOf(listener);
    if (at >= 0) noticeListeners.splice(at, 1);
  };
}

function later(run: () => void): void {
  Promise.resolve().then(run);
}

function deliver(): void {
  if (told || !notice || noticeListeners.length === 0) return;
  told = true;
  try {
    noticeListeners[0](notice);
  } catch {
    // A view that fails to say it does not stop the others from working.
  }
}

// ------------------------------------------------------------------ loading

let probeContext: { font: string; measureText(text: string): { width: number } } | null | undefined;

const PROBE_TEXT = 'mmmmmmmmmmlliWW@#01';

/**
 * Whether the device can show a family (sanitized), as the measurer of a
 * document finds out. True where it cannot be found out: nothing is fetched
 * on a guess.
 */
export function deviceHasFamily(name: string): boolean {
  if (hasFamily) return hasFamily(name) === true;
  if (typeof document === 'undefined') return true;
  const fonts = (document as { fonts?: { check?(font: string): boolean } }).fonts;
  if (fonts && typeof fonts.check === 'function') {
    try {
      // False is reliable; true is not in every browser.
      if (!fonts.check('16px "' + name + '"')) return false;
    } catch {
      // The widths below decide.
    }
  }
  if (probeContext === undefined) {
    try {
      const canvas = typeof document.createElement === 'function' ? document.createElement('canvas') : undefined;
      probeContext = (canvas && canvas.getContext('2d')) || null;
    } catch {
      probeContext = null;
    }
  }
  const context = probeContext;
  if (!context) return true;
  const generics = ['monospace', 'serif'];
  for (let i = 0; i < generics.length; i++) {
    context.font = '72px ' + generics[i];
    const plain = context.measureText(PROBE_TEXT).width;
    context.font = '72px "' + name + '", ' + generics[i];
    if (context.measureText(PROBE_TEXT).width !== plain) return true;
  }
  return false;
}

const reasonOf = (error: unknown): string => (error instanceof Error && error.message ? error.message : String(error));

function fail(family: string, reason: string): void {
  if (notice || told) return;
  notice = { family, reason };
  later(deliver);
}

/**
 * The load has ended: `read` are the four faces, or undefined when one of
 * them could not be read (`reason`). Called once per load, a moment after
 * the last file, never from inside the measuring that asked.
 */
function ended(read: PageFont[] | undefined, reason: string | undefined, family: string, mine: number): void {
  if (read) {
    // The same files wherever they came from: also a load that began under other settings is used.
    faces = read;
    if (configured === null) {
      // Turned off meanwhile: kept, not put on the page, and no view is told.
      status = 'idle';
      return;
    }
    try {
      for (let i = 0; i < read.length; i++) read[i].add();
      onPage = true;
      status = 'loaded';
      notifyFontsChanged();
      return;
    } catch (error) {
      onPage = true;
      leavePage();
      faces = undefined;
      reason = reasonOf(error);
    }
  }
  if (mine !== generation) {
    // It failed with settings that are not the settings any more: the views measure again, and the first text in
    // the family asks anew.
    status = 'idle';
    if (configured !== null) notifyFontsChanged();
    return;
  }
  status = 'failed';
  fail(family, reason === undefined ? 'The font files could not be loaded' : reason);
}

/** Gets the four faces: from what was read before, else from the files. `family` is the font name that asked. */
function start(family: string): void {
  status = 'loading';
  const mine = generation;
  if (faces) {
    const kept = faces;
    later(() => ended(kept, undefined, family, mine));
    return;
  }
  const folder = base();
  if (typeof source !== 'function' && folder === null) {
    later(() => ended(undefined, 'Where the font files of the package are is not known; set it with fonts.configure({ baseUrl }) or the fontsUrl of mount()', family, mine));
    return;
  }
  const give = source;
  const jobs = EMBEDDED_FONT_FACES.map((face) => {
    let given: ReturnType<NonNullable<typeof give>>;
    try {
      given = typeof give === 'function' ? give(face.file) : undefined;
    } catch (error) {
      return Promise.reject(error);
    }
    return Promise.resolve(given).then((data) => {
      if (data === undefined || data === null) {
        if (folder === null) throw new Error(`The font file ${face.file} was not given and there is no address to fetch it from`);
        data = folder + FOLDER + face.file;
      }
      return loadPageFont(EMBEDDED_FONT_FAMILY, data, face.weight, face.style);
    });
  });
  // All four or none: a face the browser makes up for a missing file would be drawn on lines measured for the real one.
  const read: PageFont[] = [];
  let open = jobs.length;
  let reason: string | undefined;
  const done = (): void => {
    if (--open > 0) return;
    ended(reason === undefined ? read : undefined, reason, family, mine);
  };
  for (let i = 0; i < jobs.length; i++) {
    jobs[i].then(
      (face) => {
        read[i] = face;
        done();
      },
      (error: unknown) => {
        if (reason === undefined) reason = reasonOf(error);
        done();
      },
    );
  }
}

/**
 * The embedded family when it stands in for a font name of a file, given as
 * splitFontFamily() gives it, and is on the page; undefined otherwise.
 * Asking is what loads the font: the first time a name it stands in for is
 * asked about where nothing else shows that name, the files are fetched.
 *
 * `covered` tells whether something in front of the embedded font in the
 * family list can be shown (the measurer of a document knows its list);
 * without it the device is tried for the two names.
 */
export function embeddedFontForNames(full: string, base: string, covered?: () => boolean): string | undefined {
  if (configured === null || full === '') return undefined;
  const fullLower = full.toLowerCase();
  if (aliasesLower.indexOf(fullLower) < 0 && aliasesLower.indexOf(base.toLowerCase()) < 0) return undefined;
  if (status === 'loaded') return EMBEDDED_FONT_FAMILY;
  if (status !== 'idle' || !canLoadFonts()) return undefined;
  const now = fontsVersion();
  // What a measurer found for its own list (its document may name an alternative) is not what another one finds:
  // only the answer of the device itself is kept.
  if (!covered && notNeeded.get(fullLower) === now) return undefined;
  let skip: boolean;
  try {
    skip = appFontFor(full, base, EMBEDDED_FONT_FAMILY) || (covered ? covered() : deviceHasFamily(full) || (base !== full && deviceHasFamily(base)));
  } catch {
    // The device could not be tried: nothing is fetched on a guess.
    skip = true;
  }
  if (!skip) start(base);
  else if (!covered) {
    if (notNeeded.size >= NOT_NEEDED_LIMIT) notNeeded.clear();
    notNeeded.set(fullLower, now);
  }
  // Not on the page yet: the lists are built anew when it is.
  return undefined;
}

/** embeddedFontForNames() for a font name as a file gives it. */
export function embeddedFontFor(name: unknown, covered?: () => boolean): string | undefined {
  if (configured === null || typeof name !== 'string') return undefined;
  const { full, base } = splitFontFamily(name);
  return embeddedFontForNames(full, base, covered);
}

/**
 * Forgets everything, for tests: the settings, what was loaded and what was
 * said. With `env`, the address of the script and of the page are the ones
 * given (undefined: none) instead of the real ones.
 */
export function resetEmbeddedFonts(env?: { scriptUrl?: string; pageUrl?: string }): void {
  configured = undefined;
  aliases = DEFAULT_AS.slice();
  aliasesLower = aliases.map((name) => name.toLowerCase());
  source = undefined;
  hasFamily = undefined;
  leavePage();
  faces = undefined;
  status = 'idle';
  generation++;
  notNeeded.clear();
  notice = undefined;
  told = false;
  noticeListeners.length = 0;
  probeContext = undefined;
  if (env) {
    script = env.scriptUrl;
    pageOverride = { page: env.pageUrl };
  } else {
    script = scriptAddress();
    pageOverride = undefined;
  }
}
