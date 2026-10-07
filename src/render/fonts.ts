// Font names and the fonts an application registers.
//
// Font names come from files and are never trusted: a name only reaches a CSS
// font string after sanitizeFontFamily() removed everything but letters,
// digits, spaces and hyphens from it. That holds for the names an application
// registers too, so that every family list is built the same way.
//
// Nothing here changes a model: the names a file uses are shown and measured
// through these functions and written back as they were read.

const MAX_FAMILY_LENGTH = 128;

/** Keeps letters, combining marks, digits, spaces and hyphens; the result is safe inside a quoted CSS family name. */
export function sanitizeFontFamily(name: unknown): string {
  if (typeof name !== 'string') return '';
  // Cut first: a surrogate pair split by the cut is removed by the filter below.
  const head = name.length > MAX_FAMILY_LENGTH ? name.slice(0, MAX_FAMILY_LENGTH) : name;
  return head
    // Everything in ASCII except letters, digits, space and hyphen goes; that covers the
    // characters that mean something in a font string. No \p{..} classes: old WebViews
    // fail to parse them and would lose the whole bundle.
    .replace(/[^A-Za-z0-9 \-\u00C0-\uFFFF]/g, '')
    // Halves of surrogate pairs left alone by the cut above.
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')
    .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1')
    .replace(/ +/g, ' ')
    .trim();
}

/**
 * The names old Windows programs give one font per code page: "Arial Tur",
 * "Arial CE", "Arial Cyr", "Arial Greek", "Arial Baltic" (the font
 * substitutes of Windows) and "Arial (Hebrew)", "Arial (Arabic)", ... (what
 * Word writes into the font table of an RTF file). No device has a font of
 * such a name; the font meant is the one in front of the suffix.
 *
 * A bare word is a suffix only when it is one of the five that Windows uses;
 * the others must be in brackets, because "Arial Hebrew" and "Noto Sans Thai"
 * are fonts of their own.
 */
const BARE_SUFFIX = /^(.*[^\s(])\s+(?:CE|Cyr|Greek|Tur|Baltic)$/i;
const BRACKET_SUFFIX = /^(.*[^\s(])\s*\((?:Turkish|Hebrew|Arabic|Vietnamese|Thai|Greek|Cyrillic|Baltic|Central European|Western)\)$/i;

/**
 * A font name of a file as it is and without its code page suffix, both
 * sanitized. `base` equals `full` for a name without such a suffix.
 */
export function splitFontFamily(name: unknown): { full: string; base: string } {
  const full = sanitizeFontFamily(name);
  if (full === '' || typeof name !== 'string') return { full, base: full };
  // The brackets do not survive sanitizing: the suffix is looked for in the name as it came.
  const raw = (name.length > MAX_FAMILY_LENGTH ? name.slice(0, MAX_FAMILY_LENGTH) : name).trim();
  const match = BRACKET_SUFFIX.exec(raw) || BARE_SUFFIX.exec(raw);
  if (!match) return { full, base: full };
  const base = sanitizeFontFamily(match[1]);
  // Never an empty family: such a name is left as it is.
  return { full, base: base === '' ? full : base };
}

/** The family a font name of a file stands for: the name without its code page suffix ("Arial Tur" is Arial), sanitized. */
export function baseFontFamily(name: unknown): string {
  return splitFontFamily(name).base;
}

// ------------------------------------------------------------------ registry

/** Where the data of a font comes from: an address, or the bytes of a font file (TrueType, OpenType, WOFF, WOFF2). */
export type FontSource = string | ArrayBuffer | Uint8Array | Blob;

/** A font an application adds to the ones the device has. */
export interface FontRegistration {
  /** The family name of the font. Only letters, digits, spaces and hyphens of it are kept. */
  family: string;
  /** The font file: its address, or its bytes. */
  source: FontSource;
  /** Default "normal". A number is a CSS weight from 1 to 1000. */
  weight?: 'normal' | 'bold' | number;
  /** Default "normal". */
  style?: 'normal' | 'italic';
  /**
   * Font names of files this font stands in for when the device does not
   * have them: with `as: ['Bookman Old Style']`, text a file sets in Bookman
   * Old Style (or "Bookman Old Style Tur") is shown and measured with this
   * font wherever the device lacks Bookman Old Style.
   */
  as?: string | string[];
  /** With `as`: this font is used even where the device has the font of that name. */
  prefer?: boolean;
}

/** One registered face, as registeredFonts() lists it. */
export interface RegisteredFont {
  family: string;
  /** The CSS weight: 400 for "normal", 700 for "bold". */
  weight: number;
  style: 'normal' | 'italic';
  /** The sanitized names of `as`. */
  as: string[];
  prefer: boolean;
  /** "loading" until the font file has been read; a font that failed is not listed. */
  status: 'loading' | 'loaded';
}

/** What is used of the browser's FontFace. */
interface Face {
  load(): Promise<unknown>;
}

interface FaceSet {
  add(face: Face): unknown;
  delete(face: Face): unknown;
}

interface Entry {
  family: string;
  lower: string;
  weight: number;
  style: 'normal' | 'italic';
  as: string[];
  asLower: string[];
  prefer: boolean;
  loaded: boolean;
  face: Face | undefined;
}

/** Faces kept at once; more are refused. */
const MAX_FACES = 64;
const MAX_ALIASES = 32;

let entries: Entry[] = [];
let version = 0;
const listeners: (() => void)[] = [];

/** Changes whenever the family lists built here may have changed: caches of font strings are keyed by it. */
export function fontsVersion(): number {
  return version;
}

/**
 * Calls `listener` whenever a registered font became usable or was removed:
 * text measured before is then measured anew. Returns what takes the listener off.
 */
export function onFontsChanged(listener: () => void): () => void {
  listeners.push(listener);
  return () => {
    const at = listeners.indexOf(listener);
    if (at >= 0) listeners.splice(at, 1);
  };
}

function changed(): void {
  version++;
  const all = listeners.slice();
  for (let i = 0; i < all.length; i++) {
    try {
      all[i]();
    } catch {
      // One view that fails must not keep the others from hearing of it.
    }
  }
}

/**
 * The registered fonts that stand in for a font name of a file, given as
 * splitFontFamily() gives it: `first` come before the name itself (they were
 * registered with `prefer`), `then` after it. Only fonts that have been loaded.
 */
export function replacementFonts(full: string, base: string): { first: string[]; then: string[] } | undefined {
  if (entries.length === 0) return undefined;
  const fullLower = full.toLowerCase();
  const baseLower = base.toLowerCase();
  let out: { first: string[]; then: string[] } | undefined;
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    if (!entry.loaded || (entry.asLower.indexOf(fullLower) < 0 && entry.asLower.indexOf(baseLower) < 0)) continue;
    if (!out) out = { first: [], then: [] };
    const list = entry.prefer ? out.first : out.then;
    // Several faces of one family name it once.
    if (list.indexOf(entry.family) < 0) list.push(entry.family);
  }
  return out;
}

/**
 * The families to try for a font name of a file, in order, sanitized and
 * without repeats: the registered fonts that are preferred to it, the name
 * itself, the family without the code page suffix, the other registered
 * fonts that stand in for it.
 */
export function fontFamilyChain(name: unknown): string[] {
  const { full, base } = splitFontFamily(name);
  if (full === '') return [];
  const registered = replacementFonts(full, base);
  if (!registered && base === full) return [full];
  const names: string[] = [];
  const add = (family: string) => {
    const lower = family.toLowerCase();
    for (let i = 0; i < names.length; i++) if (names[i].toLowerCase() === lower) return;
    names.push(family);
  };
  if (registered) for (let i = 0; i < registered.first.length; i++) add(registered.first[i]);
  add(full);
  add(base);
  if (registered) for (let i = 0; i < registered.then.length; i++) add(registered.then[i]);
  return names;
}

const describe = (entry: Entry): RegisteredFont => ({
  family: entry.family,
  weight: entry.weight,
  style: entry.style,
  as: entry.as.slice(),
  prefer: entry.prefer,
  status: entry.loaded ? 'loaded' : 'loading',
});

/** The fonts registered with registerFont(), in the order they were registered, a row per face. */
export function registeredFonts(): RegisteredFont[] {
  return entries.map(describe);
}

function faceSet(): FaceSet | undefined {
  const fonts = typeof document !== 'undefined' ? (document as { fonts?: Partial<FaceSet> }).fonts : undefined;
  return fonts && typeof fonts.add === 'function' && typeof fonts.delete === 'function' ? (fonts as FaceSet) : undefined;
}

function drop(entry: Entry): void {
  const at = entries.indexOf(entry);
  if (at >= 0) entries.splice(at, 1);
  const fonts = faceSet();
  if (entry.face && fonts) {
    try {
      fonts.delete(entry.face);
    } catch {
      // Nothing to take back.
    }
  }
}

/**
 * Takes every face of a family out again. Returns whether there was one.
 * Views that are open measure and lay out their text anew.
 */
export function unregisterFont(family: string): boolean {
  const lower = sanitizeFontFamily(family).toLowerCase();
  let found = false;
  let usable = false;
  const all = entries.slice();
  for (let i = 0; i < all.length; i++) {
    if (all[i].lower !== lower) continue;
    found = true;
    if (all[i].loaded) usable = true;
    drop(all[i]);
  }
  if (usable) changed();
  return found;
}

/** The weight as a number, so that "bold" and 700 are the same face. */
function weightOf(weight: unknown): number {
  if (weight === undefined || weight === 'normal') return 400;
  if (weight === 'bold') return 700;
  if (typeof weight === 'number' && weight >= 1 && weight <= 1000) return Math.round(weight * 100) / 100;
  throw new TypeError('The weight of a font must be "normal", "bold" or a number from 1 to 1000');
}

function styleOf(style: unknown): 'normal' | 'italic' {
  if (style === undefined || style === 'normal') return 'normal';
  if (style === 'italic') return 'italic';
  throw new TypeError('The style of a font must be "normal" or "italic"');
}

function aliasesOf(as: unknown): string[] {
  if (as === undefined || as === null) return [];
  const given = typeof as === 'string' ? [as] : as;
  if (!Array.isArray(given)) throw new TypeError('"as" must be a font name or a list of font names');
  if (given.length > MAX_ALIASES) throw new RangeError(`"as" names more than ${MAX_ALIASES} fonts`);
  const out: string[] = [];
  for (let i = 0; i < given.length; i++) {
    const clean = sanitizeFontFamily(given[i]);
    if (clean !== '' && out.indexOf(clean) < 0) out.push(clean);
  }
  return out;
}

/** The bytes of a Blob; old WebViews have no Blob.arrayBuffer(). */
function blobBytes(blob: Blob): Promise<ArrayBuffer> {
  if (typeof blob.arrayBuffer === 'function') return blob.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => (reader.result instanceof ArrayBuffer ? resolve(reader.result) : reject(new Error('The font file could not be read')));
    reader.onerror = () => reject(reader.error || new Error('The font file could not be read'));
    reader.readAsArrayBuffer(blob);
  });
}

/** What FontFace takes: the CSS `src` of an address, or the bytes. */
function faceSource(source: unknown): Promise<string | ArrayBuffer | Uint8Array> {
  if (typeof source === 'string') {
    if (source === '') return Promise.reject(new TypeError('The source of a font is an empty address'));
    // The address is the application's own; it only must not end the string it is put in.
    const escaped = source.replace(/[\\"\n\r\f]/g, (char) => (char === '"' ? '%22' : char === '\\' ? '%5C' : ''));
    return Promise.resolve('url("' + escaped + '")');
  }
  if (typeof ArrayBuffer !== 'undefined' && source instanceof ArrayBuffer) return Promise.resolve(source);
  if (typeof Uint8Array !== 'undefined' && source instanceof Uint8Array) return Promise.resolve(source);
  if (typeof Blob !== 'undefined' && source instanceof Blob) return blobBytes(source);
  return Promise.reject(new TypeError('The source of a font must be an address, an ArrayBuffer, a Uint8Array or a Blob'));
}

type FaceMaker = new (family: string, source: string | ArrayBuffer | Uint8Array, descriptors: { weight: string; style: string }) => Face;

/** Makes the face and reads the font file. */
function loadFace(make: FaceMaker, family: string, source: string | ArrayBuffer | Uint8Array, descriptors: { weight: string; style: string }): Promise<Face> {
  const attempt = (name: string): Promise<Face> => {
    let face: Face;
    try {
      face = new make(name, source, descriptors);
    } catch (error) {
      return Promise.reject(error);
    }
    return Promise.resolve(face.load()).then(() => face);
  };
  // The name is given as it is: browsers take it either literally or as a CSS family name, and both read a
  // plain name the same way. A browser that reads it as CSS refuses what is no list of identifiers (a word that
  // starts with a digit, a keyword such as "serif" or "inherit"); the name is given in quotes then. A font file
  // that is refused with the same error is refused again.
  return attempt(family).catch((error: unknown) => {
    if (!(error instanceof Error) || error.name !== 'SyntaxError') throw error;
    return attempt('"' + family + '"');
  });
}

/**
 * Adds a font to the ones the device has, for this page: text in a family of
 * that name is shown and measured with it, in documents and in spreadsheets,
 * on screen and in print, and so is text in the families named by `as`.
 *
 * Resolves once the font file has been read and the font can be used; views
 * that are open then measure and lay out their text anew. Rejects when the
 * file cannot be loaded or is no font, when the registration is not valid,
 * and where there are no web fonts (outside a browser, or in a browser
 * without the FontFace interface). Registering the same family, weight
 * ("bold" is 700) and style again replaces the face once the new one has
 * loaded, with one notice to the views: until then, and when the new one
 * cannot be loaded, the old one stays in use.
 */
export function registerFont(font: FontRegistration): Promise<RegisteredFont> {
  let entry: Entry;
  let make: FaceMaker;
  let fonts: FaceSet;
  try {
    if (typeof font !== 'object' || font === null) throw new TypeError('A font registration must be an object');
    const family = sanitizeFontFamily(font.family);
    if (family === '') throw new TypeError('A font needs a family name of letters, digits, spaces or hyphens');
    const weight = weightOf(font.weight);
    const style = styleOf(font.style);
    const as = aliasesOf(font.as);
    const maker = (typeof globalThis !== 'undefined' ? (globalThis as { FontFace?: unknown }).FontFace : undefined) as FaceMaker | undefined;
    const set = faceSet();
    if (typeof maker !== 'function' || !set) throw new Error('Fonts cannot be registered here: there is no FontFace and document.fonts (not a browser, or too old a one)');
    make = maker;
    fonts = set;
    const lower = family.toLowerCase();
    // A registration of the same face that is still loading is given up for this one. The face that works
    // stays until this one has loaded.
    const all = entries.slice();
    let faces = 1;
    for (let i = 0; i < all.length; i++) {
      const same = all[i].lower === lower && all[i].weight === weight && all[i].style === style;
      if (same && !all[i].loaded) drop(all[i]);
      else if (!same) faces++;
    }
    if (faces > MAX_FACES) throw new RangeError(`No more than ${MAX_FACES} fonts can be registered`);
    entry = { family, lower, weight, style, as, asLower: as.map((name) => name.toLowerCase()), prefer: font.prefer === true && as.length > 0, loaded: false, face: undefined };
    entries.push(entry);
  } catch (error) {
    return Promise.reject(error);
  }
  const current = () => entries.indexOf(entry) >= 0;
  const gone = () => new Error(`The font "${entry.family}" was unregistered or registered again while it was loading`);
  return faceSource(font.source)
    .then((source) => {
      if (!current()) throw gone();
      return loadFace(make, entry.family, source, { weight: String(entry.weight), style: entry.style });
    })
    .then(
      (face) => {
        if (!current()) throw gone();
        // The face this one replaces goes now, so that the views hear of the swap once.
        const others = entries.slice();
        for (let i = 0; i < others.length; i++) {
          if (others[i] !== entry && others[i].loaded && others[i].lower === entry.lower && others[i].weight === entry.weight && others[i].style === entry.style) drop(others[i]);
        }
        fonts.add(face);
        entry.face = face;
        entry.loaded = true;
        changed();
        return describe(entry);
      },
      (error: unknown) => {
        drop(entry);
        throw error;
      },
    );
}

/** A font of a list that is not usable when the wait for the list ended. */
export interface FontFailure {
  family: string;
  /** Why, in English. */
  reason: string;
  /** True when the font was still loading: it is used as soon as it arrives. */
  pending: boolean;
}

/**
 * Registers a list of fonts and waits for them, for `ms` at most. Never
 * rejects: resolves to the fonts that failed or did not arrive in time.
 * `timer` and `cancel` stand in for setTimeout and clearTimeout.
 */
export function settleFonts(
  list: readonly FontRegistration[],
  ms: number,
  register: (font: FontRegistration) => Promise<unknown> = registerFont,
  timer?: (callback: () => void, ms: number) => unknown,
  cancel?: (handle: unknown) => void,
): Promise<FontFailure[]> {
  if (list.length === 0) return Promise.resolve([]);
  return new Promise((resolve) => {
    const failures: (FontFailure | undefined)[] = [];
    let open = list.length;
    let over = false;
    let handle: unknown;
    let waiting = false;
    const end = () => {
      if (over) return;
      over = true;
      // The fonts were there first: the timer has nothing left to do.
      if (waiting) {
        if (!timer) clearTimeout(handle as ReturnType<typeof setTimeout>);
        else if (cancel) cancel(handle);
      }
      const out: FontFailure[] = [];
      for (let i = 0; i < list.length; i++) {
        const failure = failures[i];
        if (failure) out.push(failure);
      }
      resolve(out);
    };
    const name = (font: FontRegistration): string => (font && typeof font === 'object' ? sanitizeFontFamily(font.family) : '');
    for (let i = 0; i < list.length; i++) {
      failures[i] = { family: name(list[i]), reason: 'The font did not load in time', pending: true };
      let loading: Promise<unknown>;
      try {
        loading = Promise.resolve(register(list[i]));
      } catch (error) {
        loading = Promise.reject(error);
      }
      loading.then(
        () => {
          failures[i] = undefined;
          if (--open === 0) end();
        },
        (error: unknown) => {
          failures[i] = { family: name(list[i]), reason: error instanceof Error && error.message ? error.message : String(error), pending: false };
          if (--open === 0) end();
        },
      );
    }
    const late = () => {
      waiting = false;
      end();
    };
    handle = timer ? timer(late, ms) : setTimeout(late, ms);
    waiting = true;
  });
}
