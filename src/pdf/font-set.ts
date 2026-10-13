// Which font file the text of a canvas is painted with. A painter names its
// font the way CSS does ("italic bold 14px "Bookman Old Style", "TeX Gyre
// Bonum", serif"); a PDF needs the file. The files that are known are those
// the application registered and the font the package brings along; a font
// of the device has no file here, and text in it cannot be written as text.
//
// The answer must be the font the browser paints with, or none: the families
// are tried in the order of the list, as a browser does, and the first one
// that can be shown decides - a family the device has ends the search without
// a font, because the browser paints with that one.

import { FontError, fontLicence, readFont, type PdfFontFace } from './sfnt';

/** A face whose file is known. */
export interface PdfFontSource {
  family: string;
  /** The CSS weight the face was registered with. */
  weight: number;
  style: 'normal' | 'italic';
  /** The font file (TrueType or OpenType); undefined for a face whose file is not at hand. */
  bytes: Uint8Array | undefined;
}

export interface ParsedFont {
  style: 'normal' | 'italic';
  weight: number;
  /** In px. */
  size: number;
  families: string[];
}

const GENERIC = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', 'ui-serif', 'ui-sans-serif', 'ui-monospace', 'ui-rounded', 'emoji', 'math', 'fangsong', '-apple-system', 'blinkmacsystemfont'];
const MAX_FONT_STRING = 4096;
const MAX_FAMILIES = 64;

/**
 * The value of `font` of a canvas as its parts: "[italic] [bold | <number>]
 * <size>px <families>". Undefined for a value that is not of that form
 * (another unit, small caps, a line height that is not "normal", a system
 * font): what the engine's painters write is of that form.
 */
export function parseCssFont(font: string): ParsedFont | undefined {
  if (typeof font !== 'string' || font.length > MAX_FONT_STRING) return undefined;
  const parts = /^\s*((?:(?:normal|italic|oblique|bold|bolder|lighter|[1-9]\d{0,3}(?:\.\d+)?)\s+)*)(\d+(?:\.\d+)?|\.\d+)px(?:\/normal)?\s+(.+?)\s*$/.exec(font);
  if (!parts) return undefined;
  let style: 'normal' | 'italic' = 'normal';
  let weight = 400;
  const words = parts[1].split(/\s+/);
  for (let i = 0; i < words.length; i++) {
    const word = words[i];
    if (word === '' || word === 'normal') continue;
    if (word === 'italic') style = 'italic';
    else if (word === 'bold') weight = 700;
    // A slanted upright face, or a weight that depends on what surrounds the text: not what a file says.
    else if (word === 'oblique' || word === 'bolder' || word === 'lighter') return undefined;
    else {
      weight = Number(word);
      if (!(weight >= 1 && weight <= 1000)) return undefined;
    }
  }
  const size = Number(parts[2]);
  if (!(size > 0) || !isFinite(size)) return undefined;
  // The families: quoted, or plain words; a comma inside quotes is part of the name.
  const families: string[] = [];
  const list = parts[3];
  let at = 0;
  while (at < list.length) {
    if (families.length >= MAX_FAMILIES) return undefined;
    while (at < list.length && /\s/.test(list.charAt(at))) at++;
    const quote = list.charAt(at);
    let family: string;
    if (quote === '"' || quote === "'") {
      const end = list.indexOf(quote, at + 1);
      if (end < 0) return undefined;
      family = list.slice(at + 1, end);
      // An escape inside a name is not read here.
      if (family.indexOf('\\') >= 0) return undefined;
      at = end + 1;
      while (at < list.length && /\s/.test(list.charAt(at))) at++;
      if (at < list.length && list.charAt(at) !== ',') return undefined;
    } else {
      let end = list.indexOf(',', at);
      if (end < 0) end = list.length;
      family = list.slice(at, end).trim().replace(/\s+/g, ' ');
      if (!/^-?[A-Za-z_][\w-]*( -?[A-Za-z_][\w-]*)*$/.test(family)) return undefined;
      at = end;
    }
    if (family === '') return undefined;
    families.push(family);
    at++;
  }
  if (families.length === 0) return undefined;
  return { style, weight, size, families };
}

interface Loaded {
  weight: number;
  style: 'normal' | 'italic';
  face?: PdfFontFace;
  /** Why the file cannot be used. */
  reason?: string;
}

export type ResolvedFont = { face: PdfFontFace; size: number } | { reason: string };

const cachedFaces: { bytes: Uint8Array; face?: PdfFontFace; reason?: string }[] = [];
const MAX_CACHED_FACES = 64;

/** A font file, read once: the same bytes are the same face (what a file embeds a font by). */
function faceOf(bytes: Uint8Array): { face?: PdfFontFace; reason?: string } {
  for (let i = 0; i < cachedFaces.length; i++) if (cachedFaces[i].bytes === bytes) return cachedFaces[i];
  const entry: { bytes: Uint8Array; face?: PdfFontFace; reason?: string } = { bytes };
  try {
    const face = readFont(bytes);
    const licence = fontLicence(face);
    if (licence.embed) entry.face = face;
    else entry.reason = licence.reason;
  } catch (error) {
    entry.reason = error instanceof FontError ? error.message : 'the font file could not be read';
  }
  if (cachedFaces.length >= MAX_CACHED_FACES) cachedFaces.shift();
  cachedFaces.push(entry);
  return entry;
}

/**
 * The fonts a PDF can be written with. `deviceHas` says whether the device
 * shows a family by itself (true where that cannot be told: a font is then
 * taken to be the device's, and its text stays a picture).
 */
export class PdfFontSet {
  private readonly families = new Map<string, Loaded[]>();
  private readonly resolved = new Map<string, ResolvedFont>();

  constructor(
    sources: readonly PdfFontSource[],
    private readonly deviceHas: (family: string) => boolean = () => true,
  ) {
    for (let i = 0; i < sources.length; i++) {
      const source = sources[i];
      if (!source || typeof source.family !== 'string') continue;
      const key = source.family.toLowerCase();
      let list = this.families.get(key);
      if (!list) {
        list = [];
        this.families.set(key, list);
      }
      const read = source.bytes instanceof Uint8Array ? faceOf(source.bytes) : { face: undefined, reason: 'its font file is not at hand' };
      const entry: Loaded = { weight: source.weight >= 1 && source.weight <= 1000 ? source.weight : 400, style: source.style === 'italic' ? 'italic' : 'normal', face: read.face, reason: read.reason };
      // A face registered again takes the place of the one before it.
      let replaced = false;
      for (let k = 0; k < list.length; k++) {
        if (list[k].weight === entry.weight && list[k].style === entry.style) {
          list[k] = entry;
          replaced = true;
        }
      }
      if (!replaced) list.push(entry);
    }
  }

  /** Whether there is any font a page could be written with. */
  get empty(): boolean {
    return this.families.size === 0;
  }

  /** The face of a family for a weight and a style, as CSS chooses among the faces of a family; a reason when the browser would make one up. */
  private pick(list: Loaded[], parsed: ParsedFont, family: string): ResolvedFont {
    const same = list.filter((entry) => entry.style === parsed.style);
    if (same.length === 0) return { reason: `the font "${family}" has no ${parsed.style === 'italic' ? 'italic' : 'upright'} face: the browser makes one up` };
    const wanted = parsed.weight;
    const rank = (weight: number): number => {
      // Nearest first, in the order CSS looks: 400 and 500 try each other, lighter weights look down first, bolder ones up.
      if (weight === wanted) return 0;
      if (wanted >= 400 && wanted <= 500) return weight >= wanted && weight <= 500 ? weight - wanted : weight < wanted ? 1000 + wanted - weight : 2000 + weight - wanted;
      if (wanted < 400) return weight < wanted ? wanted - weight : 1000 + weight - wanted;
      return weight > wanted ? weight - wanted : 1000 + wanted - weight;
    };
    let best = same[0];
    for (let i = 1; i < same.length; i++) if (rank(same[i].weight) < rank(best.weight)) best = same[i];
    if (wanted >= 600 && best.weight < 600) return { reason: `the font "${family}" has no bold face: the browser makes one up` };
    if (!best.face) return { reason: `the font "${family}" cannot be embedded: ${best.reason || 'its file could not be read'}` };
    return { face: best.face, size: parsed.size };
  }

  /** The face and the size (px) text in a CSS font is painted with, or why no font file stands for it. */
  resolve(font: string): ResolvedFont {
    const known = this.resolved.get(font);
    if (known) return known;
    const parsed = parseCssFont(font);
    let out: ResolvedFont | undefined;
    if (!parsed) out = { reason: 'the font of the text is given in a form that is not read' };
    else {
      for (let i = 0; i < parsed.families.length && !out; i++) {
        const family = parsed.families[i];
        const list = this.families.get(family.toLowerCase());
        if (list) out = this.pick(list, parsed, family);
        else if (GENERIC.indexOf(family.toLowerCase()) >= 0) out = { reason: `text is painted with a font of the device ("${family}")` };
        else {
          let has = true;
          try {
            has = this.deviceHas(family) !== false;
          } catch {
            has = true;
          }
          if (has) out = { reason: `text is painted with a font of the device ("${family.slice(0, 64)}")` };
        }
      }
      if (!out) out = { reason: 'no font of the text is known' };
    }
    if (this.resolved.size > 2000) this.resolved.clear();
    this.resolved.set(font, out);
    return out;
  }
}
