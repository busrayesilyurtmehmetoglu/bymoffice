// Shared pieces of the docx reader: namespace handling, attribute readers
// that validate what they return, and the bookkeeping behind the lossless
// principle (which children and attributes were understood).

import type { AttrList, DocRelationship, ExtraXml, MiscXml, ParaProps, RunProps, Wrapper } from '../../model/doc/types';
import type { XmlComment, XmlElement, XmlNode, XmlPI } from '../../xml';

/** Tables, content controls and text boxes nested deeper than this are refused. */
export const MAX_NESTING = 64;
/** Depth limit for XML trees the model keeps as read. */
export const MAX_XML_DEPTH = 256;
/** Largest length accepted, in twips (about 17 m; Word itself stops at 22 inches). */
export const MAX_TWIPS = 1000000;
/** Largest drawing coordinate accepted, in EMU. */
export const MAX_EMU = 0x7fffffff;

const NAMESPACES: [uri: string, prefix: string][] = [
  ['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w'],
  ['http://purl.oclc.org/ooxml/wordprocessingml/main', 'w'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r'],
  ['http://purl.oclc.org/ooxml/officeDocument/relationships', 'r'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/math', 'm'],
  ['http://purl.oclc.org/ooxml/officeDocument/math', 'm'],
  ['http://schemas.openxmlformats.org/markup-compatibility/2006', 'mc'],
  ['http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'wp'],
  ['http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing', 'wp'],
  ['http://schemas.openxmlformats.org/drawingml/2006/main', 'a'],
  ['http://purl.oclc.org/ooxml/drawingml/main', 'a'],
  ['urn:schemas-microsoft-com:vml', 'v'],
  ['urn:schemas-microsoft-com:office:office', 'o'],
];

/** A namespace of the newer chart parts: 2014/chartex, and the later ones that mark a version (2015/9/8/chartex, ...). */
export const CHARTEX_SPACE = /^http:\/\/schemas\.microsoft\.com\/office\/drawing\/20\d\d(\/\d{1,2}){0,3}\/chartex$/;

export const STRICT_W_NAMESPACE = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
/** Every namespace the reader takes for WordprocessingML. */
export const WORD_NAMESPACES: readonly string[] = NAMESPACES.filter((entry) => entry[1] === 'w').map((entry) => entry[0]);

/** State shared by all parts of one document. */
export interface ReaderState {
  /** Property bags by their source text: equal properties share one object. */
  runProps: Map<string, RunProps>;
  paraProps: Map<string, ParaProps>;
  warnings: string[];
  unknown: Map<string, number>;
  strict: boolean;
  /** Chart parts the frames of the document name, as the package spells them, each once, in the order the frames were met. */
  chartParts: Set<string>;
  /**
   * The name the package has a part under (part names are compared without regard to case, so a relationship may spell
   * one otherwise); undefined for a part it does not have. Set by whoever reads a package. Without it every part a
   * relationship names is taken to be there, under the name the relationship gives it.
   */
  partName?(part: string): string | undefined;
  /** Warnings about frames of charts said so far: see MAX_CHART_FRAME_WARNINGS. */
  chartWarnings: number;
  /**
   * The parts of the newer chart types (chartex) that were read into charts before the stories
   * were, as the package spells them: a Choice of alternate content that holds a frame of such a
   * chart is taken only when its chart is among them, else the Fallback is.
   */
  chartExDrawn?: Set<string>;
}

/** Warnings about frames of charts that lead nowhere, for one document; further ones are not said. */
export const MAX_CHART_FRAME_WARNINGS = 20;

export function createReaderState(): ReaderState {
  return { runProps: new Map(), paraProps: new Map(), warnings: [], unknown: new Map(), strict: false, chartParts: new Set(), chartWarnings: 0 };
}

/**
 * What is needed to read one part.
 *
 * Namespace prefixes are taken from the root element of the part. A prefix
 * that is declared further down is taken at face value, which is right for
 * the conventional prefixes every producer uses (a:, pic:, wp:, v:, ...).
 */
export interface PartContext {
  state: ReaderState;
  part: string;
  relationships: Map<string, DocRelationship>;
  /** Qualified element name -> name the reader switches on: local name for WordprocessingML, "prefix:local" otherwise. */
  kinds: Map<string, string>;
  /** Declared prefix -> conventional prefix, for the namespaces the reader knows. */
  prefixes: Map<string, string>;
  /** Local attribute name -> the qualified spellings to try. */
  attrNames: Map<string, string[]>;
  /** Prefixes (with the colon) bound to WordprocessingML and to the relationships namespace. */
  wPrefixes: string[];
  rPrefixes: string[];
  /** Attributes found by the attribute readers since the counter was last reset. */
  hits: number;
}

export function createContext(
  state: ReaderState,
  part: string,
  rootAttrs: AttrList,
  relationships: Map<string, DocRelationship>,
): PartContext {
  const prefixes = new Map<string, string>();
  const wPrefixes: string[] = [];
  const rPrefixes: string[] = [];
  for (const [name, value] of rootAttrs) {
    if (name !== 'xmlns' && !name.startsWith('xmlns:')) continue;
    // The namespaces of the newer chart parts (chartex) in all their versions: what a Choice of alternate content may require.
    const known = NAMESPACES.find(([uri]) => uri === value) || (CHARTEX_SPACE.test(value) ? ['', 'cx'] : undefined);
    if (!known) continue;
    const prefix = name === 'xmlns' ? '' : name.slice(6);
    prefixes.set(prefix, known[1]);
    if (known[1] === 'w') {
      if (prefix) wPrefixes.push(prefix + ':');
      if (value === STRICT_W_NAMESPACE) state.strict = true;
    } else if (known[1] === 'r' && prefix) rPrefixes.push(prefix + ':');
  }
  // A part that declares nothing (hand-made or damaged) is read with the conventional prefixes.
  if (!prefixes.size) {
    prefixes.set('w', 'w');
    wPrefixes.push('w:');
  }
  if (!rPrefixes.length) rPrefixes.push('r:');
  return { state, part, relationships, kinds: new Map(), prefixes, attrNames: new Map(), wPrefixes, rPrefixes, hits: 0 };
}

/** The name the reader switches on: "p" for <w:p>, "m:oMath" for math, "wp:inline", ... */
export function kindOf(ctx: PartContext, name: string): string {
  let kind = ctx.kinds.get(name);
  if (kind === undefined) {
    const colon = name.indexOf(':');
    const prefix = colon < 0 ? '' : name.slice(0, colon);
    const local = name.slice(colon + 1);
    const known = ctx.prefixes.get(prefix);
    if (known === 'w') kind = local;
    else if (known !== undefined) kind = `${known}:${local}`;
    // An element without prefix in a part whose default namespace is not a known one.
    else kind = prefix ? name : `:${local}`;
    // The cache only ever holds the distinct names of one part, but a hostile part can have many.
    if (ctx.kinds.size < 4096) ctx.kinds.set(name, kind);
  }
  return kind;
}

/** WordprocessingML attribute by local name; counts as understood when found. */
export function attr(ctx: PartContext, el: XmlElement, local: string): string | undefined {
  if (el.attrs.size === 0) return undefined;
  let names = ctx.attrNames.get(local);
  if (names === undefined) {
    names = ctx.wPrefixes.map((prefix) => prefix + local);
    // Unqualified last: wrong by the schema, but written by some producers.
    names.push(local);
    ctx.attrNames.set(local, names);
  }
  for (let i = 0; i < names.length; i++) {
    const value = el.attrs.get(names[i]);
    if (value !== undefined) {
      ctx.hits++;
      return value;
    }
  }
  return undefined;
}

/** Relationship id attribute (r:id, r:embed, ...) by local name. */
export function relAttr(ctx: PartContext, el: XmlElement, local: string): string | undefined {
  for (const prefix of ctx.rPrefixes) {
    const value = el.attrs.get(prefix + local);
    if (value !== undefined) {
      ctx.hits++;
      return value;
    }
  }
  return undefined;
}

const INTEGER = /^-?\d{1,10}$/;
const MEASURE = /^(-?(?:\d{1,10}(?:\.\d{0,10})?|\.\d{1,10}))(mm|cm|in|pt|pc|pi)?$/;
const POINTS_PER_UNIT: Record<string, number> = { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 };

/** Integer in [min, max]; undefined (and not counted as understood) for anything else. */
export function intAttr(ctx: PartContext, el: XmlElement, local: string, min: number, max: number): number | undefined {
  const text = attr(ctx, el, local);
  if (text === undefined) return undefined;
  if (INTEGER.test(text)) {
    const value = Number(text);
    if (value >= min && value <= max) return value;
  }
  ctx.hits--;
  return undefined;
}

/**
 * A measurement: a plain number in the unit the attribute is defined in, or a
 * number with a unit ("2.5cm", "12pt"), which strict documents use. The
 * result is in the attribute's own unit; `perPoint` says how many make a point.
 */
export function parseMeasure(text: string, perPoint: number, min: number, max: number): number | undefined {
  const m = MEASURE.exec(text.trim());
  if (!m) return undefined;
  let value = Number(m[1]);
  if (m[2] !== undefined) value = Math.round(value * POINTS_PER_UNIT[m[2]] * perPoint);
  return value >= min && value <= max ? value : undefined;
}

export function measureAttr(
  ctx: PartContext,
  el: XmlElement,
  local: string,
  perPoint: number,
  min: number,
  max: number,
): number | undefined {
  const text = attr(ctx, el, local);
  if (text === undefined) return undefined;
  const value = parseMeasure(text, perPoint, min, max);
  if (value === undefined) ctx.hits--;
  return value;
}

/** Length in twips, either sign. */
export const twipsAttr = (ctx: PartContext, el: XmlElement, local: string) =>
  measureAttr(ctx, el, local, 20, -MAX_TWIPS, MAX_TWIPS);

/** Length in twips that cannot be negative. */
export const utwipsAttr = (ctx: PartContext, el: XmlElement, local: string) => measureAttr(ctx, el, local, 20, 0, MAX_TWIPS);

export function parseOnOff(text: string | undefined): boolean | undefined {
  if (text === undefined) return true;
  if (text === '1' || text === 'true' || text === 'on') return true;
  if (text === '0' || text === 'false' || text === 'off') return false;
  return undefined;
}

/** ST_OnOff attribute; undefined when absent or not a boolean. */
export function boolAttr(ctx: PartContext, el: XmlElement, local: string): boolean | undefined {
  const text = attr(ctx, el, local);
  if (text === undefined) return undefined;
  const value = parseOnOff(text);
  if (value === undefined) ctx.hits--;
  return value;
}

/** Value of an on/off element (<w:b/>, <w:b w:val="0"/>); undefined when the value is not a boolean. */
export function onOff(ctx: PartContext, el: XmlElement): boolean | undefined {
  const text = attr(ctx, el, 'val');
  const value = parseOnOff(text);
  if (value === undefined) ctx.hits--;
  return value;
}

/** Element children in order. */
export function elements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const node of el.children) if (node.type === 'element') out.push(node);
  return out;
}

const isBlank = (text: string): boolean => {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c !== 32 && c !== 9 && c !== 10 && c !== 13) return false;
  }
  return true;
};

/** A comment or a processing instruction. */
export const isMisc = (node: XmlNode): node is XmlComment | XmlPI => node.type === 'comment' || node.type === 'pi';

/**
 * True when an element has, directly in it, something a reader of its child
 * elements does not see: a comment, a processing instruction or text that is
 * more than whitespace.
 */
export function hasLooseContent(el: XmlElement): boolean {
  for (const node of el.children) {
    if (node.type === 'comment' || node.type === 'pi' || (node.type === 'text' && !isBlank(node.value))) return true;
  }
  return false;
}

/** True when an element has content at all: child elements, comments, processing instructions or text. */
export function hasContent(el: XmlElement): boolean {
  for (const node of el.children) {
    if (node.type !== 'text' || !isBlank(node.value)) return true;
  }
  return false;
}

/** Appends a comment or processing instruction to a `misc` list; `index` is the number of elements before it. */
export function pushMisc(list: MiscXml[] | undefined, index: number, node: XmlComment | XmlPI, wrap?: Wrapper): MiscXml[] {
  const item: MiscXml = { index, node };
  if (wrap) item.wrap = wrap;
  if (list) {
    list.push(item);
    return list;
  }
  return [item];
}

/** First child with the given kind (see kindOf). */
export function childOfKind(ctx: PartContext, el: XmlElement, kind: string): XmlElement | undefined {
  for (const node of el.children) {
    if (node.type === 'element' && kindOf(ctx, node.name) === kind) return node;
  }
  return undefined;
}

/** Direct text of an element (CDATA included), child elements ignored. */
export function textOf(el: XmlElement): string {
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

export const attrList = (el: XmlElement): AttrList | undefined => (el.attrs.size ? Array.from(el.attrs) : undefined);

/** File content quoted in an error message, cut so a hostile value cannot produce a huge message. */
export const quote = (text: string) => (text.length > 40 ? `"${text.slice(0, 40)}…"` : `"${text}"`);

function checkDepth(el: XmlElement, depth: number): void {
  if (depth > MAX_XML_DEPTH) throw new Error(`Invalid docx: XML is nested deeper than ${MAX_XML_DEPTH} levels`);
  for (const node of el.children) if (node.type === 'element') checkDepth(node, depth + 1);
}

/** Refuses a tree nested deeper than MAX_XML_DEPTH; returns it otherwise. */
export function bounded(el: XmlElement): XmlElement {
  checkDepth(el, 0);
  return el;
}

/**
 * Accepts a tree into the model as it is. Trees kept this way are bounded in
 * depth, so code that walks them later can recurse.
 */
export function keep(ctx: PartContext, el: XmlElement, context: string, partial = false): XmlElement {
  checkDepth(el, 0);
  const key = `${context}/${el.name}${partial ? ' (partial)' : ''}`;
  const unknown = ctx.state.unknown;
  // Names come from the file; the table must not grow without bound.
  if (unknown.has(key) || unknown.size < 2000) unknown.set(key, (unknown.get(key) || 0) + 1);
  return el;
}

export function pushExtra(
  ctx: PartContext,
  list: ExtraXml[] | undefined,
  index: number,
  el: XmlElement,
  context: string,
  partial: boolean,
  wrap?: Wrapper,
): ExtraXml[] {
  const extra: ExtraXml = { index, el: keep(ctx, el, context, partial) };
  if (partial) extra.partial = true;
  if (wrap) extra.wrap = wrap;
  if (list) {
    list.push(extra);
    return list;
  }
  return [extra];
}

const TEXT_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#13;' };
const ATTR_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '"': '&quot;', '\t': '&#9;', '\n': '&#10;', '\r': '&#13;' };

function writeElement(el: XmlElement, out: string[], depth: number): void {
  if (depth > MAX_XML_DEPTH) throw new Error(`Invalid docx: XML is nested deeper than ${MAX_XML_DEPTH} levels`);
  out.push('<', el.name);
  for (const [name, value] of el.attrs) out.push(' ', name, '="', value.replace(/[&<"\t\n\r]/g, (ch) => ATTR_ESCAPES[ch]), '"');
  if (el.children.length === 0) {
    out.push('/>');
    return;
  }
  out.push('>');
  for (const node of el.children) {
    if (node.type === 'element') writeElement(node, out, depth + 1);
    else if (node.type === 'text') out.push(node.value.replace(/[&<>\r]/g, (ch) => TEXT_ESCAPES[ch]));
    else if (node.type === 'comment') out.push('<!--', node.value, '-->');
    else out.push('<?', node.target, node.value ? ' ' + node.value : '', '?>');
  }
  out.push('</', el.name, '>');
}

/** XML text of an element, comments and processing instructions included. Also the key property bags are shared by. */
export function serializeElement(el: XmlElement): string {
  const out: string[] = [];
  writeElement(el, out, 0);
  return out.join('');
}

/** Descendants (not the element itself) with the given local name, without entering matches; bounded in depth. */
export function findDescendants(el: XmlElement, local: string, out: XmlElement[] = [], depth = 0): XmlElement[] {
  if (depth > MAX_XML_DEPTH) throw new Error(`Invalid docx: XML is nested deeper than ${MAX_XML_DEPTH} levels`);
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (node.name.slice(node.name.indexOf(':') + 1) === local) out.push(node);
    else findDescendants(node, local, out, depth + 1);
  }
  return out;
}

/** Attribute by local name whatever its prefix (for DrawingML and VML, whose attributes the reader only picks from). */
export function anyAttr(el: XmlElement, local: string): string | undefined {
  const direct = el.attrs.get(local);
  if (direct !== undefined) return direct;
  for (const [name, value] of el.attrs) {
    const colon = name.indexOf(':');
    if (colon >= 0 && name.length - colon - 1 === local.length && name.endsWith(local) && !name.startsWith('xmlns:')) return value;
  }
  return undefined;
}
