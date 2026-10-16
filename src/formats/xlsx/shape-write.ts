// A shape of a drawing written from its model: the one way the engine writes a shape, for a shape
// that is added to a drawing part, one of the file the reader found to be exactly this
// (DrawingShape.known), and every shape of a workbook built anew (build.ts).
//
// What is written is only what the model holds, so that the reader reads the very same model back
// from it: a preset geometry, a solid fill or none, a solid outline or none (with its dash and the
// ends of a line), and a text body whose runs say their look in full. That is what makes a shape
// the engine wrote "known" when its file is opened again (see sameShapeElement()).
//
// The text is written with the prefixes "xdr" and "a"; whoever puts it into a part maps them onto
// the namespaces of that part.

import type { DrawingShape, DrawingText, DrawingTextRun, ShapeTextLook } from '../../model/sheet/types';
import { localName, type XmlElement, type XmlNode } from '../../xml';

const EMU_PER_PIXEL = 9525;
const MAX_EMU = 2e12;
const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const PRESET = /^[A-Za-z0-9]{1,64}$/;
/** The ends of a line a file may name (ST_LineEndType but "none"). */
export const LINE_ENDS = ['triangle', 'stealth', 'diamond', 'oval', 'arrow'];

/** How a writer that counts what it changes wants text made writable (see ShapeWriteOptions); set while a shape is written. */
let writing: ShapeWriteOptions | undefined;
const writable = (value: string): string => (writing && writing.clean ? writing.clean(String(value)) : clean(value));

const escapeText = (value: string): string => writable(value).replace(/[&<>\r]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&#13;'));
const escapeAttr = (value: string): string => writable(value).replace(/[&<"\t\n\r]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '"' ? '&quot;' : c === '\t' ? '&#9;' : c === '\n' ? '&#10;' : '&#13;'));
/** What XML cannot hold is left out: controls, non-characters, halves of characters. */
const clean = (text: string): string => String(text).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, (m) => (m.length === 2 ? m.charAt(0) : ''));

const emu = (value: number): number => (typeof value === 'number' && value > 0 && value < Infinity ? Math.round(Math.min(value, MAX_EMU)) : 0);

/** A box in EMU: where a shape is drawn on the sheet before it is turned. */
export interface ShapeBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Whether a shape is written as a connector (`xdr:cxnSp`): a line without text. */
export const writtenAsConnector = (shape: DrawingShape): boolean => shape.connector === true && !shape.text;

const solidFill = (color: string | undefined): string => (color !== undefined && HEX_COLOR.test(color) ? `<a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill>` : '');

/** `rot` of `a:xfrm`: 60,000ths of a degree from 1 to 21,599,999; undefined for no turn. */
function rotation(degrees: number | undefined): number | undefined {
  if (typeof degrees !== 'number' || !isFinite(degrees)) return undefined;
  const steps = Math.round((((degrees % 360) + 360) % 360) * 60000);
  return steps <= 0 || steps >= 21600000 ? undefined : steps;
}

/** The space around the text of a shape a body states nothing of, EMU: 0.1 inch at the sides, 0.05 inch above and below (as the reader takes it). */
const DEFAULT_INSETS = [91440, 45720, 91440, 45720];

const TEXT_ALIGN: { [align: string]: string } = { center: 'ctr', right: 'r', justify: 'just' };
const TEXT_ANCHOR: { [anchor: string]: string } = { top: 't', middle: 'ctr', bottom: 'b' };

/** The attributes and children of `a:rPr` (and `a:endParaRPr`) for the look of a run. */
function runLookXml(name: string, run: Omit<DrawingTextRun, 'text'>): string {
  let props = `<a:${name} lang="en-US"`;
  if (typeof run.size === 'number' && run.size >= 1 && run.size <= 4000) props += ` sz="${Math.round(run.size * 100)}"`;
  if (run.bold !== undefined) props += ` b="${run.bold ? 1 : 0}"`;
  if (run.italic !== undefined) props += ` i="${run.italic ? 1 : 0}"`;
  if (run.underline !== undefined) props += ` u="${run.underline ? 'sng' : 'none'}"`;
  if (run.strike !== undefined) props += ` strike="${run.strike ? 'sngStrike' : 'noStrike'}"`;
  let inner = solidFill(run.color);
  if (typeof run.font === 'string' && run.font !== '' && run.font.charAt(0) !== '+') inner += `<a:latin typeface="${escapeAttr(run.font.slice(0, 255))}"/>`;
  return inner === '' ? `${props}/>` : `${props}>${inner}</a:${name}>`;
}

/** The text body of a shape: from its text, or (a shape with none) one empty paragraph in the look text typed into it gets. */
function textBodyXml(text: DrawingText | undefined, look: ShapeTextLook | undefined): string {
  const frame = text || look || {};
  const wrap = (frame as { wrap?: boolean }).wrap;
  const insets = (frame as { insets?: number[] }).insets;
  const anchor = (frame as { anchor?: string }).anchor;
  let body = '<a:bodyPr';
  if (wrap === false) body += ' wrap="none"';
  const names = ['lIns', 'tIns', 'rIns', 'bIns'];
  // Always stated: a reader takes a body without them as the defaults, and the model it reads states them.
  for (let i = 0; i < 4; i++) {
    const inset = insets && typeof insets[i] === 'number' && insets[i] >= 0 && insets[i] < Infinity ? insets[i] * EMU_PER_PIXEL : DEFAULT_INSETS[i];
    body += ` ${names[i]}="${emu(inset)}"`;
  }
  body += ` anchor="${(anchor && TEXT_ANCHOR[anchor]) || 't'}"/>`;
  const end = runLookXml('endParaRPr', look ? look.run : {});
  let out = `<xdr:txBody>${body}<a:lstStyle/>`;
  const paragraphs = text ? text.paragraphs : [{ runs: [], align: look ? look.align : undefined }];
  for (const paragraph of paragraphs) {
    out += '<a:p>';
    const align = paragraph.align === undefined ? undefined : TEXT_ALIGN[paragraph.align];
    if (align) out += `<a:pPr algn="${align}"/>`;
    for (const run of paragraph.runs) {
      const props = runLookXml('rPr', run);
      // A line break inside a run is an element of its own in a drawing.
      const lines = String(run.text).split('\n');
      for (let i = 0; i < lines.length; i++) {
        if (i > 0) out += `<a:br>${props}</a:br>`;
        if (lines[i] !== '') out += `<a:r>${props}<a:t>${escapeText(lines[i])}</a:t></a:r>`;
      }
    }
    out += `${end}</a:p>`;
  }
  return `${out}</xdr:txBody>`;
}

/** The outline of a shape: a solid line with its width, dash and ends, or none. */
function outlineXml(shape: DrawingShape): string {
  const line = shape.line;
  const stroke = line ? solidFill(line.color) : '';
  if (!line || stroke === '') return '<a:ln><a:noFill/></a:ln>';
  const width = line.width > 0 && line.width < Infinity ? line.width : 1;
  const end = (name: string, type: string | undefined) => (type !== undefined && LINE_ENDS.indexOf(type) >= 0 ? `<a:${name} type="${type}"/>` : '');
  return `<a:ln w="${emu(width * EMU_PER_PIXEL)}">${stroke}${line.dash ? '<a:prstDash val="dash"/>' : ''}${end('headEnd', line.head)}${end('tailEnd', line.tail)}</a:ln>`;
}

/** What a writer that reports what it changes gives shapeElementXml(). */
export interface ShapeWriteOptions {
  /** Makes a text writable (its text, its name, its alternative text, a font): what a file cannot hold is replaced, and counted by the caller. Without it such characters are left out. */
  clean?(text: string): string;
  /** Called when the name of a shape is longer than a file holds (255 characters) and is cut. */
  onNameCut?(): void;
}

/** The longest name of a shape that is written (what the reader keeps of one). */
export const MAX_SHAPE_NAME = 255;

/**
 * The element of a shape (`xdr:sp`, or `xdr:cxnSp` for a line without text) with the id `id` its
 * part gives it, drawn in `box` (EMU; see ShapeBox). A shape without a name is named `fallback`.
 */
export function shapeElementXml(shape: DrawingShape, id: number, box: ShapeBox, fallback = 'Shape', options?: ShapeWriteOptions): string {
  const before = writing;
  writing = options;
  try {
    return elementXml(shape, id, box, fallback);
  } finally {
    writing = before;
  }
}

function elementXml(shape: DrawingShape, id: number, box: ShapeBox, fallback: string): string {
  const own = typeof shape.name === 'string' && shape.name !== '';
  if (own && shape.name!.length > MAX_SHAPE_NAME && writing && writing.onNameCut) writing.onNameCut();
  const name = own ? (shape.name as string).slice(0, MAX_SHAPE_NAME) : fallback;
  let names = `<xdr:cNvPr id="${Math.max(1, Math.floor(id) || 1)}" name="${escapeAttr(name)}"`;
  if (shape.description !== undefined && shape.description !== '') names += ` descr="${escapeAttr(shape.description)}"`;
  if (shape.hidden) names += ' hidden="1"';
  names += '/>';
  const rot = rotation(shape.rotation);
  const transform = `<a:xfrm${rot !== undefined ? ` rot="${rot}"` : ''}${shape.flipH ? ' flipH="1"' : ''}${shape.flipV ? ' flipV="1"' : ''}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(box.width)}" cy="${emu(box.height)}"/></a:xfrm>`;
  const geometry = `<a:prstGeom prst="${PRESET.test(shape.geometry) ? shape.geometry : 'rect'}"><a:avLst/></a:prstGeom>`;
  if (writtenAsConnector(shape)) {
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr>${names}<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${transform}${geometry}${outlineXml(shape)}</xdr:spPr></xdr:cxnSp>`;
  }
  const props = `<xdr:spPr>${transform}${geometry}${solidFill(shape.fill) || '<a:noFill/>'}${outlineXml(shape)}</xdr:spPr>`;
  const body = shape.text || shape.textLook ? textBodyXml(shape.text, shape.textLook) : '';
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>${names}<xdr:cNvSpPr${shape.textBox ? ' txBox="1"' : ''}/></xdr:nvSpPr>${props}${body}</xdr:sp>`;
}

/** The plain text of a text: its paragraphs one to a line. */
export function shapeTextOf(text: DrawingText | undefined): string {
  if (!text) return '';
  return text.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
}

const MAX_COMPARED = 20000;

/**
 * Whether two elements say the same: the same local names, attributes and children, in order;
 * text alike where an element holds no element (white space between elements is no content). The
 * place a transform gives (`a:off`, `a:ext` of `a:xfrm`) is not compared: the anchor says where a
 * shape is, and the writer gives the transform again from it. What is nested beyond `limit`
 * elements is not looked at: the answer is no.
 */
export function sameShapeElement(a: XmlElement, b: XmlElement): boolean {
  let left = MAX_COMPARED;
  const attrs = (el: XmlElement): string[] => {
    const out: string[] = [];
    for (const [name, value] of el.attrs) if (name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:') out.push(`${localName(name)}\u0000${value}`);
    return out.sort();
  };
  const content = (el: XmlElement): XmlNode[] => {
    const hasElement = el.children.some((node) => node.type === 'element');
    return el.children.filter((node) => !(hasElement && node.type === 'text' && node.value.trim() === ''));
  };
  const same = (x: XmlElement, y: XmlElement, placed: boolean, depth: number): boolean => {
    if (--left < 0 || depth > 64) return false;
    if (localName(x.name) !== localName(y.name)) return false;
    if (!placed) {
      const ax = attrs(x);
      const ay = attrs(y);
      if (ax.length !== ay.length || ax.some((value, i) => value !== ay[i])) return false;
    }
    const cx = content(x);
    const cy = content(y);
    if (cx.length !== cy.length) return false;
    const transform = localName(x.name) === 'xfrm';
    for (let i = 0; i < cx.length; i++) {
      const nx = cx[i];
      const ny = cy[i];
      if (nx.type !== ny.type) return false;
      if (nx.type === 'text') {
        if (nx.value !== (ny as typeof nx).value) return false;
      } else if (nx.type === 'element') {
        const local = localName(nx.name);
        if (!same(nx, ny as XmlElement, transform && (local === 'off' || local === 'ext'), depth + 1)) return false;
      } else return false;
    }
    return true;
  };
  return same(a, b, false, 0);
}
