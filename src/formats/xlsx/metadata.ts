// Cell metadata (xl/metadata.xml): which cells hold dynamic array formulas.
//
// Excel marks the cell of a dynamic array formula with `cm="n"`: the n-th block of
// <cellMetadata>, whose record points at a <futureMetadata name="XLDAPR"> block that says
// `fDynamic="1"`. Only that is read here; the part is otherwise kept as the file has it. When
// a formula of the model becomes a dynamic array formula, the writer reuses such a block, adds
// one to a part that has none, or makes the part.

import { decodeXmlBytes, localName, parseXml, type XmlElement } from '../../xml';

export const METADATA_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
/** Last part of the type of the relationship from the workbook to the part. */
export const METADATA_RELATIONSHIP = 'sheetMetadata';

const DYNAMIC_ARRAY_TYPE = 'XLDAPR';
const DYNAMIC_ARRAY_URI = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';
const NS_DYNAMIC_ARRAY = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Blocks looked at in one list; a part with more (a damaged one) is taken to mark no dynamic arrays beyond them. */
const MAX_BLOCKS = 1 << 20;

/** The attributes of the metadata type, as Excel writes them. */
const TYPE_ATTRS = 'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"';

function children(parent: XmlElement, local: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const node of parent.children) {
    if (node.type === 'element' && localName(node.name) === local) {
      out.push(node);
      if (out.length >= MAX_BLOCKS) break;
    }
  }
  return out;
}

function attr(el: XmlElement, local: string): string | undefined {
  for (const [name, value] of el.attrs) if (localName(name) === local) return value;
  return undefined;
}

/** Whether a block of <futureMetadata name="XLDAPR"> says its cells hold dynamic array formulas. */
function isDynamic(bk: XmlElement): boolean {
  for (const extLst of children(bk, 'extLst')) {
    for (const ext of children(extLst, 'ext')) {
      for (const props of children(ext, 'dynamicArrayProperties')) {
        const flag = attr(props, 'fDynamic');
        if (flag === '1' || flag === 'true') return true;
      }
    }
  }
  return false;
}

/** The values of `cm` (from 1) that mark a dynamic array formula. Empty for a part that cannot be read. */
export function dynamicArrayCells(bytes: Uint8Array | undefined): Set<number> {
  const found = new Set<number>();
  if (!bytes) return found;
  let root: XmlElement;
  try {
    root = parseXml(bytes).root;
  } catch (error) {
    return found;
  }
  if (localName(root.name) !== 'metadata') return found;
  // Types by their place in the list, from 1.
  const types: string[] = [];
  for (const list of children(root, 'metadataTypes')) for (const type of children(list, 'metadataType')) types.push(attr(type, 'name') || '');
  const dynamicBlocks = new Set<number>();
  for (const future of children(root, 'futureMetadata')) {
    if (attr(future, 'name') !== DYNAMIC_ARRAY_TYPE) continue;
    children(future, 'bk').forEach((bk, index) => {
      if (isDynamic(bk)) dynamicBlocks.add(index);
    });
    break;
  }
  if (dynamicBlocks.size === 0) return found;
  for (const list of children(root, 'cellMetadata')) {
    children(list, 'bk').forEach((bk, index) => {
      for (const rc of children(bk, 'rc')) {
        const t = Number(attr(rc, 't'));
        const v = Number(attr(rc, 'v'));
        if (t >= 1 && t <= types.length && types[t - 1] === DYNAMIC_ARRAY_TYPE && dynamicBlocks.has(v)) {
          found.add(index + 1);
          break;
        }
      }
    });
    break;
  }
  return found;
}

/** A new metadata part with one block that marks dynamic array formulas: `cm="1"`. */
export function newMetadataXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<metadata xmlns="${NS_MAIN}" xmlns:xda="${NS_DYNAMIC_ARRAY}">` +
    `<metadataTypes count="1"><metadataType name="${DYNAMIC_ARRAY_TYPE}" ${TYPE_ATTRS}/></metadataTypes>` +
    `<futureMetadata name="${DYNAMIC_ARRAY_TYPE}" count="1"><bk><extLst><ext uri="${DYNAMIC_ARRAY_URI}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
    '</metadata>'
  );
}

/** Offset of the start tag of the first child of the root with this local name, or -1. */
function tagAt(s: string, prefix: string, local: string, from: number): number {
  const needle = `<${prefix}${local}`;
  let at = s.indexOf(needle, from);
  while (at >= 0) {
    const next = s.charCodeAt(at + needle.length);
    if (next === 32 || next === 62 || next === 47 || next === 9 || next === 10 || next === 13) return at;
    at = s.indexOf(needle, at + 1);
  }
  return -1;
}

/**
 * The text with its comments and CDATA sections blanked out (as many spaces): what looks like an element in
 * them is not found there, and every offset is that of the text itself. Undefined for one that never ends.
 */
function masked(s: string): string | undefined {
  const parts: string[] = [];
  let from = 0;
  // The next of each, searched again only once passed: a part of many comments and no CDATA is read once, not once per comment.
  let comment = s.indexOf('<!--');
  let cdata = s.indexOf('<![CDATA[');
  for (;;) {
    if (comment >= 0 && comment < from) comment = s.indexOf('<!--', from);
    if (cdata >= 0 && cdata < from) cdata = s.indexOf('<![CDATA[', from);
    const at = comment < 0 ? cdata : cdata < 0 ? comment : Math.min(comment, cdata);
    if (at < 0) {
      parts.push(s.slice(from));
      return parts.join('');
    }
    const close = at === comment ? s.indexOf('-->', at + 4) : s.indexOf(']]>', at + 9);
    if (close < 0) return undefined;
    const end = close + 3;
    parts.push(s.slice(from, at), ' '.repeat(end - at));
    from = end;
  }
}

/** The text of a part and its masked twin (see masked()), changed together. */
interface Twin {
  s: string;
  m: string;
}

function splice(t: Twin, at: number, end: number, text: string): void {
  t.s = t.s.slice(0, at) + text + t.s.slice(end);
  t.m = t.m.slice(0, at) + text + t.m.slice(end);
}

/** Sets the count attribute of the start tag at `at` to `count`. */
function withCount(t: Twin, at: number, count: number): void {
  const end = t.m.indexOf('>', at);
  const tag = t.s.slice(at, end);
  const replaced = /\scount="[^"]*"/.test(tag) ? tag.replace(/\scount="[^"]*"/, ` count="${count}"`) : `${tag} count="${count}"`;
  splice(t, at, end, replaced);
}

/**
 * The value of `cm` that marks a dynamic array formula in a workbook whose metadata part is
 * `bytes` (undefined: it has none), with the part to write when it has to change. Undefined
 * when the part cannot be read or added to: the formula is then written without one.
 */
export function dynamicArrayMetadata(bytes: Uint8Array | undefined): { cm: number; xml?: string } | undefined {
  if (!bytes) return { cm: 1, xml: newMetadataXml() };
  const known = dynamicArrayCells(bytes);
  if (known.size > 0) {
    // The first of them; a part may list very many (no spread into an argument list).
    let first = Infinity;
    known.forEach((cm) => {
      if (cm < first) first = cm;
    });
    return { cm: first };
  }
  let s: string;
  let root: XmlElement;
  try {
    s = decodeXmlBytes(bytes);
    root = parseXml(bytes).root;
  } catch (error) {
    return undefined;
  }
  if (localName(root.name) !== 'metadata') return undefined;
  // The part is written back as UTF-8: one that says it is in another encoding is not changed.
  const declaration = /^﻿?<\?xml[^>]*\?>/.exec(s);
  if (declaration && /encoding\s*=\s*["'](?!utf-8["'])/i.test(declaration[0])) return undefined;
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
  // Elements are found by their text, in which comments and CDATA sections are blanked out: what looks like an
  // element there is not one. They stay in the part as they are.
  const m = masked(s);
  if (m === undefined || m.length !== s.length) return undefined;
  const t: Twin = { s, m };
  const prefix = root.name.slice(0, root.name.length - 'metadata'.length);
  const rootAt = tagAt(t.m, prefix, 'metadata', 0);
  if (rootAt < 0) return undefined;
  const rootEnd = t.m.indexOf('>', rootAt);
  if (rootEnd < 0 || t.m.charCodeAt(rootEnd - 1) === 47) return undefined;
  const types = children(root, 'metadataTypes');
  const typeCount = types.length > 0 ? children(types[0], 'metadataType').length : 0;
  const cells = children(root, 'cellMetadata');
  const cellCount = cells.length > 0 ? children(cells[0], 'bk').length : 0;
  if (typeCount >= MAX_BLOCKS || cellCount >= MAX_BLOCKS) return undefined;
  if (children(root, 'futureMetadata').some((future) => attr(future, 'name') === DYNAMIC_ARRAY_TYPE)) return undefined;
  const type = typeCount + 1;
  const xda = `xmlns:xda="${NS_DYNAMIC_ARRAY}"`;
  const future =
    `<${prefix}futureMetadata name="${DYNAMIC_ARRAY_TYPE}" count="1"><${prefix}bk><${prefix}extLst><${prefix}ext uri="${DYNAMIC_ARRAY_URI}">` +
    `<xda:dynamicArrayProperties ${xda} fDynamic="1" fCollapsed="0"/></${prefix}ext></${prefix}extLst></${prefix}bk></${prefix}futureMetadata>`;
  const block = `<${prefix}bk><${prefix}rc t="${type}" v="0"/></${prefix}bk>`;
  // The type goes last in its list, so that every record of the part keeps pointing at its own.
  const typeXml = `<${prefix}metadataType name="${DYNAMIC_ARRAY_TYPE}" ${TYPE_ATTRS}/>`;
  if (types.length > 0) {
    const at = tagAt(t.m, prefix, 'metadataTypes', rootEnd);
    if (at < 0) return undefined;
    const close = t.m.indexOf(`</${prefix}metadataTypes>`, at);
    const tagEnd = t.m.indexOf('>', at);
    if (t.m.charCodeAt(tagEnd - 1) === 47) splice(t, tagEnd - 1, tagEnd + 1, `>${typeXml}</${prefix}metadataTypes>`);
    else if (close < 0) return undefined;
    else splice(t, close, close, typeXml);
    withCount(t, at, type);
  } else {
    splice(t, rootEnd + 1, rootEnd + 1, `<${prefix}metadataTypes count="1">${typeXml}</${prefix}metadataTypes>`);
  }
  // <futureMetadata> comes before <cellMetadata>, <valueMetadata> and <extLst>.
  let cellsAt = tagAt(t.m, prefix, 'cellMetadata', 0);
  let place = cellsAt;
  if (place < 0) place = tagAt(t.m, prefix, 'valueMetadata', 0);
  // The <extLst> of the part is its last child: blocks of other types hold ones of their own before it.
  if (place < 0 && children(root, 'extLst').length > 0) place = t.m.lastIndexOf(`<${prefix}extLst`);
  if (place < 0) place = t.m.lastIndexOf(`</${prefix}metadata>`);
  if (place < 0) return undefined;
  splice(t, place, place, future);
  if (cellsAt >= 0) {
    cellsAt = tagAt(t.m, prefix, 'cellMetadata', place + future.length);
    if (cellsAt < 0) return undefined;
    const tagEnd = t.m.indexOf('>', cellsAt);
    if (t.m.charCodeAt(tagEnd - 1) === 47) splice(t, tagEnd - 1, tagEnd + 1, `>${block}</${prefix}cellMetadata>`);
    else {
      const close = t.m.indexOf(`</${prefix}cellMetadata>`, cellsAt);
      if (close < 0) return undefined;
      splice(t, close, close, block);
    }
    withCount(t, cellsAt, cellCount + 1);
  } else {
    const after = place + future.length;
    splice(t, after, after, `<${prefix}cellMetadata count="1">${block}</${prefix}cellMetadata>`);
  }
  // What came out has to say what was meant, or nothing is written.
  const cm = cellCount + 1;
  const check = dynamicArrayCells(new TextEncoder().encode(t.s));
  return check.has(cm) ? { cm, xml: t.s } : undefined;
}
