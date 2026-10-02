// Helpers shared by the docx writer tests, with tests of their own: a
// canonical form of XML to compare two spellings of the same content, a
// picture of a document model that leaves out where things were read from,
// and a checker for the order of the children of property elements.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LVL_ORDER, STYLE_ORDER } from '../src/formats/docx/write-package';
import { PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TBLPR_ORDER, TCPR_ORDER, TRPR_ORDER } from '../src/formats/docx/write-props';
import type { DocModel } from '../src/model/doc/types';
import { parseXml, type XmlElement } from '../src/xml';

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * The element as text that does not depend on how it was spelled: attributes
 * sorted by name, white space between elements dropped, empty elements closed
 * in one way.
 */
export function canonical(el: XmlElement): string {
  const out: string[] = [];
  const walk = (node: XmlElement) => {
    out.push('<', node.name);
    for (const name of Array.from(node.attrs.keys()).sort()) out.push(' ', name, '="', escape(node.attrs.get(name) as string), '"');
    const mixed = node.children.some((child) => child.type === 'element');
    const kept = node.children.filter((child) => !(child.type === 'text' && mixed && child.value.trim() === ''));
    if (!kept.length) {
      out.push('/>');
      return;
    }
    out.push('>');
    for (const child of kept) {
      if (child.type === 'element') walk(child);
      else if (child.type === 'text') out.push(escape(child.value));
      else if (child.type === 'comment') out.push('<!--', child.value, '-->');
      else out.push('<?', child.target, ' ', child.value, '?>');
    }
    out.push('</', node.name, '>');
  };
  walk(el);
  return out.join('');
}

export const canonicalText = (xml: string | Uint8Array): string => canonical(parseXml(xml).root);

const SKIPPED = new Set(['srcStart', 'srcEnd', 'sectionSpan', 'modified']);

/**
 * The model as plain data, without what only says where things were read
 * from: source spans, the source text of the parts, the modified flag.
 * Wrappers are numbered in the order they are met, so that two models with
 * the same structure give the same picture. Without `shared`, every item
 * shows its wrappers in full instead: two items in one wrapper then look like
 * two items in two equal wrappers, which is what a wrapper that an edit split
 * over two paragraphs becomes in the file.
 */
export function modelPicture(value: unknown, shared = true): unknown {
  const wrappers = new Map<object, number>();
  const walk = (item: unknown, key: string, depth: number): unknown => {
    if (depth > 2000) throw new Error('modelPicture: too deep');
    if (item === null || typeof item !== 'object') return item;
    if (item instanceof Uint8Array) return { bytes: item.length };
    if (Array.isArray(item)) return item.map((entry) => walk(entry, '', depth + 1));
    if (item instanceof Map) return { map: Array.from(item.entries()).map(([k, v]) => [k, walk(v, '', depth + 1)]) };
    if (item instanceof Set) return { set: Array.from(item.values()) };
    if (shared && (key === 'wrap' || key === 'parent')) {
      const seen = wrappers.get(item);
      if (seen !== undefined) return { wrapper: seen };
      wrappers.set(item, wrappers.size);
    }
    const record = item as Record<string, unknown>;
    const story = 'rootName' in record;
    const out: Record<string, unknown> = {};
    for (const name of Object.keys(record).sort()) {
      if (SKIPPED.has(name) || (story && name === 'source')) continue;
      const inner = record[name];
      if (inner !== undefined) out[name] = walk(inner, name, depth + 1);
    }
    return out;
  };
  return walk(value, '', 0);
}

/** Where two pictures first differ, as a path of keys; undefined when they are equal. */
export function firstDifference(a: unknown, b: unknown, path = ''): string | undefined {
  if (a === b) return undefined;
  if (typeof a !== typeof b || a === null || b === null || typeof a !== 'object') return path || '.';
  if (Array.isArray(a) !== Array.isArray(b)) return path || '.';
  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) return `${path}.length(${a.length}/${b.length})`;
    for (let i = 0; i < a.length; i++) {
      const found = firstDifference(a[i], b[i], `${path}[${i}]`);
      if (found) return found;
    }
    return undefined;
  }
  const x = a as Record<string, unknown>;
  const y = b as Record<string, unknown>;
  const keys = Array.from(new Set(Object.keys(x).concat(Object.keys(y)))).sort();
  for (const key of keys) {
    const found = firstDifference(x[key], y[key], `${path}.${key}`);
    if (found) return found;
  }
  return undefined;
}

/** Path of a difference without the indexes, to count kinds of differences. */
export const differenceKind = (path: string): string => path.replace(/\[\d+\]/g, '[]').replace(/\(\d+\/\d+\)/g, '');

export function documentPicture(doc: DocModel, shared = true): unknown {
  return modelPicture(
    {
      body: doc.body,
      headers: doc.headers,
      footers: doc.footers,
      footnotes: doc.footnotes,
      endnotes: doc.endnotes,
      comments: doc.comments,
      styles: doc.styles,
      numbering: doc.numbering,
      settings: doc.settings,
      fonts: doc.fonts,
      strict: doc.strict,
    },
    shared,
  );
}

type Plain = Record<string, unknown>;
const isPlain = (value: unknown): value is Plain => value !== null && typeof value === 'object' && !Array.isArray(value);

/**
 * The picture of a model without what a package made from it necessarily
 * states differently, to compare a model read from a doc or RTF file with the
 * docx built from it:
 *   - the names of parts, root elements and relationships;
 *   - the ids of headers and footers (numbered in their order instead);
 *   - of a drawing: its tree, the name the file has to give it, the position
 *     of its picture inside it, and the anchor settings the schema requires
 *     and the source left open;
 *   - which spelling a table look was read from.
 */
export function convertedPicture(doc: DocModel, shared = true): unknown {
  const picture = documentPicture(doc, shared) as Plain;
  const ids = new Map<string, string>();
  for (const list of ['headers', 'footers']) {
    const map = (picture[list] as Plain).map as [string, unknown][];
    map.forEach((entry, i) => {
      ids.set(entry[0], `${list}#${i}`);
      entry[0] = `${list}#${i}`;
    });
  }
  const fileName = (ref: unknown): string | undefined => (isPlain(ref) && typeof ref.target === 'string' ? ref.target.slice(ref.target.lastIndexOf('/') + 1) : undefined);
  const walk = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(walk);
    if (!isPlain(value)) return value;
    if (value.type === 'drawing') {
      const anchor = value.anchor as Plain | undefined;
      const transform = value.transform as Plain | undefined;
      const turned = transform ? [transform.rotation || 0, !!transform.flipH, !!transform.flipV] : [0, false, false];
      const out: Plain = {
        type: 'drawing',
        placement: value.placement,
        kind: value.kind,
        width: value.width,
        height: value.height,
        description: value.description,
        image: fileName(value.image),
        images: Array.isArray(value.images) ? value.images.map((item) => [fileName((item as Plain).image), (item as Plain).crop]) : undefined,
        crop: value.crop,
        // Distances of zero say what no distance says.
        inline: isPlain(value.inline) && Object.keys(value.inline).some((side) => (value.inline as Plain)[side] !== 0) ? value.inline : undefined,
        effectExtent: value.effectExtent,
        turned,
        textBoxes: walk(value.textBoxes),
      };
      if (anchor) {
        const sided = anchor.wrap === 'square' || anchor.wrap === 'tight' || anchor.wrap === 'through';
        out.anchor = {
          // A position is required in the file; one the source left open is written as the default.
          horizontal: anchor.horizontal || { relativeFrom: 'column', offset: 0 },
          vertical: anchor.vertical || { relativeFrom: 'paragraph', offset: 0 },
          wrap: anchor.wrap,
          wrapText: sided ? anchor.wrapText || 'bothSides' : undefined,
          behindDoc: !!anchor.behindDoc,
          locked: !!anchor.locked,
          distT: anchor.distT,
          distB: anchor.distB,
          distL: anchor.distL,
          distR: anchor.distR,
        };
      }
      return JSON.parse(JSON.stringify(out));
    }
    const out: Plain = {};
    const story = 'rootName' in value;
    for (const name of Object.keys(value)) {
      if (story && (name === 'part' || name === 'rootName' || name === 'rootAttrs' || name === 'relationships' || name === 'prolog' || name === 'epilog')) continue;
      const inner = value[name];
      if ((name === 'headers' || name === 'footers') && isPlain(inner) && !('map' in inner)) {
        const refs: Plain = {};
        for (const kind of Object.keys(inner)) refs[kind] = ids.get(inner[kind] as string) || 'missing';
        out[name] = refs;
      } else if (name === 'look' && isPlain(inner)) {
        const look: Plain = { ...inner };
        delete look.val;
        delete look.stated;
        out[name] = look;
      } else out[name] = walk(inner);
    }
    return out;
  };
  return walk(picture);
}

// Left and start share a place, as do right and end.
const EDGES = new Map<string, number>([
  ['top', 0],
  ['left', 1],
  ['start', 1],
  ['bottom', 2],
  ['right', 3],
  ['end', 3],
  ['insideH', 4],
  ['insideV', 5],
  ['tl2br', 6],
  ['tr2bl', 7],
  ['between', 8],
  ['bar', 9],
]);

const first = (...names: string[]): Map<string, number> => new Map(names.map((name, i) => [name, i] as [string, number]));

/** Order of the children of the elements whose order the writer is responsible for, by local name. */
const ORDERS = new Map<string, Map<string, number>>([
  ['rPr', RPR_ORDER],
  ['pPr', PPR_ORDER],
  ['tblPr', TBLPR_ORDER],
  ['tblPrEx', TBLPR_ORDER],
  ['trPr', TRPR_ORDER],
  ['tcPr', TCPR_ORDER],
  ['sectPr', SECTPR_ORDER],
  ['lvl', LVL_ORDER],
  ['style', STYLE_ORDER],
  ['pBdr', EDGES],
  ['tblBorders', EDGES],
  ['tcBorders', EDGES],
  ['tcMar', EDGES],
  ['tblCellMar', EDGES],
  ['pgBorders', EDGES],
  ['numPr', first('ilvl', 'numId')],
  ['tbl', first('tblPr', 'tblGrid')],
  ['tr', first('tblPrEx', 'trPr')],
]);

/** Elements that have to be the first child of their parent. */
const LEADING = new Map<string, string>([
  ['p', 'pPr'],
  ['r', 'rPr'],
  ['tc', 'tcPr'],
]);

/**
 * Checks the order of the children of property elements against the sequence
 * of their complex type. Returns one line for each kind of violation
 * ("pPr: jc before spacing"); empty when the order is right. Only elements of
 * the WordprocessingML prefix are looked at; children the schema tables do
 * not know are skipped.
 */
export function schemaOrderViolations(xml: string | Uint8Array, prefix = 'w:'): string[] {
  const found = new Set<string>();
  const local = (name: string): string | undefined => (name.slice(0, prefix.length) === prefix ? name.slice(prefix.length) : undefined);
  const walk = (el: XmlElement, depth: number) => {
    if (depth > 400) return;
    const name = local(el.name);
    const order = name === undefined ? undefined : ORDERS.get(name);
    const leading = name === undefined ? undefined : LEADING.get(name);
    let last = -1;
    let lastName = '';
    let position = 0;
    for (const child of el.children) {
      if (child.type !== 'element') continue;
      const childName = local(child.name);
      if (leading !== undefined && childName === leading && position > 0) found.add(`${name}: ${leading} is not first`);
      position++;
      if (order && childName !== undefined) {
        const rank = order.get(childName);
        if (rank !== undefined) {
          if (rank < last) found.add(`${name}: ${childName} after ${lastName}`);
          else {
            last = rank;
            lastName = childName;
          }
        }
      }
      walk(child, depth + 1);
    }
  };
  walk(parseXml(xml).root, 0);
  return Array.from(found).sort();
}

// The tests of the helpers themselves run once, not again in every file that imports them.
if ((process.argv[1] || '').slice(-'docx-write-helpers.test.ts'.length) === 'docx-write-helpers.test.ts') {
  test('canonical form ignores attribute order, quoting and white space between elements', () => {
    const a = parseXml('<a y="2" x=\'1\'>\n  <b/>\n  <c> text </c>\n</a>').root;
    const b = parseXml('<a x="1" y="2"><b></b><c> text </c></a>').root;
    assert.equal(canonical(a), canonical(b));
    assert.notEqual(canonical(a), canonical(parseXml('<a x="1" y="2"><b/><c>text</c></a>').root));
  });

  test('schema order checker finds a misplaced child and accepts the right order', () => {
    const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    assert.deepEqual(schemaOrderViolations(`<w:p ${ns}><w:pPr><w:pStyle w:val="a"/><w:spacing/><w:jc w:val="left"/></w:pPr></w:p>`), []);
    assert.deepEqual(schemaOrderViolations(`<w:p ${ns}><w:pPr><w:jc w:val="left"/><w:spacing/></w:pPr></w:p>`), ['pPr: spacing after jc']);
    assert.deepEqual(schemaOrderViolations(`<w:p ${ns}><w:r><w:t>x</w:t><w:rPr/></w:r></w:p>`), ['r: rPr is not first']);
    assert.deepEqual(schemaOrderViolations(`<w:tcBorders ${ns}><w:left/><w:top/></w:tcBorders>`), ['tcBorders: top after left']);
  });

  test('model picture leaves out source spans and numbers shared wrappers', () => {
    const wrapper = { kind: 'ins', srcStart: 3 };
    const a = modelPicture({ rootName: 'w:document', source: 'text', blocks: [{ type: 'empty', wrap: wrapper, srcStart: 1 }, { type: 'empty', wrap: wrapper }] });
    assert.deepEqual(a, { rootName: 'w:document', blocks: [{ type: 'empty', wrap: { kind: 'ins' } }, { type: 'empty', wrap: { wrapper: 0 } }] });
    assert.equal(firstDifference(a, a), undefined);
    assert.equal(differenceKind(firstDifference({ a: [1, { b: 2 }] }, { a: [1, { b: 3 }] }) as string), '.a[].b');
  });
}
