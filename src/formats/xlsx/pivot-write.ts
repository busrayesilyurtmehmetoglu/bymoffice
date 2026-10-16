// Writes the parts of refreshed pivot tables: the table definitions, the cache definitions and
// the records of the caches.
//
// A part is written from the part as the file has it: what a refresh changes is put in its place
// (where the table stands, the items of its fields, its lines of rows and columns, the item a
// page field shows; the items of the cache fields and how many records there are; the records),
// and every other element and attribute stays as it was. A table or cache that is as the file
// has it (never refreshed, or the refresh was undone) is not written; one that an earlier save
// wrote goes back to the bytes of the file.

import { serialToDate } from '../../format/number-format';
import type { PivotCache, PivotCacheState, PivotItem, PivotLine, PivotTable, PivotTableDef, PivotTableState, PivotValue } from '../../model/sheet/pivot';
import { rangeName, type Workbook } from '../../model/sheet/types';
import { childElements, decodeXmlBytes, localName, parseXml, type XmlDocument, type XmlElement, type XmlNode } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { encodeXstring, escapeAttr } from './write';

const encoder = new TextEncoder();

const prefixOf = (el: XmlElement) => el.name.slice(0, el.name.length - localName(el.name).length);
const child = (parent: XmlElement, local: string): XmlElement | undefined => childElements(parent, local)[0];
const element = (name: string, attrs: [string, string][] = [], children: XmlNode[] = []): XmlElement => ({ type: 'element', name, attrs: new Map(attrs), children });

/** An attribute of ST_Xstring: "_xHHHH_" escapes, then XML escapes. */
const xstringAttr = (text: string) => escapeAttr(encodeXstring(text));

/** A date of a cache as the file writes it: "2024-03-01T00:00:00". */
export function isoOfSerial(serial: number, date1904: boolean): string {
  const d = serialToDate(serial, date1904);
  if (!isFinite(d.year)) return '1899-12-31T00:00:00';
  const two = (n: number) => (n < 10 ? `0${n}` : String(n));
  const year = String(d.year);
  return `${'0000'.slice(year.length)}${year}-${two(d.month)}-${two(d.day)}T${two(d.hour)}:${two(d.minute)}:${two(d.second)}`;
}

/** A number as the file writes it. */
const numberText = (n: number) => String(n);

// ---------------------------------------------------------------------------------------------
// Table definitions

/** Puts `el` among the children of `parent` after the last of those named by `before` (in their order of the schema), or first. */
function placeAfter(parent: XmlElement, el: XmlElement, before: string[]): void {
  let at = -1;
  parent.children.forEach((node, i) => {
    if (node.type === 'element' && before.indexOf(localName(node.name)) >= 0) at = i;
  });
  parent.children.splice(at + 1, 0, el);
}

/** Replaces the child named `local`, or puts the new one after the children named by `before`. */
function replaceChild(parent: XmlElement, local: string, el: XmlElement, before: string[]): void {
  const at = parent.children.findIndex((node) => node.type === 'element' && localName(node.name) === local);
  if (at >= 0) parent.children[at] = el;
  else placeAfter(parent, el, before);
}

function linesElement(name: string, lines: PivotLine[], prefix: string): XmlElement {
  const kids: XmlNode[] = lines.map((line) => {
    const attrs: [string, string][] = [];
    if (line.t !== undefined) attrs.push(['t', line.t]);
    if (line.r > 0) attrs.push(['r', String(line.r)]);
    if (line.i > 0) attrs.push(['i', String(line.i)]);
    return element(`${prefix}i`, attrs, line.x.map((v) => element(`${prefix}x`, v === 0 ? [] : [['v', String(v)]])));
  });
  return element(name, [['count', String(lines.length)]], kids);
}

function itemsElement(items: PivotItem[], prefix: string, original: XmlElement | undefined): XmlElement {
  const attrs: [string, string][] = original ? [...original.attrs].filter(([name]) => name !== 'count') : [];
  attrs.unshift(['count', String(items.length)]);
  return element(`${prefix}items`, attrs, items.map((item) => element(`${prefix}item`, item.attrs)));
}

/** The children of <pivotTableDefinition> in the order of the schema, up to those the fields are listed in. */
const TABLE_ORDER = ['location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields'];

/** Sets an attribute to a value, or takes it away for the value that is its default. */
function setOrDrop(el: XmlElement, name: string, value: string | undefined): void {
  if (value === undefined) el.attrs.delete(name);
  else el.attrs.set(name, value);
}

/** Puts a list of fields (<rowFields>, <colFields>) in the place of the one the part has; none for an empty list. */
function fieldList(root: XmlElement, local: string, list: number[], p: string): void {
  const at = root.children.findIndex((node) => node.type === 'element' && localName(node.name) === local);
  if (list.length === 0) {
    if (at >= 0) root.children.splice(at, 1);
    return;
  }
  const made = element(`${p}${local}`, [['count', String(list.length)]], list.map((x) => element(`${p}field`, [['x', String(x)]])));
  replaceChild(root, local, made, TABLE_ORDER.slice(0, TABLE_ORDER.indexOf(local)));
}

/**
 * Puts the fields of a table as they were changed (PivotTableDef) into its definition: where each
 * field stands and how it is totalled and sorted, the lists of fields of rows, columns, page and
 * values, the grand totals. Elements of the part that stand for the same thing (a field of the
 * page, a value of the same field) keep their other attributes and their children.
 */
function writeDef(root: XmlElement, def: PivotTableDef, location: XmlElement, p: string): void {
  setOrDrop(root, 'rowGrandTotals', def.rowGrandTotals ? undefined : '0');
  setOrDrop(root, 'colGrandTotals', def.colGrandTotals ? undefined : '0');
  // The values stand last in the columns (see the refresh): no other place is stated.
  root.attrs.delete('dataPosition');
  setOrDrop(location, 'rowPageCount', def.pages ? String(def.pages.rows) : undefined);
  setOrDrop(location, 'colPageCount', def.pages ? String(def.pages.cols) : undefined);
  const fieldsEl = child(root, 'pivotFields');
  const list = fieldsEl ? childElements(fieldsEl, 'pivotField') : [];
  const data = new Set(def.dataFields.map((value) => value.fld));
  def.fields.forEach((field, f) => {
    const el = list[f];
    if (!el) return;
    setOrDrop(el, 'axis', field.axis);
    setOrDrop(el, 'dataField', data.has(f) ? '1' : undefined);
    setOrDrop(el, 'defaultSubtotal', field.defaultSubtotal ? undefined : '0');
    setOrDrop(el, 'sortType', field.sortType === 'manual' ? undefined : field.sortType);
    setOrDrop(el, 'compact', field.compact ? undefined : '0');
    setOrDrop(el, 'outline', field.outline ? undefined : '0');
  });
  fieldList(root, 'rowFields', def.rowFields, p);
  fieldList(root, 'colFields', def.colFields, p);
  // The fields of the page: the element of the part for the same field keeps what else it says (hier, name, cap).
  const pagesWas = child(root, 'pageFields');
  const oldPages = pagesWas ? childElements(pagesWas, 'pageField') : [];
  if (def.pageFields.length === 0) {
    if (pagesWas) root.children.splice(root.children.indexOf(pagesWas), 1);
  } else {
    const kids = def.pageFields.map((page) => {
      const had = oldPages.find((el) => el.attrs.get('fld') === String(page.fld));
      const el = had ? element(had.name, [...had.attrs], had.children) : element(`${p}pageField`, [['fld', String(page.fld)], ['hier', '-1']]);
      return el;
    });
    replaceChild(root, 'pageFields', element(`${p}pageFields`, [['count', String(kids.length)]], kids), TABLE_ORDER.slice(0, TABLE_ORDER.indexOf('pageFields')));
  }
  // The values: the element of the part for a value of the same field keeps what else it says (base field and item, extensions).
  const dataWas = child(root, 'dataFields');
  const oldData = dataWas ? childElements(dataWas, 'dataField') : [];
  const used = new Set<XmlElement>();
  const values = def.dataFields.map((value) => {
    const had = oldData.find((el) => !used.has(el) && el.attrs.get('fld') === String(value.fld));
    let el: XmlElement;
    if (had) {
      used.add(had);
      el = element(had.name, [...had.attrs], had.children);
    } else el = element(`${p}dataField`, [['name', ''], ['fld', ''], ['baseField', '0'], ['baseItem', '0']]);
    // A name of ST_Xstring, as the reader takes it.
    el.attrs.set('name', encodeXstring(value.name));
    el.attrs.set('fld', String(value.fld));
    setOrDrop(el, 'subtotal', value.subtotal === 'sum' ? undefined : value.subtotal);
    setOrDrop(el, 'numFmtId', value.numFmtId === undefined ? undefined : String(value.numFmtId));
    return el;
  });
  replaceChild(root, 'dataFields', element(`${p}dataFields`, [['count', String(values.length)]], values), TABLE_ORDER.slice(0, TABLE_ORDER.indexOf('dataFields')));
}

/** The table definition with what the state says, from the part as the file has it. */
export function writeTableDefinition(original: Uint8Array, table: PivotTable, state: PivotTableState): XmlDocument {
  const doc = parseXml(original);
  const root = doc.root;
  if (localName(root.name) !== 'pivotTableDefinition') throw new Error(`Cannot write the pivot table "${table.name}": its part is not a table definition`);
  const p = prefixOf(root);
  const location = child(root, 'location');
  if (!location) throw new Error(`Cannot write the pivot table "${table.name}": its part does not say where it stands`);
  if (state.def) writeDef(root, state.def, location, p);
  location.attrs.set('ref', rangeName(state.location.ref));
  location.attrs.set('firstHeaderRow', String(state.location.firstHeaderRow));
  location.attrs.set('firstDataRow', String(state.location.firstDataRow));
  location.attrs.set('firstDataCol', String(state.location.firstDataCol));
  const fields = child(root, 'pivotFields');
  const list = fields ? childElements(fields, 'pivotField') : [];
  state.items.forEach((items, f) => {
    const field = list[f];
    if (!items || !field) return;
    const was = child(field, 'items');
    const made = itemsElement(items, p, was);
    if (was) field.children[field.children.indexOf(was)] = made;
    else field.children.unshift(made);
  });
  replaceChild(root, 'rowItems', linesElement(`${p}rowItems`, state.rowItems, p), TABLE_ORDER.slice(0, 3));
  replaceChild(root, 'colItems', linesElement(`${p}colItems`, state.colItems, p), TABLE_ORDER.slice(0, 5));
  const pages = child(root, 'pageFields');
  if (pages) {
    childElements(pages, 'pageField').forEach((page, i) => {
      const item = state.pageItems[i];
      if (item === undefined) page.attrs.delete('item');
      else page.attrs.set('item', String(item));
    });
  }
  return doc;
}

// ---------------------------------------------------------------------------------------------
// Cache definitions

/** What the values of a field hold, for the attributes of <sharedItems>. */
interface FieldStats {
  s: boolean;
  n: boolean;
  d: boolean;
  b: boolean;
  e: boolean;
  m: boolean;
  integers: boolean;
  long: boolean;
  min: number;
  max: number;
  minDate: number;
  maxDate: number;
}

function statsOf(values: Iterable<PivotValue>): FieldStats {
  const out: FieldStats = { s: false, n: false, d: false, b: false, e: false, m: false, integers: true, long: false, min: Infinity, max: -Infinity, minDate: Infinity, maxDate: -Infinity };
  for (const value of values) {
    switch (value.t) {
      case 's':
        out.s = true;
        if (value.v.length > 255) out.long = true;
        break;
      case 'n':
        out.n = true;
        if (Math.floor(value.v) !== value.v) out.integers = false;
        if (value.v < out.min) out.min = value.v;
        if (value.v > out.max) out.max = value.v;
        break;
      case 'd':
        out.d = true;
        if (value.v < out.minDate) out.minDate = value.v;
        if (value.v > out.maxDate) out.maxDate = value.v;
        break;
      case 'b':
        out.b = true;
        break;
      case 'e':
        out.e = true;
        break;
      default:
        out.m = true;
    }
  }
  return out;
}

/** The attributes of <sharedItems> the writer states; any other one of the file is kept. */
const SHARED_ITEMS_ATTRS = ['containsSemiMixedTypes', 'containsNonDate', 'containsDate', 'containsString', 'containsBlank', 'containsMixedTypes', 'containsNumber', 'containsInteger', 'minValue', 'maxValue', 'minDate', 'maxDate', 'count', 'longText'];

/** The attributes Excel writes on <sharedItems> for values like these: only those that differ from the defaults of the schema. */
function sharedItemsAttrs(stats: FieldStats, count: number | undefined, date1904: boolean): [string, string][] {
  const out: [string, string][] = [];
  const kinds = [stats.s, stats.n, stats.d, stats.b, stats.e].filter(Boolean).length;
  if (!(stats.s || stats.m || stats.b || stats.e)) out.push(['containsSemiMixedTypes', '0']);
  if (stats.d && !(stats.s || stats.n || stats.b || stats.e)) out.push(['containsNonDate', '0']);
  if (stats.d) out.push(['containsDate', '1']);
  if (!stats.s) out.push(['containsString', '0']);
  if (stats.m) out.push(['containsBlank', '1']);
  if (kinds > 1) out.push(['containsMixedTypes', '1']);
  if (stats.n) out.push(['containsNumber', '1']);
  if (stats.n && stats.integers) out.push(['containsInteger', '1']);
  if (stats.n) {
    out.push(['minValue', numberText(stats.min)]);
    out.push(['maxValue', numberText(stats.max)]);
  }
  if (stats.d) {
    out.push(['minDate', isoOfSerial(stats.minDate, date1904)]);
    out.push(['maxDate', isoOfSerial(stats.maxDate, date1904)]);
  }
  if (count !== undefined) out.push(['count', String(count)]);
  if (stats.long) out.push(['longText', '1']);
  return out;
}

/** A value as an element of the cache (<s>, <n>, ...), as text. */
function valueXml(value: PivotValue, prefix: string, date1904: boolean): string {
  switch (value.t) {
    case 's':
      return `<${prefix}s v="${xstringAttr(value.v)}"/>`;
    case 'n':
      return `<${prefix}n v="${numberText(value.v)}"/>`;
    case 'd':
      return `<${prefix}d v="${isoOfSerial(value.v, date1904)}"/>`;
    case 'b':
      return `<${prefix}b v="${value.v ? 1 : 0}"/>`;
    case 'e':
      return `<${prefix}e v="${xstringAttr(value.v)}"/>`;
    default:
      return `<${prefix}m/>`;
  }
}

/** A value as an element of the cache (<s>, <n>, ...). The attributes of a tree are written escaped. */
function valueElement(value: PivotValue, prefix: string, date1904: boolean): XmlElement {
  switch (value.t) {
    case 's':
      return element(`${prefix}s`, [['v', encodeXstring(value.v)]]);
    case 'n':
      return element(`${prefix}n`, [['v', numberText(value.v)]]);
    case 'd':
      return element(`${prefix}d`, [['v', isoOfSerial(value.v, date1904)]]);
    case 'b':
      return element(`${prefix}b`, [['v', value.v ? '1' : '0']]);
    case 'e':
      return element(`${prefix}e`, [['v', encodeXstring(value.v)]]);
    default:
      return element(`${prefix}m`);
  }
}

/** The cache definition with what the state says, from the part as the file has it. */
export function writeCacheDefinition(original: Uint8Array, cache: PivotCache, state: PivotCacheState, date1904: boolean): XmlDocument {
  const doc = parseXml(original);
  const root = doc.root;
  if (localName(root.name) !== 'pivotCacheDefinition') throw new Error('Cannot write a pivot cache: its part is not a cache definition');
  // Where its source stands now, when rows or columns moved it.
  if (cache.sourceRef !== undefined) sourceElement(root).attrs.set('ref', cache.sourceRef);
  const p = prefixOf(root);
  if (cache.recordsPart !== undefined || root.attrs.has('recordCount')) root.attrs.set('recordCount', String(state.records.length));
  const fieldsEl = child(root, 'cacheFields');
  const list = fieldsEl ? childElements(fieldsEl, 'cacheField') : [];
  state.fields.forEach((field, f) => {
    const el = list[f];
    if (!el) return;
    let stats: FieldStats;
    if (field.items) stats = statsOf(field.items);
    else {
      const values: PivotValue[] = [];
      for (const record of state.records) {
        const value = record[f];
        if (typeof value !== 'number') values.push(value);
      }
      stats = statsOf(values);
    }
    const was = child(el, 'sharedItems');
    const attrs = sharedItemsAttrs(stats, field.items ? field.items.length : undefined, date1904);
    if (was) for (const [name, value] of was.attrs) if (SHARED_ITEMS_ATTRS.indexOf(name) < 0) attrs.push([name, value]);
    const kids: XmlNode[] = field.items ? field.items.map((value) => valueElement(value, p, date1904)) : [];
    const made = element(`${p}sharedItems`, attrs, kids);
    if (was) el.children[el.children.indexOf(was)] = made;
    else el.children.unshift(made);
  });
  return doc;
}

// ---------------------------------------------------------------------------------------------
// Records

/** The records of a cache, from the part as the file has it: its root keeps its attributes (the count aside), and its extensions stay. */
export function writeCacheRecords(original: Uint8Array, state: PivotCacheState, date1904: boolean): Uint8Array {
  const scanner = new XmlScanner(original);
  let root: { name: string; attrs: [string, string][] } | undefined;
  const kept: string[] = [];
  while (scanner.next()) {
    if (scanner.kind !== 'open') continue;
    if (!root) {
      root = { name: scanner.name, attrs: scanner.attrs() };
      if (localName(root.name) !== 'pivotCacheRecords') throw new Error('Cannot write a pivot cache: the part of its records is not one');
      if (scanner.selfClosing) break;
      continue;
    }
    // A child of the root: the records are made anew, everything else is kept as it is.
    if (scanner.local === 'r') scanner.skip();
    else kept.push(scanner.readRaw());
  }
  if (!root) throw new Error('Cannot write a pivot cache: the part of its records is empty');
  const name = root.name;
  const prefix = prefixOf({ type: 'element', name, attrs: new Map(), children: [] });
  const attrs = root.attrs.filter(([attr]) => attr !== 'count');
  attrs.push(['count', String(state.records.length)]);
  const head = decodeXmlBytes(original).match(/^﻿?(<\?xml[^?]*\?>)/);
  const out: string[] = [head ? `${head[1].replace(/encoding\s*=\s*(["'])[^"']*\1/, 'encoding="UTF-8"')}\n` : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'];
  out.push(`<${name}${attrs.map(([attr, value]) => ` ${attr}="${escapeAttr(value)}"`).join('')}>`);
  for (const record of state.records) {
    let line = `<${prefix}r>`;
    for (let f = 0; f < record.length; f++) {
      const value = record[f];
      if (typeof value === 'number') line += value === 0 ? `<${prefix}x/>` : `<${prefix}x v="${value}"/>`;
      else line += valueXml(value, prefix, date1904);
    }
    out.push(`${line}</${prefix}r>`);
  }
  for (const raw of kept) out.push(raw);
  out.push(`</${name}>`);
  return encoder.encode(out.join(''));
}

// ---------------------------------------------------------------------------------------------
// Places that moved
//
// Rows and columns inserted or deleted move the range of a source and where a table stands (see
// src/ops/pivot-places.ts). A part that was not refreshed keeps every byte but those of the value
// of that one attribute; a part whose text cannot be changed in place that way (not UTF-8, or the
// attribute not found once) is written out from its tree, with the attribute changed.

/** The element of a cache definition that states the range of its source (cacheSource/worksheetSource). */
function sourceElement(root: XmlElement): XmlElement {
  const source = child(root, 'cacheSource');
  const sheet = source ? child(source, 'worksheetSource') : undefined;
  if (!sheet || !sheet.attrs.has('ref')) throw new Error('Cannot write a pivot cache: its part does not state the range of its source');
  return sheet;
}

/** The element at a path of local names from the root (the first of each name, as the reader takes it). */
function elementAt(root: XmlElement, path: string[]): XmlElement | undefined {
  if (localName(root.name) !== path[0]) return undefined;
  let el: XmlElement | undefined = root;
  for (let i = 1; i < path.length && el; i++) el = child(el, path[i]);
  return el;
}

/** The text of a part that is UTF-8 and says nothing else; undefined for any other. */
function utf8Text(bytes: Uint8Array): string | undefined {
  if (bytes.length >= 2 && ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))) return undefined;
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes);
  } catch (error) {
    return undefined;
  }
  const declared = /^﻿?<\?xml\b[^?>]*?\bencoding\s*=\s*["']([^"']*)["']/.exec(text);
  return declared && !/^utf-?8$/i.test(declared[1]) ? undefined : text;
}

const isNameEnd = (c: number) => c === 32 || c === 9 || c === 10 || c === 13 || c === 47 || c === 62 || c === 61;
const isBlank = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * Where the value of an attribute (by its name as written, without a prefix) of the element at a
 * path of local names stands in the text of a well-formed part: between its quotes. Undefined
 * unless there is exactly one such attribute. Comments, CDATA and instructions are passed over.
 */
function attributeSpan(text: string, path: string[], attr: string): { start: number; end: number } | undefined {
  const stack: string[] = [];
  const n = text.length;
  let found: { start: number; end: number } | undefined;
  let count = 0;
  let i = 0;
  while (i < n) {
    const lt = text.indexOf('<', i);
    if (lt < 0) break;
    const skipTo = (close: string, from: number): number => {
      const at = text.indexOf(close, from);
      return at < 0 ? -1 : at + close.length;
    };
    if (text.startsWith('<!--', lt) || text.startsWith('<![CDATA[', lt) || text.startsWith('<?', lt)) {
      i = skipTo(text.startsWith('<!--', lt) ? '-->' : text.startsWith('<?', lt) ? '?>' : ']]>', lt + 2);
      if (i < 0) return undefined;
      continue;
    }
    if (text.charCodeAt(lt + 1) === 33) return undefined;
    if (text.charCodeAt(lt + 1) === 47) {
      i = skipTo('>', lt);
      if (i < 0) return undefined;
      stack.pop();
      continue;
    }
    let j = lt + 1;
    while (j < n && !isNameEnd(text.charCodeAt(j))) j++;
    const name = text.slice(lt + 1, j);
    const local = name.slice(name.indexOf(':') + 1);
    const depth = stack.length;
    const onPath = depth === path.length - 1 && local === path[depth] && stack.every((one, k) => one === path[k]);
    let self = false;
    for (;;) {
      while (j < n && isBlank(text.charCodeAt(j))) j++;
      if (j >= n) return undefined;
      const c = text.charCodeAt(j);
      if (c === 62) {
        j++;
        break;
      }
      if (c === 47) {
        self = true;
        j += 2;
        break;
      }
      let k = j;
      while (k < n && !isNameEnd(text.charCodeAt(k))) k++;
      const attrName = text.slice(j, k);
      while (k < n && text.charCodeAt(k) !== 61) k++;
      k++;
      while (k < n && isBlank(text.charCodeAt(k))) k++;
      const quote = text.charAt(k);
      if (quote !== '"' && quote !== "'") return undefined;
      const end = text.indexOf(quote, k + 1);
      if (end < 0) return undefined;
      if (onPath && attrName === attr) {
        count++;
        found = { start: k + 1, end };
      }
      j = end + 1;
    }
    if (!self) stack.push(local);
    i = j;
  }
  return count === 1 ? found : undefined;
}

/**
 * The part with the value of one attribute of the element at `path` (local names from the root)
 * replaced: its bytes with only that value changed where that can be done, else its tree.
 */
export function patchedAttribute(original: Uint8Array, path: string[], attr: string, value: string): Uint8Array | XmlDocument {
  const text = utf8Text(original);
  const span = text === undefined ? undefined : attributeSpan(text, path, attr);
  if (text !== undefined && span) {
    const bytes = encoder.encode(text.slice(0, span.start) + escapeAttr(value) + text.slice(span.end));
    // What was changed is read back: the attribute of the element the reader takes, and nothing else of the part.
    try {
      const el = elementAt(parseXml(bytes).root, path);
      if (el && el.attrs.get(attr) === value) return bytes;
    } catch (error) {
      // Written from its tree below.
    }
  }
  const doc = parseXml(original);
  const el = elementAt(doc.root, path);
  if (!el || !el.attrs.has(attr)) throw new Error(`Cannot write a pivot part: it has no ${path.join('/')} with ${attr}`);
  el.attrs.set(attr, value);
  return doc;
}

const SOURCE_PATH = ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'];
const LOCATION_PATH = ['pivotTableDefinition', 'location'];

// ---------------------------------------------------------------------------------------------
// One save

/** What earlier saves of a workbook left of its pivot parts: the bytes the file had, and the parts the last save wrote. */
interface PivotSaves {
  originals: Map<string, Uint8Array>;
  written: Set<string>;
}

const SAVES = new WeakMap<Workbook, PivotSaves>();

/** The bytes the file had for a pivot part that an earlier save of the workbook rewrote; undefined for a part no save rewrote. */
export function pivotFileBytes(workbook: Workbook, part: string): Uint8Array | undefined {
  const saved = SAVES.get(workbook);
  return saved ? saved.originals.get(part) : undefined;
}

/** The parts of the pivot tables and caches a save of the workbook writes from the file's bytes: those refreshed, and those an earlier save wrote. */
export function pivotPartsToRead(workbook: Workbook): string[] {
  const pivots = workbook.pivots;
  if (!pivots || !workbook.pivotsModified) return [];
  const saved = SAVES.get(workbook);
  const out: string[] = [];
  const want = (part: string) => {
    if (!(saved && saved.originals.has(part)) && out.indexOf(part) < 0) out.push(part);
  };
  // Tables made here are in no file: the writer makes their parts (see writeXlsx()).
  for (const cache of pivots.caches) {
    if (cache.created !== undefined) continue;
    if (cache.now || cache.sourceRef !== undefined) want(cache.part);
    if (cache.now && cache.recordsPart !== undefined) want(cache.recordsPart);
  }
  for (const table of pivots.tables) if (table.created === undefined && (table.now || table.placed)) want(table.part);
  return out;
}

/** What a save writes of the pivot parts, and what it keeps in mind for the next one once it is done. */
export interface PivotWrites {
  writes: [part: string, data: Uint8Array | XmlDocument][];
  /** Called once the save has written everything: the next save knows what this one wrote. */
  done(): void;
}

/**
 * The pivot parts a save writes. `bytesOf` gives a part as the package has it (the file's bytes:
 * a part is asked for before any save wrote it). Nothing is remembered before done() is called.
 */
export function pivotWrites(workbook: Workbook, bytesOf: (part: string) => Uint8Array | undefined): PivotWrites {
  const pivots = workbook.pivots;
  const saved = SAVES.get(workbook);
  if (!pivots || !workbook.pivotsModified) return { writes: [], done: () => undefined };
  const originals = new Map<string, Uint8Array>();
  const original = (part: string): Uint8Array => {
    const kept = (saved && saved.originals.get(part)) || originals.get(part);
    if (kept) return kept;
    const bytes = bytesOf(part);
    if (!bytes) throw new Error(`Cannot write a pivot table: the part ${part} is missing`);
    originals.set(part, bytes);
    return bytes;
  };
  const writes: [string, Uint8Array | XmlDocument][] = [];
  const written = new Set<string>();
  for (const cache of pivots.caches) {
    // Tables made here are in no file: the writer makes their parts (see writeXlsx()).
    if (cache.created !== undefined) continue;
    if (!cache.now) {
      // Only the range of its source moved: its records stay as they are, as Excel keeps them.
      if (cache.sourceRef !== undefined) {
        writes.push([cache.part, patchedAttribute(original(cache.part), SOURCE_PATH, 'ref', cache.sourceRef)]);
        written.add(cache.part);
      }
      continue;
    }
    writes.push([cache.part, writeCacheDefinition(original(cache.part), cache, cache.now, workbook.date1904)]);
    written.add(cache.part);
    if (cache.recordsPart !== undefined) {
      writes.push([cache.recordsPart, writeCacheRecords(original(cache.recordsPart), cache.now, workbook.date1904)]);
      written.add(cache.recordsPart);
    }
  }
  for (const table of pivots.tables) {
    if (table.created !== undefined) continue;
    if (table.now) writes.push([table.part, writeTableDefinition(original(table.part), table, table.now)]);
    else if (table.placed) writes.push([table.part, patchedAttribute(original(table.part), LOCATION_PATH, 'ref', rangeName(table.placed))]);
    else continue;
    written.add(table.part);
  }
  // What an earlier save wrote and this one does not goes back to the bytes of the file.
  if (saved) for (const part of saved.written) if (!written.has(part)) writes.push([part, original(part)]);
  return {
    writes,
    done: () => {
      let state = SAVES.get(workbook);
      if (!state) SAVES.set(workbook, (state = { originals: new Map(), written: new Set() }));
      for (const [part, bytes] of originals) if (!state.originals.has(part)) state.originals.set(part, bytes);
      state.written = written;
    },
  };
}
