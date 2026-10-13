// What the tests of chart editing share: a workbook of the file with some numbers, loading
// and saving, and a check that a package is whole (every part has a content type, every
// relationship leads to a part, no part is left without a relationship that leads to it).

import assert from 'node:assert/strict';
import { writeXlsx } from '../src/formats/xlsx/write';
import { createSheet, emptyStyleSheet, open, parseXml, ZipArchive, type Cell, type Chart, type DrawingChart, type OfficeFile, type Sheet, type SheetDrawing, type Workbook, type XmlElement } from '../src/index';
import { History } from '../src/ops/ops';
import { catAx, chartWorkbook, chartXml, group, numRef, ser, sheetData, strRef, valAx, type ChartBookSpec } from './chart-fixtures';

const decoder = new TextDecoder();

export const SALES: (number | string | null)[][] = [
  ['Month', 'North', 'South'],
  ['Jan', 10, 30],
  ['Feb', 20, 20],
  ['Mar', 30, 10],
];

/** A workbook with the numbers of SALES on its first sheet ("Sayfa1"), and whatever else the spec says. */
export function salesBook(spec: ChartBookSpec = {}): Promise<Uint8Array> {
  return chartWorkbook({ sheet: `${sheetData(SALES)}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`, ...spec });
}

export interface Loaded {
  bytes: Uint8Array;
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

export async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { bytes, file, wb, history: new History(wb) };
}

export async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

export const reopen = async (bytes: Uint8Array): Promise<Workbook> => (await open(bytes)).workbook();
export const names = (bytes: Uint8Array): string[] => ZipArchive.read(bytes).names().sort();
export const text = async (bytes: Uint8Array, name: string): Promise<string> => decoder.decode(await ZipArchive.read(bytes).read(name));
export const part = (bytes: Uint8Array, name: string): Promise<Uint8Array> => ZipArchive.read(bytes).read(name);

export function same(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

export const chartsOf = (sheet: Sheet): DrawingChart[] => (sheet.drawings || []).map((drawing: SheetDrawing) => drawing.object).filter((object): object is DrawingChart => object.type === 'chart');

/** A chart without what differs between the model and what is read back of it: the cached values (the cells), that it follows edits, and what the reader remarked about the part it read. */
export function bare(chart: Chart): unknown {
  const strip = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(strip);
    if (!value || typeof value !== 'object') return value;
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value as object)) {
      if (key === 'cache' || key === 'count' || key === 'formatCode' || key === 'live' || key === 'notes') continue;
      const item = (value as Record<string, unknown>)[key];
      // The text a name cached is the cell, too.
      if (key === 'name' && item && typeof item === 'object' && (item as { ref?: string }).ref !== undefined) out[key] = { ref: (item as { ref: string }).ref };
      else out[key] = strip(item);
    }
    return out;
  };
  return strip(chart);
}

function resolve(source: string, target: string): string {
  const base = target.charAt(0) === '/' ? target : `${source.slice(0, source.lastIndexOf('/'))}/${target}`;
  const segments: string[] = [];
  for (const segment of base.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return segments.join('/');
}

const elementsOf = (root: XmlElement): XmlElement[] => root.children.filter((node): node is XmlElement => node.type === 'element');

/**
 * What is wrong with a package, one line each; none for a package that is whole: every part is
 * well-formed XML where it is XML, has a content type, is led to by a relationship, and every
 * relationship and every override leads to a part that is there.
 */
export async function packageProblems(bytes: Uint8Array): Promise<string[]> {
  const problems: string[] = [];
  const zip = ZipArchive.read(bytes);
  const all = zip.names();
  const held = new Set(all);
  const types = parseXml(await zip.read('[Content_Types].xml')).root;
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  for (const el of elementsOf(types)) {
    if (el.name === 'Default') defaults.add((el.attrs.get('Extension') || '').toLowerCase());
    if (el.name === 'Override') {
      const name = (el.attrs.get('PartName') || '').slice(1);
      if (overrides.has(name)) problems.push(`two overrides for ${name}`);
      overrides.add(name);
      if (!held.has(name)) problems.push(`an override for ${name}, which is not in the package`);
    }
  }
  const reached = new Set<string>();
  const queue = [''];
  const relsOf = (source: string) => (source === '' ? '_rels/.rels' : `${source.slice(0, source.lastIndexOf('/') + 1)}_rels/${source.slice(source.lastIndexOf('/') + 1)}.rels`);
  for (let i = 0; i < queue.length; i++) {
    const source = queue[i];
    const rels = relsOf(source);
    if (!held.has(rels)) continue;
    const ids = new Set<string>();
    for (const rel of elementsOf(parseXml(await zip.read(rels)).root)) {
      const id = rel.attrs.get('Id') || '';
      if (ids.has(id)) problems.push(`${rels}: the id ${id} twice`);
      ids.add(id);
      if (rel.attrs.get('TargetMode') === 'External') continue;
      const target = rel.attrs.get('Target') || '';
      if (target.charAt(0) === '#') continue;
      const name = resolve(`/${source}`, target);
      if (!held.has(name)) problems.push(`${rels}: ${id} leads to ${name}, which is not in the package`);
      else if (!reached.has(name)) {
        reached.add(name);
        queue.push(name);
      }
    }
  }
  for (const name of all) {
    if (name === '[Content_Types].xml') continue;
    const isRels = /(^|\/)_rels\/[^/]*\.rels$/.test(name);
    if (isRels) {
      const owner = name.replace(/_rels\/([^/]*)\.rels$/, '$1');
      if (owner !== '' && !held.has(owner)) problems.push(`${name}: the relationships of a part that is not in the package`);
    } else if (!reached.has(name)) problems.push(`${name}: no relationship leads to it`);
    const extension = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
    if (!isRels && !overrides.has(name) && !defaults.has(extension)) problems.push(`${name}: no content type`);
    if (/\.(xml|rels|vml)$/.test(name)) {
      try {
        const raw = await zip.read(name);
        const line = encodingProblem(raw);
        if (line) problems.push(`${name}: ${line}`);
        const root = parseXml(raw).root;
        const ids = new Set<string>();
        const rels = relsOf(name);
        if (!isRels && held.has(rels)) for (const rel of elementsOf(parseXml(await zip.read(rels)).root)) ids.add(rel.attrs.get('Id') || '');
        for (const line of namespaceProblems(root, isRels ? undefined : ids)) problems.push(`${name}: ${line}`);
      } catch (error) {
        problems.push(`${name}: not well-formed (${(error as Error).message})`);
      }
    }
  }
  return problems;
}

/**
 * What is wrong with the bytes of an XML part for the encoding it declares; undefined when a
 * parser that honours the declaration reads the text that was meant. A part in UTF-16 is known
 * by its byte order mark; every other part has to be UTF-8, and one that declares anything else
 * may hold no byte beyond ASCII (such a byte is two letters, or none, to the parser).
 */
export function encodingProblem(bytes: Uint8Array): string | undefined {
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) return undefined;
  let head = '';
  for (let i = 0; i < Math.min(bytes.length, 200); i++) head += String.fromCharCode(bytes[i]);
  const declared = /^(?:\u00ef\u00bb\u00bf)?<\?xml[^?]*?\sencoding\s*=\s*["']([^"']*)["']/.exec(head);
  if (declared && !/^utf-?8$/i.test(declared[1])) {
    for (let i = 0; i < bytes.length; i++) if (bytes[i] >= 0x80) return `declares the encoding ${declared[1]} and holds bytes that are not ASCII`;
    return undefined;
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch (error) {
    return 'is not UTF-8';
  }
  return undefined;
}

const RELATIONSHIPS = /^http:\/\/(schemas\.openxmlformats\.org\/officeDocument\/2006|purl\.oclc\.org\/ooxml\/officeDocument)\/relationships$/;
const CHART_SPACE = /^http:\/\/(schemas\.openxmlformats\.org\/drawingml\/2006|purl\.oclc\.org\/ooxml\/drawingml)\/chart$/;

/**
 * What is wrong with the namespaces of a part, one line each: a prefix that is bound to nothing
 * (which a parser that knows namespaces refuses the part for), elements of a chart in two
 * namespaces of charts at once (the transitional one and the strict one), and - with the ids of
 * the relationships of the part - an attribute that names a relationship the part does not have.
 */
export function namespaceProblems(root: XmlElement, ids?: Set<string>): string[] {
  const problems: string[] = [];
  const spaces = new Set<string>();
  const walk = (el: XmlElement, scope: Map<string, string>, depth: number): void => {
    let own = scope;
    for (const [name, value] of el.attrs) {
      if (name !== 'xmlns' && name.slice(0, 6) !== 'xmlns:') continue;
      if (own === scope) own = new Map(scope);
      own.set(name === 'xmlns' ? '' : name.slice(6), value);
    }
    const colon = el.name.indexOf(':');
    const prefix = colon < 0 ? '' : el.name.slice(0, colon);
    const space = own.get(prefix);
    if (prefix !== '' && space === undefined && problems.length < 20) problems.push(`the prefix of <${el.name}> is bound to no namespace`);
    if (space !== undefined && CHART_SPACE.test(space)) spaces.add(space);
    for (const [name, value] of el.attrs) {
      const at = name.indexOf(':');
      if (at < 0 || name.slice(0, 6) === 'xmlns:' || name.slice(0, 4) === 'xml:') continue;
      const bound = own.get(name.slice(0, at));
      if (bound === undefined && problems.length < 20) problems.push(`the prefix of the attribute ${name} of <${el.name}> is bound to no namespace`);
      if (bound !== undefined && ids && RELATIONSHIPS.test(bound) && !ids.has(value) && problems.length < 20) problems.push(`<${el.name}> names the relationship ${value}, which the part does not have`);
    }
    if (depth < 200) for (const node of el.children) if (node.type === 'element') walk(node, own, depth + 1);
  };
  walk(root, new Map(), 0);
  if (spaces.size > 1) problems.push(`its elements are of several namespaces of charts: ${[...spaces].join(', ')}`);
  return problems;
}

/** The relationship ids a part uses (r:id, r:embed, ...) that its relationships do not have, and the relationships nothing in the part names. */
export async function relationProblems(bytes: Uint8Array, name: string): Promise<string[]> {
  const zip = ZipArchive.read(bytes);
  const xml = decoder.decode(await zip.read(name));
  const relsName = `${name.slice(0, name.lastIndexOf('/') + 1)}_rels/${name.slice(name.lastIndexOf('/') + 1)}.rels`;
  const ids = new Set<string>();
  if (zip.names().indexOf(relsName) >= 0) for (const rel of elementsOf(parseXml(await zip.read(relsName)).root)) ids.add(rel.attrs.get('Id') || '');
  const used = new Set<string>();
  const pattern = /\br:(?:id|embed|link)="([^"]*)"/g;
  for (let match = pattern.exec(xml); match; match = pattern.exec(xml)) used.add(match[1]);
  const problems: string[] = [];
  for (const id of used) if (!ids.has(id)) problems.push(`${name}: names the relationship ${id}, which it does not have`);
  for (const id of ids) if (!used.has(id)) problems.push(`${name}: the relationship ${id} is named by nothing`);
  return problems;
}

/** A chart as another application writes one: with a style in an alternate content, text properties, a gradient, extensions and print settings, none of which a model holds. */
export const FOREIGN = chartXml({
  head: '<c:title><c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="tr-TR" sz="1400" b="1"/><a:t>Old</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="tr-TR"/></a:p></c:txPr></c:title><c:autoTitleDeleted val="0"/>',
  plot:
    group(
      'barChart',
      '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
      ser(0, {
        tx: strRef('Sayfa1!$B$1', ['North']),
        cat: strRef('Sayfa1!$A$2:$A$4', ['Jan', 'Feb', 'Mar']),
        val: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]),
        tail: '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{00000000-AAAA}"/></c:ext></c:extLst>',
      }) +
        ser(1, {
          tx: strRef('Sayfa1!$C$1', ['South']),
          cat: strRef('Sayfa1!$A$2:$A$4', ['Jan', 'Feb', 'Mar']),
          val: numRef('Sayfa1!$C$2:$C$4', [9, 9, 9]),
          props: '<c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs></a:gsLst></a:gradFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:effectLst/></c:spPr>',
        }),
      '<c:gapWidth val="219"/>',
    ) +
    catAx('<c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr>') +
    valAx(),
  tail: '<c:legend><c:legendPos val="r"/><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>',
  after:
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>' +
    '<c:extLst><c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}"><c14:pivotOptions xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"/></c:ext></c:extLst>',
  before: '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>',
});


/** A workbook of the model alone (no file behind it): a sheet "Data" with these cells, and whatever other sheets are asked for. */
export function bookOf(rows: (number | string | null)[][], more: { name: string; rows: (number | string | null)[][] }[] = []): Workbook {
  const make = (name: string, id: number, data: (number | string | null)[][]): Sheet => {
    const sheet = createSheet({ name, sheetId: id, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${id}.xml` });
    data.forEach((row, r) => {
      const cells: Cell[] = [];
      row.forEach((value, c) => {
        if (value === null) return;
        cells[c] = typeof value === 'number' ? { kind: 'number', value, style: 0 } : { kind: 'string', value, style: 0 };
      });
      sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
      sheet.rowCount = Math.max(sheet.rowCount, r + 1);
      sheet.colCount = Math.max(sheet.colCount, row.length);
    });
    return sheet;
  };
  const sheets = [make('Data', 1, rows)].concat(more.map((other, i) => make(other.name, i + 2, other.rows)));
  return { sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}
