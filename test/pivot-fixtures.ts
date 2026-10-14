// Workbooks with pivot tables, made by code the way Excel lays them out: a sheet "Data" with the
// source, a sheet "Pivot" with the table, a cache with its records and a table definition. The
// table that stands in "Pivot" is a stand-in (a header and a grand total that are out of date):
// what a refresh makes of it is what the tests look at.

import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { writeXlsx } from '../src/formats/xlsx/write';
import { open, ZipArchive, type OfficeFile, type Workbook } from '../src/index';
import { History } from '../src/ops/ops';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const encoder = new TextEncoder();

export type Datum = string | number | boolean | null | { date: number };
export type Form = 'compact' | 'outline' | 'tabular';

export interface PivotSpec {
  /** Header and rows of the source, from A1 of "Data". */
  data: Datum[][];
  rows: number[];
  cols?: number[];
  values: { fld: number; subtotal?: string; name?: string; numFmtId?: number }[];
  pages?: { fld: number; item?: Datum }[];
  form?: Form;
  /** Items hidden, by field and value. */
  hidden?: { fld: number; values: Datum[] }[];
  sortType?: { fld: number; type: 'ascending' | 'descending' }[];
  defaultSubtotal?: boolean;
  subtotalTop?: boolean;
  showAll?: boolean;
  rowGrandTotals?: boolean;
  colGrandTotals?: boolean;
  /** Where the body of the table starts on "Pivot" (0-based); default row 2, column 0 (A3). */
  at?: { row: number; col: number };
  /** Words of the stand-in table (as a Turkish Excel writes them). */
  words?: { rowHeader?: string; grand?: string; colHeader?: string };
  /** XML put at the end of the table definition (before </pivotTableDefinition>), of the cache definition, and of the records. */
  tableTail?: string;
  cacheTail?: string;
  recordsTail?: string;
  /** Attributes put on <pivotTableDefinition>. */
  tableAttrs?: string;
  /** XML put in the cacheField of a field (before </cacheField>). */
  fieldTail?: Record<number, string>;
  /** Source by a defined name instead of a range. */
  sourceName?: string;
  /** Cells of "Pivot" beside the table: [row, col, value]. */
  around?: [number, number, Datum][];
  /** Without a records part (saveData="0"). */
  noRecords?: boolean;
  /** A second table on "Pivot" made from the same cache, starting at this place. */
  twin?: { row: number; col: number };
}

const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const letter = (col: number): string => {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};
export const ref = (row: number, col: number) => `${letter(col)}${row + 1}`;

/** A cell of a sheet as Excel writes it (inline strings: no shared strings part). */
export function cellXml(row: number, col: number, value: Datum, style = 0): string {
  const s = style ? ` s="${style}"` : '';
  if (value === null) return '';
  if (typeof value === 'number') return `<c r="${ref(row, col)}"${s}><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref(row, col)}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  if (typeof value === 'object') return `<c r="${ref(row, col)}" s="2"><v>${value.date}</v></c>`;
  if (value.charAt(0) === '#') return `<c r="${ref(row, col)}"${s} t="e"><v>${esc(value)}</v></c>`;
  return `<c r="${ref(row, col)}"${s} t="inlineStr"><is><t>${esc(value)}</t></is></c>`;
}

export function sheetXml(cells: [number, number, Datum, number?][]): string {
  const rows = new Map<number, string[]>();
  const sorted = cells.slice().sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (const [r, c, v, style] of sorted) {
    const xml = cellXml(r, c, v, style);
    if (xml === '') continue;
    if (!rows.has(r)) rows.set(r, []);
    (rows.get(r) as string[]).push(xml);
  }
  const body = [...rows.entries()].map(([r, list]) => `<row r="${r + 1}">${list.join('')}</row>`).join('');
  return `${XML}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetData>${body}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
}

type Value = { t: 's' | 'n' | 'b' | 'e' | 'd' | 'm'; v?: string | number | boolean };
const valueOf = (d: Datum): Value =>
  d === null ? { t: 'm' } : typeof d === 'number' ? { t: 'n', v: d } : typeof d === 'boolean' ? { t: 'b', v: d } : typeof d === 'object' ? { t: 'd', v: d.date } : d.charAt(0) === '#' ? { t: 'e', v: d } : { t: 's', v: d };
const keyOf = (v: Value) => (v.t === 's' ? `s${String(v.v).toLowerCase()}` : `${v.t}${String(v.v)}`);

/** 1900 date system serial -> ISO text, as a cache writes it. */
export function iso(serial: number): string {
  const ms = Math.round((serial - 25569) * 86400000);
  return new Date(ms).toISOString().slice(0, 19);
}

const valueXml = (v: Value): string =>
  v.t === 'm' ? '<m/>' : v.t === 'd' ? `<d v="${iso(v.v as number)}"/>` : v.t === 'b' ? `<b v="${v.v ? 1 : 0}"/>` : `<${v.t} v="${esc(String(v.v))}"/>`;

function compare(a: Value, b: Value): number {
  const rank = (v: Value) => (v.t === 'n' || v.t === 'd' ? 0 : v.t === 's' ? 1 : v.t === 'b' ? 2 : v.t === 'e' ? 3 : 4);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a.v === 'number' && typeof b.v === 'number') return a.v - b.v;
  const x = String(a.v).toLowerCase();
  const y = String(b.v).toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

const AGG_WORDS: Record<string, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' };

/** A workbook with one pivot table (two with `twin`) made from `spec.data`. */
export async function pivotBook(spec: PivotSpec): Promise<Uint8Array> {
  const header = spec.data[0] as string[];
  const rows = spec.data.slice(1);
  const cols = spec.cols || [];
  const pages = spec.pages || [];
  const axis = new Map<number, string>();
  for (const f of spec.rows) axis.set(f, 'axisRow');
  for (const f of cols) axis.set(f, 'axisCol');
  for (const p of pages) axis.set(p.fld, 'axisPage');
  // The cache: items of every field that is not all numbers, or that stands on an axis; in the order they come.
  const shared: (Value[] | undefined)[] = header.map((_, f) => {
    const values = rows.map((row) => valueOf(row[f] === undefined ? null : row[f]));
    const plain = values.every((v) => v.t === 'n' || v.t === 'd' || v.t === 'm');
    if (plain && !axis.has(f)) return undefined;
    const seen = new Map<string, Value>();
    for (const v of values) if (!seen.has(keyOf(v))) seen.set(keyOf(v), v);
    return [...seen.values()];
  });
  // Where each value stands among the items of its field: looked up once for every record.
  const placeOf = shared.map((list) => (list ? new Map(list.map((v, x) => [keyOf(v), x] as [string, number])) : undefined));
  const cacheFields = header
    .map((name, f) => {
      const values = rows.map((row) => valueOf(row[f] === undefined ? null : row[f]));
      const list = shared[f];
      const nums = values.filter((v) => v.t === 'n').map((v) => v.v as number);
      const hasS = values.some((v) => v.t === 's');
      const attrs = hasS ? (values.some((v) => v.t === 'm') ? ' containsBlank="1"' : '') : ` containsSemiMixedTypes="0" containsString="0" containsNumber="1"${nums.every((n) => Math.floor(n) === n) ? ' containsInteger="1"' : ''} minValue="${nums.reduce((a, b) => Math.min(a, b), Infinity)}" maxValue="${nums.reduce((a, b) => Math.max(a, b), -Infinity)}"`;
      const inner = list ? `<sharedItems${attrs} count="${list.length}">${list.map(valueXml).join('')}</sharedItems>` : `<sharedItems${attrs}/>`;
      return `<cacheField name="${esc(name)}" numFmtId="0">${inner}${(spec.fieldTail || {})[f] || ''}</cacheField>`;
    })
    .join('');
  const lastRow = spec.data.length;
  const source = spec.sourceName ? `<worksheetSource name="${spec.sourceName}"/>` : `<worksheetSource ref="A1:${letter(header.length - 1)}${lastRow}" sheet="Data"/>`;
  const cacheDef =
    `${XML}<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"${spec.noRecords ? ' saveData="0"' : ' r:id="rId1"'} refreshedBy="Author" refreshedDate="45000.5" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${rows.length}">` +
    `<cacheSource type="worksheet">${source}</cacheSource><cacheFields count="${header.length}">${cacheFields}</cacheFields>${spec.cacheTail || ''}</pivotCacheDefinition>`;
  const records =
    `${XML}<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" count="${rows.length}">` +
    rows
      .map((row) => {
        const kids = header.map((_, f) => {
          const v = valueOf(row[f] === undefined ? null : row[f]);
          const list = shared[f];
          if (!list) return valueXml(v);
          const x = (placeOf[f] as Map<string, number>).get(keyOf(v)) as number;
          return x === 0 ? '<x/>' : `<x v="${x}"/>`;
        });
        return `<r>${kids.join('')}</r>`;
      })
      .join('') +
    `${spec.recordsTail || ''}</pivotCacheRecords>`;

  const form = spec.form || 'compact';
  const formAttrs = form === 'compact' ? '' : form === 'outline' ? ' compact="0"' : ' compact="0" outline="0"';
  const hiddenOf = (f: number) => (spec.hidden || []).filter((h) => h.fld === f).flatMap((h) => h.values.map((d) => keyOf(valueOf(d))));
  const pivotFields = header
    .map((_, f) => {
      const list = shared[f];
      const ax = axis.get(f);
      const sort = (spec.sortType || []).find((s) => s.fld === f);
      let attrs = ax ? ` axis="${ax}"` : '';
      if (spec.values.some((v) => v.fld === f)) attrs += ' dataField="1"';
      attrs += `${formAttrs} showAll="${spec.showAll ? 1 : 0}"`;
      if (sort) attrs += ` sortType="${sort.type}"`;
      if (spec.defaultSubtotal === false) attrs += ' defaultSubtotal="0"';
      if (spec.subtotalTop === false) attrs += ' subtotalTop="0"';
      if (!ax || !list) return `<pivotField${attrs}/>`;
      const order = list.map((v, x) => ({ v, x })).sort((a, b) => compare(a.v, b.v));
      const hidden = hiddenOf(f);
      const items = order.map(({ v, x }) => `<item${hidden.indexOf(keyOf(v)) >= 0 ? ' h="1"' : ''} x="${x}"/>`).join('') + (spec.defaultSubtotal === false ? '' : '<item t="default"/>');
      return `<pivotField${attrs}><items count="${order.length + (spec.defaultSubtotal === false ? 0 : 1)}">${items}</items></pivotField>`;
    })
    .join('');
  const D = spec.values.length;
  const m = cols.length;
  const n = spec.rows.length;
  const L = form === 'compact' ? 1 : n;
  const at = spec.at || { row: 2, col: 0 };
  const words = spec.words || {};
  // The stand-in: header, grand total.
  const stand: [number, number, Datum][] = [];
  let width: number;
  let height: number;
  let firstHeaderRow: number;
  let firstDataRow: number;
  let colItems: string;
  const nameOf = (v: PivotSpec['values'][number]) => v.name || `${AGG_WORDS[v.subtotal || 'sum']} of ${header[v.fld]}`;
  if (m === 0) {
    firstDataRow = 1;
    firstHeaderRow = D > 1 ? 0 : 1;
    width = L + D;
    height = 2;
    colItems = D === 1 ? '<colItems count="1"><i/></colItems>' : `<colItems count="${D}">${spec.values.map((_, d) => (d === 0 ? '<i><x/></i>' : `<i i="${d}"><x v="${d}"/></i>`)).join('')}</colItems>`;
    spec.values.forEach((v, d) => stand.push([0, L + d, nameOf(v)]));
    stand.push([1, L, 0]);
  } else {
    firstDataRow = m + 1;
    firstHeaderRow = 1;
    width = L + 1;
    height = m + 2;
    colItems = '<colItems count="1"><i t="grand"><x/></i></colItems>';
    stand.push([0, 0, nameOf(spec.values[0])]);
    stand.push([0, L, words.colHeader || 'Column Labels']);
    stand.push([1, L, words.grand || 'Grand Total']);
    stand.push([firstDataRow, L, 0]);
  }
  if (form === 'compact') stand.push([firstDataRow - 1, 0, words.rowHeader || 'Row Labels']);
  else spec.rows.forEach((f, j) => stand.push([firstDataRow - 1, j, header[f]]));
  stand.push([firstDataRow, 0, words.grand || 'Grand Total']);
  const location = (row: number, col: number) => `<location ref="${ref(row, col)}:${ref(row + height - 1, col + width - 1)}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${firstDataRow}" firstDataCol="${L}"${pages.length > 0 ? ` rowPageCount="${pages.length}" colPageCount="1"` : ''}/>`;
  const pageItems = pages.map((p) => {
    if (p.item === undefined) return `<pageField fld="${p.fld}" hier="-1"/>`;
    const list = shared[p.fld] as Value[];
    const order = list.map((v, x) => ({ v, x })).sort((a, b) => compare(a.v, b.v));
    const i = order.findIndex((o) => keyOf(o.v) === keyOf(valueOf(p.item as Datum)));
    return `<pageField fld="${p.fld}" item="${i}" hier="-1"/>`;
  });
  const fieldsList = (list: number[]) => list.map((f) => `<field x="${f}"/>`).join('');
  const colFieldsXml = m > 0 ? `<colFields count="${m}">${fieldsList(cols)}</colFields>` : D > 1 ? '<colFields count="1"><field x="-2"/></colFields>' : '';
  const table = (name: string, row: number, col: number) =>
    `${XML}<pivotTableDefinition xmlns="${NS_MAIN}" name="${name}" cacheId="7" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0"${form === 'compact' ? '' : ' compact="0" compactData="0"'}${form === 'tabular' ? ' outline="1" outlineData="1"' : form === 'outline' ? ' outline="1" outlineData="1"' : ' outline="1" outlineData="1"'} multipleFieldFilters="0"${spec.rowGrandTotals === false ? ' rowGrandTotals="0"' : ''}${spec.colGrandTotals === false ? ' colGrandTotals="0"' : ''}${spec.tableAttrs || ''}>` +
    `${location(row, col)}<pivotFields count="${header.length}">${pivotFields}</pivotFields>` +
    `<rowFields count="${n}">${fieldsList(spec.rows)}</rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems>` +
    `${colFieldsXml}${colItems}` +
    (pages.length > 0 ? `<pageFields count="${pages.length}">${pageItems.join('')}</pageFields>` : '') +
    `<dataFields count="${D}">${spec.values.map((v) => `<dataField name="${esc(nameOf(v))}" fld="${v.fld}"${v.subtotal && v.subtotal !== 'sum' ? ` subtotal="${v.subtotal}"` : ''} baseField="0" baseItem="0"${v.numFmtId !== undefined ? ` numFmtId="${v.numFmtId}"` : ''}/>`).join('')}</dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>${spec.tableTail || ''}</pivotTableDefinition>`;

  const dataCells: [number, number, Datum, number?][] = [];
  spec.data.forEach((row, r) => row.forEach((v, c) => dataCells.push([r, c, v === undefined ? null : v])));
  const pivotCells: [number, number, Datum, number?][] = stand.map(([r, c, v]) => [at.row + r, at.col + c, v]);
  if (spec.twin) for (const [r, c, v] of stand) pivotCells.push([spec.twin.row + r, spec.twin.col + c, v]);
  for (const [r, c, v] of spec.around || []) pivotCells.push([r, c, v]);
  if (pages.length > 0) {
    pages.forEach((p, i) => {
      pivotCells.push([at.row - pages.length - 1 + i, at.col, header[p.fld]]);
      pivotCells.push([at.row - pages.length - 1 + i, at.col + 1, p.item === undefined ? '(All)' : (p.item as Datum)]);
    });
  }

  const zip = new ZipArchive();
  const tables = spec.twin ? 2 : 1;
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="${CT}.sheet.main+xml"/>`,
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT}.worksheet+xml"/>`,
    `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="${CT}.worksheet+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="${CT}.styles+xml"/>`,
    `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${CT}.pivotCacheDefinition+xml"/>`,
    ...(spec.noRecords ? [] : [`<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="${CT}.pivotCacheRecords+xml"/>`]),
    ...Array.from({ length: tables }, (_, i) => `<Override PartName="/xl/pivotTables/pivotTable${i + 1}.xml" ContentType="${CT}.pivotTable+xml"/>`),
  ];
  zip.set(
    '[Content_Types].xml',
    encoder.encode(`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.join('')}</Types>`),
  );
  zip.set('_rels/.rels', encoder.encode(`${XML}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`));
  const names = spec.sourceName ? `<definedNames><definedName name="${spec.sourceName}">Data!$A$1:$${letter(header.length - 1)}$${lastRow}</definedName></definedNames>` : '';
  zip.set(
    'xl/workbook.xml',
    encoder.encode(
      `${XML}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Pivot" sheetId="2" r:id="rId2"/></sheets>${names}<calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches></workbook>`,
    ),
  );
  zip.set(
    'xl/_rels/workbook.xml.rels',
    encoder.encode(
      `${XML}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${NS_REL}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="${NS_REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/><Relationship Id="rId4" Type="${NS_REL}/styles" Target="styles.xml"/></Relationships>`,
    ),
  );
  zip.set(
    'xl/styles.xml',
    encoder.encode(
      `${XML}<styleSheet xmlns="${NS_MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    ),
  );
  zip.set('xl/worksheets/sheet1.xml', encoder.encode(sheetXml(dataCells)));
  zip.set('xl/worksheets/sheet2.xml', encoder.encode(sheetXml(pivotCells)));
  zip.set(
    'xl/worksheets/_rels/sheet2.xml.rels',
    encoder.encode(
      `${XML}<Relationships xmlns="${NS_PKG}">${Array.from({ length: tables }, (_, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/pivotTable" Target="../pivotTables/pivotTable${i + 1}.xml"/>`).join('')}</Relationships>`,
    ),
  );
  zip.set('xl/pivotTables/pivotTable1.xml', encoder.encode(table('PivotTable1', at.row, at.col)));
  if (spec.twin) zip.set('xl/pivotTables/pivotTable2.xml', encoder.encode(table('PivotTable2', spec.twin.row, spec.twin.col)));
  for (let i = 1; i <= tables; i++) {
    zip.set(`xl/pivotTables/_rels/pivotTable${i}.xml.rels`, encoder.encode(`${XML}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`));
  }
  zip.set('xl/pivotCache/pivotCacheDefinition1.xml', encoder.encode(cacheDef));
  if (!spec.noRecords) {
    zip.set('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', encoder.encode(`${XML}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>`));
    zip.set('xl/pivotCache/pivotCacheRecords1.xml', encoder.encode(records));
  }
  return zip.write();
}

export interface Loaded {
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

export async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { file, wb, history: new History(wb) };
}

export async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

/** The cells of a range of a sheet as values (undefined for none). */
export function grid(wb: Workbook, sheet: number, top: number, left: number, bottom: number, right: number): unknown[][] {
  const out: unknown[][] = [];
  for (let r = top; r <= bottom; r++) {
    const row: unknown[] = [];
    for (let c = left; c <= right; c++) {
      const cell = wb.sheets[sheet].rows.get(r)?.cells[c];
      row.push(cell ? cell.value : undefined);
    }
    out.push(row);
  }
  return out;
}

/** Sets a cell of "Data" (sheet 0) through an operation. */
export function setData(doc: Loaded, row: number, col: number, value: Datum): void {
  const cell =
    value === null
      ? undefined
      : typeof value === 'number'
        ? { kind: 'number' as const, value, style: 0 }
        : typeof value === 'boolean'
          ? { kind: 'boolean' as const, value, style: 0 }
          : typeof value === 'object'
            ? { kind: 'number' as const, value: value.date, style: 2 }
            : { kind: 'string' as const, value, style: 0 };
  doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row, col, cell }] }], undefined, undefined);
}

export const text = async (bytes: Uint8Array, name: string): Promise<string> => new TextDecoder().decode(await ZipArchive.read(bytes).read(name));

export function same(actual: Uint8Array, expected: Uint8Array): boolean {
  if (actual.length !== expected.length) return false;
  for (let i = 0; i < actual.length; i++) if (actual[i] !== expected[i]) return false;
  return true;
}

/** Whether python3 can run the independent checks. */
export function hasPython(): boolean {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
    return true;
  } catch {
    return false;
  }
}

/** Whether python3 has openpyxl (installed for the user or the system). */
export function hasOpenpyxl(): boolean {
  try {
    execFileSync('python3', ['-c', 'import openpyxl'], { stdio: 'ignore', timeout: 60000 });
    return true;
  } catch {
    return false;
  }
}

/** What the independent check of pivot tables (test/tools/pivot_check.py) and the package check find wrong with files; empty: all is well. */
export function independentPivots(files: { [name: string]: Uint8Array }): string[] {
  const dir = mkdtempSync(join(process.cwd(), 'test', '.pivot-'));
  try {
    for (const name of Object.keys(files)) writeFileSync(join(dir, `${name}.xlsx`), files[name]);
    const run = (args: string[]) => JSON.parse(execFileSync('python3', args, { encoding: 'utf8', timeout: 300000, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'inherit'] })) as string[];
    const tools = join(process.cwd(), 'test', 'tools');
    // The pivot check also opens each file with openpyxl, which may be installed for the user only: not isolated (-I) then.
    const pivots = [join(tools, 'pivot_check.py'), dir];
    if (hasOpenpyxl()) pivots.push('--require-openpyxl');
    return run(['-I', join(tools, 'package_check.py'), dir]).concat(run(pivots));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
