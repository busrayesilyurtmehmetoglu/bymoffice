// The definitions of a pivot table that is made here (src/ops/pivot-create.ts), and of its cache:
// the smallest parts Excel writes for a new table, before anything is in them. They are read into
// the model by the reader (readMadePivot()), and a refresh fills them in as it fills in the parts
// of a file: the items of the fields, the records, the lines of rows and columns, where the table
// stands. The writer starts from them in the place of the bytes of a file.
//
// What they say, as Excel 2016 writes a new table (UNSURE: not compared with a file of Excel byte
// for byte): the attributes of <pivotTableDefinition> Excel puts on a table it makes (dataCaption
// is required by the schema), compact form ("outline" and "compact" left at their defaults),
// showAll="0" on every field (only items with data are shown), the style Excel gives a new table
// (PivotStyleLight16), and a cache with its records saved (saveData left at its default).

import { escapeAttr, encodeXstring } from './write';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** The namespaces the definitions are made with; a workbook of the strict kind gets its own at save (see madePartText()). */
export const MADE_NAMESPACES = { main: NS_MAIN, rel: NS_REL };

/** What a table is made of. Fields are places in `fields` (the columns of the source). */
export interface MadeTableSpec {
  name: string;
  fields: string[];
  rows: number[];
  cols: number[];
  pages: number[];
  values: { fld: number; subtotal: string; name: string; numFmtId?: number }[];
  /** The cell the body starts at, and the rows the fields of the page take above it. */
  at: { row: number; col: number };
  rowGrandTotals: boolean;
  colGrandTotals: boolean;
}

/** Where a cache is read from: a range of a sheet, or a defined name. */
export type MadeSource = { kind: 'range'; sheet: string; ref: string } | { kind: 'name'; name: string };

const letters = (col: number): string => {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};

/**
 * The definition of a cache: its source and the names of its fields, without items (a refresh lists
 * them). `formats` are the number formats of the fields (that of the dates of a column of dates, as
 * Excel takes it from the source); 0 by default.
 */
export function madeCacheXml(source: MadeSource, fields: string[], formats: number[] = []): string {
  const where = source.kind === 'range' ? `<worksheetSource ref="${escapeAttr(source.ref)}" sheet="${escapeAttr(source.sheet)}"/>` : `<worksheetSource name="${escapeAttr(source.name)}"/>`;
  const list = fields.map((name, f) => `<cacheField name="${escapeAttr(encodeXstring(name))}" numFmtId="${formats[f] || 0}"><sharedItems/></cacheField>`).join('');
  return (
    `${XML}<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" r:id="rId1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">` +
    `<cacheSource type="worksheet">${where}</cacheSource><cacheFields count="${fields.length}">${list}</cacheFields></pivotCacheDefinition>`
  );
}

/** The records of a cache, before there are any. */
export function madeRecordsXml(): string {
  return `${XML}<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" count="0"/>`;
}

/** The definition of a table: its fields and where they stand; the rest (items, lines, its size) is a refresh's. */
export function madeTableXml(spec: MadeTableSpec): string {
  const axisOf = new Map<number, string>();
  for (const f of spec.rows) axisOf.set(f, 'axisRow');
  for (const f of spec.cols) axisOf.set(f, 'axisCol');
  for (const f of spec.pages) axisOf.set(f, 'axisPage');
  const data = new Set(spec.values.map((value) => value.fld));
  const fields = spec.fields
    .map((name, f) => {
      const axis = axisOf.get(f);
      const attrs = `${axis ? ` axis="${axis}"` : ''}${data.has(f) ? ' dataField="1"' : ''} showAll="0"`;
      // A field in rows, columns or the page lists its items (filled in by a refresh).
      return axis ? `<pivotField${attrs}><items count="0"/></pivotField>` : `<pivotField${attrs}/>`;
    })
    .join('');
  const list = (name: string, items: number[]) => (items.length === 0 ? '' : `<${name} count="${items.length}">${items.map((x) => `<field x="${x}"/>`).join('')}</${name}>`);
  const cols = spec.values.length > 1 ? spec.cols.concat([-2]) : spec.cols;
  const pages = spec.pages.length === 0 ? '' : `<pageFields count="${spec.pages.length}">${spec.pages.map((f) => `<pageField fld="${f}" hier="-1"/>`).join('')}</pageFields>`;
  const values = spec.values
    .map((value) => `<dataField name="${escapeAttr(encodeXstring(value.name))}" fld="${value.fld}"${value.subtotal === 'sum' ? '' : ` subtotal="${value.subtotal}"`} baseField="0" baseItem="0"${value.numFmtId !== undefined ? ` numFmtId="${value.numFmtId}"` : ''}/>`)
    .join('');
  const cell = `${letters(spec.at.col)}${spec.at.row + 1}`;
  const pageCounts = spec.pages.length > 0 ? ` rowPageCount="${spec.pages.length}" colPageCount="1"` : '';
  return (
    `${XML}<pivotTableDefinition xmlns="${NS_MAIN}" name="${escapeAttr(encodeXstring(spec.name))}" cacheId="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values"` +
    `${spec.rowGrandTotals ? '' : ' rowGrandTotals="0"'}${spec.colGrandTotals ? '' : ' colGrandTotals="0"'} updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${cell}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"${pageCounts}/>` +
    `<pivotFields count="${spec.fields.length}">${fields}</pivotFields>` +
    `${list('rowFields', spec.rows)}${list('colFields', cols)}${pages}<dataFields count="${spec.values.length}">${values}</dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`
  );
}

/** A definition made here as the workbook it is saved in has its namespaces (the strict kind has others). */
export function madePartText(xml: string, main: string, rel: string): string {
  if (main === NS_MAIN && rel === NS_REL) return xml;
  const head = xml.indexOf('>', xml.indexOf('<pivot'));
  return xml.slice(0, head).split(`"${NS_MAIN}"`).join(`"${escapeAttr(main)}"`).split(`"${NS_REL}"`).join(`"${escapeAttr(rel)}"`) + xml.slice(head);
}
