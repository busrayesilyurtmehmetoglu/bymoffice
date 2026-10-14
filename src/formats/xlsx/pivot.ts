// Reads the pivot tables of an xlsx package and their caches into the model (src/model/sheet/pivot.ts).
//
// Nothing here changes what is shown: the cells a table shows are in its sheet. What is read is
// what a refresh needs to build those cells again. A table or a cache that holds something a
// refresh here cannot build (grouped items, calculated fields, filters by label or value, data
// from outside the workbook, ...) is read all the same and marked with the reason (`refusal`):
// its refresh is refused instead of making a table Excel would not. The parts stay in the
// package as the file has them until a table is refreshed.

import type { OpcPackage, Relationship } from '../../opc';
import {
  MAX_PIVOT_FIELDS,
  MAX_PIVOT_ITEMS,
  MAX_PIVOT_PART_BYTES,
  MAX_PIVOT_TABLES,
  PIVOT_SCAN_LIMITS,
  PIVOT_AGGREGATES,
  type PivotAggregate,
  type PivotAxis,
  type PivotCache,
  type PivotCacheField,
  type PivotDataField,
  type PivotField,
  type PivotItem,
  type PivotLine,
  type PivotLocation,
  type PivotPageField,
  type PivotTable,
  type PivotValue,
  type WorkbookPivots,
} from '../../model/sheet/pivot';
import { parseRange, type Sheet, type Workbook } from '../../model/sheet/types';
import { isoToSerial } from '../../format/cell-format';
import { relationshipsOf } from '../../opc';
import { inflateRaw } from '../../deflate';
import { childElements, localName, parseXml, type XmlElement } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { decodeXstring } from './shared-strings';

const child = (parent: XmlElement, local: string): XmlElement | undefined => childElements(parent, local)[0];

/** Attribute by local name, for attributes that carry a namespace prefix such as r:id. */
function prefixedAttr(element: XmlElement, local: string): string | undefined {
  for (const [name, value] of element.attrs) if (name.indexOf(':') >= 0 && localName(name) === local) return value;
  return undefined;
}

function bool(value: string | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return fallback;
}

/** A whole number of the file within bounds; undefined for anything else. */
function int(value: string | undefined, min: number, max: number): number | undefined {
  if (value === undefined || !/^-?\d{1,10}$/.test(value)) return undefined;
  const n = Number(value);
  return n >= min && n <= max ? n : undefined;
}

/** A part of the file the reader cannot take apart. */
class Unreadable extends Error {}

/** Counts what the pivot parts of a workbook list, against MAX_PIVOT_ITEMS. */
interface Budget {
  items: number;
}

function spend(budget: Budget, count: number): void {
  budget.items += count;
  if (budget.items > MAX_PIVOT_ITEMS) throw new Unreadable('the pivot tables of the workbook list too many items');
}

/** A value of a cache as an element of the file has it (<s>, <n>, <b>, <e>, <d>, <m>); undefined for an element that is none. */
export function cacheValue(element: XmlElement, date1904: boolean): PivotValue | undefined {
  const v = element.attrs.get('v');
  switch (localName(element.name)) {
    case 's':
      return { t: 's', v: v === undefined ? '' : decodeXstring(v) };
    case 'n': {
      const n = v === undefined ? NaN : Number(v);
      if (!isFinite(n) || v === undefined || !/^[\s]*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?[\s]*$/.test(v)) throw new Unreadable(`a number of the cache is not one: "${String(v).slice(0, 40)}"`);
      return { t: 'n', v: n };
    }
    case 'b':
      return { t: 'b', v: v === '1' || v === 'true' };
    case 'e':
      return { t: 'e', v: v === undefined ? '#N/A' : v };
    case 'd': {
      const serial = v === undefined ? undefined : isoToSerial(v, date1904);
      if (serial === undefined) throw new Unreadable(`a date of the cache is not one: "${String(v).slice(0, 40)}"`);
      return { t: 'd', v: serial };
    }
    case 'm':
      return { t: 'm' };
    default:
      return undefined;
  }
}

/** Elements of a cache definition that only caches of OLAP and other outside sources have. */
const OLAP_ELEMENTS = ['cacheHierarchies', 'kpis', 'tupleCache', 'calculatedMembers', 'dimensions', 'measureGroups', 'maps'];

/** Reads a cache definition. Throws Unreadable for a part whose content is not one. */
function parseCache(root: XmlElement, cacheId: number, part: string, recordsPart: string | undefined, date1904: boolean, budget: Budget): PivotCache {
  if (localName(root.name) !== 'pivotCacheDefinition') throw new Unreadable('the part is not a pivot cache definition');
  const cache: PivotCache = { cacheId, part, fields: [] };
  if (recordsPart !== undefined) cache.recordsPart = recordsPart;
  const refuse = (reason: string) => {
    if (cache.refusal === undefined) cache.refusal = reason;
  };
  const source = child(root, 'cacheSource');
  const type = source ? source.attrs.get('type') : undefined;
  if (!source || type !== 'worksheet') refuse('its data does not come from a range of the workbook (external data, OLAP, consolidation or scenarios)');
  const sheetSource = source ? child(source, 'worksheetSource') : undefined;
  if (source && type === 'worksheet') {
    if (!sheetSource) refuse('its source is not stated');
    else if (prefixedAttr(sheetSource, 'id') !== undefined) refuse('its data comes from another workbook');
    else {
      const name = sheetSource.attrs.get('name');
      const ref = sheetSource.attrs.get('ref');
      const sheet = sheetSource.attrs.get('sheet');
      if (name !== undefined && name !== '') cache.source = sheet === undefined ? { kind: 'name', name } : { kind: 'name', name, sheet };
      else if (ref !== undefined && sheet !== undefined) cache.source = { kind: 'range', sheet, ref };
      else refuse('its source is not stated');
    }
  }
  if (source && child(source, 'consolidation')) refuse('its data is a consolidation of several ranges');
  for (const local of OLAP_ELEMENTS) if (child(root, local)) refuse('its data comes from an OLAP source');
  const fields = child(root, 'cacheFields');
  const list = fields ? childElements(fields, 'cacheField') : [];
  if (list.length > MAX_PIVOT_FIELDS) throw new Unreadable('the cache has too many fields');
  for (const element of list) {
    const field: PivotCacheField = { name: decodeXstring(element.attrs.get('name') || '') };
    if (element.attrs.get('formula') !== undefined || element.attrs.get('databaseField') === '0') refuse('it has a calculated field');
    if (child(element, 'fieldGroup')) refuse('it groups items (dates or numbers in ranges, or items put together by hand)');
    const shared = child(element, 'sharedItems');
    if (shared) {
      const kids = childElements(shared);
      if (kids.length > 0) {
        spend(budget, kids.length);
        const items: PivotValue[] = [];
        for (const kid of kids) {
          const value = cacheValue(kid, date1904);
          if (value) items.push(value);
        }
        field.items = items;
      }
    }
    cache.fields.push(field);
  }
  if (cache.fields.length === 0) refuse('its fields cannot be read');
  return cache;
}

/** Attributes of <pivotField> that ask for subtotals other than the one by the summary function of the values. */
const SUBTOTAL_FLAGS = ['sumSubtotal', 'countASubtotal', 'avgSubtotal', 'maxSubtotal', 'minSubtotal', 'productSubtotal', 'countSubtotal', 'stdDevSubtotal', 'stdDevPSubtotal', 'varSubtotal', 'varPSubtotal'];

const AXES: PivotAxis[] = ['axisRow', 'axisCol', 'axisPage', 'axisValues'];

/** Fields a table may have in its rows and columns together to be refreshed: each is a level of the lines it is built of. */
const MAX_AXIS_FIELDS = 64;

/** Whether an element or anything inside it has the attribute (by local name) set to true. */
function anyTrue(el: XmlElement, local: string, depth = 0): boolean {
  if (depth > 32) return false;
  for (const [name, value] of el.attrs) if (localName(name) === local && bool(value, false)) return true;
  for (const kid of childElements(el)) if (anyTrue(kid, local, depth + 1)) return true;
  return false;
}

/** Whether a list of formats (formats, conditional formats, chart formats) names items of fields. */
function namesItems(el: XmlElement, depth = 0): boolean {
  if (depth > 32) return true;
  for (const kid of childElements(el)) {
    if (localName(kid.name) === 'reference' && childElements(kid, 'x').length > 0) return true;
    if (namesItems(kid, depth + 1)) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------
// Settings
//
// What a table shows depends on its settings. A refresh builds what it knows the meaning of; a
// setting it does not apply is allowed only with the value that changes nothing (its default),
// and an attribute or an element it does not know at all is refused - except what the part says
// a reader may pass over (mc:Ignorable). ANY: the value does not matter here (applied, or about
// something else than the cells: styles, menus, printing, OLAP-only features of a range source).

const ANY = null;
type Rules = Record<string, string | null>;

const TABLE_ATTRS: Rules = {
  name: ANY, cacheId: ANY, dataOnRows: ANY, dataPosition: ANY, autoFormatId: ANY, applyNumberFormats: ANY, applyBorderFormats: ANY, applyFontFormats: ANY,
  applyPatternFormats: ANY, applyAlignmentFormats: ANY, applyWidthHeightFormats: ANY, dataCaption: ANY, grandTotalCaption: ANY, errorCaption: ANY, showError: ANY,
  missingCaption: ANY, showMissing: ANY, pageStyle: ANY, pivotTableStyle: ANY, vacatedStyle: ANY, tag: ANY, updatedVersion: ANY, minRefreshableVersion: ANY,
  asteriskTotals: ANY, showItems: ANY, editData: ANY, disableFieldList: ANY, showCalcMbrs: ANY, showDrill: ANY, printDrill: ANY, showMemberPropertyTips: ANY,
  showDataTips: ANY, enableWizard: ANY, enableDrill: ANY, enableFieldProperties: ANY, preserveFormatting: ANY, useAutoFormatting: ANY, pageWrap: ANY,
  pageOverThenDown: ANY, rowGrandTotals: ANY, colGrandTotals: ANY, fieldPrintTitles: ANY, itemPrintTitles: ANY, mergeItem: ANY, showDropZones: ANY,
  createdVersion: ANY, indent: ANY, showHeaders: ANY, compact: ANY, outline: ANY, outlineData: ANY, compactData: ANY, published: ANY, gridDropZones: ANY,
  immersive: ANY, multipleFieldFilters: ANY, chartFormat: ANY, rowHeaderCaption: ANY, colHeaderCaption: ANY, fieldListSortAscending: ANY, mdxSubqueries: ANY,
  customListSort: ANY, showDataDropDown: ANY, showMultipleLabel: ANY,
  visualTotals: '1', subtotalHiddenItems: '0', showEmptyRow: '0', showEmptyCol: '0',
};
const TABLE_CHILDREN = ['location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'formats', 'conditionalFormats', 'chartFormats', 'pivotHierarchies', 'pivotTableStyleInfo', 'filters', 'rowHierarchiesUsage', 'colHierarchiesUsage', 'extLst', 'calculatedItems', 'calculatedMembers'];
const FIELD_ATTRS: Rules = {
  name: ANY, axis: ANY, dataField: ANY, showDropDowns: ANY, compact: ANY, outline: ANY, subtotalTop: ANY, dragToRow: ANY, dragToCol: ANY, dragToPage: ANY,
  dragToData: ANY, dragOff: ANY, showAll: ANY, insertPageBreak: ANY, sortType: ANY, defaultSubtotal: ANY, itemPageCount: ANY, includeNewItemsInFilter: ANY,
  multipleItemSelectionAllowed: ANY, hideNewItems: ANY, numFmtId: ANY, nonAutoSortDefault: ANY, showPropCell: ANY, showPropTip: ANY, showPropAsCaption: ANY,
  defaultAttributeDrillState: ANY, hiddenLevel: ANY, uniqueMemberProperty: ANY, allDrilled: ANY, rankBy: ANY, topAutoShow: ANY, insertBlankRow: ANY,
  sumSubtotal: ANY, countASubtotal: ANY, avgSubtotal: ANY, maxSubtotal: ANY, minSubtotal: ANY, productSubtotal: ANY, countSubtotal: ANY, stdDevSubtotal: ANY,
  stdDevPSubtotal: ANY, varSubtotal: ANY, varPSubtotal: ANY,
  autoShow: '0', dataSourceSort: '0', measureFilter: '0', serverField: '0',
};
const ITEM_ATTRS: Rules = { n: ANY, t: ANY, h: ANY, s: ANY, c: ANY, x: ANY, d: ANY, e: ANY, f: ANY, m: ANY, sd: '1' };
const DATA_ATTRS: Rules = { name: ANY, fld: ANY, subtotal: ANY, showDataAs: ANY, baseField: ANY, baseItem: ANY, numFmtId: ANY };
const PAGE_ATTRS: Rules = { fld: ANY, item: ANY, hier: ANY, name: ANY, cap: ANY };
/** Attributes of the extensions of a table (x14:pivotTableDefinition), of a field (x14:pivotField) and of a value (x14:dataField). */
const TABLE_EXT_ATTRS: Rules = {
  hideValuesRow: ANY, altText: ANY, altTextSummary: ANY, enableEdit: ANY, autoApply: ANY, allocationMethod: ANY, allocationValue: ANY, allocationWeightExpression: ANY,
  weightExpression: ANY, visualTotalsForSets: ANY, calculatedMembersInFilters: ANY, fillDownLabelsDefault: '0',
};
const FIELD_EXT_ATTRS: Rules = { fillDownLabels: '0', ignore: '0' };
const DATA_EXT_ATTRS: Rules = { sourceField: ANY, uniqueName: ANY, pivotShowAs: 'normal' };
/** Elements of extensions that say nothing about the cells. */
const HARMLESS_EXTENSIONS = ['pivotTableUISettings', 'pivotTableServerFormats'];

/** The prefixes a part says a reader may pass over (mc:Ignorable of its root). */
function ignorablePrefixes(root: XmlElement): Set<string> {
  const out = new Set<string>();
  for (const [name, value] of root.attrs) if (localName(name) === 'Ignorable') for (const prefix of value.split(/\s+/)) if (prefix) out.add(prefix);
  return out;
}

/** The first setting of an element a refresh does not apply, as a reason; undefined when there is none. */
function unknownSetting(el: XmlElement, rules: Rules, passOver: Set<string>): string | undefined {
  for (const [name, value] of el.attrs) {
    if (name === 'xmlns' || name.slice(0, 6) === 'xmlns:') continue;
    const colon = name.indexOf(':');
    if (colon >= 0) {
      const prefix = name.slice(0, colon);
      if (passOver.has(prefix) || localName(name) === 'Ignorable' || localName(name) === 'id') continue;
      return `it has a setting that is not known here (${name.slice(0, 40)})`;
    }
    if (!Object.prototype.hasOwnProperty.call(rules, name)) return `it has a setting that is not known here (${name.slice(0, 40)})`;
    const wanted = rules[name];
    if (wanted === null) continue;
    const plain = value === 'true' ? '1' : value === 'false' ? '0' : value;
    if (plain !== wanted) return `it has a setting that is not applied here (${name}="${value.slice(0, 20)}")`;
  }
  return undefined;
}

/** What the extension list of an element holds that a refresh does not apply; undefined when nothing. */
function unknownExtension(el: XmlElement, own: string, rules: Rules, passOver: Set<string>): string | undefined {
  const list = childElements(el, 'extLst')[0];
  if (!list) return undefined;
  for (const ext of childElements(list, 'ext')) {
    for (const kid of childElements(ext)) {
      const local = localName(kid.name);
      if (HARMLESS_EXTENSIONS.indexOf(local) >= 0) continue;
      // Excel 365's defaults of the layout of fields that are put into the table from now on ([MS-XLSX]
      // pivotTableDefinition16): a refresh puts in no field, so they change nothing it builds. Only such defaults.
      if (local === 'pivotTableDefinition16' && own === 'pivotTableDefinition') {
        for (const [name] of kid.attrs) {
          if (name === 'xmlns' || name.slice(0, 6) === 'xmlns:' || /Default$/.test(name)) continue;
          return `it has a setting that is not known here (${name.slice(0, 40)})`;
        }
        if (childElements(kid).length > 0) return `it has an extension that is not known here (${local})`;
        continue;
      }
      if (local !== own) return `it has an extension that is not known here (${local.slice(0, 40)})`;
      const problem = unknownSetting(kid, rules, passOver);
      if (problem) return problem;
      if (childElements(kid).length > 0) return `it has an extension that is not known here (${local.slice(0, 40)})`;
    }
  }
  return undefined;
}

/**
 * The lines of rows or columns. A line names at most as many items as its direction has fields
 * (`depth`; a grand total names one): a file that names more is not read, whatever it costs.
 */
function parseLines(el: XmlElement | undefined, budget: Budget, depth: number): PivotLine[] {
  if (!el) return [];
  const list = childElements(el, 'i');
  spend(budget, list.length);
  const most = Math.max(1, depth);
  let previous = 0;
  return list.map((i) => {
    // The items a line names count as well: one line may name many.
    spend(budget, i.children.length);
    const line: PivotLine = { r: int(i.attrs.get('r'), 0, MAX_PIVOT_FIELDS) || 0, i: int(i.attrs.get('i'), 0, MAX_PIVOT_FIELDS) || 0, x: [] };
    const t = i.attrs.get('t');
    if (t !== undefined && t !== 'data') line.t = t;
    for (const x of childElements(i, 'x')) line.x.push(int(x.attrs.get('v'), 0, 1048576) || 0);
    if (line.r > previous || line.r + line.x.length > most) throw new Unreadable('a line of the table names more fields than it has');
    previous = line.r + line.x.length;
    return line;
  });
}

/** The words Excel uses for the function a value is summarized by, in its default names ("Sum of Sales"). */
const AGGREGATE_WORDS: Record<PivotAggregate, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
  product: 'Product',
  countNums: 'Count',
  stdDev: 'StdDev',
  stdDevp: 'StdDevp',
  var: 'Var',
  varp: 'Varp',
};

/** Reads a table definition. Throws Unreadable for a part whose content is not one. */
function parseTable(root: XmlElement, part: string, sheet: Sheet, cache: PivotCache, budget: Budget): PivotTable {
  if (localName(root.name) !== 'pivotTableDefinition') throw new Unreadable('the part is not a pivot table definition');
  const locationEl = child(root, 'location');
  const ref = locationEl ? parseRange((locationEl.attrs.get('ref') || '').replace(/\$/g, '')) : undefined;
  if (!locationEl || !ref) throw new Unreadable('the table does not say where it stands');
  const location: PivotLocation = {
    ref,
    firstHeaderRow: int(locationEl.attrs.get('firstHeaderRow'), 0, 1048576) || 0,
    firstDataRow: int(locationEl.attrs.get('firstDataRow'), 0, 1048576) || 0,
    firstDataCol: int(locationEl.attrs.get('firstDataCol'), 0, 16384) || 0,
  };
  const table: PivotTable = {
    name: root.attrs.get('name') || '',
    part,
    sheet,
    cache,
    location,
    fields: [],
    rowFields: [],
    colFields: [],
    pageFields: [],
    dataFields: [],
    rowItems: [],
    colItems: [],
    rowGrandTotals: bool(root.attrs.get('rowGrandTotals'), true),
    colGrandTotals: bool(root.attrs.get('colGrandTotals'), true),
    captions: {},
    customListSort: true,
  };
  const refuse = (reason: string) => {
    if (table.refusal === undefined) table.refusal = reason;
  };
  const captions: [keyof PivotTable['captions'], string][] = [
    ['rowHeader', 'rowHeaderCaption'],
    ['colHeader', 'colHeaderCaption'],
    ['grandTotal', 'grandTotalCaption'],
    ['data', 'dataCaption'],
  ];
  for (const [key, attr] of captions) {
    const value = root.attrs.get(attr);
    if (value !== undefined) table.captions[key] = value;
  }
  const passOver = ignorablePrefixes(root);
  const setting = (problem: string | undefined) => {
    if (problem !== undefined) refuse(problem);
  };
  setting(unknownSetting(root, TABLE_ATTRS, passOver));
  for (const kid of childElements(root)) if (TABLE_CHILDREN.indexOf(localName(kid.name)) < 0) refuse(`it has a part that is not known here (${localName(kid.name).slice(0, 40)})`);
  setting(unknownExtension(root, 'pivotTableDefinition', TABLE_EXT_ATTRS, passOver));
  if (child(root, 'rowHierarchiesUsage') || child(root, 'colHierarchiesUsage')) refuse('its data comes from an OLAP source');
  table.customListSort = bool(root.attrs.get('customListSort'), true);
  const missing = root.attrs.get('missingCaption');
  if (bool(root.attrs.get('showMissing'), true) && missing !== undefined) table.missingCaption = missing;
  if (bool(root.attrs.get('showError'), false)) table.errorCaption = root.attrs.get('errorCaption') ?? '';
  if (bool(root.attrs.get('dataOnRows'), false)) refuse('its values stand in rows');
  if (!bool(root.attrs.get('showHeaders'), true)) refuse('it shows no headers');
  if (bool(root.attrs.get('gridDropZones'), false)) refuse('it has the classic layout');
  if (bool(root.attrs.get('mergeItem'), false)) refuse('it merges the cells of its labels');

  const fieldsEl = child(root, 'pivotFields');
  const fieldList = fieldsEl ? childElements(fieldsEl, 'pivotField') : [];
  if (fieldList.length > MAX_PIVOT_FIELDS) throw new Unreadable('the table has too many fields');
  for (const el of fieldList) {
    const axisAttr = el.attrs.get('axis');
    const sort = el.attrs.get('sortType');
    const field: PivotField = {
      compact: bool(el.attrs.get('compact'), true),
      outline: bool(el.attrs.get('outline'), true),
      subtotalTop: bool(el.attrs.get('subtotalTop'), true),
      defaultSubtotal: bool(el.attrs.get('defaultSubtotal'), true),
      showAll: bool(el.attrs.get('showAll'), true),
      sortType: sort === 'ascending' || sort === 'descending' ? sort : 'manual',
    };
    if (axisAttr !== undefined && (AXES as string[]).indexOf(axisAttr) >= 0) field.axis = axisAttr as PivotAxis;
    const name = el.attrs.get('name');
    if (name !== undefined) field.name = name;
    if (sort !== undefined && sort !== 'manual' && sort !== 'ascending' && sort !== 'descending') refuse('a field is sorted in a way that is not known here');
    if (SUBTOTAL_FLAGS.some((flag) => bool(el.attrs.get(flag), false))) refuse('a field has subtotals by functions of their own');
    if (bool(el.attrs.get('insertBlankRow'), false)) refuse('a field puts blank rows after its items');
    if (child(el, 'autoSortScope')) refuse('a field is sorted by values');
    if (anyTrue(el, 'fillDownLabels')) refuse('a field repeats the labels of its items');
    if (el.attrs.get('subtotalCaption') !== undefined) refuse('a field has a caption of its own for its subtotals (subtotalCaption)');
    else setting(unknownSetting(el, FIELD_ATTRS, passOver));
    for (const kid of childElements(el)) if (['items', 'autoSortScope', 'extLst'].indexOf(localName(kid.name)) < 0) refuse(`a field has a part that is not known here (${localName(kid.name).slice(0, 40)})`);
    setting(unknownExtension(el, 'pivotField', FIELD_EXT_ATTRS, passOver));
    if (bool(el.attrs.get('hideNewItems'), false)) field.hideNewItems = true;
    const labelFormat = int(el.attrs.get('numFmtId'), 0, 65535);
    if (labelFormat !== undefined && labelFormat !== 0) field.numFmtId = labelFormat;
    const itemsEl = child(el, 'items');
    if (itemsEl) {
      const kids = childElements(itemsEl, 'item');
      spend(budget, kids.length);
      field.items = kids.map((kid) => {
        const item: PivotItem = { attrs: [...kid.attrs] };
        const x = int(kid.attrs.get('x'), 0, 1048576);
        if (x !== undefined) item.x = x;
        const t = kid.attrs.get('t');
        if (t !== undefined && t !== 'data') item.t = t;
        if (bool(kid.attrs.get('h'), false)) item.h = true;
        if (bool(kid.attrs.get('f'), false)) refuse('it has calculated items');
        if (bool(kid.attrs.get('m'), false)) refuse('an item of a field is missing from its source');
        if (!bool(kid.attrs.get('sd'), true)) refuse('an item of a field is collapsed (its details are hidden)');
        if (t !== undefined && t !== 'data' && t !== 'default') refuse(`a field has an item of a kind that is not built here ("${t.slice(0, 20)}")`);
        setting(unknownSetting(kid, ITEM_ATTRS, passOver));
        return item;
      });
    }
    table.fields.push(field);
  }
  const fieldsOf = (el: XmlElement | undefined): number[] =>
    el ? childElements(el, 'field').map((f) => int(f.attrs.get('x'), -2, MAX_PIVOT_FIELDS) ?? -3) : [];
  table.rowFields = fieldsOf(child(root, 'rowFields'));
  table.colFields = fieldsOf(child(root, 'colFields'));
  spend(budget, table.rowFields.length + table.colFields.length);
  const named = new Set<number>();
  for (const x of table.rowFields.concat(table.colFields)) {
    if (x === -3 || x === -1 || (x >= 0 && x >= table.fields.length)) throw new Unreadable('the table names a field it does not have');
    // A field stands in one place: in the rows or in the columns, once (the values, -2, as well).
    if (named.has(x)) throw new Unreadable('the table names a field twice in its rows and columns');
    named.add(x);
  }
  if (table.rowFields.length + table.colFields.length > MAX_AXIS_FIELDS) refuse(`it has more than ${MAX_AXIS_FIELDS} fields in its rows and columns`);
  table.rowItems = parseLines(child(root, 'rowItems'), budget, table.rowFields.length);
  table.colItems = parseLines(child(root, 'colItems'), budget, table.colFields.length);
  const pagesEl = child(root, 'pageFields');
  for (const el of pagesEl ? childElements(pagesEl, 'pageField') : []) {
    const fld = int(el.attrs.get('fld'), 0, table.fields.length - 1);
    if (fld === undefined) throw new Unreadable('a page field names a field the table does not have');
    if (named.has(fld)) throw new Unreadable('a page field stands in the rows or columns as well');
    named.add(fld);
    setting(unknownSetting(el, PAGE_ATTRS, passOver));
    const page: PivotPageField = { fld };
    const item = int(el.attrs.get('item'), 0, 1048576);
    if (item !== undefined) page.item = item;
    table.pageFields.push(page);
  }
  const dataEl = child(root, 'dataFields');
  for (const el of dataEl ? childElements(dataEl, 'dataField') : []) {
    const fld = int(el.attrs.get('fld'), 0, table.fields.length - 1);
    if (fld === undefined) throw new Unreadable('a value names a field the table does not have');
    const subtotal = el.attrs.get('subtotal') || 'sum';
    const known = (PIVOT_AGGREGATES as readonly string[]).indexOf(subtotal) >= 0;
    if (!known) refuse(`a value is summarized by a function that is not known here ("${subtotal.slice(0, 20)}")`);
    setting(unknownSetting(el, DATA_ATTRS, passOver));
    setting(unknownExtension(el, 'dataField', DATA_EXT_ATTRS, passOver));
    for (const kid of childElements(el)) if (localName(kid.name) !== 'extLst') refuse(`a value has a part that is not known here (${localName(kid.name).slice(0, 40)})`);
    const showAs = el.attrs.get('showDataAs');
    if (showAs !== undefined && showAs !== 'normal') refuse('a value is shown as a calculation (a share, a difference, a running total, ...)');
    const aggregate = (known ? subtotal : 'sum') as PivotAggregate;
    const fieldName = cache.fields[fld] ? cache.fields[fld].name : '';
    const data: PivotDataField = { name: el.attrs.get('name') || `${AGGREGATE_WORDS[aggregate]} of ${fieldName}`, fld, subtotal: aggregate };
    const numFmtId = int(el.attrs.get('numFmtId'), 0, 65535);
    if (numFmtId !== undefined) data.numFmtId = numFmtId;
    table.dataFields.push(data);
  }
  for (const local of ['formats', 'conditionalFormats']) {
    const el = child(root, local);
    if (el && namesItems(el)) refuse('it formats some of its items');
  }
  if (child(root, 'chartFormats')) refuse('a pivot chart is made from it');
  if (child(root, 'filters')) refuse('it filters by labels or by values');
  if (child(root, 'calculatedItems') || child(root, 'calculatedMembers')) refuse('it has calculated items');
  if (child(root, 'pivotHierarchies')) refuse('its data comes from an OLAP source');
  if (table.fields.length !== cache.fields.length) refuse('its fields are not those of its cache');
  return table;
}

/** The pivot tables a sheet shows, by the relationships of its part. */
export function pivotTableParts(sheetRels: Relationship[]): string[] {
  const out: string[] = [];
  for (const rel of sheetRels) if (rel.type.slice(rel.type.lastIndexOf('/') + 1) === 'pivotTable' && !rel.external && rel.part && out.indexOf(rel.part) < 0) out.push(rel.part);
  return out;
}

/**
 * Reads the pivot tables of the sheets and their caches. `tables` are the table parts of each
 * sheet; `cacheIds` the caches the workbook part lists, by the part each is in. A table or cache
 * that cannot be read is left out: it stays in the package as it is, and is not refreshed.
 * Returns undefined for a workbook without a table that could be read.
 */
export async function readPivots(
  pkg: OpcPackage,
  tables: { sheet: Sheet; parts: string[] }[],
  cacheIds: Map<string, number>,
  date1904: boolean,
  main = '/xl/workbook.xml',
  slicers = false,
): Promise<WorkbookPivots | undefined> {
  const budget: Budget = { items: 0 };
  const caches = new Map<string, PivotCache | null>();
  const out: WorkbookPivots = { caches: [], tables: [] };
  const read = async (part: string): Promise<XmlElement> => {
    const bytes = await pkg.readPart(part);
    if (bytes.length > MAX_PIVOT_PART_BYTES) throw new Unreadable('the part is too large');
    return parseXml(bytes).root;
  };
  const cacheOf = async (part: string): Promise<PivotCache | null> => {
    const key = part.toLowerCase();
    const known = caches.get(key);
    if (known !== undefined) return known;
    let cache: PivotCache | null = null;
    try {
      const rels = await pkg.relationships(part);
      const records = rels.find((rel) => rel.type.slice(rel.type.lastIndexOf('/') + 1) === 'pivotCacheRecords' && !rel.external && rel.part && pkg.hasPart(rel.part));
      const id = cacheIds.get(key);
      cache = parseCache(await read(part), id === undefined ? -1 : id, part, records && records.part ? records.part : undefined, date1904, budget);
      if (id === undefined && cache.refusal === undefined) cache.refusal = 'the workbook does not list its cache';
      out.caches.push(cache);
    } catch (error) {
      cache = null;
    }
    caches.set(key, cache);
    return cache;
  };
  every: for (const entry of tables) {
    for (const part of entry.parts) {
      // The tables beyond are left out; their parts refer to their caches, which are then not refreshed (lookThrough()).
      if (out.tables.length >= MAX_PIVOT_TABLES) break every;
      try {
        if (!pkg.hasPart(part)) continue;
        const rels = await pkg.relationships(part);
        const cacheRel = rels.find((rel) => rel.type.slice(rel.type.lastIndexOf('/') + 1) === 'pivotCacheDefinition' && !rel.external && rel.part && pkg.hasPart(rel.part));
        if (!cacheRel || !cacheRel.part) continue;
        const cache = await cacheOf(cacheRel.part);
        if (!cache) continue;
        const root = await read(part);
        const table = parseTable(root, part, entry.sheet, cache, budget);
        const id = int(root.attrs.get('cacheId'), 0, 4294967295);
        if (id !== cache.cacheId && table.refusal === undefined) table.refusal = 'it names a cache the workbook does not have under that number';
        // A part two sheets name is one table: it is refreshed for neither.
        if (out.tables.some((other) => other.part.toLowerCase() === part.toLowerCase())) {
          for (const other of out.tables) if (other.part.toLowerCase() === part.toLowerCase() && other.refusal === undefined) other.refusal = 'two sheets show the same table';
          continue;
        }
        out.tables.push(table);
      } catch (error) {
        // The table stays as the file has it, and is not refreshed.
        if (!(error instanceof Unreadable) && !(error instanceof Error)) throw error;
      }
    }
  }
  if (out.tables.length === 0) return undefined;
  // What else refers to the caches is looked through when a refresh is asked for: opening costs nothing more.
  out.lookThrough = (workbook: Workbook) => lookThrough(pkg, out, main, slicers, workbook);
  return out;
}

/** What cannot be looked through: every cache is refused with this, since what refers to it is not known. */
class CannotTell extends Error {}

const UNKNOWN = 'it cannot tell what refers to its cache';

/**
 * The bytes of a part as the package has them now, read at once (without waiting) and only when
 * the archive says they are no more than `limit`: what does not fit is not inflated at all.
 */
function partNow(pkg: OpcPackage, names: Map<string, string>, part: string, limit: number): Uint8Array | undefined {
  const name = names.get(nameKey(part));
  if (name === undefined) return undefined;
  // The archive names its entries without the slash a part name starts with.
  const entry = pkg.zip.entry(name.charAt(0) === '/' ? name.slice(1) : name);
  if (!entry) return undefined;
  if (entry.data) {
    if (entry.data.length > limit) throw new CannotTell(`${part} is too large to look through`);
    return entry.data;
  }
  if (entry.uncompressedSize > limit) throw new CannotTell(`${part} is too large to look through`);
  const data = entry.method === 0 ? entry.raw : inflateRaw(entry.raw, entry.uncompressedSize);
  if (data.length !== entry.uncompressedSize) throw new CannotTell(`${part} cannot be read`);
  return data;
}

/**
 * One key for a part name however it is spelled: without regard to case (as OPC compares them)
 * and with what is written with percent signs decoded - on both sides, the part that is named and
 * the name a relationship gives. A name with a percent sign that cannot be decoded is its own key.
 */
function nameKey(name: string): string {
  let plain = name;
  if (name.indexOf('%') >= 0) {
    try {
      plain = decodeURIComponent(name);
    } catch (error) {
      plain = name;
    }
  }
  return plain.toLowerCase();
}
/** The part a relationships part holds the relationships of: "/xl/_rels/workbook.xml.rels" -> "/xl/workbook.xml". */
function sourceOfRels(name: string): string | undefined {
  const match = /^(.*)\/_rels\/([^/]+)\.rels$/i.exec(name);
  if (!match) return undefined;
  return match[1] === '' && match[2] === '' ? '/' : `${match[1]}/${match[2]}`;
}

/**
 * The tables the pivot sources of a chart part name (c:pivotSource/c:name), read in one pass over
 * the whole part, wherever they stand (after the chart, in alternate content, ...); '' for a
 * source that names none. Empty for a chart that is no pivot chart.
 */
function pivotSourcesOf(bytes: Uint8Array): string[] {
  const scanner = new XmlScanner(bytes);
  const out: string[] = [];
  let inSource = false;
  let named = false;
  while (scanner.next()) {
    const kind: string = scanner.kind;
    const local: string = scanner.local;
    if (kind === 'close' && local === 'pivotSource') {
      if (!named) out.push('');
      inSource = false;
      continue;
    }
    if (kind !== 'open') continue;
    if (local === 'pivotSource') {
      if (scanner.selfClosing) out.push('');
      else {
        inSource = true;
        named = false;
      }
    } else if (inSource && local === 'name' && !named) {
      named = true;
      out.push(scanner.readText());
    }
  }
  return out;
}

/** The table a pivot chart names, among the tables of the model; undefined when none is found. */
function tableNamed(pivots: WorkbookPivots, source: string): PivotTable | undefined {
  const bang = source.lastIndexOf('!');
  const tableName = source.slice(bang + 1).toLowerCase();
  const sheetName = (bang < 0 ? '' : source.slice(0, bang)).replace(/^\[[^\]]*\]/, '').replace(/^'|'$/g, '').replace(/''/g, "'").toLowerCase();
  return pivots.tables.find((table) => table.name.toLowerCase() === tableName && table.sheet.name.toLowerCase() === sheetName);
}

/**
 * A refresh rebuilds the items of a cache, which other parts name by their place: a table the
 * model could not take in, a slicer or a timeline, a pivot chart. A cache such a part refers to
 * is not refreshed: what refers to it is said on it (`referred`, found anew each time). This is
 * done when a refresh is asked for, not when the file is opened, and at once: the parts are read
 * from the archive without waiting, within the limits of PIVOT_SCAN_LIMITS; what goes past them
 * refuses every cache.
 *
 * What is looked at is the workbook as it is now: the charts of the model (Chart.pivotSource: a
 * chart that is not in the package any more, or was not yet, counts as well) and the parts of the
 * package as it is now. Every relationship of the package is looked at; a part the model holds no
 * table for that refers to a cache (or to its records) marks it - part names compared without
 * regard to case and with percent signs decoded, on both sides. Slicers and timelines name a
 * cache by an id of its extension, not by a relationship: while the workbook has any, no cache is
 * refreshed. A chart of the package the model does not show (a chart sheet it could not read, a
 * part found by its type, its place or the relationship that leads to it) that is a pivot chart
 * marks the cache of the table it names; every cache, for a name that is not found.
 */
function lookThrough(pkg: OpcPackage, pivots: WorkbookPivots, main: string, slicers: boolean, workbook: Workbook): void {
  for (const cache of pivots.caches) delete cache.referred;
  const mark = (cache: PivotCache, reason: string) => {
    if (cache.referred === undefined) cache.referred = reason;
  };
  const markAll = (reason: string) => {
    for (const cache of pivots.caches) mark(cache, reason);
  };
  const markSource = (source: string) => {
    const found = source === '' ? undefined : tableNamed(pivots, source);
    if (found) mark(found.cache, 'a pivot chart is made from a table of its cache');
    else markAll('a pivot chart is made from a table that is not known here');
  };
  const limits = PIVOT_SCAN_LIMITS;
  try {
    if (slicers) markAll('a slicer or a timeline reads the caches of the workbook');
    // The charts of the model, as they are now.
    const shown = new Set<string>();
    for (const sheet of workbook.sheets) {
      for (const drawing of sheet.drawings || []) {
        const object = drawing.object;
        if (object.type !== 'chart') continue;
        if (object.part !== undefined) shown.add(nameKey(object.part));
        if (object.chart.pivotSource !== undefined) markSource(object.chart.pivotSource);
      }
    }
    const names = new Map<string, string>();
    for (const name of pkg.partNames()) if (!names.has(nameKey(name))) names.set(nameKey(name), name);
    const tables = new Set(pivots.tables.map((table) => nameKey(table.part)));
    const byPart = new Map<string, PivotCache>();
    const byRecords = new Map<string, PivotCache>();
    for (const cache of pivots.caches) {
      byPart.set(nameKey(cache.part), cache);
      if (cache.recordsPart !== undefined) byRecords.set(nameKey(cache.recordsPart), cache);
    }
    const mainKey = nameKey(main);
    const charts = new Set<string>();
    let rels = 0;
    let relsBytes = 0;
    for (const name of pkg.partNames()) {
      const lower = name.toLowerCase();
      const type = (pkg.contentTypeOf(name) || '').toLowerCase();
      if (/slicer|timeline/.test(type) || /^\/xl\/(slicers|slicercaches|timelines|timelinecaches)\//.test(lower)) markAll('a slicer or a timeline reads the caches of the workbook');
      if (/drawingml\.chart\+xml$/.test(type) || /^\/xl\/charts\/chart[^/]*\.xml$/.test(lower)) charts.add(nameKey(name));
      const source = sourceOfRels(name);
      if (source === undefined) continue;
      if (++rels > limits.relsParts) throw new CannotTell('the workbook has too many relationship parts');
      const data = partNow(pkg, names, name, limits.relsPartBytes);
      if (!data) continue;
      relsBytes += data.length;
      if (relsBytes > limits.relsBytes || relsBytes > limits.totalBytes) throw new CannotTell('the workbook has too many relationships');
      let list: Relationship[];
      try {
        list = relationshipsOf(source, parseXml(data));
      } catch (error) {
        throw new CannotTell(`the relationships of ${source} cannot be read`);
      }
      const from = nameKey(source);
      for (const rel of list) {
        // A link (to a place of the workbook, "#Sheet2!A1", or elsewhere) leads to no part.
        if (rel.external || !rel.part || /\/hyperlink$/.test(rel.type)) continue;
        // A target written with percent signs that cannot be decoded, or that holds what no part name holds (a query, a
        // fragment, a backslash, a control character, spaces at an end - which some readers drop): what it leads to is not known.
        let plain = rel.part;
        if (plain.indexOf('%') >= 0) {
          try {
            plain = decodeURIComponent(plain);
          } catch (error) {
            throw new CannotTell('a relationship leads to a part whose name cannot be read');
          }
        }
        if (/[?#\\\u0000-\u001f]/.test(plain) || plain !== plain.trim()) throw new CannotTell('a relationship leads to a part whose name cannot be read');
        const key = nameKey(rel.part);
        const kind = rel.type.slice(rel.type.lastIndexOf('/') + 1);
        if (kind === 'chart' || kind === 'chartEx') charts.add(key);
        const cache = byPart.get(key);
        if (cache && from !== mainKey && !tables.has(from)) mark(cache, `another part of the workbook refers to its cache (${source.slice(0, 60)})`);
        const owner = byRecords.get(key);
        if (owner && from !== nameKey(owner.part)) mark(owner, `another part of the workbook refers to its records (${source.slice(0, 60)})`);
      }
    }
    let chartBytes = 0;
    for (const chart of charts) {
      // A chart the model shows is judged by the model (above): it may have been changed or taken out since.
      if (shown.has(chart)) continue;
      const data = partNow(pkg, names, chart, limits.chartPartBytes);
      if (!data) continue;
      chartBytes += data.length;
      if (chartBytes > limits.chartBytes || relsBytes + chartBytes > limits.totalBytes) throw new CannotTell('the charts of the workbook are too large');
      let sources: string[];
      try {
        sources = pivotSourcesOf(data);
      } catch (error) {
        throw new CannotTell(`${chart} cannot be read`);
      }
      for (const source of sources) markSource(source);
    }
  } catch (error) {
    if (!(error instanceof CannotTell) && !(error instanceof Error)) throw error;
    markAll(UNKNOWN);
  }
}
