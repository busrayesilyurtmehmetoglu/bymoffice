// Pivot tables after their second review: the parts that may refer to a cache are looked
// through only when a refresh is asked for, in bounded linear time; the order of text is that
// of a Turkish collation on three levels; settings and limits each have a test of their own.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { PIVOT_LIMITS, compareText, refreshPivot } from '../src/ops/pivots';
import { tableState } from '../src/model/sheet/pivot';
import { grid, hasPython, independentPivots, load, pivotBook, save, setData, type Datum, type PivotSpec } from './pivot-fixtures';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
  ['East', 'A', 1],
];
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function edit(bytes: Uint8Array, change: (zip: ZipArchive, types: (xml: string) => void) => void | Promise<void>): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  let types = decoder.decode(await zip.read('[Content_Types].xml'));
  await change(zip, (xml) => (types = types.replace('</Types>', `${xml}</Types>`)));
  zip.set('[Content_Types].xml', encoder.encode(types));
  return zip.write();
}

async function patchTable(bytes: Uint8Array, change: (s: string) => string): Promise<Uint8Array> {
  return edit(bytes, async (zip) => zip.set('xl/pivotTables/pivotTable1.xml', encoder.encode(change(decoder.decode(await zip.read('xl/pivotTables/pivotTable1.xml'))))));
}

const reason = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof RangeError, String(error));
    return (error as Error).message;
  }
  return '';
};

const base = () => pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });

// ---------------------------------------------------------------------------------------------
// E1, E2, Ö3: looked through when a refresh is asked for, linear, bounded

test('opening reads nothing beyond the pivot parts: a chart of many tags opens at once', async () => {
  const bytes = await edit(await base(), (zip) => zip.set('xl/charts/chart1.xml', encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}">${'<c:pivotSource>'.repeat(20000)}</c:chartSpace>`)));
  const started = Date.now();
  const doc = await load(bytes);
  const opened = Date.now() - started;
  assert.equal(doc.wb.pivots!.caches[0].refusal, undefined, 'nothing was looked through yet');
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /pivot chart|cannot tell/);
  assert.ok(Date.now() - started < 1500, `${Date.now() - started} ms (open ${opened} ms)`);
});

test('a chart part too large to look through is not read: the refresh is refused', async () => {
  const big = new Uint8Array(40 * 1024 * 1024).fill(32);
  const bytes = await edit(await base(), (zip) => zip.set('xl/charts/chart1.xml', big));
  const doc = await load(bytes);
  const started = Date.now();
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /cannot tell what refers to its cache/);
  assert.ok(Date.now() - started < 300, `${Date.now() - started} ms`);
});

test('more relationship parts than are looked through: the refresh is refused', async () => {
  const rel = encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="r" Type="x/y" Target="../a.xml"/></Relationships>`);
  const bytes = await edit(await base(), (zip) => {
    for (let i = 0; i < 10001; i++) zip.set(`x/_rels/p${i}.xml.rels`, rel);
  });
  const doc = await load(bytes);
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /cannot tell what refers to its cache/);
});

test('a chart found by the relationship that leads to it, a target written with percent signs, a part that refers to the records', async () => {
  const chart = (name: string) => encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}"><c:pivotSource><c:name>${name}</c:name><c:fmtId val="0"/></c:pivotSource><c:chart/></c:chartSpace>`);
  const rels = (target: string, type: string) => encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="r1" Type="${type}" Target="${target}"/></Relationships>`);
  // A chart at a place of no chart and without a type of its own, found through a drawing.
  const hidden = await edit(await base(), (zip) => {
    zip.set('xl/foo/g.xml', chart('Pivot!PivotTable1'));
    zip.set('xl/foo/d.xml', encoder.encode('<d/>'));
    zip.set('xl/foo/_rels/d.xml.rels', rels('g.xml', `${OFFICE_REL}/chart`));
  });
  const one = await load(hidden);
  assert.match(reason(() => refreshPivot(one.wb, 0)), /pivot chart is made from a table of its cache/);
  // A pivot chart of a table that is not known.
  const other = await load(await edit(await base(), (zip) => zip.set('xl/charts/chart1.xml', chart('Pivot!Other'))));
  assert.match(reason(() => refreshPivot(other.wb, 0)), /pivot chart is made from a table that is not known here/);
  // The cache named with percent signs.
  const percent = await load(
    await edit(await base(), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', rels('../pivotCache/pivotCache%44efinition1.xml', `${OFFICE_REL}/pivotCacheDefinition`));
    }),
  );
  assert.match(reason(() => refreshPivot(percent.wb, 0)), /another part of the workbook refers to its cache/);
  // A target that cannot be decoded.
  const broken = await load(
    await edit(await base(), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', rels('../pivotCache/%E0%A4%A.xml', 'x/y'));
    }),
  );
  assert.match(reason(() => refreshPivot(broken.wb, 0)), /cannot tell what refers to its cache/);
  // The records.
  const records = await load(
    await edit(await base(), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', rels('../pivotCache/pivotCacheRecords1.xml', 'x/y'));
    }),
  );
  assert.match(reason(() => refreshPivot(records.wb, 0)), /refers to its records/);
});

// ---------------------------------------------------------------------------------------------
// Ö1: the order of text

test('text is sorted on three levels: letters, then accents, then case', () => {
  const sorted = (list: string[]) => list.slice().sort(compareText);
  // Written out by hand, as a Turkish collation orders them.
  assert.deepEqual(sorted(['élan', 'Émile', 'eva']), ['élan', 'Émile', 'eva']);
  assert.deepEqual(sorted(['âcil', 'Âdem', 'abla']), ['abla', 'âcil', 'Âdem']);
  assert.deepEqual(sorted(['бета', 'Альфа', 'Вега', 'гамма']), ['Альфа', 'бета', 'Вега', 'гамма']);
  assert.deepEqual(sorted(['ωμέγα', 'Άλφα', 'βήτα']), ['Άλφα', 'βήτα', 'ωμέγα']);
  assert.deepEqual(sorted(['Kâğıt', 'kağıt', 'Kalem']), ['kağıt', 'Kâğıt', 'Kalem']);
  assert.deepEqual(sorted(['Çay', 'cevap', 'Ceviz', 'çam']), ['cevap', 'Ceviz', 'çam', 'Çay']);
  assert.deepEqual(sorted(['Ilık', 'ılık', 'iğne', 'İğne', 'Hayır']), ['Hayır', 'ılık', 'Ilık', 'iğne', 'İğne']);
  // An accent written as a mark of its own weighs as the accent of a precomposed letter; the two forms are then told apart by their code units.
  const decomposed = 'e\u0301lan';
  assert.deepEqual(sorted([decomposed, '\u00e9lan', 'elan', 'elam']), ['elam', 'elan', decomposed, '\u00e9lan']);
  // A total order: the opposite answer when asked the other way round, and one that goes on.
  const alphabet = ['a', 'A', 'é', 'É', 'ı', 'I', 'i', 'İ', 'ç', 'Ç', '1', '-', ' ', 'ß', 'Б', 'б', '€', 'é', '😀', '\ud800', 'ş', 'Ş', 'q', 'w', 'â'];
  let seed = 7;
  const random = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  const word = () => {
    let out = '';
    for (let n = Math.floor(random() * 4); n > 0; n--) out += alphabet[Math.floor(random() * alphabet.length)];
    return out;
  };
  for (let t = 0; t < 50000; t++) {
    const a = word();
    const b = word();
    const c = word();
    const ab = Math.sign(compareText(a, b));
    assert.equal(ab + Math.sign(compareText(b, a)), 0, JSON.stringify([a, b]));
    if (ab <= 0 && Math.sign(compareText(b, c)) <= 0) assert.ok(compareText(a, c) <= 0, JSON.stringify([a, b, c]));
    if (a !== b) assert.notEqual(ab, 0, JSON.stringify([a, b]));
  }
});

// ---------------------------------------------------------------------------------------------
// Ö2, K1, K3, K5: settings and limits

test('an extension that is not known, more than 64 fields in rows and columns, too much to add up, a label format the style sheet lacks', async () => {
  const ext = await load(await patchTable(await base(), (s) => s.replace('</pivotTableDefinition>', '<extLst><ext uri="{0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:somethingNew/></ext></extLst></pivotTableDefinition>')));
  assert.match(reason(() => refreshPivot(ext.wb, 0)), /extension that is not known here \(somethingNew\)/);
  const wide: Datum[][] = [Array.from({ length: 66 }, (_, i) => `F${i}`), Array.from({ length: 66 }, (_, i) => (i === 65 ? 1 : `v${i}`))];
  const many = await load(await pivotBook({ data: wide, rows: Array.from({ length: 65 }, (_, i) => i), values: [{ fld: 65 }] }));
  assert.match(reason(() => refreshPivot(many.wb, 0)), /more than 64 fields in its rows and columns/);
  const work = await load(await base());
  const before = PIVOT_LIMITS.work;
  PIVOT_LIMITS.work = 5;
  try {
    assert.match(reason(() => refreshPivot(work.wb, 0)), /too long to add up/);
  } finally {
    PIVOT_LIMITS.work = before;
  }
  const dates: Datum[][] = [['Day', 'V'], [{ date: 45000 }, 1], [{ date: 45001 }, 2]];
  const format = await load(await patchTable(await pivotBook({ data: dates, rows: [0], values: [{ fld: 1 }] }), (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" numFmtId="22"')));
  assert.match(reason(() => refreshPivot(format.wb, 0)), /number format the style sheet has no style with/);
});

test('the default layout of Excel 365 (pivotTableDefinition16) is passed over; anything else in it is refused', async () => {
  const xpdl = (attrs: string) => `<extLst><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16 ${attrs}/></ext></extLst>`;
  const plain = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], tableTail: xpdl('SubtotalsOnTopDefault="0"') }));
  assert.equal(reason(() => refreshPivot(plain.wb, 0)), '');
  const odd = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], tableTail: xpdl('Unheard="1"') }));
  assert.match(reason(() => refreshPivot(odd.wb, 0)), /Unheard/);
});

test('a field that hides new items and would show none is not refreshed', async () => {
  const doc = await load(await patchTable(await pivotBook({ data: [['Region', 'Sales'], ['East', 10], ['West', 20]], rows: [0], values: [{ fld: 1 }] }), (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" hideNewItems="1"')));
  setData(doc, 1, 0, 'North');
  setData(doc, 2, 0, 'South');
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /every item of the field "Region" would be hidden/);
});

test('names of months and days: a full name among other items is refused as well (the rule of the third review)', async () => {
  const months: Datum[][] = [['M', 'V'], ['Ocak', 1], ['Şubat', 2]];
  const all = await load(await pivotBook({ data: months, rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
  assert.match(reason(() => refreshPivot(all.wb, 0)), /names of months or days/);
  const mixed = await load(await pivotBook({ data: [['M', 'V'], ['Mart', 1], ['Ankara', 2]], rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
  assert.match(reason(() => refreshPivot(mixed.wb, 0)), /names of months or days/);
  const plain = await load(await pivotBook({ data: [['M', 'V'], ['Ankara', 1], ['Bursa', 2]], rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
  assert.equal(reason(() => refreshPivot(plain.wb, 0)), '');
});

// ---------------------------------------------------------------------------------------------
// The independent check knows the captions and the order

test('the independent reader checks captions of missing values and errors, hidden new items and the order of sorted items', { skip: !hasPython() }, async () => {
  const files: { [name: string]: Uint8Array } = {};
  const make = async (spec: PivotSpec, fix: (s: string) => string, change?: (doc: Awaited<ReturnType<typeof load>>) => void) => {
    const doc = await load(await patchTable(await pivotBook(spec), fix));
    if (change) change(doc);
    doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
    return save(doc);
  };
  files.missing = await make({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] }, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition missingCaption="-" '), (doc) => setData(doc, 2, 0, 'North'));
  files.error = await make({ data: DATA, rows: [0], values: [{ fld: 2 }] }, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition showError="1" errorCaption="ERR" '), (doc) =>
    doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 2, cell: { kind: 'error', value: '#N/A', style: 0 } }] }], undefined, undefined),
  );
  files.hide = await make({ data: DATA, rows: [0], values: [{ fld: 2 }] }, (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" hideNewItems="1"'), (doc) => setData(doc, 1, 0, 'North'));
  const words: Datum[][] = [['R', 'V'], ['Kâğıt', 1], ['Kalem', 2], ['kağıt', 3], ['Âdem', 4], ['âcil', 5], ['Çay', 6], ['Ilık', 7], ['iğne', 8], [12, 9]];
  files.up = await make({ data: words, rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }, (s) => s);
  files.down = await make({ data: words, rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'descending' }] }, (s) => s);
  assert.deepEqual(independentPivots(files), []);
  // And it sees an order that is not the right one.
  const wrong = await patchTable(files.up, (s) => s.replace(/<items count="(\d+)"><item x="(\d+)"\/><item x="(\d+)"\/>/, '<items count="$1"><item x="$3"/><item x="$2"/>'));
  assert.equal(independentPivots({ wrong }).some((problem) => /not in the order/.test(problem)), true);
});

test('the refreshed sorted table reads in that order', async () => {
  const words: Datum[][] = [['R', 'V'], ['Kâğıt', 1], ['Kalem', 2], ['Âdem', 4], ['âcil', 5], ['abla', 6]];
  const doc = await load(await pivotBook({ data: words, rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  assert.deepEqual(grid(doc.wb, 1, ref.top + 1, ref.left, ref.bottom - 1, ref.left).map((row) => row[0]), ['abla', 'âcil', 'Âdem', 'Kâğıt', 'Kalem']);
});
