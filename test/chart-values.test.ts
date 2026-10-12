// The numbers a chart is drawn from: the references of its series followed
// to the cells of the workbook as they are now, and the values the file
// cached wherever a reference cannot be followed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { open, type Chart, type DrawingChart, type Workbook } from '../src/index';
import { MAX_CHART_CELLS, chartValues } from '../src/layout/chart';
import { MAX_CHART_POINTS, chartRefAreas } from '../src/model/sheet/chart';
import { History } from '../src/ops/ops';
import { setCellInput } from '../src/ops/commands';
import { setColumnsHidden, setRowsHidden } from '../src/ops/format';
import { insertSheet } from '../src/ops/sheets';
import { catAx, chartFrameOf, chartWorkbook, chartXml, group, numLit, numRef, readChart, ser, sheetData, strLit, strRef, valAx, xyAxes, xySer } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Sayfa1: A1:C5 (a header row and four rows); "Veri 'iki'": numbers in A1:D2. */
const CELLS = sheetData([
  ['Ürün', 'Adet', 'Tutar'],
  ['Çay', 10, 1.5],
  ['Kahve', 20, 2.5],
  ['Süt', 30, 3.5],
  ['Su', 40, 4.5],
]);
const OTHER = sheetData([
  [1, 2, 3, 4],
  [5, 6, 7, 8],
]);

async function book(chart: string, options: { sheet?: string; styles?: string; head?: string } = {}): Promise<{ workbook: Workbook; chart: Chart }> {
  const bytes = await chartWorkbook({
    sheet: options.sheet ?? CELLS,
    sheets: [{ name: "Veri 'iki'", xml: OTHER }],
    drawing: twoCell([4, 0], [10, 14], chartFrameOf('rIdC1')),
    charts: [chart],
    styles: options.styles,
    workbookHead: options.head,
  });
  const workbook = await (await open(bytes)).workbook();
  return { workbook, chart: (workbook.sheets[0].drawings![0].object as DrawingChart).chart };
}

const bar = (series: string, tail?: string) => chartXml({ plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', series) + catAx() + valAx(), tail });
/** The cache differs from the cells on purpose: which of the two is drawn is what the tests are about. */
const usual = ser(0, { tx: strRef('Sayfa1!$B$1', ['eski ad']), cat: strRef('Sayfa1!$A$2:$A$5', ['w', 'x', 'y', 'z']), val: numRef('Sayfa1!$B$2:$B$5', [1, 2, 3, 4]) });

test('references are followed: plain and quoted sheet names, single cells, rows, several areas', async () => {
  const { workbook } = await book(bar(usual));
  const shape = (reference: string) => {
    const areas = chartRefAreas(workbook, reference);
    return areas ? areas.map((area) => `${area.sheet.name}:${area.range.top},${area.range.left},${area.range.bottom},${area.range.right}`) : undefined;
  };
  assert.deepEqual(shape('Sayfa1!$A$2:$A$5'), ['Sayfa1:1,0,4,0']);
  assert.deepEqual(shape('=Sayfa1!A2:A5'), ['Sayfa1:1,0,4,0'], 'a leading "=" and references without "$"');
  assert.deepEqual(shape('sayfa1!$B$1'), ['Sayfa1:0,1,0,1'], 'a single cell; names are compared without case');
  assert.deepEqual(shape('Sayfa1!$C$5:$A$2'), ['Sayfa1:1,0,4,2'], 'corners in any order');
  assert.deepEqual(shape("'Veri ''iki'''!$A$1:$D$1"), ["Veri 'iki':0,0,0,3"], 'a name in quotes, with a quote in it');
  assert.deepEqual(shape("(Sayfa1!$B$2:$B$3,'Veri ''iki'''!$A$2:$B$2,Sayfa1!$B$5)"), ['Sayfa1:1,1,2,1', "Veri 'iki':1,0,1,1", 'Sayfa1:4,1,4,1'], 'areas in the order they are written');
  assert.deepEqual(shape('Sayfa1!$B:$B'), ['Sayfa1:0,1,4,1'], 'a whole column ends where the sheet does');
  assert.deepEqual(shape('Sayfa1!$2:$3'), ['Sayfa1:1,0,2,2'], 'whole rows end where the sheet does');
});

test('references that cannot be followed are not followed', async () => {
  const { workbook } = await book(bar(usual));
  const cases = [
    '',
    'A1:A5',
    '$B$2',
    'Yok!$A$1',
    '[1]Sayfa1!$A$1:$A$5',
    "'[Kitap.xlsx]Sayfa1'!$A$1",
    'Sayfa1:Sayfa2!$A$1',
    'Sayfa1!#REF!',
    '#REF!$A$1',
    '#REF!',
    'Toplam',
    'OFFSET(Sayfa1!$A$1,0,0,5,1)',
    'Sayfa1!$A$1:$XFE$1',
    'Sayfa1!$A$0',
    'Sayfa1!$A$1048577',
    "'Sayfa1!$A$1",
    '(Sayfa1!$A$1,Yok!$A$1)',
    'Sayfa1!$A$1:$B',
    'Sayfa1!' + 'A'.repeat(9000),
  ];
  for (const reference of cases) assert.equal(chartRefAreas(workbook, reference), undefined, JSON.stringify(reference.slice(0, 40)));
  assert.equal(chartRefAreas(workbook, undefined), undefined);
  assert.equal(chartRefAreas(workbook, 42 as never), undefined);
});

test('a chart sheet is no sheet a series can read cells from', async () => {
  const chart = bar(usual);
  const workbook = await (await open(await chartWorkbook({ sheet: CELLS, chartSheet: chart }))).workbook();
  assert.equal(chartRefAreas(workbook, 'Grafik1!$A$1'), undefined);
  assert.ok(chartRefAreas(workbook, 'Sayfa1!$A$1'), 'the worksheet beside it is');
});

test('the cells as they are now are drawn, not what the file cached: values, categories, the name', async () => {
  const { workbook, chart } = await book(bar(usual));
  const values = chartValues(chart, { workbook });
  assert.equal(values.series.length, 1);
  assert.deepEqual(values.series[0].values.toArray(), [10, 20, 30, 40]);
  assert.equal(values.series[0].live, true);
  assert.equal(values.series[0].name, 'Adet');
  assert.equal(values.count, 4);
  assert.deepEqual([0, 1, 2, 3, 4, -1].map((i) => values.label(i)), ['Çay', 'Kahve', 'Süt', 'Su', '', '']);

  // An edit shows at once, and undo brings the old numbers back.
  const history = new History(workbook);
  history.apply(setCellInput(workbook, 0, 2, 1, '99').concat(setCellInput(workbook, 0, 1, 0, 'Ihlamur')), undefined, undefined);
  const edited = chartValues(chart, { workbook });
  assert.deepEqual(edited.series[0].values.toArray(), [10, 99, 30, 40]);
  assert.equal(edited.label(0), 'Ihlamur');
  history.undo();
  assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [10, 20, 30, 40]);
});

test('without a workbook, and for references that cannot be followed, the cache of the file is drawn', async () => {
  const { workbook, chart } = await book(
    bar(
      usual +
        ser(1, { tx: strRef('[1]Dış!$B$1', ['Dış ad']), cat: strRef('Toplam', ['p', 'q']), val: numRef('[1]Dış!$B$2:$B$3', [7, 8]) }) +
        ser(2, { val: numRef('Silinmiş!$B$2:$B$3', [5, null, 6]) }) +
        ser(3, { tx: '<c:v>Elle</c:v>', val: numLit([1, 2]) }) +
        ser(4, { val: numRef('Sayfa1!#REF!', [3]) }),
    ),
  );
  const cached = chartValues(chart);
  assert.deepEqual(cached.series[0].values.toArray(), [1, 2, 3, 4], 'no workbook: the cache');
  assert.equal(cached.series[0].live, false);
  assert.equal(cached.series[0].name, 'eski ad');
  assert.deepEqual([0, 3].map((i) => cached.label(i)), ['w', 'z']);

  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series.map((entry) => entry.live), [true, false, false, false, false]);
  assert.deepEqual(values.series[1].values.toArray(), [7, 8], 'another workbook: the cache');
  assert.equal(values.series[1].name, 'Dış ad');
  assert.deepEqual(values.series[2].values.toArray(), [5, null, 6], 'a sheet that is not there: the cache, with its gap');
  assert.equal(values.series[2].name, 'Series3', 'a series without a name is named by its place');
  assert.equal(values.series[3].name, 'Elle');
  assert.deepEqual(values.series[4].values.toArray(), [3], '#REF!: the cache');
  assert.equal(chartValues(chart, { workbook, seriesName: (n) => `Seri${n}` }).series[2].name, 'Seri3');
});

test('a sheet added later does not get in the way; a deleted sheet is not followed', async () => {
  const { workbook, chart } = await book(bar(ser(0, { val: numRef("'Veri ''iki'''!$A$1:$D$1", [0, 0, 0, 0]) })));
  const history = new History(workbook);
  history.apply(insertSheet(workbook, 0), undefined, undefined);
  assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [1, 2, 3, 4]);
  // Charts keep sheets from being deleted through the operations; the model is changed by hand to stand for a workbook that lost the sheet.
  workbook.sheets = workbook.sheets.filter((sheet) => sheet.name !== "Veri 'iki'");
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].values.toArray(), [0, 0, 0, 0], 'the cache');
  assert.equal(values.series[0].live, false);
});

test('several areas are read one after the other; a block is read row by row', async () => {
  const { workbook, chart } = await book(bar(ser(0, { val: numRef("(Sayfa1!$B$2:$B$3,'Veri ''iki'''!$A$2:$B$2,Sayfa1!$B$5)", []) }) + ser(1, { val: numRef("'Veri ''iki'''!$A$1:$D$2", []) })));
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].values.toArray(), [10, 20, 5, 6, 40]);
  assert.deepEqual(values.series[1].values.toArray(), [1, 2, 3, 4, 5, 6, 7, 8]);
  assert.equal(values.count, 8);
});

test('hidden rows and columns are left out when the chart shows visible cells only, and kept otherwise', async () => {
  const series = ser(0, { cat: strRef('Sayfa1!$A$2:$A$5', []), val: numRef('Sayfa1!$B$2:$B$5', []) }) + ser(1, { val: numRef("'Veri ''iki'''!$A$1:$D$1", []) });
  for (const visibleOnly of [true, false]) {
    const { workbook, chart } = await book(bar(series, `<c:plotVisOnly val="${visibleOnly ? 1 : 0}"/><c:dispBlanksAs val="gap"/>`));
    const history = new History(workbook);
    history.apply(setRowsHidden(0, 2, 3, true).concat(setColumnsHidden(workbook, 1, 1, 1, true)), undefined, undefined);
    const values = chartValues(chart, { workbook });
    if (visibleOnly) {
      assert.deepEqual(values.series[0].values.toArray(), [10, 40], 'rows 3 and 4 are hidden');
      assert.deepEqual([values.label(0), values.label(1)], ['Çay', 'Su'], 'their categories go with them');
      assert.deepEqual(values.series[1].values.toArray(), [1, 3, 4], 'column B is hidden');
    } else {
      assert.deepEqual(values.series[0].values.toArray(), [10, 20, 30, 40]);
      assert.deepEqual(values.series[1].values.toArray(), [1, 2, 3, 4]);
    }
    history.undo();
    assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [10, 20, 30, 40], 'shown again');
  }
});

test('cells without a number are a gap, zero, or spanned, as the chart says', async () => {
  // B3 is a text, B4 is empty, B5 an error.
  const sheet = `<sheetData><row r="2"><c r="B2"><v>1</v></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>yok</t></is></c></row><row r="5"><c r="B5" t="e"><v>#N/A</v></c></row><row r="6"><c r="B6" t="b"><v>1</v></c></row><row r="7"><c r="B7"><v>7</v></c></row></sheetData>`;
  for (const mode of ['gap', 'zero', 'span'] as const) {
    const { workbook, chart } = await book(bar(ser(0, { val: numRef('Sayfa1!$B$2:$B$7', [1, null, null, null, null, 7]) }) + ser(1, { val: numRef('Yok!$B$2:$B$7', [1, null, null, null, null, 7]) }), `<c:plotVisOnly val="1"/><c:dispBlanksAs val="${mode}"/>`), { sheet });
    assert.equal(chart.blanks, mode);
    const values = chartValues(chart, { workbook });
    // UNSURE: Excel draws a text as zero; here text, an error and TRUE are cells without a number like an empty one.
    const expected = mode === 'zero' ? [1, 0, 0, 0, 0, 7] : [1, null, null, null, null, 7];
    assert.deepEqual(values.series[0].values.toArray(), expected, `${mode}: from cells`);
    assert.deepEqual(values.series[1].values.toArray(), expected, `${mode}: from the cache`);
  }
});

test('dates in cells of type "d" are numbers; categories show what their cells show, in the language asked for', async () => {
  const styles =
    `${XML}<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>' +
    '<cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>';
  const sheet =
    '<sheetData><row r="1"><c r="A1" s="1"><v>1234.5</v></c><c r="B1" s="1"><v>1234.5</v></c></row>' +
    '<row r="2"><c r="A2" s="2"><v>45000</v></c><c r="B2" t="d"><v>2024-01-02T00:00:00</v></c></row></sheetData>';
  const { workbook, chart } = await book(bar(ser(0, { cat: numRef('Sayfa1!$A$1:$A$2', [1, 2]), val: numRef('Sayfa1!$B$1:$B$2', [1, 2]) })), { sheet, styles });
  const english = chartValues(chart, { workbook });
  assert.deepEqual(english.series[0].values.toArray(), [1234.5, 45293]);
  assert.deepEqual(english.series[0].format, { code: '#,##0.00', builtinId: undefined }, 'the format of the first cell with a number');
  assert.deepEqual([english.label(0), english.label(1)], ['1,234.50', '3/15/2023']);
  const turkish = chartValues(chart, { workbook, locale: LOCALE_TR });
  assert.deepEqual([turkish.label(0), turkish.label(1)], ['1.234,50', '15.03.2023']);
});

test('categories from the cache: texts as they are, numbers in their format, literals, and 1, 2, 3 for none', () => {
  const read = (cat: string | undefined) => {
    const chart = chartXmlModel(bar(ser(0, { cat, val: numLit([5, 6, 7]) })));
    const values = chartValues(chart, {});
    return [values.label(0), values.label(1), values.label(2)];
  };
  assert.deepEqual(read(strRef('X!$A$1:$A$3', ['a', null, 'c'])), ['a', '', 'c']);
  assert.deepEqual(read(numRef('X!$A$1:$A$3', [0.5, 1, 1.5], '0.0%')), ['50.0%', '100.0%', '150.0%']);
  assert.deepEqual(read(strLit(['bir', 'iki', 'üç'])), ['bir', 'iki', 'üç']);
  assert.deepEqual(read(undefined), ['1', '2', '3']);
});

function chartXmlModel(xml: string): Chart {
  const chart = readChart(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

test('several levels of categories: the level next to the axis is shown, from the cache and from the cells', async () => {
  const levels =
    '<c:multiLvlStrRef><c:f>Sayfa1!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/>' +
    '<c:lvl><c:pt idx="0"><c:v>iç0</c:v></c:pt><c:pt idx="1"><c:v>iç1</c:v></c:pt><c:pt idx="2"><c:v>iç2</c:v></c:pt><c:pt idx="3"><c:v>iç3</c:v></c:pt></c:lvl>' +
    '<c:lvl><c:pt idx="0"><c:v>dış0</c:v></c:pt><c:pt idx="2"><c:v>dış2</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef>';
  const { workbook, chart } = await book(bar(ser(0, { cat: levels, val: numRef('Sayfa1!$C$2:$C$5', [1, 2, 3, 4]) })));
  const categories = chart.groups[0].series[0].categories;
  assert.equal(categories?.multiLevel, true);
  assert.deepEqual(categories?.cache, { index: [0, 1, 2, 3], values: ['iç0', 'iç1', 'iç2', 'iç3'] });
  const cached = chartValues(chart);
  assert.deepEqual([cached.label(0), cached.label(3)], ['iç0', 'iç3']);
  // UNSURE: of the block A2:B5 the last column is taken for the level next to the axis.
  const live = chartValues(chart, { workbook });
  assert.deepEqual([live.label(0), live.label(3)], ['10', '40']);
  assert.deepEqual(categories?.levels, [{ index: [0, 2], values: ['dış0', 'dış2'] }], 'the level beyond it is kept too');
  assert.deepEqual([0, 1, 2, 3].map((i) => cached.outerLabel?.(i)), ['dış0', '', 'dış2', ''], 'and is what a second row of labels shows');
  assert.ok(!chart.notes.some((note) => /levels of categories/.test(note)), 'two levels are both shown: nothing to note');
});

test('a scatter series has a place for every point: its x values, or 1, 2, 3 when they are no numbers', async () => {
  const scatter = (series: string) => chartXml({ plot: group('scatterChart', '<c:scatterStyle val="lineMarker"/>', series) + xyAxes() });
  const { workbook, chart } = await book(
    scatter(
      xySer(0, { cat: numRef('Sayfa1!$B$2:$B$5', []), val: numRef('Sayfa1!$C$2:$C$5', []) }) +
        xySer(1, { cat: strRef('Sayfa1!$A$2:$A$5', ['a', 'b', 'c', 'd']), val: numRef('Sayfa1!$C$2:$C$5', []) }) +
        xySer(2, { val: numLit([7, 8]) }) +
        xySer(3, { cat: numLit([2, 4, 6]), val: numLit([1, null, 3]) }),
    ),
  );
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].xs?.toArray(), [10, 20, 30, 40]);
  assert.deepEqual(values.series[0].values.toArray(), [1.5, 2.5, 3.5, 4.5]);
  assert.deepEqual(values.series[1].xs?.toArray(), [1, 2, 3, 4], 'texts for x');
  assert.deepEqual(values.series[2].xs?.toArray(), [1, 2], 'no x at all');
  assert.deepEqual(values.series[3].xs?.toArray(), [2, 4, 6]);
  assert.deepEqual(values.series[3].values.toArray(), [1, null, 3]);
});

test('a title that is read from a cell follows the cell', async () => {
  const xml = chartXml({ head: '<c:title><c:tx><c:strRef><c:f>Sayfa1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>eski başlık</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>', plot: group('pieChart', '', usual) });
  const { workbook, chart } = await book(xml);
  assert.equal(chartValues(chart, { workbook }).title, 'Ürün');
  assert.equal(chartValues(chart).title, undefined, 'without a workbook the cached text of the model is what there is');
  assert.equal(chart.title?.text, 'eski başlık');
});

test('a reference to more cells than a series may have is read up to the limit', async () => {
  // 120,000 rows with a number in column A.
  let rows = '';
  for (let r = 1; r <= 120000; r++) rows += `<row r="${r}"><c r="A${r}"><v>${r % 7}</v></c></row>`;
  const { workbook, chart } = await book(bar(ser(0, { val: numRef('Sayfa1!$A$1:$A$120000', []) }) + ser(1, { val: numRef('Sayfa1!$A:$A', []) })), { sheet: `<sheetData>${rows}</sheetData>` });
  const started = Date.now();
  const values = chartValues(chart, { workbook });
  assert.equal(values.series[0].values.length, MAX_CHART_POINTS);
  assert.equal(values.series[1].values.length, MAX_CHART_POINTS);
  assert.equal(values.count, MAX_CHART_POINTS);
  assert.equal(values.series[0].values.at(MAX_CHART_POINTS - 1), MAX_CHART_POINTS % 7);
  // A guard against hanging, not a measurement.
  assert.ok(Date.now() - started < 5000, 'two hundred thousand cells are read in well under five seconds');
});

test('the cells read for one chart are limited whatever its series ask for', async () => {
  // 60 series that each ask for a hundred thousand cells: six million.
  let rows = '';
  for (let r = 1; r <= 100000; r++) rows += `<row r="${r}"><c r="A${r}"><v>1</v></c></row>`;
  let series = '';
  for (let i = 0; i < 60; i++) series += ser(i, { val: numRef('Sayfa1!$A$1:$A$100000', [i]) });
  const { workbook, chart } = await book(bar(series), { sheet: `<sheetData>${rows}</sheetData>` });
  const values = chartValues(chart, { workbook });
  let read = 0;
  for (const entry of values.series) read += entry.values.length;
  assert.equal(read, MAX_CHART_CELLS);
  assert.equal(values.series[0].values.length, 100000, 'the first series are read in full');
  assert.equal(values.series[59].values.length, 0, 'those beyond the limit are drawn without values');
});

test('very many hidden columns do not slow a chart down', async () => {
  // 8,000 runs of hidden columns, and a series that reads a row across all of them.
  let cols = '<cols>';
  for (let c = 1; c <= 16000; c += 2) cols += `<col min="${c}" max="${c}" width="9" hidden="1"/>`;
  cols += '</cols>';
  const { workbook, chart } = await book(bar(ser(0, { val: numRef('Sayfa1!$1:$1', []) }) + ser(1, { val: numRef('Sayfa1!$A$1:$XFD$40', []) })), {
    sheet: `${cols}<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="D1"><v>4</v></c></row></sheetData>`,
  });
  const started = Date.now();
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].values.toArray(), [2, 4], 'A and C are hidden: B and D are read');
  assert.equal(values.series[1].values.length, 100000, 'a block of many columns is read up to the limit of a series');
  // A guard against hanging, not a measurement.
  assert.ok(Date.now() - started < 5000, 'in well under five seconds');
});
