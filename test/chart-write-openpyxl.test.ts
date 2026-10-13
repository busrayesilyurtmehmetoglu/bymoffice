// Cross-check of the chart writer against an independent reader: python3 with openpyxl opens
// a workbook the engine added charts to and says what it finds (the type of each chart, its
// series and their references), and a script with zipfile and xml.etree checks that the
// children of every element of each chart part stand in the order the schema states.
// Skipped where python3 or openpyxl is not installed.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { addChart, updateChart, CHART_TYPES, type ChartType } from '../src/ops/charts';
import { FOREIGN, chartsOf, load, packageProblems, salesBook, save } from './chart-edit-fixtures';
import { chartFrameOf } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const PYTHON = String.raw`
import json, sys, zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

C = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'
# The children of the elements the engine writes, in the order of the schema (CT_* of dml-chart.xsd).
ORDER = {
    'chartSpace': ['date1904', 'lang', 'roundedCorners', 'AlternateContent', 'style', 'clrMapOvr', 'pivotSource', 'protection', 'chart', 'spPr', 'txPr', 'externalData', 'printSettings', 'userShapes', 'extLst'],
    'chart': ['title', 'autoTitleDeleted', 'pivotFmts', 'view3D', 'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly', 'dispBlanksAs', 'showDLblsOverMax', 'extLst'],
    # The chart types are a choice that repeats, and so are the axes: each group is one place.
    'plotArea': ['layout', ('areaChart', 'area3DChart', 'lineChart', 'line3DChart', 'stockChart', 'radarChart', 'scatterChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'barChart', 'bar3DChart', 'ofPieChart', 'surfaceChart', 'surface3DChart', 'bubbleChart'), ('valAx', 'catAx', 'dateAx', 'serAx'), 'dTable', 'spPr', 'extLst'],
    'barChart': ['barDir', 'grouping', 'varyColors', 'ser', 'dLbls', 'gapWidth', 'overlap', 'serLines', 'axId', 'extLst'],
    'lineChart': ['grouping', 'varyColors', 'ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars', 'marker', 'smooth', 'axId', 'extLst'],
    'areaChart': ['grouping', 'varyColors', 'ser', 'dLbls', 'dropLines', 'axId', 'extLst'],
    'pieChart': ['varyColors', 'ser', 'dLbls', 'firstSliceAng', 'extLst'],
    'doughnutChart': ['varyColors', 'ser', 'dLbls', 'firstSliceAng', 'holeSize', 'extLst'],
    'scatterChart': ['scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId', 'extLst'],
    'ser': ['idx', 'order', 'tx', 'spPr', 'invertIfNegative', 'pictureOptions', 'marker', 'explosion', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'xVal', 'val', 'yVal', 'smooth', 'shape', 'bubbleSize', 'bubble3D', 'extLst'],
    'catAx': ['axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'minorGridlines', 'title', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos', 'spPr', 'txPr', 'crossAx', 'crosses', 'crossesAt', 'auto', 'lblAlgn', 'lblOffset', 'tickLblSkip', 'tickMarkSkip', 'noMultiLvlLbl', 'extLst'],
    'valAx': ['axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'minorGridlines', 'title', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos', 'spPr', 'txPr', 'crossAx', 'crosses', 'crossesAt', 'crossBetween', 'majorUnit', 'minorUnit', 'dispUnits', 'extLst'],
    'scaling': ['logBase', 'orientation', 'max', 'min', 'extLst'],
    'title': ['tx', 'layout', 'overlay', 'spPr', 'txPr', 'extLst'],
    'legend': ['legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr', 'extLst'],
    'dLbls': ['dLbl', 'delete', 'numFmt', 'spPr', 'txPr', 'dLblPos', 'showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize', 'separator', 'showLeaderLines', 'leaderLines', 'extLst'],
    'numRef': ['f', 'numCache', 'extLst'],
    'strRef': ['f', 'strCache', 'extLst'],
    'numCache': ['formatCode', 'ptCount', 'pt', 'extLst'],
    'strCache': ['ptCount', 'pt', 'extLst'],
    'marker': ['symbol', 'size', 'spPr', 'extLst'],
}

def local(tag):
    return tag.split('}')[-1]

def check(el, problems, where):
    name = local(el.tag)
    kids = [local(k.tag) for k in el]
    if name in ORDER and el.tag.startswith(C):
        order = ORDER[name]
        last = -1
        places = {}
        for i, entry in enumerate(order):
            for item in (entry if isinstance(entry, tuple) else (entry,)):
                places[item] = i
        for kid in kids:
            if kid not in places:
                problems.append('%s: <%s> is not a child of <%s>' % (where, kid, name))
                continue
            at = places[kid]
            if at < last:
                problems.append('%s: <%s> stands too late in <%s> (%s)' % (where, kid, name, ' '.join(kids)))
            last = max(last, at)
    for k in el:
        check(k, problems, where)

problems = []
with zipfile.ZipFile(sys.argv[1]) as z:
    for name in sorted(z.namelist()):
        if name.startswith('xl/charts/chart') and name.endswith('.xml'):
            check(ET.fromstring(z.read(name)), problems, name)

wb = load_workbook(sys.argv[1])
charts = []
for ws in wb.worksheets:
    for chart in ws._charts:
        def ref(source):
            if source is None:
                return None
            for kind in ('numRef', 'strRef'):
                held = getattr(source, kind, None)
                if held is not None:
                    return held.f
            return None
        series = []
        for s in chart.series:
            series.append({
                'name': s.tx.strRef.f if s.tx is not None and s.tx.strRef is not None else (s.tx.v if s.tx is not None else None),
                'cat': ref(getattr(s, 'cat', None)) or ref(getattr(s, 'xVal', None)),
                'val': ref(getattr(s, 'val', None)) or ref(getattr(s, 'yVal', None)),
            })
        charts.append({
            'sheet': ws.title,
            'tag': chart.tagname,
            'dir': getattr(chart, 'barDir', None),
            'grouping': getattr(chart, 'grouping', None),
            'title': None if chart.title is None or chart.title.tx is None or chart.title.tx.rich is None else ''.join(r.t for p in chart.title.tx.rich.p for r in (p.r or [])),
            'legend': None if chart.legend is None else chart.legend.position,
            'series': series,
        })
print(json.dumps({'problems': problems, 'charts': charts}))
`;

const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const TAGS: Record<ChartType, string> = { column: 'barChart', bar: 'barChart', line: 'lineChart', area: 'areaChart', pie: 'pieChart', doughnut: 'doughnutChart', scatter: 'scatterChart' };

test('chart write: openpyxl reads the charts the engine wrote, and their parts are in the order of the schema', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl; print(openpyxl.__version__)'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  // One chart of every type on the sheet, some changed after they were added, beside a chart of another application that was changed in place.
  const doc = await load(await salesBook({ drawing: twoCell([4, 1], [9, 12], chartFrameOf('rIdC1', 'Grafik 1')), charts: [FOREIGN] }));
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'Changed', legend: 'b', axisTitles: { category: 'Ay', value: 'Adet' }, axisBounds: { min: 0, max: 50 }, dataLabels: true, grouping: 'stacked', series: [{ from: 1, color: '#AA0000' }, { from: 0 }, { values: 'Sayfa1!$B$2:$B$4', name: 'Third' }] }), undefined, undefined);
  CHART_TYPES.forEach((type, i) => {
    doc.history.apply(addChart(doc.wb, 0, { type, range: ALL, title: `Chart of ${type}`, grouping: i % 2 ? 'stacked' : 'standard', legend: i % 2 ? 't' : 'r' }), undefined, undefined);
    if (type !== 'pie' && type !== 'doughnut') doc.history.apply(updateChart(doc.wb, 0, i + 1, { axisTitles: { category: 'X', value: 'Y' }, axisBounds: { max: 100 }, dataLabels: i % 2 === 0, series: [{ from: 1, color: '#00AA00' }, { from: 0 }] }), undefined, undefined);
  });
  const saved = await save(doc);
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-chart-write-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  let out: { problems: string[]; charts: { sheet: string; tag: string; dir: string | null; grouping: string | null; title: string | null; legend: string | null; series: { name: string | null; cat: string | null; val: string | null }[] }[] };
  try {
    const file = join(folder, 'charts.xlsx');
    writeFileSync(file, saved);
    const run = spawnSync('python3', ['-c', PYTHON, file], { encoding: 'utf8' });
    assert.equal(run.status, 0, `openpyxl opens the file: ${run.stderr.slice(-2000)}`);
    out = JSON.parse(run.stdout);
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
  assert.deepEqual(out.problems, [], 'every element of the chart parts has its children in the order of the schema');
  assert.equal(out.charts.length, CHART_TYPES.length + 1);
  const changed = out.charts[0];
  assert.deepEqual(
    [changed.tag, changed.grouping, changed.title, changed.legend, changed.series],
    [
      'barChart',
      'stacked',
      'Changed',
      'b',
      [
        { name: 'Sayfa1!$C$1', cat: 'Sayfa1!$A$2:$A$4', val: 'Sayfa1!$C$2:$C$4' },
        { name: 'Sayfa1!$B$1', cat: 'Sayfa1!$A$2:$A$4', val: 'Sayfa1!$B$2:$B$4' },
        { name: 'Third', cat: 'Sayfa1!$A$2:$A$4', val: 'Sayfa1!$B$2:$B$4' },
      ],
    ],
    'the chart of the file as it was changed: the series in their new order',
  );
  CHART_TYPES.forEach((type, i) => {
    const chart = out.charts[i + 1];
    assert.equal(chart.tag, TAGS[type], type);
    assert.equal(chart.title, `Chart of ${type}`, type);
    assert.equal(chart.legend, i % 2 ? 't' : 'r', type);
    if (type === 'column' || type === 'bar') assert.equal(chart.dir, type === 'bar' ? 'bar' : 'col', type);
    if (type === 'column' || type === 'bar' || type === 'line' || type === 'area') assert.equal(chart.grouping, i % 2 ? 'stacked' : type === 'line' || type === 'area' ? 'standard' : 'clustered', type);
    const round = type === 'pie' || type === 'doughnut';
    const north = { name: 'Sayfa1!$B$1', cat: 'Sayfa1!$A$2:$A$4', val: 'Sayfa1!$B$2:$B$4' };
    const south = { name: 'Sayfa1!$C$1', cat: 'Sayfa1!$A$2:$A$4', val: 'Sayfa1!$C$2:$C$4' };
    assert.deepEqual(chart.series, round ? [north, south] : [south, north], `${type}: the series and their references`);
  });
});

// Writes a workbook with four charts the way openpyxl does: the namespace of charts is the default one of each part.
const MAKE = String.raw`
import openpyxl, sys
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference, Series
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sayfa1'
for r in [['Month','North','South'],['Jan',10,30],['Feb',20,20],['Mar',30,10]]: ws.append(r)
ws2 = wb.create_sheet('Other')
for r in [['M','N','X'],['a',1,7],['b',2,8],['c',3,9],['d',4,10]]: ws2.append(r)
b = BarChart(); b.title = 'Bar'; b.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=4), titles_from_data=True); b.set_categories(Reference(ws, min_col=1, min_row=2, max_row=4)); ws.add_chart(b, 'E2')
l = LineChart(); l.add_data(Reference(ws2, min_col=2, min_row=1, max_col=3, max_row=5), titles_from_data=True); ws.add_chart(l, 'E20')
p = PieChart(); p.add_data(Reference(ws, min_col=2, min_row=1, max_row=4), titles_from_data=True); p.set_categories(Reference(ws, min_col=1, min_row=2, max_row=4)); ws2.add_chart(p, 'F2')
s = ScatterChart(); s.series.append(Series(Reference(ws2, min_col=3, min_row=2, max_row=5), Reference(ws2, min_col=2, min_row=2, max_row=5), title='xy')); ws2.add_chart(s, 'F20')
wb.save(sys.argv[1])
`;

// Opens a workbook with openpyxl and parses every part with a parser that knows namespaces (an unbound prefix
// is an error there); says how many charts each sheet has, what the series of each read, and the namespaces of
// charts each chart part uses.
const READ = String.raw`
import json, sys, re, zipfile, warnings
import xml.etree.ElementTree as ET
import openpyxl
warnings.simplefilter('ignore')
problems = []; spaces = {}
z = zipfile.ZipFile(sys.argv[1])
for name in z.namelist():
    if not re.search(r'\.(xml|rels|vml)$', name): continue
    try:
        root = ET.fromstring(z.read(name))
    except ET.ParseError as e:
        problems.append('%s: %s' % (name, e)); continue
    if '/charts/chart' in name:
        spaces[name] = sorted(set(el.tag[1:].split('}')[0] for el in root.iter() if el.tag[0] == '{' and re.search(r'drawingml/(2006/)?chart$', el.tag[1:].split('}')[0])))
wb = openpyxl.load_workbook(sys.argv[1])
def ref(source):
    for kind in ('numRef', 'strRef'):
        held = getattr(source, kind, None) if source is not None else None
        if held is not None: return held.f
    return None
charts = [[{'tag': c.tagname, 'series': [[s.tx.strRef.f if s.tx is not None and s.tx.strRef is not None else (s.tx.v if s.tx is not None else None), ref(getattr(s, 'cat', None)) or ref(getattr(s, 'xVal', None)), ref(getattr(s, 'val', None)) or ref(getattr(s, 'yVal', None))] for s in c.series]} for c in ws._charts] for ws in wb.worksheets]
print(json.dumps({'problems': problems, 'spaces': spaces, 'charts': charts}))
`;

test('chart write: charts openpyxl wrote are changed in their parts and by rows that are inserted, and openpyxl opens the result', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl; print(openpyxl.__version__)'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-chart-edit-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  try {
    const start = join(folder, 'start.xlsx');
    const made = spawnSync('python3', ['-c', MAKE, start], { encoding: 'utf8' });
    assert.equal(made.status, 0, `openpyxl writes the file: ${made.stderr.slice(-1000)}`);
    const bytes = new Uint8Array(readFileSync(start));
    const read = (file: string) => {
      const run = spawnSync('python3', ['-c', READ, file], { encoding: 'utf8' });
      assert.equal(run.status, 0, `openpyxl opens ${file}: ${run.stderr.slice(-1500)}`);
      return JSON.parse(run.stdout) as { problems: string[]; spaces: Record<string, string[]>; charts: { tag: string; series: (string | null)[][] }[][] };
    };
    const doc = await load(bytes);
    assert.deepEqual(chartsOf(doc.wb.sheets[0]).map((chart) => chart.known), ['refs', 'refs'], 'the charts of openpyxl are changed in their parts');
    assert.equal(doc.wb.chartsFollow, true);
    const steps: [string, () => void][] = [
      ['a row above the data', () => doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined)],
      ['the legend at the bottom', () => doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b' }), undefined, undefined)],
      ['a title, axis titles, labels, a colour and a series more', () => doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'Changed', axisTitles: { category: 'Ay', value: 'Adet' }, dataLabels: true, series: [{ from: 1, color: '#AA0000' }, { from: 0, name: 'N' }, { values: 'Other!$B$2:$B$5' }] }), undefined, undefined)],
      ['the pie of the other sheet', () => doc.history.apply(updateChart(doc.wb, 1, 0, { title: 'Pie', legend: 'none' }), undefined, undefined)],
      ['a chart added', () => doc.history.apply(addChart(doc.wb, 1, { type: 'area', range: { top: 0, left: 0, bottom: 4, right: 2 } }), undefined, undefined)],
      ['the sheet renamed', () => doc.history.apply([{ type: 'renameSheet', sheet: 0, name: 'Satış 1' }], undefined, undefined)],
    ];
    for (const [what, act] of steps) {
      act();
      const out = await save(doc);
      assert.deepEqual(await packageProblems(out), [], `${what}: the package is whole`);
      const file = join(folder, 'out.xlsx');
      writeFileSync(file, out);
      const found = read(file);
      assert.deepEqual(found.problems, [], `${what}: every part is read by a parser that knows namespaces`);
      for (const name of Object.keys(found.spaces)) assert.equal(found.spaces[name].length, 1, `${what}: ${name} has one namespace of charts: ${found.spaces[name].join(', ')}`);
      assert.deepEqual(found.charts.map((list) => list.length), [2, what === 'a chart added' || what === 'the sheet renamed' ? 3 : 2], `${what}: openpyxl finds every chart`);
    }
    const last = read(join(folder, 'out.xlsx'));
    assert.deepEqual(last.charts[0][0], {
      tag: 'barChart',
      series: [
        ["'Satış 1'!C2", "'Satış 1'!$A$3:$A$5", "'Satış 1'!$C$3:$C$5"],
        ['N', "'Satış 1'!$A$3:$A$5", "'Satış 1'!$B$3:$B$5"],
        [null, "'Satış 1'!$A$3:$A$5", 'Other!$B$2:$B$5'],
      ],
    });
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
});
