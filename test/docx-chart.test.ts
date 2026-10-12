// Charts of a Word document: the reader finds the chart part of a frame and
// reads it with the reader of the workbook charts (from the cached values; the
// embedded workbook is never opened), the layout gives the frame the chart,
// the painter draws it inside the frame, and nothing of it is written anew.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_DOC_CHARTS, MAX_DOC_CHART_POINTS } from '../src/formats/docx/charts';
import { readDocx, type DocxReadOptions } from '../src/formats/docx/document';
import { emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { open as openFile } from '../src/index';
import type { ImageItem, LayoutFont, LayoutPage, PageItem } from '../src/layout/doc/types';
import { chartCached, MAX_WORKBOOK_CHARTS, MAX_WORKBOOK_CHART_POINTS, type Chart } from '../src/model/sheet/chart';
import type { Block, DocModel, DrawingItem, Paragraph, Table } from '../src/model/doc/types';
import { DocHistory, insertText } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { renderPage } from '../src/render/doc';
import { DocChartPainter } from '../src/render/doc-chart';
import { cutLoses } from '../src/ui/doc-clipboard';
import { documentText } from '../src/ui/doc-view';
import { ZipArchive } from '../src/zip';
import { title } from './chart-fixtures';
import { lay, p, r, tbl, tc, texts, tr, words } from './doc-layout-fixtures';
import { chartParts, columnsChart, floatingChart, inlineChart, inlineChartDrawing, partRels } from './docx-chart-fixtures';
import { buildDocx as buildSpec, THEME, type DocxSpec } from './docx-fixtures';

const CHART1 = '/word/charts/chart1.xml';

async function read(spec: DocxSpec, options?: DocxReadOptions): Promise<DocModel> {
  return readDocx(await OpcPackage.open(await buildSpec(spec)), options);
}

/** Every drawing of the blocks, in document order: those in tables, text boxes and alternate content too. */
function drawings(blocks: Block[], out: DrawingItem[] = []): DrawingItem[] {
  for (const block of blocks) {
    if (block.type === 'table') {
      for (const row of (block as Table).rows) for (const cell of row.cells) drawings(cell.blocks, out);
      continue;
    }
    if (block.type !== 'paragraph') continue;
    for (const inline of (block as Paragraph).children) {
      if (inline.type !== 'run') continue;
      const visit = (items: typeof inline.content) => {
        for (const item of items) {
          if (item.type === 'drawing') {
            out.push(item);
            for (const box of item.textBoxes || []) drawings(box, out);
          } else if (item.type === 'alternate') visit(item.items);
        }
      };
      visit(inline.content);
    }
  }
  return out;
}

const images = (pages: LayoutPage[]): ImageItem[] => pages.reduce<PageItem[]>((out, page) => out.concat(page.items), []).filter((item): item is ImageItem => item.type === 'image');

const firstValues = (chart: Chart | undefined): (number | string | null)[] => {
  assert.ok(chart, 'the chart was read');
  const data = chart.groups[0].series[0].values;
  return [0, 1, 2].map((i) => chartCached(data, i));
};

// ------------------------------------------------------------------- reader

test('reader: a chart in its line is read from the cache of its part', async () => {
  const doc = await read({ body: p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const [drawing] = drawings(doc.body.blocks);
  assert.equal(drawing.kind, 'chart');
  assert.equal(drawing.placement, 'inline');
  assert.equal(drawing.chartPart, CHART1);
  assert.ok(doc.charts, 'the document has charts');
  assert.deepEqual(Array.from(doc.charts.keys()), [CHART1]);
  const chart = doc.charts.get(CHART1);
  assert.equal(chart && chart.groups[0].kind, 'bar');
  assert.deepEqual(firstValues(chart), [50, 100, 25]);
  assert.deepEqual(doc.warnings, []);
});

test('reader: a chart text flows around is read as well', async () => {
  const doc = await read({ body: p(floatingChart('rIdC1', 10, 0, 200, 150) + r('text')), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const [drawing] = drawings(doc.body.blocks);
  assert.equal(drawing.placement, 'anchor');
  assert.equal(drawing.chartPart, CHART1);
  assert.deepEqual(firstValues(doc.charts && doc.charts.get(CHART1)), [50, 100, 25]);
});

test('reader: a chart inside alternate content is found in the branch that is read', async () => {
  const body = p(
    `<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps">${inlineChartDrawing('rIdC1', 300, 200)}</mc:Choice>` +
      `<mc:Fallback>${inlineChartDrawing('rIdC2', 300, 200)}</mc:Fallback></mc:AlternateContent></w:r>`,
  );
  const doc = await read({ body, ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }, { rel: 'rIdC2', xml: columnsChart({ values: [1, 2, 3] }) }])) });
  const found = drawings(doc.body.blocks);
  assert.equal(found.length, 1);
  assert.equal(found[0].chartPart, CHART1);
  // The branch that is not read costs nothing: its chart part is not parsed.
  assert.deepEqual(Array.from((doc.charts as Map<string, Chart>).keys()), [CHART1]);
});

test('reader: the chart of a header is found through the relationships of the header', async () => {
  const charts = await chartParts([{ rel: 'rIdC1', xml: columnsChart() }, { rel: 'unused', xml: columnsChart({ values: [7, 8, 9] }) }]);
  // The same id means another part in the header than in the document.
  const doc = await read({
    body: p(inlineChart('rIdC1', 300, 200)) + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
    headers: { rIdH: p(inlineChart('rIdC1', 100, 60)) },
    rels: charts.rels,
    parts: charts.parts.concat([partRels('header1.xml', [['rIdC1', 'chart', 'charts/chart2.xml']])]),
  });
  const header = doc.headers.get('rIdH');
  assert.ok(header, 'the header was read');
  assert.equal(drawings(header.blocks)[0].chartPart, '/word/charts/chart2.xml');
  assert.equal(drawings(doc.body.blocks)[0].chartPart, CHART1);
  assert.deepEqual(firstValues((doc.charts as Map<string, Chart>).get('/word/charts/chart2.xml')), [7, 8, 9]);
});

test('reader: the chart of a footnote is found through the relationships of the notes part', async () => {
  const charts = await chartParts([{ rel: 'unused', xml: columnsChart() }]);
  const doc = await read({
    body: p('<w:r><w:footnoteReference w:id="2"/></w:r>'),
    footnotes: `<w:footnote w:id="2">${p(inlineChart('rIdN', 100, 60))}</w:footnote>`,
    parts: charts.parts.concat([partRels('footnotes.xml', [['rIdN', 'chart', 'charts/chart1.xml']])]),
  });
  const note = doc.footnotes && doc.footnotes.items.get(2);
  assert.ok(note, 'the footnote was read');
  assert.equal(drawings(note.blocks)[0].chartPart, CHART1);
  assert.ok(doc.charts && doc.charts.has(CHART1), 'the chart of the footnote is in the model');
});

test('reader: charts in a table cell and in a text box are read', async () => {
  const box =
    '<w:r><w:drawing><wp:inline><wp:extent cx="3810000" cy="2857500"/><wp:docPr id="3" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    `<wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>${p(inlineChart('rIdC2', 200, 100))}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const body = tbl(tr(tc(p(inlineChart('rIdC1', 100, 80)))), [3000]) + p(box);
  const doc = await read({ body, ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }, { rel: 'rIdC2', xml: columnsChart() }])) });
  const parts = drawings(doc.body.blocks)
    .filter((drawing) => drawing.kind === 'chart')
    .map((drawing) => drawing.chartPart);
  assert.deepEqual(parts, [CHART1, '/word/charts/chart2.xml']);
  assert.equal((doc.charts as Map<string, Chart>).size, 2);
});

test('reader: two frames of one chart part read it once', async () => {
  const doc = await read({ body: p(inlineChart('rIdC1', 300, 200)) + p(inlineChart('rIdC1', 150, 100)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) }, { maxCharts: 1 });
  assert.deepEqual(drawings(doc.body.blocks).map((drawing) => drawing.chartPart), [CHART1, CHART1]);
  assert.equal((doc.charts as Map<string, Chart>).size, 1);
  assert.deepEqual(doc.warnings, []);
});

test('reader: a series without cached values is read, drawn empty and noted; the embedded workbook is not opened', async () => {
  const doc = await read({ body: p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart({ uncached: true }) }])) });
  const chart = (doc.charts as Map<string, Chart>).get(CHART1);
  assert.deepEqual(firstValues(chart), [null, null, null]);
  assert.ok(
    (chart as Chart).notes.some((note) => /cached/.test(note) && /embedded workbook/.test(note)),
    `a note says that the series has no cached values: ${JSON.stringify((chart as Chart).notes)}`,
  );
});

const broken: [name: string, xml: string | Uint8Array, warning: RegExp][] = [
  ['not XML', '<c:chartSpace><c:chart>', /chart1\.xml.*could not be read/],
  ['a DTD', '<?xml version="1.0"?><!DOCTYPE c [<!ENTITY a "aaaa">]>' + columnsChart().replace(/^<\?xml[^>]*\?>/, ''), /chart1\.xml.*could not be read/],
  ['not text', new Uint8Array([0xff, 0xfe, 0x00, 0x01, 0x02, 0x80]), /chart1\.xml.*could not be read/],
  ['a type that is not drawn', columnsChart({ type: 'surfaceChart' }), /chart1\.xml.*nothing that is drawn/],
  ['another root', '<?xml version="1.0"?><root/>', /chart1\.xml.*nothing that is drawn/],
];
for (const [name, xml, warning] of broken) {
  test(`reader: a chart part that is ${name} leaves a placeholder and a warning; the document opens`, async () => {
    const doc = await read({ body: p('before') + p(inlineChart('rIdC1', 300, 200)) + p('after'), ...(await chartParts([{ rel: 'rIdC1', xml }])) });
    assert.equal(documentText(doc), 'before\n\nafter');
    const [drawing] = drawings(doc.body.blocks);
    assert.equal(drawing.kind, 'chart');
    assert.ok(!doc.charts || !doc.charts.has(CHART1), 'no chart is in the model');
    assert.ok(doc.warnings.some((line) => warning.test(line)), `a warning says so: ${JSON.stringify(doc.warnings)}`);
    const laid = await lay({ body: p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml }])) });
    const [image] = images(laid.pages);
    assert.equal(image.placeholder, 'chart');
    assert.equal(image.chart, undefined);
  });
}

test('reader: a chart whose relationship leads out of the package, to no part or nowhere stays a placeholder', async () => {
  const body = p(inlineChart('rIdOut', 300, 200)) + p(inlineChart('rIdGone', 300, 200)) + p(inlineChart('rIdNone', 300, 200)) + p(inlineChart('rIdImage', 300, 200));
  const doc = await read({
    body,
    rels: [
      ['rIdOut', 'chart', 'https://example.invalid/chart1.xml', true],
      ['rIdGone', 'chart', 'charts/chart9.xml'],
      // Not a chart by its type: a frame may not make the reader parse any part of the package as a chart.
      ['rIdImage', 'image', 'charts/chart1.xml'],
    ],
    parts: [['word/charts/chart1.xml', columnsChart()]],
  });
  const found = drawings(doc.body.blocks);
  assert.deepEqual(found.map((drawing) => drawing.kind), ['chart', 'chart', 'chart', 'chart']);
  assert.deepEqual(found.map((drawing) => drawing.chartPart), [undefined, undefined, undefined, undefined]);
  assert.equal(doc.charts, undefined);
  assert.ok(doc.warnings.some((line) => /external/.test(line)), `the external target is reported: ${JSON.stringify(doc.warnings)}`);
  assert.ok(doc.warnings.some((line) => /chart9\.xml.*missing/.test(line)), `the missing part is reported: ${JSON.stringify(doc.warnings)}`);
});

test('reader: the budgets of a document are those of a workbook', () => {
  assert.equal(MAX_DOC_CHARTS, MAX_WORKBOOK_CHARTS);
  assert.equal(MAX_DOC_CHART_POINTS, MAX_WORKBOOK_CHART_POINTS);
});

test('reader: charts beyond the limit of the document stay placeholders, with one warning', async () => {
  const charts = await chartParts([1, 2, 3, 4].map((n) => ({ rel: `rIdC${n}`, xml: columnsChart() })));
  const doc = await read({ body: [1, 2, 3, 4].map((n) => p(inlineChart(`rIdC${n}`, 100, 80))).join(''), ...charts }, { maxCharts: 2 });
  assert.deepEqual(Array.from((doc.charts as Map<string, Chart>).keys()), [CHART1, '/word/charts/chart2.xml']);
  const lines = doc.warnings.filter((line) => /limit/.test(line));
  assert.equal(lines.length, 1, JSON.stringify(doc.warnings));
  assert.match(lines[0], /2\b.*: 2$/);
});

test('reader: a chart whose cached values do not fit what is left for the document is not shown', async () => {
  const charts = await chartParts([
    { rel: 'rIdC1', xml: columnsChart() },
    { rel: 'rIdC2', xml: columnsChart() },
  ]);
  // Each chart holds six cached values (three names, three numbers).
  const doc = await read({ body: p(inlineChart('rIdC1', 100, 80)) + p(inlineChart('rIdC2', 100, 80)), ...charts }, { maxChartPoints: 8 });
  assert.deepEqual(Array.from((doc.charts as Map<string, Chart>).keys()), [CHART1]);
  assert.ok(doc.warnings.some((line) => /chart2\.xml.*too many cached values/.test(line)), JSON.stringify(doc.warnings));
});

test('reader: a chart part larger than the limit is not parsed', async () => {
  const doc = await read({ body: p(inlineChart('rIdC1', 100, 80)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) }, { maxChartPartBytes: 100 });
  assert.equal(doc.charts, undefined);
  assert.ok(doc.warnings.some((line) => /chart1\.xml.*too large/.test(line)), JSON.stringify(doc.warnings));
});

test('reader: series without a colour get the accent colours of the theme of the document', async () => {
  const plain = columnsChart().replace(/<c:spPr>.*?<\/c:spPr>/, '');
  const doc = await read({ body: p(inlineChart('rIdC1', 300, 200)), theme: THEME.replace(/(<a:accent1><a:srgbClr val=")[0-9A-Fa-f]{6}/, '$1AB12CD'), ...(await chartParts([{ rel: 'rIdC1', xml: plain }])) });
  const chart = (doc.charts as Map<string, Chart>).get(CHART1) as Chart;
  assert.equal(chart.palette[0], '#AB12CD');
});

// ----------------------------------------------------------------- the text

test('text: the title of a chart is no text of the document; the description of the frame is the alternative text', async () => {
  const spec = { body: p('one') + p(inlineChart('rIdC1', 300, 200, 'id="7" name="Chart 1" descr="Sales by month"')) + p('two'), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart({ head: title('Hidden title') }) }])) };
  const laid = await lay(spec);
  assert.equal(documentText(laid.doc), 'one\n\ntwo');
  const shown = laid.pages.reduce<string[]>((out, page) => out.concat(texts(page).map((item) => item.text)), []);
  assert.ok(!shown.some((text) => /Hidden/.test(text)), `no text item of the page holds the title: ${JSON.stringify(shown)}`);
  const [image] = images(laid.pages);
  assert.equal(image.alt, 'Sales by month');
  assert.equal(drawings(laid.doc.body.blocks)[0].description, 'Sales by month');
});

test('editing: a selection that holds a chart still cannot be cut', async () => {
  const doc = await read({ body: p('one') + p(inlineChart('rIdC1', 300, 200)) + p('two'), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  assert.equal(cutLoses(doc, { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [2], offset: 3 } }, undefined), true);
});

// ------------------------------------------------------------------- layout

test('layout: the frame has the size the file states and holds the chart of the model', async () => {
  const laid = await lay({ body: p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const [image] = images(laid.pages);
  assert.equal(image.width, 300);
  assert.equal(image.height, 200);
  assert.equal(image.chart, (laid.doc.charts as Map<string, Chart>).get(CHART1));
  // Whoever paints without charts still knows what stands there.
  assert.equal(image.placeholder, 'chart');
  assert.equal(image.part, undefined);
});

test('layout: text flows around a chart as it flowed around its placeholder', async () => {
  const body = p(floatingChart('rIdC1', 20, 0, 200, 150) + r(words(120))) + tbl(tr(tc(p(inlineChart('rIdC1', 100, 80) + r('cell')))), [3000]) + p(words(30));
  const place = (item: PageItem) => (item.type === 'text' ? `${item.text}@${item.x},${item.y}` : item.type === 'image' ? `image@${item.x},${item.y},${item.width},${item.height}` : item.type);
  const drawn = await lay({ body, ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const held = await lay({ body, ...(await chartParts([{ rel: 'rIdC1', xml: '<broken' }])) });
  assert.equal(images(drawn.pages).filter((image) => image.chart).length, 2);
  assert.equal(images(held.pages).filter((image) => image.chart).length, 0);
  assert.deepEqual(drawn.pages.map((page) => page.items.map(place)), held.pages.map((page) => page.items.map(place)));
  const [floating] = images(drawn.pages);
  assert.equal(floating.width, 200);
  assert.equal(floating.height, 150);
  // No text of the page lies in the box of the chart: lines beside it begin after it or end before it, the others stand below.
  const all = texts(drawn.pages[0]);
  assert.ok(all.length > 0, 'the page has text');
  for (const item of all) {
    if (!(item.y > floating.y && item.y < floating.y + floating.height)) continue;
    assert.ok(item.x >= floating.x + floating.width || item.x + item.width <= floating.x + 0.5, `a line runs into the chart: ${item.text.slice(0, 20)}@${item.x}`);
  }
});

test('layout: the chart of a header is in the items of every page', async () => {
  const charts = await chartParts([{ rel: 'rIdC9', xml: columnsChart() }]);
  const laid = await lay({
    body: p('page one') + p('<w:r><w:br w:type="page"/></w:r>') + p('page two') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="2880" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
    headers: { rIdH: p(inlineChart('rIdC1', 120, 60)) },
    parts: charts.parts.concat([partRels('header1.xml', [['rIdC1', 'chart', 'charts/chart1.xml']])]),
  });
  assert.equal(laid.pages.length, 2);
  for (const page of laid.pages) assert.equal(images([page]).filter((image) => image.chart).length, 1);
});

// ----------------------------------------------------------------- painting

interface Recorded {
  ctx: CanvasRenderingContext2D;
  rects: { x: number; y: number; width: number; height: number; fill: string }[];
  texts: { text: string; x: number; y: number; font: string }[];
  /** The bounding boxes of the paths that were made clips, canvas pixels. */
  clips: { x: number; y: number; width: number; height: number }[];
  balance(): number;
  failFrom: number;
}

/** A context that records what is drawn in canvas pixels, under the translations and scalings in effect. Each character is half the font size wide. */
function recorder(): Recorded {
  interface State {
    sx: number;
    sy: number;
    dx: number;
    dy: number;
    fillStyle: string;
    font: string;
  }
  let state: State = { sx: 1, sy: 1, dx: 0, dy: 0, fillStyle: '#000000', font: '10px sans-serif' };
  const stack: State[] = [];
  let path: [number, number][] = [];
  const at = (x: number, y: number): [number, number] => [state.sx * x + state.dx, state.sy * y + state.dy];
  const out: Recorded = { ctx: undefined as unknown as CanvasRenderingContext2D, rects: [], texts: [], clips: [], balance: () => stack.length, failFrom: Infinity };
  const point = (x: number, y: number) => void path.push(at(x, y));
  const ctx: { [name: string]: unknown } = {
    strokeStyle: '#000000',
    lineWidth: 1,
    lineJoin: 'miter',
    lineCap: 'butt',
    globalAlpha: 1,
    textAlign: 'left',
    textBaseline: 'alphabetic',
    save: () => void stack.push({ ...state }),
    restore: () => {
      const top = stack.pop();
      assert.ok(top, 'restore() without save()');
      state = top;
    },
    translate: (x: number, y: number) => {
      state.dx += state.sx * x;
      state.dy += state.sy * y;
    },
    scale: (x: number, y: number) => {
      state.sx *= x;
      state.sy *= y;
    },
    rotate: () => undefined,
    setTransform: () => undefined,
    setLineDash: () => undefined,
    beginPath: () => void (path = []),
    closePath: () => undefined,
    moveTo: point,
    lineTo: point,
    quadraticCurveTo: (_cx: number, _cy: number, x: number, y: number) => point(x, y),
    bezierCurveTo: (_a: number, _b: number, _c: number, _d: number, x: number, y: number) => point(x, y),
    arc: (x: number, y: number) => point(x, y),
    rect: (x: number, y: number, width: number, height: number) => {
      point(x, y);
      point(x + width, y + height);
    },
    clip: () => {
      const xs = path.map((entry) => entry[0]);
      const ys = path.map((entry) => entry[1]);
      const x = Math.min(...xs);
      const y = Math.min(...ys);
      out.clips.push({ x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y });
    },
    fill: () => undefined,
    stroke: () => undefined,
    fillRect: (x: number, y: number, width: number, height: number) => {
      if (out.rects.length >= out.failFrom) throw new Error('the canvas failed');
      const [x1, y1] = at(x, y);
      out.rects.push({ x: x1, y: y1, width: width * state.sx, height: height * state.sy, fill: state.fillStyle });
    },
    fillText: (text: string, x: number, y: number) => {
      const [tx, ty] = at(x, y);
      out.texts.push({ text, x: tx, y: ty, font: state.font });
    },
    measureText: (text: string) => {
      const match = /(\d+(?:\.\d+)?)px/.exec(state.font);
      return { width: text.length * (match ? Number(match[1]) : 10) * 0.5 };
    },
    drawImage: () => undefined,
  };
  Object.defineProperty(ctx, 'fillStyle', { get: () => state.fillStyle, set: (value: string) => void (state.fillStyle = value) });
  Object.defineProperty(ctx, 'font', { get: () => state.font, set: (value: string) => void (state.font = value) });
  out.ctx = ctx as unknown as CanvasRenderingContext2D;
  return out;
}

const cssFont = (font: LayoutFont) => `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}px "${font.family}", doc-chain`;
const GREEN = '#00B050';
const PLACEHOLDER = '#f1f3f4';

async function chartPage(scale = 1, head?: string): Promise<{ page: LayoutPage; image: ImageItem; painter: DocChartPainter; recorded: Recorded }> {
  const laid = await lay({ body: p('before') + p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart({ head }) }])) });
  const painter = new DocChartPainter({ cssFont });
  const recorded = recorder();
  renderPage(recorded.ctx, laid.pages[0], { scale, charts: painter, background: null });
  return { page: laid.pages[0], image: images(laid.pages)[0], painter, recorded };
}

test('painting: the columns of the series stand inside the frame, as high as their values, and the chart is clipped to the frame', async () => {
  const { image, recorded } = await chartPage();
  const bars = recorded.rects.filter((rect) => rect.fill === GREEN);
  assert.equal(bars.length, 3);
  for (const bar of bars) {
    assert.ok(bar.x >= image.x && bar.x + bar.width <= image.x + image.width, `a column is inside the frame horizontally: ${JSON.stringify(bar)}`);
    assert.ok(bar.y >= image.y && bar.y + bar.height <= image.y + image.height, `a column is inside the frame vertically: ${JSON.stringify(bar)}`);
  }
  // 50, 100, 25 on an axis from 0 to 100.
  assert.ok(Math.abs(bars[0].height / bars[1].height - 0.5) < 0.02, `the first column is half the second: ${bars[0].height} / ${bars[1].height}`);
  assert.ok(Math.abs(bars[2].height / bars[1].height - 0.25) < 0.02, `the third column is a quarter of the second: ${bars[2].height} / ${bars[1].height}`);
  assert.ok(
    recorded.clips.some((clip) => Math.abs(clip.x - image.x) < 0.01 && Math.abs(clip.y - image.y) < 0.01 && Math.abs(clip.width - image.width) < 0.01 && Math.abs(clip.height - image.height) < 0.01),
    `the frame is the clip: ${JSON.stringify(recorded.clips)}`,
  );
  assert.ok(!recorded.rects.some((rect) => rect.fill === PLACEHOLDER), 'no placeholder is drawn');
  assert.equal(recorded.balance(), 0);
});

test('painting: a larger scale draws the same chart larger, from the same layout', async () => {
  const one = await chartPage(1);
  const laid = one.painter.state().layouts;
  const twice = recorder();
  renderPage(twice.ctx, one.page, { scale: 2.5, charts: one.painter, background: null });
  const small = one.recorded.rects.filter((rect) => rect.fill === GREEN);
  const large = twice.rects.filter((rect) => rect.fill === GREEN);
  assert.equal(large.length, small.length);
  small.forEach((bar, i) => {
    assert.ok(Math.abs(large[i].x - bar.x * 2.5) < 0.01 && Math.abs(large[i].y - bar.y * 2.5) < 0.01, 'the place is scaled');
    assert.ok(Math.abs(large[i].width - bar.width * 2.5) < 0.01 && Math.abs(large[i].height - bar.height * 2.5) < 0.01, 'the size is scaled');
  });
  assert.equal(one.painter.state().layouts, laid, 'painting again, at another scale, lays nothing out again');
  assert.equal(one.painter.state().values, 1, 'the cached values are read once');
});

test('painting: the words of a chart are painted with the font list of the document, in the frame', async () => {
  const { image, recorded } = await chartPage(1, title('Tea and coffee'));
  const shown = recorded.texts.find((entry) => entry.text === 'Tea and coffee');
  assert.ok(shown, `the title is on the canvas: ${JSON.stringify(recorded.texts.map((entry) => entry.text))}`);
  assert.match(shown.font, /doc-chain$/);
  assert.ok(shown.x >= image.x && shown.x <= image.x + image.width && shown.y >= image.y && shown.y <= image.y + image.height, 'the title is inside the frame');
});

test('painting: without a painter of charts, and when painting a chart fails, the placeholder is drawn and the context is as it was', async () => {
  const { page, image } = await chartPage();
  const plain = recorder();
  renderPage(plain.ctx, page, { background: null });
  assert.ok(plain.rects.some((rect) => rect.fill === PLACEHOLDER && rect.width === image.width && rect.height === image.height), 'a placeholder of the size of the frame');
  assert.equal(plain.rects.filter((rect) => rect.fill === GREEN).length, 0);

  const failing = recorder();
  let asked = 0;
  renderPage(failing.ctx, page, {
    background: null,
    charts: {
      paint: (ctx) => {
        asked++;
        ctx.fillStyle = '#123456';
        throw new Error('no chart today');
      },
    },
  });
  assert.equal(asked, 1);
  assert.ok(failing.rects.some((rect) => rect.fill === PLACEHOLDER && rect.width === image.width), 'the placeholder stands for the chart that failed');
  assert.ok(!failing.rects.some((rect) => rect.fill === '#123456'), 'the fill the failed painter left is not what the placeholder is drawn with');
  assert.equal(failing.balance(), 0, 'every save has its restore');
});

test('painting: a chart without cached values paints its frame and no columns', async () => {
  const laid = await lay({ body: p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart({ uncached: true }) }])) });
  const recorded = recorder();
  renderPage(recorded.ctx, laid.pages[0], { charts: new DocChartPainter({ cssFont }), background: null });
  assert.equal(recorded.rects.filter((rect) => rect.fill === GREEN).length, 0);
  assert.ok(!recorded.rects.some((rect) => rect.fill === PLACEHOLDER), 'the chart is drawn, empty: no placeholder');
  assert.ok(recorded.clips.length > 0, 'the frame of the chart was painted');
});

// --------------------------------------------------------------- lossless

/** Compressed bytes of the entries, by name. */
function rawEntries(bytes: Uint8Array): Map<string, Uint8Array> {
  const zip = ZipArchive.read(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of zip.names()) {
    const entry = zip.entry(name);
    assert.ok(entry, name);
    out.set(name, entry.raw as Uint8Array);
  }
  return out;
}

const CHART_ENTRIES = ['word/charts/chart1.xml', 'word/charts/_rels/chart1.xml.rels', 'word/embeddings/Microsoft_Excel_Sheet1.xlsx', 'word/_rels/document.xml.rels'];

test('lossless: open and save gives the bytes of the file', async () => {
  const bytes = await buildSpec({ body: p('text') + p(inlineChart('rIdC1', 300, 200)) + p(floatingChart('rIdC1', 0, 0, 100, 80)), theme: THEME, ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const file = await openFile(bytes);
  assert.equal(file.kind, 'document');
  const doc = await file.document();
  assert.ok(doc.charts && doc.charts.size === 1, 'the chart was read');
  assert.deepEqual(await file.save(), bytes);
});

test('lossless: after an edit of the text the chart, its relationships and its workbook are the entries of the file', async () => {
  const bytes = await buildSpec({ body: p('text') + p(inlineChart('rIdC1', 300, 200)), ...(await chartParts([{ rel: 'rIdC1', xml: columnsChart() }])) });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  history.applyEdit(insertText(doc, { story: 'body', path: [0], offset: 4 }, ' added'), undefined);
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts, ['/word/document.xml']);
  const saved = await pkg.save();
  const before = rawEntries(bytes);
  const after = rawEntries(saved);
  assert.deepEqual(Array.from(after.keys()).sort(), Array.from(before.keys()).sort());
  for (const name of CHART_ENTRIES) assert.deepEqual(after.get(name), before.get(name), name);
  // The frame of the chart is written as it was read.
  const again = await readDocx(await OpcPackage.open(saved));
  assert.equal(documentText(again), 'text added\n');
  assert.deepEqual(firstValues((again.charts as Map<string, Chart>).get(CHART1)), [50, 100, 25]);
  const xml = new TextDecoder().decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf(inlineChartDrawing('rIdC1', 300, 200)) >= 0, 'the drawing is in the document part, character for character');
});
