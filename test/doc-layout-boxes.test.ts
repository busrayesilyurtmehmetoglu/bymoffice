// Text boxes in the document layout: the text of a box is laid out for one
// drawing, and all boxes of a document together only up to a limit.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { DocLayoutOptions, TextItem } from '../src/layout/doc/types';
import type { Block, DocModel, DrawingItem, Paragraph, Run } from '../src/model/doc/types';
import { simpleDoc } from './doc-binary-fixtures';
import { fakeMeasurer } from './doc-layout-fixtures';

const para = (text: string): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] });

/** An inline drawing of 2 x 0.5 inches that shows the blocks as the text of its box. */
const boxDrawing = (blocks: Block[]): DrawingItem => ({
  type: 'drawing',
  placement: 'inline',
  kind: 'shape',
  xml: { type: 'element', name: 'doc:drawing', attrs: new Map(), children: [] },
  width: 2 * 914400,
  height: 457200,
  textBoxes: [blocks],
});

/** A document of one paragraph for each drawing. */
function modelOf(drawings: DrawingItem[]): DocModel {
  const model = simpleDoc('body\r').read();
  for (const drawing of drawings) {
    const run: Run = { type: 'run', props: {}, content: [drawing] };
    model.body.blocks.push({ type: 'paragraph', props: {}, children: [run] });
  }
  return model;
}

function textsOf(model: DocModel, options?: DocLayoutOptions): { texts: string[]; items: number; elapsed: number } {
  const layout = new DocLayout(model, fakeMeasurer, options);
  const started = Date.now();
  assert.equal(layout.layout().done, true);
  const elapsed = Date.now() - started;
  const texts: string[] = [];
  let items = 0;
  for (const page of layout.pages) {
    items += page.items.length;
    for (const item of page.items) if (item.type === 'text') texts.push((item as TextItem).text);
  }
  return { texts, items, elapsed };
}

test('doc layout: the text of a box that many drawings share is laid out for the first of them only', () => {
  // 200 paragraphs of 100 characters, shown by 400 drawings: laid out for each it is eight million characters.
  const shared: Block[] = [];
  for (let i = 0; i < 200; i++) shared.push(para('w' + String(i).padStart(3, '0') + ' ' + 'x'.repeat(95)));
  const drawings: DrawingItem[] = [];
  for (let i = 0; i < 400; i++) drawings.push(boxDrawing(shared));
  const { texts, items, elapsed } = textsOf(modelOf(drawings));
  // Every paragraph of the box is on the pages once.
  for (const marker of ['w000', 'w100', 'w199']) assert.equal(texts.filter((text) => text.indexOf(marker) === 0).length, 1, marker);
  assert.ok(items < 5000, `${items} items`);
  assert.ok(elapsed < 5000, `took ${elapsed} ms`);

  // The same drawing object in two paragraphs is one drawing: it shows its text in both.
  const one = boxDrawing([para('twice')]);
  assert.equal(textsOf(modelOf([one, one])).texts.filter((text) => text === 'twice').length, 2);
  // Two drawings with a box each show both.
  assert.deepEqual(textsOf(modelOf([boxDrawing([para('first')]), boxDrawing([para('second')])])).texts, ['body', 'first', 'second']);
});

test('doc layout: the text boxes of a document are laid out up to a limit; the ones after it stay empty', () => {
  // Each box weighs 116: a paragraph and a hundred characters.
  const boxes = ['a', 'b', 'c'].map((letter) => boxDrawing([para(letter.repeat(100))]));
  const letters = (options?: DocLayoutOptions) =>
    textsOf(modelOf(boxes), options)
      .texts.filter((text) => text !== 'body')
      .map((text) => text.charAt(0))
      .filter((letter, i, all) => all.indexOf(letter) === i);
  assert.deepEqual(letters(), ['a', 'b', 'c']);
  // The limit is looked at before a box is laid out: the box that crosses it is the last one shown.
  assert.deepEqual(letters({ maxTextBoxWork: 200 }), ['a', 'b']);
  assert.deepEqual(letters({ maxTextBoxWork: 116 }), ['a']);
  assert.deepEqual(letters({ maxTextBoxWork: 117 }), ['a', 'b']);
  assert.deepEqual(letters({ maxTextBoxWork: 0 }), []);
  // The drawings keep their place and size without their text.
  const empty = new DocLayout(modelOf(boxes), fakeMeasurer, { maxTextBoxWork: 0 });
  empty.layout();
  const full = new DocLayout(modelOf(boxes), fakeMeasurer);
  full.layout();
  assert.equal(empty.pages.length, full.pages.length);

  // Tables in a box count by their rows and the text of their cells.
  const table: Block = { type: 'table', props: {}, grid: [1000], rows: [{ props: {}, cells: [{ props: {}, blocks: [para('t'.repeat(50))] }] }] };
  const withTable = [boxDrawing([table]), boxDrawing([para('z')])];
  const shown = (limit: number) => textsOf(modelOf(withTable), { maxTextBoxWork: limit }).texts.some((text) => text === 'z');
  // 16 for the table, 16 for the row, 16 and 50 for the paragraph in the cell.
  assert.equal(shown(98), false);
  assert.equal(shown(99), true);
});
