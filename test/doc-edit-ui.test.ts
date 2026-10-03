// The parts of the document editor that need no browser: the line marks of
// the layout, finding places on the pages (DocLines), the copy of a selection
// (doc-clipboard) and the mirror the typed text is read from (DocInput).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutPage, LineMark } from '../src/layout/doc/types';
import type { DocModel } from '../src/model/doc/types';
import { DocLines } from '../src/ui/doc-caret';
import { hasCopyMark, sliceHtml, sliceOf } from '../src/ui/doc-clipboard';
import { DocInput, type InputChange, type InputContext, type InputHost } from '../src/ui/doc-input';
import { fakeMeasurer, p, r, sectPr, STYLES, tbl, tc, tr, words } from './doc-layout-fixtures';
import { readSpec } from './docx-fixtures';

async function laid(body: string): Promise<{ doc: DocModel; pages: LayoutPage[]; lines: DocLines }> {
  const doc = await readSpec({ styles: STYLES, body: body + sectPr() });
  const layout = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  layout.layout();
  return { doc, pages: layout.pages, lines: new DocLines(() => layout.pages, fakeMeasurer) };
}

const marks = (page: LayoutPage): LineMark[] => page.items.filter((item): item is LineMark => item.type === 'mark');

test('line marks: every line of a paragraph, with the offsets it covers; empty paragraphs and alignment', async () => {
  // 40 characters to a line: "w000 ... w009" is 49 characters, two lines.
  const { pages } = await laid(p(words(10)) + p('') + p('centred', '<w:jc w:val="center"/>') + p('', '<w:jc w:val="right"/>'));
  const list = marks(pages[0]);
  assert.deepEqual(
    list.map((mark) => [mark.source.path.join('.'), mark.source.offset, mark.length, mark.last]),
    [
      ['0', 0, 40, false],
      ['0', 40, 9, true],
      ['1', 0, 0, true],
      ['2', 0, 7, true],
      ['3', 0, 0, true],
    ],
  );
  // Lines 16 px high at y 40, 56, ...; the text starts at the left margin of the column.
  assert.deepEqual(
    list.map((mark) => [mark.x, mark.y, mark.height]),
    [
      [0 + 40, 40, 16],
      [40, 56, 16],
      [40, 72, 16],
      [40 + (320 - 56) / 2, 88, 16],
      [360, 104, 16],
    ],
  );
  // The width leaves the spaces at the end of a line out.
  assert.equal(list[0].width, 39 * 8);
  // A view without the option has no marks: nothing changes for it.
  const doc = await readSpec({ styles: STYLES, body: p(words(10)) + sectPr() });
  const plain = new DocLayout(doc, fakeMeasurer);
  plain.layout();
  assert.equal(marks(plain.pages[0]).length, 0);
});

test('line marks of table cells, and of header rows repeated on the next page', async () => {
  let rows = tr(tc('Head A') + tc('Head B'), '<w:tblHeader/>');
  for (let i = 0; i < 30; i++) rows += tr(tc(`Row ${i}`) + tc(''));
  const { pages } = await laid(tbl(rows, [2400, 2400]) + p('after'));
  assert.ok(pages.length >= 2);
  const first = marks(pages[0]);
  assert.deepEqual(first[0].source.path, [0, 0, 0, 0]);
  assert.deepEqual(first[1].source.path, [0, 0, 1, 0]);
  // An empty cell has a mark of its own.
  assert.ok(first.some((mark) => mark.source.path.join('.') === '0.1.1.0' && mark.length === 0));
  const second = marks(pages[1]);
  assert.ok(second[0].repeat, 'the repeated header row is marked as a copy');
  assert.ok(second.some((mark) => !mark.repeat && mark.source.path[1] > 1));
});

test('DocLines: the caret of every place of a paragraph, the place of every point, the line above and below', async () => {
  const { pages, lines } = await laid(p(words(10)) + p('short') + p(words(12, 4, 100)));
  const text = words(10);
  // Every offset has a caret, left to right on a line, and the two lines meet where the first one wraps.
  let previous = -1;
  for (let offset = 0; offset <= text.length; offset++) {
    const box = lines.caret({ story: 'body', path: [0], offset });
    assert.ok(box, `offset ${offset}`);
    if (offset === 40) {
      assert.equal(box.top, 56 + 16 - 16 * 0.8 - 16 * 0.2 - 0, 'the start of the second line');
      assert.equal(box.x, 40);
      previous = box.x;
      continue;
    }
    assert.ok(box.x >= previous, `offset ${offset}: ${box.x} after ${previous}`);
    previous = box.x;
  }
  // The end of the first line, asked for as such.
  const end = lines.caret({ story: 'body', path: [0], offset: 40 }, 'end');
  assert.equal(end!.top, 40);
  assert.equal(end!.x, 40 + 39 * 8);
  // A point maps back to the place nearest to it.
  // (39, before the space that ends the first line, is where 40 is too: the space takes no room there.)
  for (const offset of [0, 3, 17, 38, 41, 48]) {
    const box = lines.caret({ story: 'body', path: [0], offset })!;
    const hit = lines.hit(0, box.x + 1, box.top + 8);
    assert.equal(hit!.place.offset, offset, `point of offset ${offset}`);
    assert.deepEqual(hit!.place.path, [0]);
  }
  // Right of a line that wraps: its end, seen as the end of that line.
  const right = lines.hit(0, 395, 45)!;
  assert.deepEqual([right.place.offset, right.affinity], [40, 'end']);
  // Below the text: the last line; in the margin: the line beside the point.
  assert.equal(lines.hit(0, 100, 390)!.place.path[0], 2);
  assert.equal(lines.hit(0, 2, 60)!.place.offset, 40);
  // Down from the first line keeps x; from the short line too, onto a longer one.
  const start = lines.lineOf({ story: 'body', path: [0], offset: 12 })!;
  const down = lines.vertical(start, 1, 40 + 12 * 8)!;
  assert.equal(lines.placeAt(down.line, 40 + 12 * 8).place.offset, 52 - 3);
  const short = lines.lineOf({ story: 'body', path: [1], offset: 2 })!;
  assert.equal(lines.vertical(short, 1, 40 + 30 * 8)!.line.mark.source.path[0], 2);
  assert.equal(lines.vertical(short, -1, 0)!.line.mark.source.offset, 40);
  assert.equal(pages.length, 1);
});

test('DocLines: the page of a place is found among many pages, by halves', async () => {
  let body = '';
  for (let i = 0; i < 400; i++) body += p(words(15, 4, i * 20));
  const { pages, lines } = await laid(body);
  assert.ok(pages.length > 30);
  for (const block of [0, 1, 150, 233, 399]) {
    const at = lines.lineOf({ story: 'body', path: [block], offset: 30 });
    assert.ok(at, `block ${block}`);
    const page = pages[at.page];
    assert.ok(page.items.some((item) => item.type === 'mark' && item.source.path[0] === block && item.source.offset <= 30));
  }
  // A paragraph that goes on over a page break: the line on the next page.
  const split = marks(pages[1])[0];
  if (split.source.offset > 0) {
    const at = lines.lineOf({ story: 'body', path: split.source.path, offset: split.source.offset + 1 });
    assert.equal(at!.page, 1);
  }
  assert.equal(lines.lineOf({ story: 'body', path: [400], offset: 0 }), undefined);
});

test('the copy of a selection: runs with the formatting in effect, alignment of whole paragraphs, HTML of escaped text', async () => {
  const doc = await readSpec({
    styles: STYLES,
    body: p(r('Plain ') + r('<b>&amp', '<w:b/>') + r(' it', '<w:i/><w:color w:val="FF0000"/>'), '<w:jc w:val="center"/>') + p(r('two', '<w:u w:val="single"/>')) + sectPr(),
  });
  const slice = sliceOf(doc, { story: 'body', path: [0], offset: 2 }, { story: 'body', path: [1], offset: 2 })!;
  assert.equal(slice.paragraphs.length, 2);
  assert.deepEqual(
    slice.paragraphs[0].runs.map((run) => [run.text, run.format.bold, run.format.italic, run.format.color]),
    [
      ['ain ', false, false, 'auto'],
      ['<b>&amp', true, false, 'auto'],
      [' it', false, true, 'FF0000'],
    ],
  );
  assert.equal(slice.paragraphs[0].align, 'center');
  assert.deepEqual(
    slice.paragraphs[1].runs.map((run) => [run.text, run.format.underline]),
    [['tw', true]],
  );
  const html = sliceHtml(slice, 'abc123');
  assert.ok(html.startsWith('<div data-bym-doc-copy="abc123">'));
  assert.ok(html.indexOf('<b>&lt;b&gt;&amp;amp</b>') >= 0, html);
  assert.ok(html.indexOf('color:#FF0000') >= 0);
  assert.ok(html.indexOf('text-align:center') >= 0);
  assert.equal(html.indexOf('<b>&amp'), -1);
  assert.ok(hasCopyMark(html, 'abc123'));
  assert.ok(!hasCopyMark(html, 'other'));
  assert.ok(!hasCopyMark(undefined, 'abc123'));
  // A font name with characters a style attribute would take as more than a name is cut down to the name.
  const evil = sliceHtml({ paragraphs: [{ runs: [{ text: 'x', format: { font: "A';background:url(x)" } }] }] }, 'm');
  assert.equal(evil.indexOf('url('), -1);
  assert.ok(evil.indexOf("font-family:'Abackgroundurlx'") >= 0, evil);
  // Cells of a table.
  assert.equal(sliceHtml({ paragraphs: [], table: [['a<', 'b'], ['c', 'd\ne']] }, 'm'), '<div data-bym-doc-copy="m"><table><tr><td>a&lt;</td><td>b</td></tr><tr><td>c</td><td>d<br>e</td></tr></table></div>');
});

/** A textarea as far as DocInput uses one. */
function fakeField(): HTMLTextAreaElement & { fire(type: string, init?: object): void } {
  const listeners = new Map<string, ((event: unknown) => void)[]>();
  const field = {
    value: '',
    selectionStart: 0,
    selectionEnd: 0,
    addEventListener(type: string, listener: (event: unknown) => void) {
      listeners.set(type, (listeners.get(type) || []).concat([listener]));
    },
    removeEventListener(type: string, listener: (event: unknown) => void) {
      listeners.set(type, (listeners.get(type) || []).filter((entry) => entry !== listener));
    },
    getRootNode() {
      return { activeElement: field };
    },
    setSelectionRange(start: number, end: number) {
      field.selectionStart = start;
      field.selectionEnd = end;
    },
    fire(type: string, init: object = {}) {
      const event = { type, cancelable: false, defaultPrevented: false, preventDefault() {}, ...init } as { preventDefault(): void; defaultPrevented: boolean; cancelable: boolean };
      event.preventDefault = () => {
        event.defaultPrevented = true;
      };
      for (const listener of listeners.get(type) || []) listener(event);
      return event;
    },
  };
  return field as unknown as HTMLTextAreaElement & { fire(type: string, init?: object): void };
}

function harness(context: InputContext | undefined) {
  const field = fakeField();
  const changes: [InputChange, string | undefined][] = [];
  const typed: string[] = [];
  const host: InputHost = {
    context: () => context,
    change: (change, composition) => changes.push([change, composition]),
    type: (text) => typed.push(text),
    remove: (direction, unit) => typed.push(`remove ${direction} ${unit}`),
    paragraph: () => typed.push('paragraph'),
    undo: () => typed.push('undo'),
    redo: () => typed.push('redo'),
    compositionStart: () => typed.push('start'),
    compositionEnd: () => typed.push('end'),
  };
  const input = new DocInput(field, host);
  input.sync();
  /** The browser puts `text` in place of [from, to) of the field and says so. */
  const edit = (from: number, to: number, text: string) => {
    field.value = field.value.slice(0, from) + text + field.value.slice(to);
    field.selectionStart = field.selectionEnd = from + text.length;
    (field as unknown as { fire(type: string): void }).fire('input');
  };
  return { field, changes, typed, input, edit };
}

test('DocInput: keys that can be cancelled are done by the host; the field mirrors the text around the caret', () => {
  const { field, typed } = harness({ before: 'Hello', after: ' world', selected: '', atStart: true, atEnd: false });
  assert.equal(field.value, '\nHello world');
  assert.deepEqual([field.selectionStart, field.selectionEnd], [6, 6]);
  const fire = (field as unknown as { fire(type: string, init: object): { defaultPrevented: boolean } }).fire;
  assert.ok(fire('beforeinput', { inputType: 'insertText', data: 'x', cancelable: true }).defaultPrevented);
  assert.ok(fire('beforeinput', { inputType: 'deleteWordBackward', cancelable: true }).defaultPrevented);
  assert.ok(fire('beforeinput', { inputType: 'insertParagraph', cancelable: true }).defaultPrevented);
  assert.ok(fire('beforeinput', { inputType: 'historyUndo', cancelable: true }).defaultPrevented);
  assert.ok(fire('beforeinput', { inputType: 'insertFromPaste', cancelable: true }).defaultPrevented);
  // What cannot be cancelled, or is a replacement, is left to the browser and read from the field afterwards.
  assert.ok(!fire('beforeinput', { inputType: 'insertReplacementText', data: 'y', cancelable: true }).defaultPrevented);
  assert.ok(!fire('beforeinput', { inputType: 'insertText', data: 'z', cancelable: false }).defaultPrevented);
  assert.deepEqual(typed, ['x', 'remove -1 word', 'paragraph', 'undo']);
});

test('DocInput: what the browser changes in the field is read back as characters removed around the caret and text put in', () => {
  const { changes, edit, field } = harness({ before: 'aaa', after: 'bb', selected: '', atStart: true, atEnd: true });
  // "\naaa|bb\n": a Backspace in a run of the same letters is the one before the caret.
  edit(3, 4, '');
  assert.deepEqual(changes[0][0], { removeBefore: 1, removeAfter: 0, replaceSelection: false, text: '', caretBack: 0, composing: 0 });
  // The mirror was filled again from the host, which says the text is unchanged.
  assert.equal(field.value, '\naaabb\n');
  edit(4, 4, 'xy');
  assert.deepEqual(changes[1][0], { removeBefore: 0, removeAfter: 0, replaceSelection: false, text: 'xy', caretBack: 0, composing: 0 });
  // A word before the caret replaced as a whole (a correction): what stays the same at its start is not part of the change.
  edit(1, 4, 'abc');
  assert.deepEqual(changes[2][0], { removeBefore: 2, removeAfter: 0, replaceSelection: false, text: 'bc', caretBack: 0, composing: 0 });
  // The line end that stands for the start of the paragraph: one character more than the paragraph has before the caret.
  field.value = '\naaabb\n';
  edit(0, 4, '');
  assert.equal(changes[3][0].removeBefore, 4);
  // Delete after the caret, up to the line end that stands for the end of the paragraph.
  edit(4, 7, '');
  assert.equal(changes[4][0].removeAfter, 3);
});

test('DocInput: a composition is read as it goes and ends with the host told; a selection is replaced as a whole', () => {
  const { changes, typed, edit, field, input } = harness({ before: 'ab', after: 'cd', selected: '', atStart: false, atEnd: false });
  const fire = (field as unknown as { fire(type: string, init?: object): void }).fire;
  fire('compositionstart');
  assert.ok(input.composing);
  fire('compositionupdate', { data: 'k' });
  edit(2, 2, 'k');
  // During a composition the field is not filled again: the browser's composition would end.
  assert.equal(field.value, 'abkcd');
  fire('compositionupdate', { data: 'ka' });
  edit(2, 3, 'ka');
  assert.equal(changes[1][0].text, 'a');
  assert.equal(changes[1][0].composing, 2);
  assert.equal(changes[1][1], changes[0][1]);
  assert.ok(changes[0][1] !== undefined && changes[0][1].indexOf('ime:') === 0);
  fire('compositionupdate', { data: 'か' });
  edit(2, 4, 'か');
  assert.deepEqual([changes[2][0].removeBefore, changes[2][0].text], [2, 'か']);
  fire('compositionend', { data: 'か' });
  assert.ok(!input.composing);
  assert.deepEqual(typed, ['start', 'end']);
  // The next composition is another step.
  fire('compositionstart');
  edit(3, 3, 'x');
  assert.notEqual(changes[3][1], changes[0][1]);
  input.commit();
  assert.ok(!input.composing);

  const selected = harness({ before: 'ab', after: 'cd', selected: ' ', atStart: false, atEnd: false });
  assert.deepEqual([selected.field.selectionStart, selected.field.selectionEnd], [2, 3]);
  selected.edit(2, 3, 'Q');
  assert.deepEqual(selected.changes[0][0], { removeBefore: 0, removeAfter: 0, replaceSelection: true, text: 'Q', caretBack: 0, composing: 0 });
  // Without anything to type into, the field holds a blank that stays selected (for the copy event of some browsers).
  const none = harness(undefined);
  assert.equal(none.field.value, ' ');
  assert.deepEqual([none.field.selectionStart, none.field.selectionEnd], [0, 1]);
});

test('text in capitals keeps the offsets of the paragraph: a letter whose capital is longer stays as it is', async () => {
  // 40 characters to a line: the first line ends inside the text, so the second line shows where offsets would go wrong.
  const text = 'straße ﬁne ß words that go on and on to the next line ßß end';
  const doc = await readSpec({ styles: STYLES, body: p(r(text, '<w:caps/>')) + sectPr() });
  const layout = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  layout.layout();
  const items = layout.pages[0].items.filter((item): item is import('../src/layout/doc/types').TextItem => item.type === 'text' && item.source !== undefined);
  assert.ok(items.length >= 2);
  let expected = 0;
  for (const item of items) {
    const source = item.source!;
    assert.equal(source.offset, expected, `"${item.text}" starts where the text before it ends`);
    assert.ok(source.offset + item.text.length <= text.length, 'inside the paragraph');
    // Every character of the item is the capital of the character of the model at its place, or that character.
    for (let i = 0; i < item.text.length; i++) {
      const model = text.charAt(source.offset + i);
      assert.ok(item.text.charAt(i) === model.toUpperCase() || item.text.charAt(i) === model, `${item.text.charAt(i)} for ${model}`);
    }
    expected = source.offset + item.text.length;
  }
  assert.equal(expected, text.length);
  assert.ok(items[0].text.indexOf('STRAßE') === 0, items[0].text);
  // The marks cover the paragraph without a gap.
  const lines = marks(layout.pages[0]);
  assert.equal(lines[0].source.offset + lines[0].length, lines[1].source.offset);
  assert.equal(lines[lines.length - 1].source.offset + lines[lines.length - 1].length, text.length);
});

test('a picture in a line: the mark knows where it is; a caret before it is on its left, after it on its right; a press on it is a press on it', async () => {
  const PX = 9525;
  const picture =
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${40 * PX}" cy="${16 * PX}"/><wp:docPr id="1" name="P"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
    `<pic:blipFill><a:blip r:embed="rIdNone"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${40 * PX}" cy="${16 * PX}"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const { pages, lines } = await laid(p(r('ab') + picture + r('cd')));
  const mark = marks(pages[0])[0];
  assert.deepEqual(mark.boxes, [{ dx: 16, width: 40, offset: 2 }]);
  // "ab" is 16 px, the picture 40, then "cd".
  const before = lines.caret({ story: 'body', path: [0], offset: 2, before: true })!;
  const after = lines.caret({ story: 'body', path: [0], offset: 2 })!;
  assert.equal(before.x, 40 + 16);
  assert.equal(after.x, 40 + 56);
  const left = lines.hit(0, 40 + 20, 50)!;
  assert.deepEqual([left.place.offset, left.place.before, left.box], [2, true, true]);
  const right = lines.hit(0, 40 + 50, 50)!;
  assert.deepEqual([right.place.offset, right.place.before, right.box], [2, undefined, true]);
  // Beside it, text as usual.
  const text = lines.hit(0, 40 + 60, 50)!;
  assert.deepEqual([text.place.offset, text.box], [2, undefined]);
  assert.equal(lines.hit(0, 40 + 76, 50)!.place.offset, 4);
});
