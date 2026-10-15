// The panel of comments of a document (src/ui/doc-comments-panel.ts): what it
// shows, in what order, what it asks of the view, how the keys move in it, and
// what it costs with many comments. Its DOM is a small one made here: only
// what the panel uses of one. What needs the pages (a thread picked in the
// panel shown in the document) is in e2e/doc-comments-panel.spec.ts.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { listComments, type DocComment } from '../src/ops/doc/index';
import {
  CommentsPanel,
  PANEL_BATCH,
  PANEL_SHORT_TEXT,
  commentDateText,
  commentQuote,
  coveredBelow,
  paragraphTexts,
  quoteText,
  shownText,
  shownThreads,
  stepIndex,
  type CommentThread,
  type CommentsPanelHost,
} from '../src/ui/doc-comments-panel';
import { STRINGS_EN, STRINGS_TR, type UiStrings } from '../src/ui/strings';
import { CommentBalloon, fillMarked } from '../src/ui/doc-comments';
import { cEnd, cRef, cStart, openCommented } from './doc-comments-fixtures';
import { p, r } from './doc-ops-fixtures';

// ------------------------------------------------------------------ fake DOM

interface FakeEvent {
  type: string;
  target: FakeElement;
  key?: string;
  code?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  altKey?: boolean;
  shiftKey?: boolean;
  defaultPrevented: boolean;
  stopped: boolean;
  preventDefault(): void;
  stopPropagation(): void;
}

class FakeStyle {
  [name: string]: unknown;
  cssText = '';
  display = '';
}

class FakeElement {
  children: FakeElement[] = [];
  parentNode: FakeElement | null = null;
  readonly attributes = new Map<string, string>();
  readonly style = new FakeStyle();
  readonly listeners = new Map<string, ((event: FakeEvent) => void)[]>();
  className = '';
  type = '';
  title = '';
  value = '';
  checked = false;
  rows = 0;
  maxLength = 0;
  tabIndex = -1;
  scrollTop = 0;
  clientHeight = 0;
  scrollHeight = 0;
  offsetTop = 0;
  offsetHeight = 0;
  offsetParent: FakeElement | null = null;
  clientWidth = 0;
  private text = '';
  constructor(readonly tagName: string) {}
  get textContent(): string {
    return this.text + this.children.map((child) => child.textContent).join('');
  }
  set textContent(value: string) {
    this.text = value;
    for (const child of this.children) child.parentNode = null;
    this.children = [];
  }
  set innerHTML(_value: string) {
    throw new Error('The panel must not set markup');
  }
  get firstChild(): FakeElement | null {
    return this.children[0] || null;
  }
  get nextSibling(): FakeElement | null {
    const parent = this.parentNode;
    if (!parent) return null;
    return parent.children[parent.children.indexOf(this) + 1] || null;
  }
  appendChild(child: FakeElement): FakeElement {
    return this.insertBefore(child, null);
  }
  insertBefore(child: FakeElement, before: FakeElement | null): FakeElement {
    if (child.parentNode) child.parentNode.removeChild(child);
    const at = before ? this.children.indexOf(before) : -1;
    if (at < 0) this.children.push(child);
    else this.children.splice(at, 0, child);
    child.parentNode = this;
    return child;
  }
  removeChild(child: FakeElement): FakeElement {
    const at = this.children.indexOf(child);
    if (at >= 0) this.children.splice(at, 1);
    child.parentNode = null;
    if (fakeDocument.activeElement && child.contains(fakeDocument.activeElement)) fakeDocument.activeElement = fakeDocument.body;
    return child;
  }
  contains(other: FakeElement | null): boolean {
    for (let at = other; at; at = at.parentNode) if (at === this) return true;
    return false;
  }
  closest(selector: string): FakeElement | null {
    for (let at: FakeElement | null = this; at; at = at.parentNode) if (at.tagName === selector) return at;
    return null;
  }
  setAttribute(name: string, value: string): void {
    this.attributes.set(name, String(value));
  }
  getAttribute(name: string): string | null {
    const value = this.attributes.get(name);
    return value === undefined ? null : value;
  }
  removeAttribute(name: string): void {
    this.attributes.delete(name);
  }
  addEventListener(type: string, listener: (event: FakeEvent) => void): void {
    const list = this.listeners.get(type) || [];
    list.push(listener);
    this.listeners.set(type, list);
  }
  removeEventListener(type: string, listener: (event: FakeEvent) => void): void {
    const list = this.listeners.get(type) || [];
    const at = list.indexOf(listener);
    if (at >= 0) list.splice(at, 1);
  }
  focus(): void {
    fakeDocument.activeElement = this;
  }
  getRootNode(): typeof fakeDocument {
    return fakeDocument;
  }
  /** Only the selector the balloon asks for its first button. */
  querySelector(selector: string): FakeElement | null {
    return all(this, (element) => element.tagName === selector)[0] || null;
  }
  /** Where the element is on the screen: what a test lays out (see `layout`), else at the top. */
  getBoundingClientRect(): { top: number; bottom: number; left: number; right: number; width: number; height: number } {
    const box = layout(this) || { top: 0, height: this.offsetHeight };
    return { top: box.top, bottom: box.top + box.height, left: 0, right: 0, width: 0, height: box.height };
  }
}

/** The layout a test gives elements on the screen; none by default. */
let layout: (element: FakeElement) => { top: number; height: number } | undefined = () => undefined;

const fakeDocument = {
  body: new FakeElement('body'),
  activeElement: null as FakeElement | null,
  createElement: (tag: string) => new FakeElement(tag),
  createTextNode: (text: string) => {
    const node = new FakeElement('#text');
    node.textContent = text;
    return node;
  },
};
fakeDocument.activeElement = fakeDocument.body;
(globalThis as unknown as { document: unknown }).document = fakeDocument;

/** Dispatches an event at an element: its listeners and those of its ancestors, until one stops it. */
function fire(target: FakeElement, type: string, init: Partial<FakeEvent> = {}): FakeEvent {
  const event: FakeEvent = {
    type,
    target,
    defaultPrevented: false,
    stopped: false,
    preventDefault() {
      this.defaultPrevented = true;
    },
    stopPropagation() {
      this.stopped = true;
    },
    ...init,
  } as FakeEvent;
  for (let at: FakeElement | null = target; at && !event.stopped; at = at.parentNode) for (const listener of (at.listeners.get(type) || []).slice()) listener(event);
  return event;
}

const key = (target: FakeElement, name: string, init: Partial<FakeEvent> = {}) => fire(target, 'keydown', { key: name, ...init });

function all(root: FakeElement, test: (element: FakeElement) => boolean, out: FakeElement[] = []): FakeElement[] {
  for (const child of root.children) {
    if (test(child)) out.push(child);
    all(child, test, out);
  }
  return out;
}
const byClass = (root: FakeElement, name: string) => all(root, (element) => element.className.split(' ').indexOf(name) >= 0);
const entriesOf = (panel: CommentsPanel) => byClass(panel.element as unknown as FakeElement, 'bym-doc-cpanel-item');
const buttonOf = (entry: FakeElement, action: string, comment?: number) =>
  all(entry, (element) => element.tagName === 'button' && element.getAttribute('data-action') === action && (comment === undefined || element.getAttribute('data-comment') === String(comment)))[0];

// --------------------------------------------------------------- the host

function comment(id: number, fields: Partial<DocComment> = {}): DocComment {
  return {
    id,
    author: 'Ann',
    date: '2026-10-01T09:05:00Z',
    text: `Comment ${id}`,
    done: false,
    story: 'body',
    start: { story: 'body', path: [id], offset: 0 },
    end: { story: 'body', path: [id], offset: 3 },
    plain: true,
    canChange: true,
    ...fields,
  };
}

function thread(id: number, replies = 0, fields: Partial<DocComment> = {}): CommentThread {
  const root = comment(id, fields);
  const answers: DocComment[] = [];
  for (let i = 1; i <= replies; i++) answers.push(comment(id * 1000 + i, { parent: id, author: 'Bob', text: `Answer ${i}`, done: root.done, start: root.start, end: root.end }));
  return { root, replies: answers };
}

interface Calls {
  reveal: [number, boolean][];
  selected: (number | undefined)[];
  submit: [string, number, string][];
  resolve: [number, boolean][];
  remove: number[];
  closed: boolean[];
}

function makeHost(threads: CommentThread[], options: { can?: boolean; strings?: UiStrings; submit?: boolean } = {}) {
  const root = new FakeElement('div');
  fakeDocument.body.appendChild(root);
  const state = { threads, can: options.can !== false, strings: options.strings || STRINGS_EN, quotes: 0 };
  const calls: Calls = { reveal: [], selected: [], submit: [], resolve: [], remove: [], closed: [] };
  const host: CommentsPanelHost = {
    root: root as unknown as HTMLElement,
    strings: () => state.strings,
    threads: () => state.threads,
    canChange: () => state.can,
    quote: (c) => {
      state.quotes++;
      return c.start ? `text of ${c.id}` : '';
    },
    date: (date) => commentDateText(date, state.strings === STRINGS_TR ? LOCALE_TR : LOCALE_EN),
    reveal: (id, toDocument) => calls.reveal.push([id, toDocument]),
    selectionChanged: (id) => calls.selected.push(id),
    submit: (mode, id, text) => {
      calls.submit.push([mode, id, text]);
      return options.submit !== false;
    },
    resolve: (id, done) => calls.resolve.push([id, done]),
    remove: (id) => calls.remove.push(id),
    closed: (refocus) => calls.closed.push(refocus),
  };
  const panel = new CommentsPanel(host);
  return { panel, host, root, state, calls };
}

// ------------------------------------------------------------------ helpers

test('dates of comments are shown in the short date and time of the language, without Intl; what is no date is shown as it is', () => {
  assert.equal(commentDateText('2026-10-01T09:05:00Z', LOCALE_EN), '10/1/2026 9:05');
  assert.equal(commentDateText('2026-10-01T09:05:00Z', LOCALE_TR), '01.10.2026 09:05');
  // The clock of the file, not turned to another zone.
  assert.equal(commentDateText('2026-10-01T23:30:00+03:00', LOCALE_TR), '01.10.2026 23:30');
  assert.equal(commentDateText('2026-10-01', LOCALE_TR), '01.10.2026');
  assert.equal(commentDateText('2026-02-30T10:00:00Z', LOCALE_EN), '2026-02-30T10:00:00Z');
  assert.equal(commentDateText('2026-10-01T25:00:00Z', LOCALE_EN), '2026-10-01T25:00:00Z');
  assert.equal(commentDateText('1601-01-01T00:00:00Z', LOCALE_EN), '1601-01-01T00:00:00Z');
  assert.equal(commentDateText('yesterday', LOCALE_EN), 'yesterday');
  assert.equal(commentDateText(undefined, LOCALE_EN), '');
  assert.equal(commentDateText('x'.repeat(10000), LOCALE_EN).length, 64);
});

test('quotes are one line, cut with an ellipsis, never inside a pair of surrogates', () => {
  assert.equal(quoteText('  one\n\ttwo  three  '), 'one two three');
  assert.equal(quoteText('abcdef', 4), 'abc…');
  const emoji = 'ab😀cd';
  const cut = quoteText(emoji, 4);
  assert.ok(!/[\ud800-\udbff]…/.test(cut), `a surrogate was cut: ${JSON.stringify(cut)}`);
  assert.ok(quoteText('x'.repeat(5_000_000)).length <= 120);
  assert.equal(shownText('y'.repeat(PANEL_SHORT_TEXT + 10), false).length, PANEL_SHORT_TEXT + 1);
  assert.equal(shownText('y'.repeat(PANEL_SHORT_TEXT + 10), true).length, PANEL_SHORT_TEXT + 10);
});

test('the keys of the list step within it, and do not wrap', () => {
  assert.equal(stepIndex(0, 3, 'ArrowDown'), 1);
  assert.equal(stepIndex(2, 3, 'ArrowDown'), 2);
  assert.equal(stepIndex(0, 3, 'ArrowUp'), 0);
  assert.equal(stepIndex(1, 3, 'Home'), 0);
  assert.equal(stepIndex(0, 3, 'End'), 2);
  assert.equal(stepIndex(-1, 3, 'ArrowDown'), 0);
  assert.equal(stepIndex(0, 0, 'ArrowDown'), undefined);
  assert.equal(stepIndex(0, 3, 'a'), undefined);
});

test('the filter hides resolved threads and keeps the order', () => {
  const threads = [thread(1), thread(2, 0, { done: true }), thread(3)];
  assert.deepEqual(shownThreads(threads, true).map((t) => t.root.id), [1, 2, 3]);
  assert.deepEqual(shownThreads(threads, false).map((t) => t.root.id), [1, 3]);
});

test('a quote is the text a comment is about, of one or two paragraphs; none for a point', async () => {
  const body =
    p(r('Before ') + cStart(1) + r('the commented words') + cEnd(1) + cRef(1) + r(' after.')) +
    p(r('One ') + cStart(2) + r('ends') + r('') ) +
    p(r('here') + cEnd(2) + cRef(2) + r(' and on.')) +
    p(r('Far ') + cStart(3) + r('start')) +
    p('middle') +
    p(r('end') + cEnd(3) + cRef(3)) +
    p(r('Point') + cStart(4) + cEnd(4) + cRef(4));
  const { doc } = await openCommented({ body, comments: [1, 2, 3, 4].map((id) => ({ id, text: `c${id}` })) });
  const quotes = new Map(listComments(doc).map((c) => [c.id, commentQuote(doc, c)]));
  assert.equal(quotes.get(1), 'the commented words');
  assert.equal(quotes.get(2), 'ends here');
  assert.equal(quotes.get(3), 'start … end');
  assert.equal(quotes.get(4), '');
  // A place the document does not have quotes nothing, and throws nothing.
  assert.equal(commentQuote(doc, comment(9, { start: { story: 'body', path: [99], offset: 0 }, end: { story: 'body', path: [99], offset: 2 } })), '');
});

// -------------------------------------------------------------------- panel

test('the panel lists every thread in order with author, date, quote, text and answers, as text; it is a labelled complementary region', () => {
  const threads = [thread(1, 2), thread(2, 0, { author: '<b>Eve</b>', text: '<img src=x onerror=alert(1)>' }), thread(3, 0, { start: undefined, end: undefined, story: undefined })];
  const { panel, root } = makeHost(threads);
  const element = panel.element as unknown as FakeElement;
  assert.equal(element.parentNode, root);
  assert.equal(element.getAttribute('role'), 'complementary');
  assert.equal(element.getAttribute('aria-label'), null, 'labelled before it was shown');
  panel.show();
  assert.equal(element.getAttribute('aria-label'), 'Comments');
  assert.equal(element.style.display, 'flex');
  const entries = entriesOf(panel);
  assert.deepEqual(entries.map((entry) => entry.getAttribute('data-thread')), ['1', '2', '3']);
  assert.match(entries[0].textContent, /^Ann · 10\/1\/2026 9:05“text of 1”Comment 12 replies$/);
  // What the file says is text: the markup is shown as it is.
  assert.match(entries[1].textContent, /<b>Eve<\/b>.*<img src=x onerror=alert\(1\)>/);
  assert.match(entries[2].textContent, /Not tied to any text/);
  assert.match(entries[0].getAttribute('aria-label') || '', /^Comment by Ann, 10\/1\/2026 9:05, 2 replies: Comment 1$/);
  assert.equal(byClass(element, 'bym-doc-cpanel-list')[0].getAttribute('role'), 'list');
  assert.ok(entries.every((entry) => entry.getAttribute('role') === 'listitem'));
  // One Tab stop in the list: the first entry while none is selected.
  assert.deepEqual(entries.map((entry) => entry.tabIndex), [0, -1, -1]);
  assert.match(byClass(element, 'bym-doc-cpanel-count')[0].textContent, /^3 threads$/);
});

test('the selected thread shows its answers and what can be done with them; a view that takes no comments shows none of it', () => {
  const { panel, calls, state } = makeHost([thread(1, 2), thread(2)]);
  panel.show();
  panel.select(1);
  assert.deepEqual(calls.selected, [1]);
  let [first, second] = entriesOf(panel);
  assert.equal(first.getAttribute('aria-current'), 'true');
  assert.equal(first.tabIndex, 0);
  assert.equal(second.tabIndex, -1);
  assert.match(first.textContent, /Answer 1.*Answer 2/);
  for (const action of ['edit', 'delete', 'reply', 'resolve']) assert.ok(buttonOf(first, action), `no ${action} button`);
  assert.ok(buttonOf(first, 'edit', 1001) && buttonOf(first, 'delete', 1002), 'the answers have no buttons');
  assert.equal(buttonOf(second, 'reply'), undefined, 'a thread that is not selected has buttons');
  state.can = false;
  panel.refresh();
  [first] = entriesOf(panel);
  assert.equal(all(first, (element) => element.tagName === 'button').length, 0, 'a read-only view offers changes');
  assert.match(first.textContent, /Answer 2/);
});

test('a thread whose marks cannot be changed is not answered or deleted from the panel; a resolved one is reopened', () => {
  const { panel } = makeHost([thread(1, 0, { canChange: false, done: true })]);
  panel.show();
  panel.select(1);
  const [entry] = entriesOf(panel);
  assert.equal(buttonOf(entry, 'reply'), undefined);
  assert.equal(buttonOf(entry, 'delete'), undefined);
  assert.ok(buttonOf(entry, 'edit'));
  assert.ok(buttonOf(entry, 'reopen'));
  assert.equal(buttonOf(entry, 'resolve'), undefined);
  assert.equal(entry.getAttribute('data-done'), 'true');
});

test('answering and editing go through a field: Save hands the text to the view, Escape gives it up, Ctrl+Enter saves', () => {
  const { panel, calls } = makeHost([thread(1, 1)]);
  panel.show();
  panel.select(1);
  let [entry] = entriesOf(panel);
  fire(buttonOf(entry, 'reply'), 'click');
  assert.equal(panel.composing, true);
  [entry] = entriesOf(panel);
  const field = all(entry, (element) => element.tagName === 'textarea')[0];
  assert.ok(field, 'no field');
  assert.equal(fakeDocument.activeElement, field);
  assert.equal(field.getAttribute('aria-label'), 'Reply text');
  assert.equal(buttonOf(entry, 'resolve'), undefined, 'buttons beside the field');
  field.value = 'Yes.';
  fire(field, 'input');
  fire(buttonOf(entry, 'save'), 'click');
  assert.deepEqual(calls.submit, [['reply', 1, 'Yes.']]);
  assert.equal(panel.composing, false);
  [entry] = entriesOf(panel);
  assert.equal(fakeDocument.activeElement, entry, 'the focus is not back on the thread');

  // Editing an answer starts from its text.
  fire(buttonOf(entry, 'edit', 1001), 'click');
  [entry] = entriesOf(panel);
  const edit = all(entry, (element) => element.tagName === 'textarea')[0];
  assert.equal(edit.value, 'Answer 1');
  const escape = key(edit, 'Escape');
  assert.equal(escape.defaultPrevented, true);
  assert.equal(escape.stopped, true, 'the key reached the document');
  assert.equal(panel.composing, false);
  assert.equal(panel.isOpen, true, 'Escape in the field closed the panel');

  fire(buttonOf(entriesOf(panel)[0], 'edit', 1), 'click');
  const again = all(entriesOf(panel)[0], (element) => element.tagName === 'textarea')[0];
  again.value = 'New text';
  fire(again, 'input');
  key(again, 'Enter', { ctrlKey: true });
  assert.deepEqual(calls.submit[1], ['edit', 1, 'New text']);
});

test('what was written stays when the view refuses it, and when the panel is brought up to date meanwhile', () => {
  const { panel, calls, state } = makeHost([thread(1), thread(2)], { submit: false });
  panel.show();
  panel.select(1);
  fire(buttonOf(entriesOf(panel)[0], 'reply'), 'click');
  const field = all(entriesOf(panel)[0], (element) => element.tagName === 'textarea')[0];
  field.value = 'Draft';
  fire(field, 'input');
  // Someone else's change: thread 2 got an answer.
  state.threads = [thread(1), thread(2, 1)];
  panel.refresh();
  const kept = all(entriesOf(panel)[0], (element) => element.tagName === 'textarea')[0];
  assert.equal(kept, field, 'the field was built anew');
  fire(buttonOf(entriesOf(panel)[0], 'save'), 'click');
  assert.deepEqual(calls.submit, [['reply', 1, 'Draft']]);
  assert.equal(panel.composing, true, 'a refused answer was given up');
  assert.equal(all(entriesOf(panel)[0], (element) => element.tagName === 'textarea')[0].value, 'Draft');
});

test('resolve, reopen and delete ask the view; the focus stays in the panel', () => {
  const { panel, calls, state } = makeHost([thread(1, 1), thread(2)]);
  panel.show();
  panel.select(1);
  const resolve = buttonOf(entriesOf(panel)[0], 'resolve');
  resolve.focus();
  fire(resolve, 'click');
  assert.deepEqual(calls.resolve, [[1, true]]);
  // The view resolved it and the panel was brought up to date: Resolve became Reopen, and has the focus.
  state.threads = [thread(1, 1, { done: true }), thread(2)];
  panel.refresh();
  assert.equal(fakeDocument.activeElement, buttonOf(entriesOf(panel)[0], 'reopen'));
  fire(buttonOf(entriesOf(panel)[0], 'delete', 1001), 'click');
  assert.deepEqual(calls.remove, [1001]);
  // The thread went with its first comment: nothing is selected, and the focus is on the list.
  fire(buttonOf(entriesOf(panel)[0], 'delete', 1), 'click');
  state.threads = [thread(2)];
  panel.refresh();
  assert.equal(panel.selectedThread, undefined);
  assert.equal(calls.selected[calls.selected.length - 1], undefined);
  assert.ok((panel.element as unknown as FakeElement).contains(fakeDocument.activeElement), 'the focus left the panel');
});

test('the keys: arrows pick the next thread and keep the focus, Enter goes to the document, Escape closes the panel', () => {
  const { panel, calls } = makeHost([thread(1), thread(2), thread(3)]);
  panel.show(true);
  let entries = entriesOf(panel);
  assert.equal(fakeDocument.activeElement, entries[0]);
  key(entries[0], 'ArrowDown');
  assert.deepEqual(calls.reveal, [[2, false]]);
  entries = entriesOf(panel);
  assert.equal(fakeDocument.activeElement, entries[1]);
  assert.equal(panel.selectedThread, 2);
  key(entries[1], 'End');
  assert.equal(panel.selectedThread, 3);
  key(entriesOf(panel)[2], 'Home');
  assert.equal(panel.selectedThread, 1);
  const enter = key(entriesOf(panel)[0], 'Enter');
  assert.equal(enter.defaultPrevented, true);
  assert.deepEqual(calls.reveal[calls.reveal.length - 1], [1, true]);
  // Keys of the panel are not the document's (it would scroll with them, or take them as shortcuts) — but Ctrl+P is.
  assert.equal(key(entriesOf(panel)[0], 'a').stopped, true);
  assert.equal(key(entriesOf(panel)[0], 'p', { ctrlKey: true }).stopped, false);
  key(entriesOf(panel)[0], 'Escape');
  assert.equal(panel.isOpen, false);
  assert.deepEqual(calls.closed, [true]);
  assert.equal(entriesOf(panel).length, 0, 'a hidden panel keeps its entries');
});

test('the shortcut that opens the panel closes it, and the Close button does too', () => {
  const { panel, calls } = makeHost([thread(1)]);
  panel.show(true);
  key(entriesOf(panel)[0], 'M', { ctrlKey: true, altKey: true, shiftKey: true, code: 'KeyM' });
  assert.equal(panel.isOpen, false);
  panel.show(true);
  const close = all(panel.element as unknown as FakeElement, (element) => element.getAttribute('data-action') === 'close')[0];
  assert.equal(close.getAttribute('aria-label'), 'Close');
  fire(close, 'click');
  assert.equal(panel.isOpen, false);
  assert.equal(calls.closed.length, 2);
});

test('a click on a thread picks it and shows it in the document; the panel keeps the focus', () => {
  const { panel, calls } = makeHost([thread(1), thread(2)]);
  panel.show();
  const text = byClass(entriesOf(panel)[1], 'bym-doc-cpanel-text')[0];
  fire(text, 'click');
  assert.deepEqual(calls.reveal, [[2, false]]);
  assert.equal(panel.selectedThread, 2);
  assert.equal(fakeDocument.activeElement, entriesOf(panel)[1]);
  // The presses of the panel are not the document's.
  assert.equal(fire(byClass(entriesOf(panel)[1], 'bym-doc-cpanel-text')[0], 'mousedown').stopped, true);
});

test('the filter hides resolved threads, says how many, and shows them again', () => {
  const { panel } = makeHost([thread(1, 0, { done: true }), thread(2), thread(3, 0, { done: true })]);
  panel.show();
  const element = panel.element as unknown as FakeElement;
  const box = all(element, (e) => e.tagName === 'input')[0];
  assert.equal(box.checked, true);
  box.checked = false;
  fire(box, 'change');
  assert.deepEqual(entriesOf(panel).map((entry) => entry.getAttribute('data-thread')), ['2']);
  assert.match(byClass(element, 'bym-doc-cpanel-count')[0].textContent, /^1 thread, 2 resolved hidden$/);
  box.checked = true;
  fire(box, 'change');
  assert.equal(entriesOf(panel).length, 3);
});

const marksIn = (root: FakeElement) => all(root, (element) => element.tagName === 'mark');

test('a match of Find is marked in a comment of the selected thread, as text; the filter gives way to it; a changed text or another thread is not marked', () => {
  const { panel, state } = makeHost([thread(1, 1, { text: 'A <b>cat</b> here' }), thread(2, 0, { done: true, text: 'cat' })]);
  const element = panel.element as unknown as FakeElement;
  panel.show();
  panel.select(1);
  panel.markText({ comment: 1, start: 5, end: 8 });
  assert.deepEqual(
    marksIn(element).map((mark) => [mark.className, mark.textContent]),
    [['bym-doc-find-mark', 'cat']],
  );
  // The text around it is text too: nothing of it was read as markup.
  assert.equal(byClass(entriesOf(panel)[0], 'bym-doc-cpanel-text')[0].textContent, 'A <b>cat</b> here');
  assert.equal(all(element, (e) => e.tagName === 'b').length, 0);
  // In an answer.
  panel.markText({ comment: 1001, start: 0, end: 6 });
  assert.deepEqual(marksIn(element).map((mark) => mark.textContent), ['Answer']);
  // A comment that is not in the selected thread, a range beyond the text: no mark.
  panel.markText({ comment: 2, start: 0, end: 3 });
  assert.equal(marksIn(element).length, 0);
  panel.markText({ comment: 1, start: 10, end: 99 });
  assert.equal(marksIn(element).length, 0);
  // A resolved thread the filter hides is shown with the others.
  const box = all(element, (e) => e.tagName === 'input')[0];
  box.checked = false;
  fire(box, 'change');
  assert.deepEqual(entriesOf(panel).map((entry) => entry.getAttribute('data-thread')), ['1']);
  panel.select(2);
  panel.markText({ comment: 2, start: 0, end: 3 });
  assert.equal(box.checked, true);
  assert.deepEqual(entriesOf(panel).map((entry) => entry.getAttribute('data-thread')), ['1', '2']);
  assert.deepEqual(marksIn(element).map((mark) => mark.textContent), ['cat']);
  // The text of the comment changed: the match is no longer where it was.
  state.threads = [thread(1, 1, { text: 'A <b>cat</b> here' }), thread(2, 0, { done: true, text: 'dog' })];
  panel.refresh();
  assert.equal(marksIn(element).length, 0);
  // Another thread picked: the mark goes.
  state.threads = [thread(1, 1, { text: 'A <b>cat</b> here' }), thread(2, 0, { done: true, text: 'cat' })];
  panel.markText({ comment: 2, start: 0, end: 3 });
  assert.equal(marksIn(element).length, 1);
  panel.select(1);
  assert.equal(marksIn(element).length, 0);
  panel.markText(undefined);
  assert.equal(marksIn(element).length, 0);
});

test('the balloon marks a match of Find in a comment of its thread; showing another thread takes the mark away', () => {
  const root = new FakeElement('div');
  const balloon = new CommentBalloon({
    root: root as unknown as HTMLElement,
    strings: () => STRINGS_EN,
    canChange: () => true,
    submit: () => true,
    resolve: () => undefined,
    remove: () => undefined,
    closed: () => undefined,
  });
  balloon.show([comment(1, { text: 'one <i>cat</i>' }), comment(2, { parent: 1, text: 'cat two' })], undefined);
  balloon.markText({ comment: 2, start: 0, end: 3 });
  assert.deepEqual(marksIn(root).map((mark) => [mark.className, mark.textContent]), [['bym-doc-find-mark', 'cat']]);
  assert.equal(all(root, (e) => e.tagName === 'i').length, 0);
  balloon.markText({ comment: 9, start: 0, end: 3 });
  assert.equal(marksIn(root).length, 0);
  balloon.markText({ comment: 1, start: 7, end: 10 });
  assert.deepEqual(marksIn(root).map((mark) => mark.textContent), ['cat']);
  balloon.show([comment(1, { text: 'one <i>cat</i>' })], undefined);
  assert.equal(marksIn(root).length, 0);
});

test('fillMarked: the text as text nodes with the range in a <mark>; never one that cuts a pair of surrogates or lies outside the text', () => {
  const box = new FakeElement('div');
  const target = box as unknown as HTMLElement;
  assert.equal(fillMarked(target, 'abc', { start: 1, end: 2 })!.textContent, 'b');
  assert.deepEqual(box.children.map((child) => [child.tagName, child.textContent]), [['#text', 'a'], ['mark', 'b'], ['#text', 'c']]);
  assert.equal(fillMarked(target, 'abc', { start: 0, end: 3 })!.textContent, 'abc');
  assert.equal(box.children.length, 1);
  for (const bad of [{ start: 2, end: 2 }, { start: -1, end: 1 }, { start: 1, end: 4 }, { start: 2, end: 1 }]) {
    assert.equal(fillMarked(target, 'abc', bad), undefined);
    assert.equal(box.textContent, 'abc');
    assert.equal(box.children.length, 0);
  }
  // "a😀b": the face is two code units; a range that begins or ends between them is not marked.
  assert.equal(fillMarked(target, 'a\u{1F600}b', { start: 2, end: 4 }), undefined);
  assert.equal(fillMarked(target, 'a\u{1F600}b', { start: 0, end: 2 }), undefined);
  assert.equal(fillMarked(target, 'a\u{1F600}b', { start: 1, end: 3 })!.textContent, '\u{1F600}');
});

test('a panel without comments says so; the Turkish one in Turkish', () => {
  const { panel, state } = makeHost([], { strings: STRINGS_TR });
  panel.show();
  const element = panel.element as unknown as FakeElement;
  assert.equal(byClass(element, 'bym-doc-cpanel-empty')[0].textContent, 'Belgede yorum yok');
  assert.equal(element.getAttribute('aria-label'), 'Yorumlar');
  state.threads = [thread(1, 1)];
  panel.refresh();
  assert.match(entriesOf(panel)[0].textContent, /01\.10\.2026 09:05.*1 yanıt/);
  state.strings = STRINGS_EN;
  panel.setStrings();
  assert.equal(element.getAttribute('aria-label'), 'Comments');
  assert.match(entriesOf(panel)[0].textContent, /10\/1\/2026 9:05.*1 reply/);
});

test('a refresh builds anew only the entries whose content changed, and keeps the order of the text', () => {
  const threads = [thread(1), thread(2), thread(3)];
  const { panel, state } = makeHost(threads);
  panel.show();
  const [a, b, c] = entriesOf(panel);
  panel.refresh();
  assert.deepEqual(entriesOf(panel), [a, b, c], 'the same threads were built anew');
  state.threads = [thread(1), thread(2, 0, { text: 'Changed' }), thread(3)];
  panel.refresh();
  const after = entriesOf(panel);
  assert.equal(after[0], a);
  assert.notEqual(after[1], b);
  assert.equal(after[2], c);
  // A thread moved before another (its text was moved): the entries follow.
  state.threads = [state.threads[2], state.threads[0], state.threads[1]];
  panel.refresh();
  assert.deepEqual(entriesOf(panel).map((entry) => entry.getAttribute('data-thread')), ['3', '1', '2']);
});

test('many threads are built a batch at a time: more as the list is scrolled or stepped through, and the selected one wherever it is', () => {
  const threads: CommentThread[] = [];
  for (let i = 0; i < 250; i++) threads.push(thread(i + 1));
  const { panel } = makeHost(threads);
  panel.show();
  assert.equal(panel.builtCount, PANEL_BATCH);
  const element = panel.element as unknown as FakeElement;
  const more = byClass(element, 'bym-doc-cpanel-more')[0];
  assert.equal(more.style.display, 'block');
  assert.equal(more.textContent, 'Show 100 more');
  fire(more, 'click');
  assert.equal(panel.builtCount, 2 * PANEL_BATCH);
  assert.equal(fakeDocument.activeElement, entriesOf(panel)[PANEL_BATCH], 'the focus did not go to the first new thread');
  const scroll = byClass(element, 'bym-doc-cpanel-scroll')[0];
  scroll.clientHeight = 400;
  scroll.scrollHeight = 5000;
  scroll.scrollTop = 4500;
  fire(scroll, 'scroll');
  assert.equal(panel.builtCount, 250);
  assert.equal(more.style.display, 'none');
  // Shown anew: a batch again, and the thread the document selects is built however far it is.
  panel.hide();
  panel.show();
  assert.equal(panel.builtCount, PANEL_BATCH);
  panel.select(240, { scroll: true });
  assert.ok(entriesOf(panel).some((entry) => entry.getAttribute('data-thread') === '240'));
  key(entriesOf(panel).filter((entry) => entry.getAttribute('data-thread') === '240')[0], 'End');
  assert.equal(panel.selectedThread, 250);
});

test('2,000 threads: what showing the panel, a change of one thread and stepping to the last one cost', (t) => {
  const threads: CommentThread[] = [];
  for (let i = 0; i < 2000; i++) threads.push(thread(i + 1, i % 3, { text: `Comment ${i} `.repeat(8) }));
  const { panel, state } = makeHost(threads);
  const time = (work: () => void): number => {
    const started = process.hrtime.bigint();
    work();
    return Number(process.hrtime.bigint() - started) / 1e6;
  };
  const shown = time(() => panel.show());
  state.quotes = 0;
  const copy = threads.slice();
  copy[5] = thread(6, 0, { text: 'Edited' });
  state.threads = copy;
  const changed = time(() => panel.refresh());
  const quotesOnRefresh = state.quotes;
  const last = time(() => panel.select(2000, { scroll: true }));
  const afterAll = time(() => panel.refresh());
  t.diagnostic(`show: ${shown.toFixed(1)} ms, ${PANEL_BATCH} entries; change of one thread: ${changed.toFixed(1)} ms (${quotesOnRefresh} quotes asked for); select the last: ${last.toFixed(1)} ms; refresh with all ${panel.builtCount} built: ${afterAll.toFixed(1)} ms`);
  assert.equal(quotesOnRefresh, PANEL_BATCH, 'a refresh asks for the quotes of more than the entries that are built');
  // Guards against something that hangs, not measures of speed: ten times what an idle machine takes.
  assert.ok(shown < 2000 && changed < 2000 && last < 5000 && afterAll < 5000);
});

test('destroy takes the panel and its listeners away; it does nothing afterwards', () => {
  const { panel, root } = makeHost([thread(1)]);
  panel.show();
  const element = panel.element as unknown as FakeElement;
  panel.destroy();
  assert.equal(element.parentNode, null);
  assert.equal(root.children.length, 0);
  for (const [type, list] of element.listeners) assert.equal(list.length, 0, `a ${type} listener stayed`);
  panel.show();
  panel.select(1);
  assert.equal(panel.isOpen, false);
});

test('2,000 comments in a document: listing them and quoting their text, what the panel asks of the view on a change', async (t) => {
  let body = '';
  const comments: { id: number; text: string }[] = [];
  for (let i = 0; i < 2000; i++) {
    body += p(r(`Paragraph ${i} `) + cStart(i) + r(`commented ${i}`) + cEnd(i) + cRef(i) + r(' rest.'));
    comments.push({ id: i, text: `Comment ${i}` });
  }
  const { doc } = await openCommented({ body, comments });
  let started = process.hrtime.bigint();
  const list = listComments(doc);
  const listed = Number(process.hrtime.bigint() - started) / 1e6;
  assert.equal(list.length, 2000);
  started = process.hrtime.bigint();
  let quoted = 0;
  for (const c of list) if (commentQuote(doc, c) === `commented ${c.id}`) quoted++;
  const quoting = Number(process.hrtime.bigint() - started) / 1e6;
  assert.equal(quoted, 2000);
  t.diagnostic(`listComments: ${listed.toFixed(1)} ms; quotes of all 2,000: ${quoting.toFixed(1)} ms`);
});

// ------------------------------------------------------------ after review

test('selecting a thread scrolls the list by where the entry is on the screen, not by offsets of the page around it', () => {
  const threads: CommentThread[] = [];
  for (let i = 1; i <= 30; i++) threads.push(thread(i));
  const { panel } = makeHost(threads);
  panel.show(true);
  const element = panel.element as unknown as FakeElement;
  const scroll = byClass(element, 'bym-doc-cpanel-scroll')[0];
  scroll.clientHeight = 400;
  scroll.offsetHeight = 400;
  // The list starts 300 px down the page; every entry is 70 px high, 80 px apart. The page around has offsets that say nothing about the list.
  layout = (e) => {
    if (e === scroll) return { top: 300, height: 400 };
    if (e.className !== 'bym-doc-cpanel-item') return undefined;
    const index = e.parentNode ? e.parentNode.children.indexOf(e) : 0;
    return { top: 300 + 8 + index * 80 - scroll.scrollTop, height: 70 };
  };
  try {
    for (const e of all(fakeDocument.body, () => true)) {
      e.offsetTop = 500;
      e.offsetParent = fakeDocument.body;
    }
    panel.select(1, { scroll: true });
    assert.equal(scroll.scrollTop, 0, 'the first entry, already in sight, scrolled the list');
    // Stepping down with the arrows keeps the entry with the focus in sight.
    for (let i = 0; i < 12; i++) key(fakeDocument.activeElement as FakeElement, 'ArrowDown');
    assert.equal(panel.selectedThread, 13);
    const box = (fakeDocument.activeElement as FakeElement).getBoundingClientRect();
    assert.ok(box.top >= 300 && box.bottom <= 700, `the entry is out of sight: ${box.top}..${box.bottom}, list scrolled to ${scroll.scrollTop}`);
    key(fakeDocument.activeElement as FakeElement, 'Home');
    assert.equal(scroll.scrollTop, 0);
  } finally {
    layout = () => undefined;
  }
});

test('the drawer of a narrow view covers the bottom of the pages: what the caret is kept above', () => {
  // The pages end at 700 (400 high); a drawer from 460 hides 240 of them; at a scale of 1 and of 2.
  assert.equal(coveredBelow({ bottom: 700, height: 400 }, 460, 400), 240);
  assert.equal(coveredBelow({ bottom: 700, height: 200 }, 600, 400), 200);
  assert.equal(coveredBelow({ bottom: 700, height: 400 }, 800, 400), 0);
  assert.equal(coveredBelow({ bottom: 700, height: 0 }, 460, 400), 0);
});

test('comments with negative ids (the file says so) are answered and deleted from the panel and the balloon', () => {
  const { panel, calls } = makeHost([thread(-5, 1)]);
  panel.show();
  panel.select(-5);
  fire(buttonOf(entriesOf(panel)[0], 'delete', -5), 'click');
  fire(buttonOf(entriesOf(panel)[0], 'resolve', -5), 'click');
  assert.deepEqual(calls.remove, [-5]);
  assert.deepEqual(calls.resolve, [[-5, true]]);
  const removed: number[] = [];
  const root = new FakeElement('div');
  const balloon = new CommentBalloon({
    root: root as unknown as HTMLElement,
    strings: () => STRINGS_EN,
    canChange: () => true,
    submit: () => true,
    resolve: () => undefined,
    remove: (id) => removed.push(id),
    closed: () => undefined,
  });
  balloon.show([comment(-5)], undefined);
  const del = all(root, (e) => e.tagName === 'button' && e.getAttribute('data-action') === 'delete')[0];
  fire(del, 'click');
  assert.deepEqual(removed, [-5]);
});

test('dates before March 1900 are the dates the file says (no day of the leap year Excel made up)', () => {
  assert.equal(commentDateText('1900-01-01T00:00:00Z', LOCALE_EN), '1/1/1900 0:00');
  assert.equal(commentDateText('1900-02-28T10:00:00Z', LOCALE_EN), '2/28/1900 10:00');
  assert.equal(commentDateText('1900-03-01T00:00:00Z', LOCALE_EN), '3/1/1900 0:00');
  assert.equal(commentDateText('1900-02-28', LOCALE_TR), '28.02.1900');
});

test('a quote cut at the length it looks at never ends in half of a pair of surrogates, and says that it was cut', () => {
  const quote = quoteText('a' + ' '.repeat(478) + '\u{1F600}x');
  assert.ok(!/[\ud800-\udbff]$/.test(quote.replace(/\u2026$/, '')), JSON.stringify(quote));
  assert.ok(!/[\ud800-\udbff](?![\udc00-\udfff])/.test(quote), JSON.stringify(quote));
  assert.equal(quote.charAt(quote.length - 1), '\u2026');
  assert.equal(quoteText('b'.repeat(200) + ' '.repeat(10000) + 'c'), 'b'.repeat(119) + '\u2026');
});

test('selecting a thread the document does not have changes nothing, and says nothing', () => {
  const { panel, calls } = makeHost([thread(1), thread(2)]);
  panel.show();
  panel.select(7);
  assert.deepEqual(calls.selected, []);
  assert.equal(panel.selectedThread, undefined);
  panel.select(1);
  panel.select(7);
  assert.deepEqual(calls.selected, [1, undefined]);
  assert.equal(panel.selectedThread, undefined);
});

test('2,000 comments in one paragraph of 1 MB: quotes read the paragraph once per change, not once per comment', async (t) => {
  let inner = '';
  const comments: { id: number; text: string }[] = [];
  for (let i = 0; i < 2000; i++) {
    inner += cStart(i) + r('w'.repeat(40)) + cEnd(i) + cRef(i) + r('z'.repeat(460));
    comments.push({ id: i, text: 'c' });
  }
  const { doc } = await openCommented({ body: p(inner), comments });
  const list = listComments(doc);
  let started = process.hrtime.bigint();
  const plain = list.slice(0, 200).map((c) => commentQuote(doc, c));
  const uncached = Number(process.hrtime.bigint() - started) / 1e6;
  started = process.hrtime.bigint();
  const texts = paragraphTexts();
  const cached = list.map((c) => commentQuote(doc, c, texts));
  const once = Number(process.hrtime.bigint() - started) / 1e6;
  assert.deepEqual(cached.slice(0, 200), plain);
  assert.ok(cached.every((quote) => quote === 'w'.repeat(40)));
  t.diagnostic(`without the cache: 200 quotes ${uncached.toFixed(1)} ms; with it: all 2,000 quotes ${once.toFixed(1)} ms`);
  assert.ok(once < 2000);
});
