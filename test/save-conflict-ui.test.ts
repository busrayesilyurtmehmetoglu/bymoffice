import assert from 'node:assert/strict';
import { test } from 'node:test';

// The question about a conflict (src/ui/save-confirm.ts) over a DOM that is made here: only what
// the question uses of one. What is sure of it without a browser: where it is put, what takes
// input while it is open, which presses are an answer, and that nothing of it is left afterwards.

interface FakeEvent {
  key?: string;
  repeat?: boolean;
  detail?: number;
  target?: FakeElement;
  stopped?: boolean;
  prevented?: boolean;
  stopPropagation?(): void;
  preventDefault?(): void;
}

class FakeStyle {
  cssText = '';
  opacity = '';
  cursor = '';
  private readonly set = new Map<string, { value: string; priority: string }>();
  setProperty(name: string, value: string, priority = ''): void {
    this.set.set(name, { value, priority });
  }
  getPropertyValue(name: string): string {
    const entry = this.set.get(name);
    return entry ? entry.value : '';
  }
  getPropertyPriority(name: string): string {
    const entry = this.set.get(name);
    return entry ? entry.priority : '';
  }
  removeProperty(name: string): void {
    this.set.delete(name);
  }
  has(name: string): boolean {
    return this.set.has(name);
  }
}

class FakeElement {
  children: FakeElement[] = [];
  parentNode: FakeElement | null = null;
  readonly attributes = new Map<string, string>();
  readonly style = new FakeStyle();
  readonly listeners = new Map<string, ((event: FakeEvent) => void)[]>();
  className = '';
  id = '';
  textContent = '';
  type = '';
  /** What a rule of the page says of its position, over what the element says inline (unless that is `important`). */
  pagePosition: string | undefined;
  constructor(
    readonly tag: string,
    readonly label = tag,
  ) {}
  get parentElement(): FakeElement | null {
    return this.parentNode;
  }
  get ownerDocument(): typeof fakeDocument {
    return fakeDocument;
  }
  get lastElementChild(): FakeElement | null {
    return this.children[this.children.length - 1] || null;
  }
  get isConnected(): boolean {
    for (let at: FakeElement | null = this; at; at = at.parentNode) if (at === fakeDocument.body) return true;
    return false;
  }
  appendChild(child: FakeElement): FakeElement {
    if (child.parentNode) child.parentNode.removeChild(child);
    this.children.push(child);
    child.parentNode = this;
    return child;
  }
  removeChild(child: FakeElement): FakeElement {
    const at = this.children.indexOf(child);
    if (at >= 0) this.children.splice(at, 1);
    child.parentNode = null;
    // As a browser does: the focus of an element that is taken out goes to the body.
    if (!fakeDocument.activeElement.isConnected) fakeDocument.activeElement = fakeDocument.body;
    return child;
  }
  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
  getAttribute(name: string): string | null {
    const value = this.attributes.get(name);
    return value === undefined ? null : value;
  }
  hasAttribute(name: string): boolean {
    return this.attributes.has(name);
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
  contains(other: FakeElement | null): boolean {
    for (let at = other; at; at = at.parentNode) if (at === this) return true;
    return false;
  }
  getRootNode(): typeof fakeDocument {
    return fakeDocument;
  }
  /** Inside something inert nothing takes the focus. */
  focus(): void {
    for (let at: FakeElement | null = this; at; at = at.parentNode) if (at.hasAttribute('inert')) return;
    if (this.isConnected) fakeDocument.activeElement = this;
  }
  /** Sends an event up from here, as a browser does; returns it with what the listeners did to it. */
  fire(type: string, event: FakeEvent = {}): FakeEvent {
    event.target = event.target || this;
    event.stopPropagation = () => void (event.stopped = true);
    event.preventDefault = () => void (event.prevented = true);
    for (let at: FakeElement | null = this; at && !event.stopped; at = at.parentNode) {
      for (const listener of (at.listeners.get(type) || []).slice()) listener(event);
    }
    return event;
  }
  listenerCount(): number {
    let count = 0;
    this.listeners.forEach((list) => (count += list.length));
    return count;
  }
  find(matches: (element: FakeElement) => boolean): FakeElement | undefined {
    if (matches(this)) return this;
    for (const child of this.children) {
      const found = child.find(matches);
      if (found) return found;
    }
    return undefined;
  }
}

const fakeDocument = {
  body: undefined as unknown as FakeElement,
  activeElement: undefined as unknown as FakeElement,
  createElement: (tag: string) => new FakeElement(tag),
};
fakeDocument.body = new FakeElement('body');
fakeDocument.activeElement = fakeDocument.body;

const computedPosition = (element: FakeElement): string => {
  const inline = element.style.getPropertyValue('position');
  if (inline !== '' && element.style.getPropertyPriority('position') === 'important') return inline;
  return element.pagePosition || inline || 'static';
};

class FakeHTMLElement {}
const globals = globalThis as unknown as { document: unknown; window: unknown; HTMLElement: unknown };
globals.document = fakeDocument;
globals.HTMLElement = FakeHTMLElement;
globals.window = { getComputedStyle: (element: FakeElement) => ({ position: computedPosition(element) }), matchMedia: () => ({ matches: false }) };
/** Whether the browser knows the `inert` attribute. */
const knowInert = (known: boolean): void => {
  if (known) (FakeHTMLElement.prototype as { inert?: boolean }).inert = false;
  else delete (FakeHTMLElement.prototype as { inert?: boolean }).inert;
};
knowInert(true);

const { askOverwrite, editorRootIn, OVERWRITE_ARM_MS } = await import('../src/ui/save-confirm');
const STRINGS = { saveConflictTitle: 'Title', saveConflictQuestion: 'Question', saveOverwrite: 'Overwrite', saveCancel: 'Cancel' };

interface Scene {
  container: FakeElement;
  root: FakeElement;
  toolbar: FakeElement;
  input: FakeElement;
  answers: boolean[];
  /** The timers that were asked for and not called off. */
  timers: { callback: () => void; ms: number }[];
  /** Lets the time pass after which "overwrite" answers. */
  arm(): void;
  close: () => void;
  backdrop(): FakeElement | undefined;
  overwrite(): FakeElement;
  cancel(): FakeElement;
}

/** An editor in a container, as the editors build theirs: a positioned root with a bar and the element that takes what is typed. */
function editor(): Omit<Scene, 'answers' | 'timers' | 'arm' | 'close' | 'backdrop' | 'overwrite' | 'cancel'> {
  fakeDocument.body.children = [];
  fakeDocument.activeElement = fakeDocument.body;
  const container = fakeDocument.body.appendChild(new FakeElement('div', 'container'));
  const root = container.appendChild(new FakeElement('div', 'root'));
  root.style.setProperty('position', 'relative');
  const toolbar = root.appendChild(new FakeElement('div', 'toolbar'));
  const input = root.appendChild(new FakeElement('textarea', 'input'));
  input.focus();
  return { container, root, toolbar, input };
}

function ask(made = editor(), more: { detail?: string; answer?: (overwrite: boolean) => void; focusView?: () => void } = {}): Scene {
  const answers: boolean[] = [];
  const timers: Scene['timers'] = [];
  const close = askOverwrite({
    root: made.root as unknown as HTMLElement,
    strings: STRINGS,
    detail: more.detail,
    answer: more.answer || ((overwrite) => void answers.push(overwrite)),
    focusView: more.focusView,
    timer: (callback, ms) => {
      const entry = { callback, ms };
      timers.push(entry);
      return () => void timers.splice(timers.indexOf(entry), 1);
    },
  });
  const action = (name: string) => () => {
    const found = fakeDocument.body.find((element) => element.getAttribute('data-bym-action') === name);
    assert.notEqual(found, undefined, `the answer "${name}" is in the document`);
    return found as FakeElement;
  };
  return {
    ...made,
    answers,
    timers,
    arm: () => {
      assert.equal(timers.length, 1, 'one timer waits');
      timers.splice(0, 1)[0].callback();
    },
    close,
    backdrop: () => fakeDocument.body.find((element) => element.className === 'bym-confirm-backdrop'),
    overwrite: action('overwrite'),
    cancel: action('cancel'),
  };
}

/** A key that goes down and up on the element with the focus; Enter and Space press a button as a browser does, unless that was prevented. */
function key(name: string, more: { repeat?: boolean; up?: boolean } = {}): { down: FakeEvent; up?: FakeEvent } {
  const target = fakeDocument.activeElement;
  const button = target.tag === 'button';
  const down = target.fire('keydown', { key: name, repeat: more.repeat === true });
  if (button && name === 'Enter' && !down.prevented) target.fire('click', { detail: 0 });
  if (more.up === false) return { down };
  const up = target.fire('keyup', { key: name });
  if (button && name === ' ' && !down.prevented && !up.prevented) target.fire('click', { detail: 0 });
  return { down, up };
}

const inertOf = (...elements: FakeElement[]): string => elements.map((element) => `${element.label}:${element.hasAttribute('inert')}`).join(' ');

test('the question is put into the root it is given and silences what is in it, whatever else the page put into the container', () => {
  const made = editor();
  // What an application adds beside the editor after it was mounted: a notice of its own, positioned.
  const toast = made.container.appendChild(new FakeElement('div', 'toast'));
  toast.pagePosition = 'absolute';
  const note = made.container.appendChild(new FakeElement('div', 'note'));
  const s = ask(made);
  const backdrop = s.backdrop() as FakeElement;
  assert.equal(backdrop.parentNode, s.root, 'the question is in the root of the editor');
  assert.equal(s.root.lastElementChild, backdrop, 'after everything else in it');
  assert.equal(inertOf(s.toolbar, s.input), 'toolbar:true input:true', 'the editor takes no input');
  assert.equal(inertOf(s.root, toast, note, s.container), 'root:false toast:false note:false container:false', 'and nothing else is silenced');
  assert.equal(fakeDocument.activeElement, s.cancel(), 'the focus is on the answer that changes nothing');
  assert.equal(s.root.style.getPropertyValue('position'), 'relative', 'a root that is positioned is left as it is');
  assert.equal(s.root.style.getPropertyPriority('position'), '', 'with the priority it had');
  s.close();
});

test('editorRootIn: the element of a view leads to what was put into the container around it', () => {
  const made = editor();
  const grid = made.root.appendChild(new FakeElement('div', 'grid'));
  const view = grid.appendChild(new FakeElement('div', 'view'));
  const as = (element: FakeElement | undefined) => element as unknown as HTMLElement;
  assert.equal(editorRootIn(as(made.container), as(view)), as(made.root), 'the root of the shell around a view');
  assert.equal(editorRootIn(as(made.container), as(made.root)), as(made.root), 'a root that is in the container itself');
  assert.equal(editorRootIn(as(made.container), undefined), as(made.container), 'a view that names no element leaves the container');
  assert.equal(editorRootIn(as(made.container), as(new FakeElement('div', 'elsewhere'))), as(made.container), 'and so does an element that is not in the container');
  made.container.appendChild(new FakeElement('div', 'toast')).pagePosition = 'absolute';
  assert.equal(editorRootIn(as(made.container), as(view)), as(made.root), 'whatever was put into the container after the editor');
});

test('a root that is not positioned is positioned while the question is open, and is as it was afterwards', () => {
  // A rule of the page undoes what the editor set inline.
  const overruled = editor();
  overruled.root.pagePosition = 'static';
  const first = ask(overruled);
  assert.equal(computedPosition(first.root), 'relative', 'the question has edges to be placed by');
  assert.equal(first.root.style.getPropertyPriority('position'), 'important', 'over the rule of the page');
  first.cancel().fire('click', { detail: 1 });
  assert.equal(first.root.style.getPropertyValue('position'), 'relative', 'what the editor had set is back');
  assert.equal(first.root.style.getPropertyPriority('position'), '', 'with its priority');

  // A root that says nothing of its position (a container that is covered itself).
  const bare = editor();
  bare.root.style.removeProperty('position');
  const second = ask(bare);
  assert.equal(computedPosition(second.root), 'relative', 'positioned while the question is open');
  second.close();
  assert.equal(second.root.style.has('position'), false, 'and nothing of it is left');
});

test('the arrow keys do not move between the answers: an arrow and Enter, on their way to the editor, do not overwrite', () => {
  for (const arrow of ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight']) {
    const s = ask();
    s.arm();
    const pressed = key(arrow);
    assert.equal(fakeDocument.activeElement, s.cancel(), `${arrow} leaves the focus on "cancel"`);
    assert.equal(pressed.down.prevented, true, `${arrow} does nothing to the page either`);
    assert.equal(pressed.down.stopped, true, `and the editor does not hear of ${arrow}`);
    for (let i = 0; i < 5; i++) key(arrow, { repeat: true, up: false });
    assert.equal(fakeDocument.activeElement, s.cancel(), `${arrow} held down moves nothing`);
    key('Enter');
    assert.deepEqual(s.answers, [false], `Enter after ${arrow} is "cancel"`);
    assert.equal(s.backdrop(), undefined, 'and the question is gone');
  }
});

test(`"overwrite" answers nothing for ${OVERWRITE_ARM_MS} ms after the question came; "cancel" and Escape answer at once`, () => {
  const s = ask();
  assert.deepEqual(
    s.timers.map((timer) => timer.ms),
    [400],
    'the time after which "overwrite" answers',
  );
  assert.equal(s.overwrite().getAttribute('aria-disabled'), 'true', 'the answer says that it is not available yet');
  assert.equal(s.overwrite().hasAttribute('disabled'), false, 'and is not disabled: it can be reached');
  // The second press of a double press on "save", wherever the question put its answer.
  s.overwrite().fire('click', { detail: 1 });
  s.overwrite().fire('click', { detail: 2 });
  s.overwrite().focus();
  assert.equal(fakeDocument.activeElement, s.overwrite(), 'the answer takes the focus');
  key('Enter');
  key(' ');
  assert.deepEqual(s.answers, [], 'no press is an answer yet');
  assert.notEqual(s.backdrop(), undefined, 'and the question stays');
  s.arm();
  assert.equal(s.overwrite().hasAttribute('aria-disabled'), false, 'the answer is available');
  assert.equal(s.overwrite().style.opacity, '1', 'and looks it');
  const overwrite = s.overwrite();
  overwrite.fire('click', { detail: 1 });
  assert.deepEqual(s.answers, [true], 'a press answers now');
  assert.equal(s.backdrop(), undefined, 'the question is gone');
  overwrite.fire('click', { detail: 1 });
  assert.deepEqual(s.answers, [true], 'and answers once');

  const early = ask();
  early.cancel().fire('click', { detail: 1 });
  assert.deepEqual(early.answers, [false], '"cancel" answers before that time');
  assert.equal(early.timers.length, 0, 'and the timer is called off');

  const escaped = ask();
  const pressed = key('Escape');
  assert.deepEqual(escaped.answers, [false], 'Escape cancels before that time');
  assert.equal(pressed.down.prevented, true, 'and is not handed on');
  assert.equal(escaped.timers.length, 0, 'the timer is called off');

  const keyed = ask();
  key('Enter');
  assert.deepEqual(keyed.answers, [false], 'Enter on "cancel", which has the focus, cancels at once');
});

test('a key that is held down never overwrites, also when it is let go; the next press does', () => {
  const s = ask();
  s.arm();
  s.overwrite().focus();
  // Enter that repeats: a browser presses the button with each repeat unless the key is prevented.
  for (let i = 0; i < 4; i++) {
    const pressed = key('Enter', { repeat: true, up: false });
    assert.equal(pressed.down.prevented, true, 'a repeated Enter is prevented');
  }
  // A browser that presses all the same.
  s.overwrite().fire('click', { detail: 0 });
  assert.deepEqual(s.answers, [], 'a repeated Enter is no answer');
  // Space that was held: the press comes when it is let go.
  key(' ', { repeat: true, up: false });
  const up = s.overwrite().fire('keyup', { key: ' ' });
  assert.equal(up.prevented, true, 'letting go of a Space that repeated presses nothing');
  s.overwrite().fire('click', { detail: 0 });
  assert.deepEqual(s.answers, [], 'and is no answer where it does');
  key('Escape', { repeat: true, up: false });
  assert.deepEqual(s.answers, [], 'an Escape that is held down cancels nothing');
  assert.notEqual(s.backdrop(), undefined, 'the question is still open');
  key('Enter');
  assert.deepEqual(s.answers, [true], 'a key that is pressed anew answers');

  // Held down since before the time was over.
  const early = ask();
  early.overwrite().focus();
  key('Enter', { up: false });
  early.arm();
  for (let i = 0; i < 3; i++) key('Enter', { repeat: true, up: false });
  early.overwrite().fire('click', { detail: 0 });
  assert.deepEqual(early.answers, [], 'a key that went down too early does not answer once the time is over');
  // The pointer is not held back by a key that repeats.
  early.overwrite().fire('click', { detail: 1 });
  assert.deepEqual(early.answers, [true], 'a press with the pointer answers');
});

test('however the question ends, the editor takes input again, the focus is back and no listener is left', () => {
  const ways: [name: string, end: (s: Scene) => void, answers: boolean[]][] = [
    ['cancel', (s) => void s.cancel().fire('click', { detail: 1 }), [false]],
    ['overwrite', (s) => (s.arm(), void s.overwrite().fire('click', { detail: 1 })), [true]],
    ['Escape', () => void key('Escape'), [false]],
    ['taken away by who asked', (s) => (s.close(), s.close()), []],
    ['taken away before "overwrite" answers', (s) => s.close(), []],
  ];
  for (const inert of [true, false]) {
    knowInert(inert);
    for (const [name, end, answers] of ways) {
      const made = editor();
      // Something of the editor that was inert before the question stays so after it.
      const menu = made.root.appendChild(new FakeElement('div', 'menu'));
      menu.setAttribute('inert', '');
      const s = ask(made);
      const backdrop = s.backdrop() as FakeElement;
      const buttons = [s.overwrite(), s.cancel()];
      assert.equal(inertOf(s.toolbar, s.input, menu), 'toolbar:true input:true menu:true', `${name}: the editor is silent while the question is open`);
      assert.equal(s.root.listenerCount(), inert ? 0 : 1, `${name}: a browser without inert has its focus watched`);
      end(s);
      assert.deepEqual(s.answers, answers, `${name}: the answer`);
      assert.equal(s.backdrop(), undefined, `${name}: the question is out of the document`);
      assert.equal(inertOf(s.toolbar, s.input, menu), 'toolbar:false input:false menu:true', `${name}: the editor takes input again, and what was inert before still is`);
      assert.equal(backdrop.listenerCount(), 0, `${name}: no listener is left on the question`);
      assert.equal(s.root.listenerCount(), 0, `${name}: none on the root`);
      assert.equal(s.timers.length, 0, `${name}: no timer waits`);
      assert.equal(fakeDocument.activeElement, s.input, `${name}: the focus is back where it was`);
      // What is left of the question answers nothing.
      for (const button of buttons) button.fire('click', { detail: 1 });
      backdrop.fire('keydown', { key: 'Escape' });
      assert.deepEqual(s.answers, answers, `${name}: nothing answers afterwards`);
    }
  }
  knowInert(true);
});

test('a press never takes the focus out of the question: on an answer it goes to that answer, beside them it stays', () => {
  const s = ask();
  // Safari gives a button that is pressed no focus and takes it from the element that had it, unless the press is prevented.
  const onAnswer = s.overwrite().fire('mousedown');
  assert.equal(onAnswer.prevented, true, 'the press does not move the focus by itself');
  assert.equal(fakeDocument.activeElement, s.overwrite(), 'the answer that is pressed has the focus, also while it does not answer yet');
  const beside = (s.backdrop() as FakeElement).fire('mousedown');
  assert.equal(beside.prevented, true, 'a press beside the answers is prevented');
  assert.equal(fakeDocument.activeElement, s.overwrite(), 'and leaves the focus where it is');
  s.cancel().fire('mousedown');
  assert.equal(fakeDocument.activeElement, s.cancel(), 'a press on "cancel" gives it the focus');
  fakeDocument.activeElement = fakeDocument.body;
  (s.backdrop() as FakeElement).fire('mousedown');
  assert.equal(fakeDocument.activeElement, s.cancel(), 'a focus that was lost comes back to "cancel"');
  key('Escape');
  assert.deepEqual(s.answers, [false], 'and Escape cancels');
});

test('a browser without inert: the focus that goes into the editor comes back to the question', () => {
  knowInert(false);
  try {
    const s = ask();
    s.input.removeAttribute('inert');
    s.input.focus();
    s.input.fire('focusin');
    assert.equal(fakeDocument.activeElement, s.cancel(), 'the focus is on "cancel" again');
    s.close();
  } finally {
    knowInert(true);
  }
});

test('an answer that throws still takes the question away and gives the editor back', () => {
  for (const overwrite of [false, true]) {
    const made = editor();
    const s = ask(made, {
      answer: () => {
        throw new Error('the page failed at its answer');
      },
    });
    const backdrop = s.backdrop() as FakeElement;
    s.arm();
    assert.throws(() => (overwrite ? s.overwrite() : s.cancel()).fire('click', { detail: 1 }), /the page failed at its answer/, 'what the page threw is not swallowed');
    assert.equal(s.backdrop(), undefined, 'the question is gone');
    assert.equal(inertOf(s.toolbar, s.input), 'toolbar:false input:false', 'the editor takes input again');
    assert.equal(backdrop.listenerCount() + s.root.listenerCount(), 0, 'no listener is left');
    assert.equal(fakeDocument.activeElement, s.input, 'the focus is back');
  }
});

test('when the element that had the focus is gone, the view is asked to take it', () => {
  const made = editor();
  let focused = 0;
  const s = ask(made, { focusView: () => void focused++ });
  made.root.removeChild(made.input);
  s.cancel().focus();
  key('Escape');
  assert.equal(focused, 1, 'the view was given the focus');
});

test('what the application said is shown as one short line of text', () => {
  const s = ask(editor(), { detail: `<b>Ada</b>${String.fromCharCode(10)}changed it ${'a'.repeat(500)}` });
  const detail = fakeDocument.body.find((element) => element.className === 'bym-confirm-detail') as FakeElement;
  assert.equal(detail.textContent.slice(0, 22), '<b>Ada</b> changed it ', 'as text, on one line');
  assert.equal(detail.textContent.length, 201, 'cut after 200 characters');
  assert.equal(detail.children.length, 0, 'and never as markup');
  s.close();
  const none = ask(editor(), { detail: '   ' });
  assert.equal(fakeDocument.body.find((element) => element.className === 'bym-confirm-detail'), undefined, 'nothing is shown when nothing was said');
  none.close();
});
