import assert from 'node:assert/strict';
import { test } from 'node:test';
import { FocusRetry, pressOf } from '../src/ui/focus-retry';

// The focus that is asked for again at the end of a press (a frame that is
// not the active one, in Firefox): only for that press, only at its release,
// only while the focus is nowhere, and never more than one at a time.

interface Listener {
  type: string;
  fn: (event: { type: string }) => void;
}

function world(shadow: 'none' | 'one' | 'nested' = 'none') {
  const listeners: Listener[] = [];
  const doc: any = {
    body: { id: 'body' },
    documentElement: { id: 'html' },
    activeElement: null as unknown,
    addEventListener: (type: string, fn: Listener['fn'], capture: unknown) => {
      assert.equal(capture, true, 'while the event goes down');
      listeners.push({ type, fn });
    },
    removeEventListener: (type: string, fn: Listener['fn']) => {
      const at = listeners.findIndex((listener) => listener.type === type && listener.fn === fn);
      if (at >= 0) listeners.splice(at, 1);
    },
  };
  doc.activeElement = doc.body;
  // The shadow roots the input is in, from the inside out; each has a host in the root around it.
  const roots: any[] = [];
  let around: any = doc;
  for (let i = 0; i < (shadow === 'none' ? 0 : shadow === 'one' ? 1 : 2); i++) {
    const outer = around;
    const host = { id: `host${i}`, getRootNode: () => outer };
    const root = { activeElement: null as unknown, host };
    roots.unshift(root);
    around = root;
  }
  const root = roots.length > 0 ? roots[0] : doc;
  const input: any = { id: 'input', ownerDocument: doc, getRootNode: () => root };
  const fire = (type: string, more: object = {}): void => {
    for (const listener of listeners.filter((entry) => entry.type === type)) listener.fn({ type, ...more });
  };
  let focused = 0;
  let alive = true;
  const retry = new FocusRetry(input, () => alive);
  return {
    doc,
    roots,
    input,
    listeners,
    fire,
    retry,
    focus: () => void focused++,
    get focused() {
      return focused;
    },
    kill: () => void (alive = false),
  };
}

const MOUSE = { pointerId: undefined, touch: false };

test('pressOf: what a mouse, a pointer and a touch event say of their press', () => {
  assert.deepEqual(pressOf({ type: 'mousedown' } as Event), { pointerId: undefined, touch: false });
  assert.deepEqual(pressOf({ type: 'pointerdown', pointerId: 1, pointerType: 'mouse' } as unknown as Event), { pointerId: 1, touch: false });
  assert.deepEqual(pressOf({ type: 'pointerup', pointerId: 7, pointerType: 'touch' } as unknown as Event), { pointerId: 7, touch: true });
  assert.deepEqual(pressOf({ type: 'touchstart' } as Event), { pointerId: undefined, touch: true });
});

test('a press whose focus did not take: asked again at its release, once, and the listeners are gone', () => {
  for (const release of ['pointerup', 'mouseup']) {
    const w = world();
    w.retry.after(MOUSE, w.focus);
    assert.equal(w.retry.waiting, true);
    w.fire(release, release === 'pointerup' ? { pointerId: 1, pointerType: 'mouse' } : {});
    assert.equal(w.focused, 1, release);
    assert.deepEqual([w.listeners.length, w.retry.waiting], [0, false], release);
    w.fire('mouseup');
    w.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
    assert.equal(w.focused, 1, 'a later release finds nothing waiting');
  }
});

test('a press whose focus took: nothing waits', () => {
  const w = world();
  w.doc.activeElement = w.input;
  w.retry.after(MOUSE, w.focus);
  assert.deepEqual([w.listeners.length, w.retry.waiting], [0, false]);
});

test('one follow-up at a time: another press takes the place of the one that waits', () => {
  const w = world();
  for (let i = 0; i < 5; i++) w.retry.after(MOUSE, w.focus);
  const types = w.listeners.map((listener) => listener.type).sort();
  assert.deepEqual(types, ['dragend', 'mouseup', 'pointercancel', 'pointerup', 'touchcancel', 'touchend', 'visibilitychange'], 'one listener of each kind, not five');
  w.fire('mouseup');
  assert.equal(w.focused, 1, 'and one focus');
});

test('a press that ends without a release, a page that is hidden, a view that is destroyed: nothing waits any more', () => {
  for (const end of ['pointercancel', 'touchcancel', 'dragend', 'visibilitychange', 'cancel()']) {
    const w = world();
    w.retry.after(MOUSE, w.focus);
    if (end === 'cancel()') w.retry.cancel();
    else w.fire(end);
    assert.deepEqual([w.listeners.length, w.retry.waiting], [0, false], end);
    // Much later the user presses and releases somewhere else in the document, where nothing takes the focus.
    w.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
    w.fire('mouseup');
    w.fire('touchend');
    assert.equal(w.focused, 0, `${end}: the unrelated release does not give the editor the focus`);
  }
  // A view that is gone by the time of the release: nothing is focused, and nothing is left behind.
  const w = world();
  w.retry.after(MOUSE, w.focus);
  w.kill();
  w.fire('mouseup');
  assert.deepEqual([w.focused, w.listeners.length], [0, 0]);
});

test('only the release of the pointer that pressed counts', () => {
  // A pointer press: the pointerup of another pointer is not its release, nor is a mouseup of a finger's press.
  const w = world();
  w.retry.after({ pointerId: 4, touch: true }, w.focus);
  w.fire('pointerup', { pointerId: 9, pointerType: 'touch' });
  w.fire('mouseup');
  assert.deepEqual([w.focused, w.retry.waiting], [0, true]);
  w.fire('pointerup', { pointerId: 4, pointerType: 'touch' });
  assert.deepEqual([w.focused, w.retry.waiting], [1, false]);

  // A finger's press seen as a touch event ends with touchend; a mouse press does not.
  const finger = world();
  finger.retry.after({ pointerId: undefined, touch: true }, finger.focus);
  finger.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
  assert.equal(finger.focused, 0);
  finger.fire('touchend');
  assert.equal(finger.focused, 1);

  const mouse = world();
  mouse.retry.after(MOUSE, mouse.focus);
  mouse.fire('touchend');
  mouse.fire('pointerup', { pointerId: 3, pointerType: 'touch' });
  assert.deepEqual([mouse.focused, mouse.retry.waiting], [0, true]);
  mouse.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
  assert.equal(mouse.focused, 1);
});

test('the focus went somewhere meanwhile: it stays there', () => {
  const w = world();
  w.retry.after(MOUSE, w.focus);
  w.doc.activeElement = { id: 'field-of-the-application' };
  w.fire('mouseup');
  assert.deepEqual([w.focused, w.listeners.length], [0, 0]);

  // Nowhere: nothing, the body, the root element.
  for (const nowhere of ['null', 'body', 'html']) {
    const again = world();
    again.retry.after(MOUSE, again.focus);
    again.doc.activeElement = nowhere === 'null' ? null : nowhere === 'body' ? again.doc.body : again.doc.documentElement;
    again.fire('mouseup');
    assert.equal(again.focused, 1, nowhere);
  }
  // The input got it by itself before the release: not focused again.
  const taken = world();
  taken.retry.after(MOUSE, taken.focus);
  taken.doc.activeElement = taken.input;
  taken.fire('mouseup');
  assert.equal(taken.focused, 0);
});

test('in a shadow root: a field outside the root keeps the focus it took; the host, the body or nothing count as nowhere', () => {
  for (const depth of ['one', 'nested'] as const) {
    // The application focused a field of the document: the shadow root says null, the document names the field.
    const outside = world(depth);
    outside.retry.after(MOUSE, outside.focus);
    outside.doc.activeElement = { id: 'field-of-the-application' };
    outside.fire('mouseup');
    assert.equal(outside.focused, 0, `${depth}: the focus is not taken back from a field outside the shadow root`);

    // Something inside the same shadow root (a menu of the view) has it.
    const inside = world(depth);
    inside.retry.after(MOUSE, inside.focus);
    inside.roots[0].activeElement = { id: 'menu' };
    inside.fire('mouseup');
    assert.equal(inside.focused, 0, `${depth}: nor from something in the root`);

    for (const nowhere of ['body', 'null', 'host']) {
      const w = world(depth);
      w.retry.after(MOUSE, w.focus);
      w.doc.activeElement = nowhere === 'body' ? w.doc.body : nowhere === 'null' ? null : w.roots[w.roots.length - 1].host;
      w.fire('mouseup');
      assert.equal(w.focused, 1, `${depth}, ${nowhere}: nowhere, so the editor takes it`);
    }
  }
  // Nested: a field in the root around the view's own root.
  const between = world('nested');
  between.retry.after(MOUSE, between.focus);
  between.roots[1].activeElement = { id: 'field-in-the-outer-root' };
  between.doc.activeElement = between.roots[1].host;
  between.fire('mouseup');
  assert.equal(between.focused, 0, 'a field in the root between the view and the document keeps the focus');
});

test('in a frame that is not the active one at the release: the focus is with the page around it, and stays there', () => {
  // Firefox, the page around the frame focused a field of its own while the press was held: the frame never became active.
  const away = world();
  away.doc.hasFocus = () => false;
  away.retry.after(MOUSE, away.focus);
  away.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
  assert.deepEqual([away.focused, away.listeners.length], [0, 0]);
  // The frame became active with the release, as it does when nothing else took the focus: the editor takes it.
  const active = world();
  let has = false;
  active.doc.hasFocus = () => has;
  active.retry.after(MOUSE, active.focus);
  has = true;
  active.fire('pointerup', { pointerId: 1, pointerType: 'mouse' });
  assert.equal(active.focused, 1);
});

test('an input without a document that can be listened to: nothing is done', () => {
  const retry = new FocusRetry({ ownerDocument: null, getRootNode: () => ({ activeElement: null }) } as unknown as HTMLElement, () => true);
  let focused = 0;
  retry.after(MOUSE, () => void focused++);
  assert.deepEqual([retry.waiting, focused], [false, 0]);
  retry.cancel();
});
