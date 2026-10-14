// Tab stops of a document that is edited together: a change of them travels as a change of
// paragraph formatting ("para"), is checked like everything that arrives, and every participant
// works it out against the same styles. The default tab stop is a setting no shared command
// changes: refused in a room, and said.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CollabDocHistory, NOT_SHARED } from '../src/collab/doc/history';
import { tracing } from '../src/collab/doc/trace';
import * as shared from '../src/collab/doc/commands';
import { decodeStep } from '../src/collab/doc/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import { paragraphTabStops } from '../src/ops/doc/index';
import { DocNet, actDoc, assertSameDoc, at, docFile, joinDoc } from './collab-doc-fixtures';

async function pair(paragraphs: string[]) {
  const bytes = await docFile(paragraphs);
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  return { net, a, b };
}

const stops = (doc: Parameters<typeof paragraphTabStops>[0], row: number) => paragraphTabStops(doc, { story: 'body', path: [row], offset: 0 });

test('collab tabs: tab stops set by one are the stops of the other, and a clear by the other reaches the first', async () => {
  const { net, a, b } = await pair(['one', 'two']);
  assert.ok(actDoc(a, [{ t: 'para', at: at(0, 0), to: at(1, 0), f: { tabs: { set: [{ pos: 1440, align: 'right', leader: 'dot' }] } } }]));
  net.drain();
  for (const member of [a, b]) for (const row of [0, 1]) assert.deepEqual(stops(member.doc, row), [{ pos: 1440, align: 'right', leader: 'dot' }], `${member.name} ${row}`);
  assertSameDoc([a, b], 'after the stops');
  assert.ok(actDoc(b, [{ t: 'para', at: at(0, 0), to: at(0, 0), f: { tabs: { clear: [1440] } } }]));
  net.drain();
  assert.deepEqual(stops(a.doc, 0), []);
  assert.deepEqual(stops(a.doc, 1).length, 1);
  assertSameDoc([a, b], 'after the clear');
});

test('collab tabs: an undo of tab stops takes them back for everybody: the paragraph travels with the stops it states', async () => {
  const { net, a, b } = await pair(['one', 'two']);
  assert.ok(actDoc(a, [{ t: 'para', at: at(0, 0), to: at(0, 0), f: { tabs: { set: [{ pos: 2000, align: 'decimal' }] } } }]));
  net.drain();
  // Stops a paragraph already had, which the undo has to bring back as they were.
  assert.ok(actDoc(a, [{ t: 'para', at: at(0, 0), to: at(0, 0), f: { tabs: { set: [{ pos: 3000, align: 'left' }] } } }]));
  net.drain();
  assert.ok(a.session.undo(), 'the paragraph travels whole, with its own stops');
  net.drain();
  assertSameDoc([a, b], 'after the undo');
  assert.deepEqual(stops(b.doc, 0), [{ pos: 2000, align: 'decimal' }]);
  assert.deepEqual(a.notices, []);
});

test('collab tabs: a change of tab stops that is not one is refused when it arrives', () => {
  for (const tabs of [{ set: [{ pos: 1.5, align: 'left' }] }, { set: [{ pos: 1, align: 'clear' }] }, { clear: ['x'] }, { other: 1 }, { set: new Array(65).fill({ pos: 1, align: 'left' }) }]) {
    assert.throws(() => decodeStep([{ t: 'para', at: at(0, 0), to: at(0, 0), f: { tabs } }]), JSON.stringify(tabs).slice(0, 50));
  }
});

test('collab tabs: the command travels as itself, the default tab stop does not travel and is refused', async () => {
  const { a } = await pair(['one']);
  tracing(true);
  try {
    const at0 = { story: 'body', path: [0], offset: 0 };
    const edit = shared.setParaFormat(a.doc, { anchor: at0, focus: at0 }, { tabs: { set: [{ pos: 720, align: 'center' }] } });
    const history = new CollabDocHistory<unknown>(a.session);
    history.apply(edit.ops, undefined, undefined);
    const setting = shared.setDefaultTabStop(a.doc, 360, at0);
    assert.throws(() => history.apply(setting.ops, undefined, undefined), (error: unknown) => error instanceof CollabUnsupportedError && error.message === NOT_SHARED);
    assert.equal(a.doc.settings.defaultTabStop, 720);
  } finally {
    tracing(false);
  }
});

test('collab tabs: a paragraph that is put states its stops as a list; a change of stops in its place is refused', () => {
  const place = ['body', [0], 0];
  const ok = { p: { tabs: [{ pos: 720, align: 'left' }, { pos: 1440, align: 'clear' }] }, r: [{ c: ['x'] }] };
  assert.equal(decodeStep([{ t: 'put', at: place, n: 1, x: [ok], h: '00000000' }]).length, 1);
  for (const tabs of [{ set: [{ pos: 720, align: 'left' }] }, { clearAll: true }, { clear: [720] }, null, 'x']) {
    assert.throws(() => decodeStep([{ t: 'put', at: place, n: 1, x: [{ p: { tabs }, r: [{ c: ['x'] }] }], h: '00000000' }]), { name: 'WireError' }, JSON.stringify(tabs));
  }
});
