// Collaborative editing, the first things: two participants and cells.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { setCellInput } from '../src/ops/commands';
import { formatRange } from '../src/ops/format';
import { getCell } from '../src/index';
import { Net, act, assertSame, join, seeded, startFile } from './collab-fixtures';

test('what one participant types, the other one has', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(1);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  assert.equal(a.session.state, 'online', 'a is online');
  assert.equal(b.session.peers.length, 1, 'b sees a');
  assert.equal(act(a, () => setCellInput(a.wb, 0, 1, 1, '42')), true, 'a typed');
  assert.equal(a.session.pending, 1, 'the step waits for the relay');
  net.drain(random);
  assert.equal(a.session.pending, 0, 'the relay placed the step');
  assert.equal(getCell(b.wb.sheets[0], 1, 1)?.value, 42, 'b has the value');
  assert.equal(b.session.rev, 1, 'b is at the step');
  assertSame([a, b], 'after one step');
});

test('two participants who write the same cell at once end up with the one the relay placed last', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(2);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 2, 1, 'from a'));
  act(b, () => setCellInput(b.wb, 0, 2, 1, 'from b'));
  net.drain(random);
  assertSame([a, b], 'the same cell');
  assert.match(String(getCell(a.wb.sheets[0], 2, 1)?.value), /^from [ab]$/, 'one of the two texts');
  assert.equal(a.notices.length + b.notices.length, 0, 'nothing was dropped');
});

test('a format and a value of one cell that are changed at once both stay', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(3);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  act(a, () => formatRange(a.wb, 0, { top: 3, left: 1, bottom: 3, right: 1 }, { bold: true }));
  act(b, () => setCellInput(b.wb, 0, 3, 1, '77'));
  net.drain(random);
  assertSame([a, b], 'format and value');
  const cell = getCell(a.wb.sheets[0], 3, 1)!;
  assert.equal(cell.value, 77, 'the value of b');
  assert.equal(a.wb.styles.fonts[a.wb.styles.cellXfs[cell.style].fontId].bold, true, 'the format of a');
});
