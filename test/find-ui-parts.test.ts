// The parts of Find and Replace and of the toolbars that need no DOM: the names a list of
// fonts offers, and work done a slice at a time.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { COMMON_FONTS, MAX_FONT_ENTRIES, fontChoices } from '../src/ui/font-list';
import { runSliced } from '../src/ui/sliced';

test('font list: the common fonts and those of the file, each once whatever its case, in order', () => {
  const { names, shown } = fontChoices(['Bookman Old Style', 'arial', 'Zeta Sans'], 'calibri');
  for (const name of COMMON_FONTS) assert.ok(names.indexOf(name) >= 0, name);
  assert.ok(names.indexOf('Zeta Sans') >= 0);
  assert.equal(names.filter((name) => name.toLowerCase() === 'arial').length, 1);
  assert.deepEqual(names.slice().sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : 1)), names);
  // The entry of the font of the selection is the one the list has, whatever the case it was asked in.
  assert.equal(shown, 'Calibri');
  assert.equal(fontChoices([], undefined).shown, undefined);
});

test('font list: names that are not usable are not offered, but the font of the selection is shown; the list is bounded', () => {
  const odd = 'Bad\u0001Name';
  const long = 'L'.repeat(200);
  const { names } = fontChoices([odd, long, ''], undefined);
  assert.equal(names.indexOf(odd), -1);
  assert.equal(names.indexOf(long), -1);
  assert.equal(names.indexOf(''), -1);
  assert.equal(fontChoices([], long).shown, long);
  const many: string[] = [];
  for (let i = 0; i < 1000; i++) many.push(`Font ${i}`);
  const bounded = fontChoices(many, 'Font 999');
  assert.ok(bounded.names.length <= MAX_FONT_ENTRIES + 1, String(bounded.names.length));
  assert.equal(bounded.shown, 'Font 999');
});

test('sliced work: small work is done before the call returns; large work in several turns; it stops when asked', async () => {
  let calls = 0;
  let resolved = false;
  const quick = runSliced(() => ++calls >= 3, () => true).then((done) => {
    resolved = true;
    return done;
  });
  assert.equal(calls, 3);
  assert.equal(await quick, true);
  assert.equal(resolved, true);

  // Work that takes about 5 ms a step, 20 steps: more than one slice of 12 ms.
  let steps = 0;
  let turns = 0;
  let last = -1;
  const slow = await runSliced(
    () => {
      if (last !== turns) last = turns;
      const until = Date.now() + 5;
      while (Date.now() < until);
      return ++steps >= 20;
    },
    () => {
      turns++;
      return true;
    },
  );
  assert.equal(slow, true);
  assert.ok(turns > 3, String(turns));

  let alive = true;
  let count = 0;
  const stopped = runSliced(
    () => {
      count++;
      const until = Date.now() + 20;
      while (Date.now() < until);
      return false;
    },
    () => alive,
  );
  alive = false;
  assert.equal(await stopped, false);
  assert.equal(count, 1);

  await assert.rejects(
    runSliced(() => {
      throw new RangeError('bad');
    }, () => true),
    RangeError,
  );
});
