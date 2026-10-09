// Fourth review of the workbook view, what it decides without a DOM: which rules a long list
// shows, and when a copy of cells no longer fits the workbook.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_MANAGER_ROWS, managerPage } from '../src/ui/rules-dialog';
import { copyOutdated } from '../src/ui/sheet-view';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { foldSheetName } from '../src/model/sheet/types';

test('review 4 (B8): a list of rules shows them in runs of 200, and every rule is in one of them', () => {
  assert.equal(MAX_MANAGER_ROWS, 200);
  assert.deepEqual(managerPage(0, 0), { from: 0, to: 0 });
  assert.deepEqual(managerPage(200, 0), { from: 0, to: 200 });
  assert.deepEqual(managerPage(201, 0), { from: 0, to: 200, next: 200 });
  assert.deepEqual(managerPage(201, 200), { from: 200, to: 201, previous: 0 });
  assert.deepEqual(managerPage(650, 200), { from: 200, to: 400, previous: 0, next: 400 });
  assert.deepEqual(managerPage(650, 600), { from: 600, to: 650, previous: 400 });
  // Every rule of a long list is reached by going on from the first run.
  const seen: number[] = [];
  let page = managerPage(1234, 0);
  for (let turns = 0; turns < 100; turns++) {
    for (let i = page.from; i < page.to; i++) seen.push(i);
    if (page.next === undefined) break;
    page = managerPage(1234, page.next);
  }
  assert.equal(seen.length, 1234);
  assert.deepEqual(seen.slice(0, 3).concat(seen.slice(-2)), [0, 1, 2, 1232, 1233]);
});

test('review 4 (B8): a place no rule is at any more (rules were deleted) gives the last run there is, and odd places a run', () => {
  assert.deepEqual(managerPage(150, 200), { from: 0, to: 150 });
  assert.deepEqual(managerPage(400, 400), { from: 200, to: 400, previous: 0 });
  assert.deepEqual(managerPage(401, 999999), { from: 400, to: 401, previous: 200 });
  assert.deepEqual(managerPage(450, 317), { from: 200, to: 400, previous: 0, next: 400 });
  for (const from of [-5, Number.NaN, Number.POSITIVE_INFINITY, 0.5]) {
    const page = managerPage(450, from);
    assert.ok(page.from % 200 === 0 && page.from >= 0 && page.from < 450 && page.to > page.from, String(from));
  }
  assert.deepEqual(managerPage(Number.NaN, 0), { from: 0, to: 0 });
  assert.deepEqual(managerPage(-3, 0), { from: 0, to: 0 });
  // Another size, as a test would use it.
  assert.deepEqual(managerPage(7, 3, 3), { from: 3, to: 6, previous: 0, next: 6 });
});

test('review 4 (B8): the words of the list name the rules shown and the way to the others, in both languages', () => {
  for (const strings of [STRINGS_EN.rules, STRINGS_TR.rules]) {
    const shown = strings.rulesShown(201, 400, 650);
    for (const number of ['201', '400', '650']) assert.ok(shown.indexOf(number) >= 0, shown);
    assert.ok(strings.nextRules(200).indexOf('200') >= 0);
    assert.ok(strings.previousRules(200).indexOf('200') >= 0);
    assert.notEqual(strings.nextRules(200), strings.previousRules(200));
    assert.ok(strings.dialogOpen.length > 10);
  }
});

const names = (...list: string[]) => list.map((name) => ({ name }));
const folded = (...list: string[]) => list.map(foldSheetName);

test('review 4 (B9): a copy of cells stays what it is while every sheet of then is still there under its name', () => {
  const then = folded('Veri', 'Özet');
  // A sheet moved, one hidden or coloured (the names are what they were), one added, one copied.
  assert.equal(copyOutdated(then, names('Özet', 'Veri')), false);
  assert.equal(copyOutdated(then, names('Veri', 'Özet')), false);
  assert.equal(copyOutdated(then, names('Veri', 'Sayfa3', 'Özet')), false);
  assert.equal(copyOutdated(then, names('Veri', 'Veri (2)', 'Özet')), false);
  // A name that differs by the case of its letters is the same sheet.
  assert.equal(copyOutdated(then, names('VERİ', 'özet')), copyOutdated(folded('VERİ', 'özet'), names('Veri', 'Özet')));
});

test('review 4 (B9): a copy is forgotten once a sheet of then is gone or has another name', () => {
  const then = folded('Veri', 'Özet');
  assert.equal(copyOutdated(then, names('Veri')), true);
  assert.equal(copyOutdated(then, names('Veri', 'Sonuç')), true);
  assert.equal(copyOutdated(then, names('Veri', 'Sonuç', 'Yeni')), true);
  // Nothing is known of when the copy was taken: it is not trusted.
  assert.equal(copyOutdated(undefined, names('Veri', 'Özet')), true);
  // A sheet that was added after the copy and removed again takes nothing the copy named.
  assert.equal(copyOutdated(then, names('Veri', 'Özet')), false);
});
