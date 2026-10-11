// Random runs of copying, cutting and pasting cells that have rules, between the
// other things that change rules (see rules-random-fixtures.ts). What a paste is
// held against is worked out apart from the ranges of the rules, cell by cell:
// what a rule says for a cell, written as if that cell were A1, is the same for
// the cell it was copied from and the cell it was pasted to - and the rules of a
// cell are held against each other in the order they are applied in, which is
// what decides how the cell looks. After every paste and cut, and every save, the
// file is read back and says what the model says; everything undone gives the
// rules of the file back as they were written.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { shiftFormula } from '../src/calc/refs';
import { DataValidations } from '../src/format/validation';
import type { CellRange, Workbook } from '../src/index';
import { copyRange, cutRange, pasteAt, type CopiedRange } from '../src/ops/commands';
import { conditionalRules, copiedRulesExtent, mainPartRule } from '../src/ops/rules';
import { checkPackage, checkSheetOrder, load, partNames, partText, reopen, save } from './rules-fixtures';
import { ownCellProblems, randomBook, randomStep, rulesNow, rulesSeen, seeded } from './rules-random-fixtures';
import { canonical } from './workbook-content';

/** The block of cells of every sheet that the runs work in, and a little beyond it. */
const ROWS = 30;
const COLS = 22;

interface CellRules {
  /** The validation of the cell, as if the cell were A1; undefined for none. */
  validation: string | undefined;
  /** Whether that validation is one of the extension list. */
  fixed: boolean;
  /** The conditional rules of the cell that a copy takes along, as the main part states them, as if the cell were A1, in the order they are applied in. */
  carried: string[];
  /** The rules only the extension list has, in that order as well. */
  kept: string[];
  /** Both together, in the order they are applied in: 'c' and 'k' with the place of each in its list. */
  order: string;
}

/** What the rules of a sheet say for each cell of the block, each cell for itself. */
function cellRules(wb: Workbook, sheetIndex: number): Map<string, CellRules> {
  const out = new Map<string, CellRules>();
  const validations = new DataValidations(wb, wb.sheets[sheetIndex]);
  const infos = conditionalRules(wb, sheetIndex);
  for (let row = 0; row < ROWS; row++) {
    for (let col = 0; col < COLS; col++) {
      const entry: CellRules = { validation: undefined, fixed: false, carried: [], kept: [], order: '' };
      const rule = validations.ruleAt(row, col);
      if (rule) {
        const home = (text: string | undefined) => (text === undefined ? undefined : shiftFormula(text, -rule.ranges[0].top, -rule.ranges[0].left, true));
        entry.fixed = rule.extended === true;
        entry.validation = canonical({ ...rule, ranges: undefined, src: undefined, live: undefined, extended: undefined, formula1: home(rule.formula1), formula2: home(rule.formula2) });
      }
      for (const info of infos) {
        if (!info.ranges.some((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right)) continue;
        if (info.rule.extended) {
          entry.order += 'k';
          entry.kept.push(canonical({ ...info.rule, src: undefined, priority: 0 }));
          continue;
        }
        const plain = mainPartRule(info.rule)!;
        entry.order += 'c';
        // The format by what it says: a file that is read back numbers its records anew.
        // A rule that looks at its own cell says so by its kind: its formula names the first cell of the rule with whatever
        // "$" (held against that cell where the file is saved, see ownCellProblems()). Where the rule was copied from is
        // a place in a file, as `src` is.
        const formulas = OWN_CELL_KINDS.indexOf(plain.type) >= 0 ? [] : plain.formulas.map((text) => shiftFormula(text, -info.ranges[0].top, -info.ranges[0].left, true));
        // A format that says nothing about the font is read back with an empty record for it: the same format.
        const format = info.format && info.format.font && Object.keys(info.format.font).length === 0 ? { ...info.format, font: undefined } : info.format;
        entry.carried.push(canonical({ ...plain, origin: undefined, priority: 0, dxfId: undefined, dxf: undefined, formulas, format }));
      }
      if (entry.validation !== undefined || entry.carried.length > 0 || entry.kept.length > 0) out.set(`${row},${col}`, entry);
    }
  }
  return out;
}

/** Two long texts are the same; the message of a failure shows where they part. */
function sameText(actual: string, expected: string, label: string): void {
  // A format whose record has a <font> that says nothing the model holds (a size alone) is read back with an empty
  // record for the font; the model, once the format was changed, has none. The same format.
  const plain = (text: string) => text.split(',"font":{}').join('').split('"font":{},').join('');
  actual = plain(actual);
  expected = plain(expected);
  if (actual === expected) return;
  let at = 0;
  while (at < actual.length && actual[at] === expected[at]) at++;
  assert.fail(`${label}\n  actual:   ...${actual.slice(Math.max(0, at - 160), at + 160)}\n  expected: ...${expected.slice(Math.max(0, at - 160), at + 160)}`);
}

const OWN_CELL_KINDS = ['timePeriod', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors'];

const NONE: CellRules = { validation: undefined, fixed: false, carried: [], kept: [], order: '' };
/** What a cell says; for a cell that was pasted over, how its own rules and those that stayed stand to each other is not part of it. */
const said = (entry: CellRules | undefined, pasted = false) => canonical(pasted ? { ...(entry || NONE), order: '' } : entry || NONE);
const rulesText = (xml: string) => (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>|<dataValidations[\s\S]*?<\/dataValidations>|<extLst>[\s\S]*<\/extLst>/g) || []).join('\n');

async function run(seed: number, steps: number, extensionOnly: boolean): Promise<void> {
  const random = seeded(seed);
  const int = (limit: number) => Math.floor(random() * limit);
  const doc = await load(await randomBook({ conditional: true, extensionOnly }));
  const original = canonical(rulesNow(doc.wb));
  const sheets = doc.wb.sheets.length;
  let held: { copied: CopiedRange; cells: Map<string, CellRules> } | undefined;
  const trail: string[] = [];
  const all = () => doc.wb.sheets.map((_, s) => cellRules(doc.wb, s));
  /** Every cell of every sheet says what it said, but the cells of `range` of `sheet`, which say what `expected` gives. */
  const check = (before: Map<string, CellRules>[], sheet: number, range: CellRange | undefined, expected: (row: number, col: number, was: CellRules) => CellRules, where: string) => {
    const after = all();
    for (let s = 0; s < sheets; s++) {
      for (let row = 0; row < ROWS; row++) {
        for (let col = 0; col < COLS; col++) {
          const key = `${row},${col}`;
          const was = before[s].get(key) || NONE;
          const inside = s === sheet && range !== undefined && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
          assert.equal(said(after[s].get(key), inside), said(inside ? expected(row, col, was) : was, inside), `${where}: sheet ${s}, cell ${key}${inside ? '' : ', which was not touched'}`);
        }
      }
    }
    return after;
  };
  /** The file that is saved now says for every cell what the model says, in the same order. */
  const checkSaved = async (model: Map<string, CellRules>[], where: string) => {
    const saved = await save(doc);
    for (const name of partNames(saved)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) assert.deepEqual(ownCellProblems(await partText(saved, name)), [], `${where}: ${name}: rules that look at their own cell`);
    // What names an element of the file (the `uid` of a rule) is stated by one element: a copy of the rule states none.
    const ids: string[] = [];
    for (const name of partNames(saved)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) for (const found of (await partText(saved, name)).match(/<cfRule\b[^>]*\buid="[^"]*"/g) || []) ids.push(found.slice(found.lastIndexOf('uid=')));
    assert.deepEqual(ids.filter((id, at) => ids.indexOf(id) !== at), [], `${where}: an identity of a rule is stated twice`);
    const back = await reopen(saved);
    for (let s = 0; s < sheets; s++) {
      const read = cellRules(back, s);
      for (const key of new Set([...read.keys(), ...model[s].keys()])) assert.equal(said(read.get(key)), said(model[s].get(key)), `${where}: sheet ${s}, cell ${key} of the saved file`);
    }
  };
  for (let i = 0; i < steps; i++) {
    const roll = random();
    const where = `seed ${seed}, step ${i} (${trail.slice(-6).join(', ')})`;
    if (roll < 0.2) {
      const sheet = int(sheets);
      const top = int(12);
      const left = int(12);
      const range = { top, left, bottom: top + int(4), right: left + int(4) };
      const cells = cellRules(doc.wb, sheet);
      const cut = random() < 0.3;
      const copied = copyRange(doc.wb, sheet, range);
      // The cells of the copy by their place in it.
      const taken = new Map<string, CellRules>();
      for (let row = range.top; row <= range.bottom; row++) for (let col = range.left; col <= range.right; col++) taken.set(`${row - range.top},${col - range.left}`, cells.get(`${row},${col}`) || NONE);
      held = { copied, cells: taken };
      trail.push(`${cut ? 'cut' : 'copy'} ${sheet}:${range.top},${range.left}-${range.bottom},${range.right}`);
      if (cut) {
        const before = all();
        const ops = cutRange(doc.wb, sheet, range);
        if (ops.length > 0) doc.history.apply(ops, undefined, undefined);
        // What only the extension list says stays; everything else of the cells is gone.
        await checkSaved(check(before, sheet, range, (_row, _col, was) => ({ validation: was.fixed ? was.validation : undefined, fixed: was.fixed, carried: [], kept: was.kept, order: '' }), `${where}, cut`), `${where}, cut`);
      }
    } else if (roll < 0.5 && held) {
      const copied = held.copied;
      const cells = held.cells;
      const sheet = int(sheets);
      const extent = copiedRulesExtent(copied.rules);
      const height = Math.max(copied.cells.length, extent.height);
      const width = Math.max(copied.cells.reduce((max, row) => Math.max(max, row.length), 0), extent.width);
      const top = int(14);
      const left = int(14);
      const selection = { top, left, bottom: top + height * (1 + int(2)) - 1, right: left + width * (1 + int(2)) - 1 };
      trail.push(`paste ${sheet}:${top},${left}-${selection.bottom},${selection.right}`);
      const before = all();
      const result = pasteAt(doc.wb, sheet, selection, copied.text, copied);
      assert.deepEqual(result.range, selection, `${where}: what was pasted over`);
      if (result.ops.length > 0) doc.history.apply(result.ops, undefined, undefined);
      const model = check(
        before,
        sheet,
        result.range,
        (row, col, was) => {
          const from = cells.get(`${(row - top) % height},${(col - left) % width}`) || NONE;
          // A cell under a validation of the extension list keeps it; that list's conditional rules stay beside what comes.
          const validation = was.fixed ? was.validation : from.fixed ? undefined : from.validation;
          return { validation, fixed: was.fixed, carried: from.carried, kept: was.kept, order: '' };
        },
        `${where}, paste`,
      );
      await checkSaved(model, `${where}, paste`);
    } else if (roll < 0.58) {
      const sheet = int(sheets);
      const axis = random() < 0.5 ? 'row' : 'col';
      const at = int(8);
      const count = 1 + int(2);
      const kind = random() < 0.7 ? 'insert' : 'delete';
      trail.push(`${kind} ${axis} ${sheet}:${at}+${count}`);
      doc.history.apply([{ type: kind, sheet, axis, at, count }], undefined, undefined);
    } else if (roll < 0.7) {
      if (!doc.history.canUndo) continue;
      trail.push('undo');
      doc.history.undo();
    } else if (roll < 0.76) {
      if (!doc.history.canRedo) continue;
      trail.push('redo');
      doc.history.redo();
    } else {
      const step = randomStep(doc.wb, doc.history, random, { rulesOnly: true, conditional: true });
      if (!step || !step.ops || step.ops.length === 0) continue;
      trail.push(step.what);
      doc.history.apply(step.ops, undefined, undefined);
    }
    if (random() < 0.2 || i === steps - 1) {
      const saved = await save(doc);
      await checkPackage(saved);
      for (const name of partNames(saved)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) checkSheetOrder(await partText(saved, name), `${where}: ${name}`);
      const back = await reopen(saved);
      sameText(canonical(rulesNow(back)), canonical(rulesNow(doc.wb)), `${where}: the saved file has other rules`);
      assert.equal(canonical(rulesSeen(back)), canonical(rulesSeen(doc.wb)), `${where}: the saved file is looked up differently`);
    }
  }
  while (doc.history.canUndo) doc.history.undo();
  assert.equal(canonical(rulesNow(doc.wb)), original, `seed ${seed}: everything undone`);
  const last = await save(doc);
  await checkPackage(last);
  for (const name of partNames(doc.bytes)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) assert.equal(rulesText(await partText(last, name)), rulesText(await partText(doc.bytes, name)), `seed ${seed}: ${name}, everything undone`);
}

test('random runs of copy, cut and paste: every cell that is written says what the cell it is a copy of said, no other cell changes, and every save reads back as the model', async () => {
  for (let seed = 1; seed <= 60; seed++) await run(seed, 50, false);
});

test('the same in a sheet with a rule only the extension list has: that rule stays, whatever is pasted over it', async () => {
  for (let seed = 201; seed <= 240; seed++) await run(seed, 50, true);
});
