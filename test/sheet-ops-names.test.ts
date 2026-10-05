// Content kept the way the file has it (conditional formats, data validations) names sheets
// as the file does. These tests rename, delete and insert sheets, move rows in them, and
// check that the rules are evaluated for the sheets as they are now.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, ConditionalFormats, DataValidations, open, type Workbook } from '../src/index';
import { currentSheetName, fileSheetName, sheetNameMap } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const STYLES =
  `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${NS}"><fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>' +
  '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>';

const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
const expression = (sqref: string, formula: string, priority: number) =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="expression" dxfId="0" priority="${priority}"><formula>${formula}</formula></cfRule></conditionalFormatting>`;

/**
 * "Sheet1" has the rules; they read "Sheet2" (A1 = 7) and "Sheet3" (A1 = 0, A2 = 5):
 * A1 is filled while Sheet3!$A$2 > 0, B1 while Sheet2!$A$1 > 0, C1 while its own $F$1 > 0,
 * D1 has a data bar that ends at Sheet3!$A$2 * 2, E1 is filled while Other!$A$1 = 3 (no such sheet).
 * Validations: A5 a whole number between Sheet3!$A$2 and Sheet2!$A$1, B5 a list from Sheet2.
 */
async function workbook(): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      styles: STYLES,
      sheets: [
        {
          name: 'Sheet1',
          xml: worksheetXml(
            `<sheetData><row r="1">${num('A1', 1)}${num('B1', 1)}${num('C1', 1)}${num('D1', 5)}${num('E1', 1)}${num('F1', 1)}</row></sheetData>` +
              expression('A1', 'Sheet3!$A$2&gt;0', 1) +
              expression('B1', 'Sheet2!$A$1&gt;0', 2) +
              expression('C1', '$F$1&gt;0', 3) +
              '<conditionalFormatting sqref="D1"><cfRule type="dataBar" priority="4"><dataBar><cfvo type="num" val="0"/><cfvo type="formula" val="Sheet3!$A$2*2"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>' +
              expression('E1', 'Other!$A$1=3', 5) +
              '<dataValidations count="2">' +
              '<dataValidation type="whole" sqref="A5" showErrorMessage="1"><formula1>Sheet3!$A$2</formula1><formula2>Sheet2!$A$1</formula2></dataValidation>' +
              '<dataValidation type="list" sqref="B5" showErrorMessage="1"><formula1>Sheet2!$A$1:$A$2</formula1></dataValidation>' +
              '</dataValidations>',
          ),
        },
        { name: 'Sheet2', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 7)}</row><row r="2">${num('A2', 8)}</row></sheetData>`) },
        { name: 'Sheet3', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 0)}</row><row r="2">${num('A2', 5)}</row></sheetData>`) },
      ],
    }),
  );
  return file.workbook();
}

/** A workbook with what a view keeps beside it: history, engine, the rules of the first sheet. */
function session(wb: Workbook) {
  const host = wb.sheets[0];
  const history = new History(wb);
  const engine = new CalcEngine(wb);
  const formats = new ConditionalFormats(wb, host, { engine: () => engine, budget: Infinity });
  const rules = new DataValidations(wb, host, { engine: () => engine });
  /** After every change, as the view does it. */
  const refresh = () => {
    engine.invalidate();
    engine.recalculate(undefined, { overwrite: true });
    formats.invalidate();
    formats.begin();
  };
  refresh();
  return {
    host,
    history,
    apply: (ops: Op[]) => {
      history.apply(ops, undefined, undefined);
      refresh();
    },
    undo: () => {
      history.undo();
      refresh();
    },
    redo: () => {
      history.redo();
      refresh();
    },
    /** Which of A1, B1, C1 and E1 are filled, and the length of the bar of D1. */
    shown: () => {
      const at = wb.sheets.indexOf(host);
      assert.ok(at >= 0);
      const filled = ['A', 'B', 'C', 'E'].filter((letter) => {
        const found = formats.resolve(0, letter.charCodeAt(0) - 65);
        return !!found && !!found.style;
      });
      const bar = formats.resolve(0, 3)?.bar;
      return { filled: filled.join(''), bar: bar ? JSON.stringify(bar) : undefined };
    },
    limits: () => {
      const rule = rules.ruleAt(4, 0);
      return rule ? [rule.formula1, rule.formula2] : undefined;
    },
    list: () => rules.ruleAt(4, 1)?.formula1,
    items: () => rules.list(4, 1)?.items,
  };
}

test('file names: helpers answer for the file at once, not rename by rename', async () => {
  const wb = await workbook();
  const history = new History(wb);
  assert.equal(sheetNameMap(wb), undefined, 'nothing to map while the sheets are as in the file');
  assert.deepEqual(wb.sheets.map(fileSheetName), ['Sheet1', 'Sheet2', 'Sheet3']);
  assert.equal(currentSheetName(wb, 'Sheet2'), undefined);

  history.apply(deleteSheet(wb, 1), undefined, undefined);
  history.apply(renameSheet(wb, 1, 'Sheet2'), undefined, undefined);
  history.apply(insertSheet(wb, 2, { name: 'Sheet3' }), undefined, undefined);
  assert.deepEqual(
    wb.sheets.map((sheet) => [sheet.name, fileSheetName(sheet)]),
    [
      ['Sheet1', 'Sheet1'],
      ['Sheet2', 'Sheet3'],
      ['Sheet3', undefined],
    ],
  );
  assert.equal(currentSheetName(wb, 'Sheet2'), null, 'the Sheet2 of the file is the deleted sheet');
  assert.equal(currentSheetName(wb, 'SHEET2'), null);
  assert.equal(currentSheetName(wb, 'Sheet3'), 'Sheet2', 'the Sheet3 of the file is called Sheet2 now, not the new Sheet3');
  assert.equal(currentSheetName(wb, 'sheet3'), 'Sheet2');
  assert.equal(currentSheetName(wb, 'Sheet1'), undefined, 'unchanged');
  assert.equal(currentSheetName(wb, 'Nowhere'), undefined, 'not a sheet of the file: left as it is');
  const map = sheetNameMap(wb)!;
  assert.deepEqual(['Sheet1', 'Sheet2', 'Sheet3', 'x'].map(map), [undefined, null, 'Sheet2', undefined]);

  // A swap of two names.
  const swapped = await workbook();
  const swap = new History(swapped);
  swap.apply(renameSheet(swapped, 1, 'Tmp'), undefined, undefined);
  swap.apply(renameSheet(swapped, 2, 'Sheet2'), undefined, undefined);
  swap.apply(renameSheet(swapped, 1, 'Sheet3'), undefined, undefined);
  assert.deepEqual(['Sheet1', 'Sheet2', 'Sheet3', 'Tmp'].map(sheetNameMap(swapped)!), [undefined, 'Sheet3', 'Sheet2', undefined]);
  for (let i = 0; i < 3; i++) swap.undo();
  assert.equal(sheetNameMap(swapped), undefined);
  assert.equal(swapped.removedSheetNames, undefined);
});

test('conditional formats and validations follow renamed sheets', async () => {
  const wb = await workbook();
  const s = session(wb);
  const start = s.shown();
  assert.equal(start.filled, 'ABC');
  assert.ok(start.bar);
  assert.deepEqual(s.limits(), ['Sheet3!$A$2', 'Sheet2!$A$1']);
  assert.deepEqual(s.items(), ['7', '8']);

  s.apply(renameSheet(wb, 2, 'Kaynak veri'));
  assert.deepEqual(s.shown(), start, 'the rules read the sheet under its new name');
  assert.deepEqual(s.limits(), ["'Kaynak veri'!$A$2", 'Sheet2!$A$1']);
  // The value the rule reads changes on the renamed sheet: the rule sees it.
  s.apply(setCellInput(wb, 2, 1, 0, '-1'));
  assert.equal(s.shown().filled, 'BC');
  s.undo();
  // The sheet of the rules itself is renamed: its own references carry no name and stay.
  s.apply(renameSheet(wb, 0, 'Ana'));
  assert.deepEqual(s.shown(), start);
  s.apply(renameSheet(wb, 1, 'Liste'));
  assert.deepEqual(s.shown(), start);
  assert.deepEqual(s.limits(), ["'Kaynak veri'!$A$2", 'Liste!$A$1']);
  assert.equal(s.list(), 'Liste!$A$1:$A$2');
  assert.deepEqual(s.items(), ['7', '8']);
  for (let i = 0; i < 3; i++) s.undo();
  assert.deepEqual(s.shown(), start);
  assert.deepEqual(s.limits(), ['Sheet3!$A$2', 'Sheet2!$A$1']);
  for (let i = 0; i < 3; i++) s.redo();
  assert.deepEqual(s.limits(), ["'Kaynak veri'!$A$2", 'Liste!$A$1']);
  assert.deepEqual(s.shown(), start);
});

test('rows moved in a renamed sheet are matched by the name of the file; a deleted sheet is #REF!; a name used again is the new sheet', async () => {
  const wb = await workbook();
  const s = session(wb);
  const start = s.shown();

  // Rows inserted above what the rules read on Sheet3: the references follow ($A$2 -> $A$3).
  s.apply([{ type: 'insert', sheet: 2, axis: 'row', at: 0, count: 1 }]);
  assert.deepEqual(s.shown(), start);
  assert.deepEqual(s.limits(), ['Sheet3!$A$3', 'Sheet2!$A$1']);

  // Sheet2 goes: what read it is an error and fills nothing; a limit that is an error lets no value pass.
  s.apply(deleteSheet(wb, 1));
  assert.deepEqual(s.shown(), { filled: 'AC', bar: start.bar });
  assert.deepEqual(s.limits(), ['Sheet3!$A$3', '#REF!']);
  assert.equal(s.list(), '#REF!');
  assert.equal(s.items(), undefined);

  // Sheet3 takes the name "Sheet2". The file's "Sheet2!$A$1" is still the deleted sheet, and the
  // file's "Sheet3!$A$2" is this sheet, with the row inserted before the rename still counted.
  s.apply(renameSheet(wb, 1, 'Sheet2'));
  s.apply(setCellInput(wb, 1, 0, 0, '9'));
  assert.deepEqual(s.shown(), { filled: 'AC', bar: start.bar }, 'B1 does not come back for the new Sheet2!A1 = 9');
  assert.deepEqual(s.limits(), ['Sheet2!$A$3', '#REF!']);

  // Another row inserted after the rename: the edit belongs to the sheet the file calls Sheet3.
  s.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 1, count: 2 }]);
  assert.deepEqual(s.shown(), { filled: 'AC', bar: start.bar });
  assert.deepEqual(s.limits(), ['Sheet2!$A$5', '#REF!']);

  // A new sheet called "Sheet3": nothing of the file names it, whatever rows move in it.
  s.apply(insertSheet(wb, 1, { name: 'Sheet3' }));
  s.apply(setCellInput(wb, 1, 4, 0, '-5'));
  s.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 3 }]);
  assert.deepEqual(
    wb.sheets.map((sheet) => sheet.name),
    ['Sheet1', 'Sheet3', 'Sheet2'],
  );
  assert.deepEqual(s.shown(), { filled: 'AC', bar: start.bar });
  assert.deepEqual(s.limits(), ['Sheet2!$A$5', '#REF!']);

  // Rows of the sheet of the rules: its own references follow, with or without the sheets changed.
  s.apply([{ type: 'insert', sheet: 0, axis: 'col', at: 5, count: 1 }]);
  assert.equal(s.shown().filled, 'AC', 'C1 reads $G$1 now');
  s.apply(setCellInput(wb, 0, 0, 6, '0'));
  assert.equal(s.shown().filled, 'A');
  s.undo();
  s.undo();

  // A name the file does not have stays as written, and finds a sheet that takes it.
  s.apply(insertSheet(wb, 3, { name: 'Other' }));
  s.apply(setCellInput(wb, 3, 0, 0, '3'));
  assert.equal(s.shown().filled, 'ACE');
  s.undo();
  s.undo();

  // Everything undone, step by step, is the start again.
  while (s.history.canUndo) s.undo();
  assert.deepEqual(s.shown(), start);
  assert.deepEqual(s.limits(), ['Sheet3!$A$2', 'Sheet2!$A$1']);
  assert.deepEqual(s.items(), ['7', '8']);
  assert.equal(sheetNameMap(wb), undefined);
  // Redone, the sheet called "Other" included.
  while (s.history.canRedo) s.redo();
  assert.deepEqual(s.limits(), ['Sheet2!$A$5', '#REF!']);
  assert.deepEqual(s.shown(), { filled: 'ACE', bar: start.bar });
});

test('rules of a sheet that is not in the file follow its own rows and no name of the file', async () => {
  const wb = await workbook();
  const history = new History(wb);
  history.apply(insertSheet(wb, 3, { name: 'New' }), undefined, undefined);
  history.apply(renameSheet(wb, 1, 'Liste'), undefined, undefined);
  const added = wb.sheets[3];
  // What an editor of rules would put there one day; today nothing does.
  added.dataValidations = [
    { ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], type: 'whole', operator: 'between', formula1: '$B$2', formula2: 'Sheet2!$A$1+New!$B$2', allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop' },
  ];
  history.apply([{ type: 'insert', sheet: 3, axis: 'row', at: 0, count: 2 }], undefined, undefined);
  const rules = new DataValidations(wb, added);
  // Its own unqualified reference follows; "Sheet2" is the name of the file for the sheet now called Liste;
  // "New!" is not a name of the file, so the rows moved in the new sheet do not touch it.
  assert.deepEqual([rules.ruleAt(2, 0)?.formula1, rules.ruleAt(2, 0)?.formula2], ['$B$4', 'Liste!$A$1+New!$B$2']);
  assert.equal(rules.ruleAt(0, 0), undefined);
});
