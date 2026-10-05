// F4 while a formula is typed: the reference at the caret goes from relative to
// absolute and back the way Excel has it. Only the text is looked at here.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { formulaReferences } from '../src/calc/refs';
import { cycleAbsolute } from '../src/edit/absolute';

/** The text with the selection marked: "[" and "]" around it, "|" for a caret. */
function marked(result: { text: string; start: number; end: number } | undefined): string | undefined {
  if (!result) return undefined;
  const { text, start, end } = result;
  if (start === end) return `${text.slice(0, start)}|${text.slice(start)}`;
  return `${text.slice(0, start)}[${text.slice(start, end)}]${text.slice(end)}`;
}

/** Presses F4 on a text whose selection is marked as above. */
function press(text: string): string | undefined {
  const caret = text.indexOf('|');
  if (caret >= 0) return marked(cycleAbsolute(text.replace('|', ''), caret, caret));
  const start = text.indexOf('[');
  const end = text.indexOf(']') - 1;
  assert.ok(start >= 0 && end >= start, `no selection in ${text}`);
  return marked(cycleAbsolute(text.replace('[', '').replace(']', ''), start, end));
}

test('cycleAbsolute: a cell goes A1, $A$1, A$1, $A1 and back; a caret ends up right behind it, a selection stays one', () => {
  assert.equal(press('=A1|'), '=$A$1|');
  assert.equal(press('=[$A$1]'), '=[A$1]');
  assert.equal(press('=[A$1]'), '=[$A1]');
  assert.equal(press('=[$A1]'), '=[A1]');
  // Four presses in a row, each on what the one before left selected.
  let text = '=B7|*2';
  const seen: string[] = [];
  for (let i = 0; i < 4; i++) {
    text = press(text) as string;
    seen.push(text);
  }
  assert.deepEqual(seen, ['=$B$7|*2', '=B$7|*2', '=$B7|*2', '=B7|*2']);
});

test('cycleAbsolute: the caret may stand before, inside or right behind the reference', () => {
  assert.equal(press('=|A1+B2'), '=$A$1|+B2');
  assert.equal(press('=A|1+B2'), '=$A$1|+B2');
  assert.equal(press('=A1|+B2'), '=$A$1|+B2');
  assert.equal(press('=A1+|B2'), '=A1+$B$2|');
  assert.equal(press('=A1+B2|'), '=A1+$B$2|');
  assert.equal(press('=SUM(AB12|;3)'), '=SUM($AB$12|;3)');
  // Lower case is written the way references are.
  assert.equal(press('=a1|'), '=$A$1|');
});

test('cycleAbsolute: both corners of a range go together, from where the first corner is', () => {
  assert.equal(press('=SUM(A1:B2|)'), '=SUM($A$1:$B$2|)');
  assert.equal(press('=SUM([$A$1:$B$2])'), '=SUM([A$1:B$2])');
  assert.equal(press('=SUM([A$1:B$2])'), '=SUM([$A1:$B2])');
  assert.equal(press('=SUM([$A1:$B2])'), '=SUM([A1:B2])');
  assert.equal(press('=SUM(A|1:B2)'), '=SUM($A$1:$B$2|)');
  assert.equal(press('=SUM(A1:|B2)'), '=SUM($A$1:$B$2|)');
  // Corners that differ: the first one says where the reference is in the round.
  assert.equal(press('=SUM($A1:B$2|)'), '=SUM(A1:B2|)');
  assert.equal(press('=SUM(A1:$B$2|)'), '=SUM($A$1:$B$2|)');
});

test('cycleAbsolute: whole columns and whole rows have two states', () => {
  assert.equal(press('=SUM(A:A|)'), '=SUM($A:$A|)');
  assert.equal(press('=SUM([$A:$A])'), '=SUM([A:A])');
  assert.equal(press('=SUM($A:$B|)'), '=SUM(A:B|)');
  assert.equal(press('=SUM(B:|D)'), '=SUM($B:$D|)');
  assert.equal(press('=SUM(1:1|)'), '=SUM($1:$1|)');
  assert.equal(press('=SUM([$1:$1])'), '=SUM([1:1])');
  assert.equal(press('=SUM(3:|7)'), '=SUM($3:$7|)');
  assert.equal(press('=SUM($A:B|)'), '=SUM(A:B|)');
});

test('cycleAbsolute: the name of a sheet stays as it is written', () => {
  assert.equal(press('=Sheet1!A1|'), '=Sheet1!$A$1|');
  assert.equal(press('=Sheet1![$A$1]'), '=Sheet1![A$1]');
  assert.equal(press("='It''s'!$A$1|"), "='It''s'!A$1|");
  assert.equal(press("='Sales 2026'!B5:D9|+1"), "='Sales 2026'!$B$5:$D$9|+1");
  // The caret in the name of the sheet is at the reference too.
  assert.equal(press('=She|et1!A1'), '=Sheet1!$A$1|');
  assert.equal(press("='It|''s'!A1"), "='It''s'!$A$1|");
  // Both corners with a name, a run of sheets, another workbook.
  assert.equal(press('=Sheet1!A1:Sheet1!B2|'), '=Sheet1!$A$1:Sheet1!$B$2|');
  assert.equal(press('=SUM(Sheet1:Sheet3!A1|)'), '=SUM(Sheet1:Sheet3!$A$1|)');
  assert.equal(press('=[1]Sheet1!A1|'), '=[1]Sheet1!$A$1|');
  assert.equal(press('=Sheet1!A:A|'), '=Sheet1!$A:$A|');
});

test('cycleAbsolute: every reference the selection touches goes on, each from where it is', () => {
  assert.equal(press('=[A1+B2]*C3'), '=[$A$1+$B$2]*C3');
  assert.equal(press('=[$A$1+$B$2]*C3'), '=[A$1+B$2]*C3');
  assert.equal(press('=[A1+$B$2]*C3'), '=[$A$1+B$2]*C3');
  // The selection keeps what it held beside the references, and grows to hold those it only touched.
  assert.equal(press('[=A1+B2+1]'), '[=$A$1+$B$2+1]');
  assert.equal(press('=A[1+B]2+1'), '=[$A$1+$B$2]+1');
  assert.equal(press('=A1[+]B2'), '=[$A$1+$B$2]');
  assert.equal(press('=SUM([A1:B2;C:C;5:5];9)'), '=SUM([$A$1:$B$2;$C:$C;$5:$5];9)');
  assert.equal(press('=A1+[1+2]+B2'), undefined);
  // A selection that ends in the name of a sheet changes the reference and keeps its own end.
  assert.equal(press('=[A1+She]et1!B2'), '=[$A$1+She]et1!$B$2');
});

test('cycleAbsolute: what is no reference is left alone', () => {
  for (const text of [
    '=1+2|',
    '=SUM(|',
    '=LOG10|(5)',
    '=LO|G10(5)',
    '=A1B|',
    '=1E5|',
    '=1e+5|',
    '="A1|"',
    '="A1"|',
    '="say ""A1|"" twice"',
    '=#REF!A1|',
    '=Table1[A1|]',
    '=MyName|',
    '=A|',
    '=5|',
    '=XFE1|',
    '=A1048577|',
    '=A0|',
    '=TRUE|',
    '=|',
    '=A1 |+B2',
  ]) {
    assert.equal(press(text), undefined, text);
  }
  // Not a formula at all.
  assert.equal(press('A1|'), undefined);
  assert.equal(press("'=A1|"), undefined);
  assert.equal(press('|'), undefined);
  // A text in quotes holds no reference, the text around it does.
  assert.equal(press('="A1"&B2|'), '="A1"&$B$2|');
  assert.equal(press('=["A1"&B2]'), '=["A1"&$B$2]');
  // Behind a quote that is never closed nothing is a reference.
  assert.equal(press('=A1&"x+B2|'), undefined);
  assert.equal(press('=A1|&"x+B2'), '=$A$1|&"x+B2');
  // LOG10 without brackets is a cell, as it is to Excel.
  assert.equal(press('=LOG10|'), '=$LOG$10|');
  // The Turkish way of writing a formula changes nothing.
  assert.equal(press('=TOPLA(A1;B2|)*1,5'), '=TOPLA(A1;$B$2|)*1,5');
});

test('cycleAbsolute: a selection that is none, and a text too long to look at', () => {
  assert.equal(cycleAbsolute('=A1', -1, 2), undefined);
  assert.equal(cycleAbsolute('=A1', 2, 1), undefined);
  assert.equal(cycleAbsolute('=A1', 0, 9), undefined);
  assert.equal(cycleAbsolute('=A1', Number.NaN, 2), undefined);
  assert.equal(cycleAbsolute('=A1', 1.5, 2), undefined);
  assert.equal(cycleAbsolute(undefined as unknown as string, 0, 0), undefined);
  assert.equal(cycleAbsolute(`=A1+"${'x'.repeat(40000)}"`, 2, 2), undefined);
  // Many references: each is rewritten once.
  const many = `=${Array.from({ length: 2000 }, (_value, index) => `A${index + 1}`).join('+')}`;
  const result = cycleAbsolute(many, 0, many.length);
  assert.ok(result);
  assert.equal(result.text.split('$').length - 1, 4000);
  assert.equal(result.start, 0);
  assert.equal(result.end, result.text.length);
});

test('formulaReferences: where the references of a text are, and how they are written with other signs', () => {
  const found = formulaReferences("=SUM(A1,'My Sheet'!$B2:C$3,D:E,$4:$4)+LOG10(x1y)+\"A9\"");
  assert.deepEqual(
    found.map((ref) => [ref.start, ref.bodyStart, ref.end, ref.kind, ref.absCol, ref.absRow]),
    [
      [5, 5, 7, 'cell', false, false],
      [8, 19, 26, 'cell', true, false],
      [27, 27, 30, 'cols', false, false],
      [31, 31, 36, 'rows', false, true],
    ],
  );
  assert.deepEqual(
    found.map((ref) => ref.write(true, true)),
    ['$A$1', '$B$2:$C$3', '$D:$E', '$4:$4'],
  );
  assert.deepEqual(
    found.map((ref) => ref.write(false, false)),
    ['A1', 'B2:C3', 'D:E', '4:4'],
  );
  assert.deepEqual(formulaReferences(''), []);
  assert.deepEqual(formulaReferences(undefined as unknown as string), []);
  assert.deepEqual(formulaReferences("='never closed!A1"), []);
});
