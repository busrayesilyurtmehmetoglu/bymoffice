// What whole columns and many spills cost: guards against work that grows with the square of the formulas. The limits
// are about ten times what a free machine takes; they catch a stall, not a slowdown.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { getCell, open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const METADATA =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

async function engineOf(rows: string[], metadata?: string) {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${rows.join('')}</sheetData>`) }], metadata });
  const workbook = await (await open(bytes)).workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 0 } });
  let start = Date.now();
  engine.recalculate(engine.uncalculated());
  const opened = Date.now() - start;
  const type = (row: number, col: number, text: string) => {
    history.apply(setCellInput(workbook, 0, row, col, text), undefined, undefined);
    start = Date.now();
    engine.recalculate([{ sheet: 0, row, col }]);
    return Date.now() - start;
  };
  return { workbook, opened, type };
}

test('20,000 formulas that look up a short whole column: one edit they all read is not quadratic (E1)', async () => {
  for (const formula of ['IFERROR(MATCH(A{r},H:H,0),0)+$J$1', 'IFERROR(VLOOKUP(A{r},H:I,2,FALSE),0)+$J$1']) {
    const rows: string[] = [];
    for (let r = 1; r <= 20000; r++) {
      let cells = `<c r="A${r}"><v>${r % 100}</v></c><c r="G${r}"><f>${formula.replace(/\{r\}/g, String(r))}</f><v>0</v></c>`;
      if (r <= 100) cells += `<c r="H${r}"><v>${r % 50}</v></c><c r="I${r}"><v>${r}</v></c>`;
      if (r === 1) cells += '<c r="J1"><v>0</v></c>';
      rows.push(`<row r="${r}">${cells}</row>`);
    }
    const t = await engineOf(rows);
    // Every result changes: the engine writes 20,000 of them, and what a whole column holds is not looked for again.
    const took = t.type(0, 9, '1');
    assert.ok(took < 1500, `${formula}: ${took} ms`);
    assert.equal(getCell(t.workbook.sheets[0], 1, 6)?.value, 3);
  }
});

test('20,000 lookups that spill three columns each: opening and an edit do not walk every spill for every range (E2)', async () => {
  const rows: string[] = [];
  for (let r = 1; r <= 20000; r++) {
    const k = (r % 100) + 1;
    let cells = `<c r="A${r}"><v>${k}</v></c><c r="B${r}" cm="1"><f t="array" ref="B${r}:D${r}">_xlfn.XLOOKUP(A${r},H:H,I:K)</f><v>${k * 10}</v></c><c r="C${r}"><v>${k * 20}</v></c><c r="D${r}"><v>${k * 30}</v></c>`;
    cells += `<c r="E${r}"><f>SUM(B${r}:D${r})</f><v>${k * 60}</v></c>`;
    if (r <= 100) cells += `<c r="H${r}"><v>${r + 1}</v></c><c r="I${r}"><v>${(r + 1) * 10}</v></c><c r="J${r}"><v>${(r + 1) * 20}</v></c><c r="K${r}"><v>${(r + 1) * 30}</v></c>`;
    rows.push(`<row r="${r}">${cells}</row>`);
  }
  const t = await engineOf(rows, METADATA);
  assert.ok(t.opened < 3000, `open: ${t.opened} ms`);
  const unrelated = t.type(999, 25, '1');
  assert.ok(unrelated < 3000, `an unrelated edit: ${unrelated} ms`);
  const again = t.type(1000, 25, '2');
  assert.ok(again < 500, `another unrelated edit: ${again} ms`);
  const read = t.type(5, 0, '7');
  assert.ok(read < 500, `an edit one lookup reads: ${read} ms`);
  assert.equal(getCell(t.workbook.sheets[0], 5, 4)?.value, 7 * 10 + 7 * 20 + 7 * 30);
});
