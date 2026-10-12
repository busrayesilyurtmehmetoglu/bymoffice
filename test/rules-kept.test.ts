// What the element of a rule holds beside the fields of the model (RuleKept) is in the model,
// with every name resolved to its namespace, and goes with the rule wherever the rule goes:
// pasted beside, pasted into another sheet, cut and pasted, copied with its sheet. What a
// sheet cannot state is known when the rule is put there, and counted; nothing is left out
// only when the file is written.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { open, parseRange, parseXml, type CellRange, type Workbook } from '../src/index';
import { buildXlsx as buildFromModel, conditionalRuleXml, keptRuleExtra, keptTarget } from '../src/formats/xlsx/build';
import { KEPT_MAIN, MAX_KEPT_DEPTH, MAX_KEPT_SIZE, packKept, unpackKept, type KeptContent, type RuleKept as PackedKept } from '../src/model/sheet/types';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import { conditionalFormatNow, conditionalRules, dataValidationNow, keptFor, ruleKeptNow, ruleSpace, setValidation, updateConditionalRule, validationKeptNow, type KeptView } from '../src/ops/rules';

/** What a rule holds beside the fields of the model, as a tree (the model holds it packed into a text). */
type RuleKept = KeptView;
const packed = (view: KeptView): PackedKept => (view.over ? { over: true } : { packed: packKept(view), ...(view.ignorable ? { ignorable: view.ignorable } : {}) });
import { copySheet, copySheetLosses, deleteSheet, renameSheet } from '../src/ops/sheets';
import { STYLES, cf } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, load, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';

const R = (text: string): CellRange => parseRange(text) as CellRange;
const apply = (doc: Loaded, ops: Parameters<Loaded['history']['apply']>[0]) => doc.history.apply(ops, undefined, undefined);
const formats = (xml: string) => xml.match(/<(?:\w+:)?conditionalFormatting[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/g) || [];
const validations = (xml: string) => xml.match(/<dataValidation [^>]*?(?:\/>|>[\s\S]*?<\/dataValidation>)/g) || [];

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const MINE = 'urn:example:mine';
const OTHER = 'urn:example:other';
const XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';

/** A sheet part with the root attributes `root` (beside the namespace of the part and that of relationships). */
const sheet = (root: string, inner: string) => worksheetXml(inner).replace('<worksheet ', `<worksheet ${root} `);
const book = (sheets: [name: string, xml: string][]) => buildXlsx({ sheets: sheets.map(([name, xml]) => ({ name, xml })), styles: STYLES });

/** The conditional rules of a sheet as they apply now, each with what it holds beside the fields of the model. */
function keptOf(wb: Workbook, index: number): (RuleKept | undefined)[] {
  const space = ruleSpace(wb, wb.sheets[index]);
  const out: (RuleKept | undefined)[] = [];
  for (const format of wb.sheets[index].conditionalFormats || []) {
    const now = conditionalFormatNow(format, space);
    if (now) for (const rule of now.rules) out.push(ruleKeptNow(rule));
  }
  return out;
}

function validationKept(wb: Workbook, index: number): (RuleKept | undefined)[] {
  const space = ruleSpace(wb, wb.sheets[index]);
  const out: (RuleKept | undefined)[] = [];
  for (const rule of wb.sheets[index].dataValidations || []) {
    const now = dataValidationNow(rule, space);
    if (now) out.push(validationKeptNow(now));
  }
  return out;
}

/** Without the word on what readers of the part may pass over, which is about the part. */
const content = (kept: RuleKept | undefined) => (kept ? { attrs: kept.attrs, children: kept.children, holder: kept.holder } : undefined);

/** Every sheet of the saved file says of its rules what the model says: the fields and what is held beside them. */
async function sameAsFile(doc: Loaded, saved: Uint8Array, label: string): Promise<Workbook> {
  const back = await reopen(saved);
  assert.equal(back.sheets.length, doc.wb.sheets.length, `${label}: the sheets of the file`);
  for (let i = 0; i < back.sheets.length; i++) {
    assert.deepEqual(keptOf(back, i), keptOf(doc.wb, i), `${label}: sheet ${i}, what the conditional rules hold beside the fields of the model`);
    assert.deepEqual(validationKept(back, i), validationKept(doc.wb, i), `${label}: sheet ${i}, what the data validations hold beside the fields of the model`);
    assert.deepEqual(
      conditionalRules(back, i).map((info) => [info.ranges, info.rule.type, info.rule.formulas, info.format]),
      conditionalRules(doc.wb, i).map((info) => [info.ranges, info.rule.type, info.rule.formulas, info.format]),
      `${label}: sheet ${i}, the rules`,
    );
  }
  return back;
}

const ICONS =
  '<cfRule type="iconSet" priority="5" xr:uid="{ICON}" my:mark="1" kept="yes"><iconSet iconSet="3Arrows" percent="0"><cfvo type="num" val="1" foo="bar" xr:uid="{CFVO}"><my:more a="1">text</my:more></cfvo><cfvo type="percent" val="33" gte="0"/><my:between/><cfvo type="percent" val="67"/></iconSet>' +
  '<extLst><ext uri="{KEEP}"><my:data/></ext></extLst></cfRule>';
const SCALE = '<cfRule type="colorScale" priority="6"><colorScale my:scale="s" plain="p"><cfvo type="min" gte="0"/><cfvo type="max"/><color rgb="FFF8696B" my:tone="t"/><color rgb="FF63BE7B"/><my:after/></colorScale></cfRule>';
const BAR = '<cfRule type="dataBar" priority="7"><dataBar minLength="5" my:bar="b"><cfvo type="min"/><cfvo type="max" my:max="m"/><color rgb="FF638EC6"><my:inColor/></color></dataBar></cfRule>';
const PLAIN = '<cfRule type="cellIs" dxfId="0" priority="8" operator="greaterThan"><formula>3</formula></cfRule>';

const KEPT_ICONS: RuleKept = {
  attrs: [
    [MINE, 'mark', '1'],
    ['', 'kept', 'yes'],
  ],
  children: [{ ns: KEPT_MAIN, name: 'extLst', children: [{ ns: KEPT_MAIN, name: 'ext', attrs: [['', 'uri', '{KEEP}']], children: [{ ns: MINE, name: 'data' }] }] }],
  holder: {
    kind: 'iconSet',
    attrs: [['', 'percent', '0']],
    between: [[], [], [{ ns: MINE, name: 'between' }], []],
    cfvo: [{ attrs: [['', 'foo', 'bar']], children: [{ ns: MINE, name: 'more', attrs: [['', 'a', '1']], children: ['text'] }] }, null, null],
  },
};
const KEPT_SCALE: RuleKept = {
  holder: {
    kind: 'colorScale',
    attrs: [
      [MINE, 'scale', 's'],
      ['', 'plain', 'p'],
    ],
    between: [[], [], [{ ns: MINE, name: 'after' }]],
    cfvo: [{ attrs: [['', 'gte', '0']] }, null],
    colors: [{ attrs: [[MINE, 'tone', 't']] }, null],
  },
};
const KEPT_BAR: RuleKept = { holder: { kind: 'dataBar', attrs: [[MINE, 'bar', 'b']], cfvo: [null, { attrs: [[MINE, 'max', 'm']] }], colors: [{ children: [{ ns: MINE, name: 'inColor' }] }] } };

const DATA = '<sheetData><row r="1"><c r="D1"><v>5</v></c></row></sheetData>';
const RULES = cf('D1:D5', ICONS) + cf('F1:F5', SCALE) + cf('H1:H5', BAR) + cf('J1:J5', PLAIN);
/** A book whose first sheet declares the namespace of the content as `how` says, and whose second declares what `target` says. */
const twoSheets = (how: 'root' | 'ignorable' | 'block', target = '', rules = RULES) => {
  const root = how === 'block' ? `xmlns:xr="${XR}"` : `xmlns:xr="${XR}" xmlns:my="${MINE}"${how === 'ignorable' ? ` ${MC} mc:Ignorable="my xr"` : ''}`;
  const inner = how === 'block' ? rules.replace(/<conditionalFormatting /g, `<conditionalFormatting xmlns:my="${MINE}" `) : rules;
  return book([
    ['Data', sheet(root, DATA + inner)],
    ['Calc', sheet(target, '<sheetData/>')],
  ]);
};

test('the reader takes what a rule holds beside the fields of the model into the model, names resolved to namespaces, without what names the element', async () => {
  for (const how of ['root', 'block', 'ignorable'] as const) {
    const wb = await reopen(await twoSheets(how));
    const rules = wb.sheets[0].conditionalFormats!.map((format) => format.rules[0]);
    const held = rules.map((rule) => ruleKeptNow(rule));
    const flagged = (kept: RuleKept): RuleKept => (how === 'ignorable' ? { ...kept, ignorable: [MINE] } : kept);
    assert.deepEqual(held[0], flagged(KEPT_ICONS), `${how}: the icon set`);
    assert.deepEqual(held[1], flagged(KEPT_SCALE), `${how}: the colour scale`);
    assert.deepEqual(held[2], flagged(KEPT_BAR), `${how}: the data bar`);
    assert.equal(rules[3].kept, undefined, `${how}: a rule that holds nothing else has no record of it`);
    // The model holds each as one text, which says the same when it is read again.
    for (const rule of rules) if (rule.kept) assert.deepEqual(packKept(unpackKept(rule.kept.packed) as KeptContent), rule.kept.packed, `${how}: the text of the model`);
    assert.ok(JSON.stringify(rules).indexOf('uid') < 0, `${how}: no identity of an element is in the model`);
    assert.deepEqual(wb.sheets[0].spaces, how === 'ignorable' ? { ignorable: [XR, MINE].sort() } : undefined, `${how}: what the root lets readers pass over`);
  }
});

test('a rule pasted into a sheet that does not declare the namespace is written with the declaration on its own element', async () => {
  for (const how of ['root', 'block'] as const) {
    // The second sheet binds the prefix to something else: the name in the file is no guide.
    const doc = await load(await twoSheets(how, `xmlns:my="${OTHER}"`));
    const copied = copyRange(doc.wb, 0, R('D1:H5'));
    const result = pasteAt(doc.wb, 1, R('B2'), copied.text, copied);
    apply(doc, result.ops);
    assert.equal(result.skippedRules, undefined, `${how}: nothing is left behind`);
    assert.deepEqual(keptOf(doc.wb, 1).map(content), [KEPT_ICONS, KEPT_SCALE, KEPT_BAR].map(content), `${how}: the pasted rules hold what the rules they are copies of hold`);
    const saved = await save(doc);
    await checkPackage(saved);
    const there = formats(await partText(saved, 'xl/worksheets/sheet2.xml'));
    assert.equal(there.length, 3, `${how}: three rules:\n${there.join('\n')}`);
    for (const text of there) assert.ok(/<cfRule [^>]*xmlns:ns1="urn:example:mine"/.test(text) && text.indexOf('my:') < 0, `${how}: the namespace is declared on the rule, under a prefix of its own, in:\n${text}`);
    assert.ok(there[0].indexOf('ns1:mark="1" kept="yes"') > 0 && there[0].indexOf('<cfvo type="num" val="1" foo="bar"><ns1:more a="1">text</ns1:more></cfvo>') > 0 && there[0].indexOf('<ns1:between/><cfvo type="percent" val="67"/>') > 0, `${how}: the icon set, in:\n${there[0]}`);
    assert.ok(there[0].indexOf('<extLst><ext uri="{KEEP}"><ns1:data/></ext></extLst></cfRule>') > 0, `${how}: the extension of the rule, in:\n${there[0]}`);
    assert.ok(there.join('').indexOf('uid') < 0 && there.join('').indexOf('{ICON}') < 0 && there.join('').indexOf('{CFVO}') < 0, `${how}: no identity of the elements they were copied from, in:\n${there.join('\n')}`);
    assert.ok(there[1].indexOf('<colorScale ns1:scale="s" plain="p"><cfvo type="min" gte="0"/><cfvo type="max"/><color rgb="FFF8696B" ns1:tone="t"/><color rgb="FF63BE7B"/><ns1:after/></colorScale>') > 0, `${how}: the scale, in:\n${there[1]}`);
    assert.ok(there[2].indexOf('<dataBar minLength="5" ns1:bar="b"><cfvo type="min"/><cfvo type="max" ns1:max="m"/><color rgb="FF638EC6"><ns1:inColor/></color></dataBar>') > 0, `${how}: the bar, in:\n${there[2]}`);
    await sameAsFile(doc, saved, how);
    while (doc.history.canUndo) doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${how}: everything undone, the file is the one that was opened`);
  }
});

test('content of a namespace readers may pass over goes only where the root says the same, and what does not go is counted when the rule is pasted', async () => {
  const cases: [what: string, target: string, skipped: number | undefined, prefix: string | undefined][] = [
    ['the other sheet declares nothing', '', 3, undefined],
    ['the other sheet declares the namespace, but lets no reader pass it over', `xmlns:my="${MINE}"`, 3, undefined],
    ['the other sheet says the same under the same prefix', `xmlns:my="${MINE}" ${MC} mc:Ignorable="my"`, undefined, 'my'],
    ['the other sheet says the same under another prefix', `xmlns:q="${MINE}" xmlns:my="${OTHER}" ${MC} mc:Ignorable="q"`, undefined, 'q'],
  ];
  for (const [what, target, skipped, prefix] of cases) {
    const doc = await load(await twoSheets('ignorable', target));
    const copied = copyRange(doc.wb, 0, R('D1:J5'));
    const result = pasteAt(doc.wb, 1, R('B2'), copied.text, copied);
    apply(doc, result.ops);
    assert.equal(result.skippedRules, skipped, `${what}: the rules that did not come as they are`);
    const saved = await save(doc);
    await checkPackage(saved);
    const there = formats(await partText(saved, 'xl/worksheets/sheet2.xml')).join('\n');
    if (prefix === undefined) {
      assert.ok(there.indexOf(MINE) < 0 && !/\b(my|ns\d):/.test(there), `${what}: nothing of the namespace is written, in:\n${there}`);
      // What is in no namespace, or in the one of the part, comes all the same.
      assert.ok(there.indexOf('kept="yes"') > 0 && there.indexOf('foo="bar"') > 0 && there.indexOf('<colorScale plain="p">') > 0 && there.indexOf('<extLst><ext uri="{KEEP}"/></extLst>') > 0, `${what}: the rest of what the rules hold, in:\n${there}`);
    } else {
      assert.ok(there.indexOf(`${prefix}:mark="1"`) > 0 && there.indexOf(`<${prefix}:between/>`) > 0 && there.indexOf(`${prefix}:tone="t"`) > 0 && there.indexOf(`<${prefix}:inColor/>`) > 0 && there.indexOf('xmlns:') < 0, `${what}: written under the prefix of the root, in:\n${there}`);
      assert.deepEqual(keptOf(doc.wb, 1), [KEPT_ICONS, KEPT_SCALE, KEPT_BAR, undefined].map((kept) => (kept ? { ...kept, ignorable: [MINE] } : kept)), `${what}: the model`);
    }
    // The model said so before the file was written: the file says what the model says.
    await sameAsFile(doc, saved, what);
  }
});

test('the same content pasted a second time gives the rule that is there the cells, also where part of it could not come', async () => {
  for (const target of ['', `xmlns:my="${MINE}" ${MC} mc:Ignorable="my"`]) {
    const doc = await load(await twoSheets('ignorable', target));
    const copied = copyRange(doc.wb, 0, R('D1:D5'));
    apply(doc, pasteAt(doc.wb, 1, R('B2'), copied.text, copied).ops);
    apply(doc, pasteAt(doc.wb, 1, R('E2'), copied.text, copied).ops);
    const there = formats(await partText(await save(doc), 'xl/worksheets/sheet2.xml'));
    assert.equal(there.length, 1, `one rule for both places (${target === '' ? 'cut down' : 'whole'}):\n${there.join('\n')}`);
    assert.ok(/sqref="B2:B6 E2:E6"/.test(there[0]), `with the cells of both, in:\n${there[0]}`);
  }
});

test('two rules that say the same in the model and beside it are one rule, whichever element each came from; two that differ beside it are not', async () => {
  const same = '<cfRule type="iconSet" priority="1" kept="yes"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0" foo="bar"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
  const other = same.replace('foo="bar"', 'foo="baz"');
  const twin = same.replace('priority="1"', 'priority="1" xr:uid="{TWIN}"');
  for (const [what, there, merged] of [
    ['the same text', same, true],
    ['the same but for the identity of the element', twin, true],
    ['another value of an attribute the model has no field for', other, false],
  ] as const) {
    const doc = await load(
      await book([
        ['Data', sheet(`xmlns:xr="${XR}"`, DATA + cf('D1:D5', same))],
        ['Calc', sheet(`xmlns:xr="${XR}"`, '<sheetData/>' + cf('A1:A5', there))],
      ]),
    );
    const copied = copyRange(doc.wb, 0, R('D1:D2'));
    apply(doc, pasteAt(doc.wb, 1, R('C1'), copied.text, copied).ops);
    const saved = formats(await partText(await save(doc), 'xl/worksheets/sheet2.xml'));
    if (merged) {
      assert.equal(saved.length, 1, `${what}: the rule of the sheet gets the cells:\n${saved.join('\n')}`);
      assert.ok(/sqref="A1:A5 C1:C2"/.test(saved[0]) && (what === 'the same text' || saved[0].indexOf('{TWIN}') > 0), `${what}: it stays the element it was, in:\n${saved[0]}`);
    } else {
      assert.equal(saved.length, 2, `${what}: a rule of its own:\n${saved.join('\n')}`);
      assert.ok(saved.some((text) => /sqref="A1:A5"/.test(text) && text.indexOf('baz') > 0) && saved.some((text) => /sqref="C1:C2"/.test(text) && text.indexOf('foo="bar"') > 0), `${what}: each with what it holds:\n${saved.join('\n')}`);
    }
  }
});

test('a sheet part that writes its names with a prefix gives its rules to one that does not, and takes theirs, with all they hold', async () => {
  const prefixed = (inner: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:worksheet xmlns:x="${NS_MAIN}" xmlns:r="${NS_REL}">${inner.replace(/<(\/?)(?![?!])/g, '<$1x:')}</x:worksheet>`;
  const rule = '<cfRule type="iconSet" priority="5" kept="yes"><iconSet iconSet="3Arrows" percent="0"><cfvo type="num" val="1" foo="bar"><extLst><ext uri="{IN}"/></extLst></cfvo><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet><extLst><ext uri="{KEEP}"/></extLst></cfRule>';
  const doc = await load(
    await book([
      ['Prefixed', prefixed(DATA + cf('D1:D5', rule))],
      ['Plain', worksheetXml(DATA + cf('D1:D5', rule.replace('{KEEP}', '{PLAIN}')))],
    ]),
  );
  assert.deepEqual(content(keptOf(doc.wb, 0)[0]), content(keptOf(doc.wb, 1)[0] && JSON.parse(JSON.stringify(keptOf(doc.wb, 1)[0]).replace('{PLAIN}', '{KEEP}'))), 'both parts read to the same model');
  let copied = copyRange(doc.wb, 0, R('D1:D5'));
  let result = pasteAt(doc.wb, 1, R('H1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, undefined, 'to the plain part: nothing is left behind');
  copied = copyRange(doc.wb, 1, R('D1:D5'));
  result = pasteAt(doc.wb, 0, R('K1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, undefined, 'to the prefixed part: nothing is left behind');
  const saved = await save(doc);
  await checkPackage(saved);
  const plain = formats(await partText(saved, 'xl/worksheets/sheet2.xml')).find((text) => /sqref="H1:H5"/.test(text)) || '';
  assert.equal(
    plain,
    '<conditionalFormatting sqref="H1:H5"><cfRule type="iconSet" priority="1" kept="yes"><iconSet iconSet="3Arrows" percent="0"><cfvo type="num" val="1" foo="bar"><extLst><ext uri="{IN}"/></extLst></cfvo><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet><extLst><ext uri="{KEEP}"/></extLst></cfRule></conditionalFormatting>',
    'in the plain part',
  );
  const other = formats(await partText(saved, 'xl/worksheets/sheet1.xml')).find((text) => /sqref="K1:K5"/.test(text)) || '';
  assert.equal(
    other,
    '<x:conditionalFormatting sqref="K1:K5"><x:cfRule type="iconSet" priority="1" kept="yes"><x:iconSet iconSet="3Arrows" percent="0"><x:cfvo type="num" val="1" foo="bar"><x:extLst><x:ext uri="{IN}"/></x:extLst></x:cfvo><x:cfvo type="percent" val="33"/><x:cfvo type="percent" val="67"/></x:iconSet><x:extLst><x:ext uri="{PLAIN}"/></x:extLst></x:cfRule></x:conditionalFormatting>',
    'in the prefixed part',
  );
  await sameAsFile(doc, saved, 'both parts');
});

test('a copy of a sheet has the rules with what they hold beside the fields of the model, and what a new part cannot state is counted before', async () => {
  const VALIDATIONS = `<dataValidations count="2"><dataValidation type="whole" imeMode="hiragana" my:note="n" xr:uid="{V1}" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation><dataValidation type="list" sqref="B1"><formula1>"a,b"</formula1></dataValidation></dataValidations>`;
  for (const how of ['root', 'ignorable'] as const) {
    const root = `xmlns:xr="${XR}" xmlns:my="${MINE}"${how === 'ignorable' ? ` ${MC} mc:Ignorable="my xr"` : ''}`;
    const doc = await load(await book([['Data', sheet(root, DATA + RULES + VALIDATIONS)]]));
    const losses = copySheetLosses(doc.wb, 0);
    assert.deepEqual(losses, how === 'ignorable' ? [{ what: 'ruleContent', count: 4 }] : [], `${how}: what the copy does not get`);
    apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
    const saved = await save(doc);
    await checkPackage(saved);
    const back = await sameAsFile(doc, saved, how);
    const copy = back.sheets.findIndex((item) => item.name === 'Copy');
    const text = await partText(saved, back.sheets[copy].part.slice(1));
    assert.ok(text.indexOf('uid') < 0 && text.indexOf('{ICON}') < 0, `${how}: no identity of an element of the sheet it is a copy of, in:\n${text}`);
    assert.ok(text.indexOf('imeMode="hiragana"') > 0 && text.indexOf('kept="yes"') > 0 && text.indexOf('<cfvo type="min" gte="0"/>') > 0 && text.indexOf('<dataBar minLength="5"') > 0, `${how}: what is in no namespace, in:\n${text}`);
    if (how === 'root') {
      assert.deepEqual(keptOf(back, copy).map(content), [KEPT_ICONS, KEPT_SCALE, KEPT_BAR, undefined].map(content), 'the conditional rules of the copy hold what those of the sheet hold');
      assert.deepEqual(validationKept(back, copy), [{ attrs: [['', 'imeMode', 'hiragana'], [MINE, 'note', 'n']] }, undefined], 'the data validations of the copy as well');
      assert.ok(text.indexOf(`xmlns:ns1="${MINE}"`) > 0, `the namespace is declared on the elements of the rules, in:\n${text}`);
    } else {
      assert.ok(text.indexOf(MINE) < 0, `nothing of the namespace is written, in:\n${text}`);
      assert.deepEqual(validationKept(back, copy), [{ attrs: [['', 'imeMode', 'hiragana']] }, undefined], 'the data validations of the copy, cut down');
    }
    // A file built from the model alone says the same of the sheet that was read.
    const built = await (await open(await buildFromModel(doc.wb))).workbook();
    assert.deepEqual(keptOf(built, 0).map(content), keptOf(back, copy).map(content), `${how}: a file built from the model`);
    while (doc.history.canUndo) doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${how}: everything undone, the file is the one that was opened`);
  }
});

test('what a scale or a bar holds beside the fields of the model stays when the rule is changed where it is', async () => {
  const doc = await load(await twoSheets('root'));
  for (let i = 0; i < 3; i++) apply(doc, updateConditionalRule(doc.wb, 0, i, { stopIfTrue: true }));
  const saved = await save(doc);
  const there = formats(await partText(saved, 'xl/worksheets/sheet1.xml')).join('\n');
  assert.ok(there.indexOf('<colorScale my:scale="s" plain="p"><cfvo type="min" gte="0"/><cfvo type="max"/><color rgb="FFF8696B" my:tone="t"/><color rgb="FF63BE7B"/><my:after/></colorScale>') > 0, `the scale, in:\n${there}`);
  assert.ok(there.indexOf('<dataBar minLength="5" my:bar="b"><cfvo type="min"/><cfvo type="max" my:max="m"/><color rgb="FF638EC6"><my:inColor/></color></dataBar>') > 0, `the bar, in:\n${there}`);
  assert.ok(there.indexOf('xr:uid="{ICON}"') > 0 && there.indexOf('xr:uid="{CFVO}"') > 0, `the rule that is changed where it is keeps the identity of its element, in:\n${there}`);
  await sameAsFile(doc, saved, 'changed where they are');
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'undone, the file is the one that was opened');
});

test('cut and pasted into another sheet, then the sheet they came from renamed, changed and deleted: the rules hold what they held', async () => {
  const doc = await load(await twoSheets('root'));
  const copied = copyRange(doc.wb, 0, R('D1:H5'));
  apply(doc, cutRange(doc.wb, 0, R('D1:H5')));
  const result = pasteAt(doc.wb, 1, R('B2'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, undefined, 'nothing is left behind');
  // In whatever order the sheet lists them.
  const sorted = (list: unknown[]) => list.map((item) => JSON.stringify(item)).sort();
  const expected = sorted([KEPT_ICONS, KEPT_SCALE, KEPT_BAR].map(content));
  apply(doc, renameSheet(doc.wb, 0, 'Renamed'));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 }]);
  let saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(sorted(keptOf(await sameAsFile(doc, saved, 'renamed and changed'), 1).map(content)), expected, 'after the sheet they came from was renamed and got rows');
  apply(doc, deleteSheet(doc.wb, 0));
  saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(sorted(keptOf(await sameAsFile(doc, saved, 'deleted'), 0).map(content)), expected, 'after the sheet they came from was deleted');
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: true }));
  saved = await save(doc);
  assert.deepEqual(sorted(keptOf(await sameAsFile(doc, saved, 'saved again'), 0).map(content)), expected, 'in a second save, when the part they were read from is in no package');
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone, the file is the one that was opened');
});

test('cells cut and pasted into a sheet that cannot state part of what their rule holds: counted, and the rule is there without it', async () => {
  const doc = await load(await twoSheets('ignorable'));
  const copied = copyRange(doc.wb, 0, R('D1:D5'));
  apply(doc, cutRange(doc.wb, 0, R('D1:D5')));
  const result = pasteAt(doc.wb, 1, R('B2'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, 1, 'the rule did not come as it is');
  const saved = await save(doc);
  const there = formats(await partText(saved, 'xl/worksheets/sheet2.xml')).join('\n');
  assert.ok(/sqref="B2:B6"/.test(there) && there.indexOf('kept="yes"') > 0 && there.indexOf(MINE) < 0, `the rule without what the sheet cannot state, in:\n${there}`);
  await sameAsFile(doc, saved, 'cut and pasted');
  // Pasted back into the sheet it came from, a copy of it holds all of it again only from the cells that still have it.
  doc.history.undo();
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'undone, the file is the one that was opened');
});

test('an element that holds more than a rule carries stays in the file as it is, and a copy of the rule comes without it, counted', async () => {
  const big = `<cfRule type="iconSet" priority="5" note="${'n'.repeat(MAX_KEPT_SIZE + 10)}"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`;
  let nested = '<deep/>';
  for (let i = 0; i < MAX_KEPT_DEPTH + 2; i++) nested = `<level>${nested}</level>`;
  const deep = `<cfRule type="cellIs" dxfId="0" priority="6" operator="equal"><formula>1</formula><extLst><ext uri="{DEEP}">${nested}</ext></extLst></cfRule>`;
  const unbound = '<cfRule type="cellIs" dxfId="0" priority="7" operator="equal" zz:what="1"><formula>2</formula></cfRule>';
  const bigValidation = `<dataValidations count="1"><dataValidation type="whole" note="${'v'.repeat(MAX_KEPT_SIZE + 10)}" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>`;
  const doc = await load(
    await book([
      ['Data', worksheetXml(DATA + cf('D1:D5', big) + cf('F1:F5', deep) + cf('H1:H5', unbound) + bigValidation)],
      ['Calc', worksheetXml('<sheetData/>')],
    ]),
  );
  assert.deepEqual(keptOf(doc.wb, 0), [{ over: true }, { over: true }, { over: true }], 'the model says that there is more, and holds none of it');
  assert.deepEqual(validationKept(doc.wb, 0), [{ over: true }], 'the same for a data validation');
  assert.ok(JSON.stringify(doc.wb.sheets[0].conditionalFormats).length < 4000, 'nothing of the content is in the model');
  sameBytes(await save(doc), doc.bytes, 'saved as it was opened');
  assert.deepEqual(copySheetLosses(doc.wb, 0), [{ what: 'ruleContent', count: 4 }], 'a copy of the sheet would not get it');
  // Changed where it is, the rule has all of it still.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: true }));
  let there = formats(await partText(await save(doc), 'xl/worksheets/sheet1.xml')).join('\n');
  assert.ok(there.indexOf('n'.repeat(MAX_KEPT_SIZE + 10)) > 0 && there.indexOf('stopIfTrue="1"') > 0, 'changed where it is');
  doc.history.undo();
  // Such a rule is kept as the file has it: a copy of its cells does not take it, in its own sheet or another, and says so.
  for (const target of [0, 1]) {
    const copied = copyRange(doc.wb, 0, R('A1:H5'));
    const result = pasteAt(doc.wb, target, R(target === 0 ? 'L1' : 'A1'), copied.text, copied);
    apply(doc, result.ops);
    assert.equal(result.skippedRules, 4, `pasted into sheet ${target}: counted, one for each rule`);
    const saved = await save(doc);
    await checkPackage(saved);
    assert.deepEqual(formats(await partText(saved, 'xl/worksheets/sheet1.xml')), formats(await partText(doc.bytes, 'xl/worksheets/sheet1.xml')), `pasted into sheet ${target}: the rules of the sheet are as they were`);
    const calc = await partText(saved, 'xl/worksheets/sheet2.xml');
    assert.ok(formats(calc).length === 0 && calc.indexOf('dataValidation') < 0, `pasted into sheet ${target}: no rule that says less than the one of the file, in:\n${calc}`);
    doc.history.undo();
  }
});

test('values and text a rule holds beside the fields of the model are written with the escapes of XML', async () => {
  const rule =
    '<cfRule type="cellIs" dxfId="0" priority="5" operator="equal" note="a &quot;b&quot; &lt;c&gt; &amp; d&#9;e&#10;f&#13;g \'h\'"><formula>1</formula><extLst><ext uri="{&amp;&lt;}"><my:text xmlns:my="urn:example:mine?a=1&amp;b=&quot;2&quot;">1 &lt; 2 &amp;&amp; 3 &gt; 2 ]]&gt;<![CDATA[<raw&>]]></my:text></ext></extLst></cfRule>';
  const doc = await load(
    await book([
      ['Data', worksheetXml(DATA + cf('D1:D5', rule))],
      ['Calc', worksheetXml('<sheetData/>')],
    ]),
  );
  const kept = keptOf(doc.wb, 0)[0]!;
  assert.deepEqual(kept.attrs, [['', 'note', 'a "b" <c> & d\te\nf\rg \'h\'']], 'the value as it is meant');
  const copied = copyRange(doc.wb, 0, R('D1:D5'));
  apply(doc, pasteAt(doc.wb, 1, R('A1'), copied.text, copied).ops);
  const saved = await save(doc);
  await checkPackage(saved);
  const text = await partText(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(text.indexOf('note="a &quot;b&quot; &lt;c> &amp; d&#9;e&#10;f&#13;g \'h\'"') > 0 || text.indexOf('note="a &quot;b&quot; &lt;c&gt; &amp; d&#9;e&#10;f&#13;g \'h\'"') > 0, `the attribute, in:\n${text}`);
  assert.ok(text.indexOf('xmlns:ns1="urn:example:mine?a=1&amp;b=&quot;2&quot;"') > 0, `the namespace, in:\n${text}`);
  assert.deepEqual(keptOf(await reopen(saved), 1).map(content), [content(kept)], 'read back, it is what it was');
  // A model put together by hand: an attribute is stated once, none declares a namespace, and a text that is no text of the model gives nothing.
  const base = conditionalRules(doc.wb, 0)[0].rule;
  const target = keptTarget('', [['xmlns', NS_MAIN]]);
  const made = '<k xmlns:n0="urn:&quot;q&quot;"><a ok="x &lt; y" n0:p="&lt;&amp;>" type="iconSet" priority="99" xmlns="urn:taken"/><c><plain>a &amp; b</plain></c></k>';
  const xml = conditionalRuleXml(base, 1, 0, undefined, '', keptRuleExtra(base, { packed: made }, target));
  const parsed = parseXml(`<conditionalFormatting xmlns="${NS_MAIN}">${xml}</conditionalFormatting>`).root;
  assert.equal(parsed.children.length, 1, `one well-formed element:\n${xml}`);
  assert.ok(xml.indexOf('ok="x &lt; y"') > 0 && xml.indexOf('<plain xmlns="">a &amp; b</plain>') > 0 && xml.indexOf('xmlns:ns1="urn:&quot;q&quot;"') > 0 && xml.indexOf('ns1:p="&lt;&amp;>"') > 0, `what can be said is said, in:\n${xml}`);
  assert.ok(xml.indexOf('iconSet') < 0 && xml.indexOf('99') < 0 && xml.indexOf('urn:taken') < 0, `no attribute is stated twice, and none declares a namespace, in:\n${xml}`);
  // What XML cannot carry is not written as it is.
  const odd = conditionalRuleXml(base, 1, 0, undefined, '', keptRuleExtra(base, { packed: '<k><a ok="x\u0001y"/></k>' }, target));
  assert.ok(odd.indexOf('\u0001') < 0, `no character XML cannot carry, in:\n${JSON.stringify(odd)}`);
  for (const broken of ['<k><a ok="1"', 'no text of the model', '<other><a ok="1"/></other>', '<!DOCTYPE k [<!ENTITY e "x">]><k><a ok="&e;"/></k>']) {
    assert.equal(conditionalRuleXml(base, 1, 0, undefined, '', keptRuleExtra(base, { packed: broken }, target)), conditionalRuleXml(base, 1, 0), `nothing is written of ${JSON.stringify(broken)}`);
  }
});

test('a file with such rules is saved as it was opened, and again after a change that is taken back', async () => {
  for (const how of ['root', 'ignorable', 'block'] as const) {
    const doc = await load(await twoSheets(how));
    sameBytes(await save(doc), doc.bytes, `${how}: opened and saved`);
    apply(doc, updateConditionalRule(doc.wb, 0, 1, { stopIfTrue: true }));
    const copied = copyRange(doc.wb, 0, R('D1:H5'));
    apply(doc, pasteAt(doc.wb, 1, R('B2'), copied.text, copied).ops);
    apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
    assert.notDeepEqual(await save(doc), doc.bytes, `${how}: changed`);
    while (doc.history.canUndo) doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${how}: changed, taken back and saved`);
  }
});

test('a data validation takes what its element holds beside the fields of the model along, to other cells, to another sheet and through a cut', async () => {
  const rules = `<dataValidations count="2"><dataValidation type="whole" imeMode="hiragana" my:note="n" xr:uid="{V1}" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2><my:more/></dataValidation><dataValidation type="whole" imeMode="off" sqref="C1:C3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>`;
  const KEPT: RuleKept = { attrs: [['', 'imeMode', 'hiragana'], [MINE, 'note', 'n']], children: [{ ns: MINE, name: 'more' }] };
  for (const how of ['root', 'ignorable'] as const) {
    const root = `xmlns:xr="${XR}" xmlns:my="${MINE}"${how === 'ignorable' ? ` ${MC} mc:Ignorable="my xr"` : ''}`;
    const doc = await load(
      await book([
        ['Data', sheet(root, DATA + rules)],
        ['Calc', worksheetXml('<sheetData/>')],
      ]),
    );
    assert.deepEqual(validationKept(doc.wb, 0), [how === 'ignorable' ? { ...KEPT, ignorable: [MINE] } : KEPT, { attrs: [['', 'imeMode', 'off']] }], `${how}: the model`);
    // Beside, in the sheet itself: the rule that says the same gets the cells - the one with the same content beside the fields.
    let copied = copyRange(doc.wb, 0, R('A1'));
    apply(doc, pasteAt(doc.wb, 0, R('E1'), copied.text, copied).ops);
    let saved = await save(doc);
    let there = validations(await partText(saved, 'xl/worksheets/sheet1.xml'));
    assert.ok(there.length === 2 && /sqref="A1:A3 E1"/.test(there[0]) && there[0].indexOf('xr:uid="{V1}"') > 0 && /sqref="C1:C3"/.test(there[1]), `${how}: beside:\n${there.join('\n')}`);
    // Into another sheet.
    copied = copyRange(doc.wb, 0, R('A1:C3'));
    const result = pasteAt(doc.wb, 1, R('B2'), copied.text, copied);
    apply(doc, result.ops);
    assert.equal(result.skippedRules, how === 'ignorable' ? 1 : undefined, `${how}: what did not come as it is`);
    saved = await save(doc);
    await checkPackage(saved);
    there = validations(await partText(saved, 'xl/worksheets/sheet2.xml'));
    assert.equal(there.length, 2, `${how}: two rules in the other sheet:\n${there.join('\n')}`);
    assert.ok(there.join('').indexOf('uid') < 0, `${how}: no identity of the element it was copied from:\n${there.join('\n')}`);
    if (how === 'root') assert.equal(there[0], `<dataValidation type="whole" xmlns:ns1="${MINE}" imeMode="hiragana" ns1:note="n" sqref="B2:B4"><formula1>1</formula1><formula2>9</formula2><ns1:more/></dataValidation>`, 'with all it holds');
    else assert.equal(there[0], '<dataValidation type="whole" imeMode="hiragana" sqref="B2:B4"><formula1>1</formula1><formula2>9</formula2></dataValidation>', 'with what the sheet can state');
    assert.equal(there[1], '<dataValidation type="whole" imeMode="off" sqref="D2:D4"><formula1>1</formula1><formula2>9</formula2></dataValidation>', `${how}: the other rule`);
    await sameAsFile(doc, saved, how);
    // A rule that is set anew for the cells holds nothing of the one that was there.
    apply(doc, setValidation(doc.wb, 1, R('B2:B4'), { type: 'whole', operator: 'between', formula1: '2', formula2: '3' }));
    there = validations(await partText(await save(doc), 'xl/worksheets/sheet2.xml'));
    assert.ok(there.some((text) => /sqref="B2:B4"/.test(text) && text.indexOf('imeMode="hiragana"') > 0), `${how}: the rule that is edited stays the rule it is:\n${there.join('\n')}`);
    while (doc.history.canUndo) doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${how}: everything undone, the file is the one that was opened`);
  }
});

test('a format that keeps a font the model has no word for says so in the model as in the file', async () => {
  // Record 2 of the styles: <font><i/><sz val="14"/></font> with a fill, an alignment and a border.
  const doc = await load(await book([['Data', worksheetXml(DATA + cf('D1:D5', '<cfRule type="cellIs" dxfId="2" priority="1" operator="equal"><formula>1</formula></cfRule>'))]]));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { format: { fill: '#00FF00' } }));
  const saved = await save(doc);
  const styles = await partText(saved, 'xl/styles.xml');
  assert.ok(styles.indexOf('<dxf><font><sz val="14"/></font><fill>') > 0, `the size of the font stays in the record, in:\n${styles}`);
  const back = await reopen(saved);
  assert.deepEqual(conditionalRules(back, 0)[0].format, conditionalRules(doc.wb, 0)[0].format, 'the format of the model is the format of the file');
  assert.deepEqual(conditionalRules(doc.wb, 0)[0].format!.font, {}, 'a font that says nothing the model has a word for');
  // Where the record has a font of nothing but what the model holds, taking that away leaves no font: in the model and in the file.
  const bold = await load(await book([['Data', worksheetXml(DATA + cf('D1:D5', '<cfRule type="cellIs" dxfId="1" priority="1" operator="equal"><formula>1</formula></cfRule>'))]]));
  apply(bold, updateConditionalRule(bold.wb, 0, 0, { format: { fill: '#00FF00' } }));
  const again = await reopen(await save(bold));
  assert.deepEqual(conditionalRules(again, 0)[0].format, conditionalRules(bold.wb, 0)[0].format, 'bold taken away: the same in the model and in the file');
  assert.equal(conditionalRules(bold.wb, 0)[0].format!.font, undefined, 'no font at all');
});

test('the tie of a rule to its counterpart in the extension list is no part of what a copy holds; an extension of its own beside it is', async () => {
  const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
  const bar = `<cfRule type="dataBar" priority="1" xr:uid="{BAR}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{OWN}"><own/></ext><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14}><x14:id>{TIE-1}</x14:id></ext></extLst></cfRule>`;
  const doc = await load(
    await book([
      ['Data', sheet(`xmlns:xr="${XR}"`, DATA + cf('D1:D5', bar))],
      ['Calc', worksheetXml('<sheetData/>')],
    ]),
  );
  assert.deepEqual(keptOf(doc.wb, 0), [{ children: [{ ns: KEPT_MAIN, name: 'extLst', children: [{ ns: KEPT_MAIN, name: 'ext', attrs: [['', 'uri', '{OWN}']], children: [{ ns: KEPT_MAIN, name: 'own' }] }] }] }], 'the model: the extension of its own, not the tie');
  assert.equal(doc.wb.sheets[0].conditionalFormats![0].rules[0].extId, '{TIE-1}', 'the tie is a field of the rule of the file');
  const copied = copyRange(doc.wb, 0, R('D1:D5'));
  const result = pasteAt(doc.wb, 1, R('A1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, 1, 'what the extension list says of the rule does not come along, which is counted');
  const saved = await save(doc);
  await checkPackage(saved);
  const there = formats(await partText(saved, 'xl/worksheets/sheet2.xml')).join('\n');
  assert.ok(there.indexOf('{TIE-1}') < 0 && there.indexOf('uid') < 0 && there.indexOf('x14') < 0, `no tie and no identity, in:\n${there}`);
  assert.ok(there.indexOf('<extLst><ext uri="{OWN}"><own/></ext></extLst>') > 0, `its own extension, in:\n${there}`);
  assert.ok((await partText(saved, 'xl/worksheets/sheet1.xml')).indexOf('<x14:id>{TIE-1}</x14:id>') > 0, 'the rule it was copied from keeps its tie');
});

test('keptFor: what a sheet cannot state is taken out and said, the rest is the object it was', () => {
  const kept = packed({
    attrs: [['', 'a', '1'], [MINE, 'b', '2']],
    children: [{ ns: KEPT_MAIN, name: 'extLst', children: [{ ns: MINE, name: 'gone' }, 'text', { ns: OTHER, name: 'stays', attrs: [[MINE, 'x', '1']] }] }],
    holder: { kind: 'iconSet', cfvo: [{ attrs: [[MINE, 'c', '3']] }, null, null] },
    ignorable: [MINE],
  });
  const view = (fitted: { kept: PackedKept | undefined; lost: boolean }) => ({ kept: validationKeptNow({ kept: fitted.kept }), lost: fitted.lost });
  assert.equal(keptFor(kept, { ignorable: [MINE] }).kept, kept, 'a sheet that says the same: the object itself');
  assert.deepEqual(view(keptFor(kept, { ignorable: [OTHER] })), { kept: { attrs: [['', 'a', '1']], children: [{ ns: KEPT_MAIN, name: 'extLst', children: ['text', { ns: OTHER, name: 'stays' }] }], ignorable: [OTHER] }, lost: true }, 'a sheet that does not, and lets readers pass over another namespace the content uses');
  assert.deepEqual(view(keptFor(packed({ attrs: [[MINE, 'b', '2']], ignorable: [MINE] }), undefined)), { kept: undefined, lost: true }, 'nothing is left');
  assert.deepEqual(view(keptFor(packed({ attrs: [[MINE, 'b', '2']] }), { ignorable: [MINE, OTHER] })), { kept: { attrs: [[MINE, 'b', '2']], ignorable: [MINE] }, lost: false }, 'what the sheet lets readers pass over is said for the sheet');
  assert.deepEqual(keptFor({ over: true }, { ignorable: [MINE] }), { kept: undefined, lost: true }, 'what is not carried');
  assert.deepEqual(keptFor(undefined, undefined), { kept: undefined, lost: false }, 'nothing to begin with');
});

test('files with rules that were pasted and copied are read by openpyxl as the file they were made from is', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not installed');
    return;
  }
  // A folder of the kind the repository ignores (tmp-*), gone when the test ends.
  const dir = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-rules-kept-${process.pid}`);
  mkdirSync(dir, { recursive: true });
  try {
    // openpyxl drops a rule with an attribute it does not know (in the file as it was made, too): beside those, rules it reads.
    const own = '<cfRule type="dataBar" priority="9"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{OWN}"><my:own/></ext></extLst></cfRule>';
    const files: string[] = [];
    for (const [how, target] of [
      ['root', `xmlns:my="${OTHER}"`],
      ['ignorable', ''],
      ['ignorable', `xmlns:q="${MINE}" ${MC} mc:Ignorable="q"`],
      ['block', ''],
    ] as const) {
      const doc = await load(await twoSheets(how, target, RULES + cf('L1:L5', own)));
      const copied = copyRange(doc.wb, 0, R('A1:L5'));
      apply(doc, pasteAt(doc.wb, 1, R('A1'), copied.text, copied).ops);
      apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
      const file = join(dir, `${files.length}.xlsx`);
      writeFileSync(file, await save(doc));
      files.push(file);
    }
    const script =
      'import sys, warnings, openpyxl\nwarnings.simplefilter("ignore")\nfor name in sys.argv[1:]:\n    wb = openpyxl.load_workbook(name)\n    print("|".join(ws.title + "=" + ",".join(sorted(str(c.sqref) + ":" + r.type for c in ws.conditional_formatting for r in c.rules)) for ws in wb.worksheets))\n';
    const result = spawnSync('python3', ['-c', script, ...files], { encoding: 'utf8' });
    assert.equal(result.status, 0, `openpyxl reads the files: ${result.stderr}`);
    const rules = 'H1:H5:dataBar,J1:J5:cellIs,L1:L5:dataBar';
    assert.deepEqual(result.stdout.trim().split('\n'), files.map(() => `Data=${rules}|Copy=${rules}|Calc=${rules}`), 'the sheet, its copy and the sheet its cells were pasted into, each with the rules openpyxl takes');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
