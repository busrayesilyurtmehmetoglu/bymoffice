// A participant, or the relay, that is not honest: what it sends is refused or skipped the
// same way by everybody, changes nobody's workbook into another one than the others have, and
// never gets into a file as anything but text.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { test } from 'node:test';
import { ZipArchive, open } from '../src/index';
import { contentVersion } from '../src/collab/protocol';
import { decodeStep } from '../src/collab/wire';
import { setCellInput } from '../src/ops/commands';
import { packageProblems } from './chart-edit-fixtures';
import { Net, act, assertSame, canon, join, saveOf, seeded, startFile, type Member } from './collab-fixtures';

/** A connection that says whatever the test tells it to. */
function intruder(net: Net, bytes: Uint8Array, key = 'evil'): { send(message: unknown): void; got: { t: string; rev?: number; id?: string }[] } {
  const raw = net.server.connect();
  const got: { t: string; rev?: number; id?: string }[] = [];
  raw.onMessage((text) => got.push(JSON.parse(text)));
  raw.onState!((state) => {
    if (state === 'open') raw.send(JSON.stringify({ t: 'join', v: [contentVersion(bytes)], key, pv: 2 }));
  });
  return { send: (message) => raw.send(JSON.stringify(message)), got };
}

const FONT = { bold: false, italic: false, strike: false, underline: 'none' };
const PLAIN = { n: 0, f: FONT, l: { type: 'pattern', pattern: 'none' }, b: { diagonalUp: false, diagonalDown: false } };

test('a step under the id of someone else\'s next step is a step of whoever sent it: nobody takes it for their own', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(1);
  const victim = await join(net, 'victim', bytes);
  const bystander = await join(net, 'bystander', bytes);
  const evil = intruder(net, bytes);
  net.drain(random);
  act(victim, () => setCellInput(victim.wb, 0, 20, 0, 'first'));
  net.drain(random);
  const seen = evil.got.filter((message) => message.t === 'op').pop()!;
  const [key, seq] = seen.id!.split(':');
  // The victim's next id, on a connection that is not the victim's.
  evil.send({ t: 'op', base: seen.rev, id: `${key}:${Number(seq) + 1}`, ops: [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'string', v: 'ATTACKER' } }]] }] });
  net.deliverUp();
  act(victim, () => setCellInput(victim.wb, 0, 21, 0, 'second'));
  net.drain(random);
  assertSame([victim, bystander], 'after the forged step');
  assert.equal(victim.session.pending, 0, 'the step of the victim was placed');
  assert.equal(canon(bystander.wb).indexOf('"second"') >= 0, true, 'and the others have it');
});

test('a "saved" that names the file the room started with does not move that file to a later step', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(2);
  const a = await join(net, 'a', bytes);
  net.drain(random);
  for (let i = 0; i < 5; i++) act(a, () => setCellInput(a.wb, 0, 20 + i, 0, `v${i}`));
  net.drain(random);
  const evil = intruder(net, bytes);
  net.drain(random);
  evil.send({ t: 'saved', v: [contentVersion(bytes)], rev: 5 });
  net.drain(random);
  const late = await join(net, 'late', bytes);
  net.drain(random);
  assert.equal(late.session.rev, 5, 'who joins with the first file gets every step');
  assertSame([a, late], 'after the false "saved"');
});

test('what a file would take as markup, or not at all, is no step', async () => {
  const { wb } = await join(new Net(), 'x', await startFile());
  const cell = (content: unknown): unknown[] => [{ t: 'cells', s: 0, c: [[0, 0, { v: content }]] }];
  const format = (change: object): unknown[] => [{ t: 'st', f: [{ ...PLAIN, ...change }] }, { t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }];
  const bad: [string, unknown[]][] = [
    ['an attribute of a formula that is none', cell({ k: 'number', v: 1, f: { x: 'A1', a: [['x="1"><v>9</v></f><f', 'y']] } })],
    ['a formula that says it is shared', cell({ k: 'number', v: 1, f: { x: 'A1', a: [['t', 'shared'], ['si', '99']] } })],
    ['a formula attribute with another value than a flag', cell({ k: 'number', v: 1, f: { x: 'A1', a: [['ca', '"><x']] } })],
    ['an array formula', cell({ k: 'number', v: 1, f: { x: 'A1:A3', t: 'array', r: 'A1:XFD1048576' } })],
    ['a data table', cell({ k: 'number', v: 1, f: { x: '', t: 'dataTable', r: 'B2:C3' } })],
    ['a date that is none', cell({ k: 'date', v: 'not a date <x>' })],
    ['an error that is none', cell({ k: 'error', v: '<b>boom</b>' })],
    ['a text with a character no XML file holds', cell({ k: 'string', v: 'a\u0000b' })],
    ['a text with half a surrogate pair', cell({ k: 'string', v: 'a\ud800b' })],
    ['a pattern that is none', format({ l: { type: 'pattern', pattern: 'evil"pattern<' } })],
    ['a gradient fill', format({ l: { type: 'gradient', xml: '<q:evil xmlns:q="urn:x"/>' } })],
    ['a border style that is none', format({ b: { diagonalUp: false, diagonalDown: false, left: { style: 'no-such-style' } } })],
    ['a font name with a control character', format({ f: { ...FONT, name: 'a\u0001b' } })],
    ['a number format with a control character', format({ n: '0\u0001' })],
    ['a sheet name with a control character', [{ t: 'name', s: 1, name: 'a\u0001b' }]],
    ['a note with a control character', [{ t: 'notes', s: 0, n: [[1, 1, { x: 't\u000b', a: 'au' }]] }]],
    ['an author with a control character', [{ t: 'notes', s: 0, n: [[1, 1, { x: 't', a: 'a\u0001' }]] }]],
    ['a filter value with a control character', [{ t: 'filter', s: 0, f: { range: { top: 0, left: 0, bottom: 3, right: 1 }, columns: [{ colId: 0, kind: 'values', values: ['a\u0002'] }] } }]],
    ['a filter operator that is none', [{ t: 'filter', s: 0, f: { range: { top: 0, left: 0, bottom: 3, right: 1 }, columns: [{ colId: 0, kind: 'custom', conditions: [{ operator: 'x"y', value: '1' }] }] } }]],
    ['an argument of a command with a control character', [{ t: 'cmd', s: 0, name: 'setValidation', a: [{ top: 0, left: 0, bottom: 0, right: 0 }, { type: 'whole', formula1: '1', promptTitle: 'a\u0001' }] }]],
    ['too many merged ranges at once', [{ t: 'merges', s: 0, add: Array.from({ length: 48000 }, (_x, i) => [i * 2, 0, i * 2, 1]), del: [] }]],
    ['too many notes at once', [{ t: 'notes', s: 0, n: Array.from({ length: 40000 }, (_x, i) => [i, 0, { x: '' }]) }]],
    ['too many runs of columns at once', [{ t: 'cols', s: 0, p: Array.from({ length: 16384 }, (_x, i) => [i, i, { cw: false, hid: false, ol: i % 2, col: false }]) }]],
    ['too many formats at once', [{ t: 'st', f: Array.from({ length: 3000 }, (_x, i) => ({ ...PLAIN, f: { ...FONT, name: `F${i}` } })) }, { t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }]],
  ];
  const before = [wb.styles.cellXfs.length, wb.styles.fonts.length, wb.styles.numFmts.size];
  for (const [what, step] of bad) {
    const started = Date.now();
    assert.throws(() => decodeStep(step, wb.styles), { name: 'WireError' }, `refused: ${what}`);
    assert.ok(Date.now() - started < 1000, `refused at once: ${what}`);
  }
  assert.deepEqual([wb.styles.cellXfs.length, wb.styles.fonts.length, wb.styles.numFmts.size], before, 'a step that is refused leaves no format behind');
  // A step whose first operation is fine and whose second is not adds nothing either.
  assert.throws(() => decodeStep([{ t: 'st', f: [{ ...PLAIN, f: { ...FONT, name: 'Only Here', bold: true } }] }, { t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }, { t: 'nothing', s: 0 }], wb.styles), { name: 'WireError' }, 'refused as a whole');
  assert.equal(wb.styles.fonts.length, before[1], 'the format of its first operation was not added');
});

test('a step that is no step of the workbook is skipped by everybody, also by who joins later, and the room goes on', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(3);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  const evil = intruder(net, bytes);
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 20, 0, 'waits'));
  // One that cannot be read, one that reads and cannot be applied (a merged range of one cell).
  evil.send({ t: 'op', base: 0, id: 'evil:1', ops: [{ t: 'no such operation', s: 0 }] });
  net.deliverUp();
  evil.send({ t: 'op', base: 1, id: 'evil:2', ops: [{ t: 'merges', s: 0, add: [[1, 1, 1, 1]], del: [] }] });
  net.drain(random);
  for (const member of [a, b]) {
    assert.equal(member.session.state, 'online', `${member.name} stays in the room`);
    assert.deepEqual(member.notices.map((notice) => notice.code), ['bad-step'], `${member.name} was told once`);
  }
  assert.equal(b.session.pending, 0, 'what waited was placed');
  act(a, () => setCellInput(a.wb, 0, 21, 0, 'goes on'));
  net.drain(random);
  const late = await join(net, 'late', bytes);
  net.drain(random);
  assert.equal(late.session.state, 'online', 'who replays the log stays as well');
  assertSame([a, b, late], 'after the steps that were skipped');
  assert.equal(late.session.rev, a.session.rev, 'at the same step');
});

/** Strings that a file must hold as text, whatever they look like. */
const NASTY = ['"><x y="', ']]>', '&amp;<&>', "it's", ' lead and trail ', '<!--', '\u00e7\u011f\u0131\u00f6\u015f\u00fc', '\ud83d\ude00', '=1+1', 'a\tb', 'x'.repeat(300), '', '%20', "'quoted'!A1", '_x0041_', '\u200f\u202e'];

function python(): boolean {
  try {
    execFileSync('python3', ['-c', 'import xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
    return true;
  } catch {
    return false;
  }
}

test('whatever steps arrive, the file that is saved is well formed and opens', async (t) => {
  const bytes = await startFile();
  const withPython = python();
  if (!withPython) t.diagnostic('python3 is not here: the files are read by the engine only');
  const random = seeded(77);
  const pick = <T>(list: readonly T[]): T => list[Math.floor(random() * list.length)];
  const int = (low: number, high: number): number => low + Math.floor(random() * (high - low + 1));
  const text = (): string => pick(NASTY);
  for (let run = 0; run < 25; run++) {
    const net = new Net();
    const a: Member = await join(net, 'a', bytes);
    const evil = intruder(net, bytes);
    net.drain(random);
    let rev = 0;
    for (let i = 0; i < 30; i++) {
      const s = int(0, 1);
      const row = int(0, 12);
      const col = int(0, 5);
      const ops: unknown[] = pick<() => unknown[]>([
        () => [{ t: 'cells', s, c: [[row, col, { v: { k: 'string', v: text() } }]] }],
        () => [{ t: 'cells', s, c: [[row, col, { v: { k: 'blank', v: null, f: { x: pick(['A1+1', `"${text().replace(/"/g, '""')}"&A1`, 'SUM(A:A)', 'Other!A1']), a: [['ca', '1']] } } }]] }],
        () => [{ t: 'cells', s, c: [[row, col, { v: { k: 'error', v: '#N/A' } }], [row, col + 1, { v: { k: 'date', v: '2024-01-05T10:00:00' } }], [row, col + 2, { v: { k: 'string', v: text(), r: [{ text: text(), font: { ...FONT, name: text().slice(0, 30) || 'Arial', bold: true } }] } }]] }],
        () => [{ t: 'st', f: [{ ...PLAIN, n: pick(['0.00', `"${text().replace(/"/g, '')}"0`, '@', 'dd.mm.yyyy']), f: { ...FONT, name: text().slice(0, 30) || 'Arial', color: { rgb: 'FFFF0000' } }, l: { type: 'pattern', pattern: 'solid', fg: { rgb: 'FF00FF00' } }, b: { diagonalUp: false, diagonalDown: false, left: { style: 'thin' } } }] }, { t: 'cells', s, c: [[row, col, { s: 0 }]] }],
        () => [{ t: 'name', s, name: (text().replace(/[\\/?*[\]:]/g, '').slice(0, 20) || 'x') + i }],
        () => [{ t: 'sheet+', s: 0, name: (text().replace(/[\\/?*[\]:]/g, '').slice(0, 20) || 'n') + i }],
        () => [{ t: 'notes', s, n: [[row, col, { x: text(), a: text().slice(0, 40) }]] }],
        () => [{ t: 'cmd', s, name: 'setValidation', a: [{ top: row, left: col, bottom: row + 1, right: col }, { type: 'list', formula1: `"${text().replace(/[",]/g, '').slice(0, 20) || 'a'},b"`, promptTitle: text().slice(0, 30), prompt: text(), errorTitle: text().slice(0, 30), error: text() }] }],
        () => [{ t: 'cmd', s, name: 'addConditionalRule', a: [{ top: row, left: col, bottom: row + 2, right: col }, { type: 'containsText', text: text().slice(0, 50) || 'a', format: { fill: '#FFFF00' } }] }],
        () => [{ t: 'cmd', s, name: 'addChart', a: [{ type: 'column', range: { top: 0, left: 0, bottom: 4, right: 2 }, title: text().slice(0, 100) }] }],
        () => [{ t: 'filter', s, f: { range: { top: 0, left: 0, bottom: 9, right: 3 }, columns: [{ colId: 1, kind: 'values', values: [text(), text()] }] } }],
        () => [{ t: 'tab', s, color: { rgb: 'FF112233' } }],
      ])();
      evil.send({ t: 'op', base: rev, id: `evil:${i + 1}`, ops });
      net.drain(random);
      rev = a.session.rev;
    }
    assert.equal(a.session.state, 'online', `run ${run}: the session is still there`);
    const saved = await saveOf(a);
    assert.deepEqual(await packageProblems(saved), [], `run ${run}: the package is whole`);
    const again = await (await open(saved)).workbook();
    assert.equal(again.sheets.length, a.wb.sheets.length, `run ${run}: the file opens with its sheets`);
    if (withPython) {
      const zip = ZipArchive.read(saved);
      const parts: { [name: string]: string } = {};
      for (const name of zip.names()) if (/\.(xml|rels|vml)$/.test(name)) parts[name] = Buffer.from(await zip.read(name)).toString('base64');
      const out = execFileSync('python3', ['-c', 'import sys,json,base64\nimport xml.etree.ElementTree as E\nbad=[]\nfor k,v in json.load(sys.stdin).items():\n  try:\n    E.fromstring(base64.b64decode(v))\n  except Exception as e:\n    bad.append(k+": "+str(e))\nprint(json.dumps(bad))'], { input: JSON.stringify(parts), encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, timeout: 60000 });
      assert.deepEqual(JSON.parse(out), [], `run ${run}: every part is well formed, namespaces and all`);
    }
  }
});

test('more of what is no step: lengths, depths and places', async () => {
  const { wb } = await join(new Net(), 'x', await startFile());
  let deep: unknown = 1;
  for (let i = 0; i < 20; i++) deep = [deep];
  const bad: [string, unknown[]][] = [
    ['a formula that is too long', [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'blank', v: null, f: { x: 'A1+'.repeat(3000) } } }]] }]],
    ['a sheet beyond any workbook', [{ t: 'tab', s: 5000, color: 0 }]],
    ['an argument that is nested too deep', [{ t: 'cmd', s: 0, name: 'addChart', a: [deep] }]],
    ['an argument of too many parts', [{ t: 'cmd', s: 0, name: 'addChart', a: [Array.from({ length: 100001 }, () => 0)] }]],
    ['a format that no table of the step has', [{ t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }]],
    ['too many peers in a hello', []],
  ];
  for (const [what, step] of bad.slice(0, 5)) assert.throws(() => decodeStep(step, wb.styles), { name: 'WireError' }, `refused: ${what}`);
  const { parseServerMessage } = await import('../src/collab/protocol');
  assert.throws(() => parseServerMessage(JSON.stringify({ t: 'hello', rev: 0, from: 0, ep: 'e', you: 'c', ro: false, v: [], peers: Array.from({ length: 1025 }, (_x, i) => ({ id: `p${i}`, user: {}, ro: false })) })), { name: 'WireError' }, 'refused: a hello with more peers than a room has');
});
