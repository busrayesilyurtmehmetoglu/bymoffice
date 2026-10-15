// The version of the contract of editing together goes up with the wire: a test that fails
// when the kinds of steps, or the fields of one, change while PROTOCOL_VERSION stays.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { PROTOCOL_VERSION } from '../src/collab/room';
import { decodeStep as decodeSheetStep } from '../src/collab/wire';
import { decodeStep as decodeDocStep } from '../src/collab/doc/wire';
import { contentVersion } from '../src/collab/protocol';
import { CollabSession } from '../src/collab/session';
import { open } from '../src/index';
import { Net, seeded, startFile } from './collab-fixtures';
import { versionProblem, wireFingerprint } from './collab-wire-fingerprint';

/**
 * The wire as version `pv` has it (see collab-wire-fingerprint.ts for how it is taken). When the
 * test below fails: raise PROTOCOL_VERSION, and put the lines it prints here with the new version.
 */
const PINNED = {
  pv: 4,
  lines: [
    'doc commands: deleteBackward deleteColumn deleteForward deleteRow indentList insertColumn insertHyperlink insertPageBreak insertRow insertTable toggleList',
    'doc kinds: cmd del para put run split text unrun',
    'doc:cmd: a at name t to',
    'doc:del: at t to',
    'doc:para: at f t to',
    'doc:para~f.indent: firstLine firstLineChars hanging hangingChars left leftChars right rightChars',
    'doc:para~f.numbering: level numId',
    'doc:para~f.spacing: after afterAuto afterLines before beforeAuto beforeLines line lineRule',
    'doc:para~f.tabs.set[]: align leader pos',
    'doc:para~f.tabs: clear clearAll set',
    'doc:para~f: align contextualSpacing indent keepLines keepNext numbering outlineLevel pageBreakBefore spacing style tabs widowControl',
    'doc:put.x[].r[]: a b c e p',
    'doc:put.x[]: a m p r',
    'doc:put: at h n t x',
    'doc:put~x[].m.cs: bold italic size',
    'doc:put~x[].m.f: bold color font highlight italic size strike style underline vertAlign',
    'doc:put~x[].m.fonts: ascii cs eastAsia hAnsi hint',
    'doc:put~x[].m.lang: bidi eastAsia val',
    'doc:put~x[].m: cs f fonts lang',
    'doc:put~x[].p.indent: firstLine firstLineChars hanging hangingChars left leftChars right rightChars',
    'doc:put~x[].p.spacing: after afterAuto afterLines before beforeAuto beforeLines line lineRule',
    'doc:put~x[].p.tabs[]: align leader pos',
    'doc:put~x[].p: align contextualSpacing indent keepLines keepNext numbering outlineLevel pageBreakBefore spacing style tabs widowControl',
    'doc:put~x[].r[].p.cs: bold italic size',
    'doc:put~x[].r[].p.f: bold color font highlight italic size strike style underline vertAlign',
    'doc:put~x[].r[].p.fonts: ascii cs eastAsia hAnsi hint',
    'doc:put~x[].r[].p.lang: bidi eastAsia val',
    'doc:put~x[].r[].p: cs f fonts lang',
    'doc:put~x[].r[]: a b c e p',
    'doc:put~x[]: a m p r',
    'doc:run: at f t to',
    'doc:run~f: bold color font highlight italic size strike style underline vertAlign',
    'doc:split: at t',
    'doc:text: at f t to x',
    'doc:text~f: bold color font highlight italic size strike style underline vertAlign',
    'doc:unrun: at k t to',
    'sheet commands: addChart addConditionalRule clearConditionalRules moveChart moveConditionalRule pasteChart removeChart removeConditionalRule setValidation updateChart updateConditionalRule',
    'sheet kinds: cells cmd cols del filter ins merges move name notes pane rows sheet+ sheet- sort state tab',
    'sheet:cells.c[][].v.f: a d r t x',
    'sheet:cells.c[][].v.r[].font.color: auto indexed rgb theme tint',
    'sheet:cells.c[][].v.r[].font: bold charset color family italic name scheme size strike underline vertAlign',
    'sheet:cells.c[][].v.r[]: font text',
    'sheet:cells.c[][].v: f k ph r v',
    'sheet:cells.c[][]: s v',
    'sheet:cells: c s t',
    'sheet:cmd: a name s t',
    'sheet:cols.p[][]: col cw hid ol s w',
    'sheet:cols: p s t',
    'sheet:del: at n s t x',
    'sheet:filter.f.columns[].conditions[]: operator value',
    'sheet:filter.f.columns[].dateGroups[]: day grouping hour minute month second year',
    'sheet:filter.f.columns[]: and blank colId conditions dateGroups dynamicType hiddenButton kind percent top value values',
    'sheet:filter.f.range: bottom left right top',
    'sheet:filter.f.sortState.conditions[].range: bottom left right top',
    'sheet:filter.f.sortState.conditions[]: descending range',
    'sheet:filter.f.sortState.range: bottom left right top',
    'sheet:filter.f.sortState: caseSensitive columnSort conditions range',
    'sheet:filter.f: columns held range sortState',
    'sheet:filter: f hide s show t',
    'sheet:ins: at n s t x',
    'sheet:merges: add del s t',
    'sheet:move: s t to',
    'sheet:name: name s t',
    'sheet:notes.n[][].r[].font: bold charset color family italic name scheme size strike underline vertAlign',
    'sheet:notes.n[][].r[]: font text',
    'sheet:notes.n[][]: a r vis x',
    'sheet:notes: n s t',
    'sheet:pane: c fv r s t tl',
    'sheet:rows.r[][]: ch h hid s',
    'sheet:rows: r s t',
    'sheet:sheet+: name s t',
    'sheet:sheet-: s t',
    'sheet:sort: g o s t',
    'sheet:st.f[].a: horizontal indent readingOrder shrinkToFit textRotation vertical wrapText',
    'sheet:st.f[].b.left.color: auto indexed rgb theme tint',
    'sheet:st.f[].b.left: color style',
    'sheet:st.f[].b: bottom diagonal diagonalDown diagonalUp left right top',
    'sheet:st.f[].f.color: auto indexed rgb theme tint',
    'sheet:st.f[].f: bold charset color family italic name scheme size strike underline vertAlign',
    'sheet:st.f[].l.bg: auto indexed rgb theme tint',
    'sheet:st.f[].l.fg: auto indexed rgb theme tint',
    'sheet:st.f[].l: bg fg pattern type',
    'sheet:st.f[].p: hidden locked',
    'sheet:st.f[]: a b f l n p q x',
    'sheet:st: f t',
    'sheet:state: s state t',
    'sheet:tab.color: auto indexed rgb theme tint',
    'sheet:tab: color s t',
  ],
};

test('the wire of editing together is the one pinned for this PROTOCOL_VERSION: pv is raised whenever a kind of step or a field of one is added', async () => {
  const lines = await wireFingerprint();
  const problem = versionProblem(lines, PROTOCOL_VERSION, PINNED);
  assert.equal(problem, undefined, problem);
});

test('the fingerprint sees a field that is added to a step, and the check then asks for a new version', async () => {
  // A decoder of workbooks that reads one field more of a pane, and one of documents that reads one more of a paragraph format.
  const sheet = (step: unknown[], styles: Parameters<typeof decodeSheetStep>[1]) => {
    for (const op of step as { t?: unknown; zz?: unknown }[]) if (op && op.t === 'pane') void op.zz;
    return decodeSheetStep(step, styles);
  };
  const doc = (step: unknown[]) => {
    for (const wire of step as { t?: unknown; yy?: unknown }[]) if (wire && wire.t === 'split') void ('yy' in wire);
    return decodeDocStep(step);
  };
  const now = await wireFingerprint();
  const more = await wireFingerprint({ sheet, doc });
  assert.ok(more.includes('sheet:pane: c fv r s t tl zz'), 'the field read of a pane is seen');
  assert.ok(more.includes('doc:split: at t yy'), 'a field asked for with "in" is seen');
  const problem = versionProblem(more, PINNED.pv, { pv: PINNED.pv, lines: now });
  assert.ok(problem && /raise PROTOCOL_VERSION/.test(problem) && /\+ sheet:pane: c fv r s t tl zz/.test(problem) && /- sheet:pane: c fv r s t tl$/m.test(problem), problem);
  // Raised, but not pinned: the fingerprint is to be updated; pinned, but not raised: the version is to be named.
  assert.match(versionProblem(more, PINNED.pv + 1, { pv: PINNED.pv, lines: now }) || '', /update the fingerprint/);
  assert.match(versionProblem(now, PINNED.pv + 1, { pv: PINNED.pv, lines: now }) || '', /set PINNED.pv/);
  assert.equal(versionProblem(more, PINNED.pv + 1, { pv: PINNED.pv + 1, lines: more }), undefined);
});

test('the fingerprint does not depend on the order of anything: taken twice it is the same', async () => {
  assert.deepEqual(await wireFingerprint(), await wireFingerprint());
});

test('the relay of the page turns a session of this version away from a room of another one, and the session says "version", not "refused"', async () => {
  const net = new Net();
  const random = seeded(7);
  const bytes = await startFile();
  // Someone of an older page holds the room.
  const older = net.server.connect();
  const said: { t: string }[] = [];
  older.onMessage((text) => said.push(JSON.parse(text)));
  older.onState!((state) => {
    if (state === 'open') older.send(JSON.stringify({ t: 'join', pv: PROTOCOL_VERSION - 1, key: 'older', v: [contentVersion(bytes)] }));
  });
  net.drain(random);
  assert.equal(said[0].t, 'hello');
  const wb = await (await open(bytes)).workbook();
  const notices: string[] = [];
  const states: string[] = [];
  const session = new CollabSession(
    { workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice.code) },
    { transport: net.server.connect(), versions: [contentVersion(bytes)], timer: () => () => undefined, onState: (state) => states.push(state) },
  );
  net.drain(random);
  assert.deepEqual([session.state, session.reason, notices], ['closed', 'version', ['version']]);
  assert.equal(states[states.length - 1], 'closed');
});
