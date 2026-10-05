// Document protection (<w:documentProtection>): read into the settings, left
// in the file byte for byte, carried into a new docx, and read from doc and
// RTF where those state it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type Bridge, type BridgeHost } from '../src/embed/bridge';
import { convertForEditing } from '../src/formats/convert';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { readRtf } from '../src/formats/rtf/document';
import { open as openFile } from '../src/index';
import type { DocumentViewOptions } from '../src/ui/doc-view';
import { enforcedProtection, hasPermissionRanges, tracksChanges } from '../src/model/doc/types';
import { insertText } from '../src/ops/doc/index';
import { applyDocOps } from '../src/ops/doc/ops';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { FC, simpleDoc } from './doc-binary-fixtures';
import { buildDocx as buildSpec, readSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const HASH = 'w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="q83vEjRWeJA=" w:salt="AAECAwQFBgc="';
const BODY = '<w:p><w:r><w:t>text</w:t></w:r></w:p>';

const protect = (edit: string, enforcement: string | undefined, more = '') =>
  `<w:documentProtection w:edit="${edit}"${enforcement === undefined ? '' : ` w:enforcement="${enforcement}"`}${more ? ' ' + more : ''}/>`;

test('every kind of protection is read, with whether it is enforced', async () => {
  for (const edit of ['readOnly', 'comments', 'trackedChanges', 'forms', 'none']) {
    for (const [value, on] of [['1', true], ['true', true], ['on', true], ['0', false], ['false', false], [undefined, false]] as const) {
      const doc = await readSpec({ body: BODY, settings: protect(edit, value) });
      const protection = doc.settings.protection;
      assert.ok(protection, `${edit} ${value}`);
      assert.equal(protection.edit, edit);
      assert.equal(protection.enforced, on, `${edit} ${value}`);
      const expected = on && edit !== 'none' ? edit : undefined;
      assert.equal(enforcedProtection(doc), expected, `${edit} ${value}`);
      assert.equal(tracksChanges(doc), on && edit === 'trackedChanges');
    }
  }
});

test('a document without protection has none', async () => {
  const doc = await readSpec({ body: BODY, settings: '<w:trackRevisions/>' });
  assert.equal(doc.settings.protection, undefined);
  assert.equal(enforcedProtection(doc), undefined);
  assert.equal(tracksChanges(doc), true);
  assert.equal(tracksChanges(await readSpec({ body: BODY })), false);
});

test('an enforced protection of a kind that is not known counts as read-only', async () => {
  const doc = await readSpec({ body: BODY, settings: protect('somethingNew', '1') });
  assert.equal(doc.settings.protection?.edit, 'somethingNew');
  assert.equal(enforcedProtection(doc), 'readOnly');
});

test('an enforced protection that restricts formatting only (no w:edit) locks nothing', async () => {
  const doc = await readSpec({ body: BODY, settings: '<w:documentProtection w:formatting="1" w:enforcement="1" w:cryptProviderType="rsaAES" w:hash="q83vEjRWeJA="/>' });
  assert.equal(doc.settings.protection?.edit, '');
  assert.equal(doc.settings.protection?.enforced, true);
  assert.equal(enforcedProtection(doc), undefined);
  assert.equal(tracksChanges(doc), false);
});

test('the element is kept as read, the password hash included', async () => {
  const doc = await readSpec({ body: BODY, settings: protect('readOnly', '1', HASH) });
  const xml = doc.settings.protection?.xml;
  assert.ok(xml);
  assert.equal(xml.attrs.get('w:hash'), 'q83vEjRWeJA=');
  assert.equal(xml.attrs.get('w:salt'), 'AAECAwQFBgc=');
  assert.equal(xml.attrs.get('w:cryptSpinCount'), '100000');
});

test('the write protection (a recommendation) is kept and never applied', async () => {
  const doc = await readSpec({ body: BODY, settings: '<w:writeProtection w:recommended="1"/>' });
  assert.ok(doc.settings.writeProtection);
  assert.equal(doc.settings.writeProtection.attrs.get('w:recommended'), '1');
  assert.equal(enforcedProtection(doc), undefined);
});

test('saving an edited protected document leaves the settings part byte for byte', async () => {
  const settings = `<w:zoom w:percent="100"/><w:trackRevisions/>${protect('trackedChanges', '1', HASH)}<w:defaultTabStop w:val="708"/>`;
  const bytes = await buildSpec({ body: BODY, settings });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  applyDocOps(doc, insertText(doc, { story: 'body', path: [0], offset: 4 }, '!', undefined, { author: 'A', date: '2026-01-01T00:00:00Z' }).ops);
  await writeDocx(pkg, doc);
  const saved = await pkg.save();
  const before = ZipArchive.read(bytes).entry('word/settings.xml');
  const after = ZipArchive.read(saved).entry('word/settings.xml');
  assert.ok(before && after);
  assert.equal(after.data, undefined);
  assert.deepEqual(after.raw, before.raw);
  const again = await readDocx(await OpcPackage.open(saved));
  assert.equal(enforcedProtection(again), 'trackedChanges');
  assert.equal(again.settings.protection?.xml?.attrs.get('w:hash'), 'q83vEjRWeJA=');
});

test('a new docx carries the protection over, in the place the schema gives it', async () => {
  const source = await readSpec({ body: BODY, settings: `<w:trackRevisions/>${protect('forms', '1', HASH)}<w:defaultTabStop w:val="708"/>` });
  const report = emptyWriteReport();
  const pkg = await OpcPackage.open(await buildDocx(source, { report }));
  const text = decoder.decode(await pkg.readPart('/word/settings.xml'));
  const at = text.indexOf('<w:documentProtection ');
  assert.ok(at > text.indexOf('<w:trackRevisions/>') && at < text.indexOf('<w:defaultTabStop'), text);
  assert.ok(text.indexOf('w:hash="q83vEjRWeJA="') > 0, text);
  const again = await readDocx(pkg);
  assert.equal(enforcedProtection(again), 'forms');
  // The element was kept whole: nothing to warn about.
  assert.deepEqual(report.warnings.filter((line) => /protection/i.test(line)), []);
});

test('a protection without its element (doc, RTF) is written without a password, and the report says so', async () => {
  const source = await readSpec({ body: BODY });
  source.settings.protection = { edit: 'readOnly', enforced: true };
  const report = emptyWriteReport();
  const pkg = await OpcPackage.open(await buildDocx(source, { report }));
  const text = decoder.decode(await pkg.readPart('/word/settings.xml'));
  assert.ok(text.indexOf('<w:documentProtection w:edit="readOnly" w:enforcement="1"/>') > 0, text);
  assert.equal(report.warnings.filter((line) => /protection/i.test(line)).length, 1, report.warnings.join('\n'));
  assert.equal(enforcedProtection(await readDocx(pkg)), 'readOnly');
});

test('ranges that are exempt from the protection are noticed', async () => {
  const exempt = await readSpec({ body: '<w:p><w:permStart w:id="1" w:edGrp="everyone"/><w:r><w:t>free</w:t></w:r><w:permEnd w:id="1"/></w:p>', settings: protect('readOnly', '1') });
  assert.equal(hasPermissionRanges(exempt), true);
  assert.equal(hasPermissionRanges(await readSpec({ body: BODY, settings: protect('readOnly', '1') })), false);
});

const rtf = (head: string) => new TextEncoder().encode(`{\\rtf1\\ansi${head}\\pard text\\par}`);

test('RTF: a protection applies when the file enforces it, or names it the old way without saying whether it is enforced', () => {
  const of = (head: string) => readRtf(rtf(head)).settings.protection;
  assert.equal(of(''), undefined);
  // The old words alone.
  assert.deepEqual(of('\\formprot'), { edit: 'forms', enforced: true });
  assert.deepEqual(of('\\revprot'), { edit: 'trackedChanges', enforced: true });
  assert.deepEqual(of('\\annotprot'), { edit: 'comments', enforced: true });
  assert.deepEqual(of('\\readprot'), { edit: 'readOnly', enforced: true });
  // Enforced, with the level that says for what.
  assert.deepEqual(of('\\enforceprot1\\protlevel0'), { edit: 'trackedChanges', enforced: true });
  assert.deepEqual(of('\\enforceprot1\\protlevel1'), { edit: 'comments', enforced: true });
  assert.deepEqual(of('\\enforceprot1\\protlevel2'), { edit: 'forms', enforced: true });
  assert.deepEqual(of('\\enforceprot1\\protlevel3'), { edit: 'readOnly', enforced: true });
  assert.deepEqual(of('\\readprot\\enforceprot1\\protlevel3'), { edit: 'readOnly', enforced: true });
  // The level wins over an old word beside it.
  assert.deepEqual(of('\\formprot\\enforceprot1\\protlevel1'), { edit: 'comments', enforced: true });
  // Said not to be enforced: kept, not applied.
  assert.deepEqual(of('\\formprot\\enforceprot0\\protlevel2'), { edit: 'forms', enforced: false });
  assert.deepEqual(of('\\enforceprot0\\protlevel3'), { edit: 'readOnly', enforced: false });
  // A level alone is no protection, nor is an enforcement of nothing, nor a level out of range.
  for (const head of ['\\protlevel3', '\\protlevel0', '\\enforceprot1', '\\enforceprot1\\protlevel7', '\\protlevel']) {
    assert.equal(enforcedProtection(readRtf(rtf(head))), undefined, head);
  }
  assert.equal(of('\\protlevel3'), undefined);
});

/** A document properties structure (Dop) of 84 bytes with the given bytes set. */
function dop(bytes: [offset: number, value: number][]): Uint8Array {
  const out = new Uint8Array(84);
  for (const [offset, value] of bytes) out[offset] = value;
  return out;
}

test('doc: the protection bits of the document properties lock only together with the bit that enables protection', () => {
  const of = (bytes: [number, number][]) => simpleDoc('text\r').table(FC.Dop, dop(bytes)).read().settings.protection;
  assert.equal(of([]), undefined);
  // fProtEnabled (byte 7, 0x02) alone: forms.
  assert.deepEqual(of([[7, 0x02]]), { edit: 'forms', enforced: true });
  // With fLockAtn (byte 6, 0x10): comments; with fLockRev (byte 7, 0x40): tracked changes.
  assert.deepEqual(of([[6, 0x10], [7, 0x02]]), { edit: 'comments', enforced: true });
  assert.deepEqual(of([[7, 0x42]]), { edit: 'trackedChanges', enforced: true });
  // The lock bits without it: kept, not applied.
  assert.deepEqual(of([[6, 0x10]]), { edit: 'comments', enforced: false });
  assert.deepEqual(of([[7, 0x40]]), { edit: 'trackedChanges', enforced: false });
  // Neighbouring bits are not protection: fMirrorMargins (byte 6, 0x20), fEmbedFonts (byte 7, 0x80), fRevMarking (byte 5, 0x80).
  const other = simpleDoc('text\r').table(FC.Dop, dop([[5, 0x80], [6, 0x20], [7, 0x80]])).read().settings;
  assert.equal(other.protection, undefined);
  assert.equal(other.trackRevisions, true);
  assert.equal(other.mirrorMargins, true);
});

test('the write protection is carried into a new docx, first among the settings', async () => {
  const source = await readSpec({ body: BODY, settings: '<w:writeProtection w:recommended="1" w:hash="q83vEjRWeJA="/><w:trackRevisions/>' });
  const pkg = await OpcPackage.open(await buildDocx(source));
  const text = decoder.decode(await pkg.readPart('/word/settings.xml'));
  const at = text.indexOf('<w:writeProtection w:recommended="1" w:hash="q83vEjRWeJA="/>');
  assert.ok(at > 0 && at < text.indexOf('<w:trackRevisions/>'), text);
  assert.ok((await readDocx(pkg)).settings.writeProtection);
});

// ------------------------------------------------------------------ the bridge

interface Shown {
  sent: Record<string, unknown>[];
  options: DocumentViewOptions[];
  readOnlyCalls: boolean[];
  bridge: Bridge;
}

/** A bridge over the real open(), with a view that only records what it is given. */
function bridged(): Shown {
  const state: Shown = { sent: [], options: [], readOnlyCalls: [], bridge: undefined as unknown as Bridge };
  const host: BridgeHost = {
    container: {} as HTMLElement,
    open: openFile,
    createView: () => {
      throw new Error('not a workbook');
    },
    createDocumentView: (_container, _doc, options) => {
      state.options.push(options);
      return { destroy: () => undefined, setReadOnly: (value: boolean) => state.readOnlyCalls.push(value) };
    },
    post: (json) => {
      state.sent.push(JSON.parse(json));
      return true;
    },
    version: 'test',
  };
  state.bridge = createBridge(host);
  return state;
}

const loadedOf = (shown: Shown): Record<string, unknown> => shown.sent.filter((message) => message.type === 'loaded' || message.type === 'error').slice(-1)[0];

test('bridge: a document protected against changes is loaded read-only, and loaded says why', async () => {
  for (const edit of ['readOnly', 'comments', 'forms']) {
    const shown = bridged();
    const data = base64Encode(await buildSpec({ body: BODY, settings: protect(edit, '1') }));
    await shown.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data, readOnly: false }));
    assert.deepEqual(loadedOf(shown), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, readOnly: true, protection: edit });
    assert.equal(shown.options[0].readOnly, true);
    // Nothing turns editing on: the answer says what holds instead.
    await shown.bridge.receive(JSON.stringify({ type: 'setReadOnly', id: 'b', readOnly: false }));
    assert.deepEqual(shown.readOnlyCalls, [true]);
    assert.deepEqual(shown.sent[shown.sent.length - 1], { type: 'done', id: 'b', command: 'setReadOnly', readOnly: true });
  }
});

test('bridge: exempt ranges are reported and not honoured', async () => {
  const shown = bridged();
  const body = '<w:p><w:permStart w:id="1" w:edGrp="everyone"/><w:r><w:t>free</w:t></w:r><w:permEnd w:id="1"/></w:p>';
  await shown.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: base64Encode(await buildSpec({ body, settings: protect('readOnly', '1') })) }));
  assert.deepEqual(loadedOf(shown), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, readOnly: true, protection: 'readOnly', protectionExceptions: true });
});

test('bridge: tracked changes are said in loaded; the author goes to the view', async () => {
  const forced = bridged();
  await forced.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: base64Encode(await buildSpec({ body: BODY, settings: protect('trackedChanges', '1') })), author: 'Ada' }));
  assert.deepEqual(loadedOf(forced), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, protection: 'trackedChanges', trackChanges: true });
  assert.equal(forced.options[0].readOnly, false);
  assert.equal(forced.options[0].author, 'Ada');

  const setting = bridged();
  await setting.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: base64Encode(await buildSpec({ body: BODY, settings: '<w:trackRevisions/>' })) }));
  assert.deepEqual(loadedOf(setting), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, trackChanges: true });

  const asked = bridged();
  await asked.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: base64Encode(await buildSpec({ body: BODY })), trackChanges: true }));
  assert.deepEqual(loadedOf(asked), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, trackChanges: true });
  assert.equal(asked.options[0].trackChanges, true);

  const plain = bridged();
  await plain.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: base64Encode(await buildSpec({ body: BODY, settings: protect('readOnly', '0') })) }));
  assert.deepEqual(loadedOf(plain), { type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.equal(plain.options[0].trackChanges, undefined);
});

test('bridge: an author that is not text, or too long, is a bad message', async () => {
  const shown = bridged();
  const data = base64Encode(await buildSpec({ body: BODY }));
  await shown.bridge.receive(JSON.stringify({ type: 'load', id: 'a', data, author: 7 }));
  assert.equal(loadedOf(shown).code, 'bad-message');
  await shown.bridge.receive(JSON.stringify({ type: 'load', id: 'b', data, author: 'x'.repeat(256) }));
  assert.equal(loadedOf(shown).code, 'bad-message');
  assert.equal(shown.options.length, 0);
});

test('a protected RTF file keeps its protection when it is converted, and the report says the password is not carried', async () => {
  const file = await openFile(rtf('\\readprot\\enforceprot1\\protlevel3'));
  const converted = await convertForEditing(file);
  assert.ok(converted);
  assert.equal(converted.warnings.filter((line) => /protection/i.test(line)).length, 1);
  const again = await (await openFile(converted.bytes)).document!();
  assert.equal(enforcedProtection(again), 'readOnly');
});
