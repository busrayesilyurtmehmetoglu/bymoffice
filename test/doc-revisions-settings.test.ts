// Turning the tracking of changes on and off: the setting is an edit that can
// be undone, and the settings part of the file changes by <w:trackRevisions>
// and nothing else.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { tracksChanges, type DocModel } from '../src/model/doc/types';
import { DocHistory, TRACK_ENFORCED, insertText, isTracking, revisionsOf, setTrackRevisions, setTracking, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { STYLES, p, pos, run, texts } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
const ME = { author: 'Me', date: '2026-10-08T10:00:00Z' };

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: p('abc') + SECT, styles: STYLES, ...extra });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

async function settingsOf(o: Opened): Promise<string | undefined> {
  await writeDocx(o.pkg, o.doc);
  if (!o.pkg.hasPart('/word/settings.xml')) return undefined;
  const text = decoder.decode(await o.pkg.readPart('/word/settings.xml'));
  parseXml(text);
  assert.deepEqual(schemaOrderViolations(text), []);
  return text;
}

const inner = (text: string | undefined): string => {
  assert.ok(text !== undefined);
  return text.slice(text.indexOf('>', text.indexOf('<w:settings')) + 1, text.lastIndexOf('</w:settings>'));
};

const refused = (make: () => unknown): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_ENFORCED);
};

test('tracking: turning it on puts <w:trackRevisions/> at its place in the settings and changes nothing else', async () => {
  const before = '<w:zoom w:percent="120"/><w:mirrorMargins/><w:revisionView w:markup="1"/>';
  const after = '<w:defaultTabStop w:val="708"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w14:docId w14:val="1234ABCD"/>';
  const o = await open({ settings: before + after });
  assert.equal(tracksChanges(o.doc), false);
  const edit = setTrackRevisions(o.doc, true);
  assert.equal(edit.ops.length, 1);
  run(o.history, edit);
  assert.equal(tracksChanges(o.doc), true);
  assert.equal(inner(await settingsOf(o)), before + '<w:trackRevisions/>' + after);
  // Nothing but the settings was written.
  assert.equal(decoder.decode(await o.pkg.readPart('/word/document.xml')).indexOf('<w:ins'), -1);

  // Off again by undoing: the file is what it was, byte for byte.
  o.history.undo();
  assert.equal(tracksChanges(o.doc), false);
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
  // And by turning it off.
  o.history.redo();
  run(o.history, setTrackRevisions(o.doc, false));
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
  // Asked for what is already so, there is nothing to do.
  assert.equal(setTrackRevisions(o.doc, false).ops.length, 0);
});

test('tracking: turning it off takes the element out, however the file wrote it, and leaves the rest', async () => {
  for (const [element, kept] of [
    ['<w:trackRevisions/>', ''],
    ['<w:trackRevisions w:val="true"/>', ''],
    ['<w:trackRevisions w:val="0"/><w:trackRevisions w:val="1"></w:trackRevisions>', ''],
  ]) {
    const o = await open({ settings: `<w:zoom w:percent="90"/>${element}<w:defaultTabStop w:val="708"/>` });
    assert.equal(tracksChanges(o.doc), true, element);
    run(o.history, setTrackRevisions(o.doc, false));
    assert.equal(inner(await settingsOf(o)), `<w:zoom w:percent="90"/>${kept}<w:defaultTabStop w:val="708"/>`, element);
    const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
    assert.equal(tracksChanges(again), false);
    o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, element);
  }
});

test('tracking: a setting the file states as off is turned on in its place', async () => {
  const o = await open({ settings: '<w:zoom w:percent="90"/><w:trackRevisions w:val="false"/><w:defaultTabStop w:val="708"/>' });
  assert.equal(tracksChanges(o.doc), false);
  run(o.history, setTrackRevisions(o.doc, true));
  assert.equal(inner(await settingsOf(o)), '<w:zoom w:percent="90"/><w:trackRevisions/><w:defaultTabStop w:val="708"/>');
});

test('tracking: a settings part that is empty, self-closing or missing', async () => {
  let o = await open({ settings: '' });
  run(o.history, setTrackRevisions(o.doc, true));
  assert.equal(inner(await settingsOf(o)), '<w:trackRevisions/>');

  o = await open({ parts: [['word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>']], rels: [['rId90', 'settings', 'settings.xml']] });
  run(o.history, setTrackRevisions(o.doc, true));
  assert.equal(inner(await settingsOf(o)), '<w:trackRevisions/>');

  // No settings part at all: one is made, with its relationship and content type, and taken out again when it is not needed.
  o = await open();
  assert.equal(o.pkg.hasPart('/word/settings.xml'), false);
  run(o.history, setTrackRevisions(o.doc, true));
  assert.equal(inner(await settingsOf(o)), '<w:trackRevisions/>');
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.equal(tracksChanges(again), true);
  assert.deepEqual(texts(again), ['abc']);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.equal(o.pkg.hasPart('/word/settings.xml'), false);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('tracking: on, typing is tracked; off, it is not; saved and opened again, both are as they were typed', async () => {
  const o = await open({ settings: '<w:zoom w:percent="100"/>' });
  setTracking(o.doc, { author: 'Me', now: () => new Date(Date.UTC(2026, 9, 8, 10, 0, 0)) });
  run(o.history, setTrackRevisions(o.doc, true));
  assert.equal(isTracking(o.doc), true);
  run(o.history, insertText(o.doc, pos(0, 3), 'X', undefined, ME));
  run(o.history, setTrackRevisions(o.doc, false));
  assert.equal(isTracking(o.doc), false);
  run(o.history, insertText(o.doc, pos(0, 4), 'Y'));
  await writeDocx(o.pkg, o.doc);
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.equal(tracksChanges(again), false);
  assert.deepEqual(texts(again), ['abcXY']);
  assert.deepEqual(
    revisionsOf(again).map((revision) => [revision.kind, revision.author, revision.text]),
    [['insert', 'Me', 'X']],
  );
  // The whole way back.
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('tracking: it cannot be turned off where a protection enforces it, nor where the application asked for it', async () => {
  const locked = await open({ settings: '<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>' });
  refused(() => setTrackRevisions(locked.doc, false));
  // Turning on what is on anyway is allowed, and writes the setting.
  assert.equal(setTrackRevisions(locked.doc, true).ops.length, 1);

  const forced = await open({ settings: '<w:trackRevisions/>' });
  setTracking(forced.doc, { force: true });
  refused(() => setTrackRevisions(forced.doc, false));
  const plain = await open();
  setTracking(plain.doc, { force: true });
  refused(() => setTrackRevisions(plain.doc, false));

  // A protection that is not enforced, and one of another kind, restrict nothing here.
  const free = await open({ settings: '<w:trackRevisions/><w:documentProtection w:edit="trackedChanges" w:enforcement="0"/>' });
  assert.equal(setTrackRevisions(free.doc, false).ops.length, 1);
  assert.throws(() => setTrackRevisions(free.doc, 'yes' as never), RangeError);
});

test('review B7: a settings part with a byte order mark and CRLF line ends changes by the element alone', async () => {
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const before = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings ${NS}>\r\n  <w:zoom w:percent="90"/>\r\n  `;
  const after = `<w:defaultTabStop w:val="708"/>\r\n  <w:compat/>\r\n</w:settings>\r\n`;
  const utf8 = (text: string): Uint8Array => new TextEncoder().encode('﻿' + text);
  const utf16 = (text: string): Uint8Array => {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = 0xff;
    out[1] = 0xfe;
    for (let i = 0; i < text.length; i++) {
      out[2 + i * 2] = text.charCodeAt(i) & 0xff;
      out[3 + i * 2] = text.charCodeAt(i) >> 8;
    }
    return out;
  };
  for (const [name, encode, declared] of [
    ['utf-8', utf8, 'UTF-8'],
    ['utf-16', utf16, 'UTF-16'],
  ] as const) {
    const head = before.replace('UTF-8', declared);
    // Turning it on: the element goes in after <w:zoom>, where the schema has it.
    let o = await open({ parts: [['word/settings.xml', encode(head + after)]], rels: [['rId90', 'settings', 'settings.xml']] });
    run(o.history, setTrackRevisions(o.doc, true));
    await writeDocx(o.pkg, o.doc);
    const place = head.indexOf('<w:zoom w:percent="90"/>') + '<w:zoom w:percent="90"/>'.length;
    assert.deepEqual(await o.pkg.readPart('/word/settings.xml'), encode(head.slice(0, place) + '<w:trackRevisions/>' + head.slice(place) + after), name);
    o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, name);

    // Turning it off: the element goes, and nothing else.
    o = await open({ parts: [['word/settings.xml', encode(head + '<w:trackRevisions/>\r\n  ' + after)]], rels: [['rId90', 'settings', 'settings.xml']] });
    assert.equal(tracksChanges(o.doc), true, name);
    run(o.history, setTrackRevisions(o.doc, false));
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.readPart('/word/settings.xml'), encode(head + '\r\n  ' + after), name);
  }
});

test('review: where the application asked for tracking, turning it on changes nothing in the file', async () => {
  const o = await open({ settings: '<w:zoom w:percent="100"/>' });
  setTracking(o.doc, { force: true });
  assert.equal(isTracking(o.doc), true);
  // Tracking is on already, by what the application said: the setting of the document is not the user's to write.
  assert.equal(setTrackRevisions(o.doc, true).ops.length, 0);
  assert.equal(o.doc.settings.trackRevisions, false);
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});
