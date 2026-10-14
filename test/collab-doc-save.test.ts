// The file of a document that is edited together: the mark of the room in it, whether two
// participants write the same file (what proves that a save of someone else holds one's own
// work), and the bridge of the WebView page.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { CollabDocHistory, holdsRevisions } from '../src/collab/doc/history';
import type { DocChange, DocCollabSession } from '../src/collab/doc/session';
import { contentVersion } from '../src/collab/protocol';
import { MemoryServer } from '../src/collab/server';
import { createBridge, type BridgeDocumentView, type BridgeHost } from '../src/embed/bridge';
import { readCollabStamp, samePackage, writeSharedFile } from '../src/formats/xlsx/collab-stamp';
import { ZipArchive, base64Decode, base64Encode, open } from '../src/index';
import type { DocModel, Paragraph } from '../src/model/doc/types';
import type { DocWire } from '../src/collab/doc/rebase';
import { traced, tracing, wiresOf } from '../src/collab/doc/trace';
import * as wireModule from '../src/collab/doc/wire';
import { decodeStep, runWire } from '../src/collab/doc/wire';
import * as raw from '../src/ops/doc/index';
import { applyDocOps } from '../src/ops/doc/ops';
import { aloneFile, make } from './collab-doc-changes';
import { buildDocx } from './docx-fixtures';
import { independent } from './one-go-fixtures';
import { DocNet, actDoc, assertSameDoc, at, bodyText, docFile, joinDoc, saveDoc, seeded, type DocMember } from './collab-doc-fixtures';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 60, 100000));

test('the file of a document says the room and the step it was saved at, and who opens it joins there', async () => {
  const bytes = await docFile(['one', 'two']);
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  net.drain();
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: ' and more' }]));
  assert.ok(actDoc(a, [{ t: 'split', at: at(1, 1) }]));
  net.drain();
  const stored = await saveDoc(a);
  const again = await open(stored);
  const stamp = await readCollabStamp(again.package!);
  assert.ok(stamp, 'the file carries the mark');
  assert.equal(stamp.rev, 2);
  assert.equal(bodyText(await again.document!()), 'one and more\nt\nwo');
  // The room hears of the file; who opens it joins at its step, and gets what came since.
  a.session.announceSaved([contentVersion(stored)], a.session.rev);
  assert.ok(actDoc(a, [{ t: 'text', at: at(2, 2), x: '!' }]));
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  assert.equal(late.session.state, 'online');
  assert.equal(late.session.rev, 3);
  assertSameDoc([a, late], 'after joining with the saved file');
  // A file that is written while a change of its author waits is the file of no step: it says nothing of itself.
  const cut = await joinDoc(net, 'cut', bytes);
  net.drain();
  cut.transport.suspend();
  assert.ok(actDoc(cut, [{ t: 'text', at: at(0, 0), x: 'unsent ' }]));
  const done = await writeSharedFile(cut.session, cut.file.package, () => cut.file.save());
  assert.equal(done.rev, undefined);
  assert.equal(await readCollabStamp((await open(done.bytes)).package!), undefined, 'no mark in it');
});

test('two participants who hold the same steps write the same file, part for part', async (t) => {
  const bytes = await aloneFile();
  let same = 0;
  let runs = 0;
  let identical = 0;
  let lateSame = 0;
  let lateRuns = 0;
  const differing: string[] = [];
  for (let seed = 1; seed <= Math.max(6, Math.floor(SEEDS / 3)); seed++) {
    const random = seeded(seed * 97);
    const net = new DocNet();
    const members: DocMember[] = [await joinDoc(net, 'a', bytes), await joinDoc(net, 'b', bytes)];
    const histories = members.map((member) => new CollabDocHistory<string>(member.session));
    net.drain(random);
    for (let tick = 0; tick < 50; tick++) {
      const who = Math.floor(random() * 2);
      if (random() < 0.4) net.step(random);
      else if (random() < 0.1) histories[who].undo();
      else {
        try {
          histories[who].apply(make(members[who].doc, seeded(seed * 1000 + tick))(members[who].doc).ops, '', '');
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
    }
    net.drain(random);
    assertSameDoc(members, `seed ${seed}`);
    const files = [await saveDoc(members[0]), await saveDoc(members[1])];
    runs++;
    if (await samePackage(files[0], files[1])) same++;
    else {
      const zips = files.map((file) => ZipArchive.read(file));
      for (const name of zips[0].names()) {
        const one = Buffer.from(await zips[0].read(name)).toString('latin1');
        const two = zips[1].names().indexOf(name) >= 0 ? Buffer.from(await zips[1].read(name)).toString('latin1') : '';
        if (one !== two && differing.length < 4) {
          let i = 0;
          while (i < one.length && one[i] === two[i]) i++;
          differing.push(`seed ${seed} ${name} at ${i}: …${one.slice(Math.max(0, i - 60), i + 60)}… against …${two.slice(Math.max(0, i - 60), i + 60)}…`);
        }
      }
    }
    if (Buffer.from(files[0]).equals(Buffer.from(files[1]))) identical++;
    // One who joins with the saved file and goes on from there: does it write what the others write? Measured, not required
    // (a document that was read from a file is not, block for block, the document that was edited into it).
    members[0].session.announceSaved([contentVersion(files[0])], members[0].session.rev);
    net.drain(random);
    const late = await joinDoc(net, 'late', files[0]);
    net.drain(random);
    for (let tick = 0; tick < 6; tick++) {
      try {
        histories[0].apply(make(members[0].doc, seeded(seed * 5000 + tick))(members[0].doc).ops, '', '');
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
      net.drain(random);
    }
    lateRuns++;
    if (late.session.state === 'online') {
      const mine = await saveDoc(members[0]);
      const theirs = await saveDoc(late);
      if (await samePackage(mine, theirs)) lateSame++;
      else assert.ok(await sameButControlIds(mine, theirs), `seed ${seed}: who joined from the saved file writes another file, and not only by the id of a content control`);
    }
    for (const member of members.concat([late])) member.session.leave();
  }
  assert.deepEqual(differing, [], 'the parts that differ');
  assert.equal(same, runs, 'every pair of files holds the same parts');
  t.diagnostic(`${runs} pairs: ${same} with the same parts, ${identical} the same bytes; with one who joined from the saved file: ${lateSame} of ${lateRuns} the same parts`);
});

// ------------------------------------------------------------------ the bridge

interface Page {
  host: BridgeHost;
  sent: { type: string; [key: string]: unknown }[];
  session?: DocCollabSession;
  doc?: DocModel;
  changes: DocChange[];
  notices: string[];
}

function page(server: MemoryServer): Page {
  const state: Page = { host: undefined as unknown as BridgeHost, sent: [], changes: [], notices: [] };
  const view = (doc: DocModel): BridgeDocumentView => {
    state.doc = doc;
    return {
      destroy: () => undefined,
      canUndo: false,
      canRedo: false,
      changeCount: 0,
      commitEditing: () => undefined,
      markSaved: () => undefined,
      attachCollab: (session) => {
        state.session = session;
      },
      collabChanged: (change) => void state.changes.push(change),
      collabNotice: (notice) => void state.notices.push(notice.code),
      collabPeersChanged: () => undefined,
    };
  };
  state.host = {
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: (bytes, options) => open(bytes, options),
    createView: () => {
      throw new Error('no workbook is shown in this test');
    },
    createDocumentView: (_container, doc) => view(doc as DocModel),
    createTransport: () => server.connect(),
    timer: () => undefined,
    frame: () => undefined,
  } as unknown as BridgeHost;
  return state;
}

const settled = (): Promise<void> => new Promise<void>((resolve) => setTimeout(resolve, 20));

test('bridge: a docx with `collab` is edited together, and the application hears of the room and of what the others change', async () => {
  const server = new MemoryServer();
  const bytes = await docFile(['one', 'two']);
  const first = page(server);
  const second = page(server);
  const a = createBridge(first.host);
  const b = createBridge(second.host);
  await a.receive({ type: 'load', id: 'a', data: base64Encode(bytes), collab: { url: 'wss://relay.example/doc', user: { name: 'Ada' } } });
  await b.receive({ type: 'load', id: 'b', data: base64Encode(bytes), collab: 'wss://relay.example/doc' });
  await settled();
  const state = first.sent.filter((message) => message.type === 'collabState').pop()!;
  assert.equal(state.state, 'online');
  assert.equal((state.people as unknown[]).length, 2, 'both are listed');
  assert.ok(first.session && second.session, 'the views were tied to their sessions');
  // What the view of the first page does (it makes its changes as steps of the session).
  first.session.submit([{ t: 'text', at: at(0, 3), x: ' typed' }]);
  await settled();
  assert.equal(bodyText(second.doc!), 'one typed\ntwo', 'the other page has it');
  assert.equal(second.changes.length, 1, 'its view was told');
  assert.equal(second.sent.filter((message) => message.type === 'changed').length, 1, 'and so was its application');
  // The file the second page hands over says where it was saved.
  await b.receive({ type: 'save', id: 's' });
  const saved = second.sent.find((message) => message.type === 'saved')!;
  const stamp = await readCollabStamp((await open(base64Decode(saved.data as string))).package!);
  assert.equal(stamp && stamp.rev, 1);
  a.destroy();
  b.destroy();
});

test('bridge: a document that tracks its changes, or is no docx, is shown without a room, and the application is told', async () => {
  const server = new MemoryServer();
  const tracked = await docFile(['one'], { settings: '<w:trackRevisions/>' });
  const f = page(server);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(tracked), collab: 'wss://relay.example/doc' });
  await settled();
  const states = f.sent.filter((message) => message.type === 'collabState');
  assert.deepEqual(
    states.map((message) => [message.state, message.reason]),
    [['closed', 'unsupported']],
  );
  assert.equal(f.session, undefined);
  assert.equal(f.sent.some((message) => message.type === 'loaded'), true, 'the document is shown all the same');
  bridge.destroy();
  // A document that does not track, and holds changes that were tracked: the numbers of those changes are given
  // from a counter of each participant when a change is cut in two. It is not edited together either.
  const kinds: [string, string][] = [
    ['an insertion', '<w:p><w:ins w:id="7" w:author="Someone" w:date="2020-01-01T00:00:00Z"><w:r><w:t>put in</w:t></w:r></w:ins></w:p>'],
    ['a deletion', '<w:p><w:del w:id="8" w:author="Someone" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>taken out</w:delText></w:r></w:del></w:p>'],
    ['a change of how a run looks', '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="9" w:author="Someone" w:date="2020-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>bold now</w:t></w:r></w:p>'],
  ];
  for (const [what, body] of kinds) {
    const holding = await docFile(['one'], { body });
    assert.equal(holdsRevisions(await (await open(holding)).document!()), true, what);
    const g = page(new MemoryServer());
    const other = createBridge(g.host);
    await other.receive({ type: 'load', id: 'a', data: base64Encode(holding), collab: 'wss://relay.example/doc' });
    await settled();
    assert.deepEqual(g.sent.filter((message) => message.type === 'collabState').map((message) => [message.state, message.reason]), [['closed', 'unsupported']], what);
    assert.equal(g.session, undefined, what);
    other.destroy();
  }
  assert.equal(holdsRevisions(await (await open(await docFile(['one', 'two']))).document!()), false);
  // Nor by a program that makes the session itself.
  const holding = await docFile(['one'], { body: kinds[0][1] });
  await assert.rejects(joinDoc(new DocNet(), 'a', holding), { name: 'CollabUnsupportedError' });
});

test('a link made after someone joined from a saved file gets the same relationship for everybody: their files are the same', async () => {
  const bytes = await docFile(['one two', 'three four']);
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  // A link that is made and whose text is deleted again: its relationship stays with who made it. And a list, whose
  // definitions get a part of their own when the file is written (and a relationship to it).
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertHyperlink', at: at(0, 0), to: at(0, 3), a: ['https://example.com/one'] }]));
  net.drain();
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 0), to: at(0, 4) }]));
  net.drain();
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'toggleList', at: at(1, 0), to: at(1, 0), a: ['bullet'] }]));
  net.drain();
  const stored = await saveDoc(a);
  a.session.announceSaved([contentVersion(stored)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  assert.equal(late.session.state, 'online');
  // A link made now: by who joined late, and by who was there.
  assert.ok(actDoc(late, [{ t: 'cmd', name: 'insertHyperlink', at: at(1, 0), to: at(1, 5), a: ['https://example.com/three'] }]));
  net.drain();
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertHyperlink', at: at(0, 0), to: at(0, 3), a: ['https://example.com/two'] }]));
  net.drain();
  assertSameDoc([a, b], 'who were there');
  const one = await saveDoc(a);
  const two = await saveDoc(late);
  const links = async (file: Uint8Array): Promise<string[]> => (new TextDecoder().decode(await ZipArchive.read(file).read('word/_rels/document.xml.rels')).match(/<Relationship [^>]*hyperlink[^>]*>/g) || []).sort();
  assert.deepEqual(await links(two), await links(one), 'the links have the same relationships');
  assert.ok(await samePackage(one, two), 'the two files are the same, part for part');
});

test('the relationship of a link travels with the command: who holds a relationship more (one that was let go of, one of a file read later) gives the link the same', async () => {
  const bytes = await docFile(['one two']);
  const plain = await (await open(bytes)).document!();
  const more = await (await open(bytes)).document!();
  // A relationship of a link whose text went: one participant holds it, one that read the file later does not.
  more.body.relationships.set('rId1', { id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target: 'https://example.com/old', external: true, part: null } as never);
  // As the view makes it: the command of the document, which the room hears of as a step.
  tracing(true);
  let wire: Extract<DocWire, { t: 'cmd' }>;
  try {
    const edit = traced(raw).insertHyperlink(plain, { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 3 } }, 'https://example.com/new');
    wire = (wiresOf(edit.ops) as DocWire[])[0] as Extract<DocWire, { t: 'cmd' }>;
  } finally {
    tracing(false);
  }
  assert.equal(wire.a.length, 2, 'the id travels');
  for (const doc of [plain, more]) applyDocOps(doc, runWire(doc, wire).ops);
  const linkOf = (doc: DocModel): string | undefined => {
    for (const inline of (doc.body.blocks[0] as Paragraph).children) if (inline.type === 'run' && inline.wrap && inline.wrap.kind === 'hyperlink') return inline.wrap.rId;
    return undefined;
  };
  assert.ok(linkOf(plain));
  assert.equal(linkOf(more), linkOf(plain), 'the same relationship');
  // A command whose relationship is one the document has already is refused, by everybody alike.
  assert.throws(() => runWire(more, { ...wire, a: [wire.a[0], 'rId1'] } as DocWire), RangeError);
  assert.throws(() => decodeStep([{ t: 'cmd', name: 'insertHyperlink', at: ['body', [0], 0], to: ['body', [0], 3], a: ['https://example.com/', 'rId"/>'] }]), { name: 'WireError' });
  // A number of more than nine digits in it: no id a file is written with.
  assert.throws(() => decodeStep([{ t: 'cmd', name: 'insertHyperlink', at: ['body', [0], 0], to: ['body', [0], 3], a: ['https://example.com/', 'rId12345678901'] }]), { name: 'WireError' });
  assert.doesNotThrow(() => decodeStep([{ t: 'cmd', name: 'insertHyperlink', at: ['body', [0], 0], to: ['body', [0], 3], a: ['https://example.com/', 'rIdC1a2b3c4d5'] }]));
});

test('the reader that knows nothing of the engine finds bookmarks, paragraph ids and rsids that a file may not hold', async (t) => {
  const good = await buildDocx({ body: '<w:p w14:paraId="1A2B3C4D" w14:textId="7FFFFFFF" w:rsidR="00AB12CD"><w:bookmarkStart w:id="1" w:name="one"/><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' });
  if (independent({ probe: good }).length > 0) assert.fail('a file that is whole is found to be broken');
  const cases: [string, string, RegExp][] = [
    ['two bookmarks of one id', '<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="1" w:name="b"/><w:bookmarkEnd w:id="1"/></w:p>', /bookmark/],
    ['two bookmarks of one name', '<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="A"/><w:bookmarkEnd w:id="2"/></w:p>', /bookmark/],
    ['a bookmark that never ends', '<w:p><w:bookmarkStart w:id="1" w:name="a"/></w:p>', /bookmark/],
    ['a bookmark that never started', '<w:p><w:bookmarkEnd w:id="4"/></w:p>', /bookmark/],
    ['two paragraphs of one id', '<w:p w14:paraId="11111111"/><w:p w14:paraId="11111111"/>', /paragraph id/],
    ['a paragraph id of the upper half', '<w:p w14:paraId="FFFFFFFF"/>', /paragraph id|text id/],
    ['a text id that is not eight digits', '<w:p w14:textId="12"/>', /paragraph id|text id/],
    ['an rsid that is not eight digits', '<w:p w:rsidR="0"><w:r w:rsidRPr="ABC"><w:t>x</w:t></w:r></w:p>', /rsid/],
    ['a bookmark that starts without a name', '<w:p><w:bookmarkStart w:id="1"/><w:bookmarkEnd w:id="1"/></w:p>', /without a name/],
    ['one id written twice in two ways', '<w:p><w:bookmarkStart w:id="9" w:name="a"/><w:bookmarkEnd w:id="9"/><w:bookmarkStart w:id="09" w:name="b"/><w:bookmarkEnd w:id="09"/></w:p>', /two bookmarks (start|end) with the id 9/],
    ['a bookmark that ends before it starts', '<w:p><w:bookmarkEnd w:id="3"/><w:r><w:t>x</w:t></w:r><w:bookmarkStart w:id="3" w:name="c"/></w:p>', /ends before it starts/],
  ];
  const found = independent(Object.fromEntries(await Promise.all(cases.map(async ([what, body]) => [what.replace(/ /g, '-'), await buildDocx({ body })]))));
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return t.skip('python3 is not there');
  }
  for (const [what, , pattern] of cases) assert.ok(found.some((line) => line.indexOf(what.replace(/ /g, '-')) >= 0 && pattern.test(line)), `${what}: ${JSON.stringify(found).slice(0, 600)}`);
});

/**
 * Whether two files are the same but for the ids of content controls: the one known way in which
 * who joined from a saved file and who was there write other files (a control that Enter cut in
 * two keeps its id in the model of who cut it, in its first half; the file holds the second half
 * without it; when the first half goes, the one writes the id, the other does not). Its only
 * consequence: the proof of a save does not hold, and the participant saves its work itself.
 */
async function sameButControlIds(one: Uint8Array, two: Uint8Array): Promise<boolean> {
  const a = ZipArchive.read(one);
  const b = ZipArchive.read(two);
  const names = a.names().slice().sort();
  if (names.join() !== b.names().slice().sort().join()) return false;
  const strip = (text: string): string => text.replace(/(<w:sdtPr>(?:(?!<\/w:sdtPr>).)*?)<w:id w:val="-?\d+"\/>/g, '$1').replace(/<w:sdtPr><\/w:sdtPr>/g, '<w:sdtPr/>');
  for (const name of names) {
    const x = new TextDecoder().decode(await a.read(name));
    const y = new TextDecoder().decode(await b.read(name));
    if (x === y) continue;
    if (name !== 'word/document.xml' || strip(x) !== strip(y)) return false;
  }
  return true;
}

test('KNOWN LIMIT: a content control that Enter cut in two, saved, and the first half deleted: who joined from the saved file writes the same file but for the id of the control', async () => {
  for (const cut of [{ t: 'split', at: at(0, 13) }, { t: 'cmd', name: 'insertPageBreak', at: at(0, 13), a: [] }] as DocWire[]) await controlCut(cut);
});

async function controlCut(cut: DocWire): Promise<void> {
  const bytes = await docFile([], { body: '<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent><w:r><w:t>inside sdt</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r></w:p>' });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  net.drain();
  assert.ok(actDoc(a, [cut]));
  net.drain();
  assert.ok(bodyText(a.doc).split('\n').indexOf(' sdt after') > 0, `${cut.t}: the control lies over two paragraphs`);
  const stored = await saveDoc(a);
  a.session.announceSaved([contentVersion(stored)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  assert.equal(late.session.state, 'online');
  // The first half of the control goes, with the text before it.
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 0), to: at(1, 0) }]));
  net.drain();
  assertSameDoc([a], 'a');
  assert.equal(bodyText(late.doc), bodyText(a.doc));
  const mine = await saveDoc(a);
  const theirs = await saveDoc(late);
  assert.ok((await samePackage(mine, theirs)) || (await sameButControlIds(mine, theirs)), `${cut.t}: the two files differ in more than the id of the control`);
}

test('what names itself in a document is what the reader that knows nothing of the engine finds: one walker each, the same things', async (t) => {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return t.skip('python3 is not there');
  }
  const box = '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="1" name="T"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="66666666"><w:bookmarkStart w:id="31" w:name="inbox"/><w:r><w:t>box</w:t></w:r><w:bookmarkEnd w:id="31"/></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><w:p w14:paraId="66666666"/></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>';
  // Both branches of an AlternateContent that name other things; bookmarks between rows and cells; a block and an inline element the model does not know; math.
  const branches = box.replace('<mc:Fallback><w:pict><w:p w14:paraId="66666666"/></w:pict></mc:Fallback>', '<mc:Fallback><w:pict><w:p w14:paraId="67676767"><w:bookmarkStart w:id="32" w:name="fallback"/><w:bookmarkEnd w:id="32"/></w:p></w:pict></mc:Fallback>');
  const hidden =
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:bookmarkStart w:id="50" w:name="rows"/><w:tr><w:bookmarkStart w:id="51" w:name="cells"/><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc><w:bookmarkEnd w:id="51"/></w:tr><w:bookmarkEnd w:id="50"/></w:tbl>' +
    '<w:customXmlFoo><w:p w14:paraId="0ABCDEF0"><w:bookmarkStart w:id="52" w:name="unknown"/><w:r><w:t>u</w:t></w:r><w:bookmarkEnd w:id="52"/></w:p></w:customXmlFoo>' +
    '<w:p><w:foo><w:bookmarkStart w:id="53" w:name="inline"/></w:foo><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="53"/><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:bookmarkStart w:id="54" w:name="math"/><m:r><m:t>x</m:t></m:r><w:bookmarkEnd w:id="54"/></m:oMath></w:p>';
  const bytes = await buildDocx({
    body: '<w:p w14:paraId="11111111"><w:bookmarkStart w:id="1" w:name="one"/><w:r><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' + box + branches + hidden + '<w:bookmarkStart w:id="40" w:name="tablemark"/><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="22222222"><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:bookmarkEnd w:id="40"/><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
    headers: { rIdH: '<w:p w14:paraId="33333333"><w:bookmarkStart w:id="7" w:name="hb"/><w:r><w:t>head</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>' },
    footnotes: '<w:footnote w:type="separator" w:id="-1"><w:p w14:paraId="44444444"><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
    comments: '<w:comment w:id="0" w:author="x"><w:p w14:paraId="55555555"><w:r><w:t>c</w:t></w:r></w:p></w:comment>',
  });
  const doc = await (await open(bytes)).document!();
  const mine: string[] = [];
  const namesOf = (wireModule as unknown as { namesOfDocument: (doc: DocModel, skip: unknown[], visit: (thing: { kind: string; id: string; name?: string }) => void) => void }).namesOfDocument;
  namesOf(doc, [], (thing) => mine.push(thing.kind === 'para' ? `para ${thing.id}` : thing.kind === 'start' ? `start ${thing.id} ${thing.name}` : `end ${thing.id}`));
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.names-'));
  try {
    writeFileSync(joinPath(dir, 'file.docx'), bytes);
    const found = JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir, '--names'], { encoding: 'utf8', timeout: 120000 })) as { [file: string]: string[][] };
    const theirs = found['file.docx'].map((thing) => (thing[0] === 'para' ? `para ${thing[1].toUpperCase()}` : thing[0] === 'start' ? `start ${thing[1]} ${thing[2]}` : `end ${thing[1]}`));
    assert.deepEqual(mine.slice().sort(), theirs.sort());
    for (const thing of ['para 66666666', 'para 67676767', 'start 32 fallback', 'para 44444444', 'para 55555555', 'start 40 tablemark', 'start 50 rows', 'start 51 cells', 'para 0ABCDEF0', 'start 52 unknown', 'start 53 inline', 'start 54 math']) assert.ok(mine.indexOf(thing) >= 0, `${thing}: ${mine.join(', ')}`);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('what the engine counts is what the saved file names, wherever it stands: properties kept as XML, wrappers, tables, notes, headers, the background', async (t) => {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return t.skip('python3 is not there');
  }
  const namesOf = (wireModule as unknown as { namesOfDocument?: (doc: DocModel, skip: unknown[], visit: (thing: { kind: string; id: string; name?: string }) => void) => void }).namesOfDocument;
  assert.ok(namesOf, 'the reading of what a document names');
  const P0 = '<w:p w14:paraId="1A2B3C4D"><w:r><w:t>zero</w:t></w:r></w:p>';
  const BM = '<w:bookmarkStart w:id="9" w:name="mk"/>';
  const BE = '<w:bookmarkEnd w:id="9"/>';
  const tbl = (inner: string, trAttr = '', tcPr = ''): string => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr${trAttr}><w:tc>${tcPr}${inner}</w:tc></w:tr></w:tbl><w:p/>`;
  const pp = (x: string): string => `<w:p>${x}</w:p>`;
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const cases: [string, Record<string, unknown>][] = [
    ['a bookmark in cell properties', { body: P0 + tbl('<w:p><w:r><w:t>c</w:t></w:r></w:p>', '', `<w:tcPr><w:tcW w:w="0" w:type="auto"/><w:foo>${BM}${BE}</w:foo></w:tcPr>`) }],
    ['the id of a row', { body: P0 + tbl('<w:p><w:r><w:t>c</w:t></w:r></w:p>', ' w14:paraId="0ABCDEF0" w14:textId="77777777"') }],
    ['a bookmark in paragraph properties', { body: P0 + `<w:p><w:pPr><w:foo>${BM}${BE}</w:foo></w:pPr><w:r><w:t>q</w:t></w:r></w:p>` }],
    ['an id in run properties', { body: P0 + `<w:p><w:r><w:rPr><w:foo w14:paraId="0ABCDEF0"/></w:rPr><w:t>q</w:t></w:r></w:p>` }],
    ['a bookmark in the properties of a control', { body: P0 + pp(`<w:sdt><w:sdtPr><w:foo>${BM}${BE}</w:foo></w:sdtPr><w:sdtContent><w:r><w:t>s</w:t></w:r></w:sdtContent></w:sdt>`) }],
    ['blocks in an mc:AlternateContent', { body: P0 + `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w14"><w:p w14:paraId="0ABCDEF0">${BM}<w:r><w:t>a</w:t></w:r>${BE}</w:p></mc:Choice><mc:Fallback><w:p w14:paraId="0ABCDEF0">${BM}<w:r><w:t>a</w:t></w:r>${BE}</w:p></mc:Fallback></mc:AlternateContent><w:p/>` }],
    ['a VML text box', { body: P0 + pp(`<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:10pt;height:10pt"><v:textbox><w:txbxContent><w:p w14:paraId="0ABCDEF0">${BM}<w:r><w:t>v</w:t></w:r>${BE}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`) }],
    ['a kept element of the last section', { body: P0 + `<w:p/><w:sectPr><w:foo>${BM}${BE}</w:foo></w:sectPr>` }],
    ['a kept element of the table grid', { body: P0 + `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:foo>${BM}${BE}</w:foo></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>` }],
    ['the background', { document: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:background w:color="FFFFFF"><w:foo w14:paraId="0ABCDEF0">${BM}${BE}</w:foo></w:background><w:body>${P0}</w:body></w:document>` }],
    ['kept elements of the notes', { body: P0, footnotes: `<w:foo><w:p w14:paraId="0ABCDEF0">${BM}${BE}</w:p></w:foo><w:footnote w:id="2"><w:p/></w:footnote>` }],
    ['a section in a header', { body: P0 + '<w:p/><w:sectPr><w:headerReference w:type="default" r:id="h1"/></w:sectPr>', headers: { h1: `<w:p/><w:sectPr><w:foo>${BM}${BE}</w:foo></w:sectPr>` } }],
    ['ids of a cell, a table, a link, a run', { body: P0 + `<w:tbl w14:paraId="0ABCDEF1"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc w14:paraId="0ABCDEF2"><w:p/></w:tc></w:tr></w:tbl>` + pp(`<w:hyperlink w:anchor="a" w14:paraId="0ABCDEF3"><w:r w14:paraId="0ABCDEF4"><w:t>1</w:t></w:r></w:hyperlink>`) }],
    ['math', { body: P0 + `<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath>${BM}<m:r><m:t>x</m:t></m:r>${BE}</m:oMath></m:oMathPara></w:p>` }],
  ];
  for (const [what, spec] of cases) {
    const net = new DocNet();
    const a = await joinDoc(net, 'a', await buildDocx(spec));
    const evil = (name: string): void => void name;
    void evil;
    // Something is changed, so that the body is written from the model as well.
    assert.ok(actDoc(a, [{ t: 'text', at: at(0, 0), x: 'w' }]), what);
    net.drain();
    const mine: string[] = [];
    namesOf(a.doc, [], (thing) => mine.push(thing.kind === 'para' ? `para ${thing.id}` : thing.kind === 'start' ? `start ${thing.id} ${thing.name}` : `end ${thing.id}`));
    const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.names-'));
    try {
      writeFileSync(joinPath(dir, 'file.docx'), await saveDoc(a));
      const found = JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir, '--names'], { encoding: 'utf8', timeout: 120000 })) as { [file: string]: string[][] };
      const theirs = found['file.docx'].map((thing) => (thing[0] === 'para' ? `para ${thing[1].toUpperCase()}` : thing[0] === 'start' ? `start ${thing[1]} ${thing[2]}` : `end ${thing[1]}`));
      assert.deepEqual(mine.slice().sort(), theirs.sort(), what);
      assert.deepEqual(JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 120000 })), [], what);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }
});

