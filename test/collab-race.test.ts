// Whatever writes the package of a file does so alone: a save, a save as, a file that is written
// only to compare, at any moment relative to each other. Every file that comes out is whole - read
// by this engine and by a reader that knows nothing of it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { proofFile, readCollabStamp, writeSharedFile } from '../src/formats/xlsx/collab-stamp';
import { OpcPackage, ZipArchive, createBlank, open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { packageTurn } from '../src/package-turn';
import { Net, act, join, saveOf, seeded, startFile } from './collab-fixtures';
import { RACE_DELAYS, changeSheet, docEdit, docStart, independent, load, moments, sameBytes, type Engine } from './one-go-fixtures';

const ticks = async (n: number): Promise<void> => {
  for (let i = 0; i < n; i++) await Promise.resolve();
};

test('a file written to compare and a save of the same file, started any number of moments apart: both come out whole, and the stored one says its step', async () => {
  const bytes = await startFile();
  const files: { [name: string]: Uint8Array } = {};
  for (const stamped of [false, true]) {
    const net = new Net();
    const random = seeded(1);
    const a = await join(net, 'a', bytes);
    net.drain(random);
    act(a, () => setCellInput(a.wb, 0, 12, 0, 'typed'));
    net.drain(random);
    // A package that was saved in the room before carries a stamp already; one that never was has none.
    if (stamped) await saveOf(a);
    const pkg = a.file.package!;
    const turn = packageTurn(a.file);
    for (const delay of [0, 1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 233, 300]) {
      for (const proofFirst of [true, false]) {
        const proof = (): Promise<{ bytes: Uint8Array; rev: number | undefined }> => proofFile(a.session, () => turn.save(), (given) => OpcPackage.open(given), turn.locked);
        const save = (): Promise<{ bytes: Uint8Array; rev: number | undefined }> => writeSharedFile(a.session, pkg, () => turn.save(), turn.locked);
        const first = (proofFirst ? proof : save)();
        await ticks(delay);
        const second = (proofFirst ? save : proof)();
        const [x, y] = await Promise.all([first, second]);
        const saved = proofFirst ? y : x;
        const proven = proofFirst ? x : y;
        const label = `${stamped ? 'stamped' : 'unstamped'} package, ${delay} moments, ${proofFirst ? 'the proof' : 'the save'} first`;
        for (const [what, out] of [['the stored file', saved], ['the file to compare', proven]] as [string, { bytes: Uint8Array; rev: number | undefined }][]) {
          const again = await open(out.bytes);
          assert.equal((await again.workbook()).sheets.length, 2, `${label}: ${what} opens`);
          assert.deepEqual(await readCollabStamp(again.package!), a.session.stampNow, `${label}: ${what} says the room and the step`);
          const archive = ZipArchive.read(out.bytes);
          for (const name of archive.names()) if (/\.(xml|rels)$/.test(name)) assert.ok((await archive.read(name)).length > 0, `${label}: ${name} of ${what} is not empty`);
        }
        files[`${stamped ? 's' : 'u'}-${delay}-${proofFirst ? 'p' : 's'}-saved`] = saved.bytes;
        files[`${stamped ? 's' : 'u'}-${delay}-${proofFirst ? 'p' : 's'}-proof`] = proven.bytes;
      }
    }
  }
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
});

test('saves of one file at once, by whatever way: each waits for the one before it, and every file is whole', async () => {
  const files: { [name: string]: Uint8Array } = {};
  // A workbook: save, save, save as, save - all at once, with a change before each.
  const book = await open(await startFile());
  const wb = await book.workbook();
  const { applyOp } = await import('../src/ops/ops');
  const outs: Promise<Uint8Array>[] = [];
  for (let i = 0; i < 6; i++) {
    for (const op of setCellInput(wb, 0, 14 + i, 0, `change ${i}`)) applyOp(wb, op);
    outs.push(i % 3 === 2 ? book.saveAsXlsx!() : book.save());
    if (i % 2 === 0) await ticks(i * 7);
  }
  const saved = await Promise.all(outs);
  for (let i = 0; i < saved.length; i++) {
    const again = await (await open(saved[i])).workbook();
    assert.equal(again.sheets.length, 2, `workbook, save ${i}: the file opens`);
    files[`xlsx-${i}`] = saved[i];
  }
  const last = await (await open(saved[saved.length - 1])).workbook();
  assert.equal(last.sheets[0].rows.get(19)!.cells[0]!.value, 'change 5', 'the last save holds the last change');
  // A document: two saves at once after its model was read.
  const doc = await open(await createBlank('docx'));
  await doc.document();
  const both = await Promise.all([doc.save(), ticks(3).then(() => doc.save()), doc.saveAsDocx!()]);
  for (let i = 0; i < both.length; i++) {
    const again = await open(both[i]);
    assert.ok(await again.document(), `document, save ${i}: the file opens`);
    files[`docx-${i}`] = both[i];
  }
  // An archive that is changed while it is written: the write holds what it started with (and says how many entries
  // it holds), the next one what came after. Entries were changed before, so that the write has work that takes moments.
  const archive = ZipArchive.read(await startFile());
  const before = archive.names();
  for (const name of before) if (/\.xml$/.test(name)) archive.set(name, new TextEncoder().encode(`${new TextDecoder().decode(await archive.read(name))}<!-- changed -->`));
  const writing = archive.write();
  await ticks(3);
  archive.set('late.xml', new TextEncoder().encode('<late/>'));
  archive.set('later.xml', new TextEncoder().encode('<later/>'));
  const second = archive.write();
  // The first is done and the second is under way (it has two new entries to compress) when an entry is deleted.
  await writing;
  await ticks(2);
  archive.delete(before[before.length - 1]);
  const third = archive.write();
  const [one, two, three] = await Promise.all([writing, second, third]);
  assert.deepEqual(ZipArchive.read(one).names().slice().sort(), before.slice().sort(), 'the first write holds the entries it started with, and says how many: what was set after it began is not in it');
  assert.equal(new TextDecoder().decode(await ZipArchive.read(two).read('late.xml')), '<late/>', 'the next write holds what came after, whole');
  assert.equal(ZipArchive.read(two).names().length, before.length + 2, 'and what was deleted while it ran, still');
  assert.equal(ZipArchive.read(three).names().indexOf(before[before.length - 1]), -1, 'the one after it is without what was deleted');
  files['zip-three'] = three;
  files['zip-one'] = one;
  files['zip-two'] = two;
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
});

test('a workbook that is changed while its file is written: the file is the workbook of before the change or of after it, never of both', async () => {
  const { applyOp } = await import('../src/ops/ops');
  const { formatRange } = await import('../src/ops/format');
  const { deleteSheet, insertSheet } = await import('../src/ops/sheets');
  const { canon } = await import('./collab-fixtures');
  const bytes = await startFile();
  const files: { [name: string]: Uint8Array } = {};
  const edits: [string, (wb: Awaited<ReturnType<Awaited<ReturnType<typeof open>>['workbook']>>) => void][] = [
    ['a new style', (wb) => formatRange(wb, 0, { top: 1, left: 0, bottom: 2, right: 1 }, { bold: true, fill: { rgb: 'FF00FF00' } }).forEach((op) => applyOp(wb, op))],
    ['a sheet that comes, with a formula that reads it', (wb) => [...insertSheet(wb, 1, { name: 'New' }), ...setCellInput(wb, 0, 12, 0, '=New!A1')].forEach((op) => applyOp(wb, op))],
    ['a sheet that goes', (wb) => deleteSheet(wb, 1).forEach((op) => applyOp(wb, op))],
  ];
  let between = 0;
  for (const [name, edit] of edits) {
    for (let delay = 0; delay <= 80; delay += delay < 30 ? 1 : 5) {
      const file = await open(bytes);
      const wb = await file.workbook();
      // Something to write: the save has work to do in several parts.
      setCellInput(wb, 0, 11, 2, 'changed before the save').forEach((op) => applyOp(wb, op));
      const was = canon(wb);
      const saving = file.save();
      await ticks(delay);
      edit(wb);
      const now = canon(wb);
      const out = await saving;
      const got = canon(await (await open(out)).workbook());
      if (got !== was && got !== now) assert.fail(`${name}, ${delay} moments into the save: the file is the workbook of no moment`);
      if (got === now) between++;
      files[`${name.replace(/[^a-z]+/g, '-')}-${delay}`] = out;
    }
  }
  assert.ok(between > 0, 'the change came while the file was written in some of the runs (or the test tests nothing)');
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine: its styles, its sheets and its directory agree');
});

test('a model that is changed while its file is written: written again once the changes pause, not into the next one; what never pauses is not saved, and says so', async () => {
  const { applyDocOp } = await import('../src/ops/doc/ops');
  const { buildEdit } = await import('../src/ops/doc/draft');
  const { editsOf, SaveInterruptedError, writeSteady } = await import('../src/model/edits');
  const file = await open(await createBlank('docx'));
  const doc = await file.document();
  const change = (): void => void applyDocOp(doc, { type: 'setTrackRevisions', on: false } as never);
  // Time is the test's: a sleep is a step of its clock, and what happens meanwhile is scripted.
  let clock = 0;
  let during: (() => void)[] = [];
  const sleeps: number[] = [];
  const options = {
    now: () => clock,
    sleep: async (ms: number) => {
      sleeps.push(ms);
      clock += ms;
      const next = during.shift();
      if (next) next();
    },
  };
  // Changed once while it was written, and twice more right after: the file is written again only after a moment without a change.
  let writes = 0;
  during = [change, change];
  const steady = await writeSteady(doc, async () => (++writes === 1 ? (change(), writes) : writes), options);
  assert.deepEqual([steady, writes, sleeps.length], [2, 2, 3], 'two writes, the second after the changes had paused (three waits: two met a change, the third met none)');
  // What the writer throws while the model changes under it is no failure of the save: it is written again.
  writes = 0;
  const after = await writeSteady(
    doc,
    async () => {
      if (++writes === 1) {
        change();
        throw new Error('tripped over half a change');
      }
      return 'whole';
    },
    options,
  );
  assert.deepEqual([after, writes], ['whole', 2], 'written again, and that file is answered');
  await assert.rejects(writeSteady(doc, async () => Promise.reject(new Error('really broken')), options), /really broken/, 'what it throws in peace is thrown');
  // A model that never pauses: given up after the time a save may wait, with an error of its own.
  during = Array.from({ length: 1000 }, () => change);
  const started = clock;
  await assert.rejects(writeSteady(doc, async () => void change(), options), SaveInterruptedError, 'a model that changes all the time is not written half');
  assert.ok(clock - started >= 10000 && clock - started < 11000, `after ten seconds of waiting (${clock - started} ms)`);
  during = [];
  // A command that is only tried on the document (and taken back at once) is no change of it.
  const counted = editsOf(doc);
  buildEdit(doc, (draft) => {
    draft.apply({ type: 'setTrackRevisions', on: true } as never);
    return { story: 'body', path: [0], offset: 0 } as never;
  });
  assert.equal(editsOf(doc), counted, 'the document counts as unchanged after a command was tried on it');
  assert.ok((await file.save()).length > 0, 'and a save in peace is a file');
  // A result the calculation writes is a change like any other.
  const { CalcEngine } = await import('../src/calc/engine');
  const { applyOp } = await import('../src/ops/ops');
  const book = await (await open(await startFile())).workbook();
  const engine = new CalcEngine(book);
  const ops = setCellInput(book, 0, 1, 1, '55');
  for (const op of ops) applyOp(book, op);
  const before = editsOf(book);
  const done = engine.recalculate([{ sheet: 0, row: 1, col: 1 }]);
  assert.ok(done.some((part) => part.cells.length > 0), 'the calculation wrote results (or the test tests nothing)');
  assert.ok(editsOf(book) > before, 'and each of them counts as a change of the workbook');
});

test('the turn of a package: a save from within it does not wait for it, and the stamp of a room is written within the turn of its file by itself', async () => {
  const file = await open(await startFile());
  const turn = packageTurn(file);
  // From within the turn, started at once: part of it.
  const nested = await Promise.race([turn.locked(() => file.save()), new Promise<string>((resolve) => setTimeout(() => resolve('never came back'), 2000))]);
  assert.ok(nested instanceof Uint8Array, 'a save that is started within the turn comes back');
  assert.ok((await Promise.race([file.save(), new Promise<string>((resolve) => setTimeout(() => resolve('never came back'), 2000))])) instanceof Uint8Array, 'and the turn is free afterwards');
  // As a program without mount() writes the file of a room: with the package and the save of the file, and nothing else.
  const net = new Net();
  const random = seeded(3);
  const a = await join(net, 'a', await startFile());
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'typed'));
  net.drain(random);
  const files: { [name: string]: Uint8Array } = {};
  for (const delay of [0, 1, 2, 3, 5, 8, 13, 21, 34, 55]) {
    const plain = a.file.save();
    await ticks(delay);
    const shared = writeSharedFile(a.session, a.file.package, () => a.file.save());
    const [x, y] = await Promise.all([plain, shared]);
    for (const [what, out] of [['the save', x], ['the save of the room', y.bytes]] as [string, Uint8Array][]) {
      assert.equal((await (await open(out)).workbook()).sheets.length, 2, `${delay} moments apart: ${what} opens`);
      files[`${delay}-${what.length}`] = out;
    }
    assert.deepEqual(await readCollabStamp((await open(y.bytes)).package!), a.session.stampNow, `${delay} moments apart: the file of the room says its step`);
  }
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
  // Two files of the room that are written at once, each for its own step: each says its own, and not the other's.
  const fake = (rev: number) => ({ state: 'online', savableRev: rev, stampNow: { ep: 'room', rev }, whenSettled: async () => undefined, pause: () => undefined, resume: () => false });
  const other = await open(await startFile());
  const { applyOp: apply } = await import('../src/ops/ops');
  const otherBook = await other.workbook();
  for (const delay of [0, 1, 2, 3, 4, 5, 6, 8, 10, 13, 16, 21, 27, 34]) {
    // Something to write, so that the save takes its moments.
    setCellInput(otherBook, 0, 11, 1, `again ${delay}`).forEach((op) => apply(otherBook, op));
    const first = writeSharedFile(fake(1), other.package, () => other.save());
    await ticks(delay);
    const second = writeSharedFile(fake(2), other.package, () => other.save());
    const [x, y] = await Promise.all([first, second]);
    assert.deepEqual([await readCollabStamp((await open(x.bytes)).package!), await readCollabStamp((await open(y.bytes)).package!)], [{ ep: 'room', rev: 1 }, { ep: 'room', rev: 2 }], `${delay} moments apart: each file says the step it was written for`);
  }
});

/** Counts how often the package of a file is packed: once for a save that writes its file once. */
function countPacks(file: { package?: { save(): Promise<Uint8Array> } }): () => number {
  const pkg = file.package!;
  const packed = pkg.save.bind(pkg);
  let count = 0;
  pkg.save = () => {
    count++;
    return packed();
  };
  return () => count;
}

test('a workbook that is changed in every way while its file is written: the file is the workbook of the moment it was read at, written once', async () => {
  const e: Engine = await load(process.cwd());
  const { applyOp } = e.ops;
  // The file to begin with holds something of everything, as a file does: a chart, notes, rules, merged cells, a filter, a hidden sheet.
  const prepared = await open(await startFile());
  const book = await prepared.workbook();
  for (const [seed, kind] of [[1, 'chart+'], [2, 'note'], [3, 'note'], [4, 'rule+'], [5, 'rule+'], [6, 'validation'], [7, 'merge'], [8, 'filter'], [9, 'sheet+'], [10, 'style'], [11, 'chart+']] as [number, string][]) {
    for (let tries = 0; tries < 50 && !changeSheet(e, book, seeded(seed * 100 + tries), kind); tries++);
  }
  const start = await prepared.save();
  const session = async (): Promise<{ file: Awaited<ReturnType<typeof open>>; wb: Awaited<ReturnType<Awaited<ReturnType<typeof open>>['workbook']>> }> => {
    const file = await open(start);
    const wb = await file.workbook();
    // Something to write: the save has work to do.
    setCellInput(wb, 0, 11, 2, 'changed before the save').forEach((op) => applyOp(wb, op));
    return { file, wb };
  };
  const before = await (await session()).file.save();
  const kinds = ['type', 'style', 'sheet+', 'sheet-', 'move sheet', 'rename', 'copy sheet', 'hide sheet', 'tab', 'insert rows', 'delete rows', 'insert columns', 'delete columns', 'note', 'validation', 'rule+', 'rule-', 'rule change', 'chart+', 'chart copy', 'chart paste', 'chart-', 'chart move', 'chart change', 'filter', 'merge', 'unmerge', 'sort', 'paste'];
  const files: { [name: string]: Uint8Array } = { before };
  const fell = { before: 0, after: 0 };
  for (const kind of kinds) {
    // A seed with which the change applies to the file, and does something to it.
    let seed = 0;
    let after: Uint8Array | undefined;
    for (let tries = 1; tries <= 300 && !after; tries++) {
      const probe = await session();
      if (!changeSheet(e, probe.wb, seeded(tries), kind)) continue;
      const bytes = await probe.file.save();
      if (sameBytes(bytes, before)) continue;
      seed = tries;
      after = bytes;
    }
    assert.ok(after, `${kind}: a change of this kind applies to the file`);
    files[`after-${kind.replace(/[^a-z]+/g, '-')}`] = after;
    let both = 0;
    for (const delay of RACE_DELAYS) {
      const { file, wb } = await session();
      const packs = countPacks(file);
      const saving = file.save();
      await moments(delay);
      assert.ok(changeSheet(e, wb, seeded(seed), kind), `${kind}, ${delay}: the change applies`);
      const out = await saving;
      const label = `${kind}, ${delay} moments into the save`;
      const isBefore = sameBytes(out, before);
      if (!isBefore && !sameBytes(out, after)) assert.fail(`${label}: the file is the workbook of no moment`);
      assert.equal(packs(), 1, `${label}: the file was written once`);
      if (isBefore) fell.before++;
      else fell.after++;
      both |= isBefore ? 1 : 2;
      // And nothing of that save is in the way of the next one, which holds the change.
      assert.ok(sameBytes(await file.save(), after), `${label}: the save after it is the workbook with the change`);
      if (!isBefore || delay === 0) files[`${kind.replace(/[^a-z]+/g, '-')}-${delay}`] = out;
    }
    assert.ok(both & 1, `${kind}: the change came after the workbook was read in some of the runs (or the test tests nothing)`);
  }
  assert.ok(fell.after > 0 && fell.before > 0, `the change came before the workbook was read in some runs and after it in others (${JSON.stringify(fell)})`);
  if (process.env.RACE_STATS) console.log('workbook', JSON.stringify(fell));
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
});

test('a document that is changed in every way while its file is written: the file is the document of the moment it was read at, written once', async () => {
  const e: Engine = await load(process.cwd());
  const { DocHistory, setTracking } = e.doc;
  const now = (): Date => new Date('2026-10-12T08:09:10Z');
  const files: { [name: string]: Uint8Array } = {};
  const fell = { before: 0, after: 0 };
  // Three files: without comments, with every part beside the comments, with tracking on (what is typed there is a tracked change).
  for (const fileKind of [0, 3, 4]) {
    const start = await docStart(e, fileKind);
    const session = async (): Promise<{ file: Awaited<ReturnType<typeof open>>; doc: unknown; apply(edit: unknown): void }> => {
      const file = await open(start);
      const doc = await file.document();
      setTracking(doc, { author: 'Me', now });
      const history = new DocHistory(doc);
      const apply = (edit: unknown): void => {
        history.applyEdit(edit, undefined);
        history.seal();
      };
      // Something to write: the save has work to do.
      apply(e.doc.insertText(doc, { story: 'body', path: [1], offset: 2 }, 'typed before the save'));
      return { file, doc, apply };
    };
    const before = await (await session()).file.save();
    files[`before-${fileKind}`] = before;
    const kinds = ['insertText', 'insertLines', 'deleteRange', 'splitParagraph', 'setRunFormat', 'setParaFormat', 'toggleList', 'insertTable', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'insertHyperlink', 'insertImage', 'add', 'track'].concat(fileKind === 0 ? [] : ['reply', 'remove', 'edit', 'resolve']);
    for (const kind of kinds) {
      const made = (doc: unknown, seed: number): { name: string; edit: { ops: unknown[] } } | undefined => {
        const d = e.docFixtures.dice(seed);
        const only = ['add', 'reply', 'remove', 'track'].indexOf(kind) >= 0 ? kind : undefined;
        const edit = only !== undefined ? docEdit(e, doc, d, {}, only) : docEdit(e, doc, d, {});
        return edit && edit.name === kind && edit.edit.ops.length > 0 ? edit : undefined;
      };
      let seed = 0;
      let after: Uint8Array | undefined;
      for (let tries = 1; tries <= 600 && !after; tries++) {
        const probe = await session();
        const edit = made(probe.doc, tries);
        if (!edit) continue;
        probe.apply(edit.edit);
        const bytes = await probe.file.save();
        if (sameBytes(bytes, before)) continue;
        seed = tries;
        after = bytes;
      }
      // Not every kind applies to every file (a row is deleted where there is a table at the place tried).
      if (!after) continue;
      files[`after-${fileKind}-${kind}`] = after;
      for (const delay of RACE_DELAYS) {
        const { file, doc, apply } = await session();
        const packs = countPacks(file);
        const saving = file.save();
        await moments(delay);
        const edit = made(doc, seed);
        assert.ok(edit, `${kind}, ${delay}: the change applies`);
        apply(edit.edit);
        const out = await saving;
        const label = `file ${fileKind}, ${kind}, ${delay} moments into the save`;
        const isBefore = sameBytes(out, before);
        if (!isBefore && !sameBytes(out, after)) assert.fail(`${label}: the file is the document of no moment`);
        assert.equal(packs(), 1, `${label}: the file was written once`);
        if (isBefore) fell.before++;
        else fell.after++;
        assert.ok(sameBytes(await file.save(), after), `${label}: the save after it is the document with the change`);
        if (!isBefore || delay === 0) files[`${fileKind}-${kind}-${delay}`] = out;
      }
    }
  }
  assert.ok(fell.after > 0 && fell.before > 0, `the change came before the document was read in some runs and after it in others (${JSON.stringify(fell)})`);
  if (process.env.RACE_STATS) console.log('document', JSON.stringify(fell), Object.keys(files).filter((name) => name.indexOf('after-') === 0).join(' '));
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
});
