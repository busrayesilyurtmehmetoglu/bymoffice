// Whatever writes the package of a file does so alone: a save, a save as, a file that is written
// only to compare, at any moment relative to each other. Every file that comes out is whole - read
// by this engine and by a reader that knows nothing of it.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { proofFile, readCollabStamp, writeSharedFile } from '../src/formats/xlsx/collab-stamp';
import { OpcPackage, ZipArchive, createBlank, open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { packageTurn } from '../src/package-turn';
import { Net, act, join, saveOf, seeded, startFile } from './collab-fixtures';

/** Every archive: all parts unpack, none is empty that should not be, and every XML part is well formed (Python's zipfile and ElementTree). */
function independent(files: { [name: string]: Uint8Array }): string[] {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return [];
  }
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.race-'));
  try {
    for (const name of Object.keys(files)) writeFileSync(joinPath(dir, `${name}.bin`), files[name]);
    const script = [
      'import sys, os, json, zipfile, struct, re',
      'import xml.etree.ElementTree as E',
      'bad = []',
      'd = sys.argv[1]',
      'for f in sorted(os.listdir(d)):',
      '  try:',
      '    z = zipfile.ZipFile(os.path.join(d, f))',
      '    if z.testzip() is not None: bad.append(f + ": a part does not unpack")',
      '    names = z.namelist()',
      '    if "[Content_Types].xml" not in names or "_rels/.rels" not in names: bad.append(f + ": no content types or relationships")',
      '    raw = open(os.path.join(d, f), "rb").read()',
      '    e = raw.rfind(b"PK\\x05\\x06")',
      '    said = struct.unpack("<H", raw[e + 10:e + 12])[0]',
      '    central = raw.count(b"PK\\x01\\x02")',
      '    local = raw.count(b"PK\\x03\\x04")',
      '    if not (said == len(names) == central == local): bad.append(f + ": the archive says %d entries, its directory has %d, %d are in it, %d are listed" % (said, central, local, len(names)))',
      '    if "xl/styles.xml" in names:',
      '      st = z.read("xl/styles.xml").decode("utf-8")',
      '      m = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, re.S)',
      '      xfs = len(re.findall(r"<xf[ >/]", m.group(1))) if m else 1',
      '      wbx = z.read("xl/workbook.xml").decode("utf-8")',
      '      sheets = len(re.findall(r"<sheet ", wbx))',
      '      parts = [n for n in names if re.match(r"xl/worksheets/[^/]+\\.xml$", n)]',
      '      if sheets != len(parts): bad.append(f + ": the workbook names %d sheets and has %d sheet parts" % (sheets, len(parts)))',
      '      for n in parts:',
      '        for sidx in re.findall(r"<c [^>]*\\bs=\\"(\\d+)\\"", z.read(n).decode("utf-8")):',
      '          if int(sidx) >= xfs: bad.append(f + ": " + n + " names style %s of %d" % (sidx, xfs)); break',
      '    for n in names:',
      '      data = z.read(n)',
      '      if n.endswith(".xml") or n.endswith(".rels"):',
      '        if len(data) == 0: bad.append(f + ": " + n + " is empty")',
      '        else: E.fromstring(data)',
      '  except Exception as e:',
      '    bad.append(f + ": " + str(e))',
      'print(json.dumps(bad))',
    ].join('\n');
    return JSON.parse(execFileSync('python3', ['-I', '-c', script, dir], { encoding: 'utf8', timeout: 120000, stdio: ['ignore', 'pipe', 'inherit'] }));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

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

test('a document that is changed while its file is written is written again; one that never stops changing is not saved, and says so', async () => {
  const { applyDocOp } = await import('../src/ops/doc/ops');
  const { editsOf, SaveInterruptedError, MAX_WRITE_TRIES, writeSteady } = await import('../src/model/edits');
  const file = await open(await createBlank('docx'));
  const doc = await file.document();
  const before = editsOf(doc);
  // An operation that changes nothing the file shows still counts as a change of the model: the writer starts over.
  const same = { type: 'setTrackRevisions', on: false } as never;
  let writes = 0;
  const steady = await writeSteady(doc, async () => {
    writes++;
    if (writes === 1) applyDocOp(doc, same);
    return writes;
  });
  assert.deepEqual([steady, editsOf(doc) > before], [2, true], 'written again after the change, and that file is answered');
  let tries = 0;
  await assert.rejects(
    writeSteady(doc, async () => {
      tries++;
      applyDocOp(doc, same);
    }),
    SaveInterruptedError,
    'a model that changes every time is not written half',
  );
  assert.equal(tries, MAX_WRITE_TRIES, 'after so many tries');
  assert.ok((await file.save()).length > 0, 'and the next save, in peace, is a file');
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
