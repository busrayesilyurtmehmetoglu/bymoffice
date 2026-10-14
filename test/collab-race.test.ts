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
      'import sys, os, json, zipfile',
      'import xml.etree.ElementTree as E',
      'bad = []',
      'd = sys.argv[1]',
      'for f in sorted(os.listdir(d)):',
      '  try:',
      '    z = zipfile.ZipFile(os.path.join(d, f))',
      '    if z.testzip() is not None: bad.append(f + ": a part does not unpack")',
      '    names = z.namelist()',
      '    if "[Content_Types].xml" not in names or "_rels/.rels" not in names: bad.append(f + ": no content types or relationships")',
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
  // An archive that is changed while it is written: the write holds what it started with, the next one what came after.
  const archive = ZipArchive.read(await startFile());
  const writing = archive.write();
  // A few moments in: the write is under way.
  await ticks(4);
  archive.set('late.xml', new TextEncoder().encode('<late/>'));
  const second = archive.write();
  const [one, two] = await Promise.all([writing, second]);
  assert.equal(ZipArchive.read(one).names().indexOf('late.xml'), -1, 'what was set after a write began is not in that write');
  assert.equal(new TextDecoder().decode(await ZipArchive.read(two).read('late.xml')), '<late/>', 'and whole in the next');
  files['zip-one'] = one;
  files['zip-two'] = two;
  assert.deepEqual(independent(files), [], 'every file is whole to a reader that is not this engine');
});
