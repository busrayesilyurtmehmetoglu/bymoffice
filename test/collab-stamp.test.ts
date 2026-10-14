// The stamp of a room in a file: one custom document property, written beside whatever the file
// has, read back as it was written, and nothing else in the package is touched.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readCollabStamp, writeCollabStamp, writeSharedFile } from '../src/formats/xlsx/collab-stamp';
import { open } from '../src/index';
import { parseXml } from '../src/xml';
import { packageProblems } from './chart-edit-fixtures';
import { startFile } from './collab-fixtures';

const decoder = new TextDecoder();
const CUSTOM = '/docProps/custom.xml';

test('a stamp is written into a file that has no custom properties, read back, changed and taken out', async () => {
  const bytes = await startFile();
  const file = await open(bytes);
  const pkg = file.package!;
  assert.equal(await readCollabStamp(pkg), undefined, 'a file that was never in a room says nothing');
  const before = pkg.partNames().slice().sort();
  await writeCollabStamp(pkg, undefined);
  assert.deepEqual(pkg.partNames().slice().sort(), before, 'taking out what is not there adds nothing');
  await writeCollabStamp(pkg, { ep: 'room-1_A', rev: 12 });
  const saved = await pkg.save();
  assert.deepEqual(await packageProblems(saved), [], 'the package is whole: every part declared, every relationship leads somewhere');
  const again = await open(saved);
  assert.deepEqual(await readCollabStamp(again.package!), { ep: 'room-1_A', rev: 12 }, 'read back as it was written');
  const xml = decoder.decode(await again.package!.readPart(CUSTOM));
  const root = parseXml(xml).root;
  assert.equal(root.attrs.get('xmlns'), 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties', 'the part is one of custom properties');
  assert.match(xml, /<property fmtid="\{D5CDD505-2E9C-101B-9397-08002B2CF9AE\}" pid="2" name="BymOfficeCollab"><vt:lpwstr>room-1_A:12<\/vt:lpwstr><\/property>/, 'one property, a text');
  assert.equal(again.package!.contentTypeOf(CUSTOM), 'application/vnd.openxmlformats-officedocument.custom-properties+xml', 'declared as what it is');
  const rels = await again.package!.relationships('/');
  assert.equal(rels.filter((rel) => rel.part === CUSTOM && rel.type.endsWith('/custom-properties')).length, 1, 'and the package leads to it once');
  assert.equal(new Set(rels.map((rel) => rel.id)).size, rels.length, 'under an id of its own');
  // The workbook is what it was.
  assert.equal((await again.workbook()).sheets.length, 2, 'the workbook still opens');
  await writeCollabStamp(again.package!, { ep: 'room-2', rev: 0 });
  assert.deepEqual(await readCollabStamp(again.package!), { ep: 'room-2', rev: 0 }, 'a later save says its own room and step');
  assert.equal(decoder.decode(await again.package!.readPart(CUSTOM)).split('BymOfficeCollab').length, 2, 'in the one property');
  await writeCollabStamp(again.package!, undefined);
  assert.equal(await readCollabStamp(again.package!), undefined, 'and a file that is of no room any more says nothing');
  await assert.rejects(writeCollabStamp(again.package!, { ep: 'no spaces', rev: 1 }), RangeError, 'what is no name of a room is not written');
});

test('the properties a file has stay as they are; what is no stamp is none', async () => {
  const file = await open(await startFile());
  const pkg = file.package!;
  await writeCollabStamp(pkg, { ep: 'e', rev: 1 });
  // A file as an application writes it: properties of its own, a prefix on every name, the stamp of someone's making.
  const theirs = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:Properties xmlns:cp="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><cp:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Department"><vt:lpwstr>Fleet &amp; Crew</vt:lpwstr></cp:property><!-- kept --><cp:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Checked"><vt:bool>true</vt:bool></cp:property></cp:Properties>';
  pkg.writePart(CUSTOM, new TextEncoder().encode(theirs));
  assert.equal(await readCollabStamp(pkg), undefined, 'no stamp among them');
  await writeCollabStamp(pkg, { ep: 'room', rev: 5 });
  const xml = decoder.decode(await pkg.readPart(CUSTOM));
  assert.ok(xml.indexOf('<cp:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="Department"><vt:lpwstr>Fleet &amp; Crew</vt:lpwstr></cp:property><!-- kept --><cp:property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Checked"><vt:bool>true</vt:bool></cp:property>') >= 0, `what was there is there, as it was: ${xml}`);
  assert.match(xml, /<cp:property [^>]*pid="2" name="BymOfficeCollab"><vt:lpwstr>room:5<\/vt:lpwstr><\/cp:property><\/cp:Properties>$/, 'the stamp comes last, under the least number nobody has, in the words of the part');
  assert.deepEqual(await readCollabStamp(pkg), { ep: 'room', rev: 5 }, 'and is read');
  await writeCollabStamp(pkg, undefined);
  assert.equal(decoder.decode(await pkg.readPart(CUSTOM)).indexOf('BymOfficeCollab'), -1, 'taken out again, it leaves the others');
  assert.ok(decoder.decode(await pkg.readPart(CUSTOM)).indexOf('name="Checked"') > 0, 'which are still there');
  // What somebody wrote under the name of the stamp and is none.
  for (const text of ['room', 'room:', 'room:-1', 'room:1.5', 'ro om:1', ':3', `${'x'.repeat(65)}:1`, 'room:99999999999999999999']) {
    pkg.writePart(CUSTOM, new TextEncoder().encode(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="BymOfficeCollab"><vt:lpwstr>${text}</vt:lpwstr></property></Properties>`));
    assert.equal(await readCollabStamp(pkg), undefined, `"${text.slice(0, 20)}" is no stamp`);
  }
  pkg.writePart(CUSTOM, new TextEncoder().encode('<Properties'));
  assert.equal(await readCollabStamp(pkg), undefined, 'a part that is no XML holds none, and nothing is thrown');
});

test('a file that is written while something is typed into the workbook says no step it may not be of', async () => {
  const file = await open(await startFile());
  const pkg = file.package!;
  let touched = 3;
  let paused = 0;
  const session = {
    state: 'online',
    savableRev: 7 as number | undefined,
    stampNow: { ep: 'room', rev: 7 } as { ep: string; rev: number } | undefined,
    whenSettled: async () => undefined,
    pause: () => void paused++,
    resume: () => touched-- > 0,
  };
  let takes = 0;
  const take = async (): Promise<Uint8Array> => (takes++, pkg.save());
  const twice = await writeSharedFile(session, pkg, take);
  assert.deepEqual([takes, paused, twice.rev], [3, 3, undefined], 'tried again while something was typed, and in the end written as the file of no step');
  assert.equal(await readCollabStamp((await open(twice.bytes)).package!), undefined, 'which says nothing of itself');
  const calm = await writeSharedFile(session, pkg, take);
  assert.deepEqual([takes, calm.rev], [4, 7], 'a file written in peace is the file of the step');
  assert.deepEqual(await readCollabStamp((await open(calm.bytes)).package!), { ep: 'room', rev: 7 }, 'and says so');
  session.state = 'closed';
  const alone = await writeSharedFile(session, pkg, take);
  assert.equal(await readCollabStamp((await open(alone.bytes)).package!), undefined, 'a workbook that left its room for good writes a file of no room');
});

test('a part of custom properties that cannot be read is not touched, and the file is saved without a stamp; two stamps are none', async () => {
  const head = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
  const property = (pid: string, name: string, text: string): string => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${name}"><vt:lpwstr>${text}</vt:lpwstr></property>`;
  for (const [name, broken] of [
    ['cut short', '<Properties xmlns="x"><property'],
    ['with a document type', `<!DOCTYPE Properties [<!ENTITY e "x">]>${head}${property('2', 'BymOfficeCollab', 'room:1')}</Properties>`],
  ] as [string, string][]) {
    const file = await open(await startFile());
    const pkg = file.package!;
    assert.equal(await writeCollabStamp(pkg, { ep: 'first', rev: 1 }), true, 'a stamp is written');
    const bytes = new TextEncoder().encode(broken);
    pkg.writePart(CUSTOM, bytes);
    assert.equal(await readCollabStamp(pkg), undefined, `${name}: no stamp is read from it`);
    assert.equal(await writeCollabStamp(pkg, { ep: 'room', rev: 2 }), false, `${name}: no stamp can be written, and that is answered`);
    assert.equal(await writeCollabStamp(pkg, undefined), true, `${name}: taking out a stamp that cannot be there is done`);
    assert.deepEqual(Array.from(await pkg.readPart(CUSTOM)), Array.from(bytes), `${name}: not a byte of the part was touched`);
    // In a room the file is saved all the same, as the file of no step, and it is said that it carries no stamp.
    let takes = 0;
    const session = { state: 'online', savableRev: 2, stampNow: { ep: 'room', rev: 2 }, whenSettled: async () => undefined, pause: () => undefined, resume: () => false };
    const done = await writeSharedFile(session, pkg, async () => (takes++, pkg.save()));
    assert.deepEqual([takes, done.rev, done.unstamped], [1, undefined, true], `${name}: saved once, of no step, and told`);
    assert.deepEqual(Array.from(await (await open(done.bytes)).package!.readPart(CUSTOM)), Array.from(bytes), `${name}: the saved file has the part as it was`);
  }
  // Two properties of the name: none is the stamp; writing leaves one, taking out leaves none; the others stay.
  const file = await open(await startFile());
  const pkg = file.package!;
  await writeCollabStamp(pkg, { ep: 'first', rev: 1 });
  pkg.writePart(CUSTOM, new TextEncoder().encode(`${head}${property('2', 'BymOfficeCollab', 'a:1')}${property('3', 'Kept', 'yes')}${property('2147483647', 'BymOfficeCollab', 'b:2')}</Properties>`));
  assert.equal(await readCollabStamp(pkg), undefined, 'two stamps say nothing');
  assert.equal(await writeCollabStamp(pkg, { ep: 'room', rev: 9 }), true, 'a stamp is written over them');
  let xml = decoder.decode(await pkg.readPart(CUSTOM));
  assert.deepEqual([xml.split('BymOfficeCollab').length - 1, xml.indexOf('name="Kept"') > 0], [1, true], 'one stamp is left, and the other property');
  assert.deepEqual(await readCollabStamp(pkg), { ep: 'room', rev: 9 }, 'which is read');
  pkg.writePart(CUSTOM, new TextEncoder().encode(`${head}${property('2', 'BymOfficeCollab', 'a:1')}${property('3', 'Kept', 'yes')}${property('4', 'BymOfficeCollab', 'b:2')}</Properties>`));
  await writeCollabStamp(pkg, undefined);
  xml = decoder.decode(await pkg.readPart(CUSTOM));
  assert.deepEqual([xml.indexOf('BymOfficeCollab'), xml.indexOf('name="Kept"') > 0], [-1, true], 'taken out, both go and the other stays');
  // The largest number a property can have is taken: the stamp gets the least free one, not one more than it.
  pkg.writePart(CUSTOM, new TextEncoder().encode(`${head}${property('2', 'A', 'x')}${property('2147483647', 'B', 'y')}</Properties>`));
  await writeCollabStamp(pkg, { ep: 'room', rev: 1 });
  assert.match(decoder.decode(await pkg.readPart(CUSTOM)), /pid="3" name="BymOfficeCollab"/, 'the least number from 2 that nobody has');
});

test('two files are the same file when their parts are, however they were packed', async () => {
  const { ZipArchive } = await import('../src/index');
  const { samePackage } = await import('../src/formats/xlsx/collab-stamp');
  const bytes = await startFile();
  const archive = ZipArchive.read(bytes);
  // The same parts, written in the opposite order (and so as other bytes).
  const turned = new ZipArchive();
  const names = archive.names();
  const contents: [string, Uint8Array][] = [];
  for (const name of names) contents.push([name, await archive.read(name)]);
  for (const [name, content] of contents.slice().reverse()) turned.set(name, content);
  const other = await turned.write();
  assert.notEqual(Buffer.compare(Buffer.from(other), Buffer.from(bytes)), 0, 'the two archives are other bytes (or the test tests nothing)');
  assert.equal(await samePackage(bytes, other), true, 'the same parts with the same content are the same file');
  // One byte of one part, one part more, one part less, and what is no archive.
  const changed = ZipArchive.read(bytes);
  const sheet = names.find((name) => /sheet1\.xml$/.test(name))!;
  const content = (await changed.read(sheet)).slice();
  content[content.length - 3] ^= 1;
  changed.set(sheet, content);
  assert.equal(await samePackage(bytes, await changed.write()), false, 'one byte of one part is another file');
  const more = ZipArchive.read(bytes);
  more.set('extra.bin', new Uint8Array([1]));
  assert.equal(await samePackage(bytes, await more.write()), false, 'a part more is another file');
  assert.equal(await samePackage(await more.write(), bytes), false, 'either way round');
  assert.equal(await samePackage(bytes, new Uint8Array([1, 2, 3])), false, 'what is no archive is no file');
  // A file that unpacks to far more than it is: turned away by what its directory says, with nothing unpacked.
  const bomb = ZipArchive.read(bytes);
  bomb.set(sheet, new Uint8Array(24 * 1024 * 1024));
  const small = await bomb.write();
  assert.ok(small.length < bytes.length + 64 * 1024, `the file is small (${small.length} bytes) and its part is 24 MiB`);
  const reads: string[] = [];
  const read = ZipArchive.prototype.read;
  ZipArchive.prototype.read = function (this: InstanceType<typeof ZipArchive>, name: string) {
    reads.push(name);
    return read.call(this, name);
  };
  try {
    assert.equal(await samePackage(small, bytes), false, 'another file');
    assert.equal(await samePackage(bytes, small), false, 'either way round');
    assert.deepEqual(reads, [], 'and not one part was unpacked to find that out');
    assert.equal(await samePackage(bytes, await more.write()), false, 'nor for a file with other parts');
    assert.deepEqual(reads, [], 'whose names say so');
    assert.equal(await samePackage(bytes, other), true, 'the same file is unpacked and compared');
    assert.equal(reads.length, names.length * 2, 'every part of both, once');
  } finally {
    ZipArchive.prototype.read = read;
  }
});

test('save after save, the file one participant stores is the file another would write at that step - and writing one to compare changes nothing of the package', async () => {
  const { Net, act, join, saveOf, seeded } = await import('./collab-fixtures');
  const { proofFile, samePackage } = await import('../src/formats/xlsx/collab-stamp');
  const { packageTurn } = await import('../src/package-turn');
  const { OpcPackage } = await import('../src/index');
  const { setCellInput } = await import('../src/ops/commands');
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(1);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  for (let round = 0; round < 4; round++) {
    const who = round % 2 === 0 ? b : a;
    act(who, () => setCellInput(who.wb, 0, 12 + round, 0, `typed in round ${round}`));
    net.drain(random);
    // The one saves (its package keeps the stamp of every save); the other writes its file to compare.
    const stored = await saveOf(a);
    const pkg = b.file.package!;
    const before = [pkg.partNames().slice().sort().join(), await readCollabStamp(pkg)];
    const turn = packageTurn(b.file);
    const own = await proofFile(b.session, () => turn.save(), (given) => OpcPackage.open(given), turn.locked);
    assert.equal(own.rev, round + 1, `round ${round}: the file of the step`);
    assert.equal(await samePackage(stored, own.bytes), true, `round ${round}: the same file, part for part`);
    assert.deepEqual([pkg.partNames().slice().sort().join(), await readCollabStamp(pkg)], before, `round ${round}: the package has the parts and the stamp it had (none: this participant never saved)`);
    // In round 2 this participant saves as well: from then on its package carries a stamp of its own, and the proof still holds.
    if (round === 1) await saveOf(b);
  }
});
