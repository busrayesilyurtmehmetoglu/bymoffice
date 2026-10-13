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
  assert.match(xml, /<cp:property [^>]*pid="8" name="BymOfficeCollab"><vt:lpwstr>room:5<\/vt:lpwstr><\/cp:property><\/cp:Properties>$/, 'the stamp comes last, under a number nobody has, in the words of the part');
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
