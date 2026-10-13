// What a program needs to edit a workbook together without mount() (in Node, or with a view of
// its own): the session, a relay, and the two functions that write a file with the step of the
// room in it and read that step back. Everything through the entry point of the package.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import * as B from '../src/index';

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 5));

interface Member {
  file: B.OfficeFile;
  workbook: B.Workbook;
  session: B.CollabSession;
}

async function join(server: B.MemoryServer, bytes: Uint8Array, name: string): Promise<Member> {
  const file = await B.open(bytes);
  const workbook = await file.workbook();
  const host: B.CollabSessionHost = { workbook: () => workbook, changed: () => undefined, notice: () => undefined };
  const stamp: B.CollabStamp | undefined = await B.readCollabStamp(file.package as B.OpcPackage);
  const options: B.CollabSessionOptions = { transport: server.connect(), versions: [B.contentVersion(bytes)], stamp, user: { name } };
  return { file, workbook, session: new B.CollabSession(host, options) };
}

async function settle(members: Member[]): Promise<void> {
  for (let i = 0; i < 400; i++) {
    await tick();
    if (members.every((m) => m.session.state === 'online' && m.session.pending === 0) && new Set(members.map((m) => m.session.rev)).size === 1) return;
  }
  assert.fail('the sessions did not settle');
}

test('a file is written with the step of its room and read back, without mount()', async () => {
  assert.equal(B.collab.readCollabStamp, B.readCollabStamp);
  assert.equal(B.collab.writeSharedFile, B.writeSharedFile);
  const base = await B.createBlank('xlsx', { sheetName: 'Shared' });
  const server = B.collab.memoryServer();
  const ann = await join(server, base, 'Ann');
  const bob = await join(server, base, 'Bob');
  try {
    await settle([ann, bob]);
    assert.equal(await B.readCollabStamp(ann.file.package as B.OpcPackage), undefined, 'a file no room wrote says nothing');
    ann.session.submit(B.setCellInput(ann.workbook, 0, 0, 0, '12'));
    bob.session.submit(B.setCellInput(bob.workbook, 0, 1, 0, '=A1*2'));
    await settle([ann, bob]);

    const saved: B.SharedFile = await B.writeSharedFile(ann.session, ann.file.package, () => ann.file.save());
    assert.equal(saved.unstamped, undefined);
    assert.equal(saved.rev, ann.session.rev);
    const again = await B.open(saved.bytes);
    const stamp = await B.collab.readCollabStamp(again.package as B.OpcPackage);
    assert.deepEqual(stamp, ann.session.stampNow);
    assert.equal(stamp && stamp.rev, saved.rev);
    assert.equal(B.getCell((await again.workbook()).sheets[0], 0, 0)!.value, 12);

    // Whoever opens that file joins at its step.
    ann.session.announceSaved([B.contentVersion(saved.bytes)], saved.rev);
    await tick();
    const cid = await join(server, saved.bytes, 'Cid');
    try {
      await settle([ann, bob, cid]);
      assert.equal(B.getCell(cid.workbook.sheets[0], 1, 0)!.formula!.text, 'A1*2');
    } finally {
      cid.session.leave();
    }

    // A session that left the room for good writes a file that is of no room.
    ann.session.leave();
    const alone = await B.writeSharedFile(ann.session, ann.file.package, () => ann.file.save());
    assert.equal(alone.rev, undefined);
    assert.equal(await B.readCollabStamp((await B.open(alone.bytes)).package as B.OpcPackage), undefined);
  } finally {
    ann.session.leave();
    bob.session.leave();
    server.close();
  }
});
