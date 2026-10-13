// A room that the relay forgot and started anew is another room, whatever its steps are
// numbered; and a file is of the step its room wrote into it, whatever anybody says of it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { setCellInput } from '../src/ops/commands';
import { contentVersion } from '../src/collab/protocol';
import { PROTOCOL_VERSION } from '../src/collab/session';
import { Net, act, assertSame, canon, int, join, saveOf, seeded, settleAll, startFile, type Member } from './collab-fixtures';

const forget = (net: Net): void => (net.server as unknown as { roomsByKey: Map<string, unknown> }).roomsByKey.clear();

/** Everybody who is online has the same workbook; the others know that they are not in the room. */
function together(members: Member[], label: string): void {
  const online = members.filter((member) => member.session.state === 'online');
  if (online.length > 1) assertSame(online, label);
  for (const member of members) {
    if (member.session.state === 'online') continue;
    assert.ok(member.notices.length > 0, `${label}: ${member.name} is ${member.session.state} and was told so`);
  }
}

test('who comes back to a room the relay forgot does not take the steps of the new room for the rest of the old one', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(1);
  let stored = bytes;
  const dirty = { a: true, b: true };
  const a = await join(net, 'a', bytes, {}, async () => stored, () => dirty.a);
  const b = await join(net, 'b', bytes, {}, async () => stored, () => dirty.b);
  net.drain(random);
  for (let i = 0; i < 3; i++) act(a, () => setCellInput(a.wb, 0, 12 + i, 0, `old ${i}`));
  net.drain(random);
  stored = await saveOf(a);
  a.session.announceSaved([contentVersion(stored)], a.session.savableRev);
  net.drain(random);
  dirty.a = dirty.b = false;
  assertSame([a, b], 'before the room is forgotten');
  // The relay starts anew. One comes back, and changes four cells.
  a.transport.suspend();
  b.transport.suspend();
  net.drain(random);
  forget(net);
  a.transport.resume();
  await settleAll(net, random, [a, b]);
  assert.equal(a.session.state, 'online', 'the first one is in the new room');
  for (let i = 0; i < 4; i++) act(a, () => setCellInput(a.wb, 0, 16 + i, 1, `new ${i}`));
  await settleAll(net, random, [a, b]);
  // The other comes back, at "step 3" of a room that is no more.
  b.transport.resume();
  await settleAll(net, random, [a, b]);
  together([a, b], 'after both came back');
  assert.equal(b.session.state, 'online', 'the second one is in the new room as well: it had nothing unsaved');
  assert.equal(canon(b.wb), canon(a.wb), 'with everything the first one changed');
});

test('who comes back at step 0 with an older file to a room that was started with a newer one is not let in with it', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(2);
  let stored = bytes;
  const a = await join(net, 'a', bytes, {}, async () => stored, () => false);
  const b = await join(net, 'b', bytes, {}, async () => stored, () => false);
  net.drain(random);
  a.transport.suspend();
  b.transport.suspend();
  net.drain(random);
  forget(net);
  // Someone else changed the file meanwhile; the first one back loads it and starts the room with it.
  const c = await join(new Net(), 'c', bytes);
  act(c, () => setCellInput(c.wb, 0, 14, 0, 'changed elsewhere'));
  stored = await saveOf(c);
  a.session.reload();
  await settleAll(net, random, [a, b]);
  a.transport.resume();
  await settleAll(net, random, [a, b]);
  act(a, () => setCellInput(a.wb, 0, 15, 0, 'in the new room'));
  await settleAll(net, random, [a, b]);
  b.transport.resume();
  await settleAll(net, random, [a, b]);
  together([a, b], 'after both came back');
  assert.equal(canon(b.wb), canon(a.wb), 'the second one has the file the room was started with, and what was changed since');
});

test('a file is of the step that was written into it: who says otherwise beforehand moves nobody to another step', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(3);
  let stored = bytes;
  const a = await join(net, 'a', bytes, {}, async () => stored, () => false);
  net.drain(random);
  // Someone who may write, and lies.
  const raw = net.server.connect();
  raw.onMessage(() => undefined);
  raw.onState!((state) => {
    if (state === 'open') raw.send(JSON.stringify({ t: 'join', pv: PROTOCOL_VERSION, v: [contentVersion(bytes)], key: 'evil' }));
  });
  net.drain(random);
  // Steps that are not the same when they are made twice.
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'first'));
  for (let i = 0; i < 2; i++) act(a, () => [{ type: 'insert', sheet: 0, axis: 'row', at: 3, count: 1 }]);
  net.drain(random);
  // Whoever has the workbook writes the same bytes: the liar knows what the file will be known by, and says "step 1" first.
  stored = await saveOf(a);
  raw.send(JSON.stringify({ t: 'saved', v: [contentVersion(stored)], rev: 1 }));
  net.drain(random);
  a.session.announceSaved([contentVersion(stored)], a.session.savableRev);
  net.drain(random);
  const c = await join(net, 'c', stored, {}, async () => stored, () => false);
  await settleAll(net, random, [a, c]);
  act(a, () => setCellInput(a.wb, 0, 12, 0, '=A14+A15'));
  await settleAll(net, random, [a, c]);
  together([a, c], 'after the newcomer joined');
  assert.equal(c.session.state, 'online', 'the newcomer is in');
  assert.equal(canon(c.wb), canon(a.wb), 'and no step was applied to it twice');
});

test('two who were detached save one after the other and come back: nobody is online with a workbook of their own', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(4);
  let stored = bytes;
  const a = await join(net, 'a', bytes, {}, async () => stored, () => true);
  const b = await join(net, 'b', bytes, {}, async () => stored, () => true);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'of a'));
  act(b, () => setCellInput(b.wb, 0, 13, 0, 'of b'));
  net.drain(random);
  a.transport.suspend();
  b.transport.suspend();
  net.drain(random);
  forget(net);
  a.transport.resume();
  b.transport.resume();
  await settleAll(net, random, [a, b]);
  assert.deepEqual([a.session.state, b.session.state], ['detached', 'detached'], 'both keep their unsaved work apart from the new room');
  for (const member of [a, b]) {
    stored = await saveOf(member);
    member.session.rejoin([contentVersion(stored)]);
    await settleAll(net, random, [a, b]);
  }
  together([a, b], 'after both saved and came back');
});

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));

test(`rooms that are forgotten at any moment: in the end everybody is in one room with one workbook, and no file was loaded over unsaved work unasked (${SEEDS} runs)`, async () => {
  const bytes = await startFile();
  let choices = 0;
  for (let seed = Number(process.env.COLLAB_FIRST) || 1; seed <= SEEDS; seed++) {
    // Time does not pass in this test: the relay asks for a save whenever someone cannot join.
    const net = new Net({ syncPerMinute: 100000 } as object);
    const random = seeded(seed);
    // The file where it is stored, and how often it was written: what a save is conditional on (as an ETag is).
    const store = { bytes, written: 0 };
    const members: Member[] = [];
    const saved = new Map<Member, string>();
    const base = new Map<Member, number>();
    const lost: string[] = [];
    let timers: (() => void)[] = [];
    const away = new Set<Member>();
    const settle = async (): Promise<void> => {
      for (let round = 0; round < 40; round++) {
        await settleAll(net, random, members);
        await chain;
        // Time passes only while everybody is connected: who waits for someone to save does not give up on someone who is away for a moment.
        if (away.size > 0 || (timers.length === 0 && net.waiting === 0)) return;
        const due = timers;
        timers = [];
        for (const run of due) run();
      }
    };
    /** Saves when the file is still the one this participant has; false on a conflict. */
    let chain: Promise<unknown> = Promise.resolve();
    const save = (member: Member): Promise<boolean> => {
      const next = chain.then(() => write(member));
      chain = next;
      return next;
    };
    const write = async (member: Member): Promise<boolean> => {
      if (member.session.state !== 'detached' && member.session.savableRev === undefined) return true;
      if (base.get(member) !== store.written) return false;
      const detached = member.session.state === 'detached';
      const rev = member.session.savableRev;
      store.bytes = await saveOf(member);
      store.written++;
      base.set(member, store.written);
      saved.set(member, canon(member.wb));
      if (detached) member.session.rejoin([contentVersion(store.bytes)]);
      else member.session.announceSaved([contentVersion(store.bytes)], rev);
      return true;
    };
    for (let i = 0; i < 2 + (seed % 2); i++) {
      const member: Member = await join(
        net,
        `m${i}`,
        bytes,
        {
          timer: (run) => {
            timers.push(run);
            return () => {
              timers = timers.filter((other) => other !== run);
            };
          },
          // Someone cannot join: the one who saves for the others does, when the file is still the one it has.
          onSync: () => {
            if (!member.session.leads || member.session.savableRev === undefined) return;
            // The file was written by someone who is not in the room (a fork): the user is asked, and says "write over it".
            void save(member).then((done) => {
              if (done) return undefined;
              choices++;
              base.set(member, store.written);
              return save(member);
            });
          },
          // What someone else saved is what this participant has, when it is at that step with nothing waiting.
          onSaved: (told) => {
            if (!told.clean) return;
            saved.set(member, canon(member.wb));
            base.set(member, store.written);
          },
        },
        async () => {
          // The workbook is given up for the file only when everything in it is in a file.
          if (saved.get(member) !== canon(member.wb)) lost.push(`seed ${seed}: ${member.name} loaded the file over unsaved work`);
          base.set(member, store.written);
          const next = store.bytes;
          queueMicrotask(() => saved.set(member, canon(member.wb)));
          return next;
        },
        () => saved.get(member) !== canon(member.wb),
      );
      saved.set(member, canon(member.wb));
      base.set(member, 0);
      members.push(member);
    }
    let typed = 0;
    for (let tick = 0; tick < 60; tick++) {
      const member = members[int(random, 0, members.length - 1)];
      const roll = random();
      if (roll < 0.5) {
        const text = `t${typed++}`;
        if (member.session.state !== 'closed' && member.session.state !== 'connecting') act(member, () => setCellInput(member.wb, 0, int(random, 10, 20), int(random, 0, 4), text));
      } else if (roll < 0.62) {
        if (member.session.state === 'detached' || member.session.savableRev !== undefined) await save(member);
      } else if (roll < 0.7) {
        member.transport.suspend();
        away.add(member);
      } else if (roll < 0.8) {
        member.transport.resume();
        away.delete(member);
      }
      else if (roll < 0.83) {
        for (const each of members) {
          each.transport.suspend();
          away.add(each);
        }
        net.drain(random);
        forget(net);
      }
      if (random() < 0.7) await settle();
    }
    for (const member of members) member.transport.resume();
    away.clear();
    await settle();
    // Whoever is on their own saves and comes back with the file; on a conflict the user is asked, and here
    // answers "load what is stored": the one loss there is, and one that was chosen.
    for (let round = 0; round < 6; round++) {
      for (const member of members) {
        if (member.session.state !== 'detached') continue;
        if (await save(member)) continue;
        choices++;
        saved.set(member, canon(member.wb));
        member.session.reload();
      }
      await settle();
    }
    const label = `seed ${seed}: ${members.map((member) => `${member.session.state}(${member.notices.map((notice) => notice.code).join('+')})`).join(' ')}`;
    assert.deepEqual(lost, [], 'no file was loaded over work that was in no file');
    assert.deepEqual(members.map((member) => member.session.state), members.map(() => 'online'), `everybody is back in the room; ${label}`);
    for (const member of members) assert.equal(member.session.pending, 0, `${label}: ${member.name} has nothing on its way`);
    assertSame(members, label);
  }
  assert.ok(choices > 0, 'conflicts happened, and were asked about');
});

test('over a transport of the application (no way to connect anew) a file the room has not heard of leaves the room and joins again, and gets in once the room has saved', async () => {
  const { CollabSession } = await import('../src/collab/session');
  const { open } = await import('../src/index');
  const { readCollabStamp } = await import('../src/formats/xlsx/collab-stamp');
  const bytes = await startFile();
  const net = new Net({ syncPerMinute: 100000 } as object);
  const random = seeded(9);
  let stored = bytes;
  let a: Member;
  // eslint-disable-next-line prefer-const
  a = await join(net, 'a', bytes, {
    onSync: () => {
      void saveOf(a).then((saved) => {
        stored = saved;
        a.session.announceSaved([contentVersion(saved)], a.session.savableRev);
      });
    },
  });
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 1, 'one'));
  net.drain(random);
  // Saved with the room and the step in it, but the room was never told of the save (or did not write it down).
  stored = await saveOf(a);
  let file = await open(stored);
  let wb = await file.workbook();
  const inner = net.server.connect();
  inner.onState!(() => undefined);
  const sent: string[] = [];
  const got: string[] = [];
  const notices: string[] = [];
  // Only what the contract asks of a transport: send, onMessage, close.
  const plain = {
    send: (text: string) => {
      sent.push(JSON.parse(text).t);
      inner.send(text);
    },
    onMessage: (listener: (text: string) => void) =>
      inner.onMessage((text) => {
        const message = JSON.parse(text) as { t: string; code?: string };
        got.push(message.code ? `${message.t}:${message.code}` : message.t);
        listener(text);
      }),
    close: () => inner.close(),
  };
  let timers: (() => void)[] = [];
  const session = new CollabSession(
    {
      workbook: () => wb,
      changed: () => undefined,
      notice: (notice) => void notices.push(notice.code),
      unsaved: () => false,
      reload: async () => {
        file = await open(stored);
        wb = await file.workbook();
        return { versions: [contentVersion(stored)], stamp: await readCollabStamp(file.package!) };
      },
    },
    { transport: plain, versions: [contentVersion(stored)], stamp: await readCollabStamp(file.package!), timer: (run) => (timers.push(run), () => undefined) },
  );
  for (let round = 0; round < 60 && session.state !== 'online'; round++) {
    net.drain(random);
    await new Promise((resolve) => setImmediate(resolve));
    const due = timers;
    timers = [];
    for (const run of due) run();
  }
  assert.equal(got.indexOf('error:bad-message'), -1, `no second "join" from a connection that is in a room: sent ${sent.join()} | got ${got.join()}`);
  assert.deepEqual(sent.slice(0, 3), ['join', 'leave', 'join'], 'out of the room it does not trust, then in again without the word of the file');
  assert.equal(session.state, 'online', `in the end it is in the room: got ${got.join()}; said ${notices.join()}`);
  act(a, () => setCellInput(a.wb, 0, 13, 1, 'two'));
  net.drain(random);
  assert.equal(canon(wb), canon(a.wb), 'with the workbook of the room');
});
