// Who is in the file: the list a page shows, from what the session knows. This participant
// first, the others as the relay named them; nobody who has left; names as text that can do no harm.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_PEOPLE, initials, listPeople, type CollabPerson } from '../src/collab/people';
import { cleanName, cleanUser, parseServerMessage } from '../src/collab/protocol';
import type { CollabPeer } from '../src/collab/session';
import { Net, join, seeded, settleAll, startFile, type Member } from './collab-fixtures';

const guest = (n: number): string => `Guest ${n}`;
const people = (member: Member): CollabPerson[] => listPeople(member.session.user, member.session.peers, member.session.state, guest).list;
const names = (member: Member): string[] => people(member).map((who) => `${who.name}${who.self ? '*' : ''}${who.readOnly ? ' (reads)' : ''}`);

test('the list: this participant first, then the others; whoever leaves is gone, however it left; out of the room it is this participant alone', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(1);
  const a = await join(net, 'Ada', bytes, { user: { name: 'Ada', color: '#D93025' } }, async () => bytes, () => false);
  net.drain(random);
  assert.deepEqual(names(a), ['Ada*'], 'alone in the file: one');
  const b = await join(net, 'Ben', bytes, { user: { name: 'Ben' } });
  net.drain(random);
  const c = await join(net, 'c', bytes, { user: {}, readOnly: true }, async () => bytes, () => false);
  net.drain(random);
  const d = await join(net, 'd', bytes, { user: {} });
  net.drain(random);
  const guestOf = (who: Member): string => people(who)[0].name;
  assert.match(guestOf(c), /^Guest [1-9][0-9]?$/, 'who gave no name is a guest with a number');
  assert.deepEqual(names(a), ['Ada*', 'Ben', `${guestOf(c)} (reads)`, guestOf(d)], 'the others in the order they came, who only reads marked');
  assert.deepEqual(names(b), ['Ben*', 'Ada', `${guestOf(c)} (reads)`, guestOf(d)], 'the same people for everybody, each first in its own list');
  assert.deepEqual(names(c)[0], `${guestOf(c)}* (reads)`, 'a guest has the name for itself that the others see');
  assert.equal(people(a)[0].color, '#D93025', 'the colour that was given');
  assert.equal(people(b)[1].color, '#D93025', 'is the colour the others see');
  assert.equal(new Set(people(a).map((who) => who.id)).size, 4, 'everybody is one row');
  // Where they are.
  b.session.setPresence({ sheet: 1, row: 2, col: 3, range: { top: 2, left: 3, bottom: 2, right: 3 } });
  (b.session as unknown as { flushPresence(): void }).flushPresence();
  net.drain(random);
  assert.equal(people(a)[1].sheet, 1, 'the sheet someone is on');
  assert.equal(people(a)[2].sheet, undefined, 'and none for who has not said');
  // One leaves the room, one loses its connection, one is taken away by the relay's page going.
  b.session.leave();
  net.drain(random);
  assert.deepEqual(names(a), ['Ada*', `${guestOf(c)} (reads)`, guestOf(d)], 'who left is not listed');
  assert.deepEqual(names(b), ['Ben*'], 'and lists itself alone');
  d.transport.suspend();
  net.drain(random);
  assert.deepEqual(names(a), ['Ada*', `${guestOf(c)} (reads)`], 'who lost the connection is not listed');
  assert.deepEqual(people(d).map((who) => who.self), [true], 'and, offline, lists itself alone');
  d.transport.resume();
  net.drain(random);
  assert.equal(names(a).length, 3, 'back, it is listed again, once');
  // The relay forgets everything (it was started anew): nobody is left over from before.
  for (const member of [a, c, d]) member.transport.suspend();
  net.drain(random);
  assert.deepEqual(names(a), ['Ada*'], 'offline: this participant alone');
  (net.server as unknown as { roomsByKey: Map<string, unknown> }).roomsByKey.clear();
  a.transport.resume();
  await settleAll(net, random, [a, c]);
  assert.deepEqual([a.session.state, names(a)], ['online', ['Ada*']], 'in the new room alone: no ghost of the room before it');
  c.transport.resume();
  await settleAll(net, random, [a, c]);
  assert.equal(names(a).length, 2, 'until someone comes');
});

test('the same connection twice, or this participant among the others, is one row; and there are no more rows than a page can show', () => {
  const peer = (id: string, name?: string): CollabPeer => ({ id, color: '#1A73E8', guest: 7, readOnly: false, ...(name === undefined ? {} : { name }) });
  const twice = listPeople(peer('me', 'Me'), [peer('x', 'X'), peer('x', 'X again'), peer('me', 'Impostor of me'), peer('y', 'X')], 'online', guest);
  assert.deepEqual(twice.list.map((who) => `${who.id}:${who.name}`), ['me:Me', 'x:X', 'y:X'], 'one row for each connection; two people of one name are two rows');
  assert.equal(twice.count, 3, 'and so is the count');
  const crowd = Array.from({ length: 5000 }, (_unused, index) => peer(`p${index}`, `P ${index}`));
  const many = listPeople(peer('me'), crowd, 'online', guest);
  assert.deepEqual([many.list.length, many.more, many.count], [MAX_PEOPLE, 5001 - MAX_PEOPLE, 5001], 'two hundred rows, and the rest as a number');
  assert.equal(many.list[0].name, 'Guest 7', 'who gave no name is a guest');
  for (const state of ['connecting', 'offline', 'detached', 'closed'] as const) assert.deepEqual(listPeople(peer('me'), crowd, state, guest).count, 1, `${state}: nobody else is known to be there`);
  assert.equal(listPeople({ ...peer('me'), color: 'red;background:url(x)' }, [], 'online', guest).list[0].color, '#5F6368', 'a colour that is none is not used');
});

test('names are text that is shown: no control or direction characters, no more than 64 characters, nothing cut in half', () => {
  assert.equal(cleanName('  Ada \u0000\n Lovelace\t'), 'Ada Lovelace', 'control characters are spaces, and spaces are one');
  assert.equal(cleanName('‮gnp.exe‬ ⁦x⁩‏​﻿'), 'gnp.exe x', 'what turns the text around or hides in it is taken out');
  assert.equal(cleanName('<img src=x onerror=alert(1)>'), '<img src=x onerror=alert(1)>', 'markup is text like any other (it is only ever shown as text)');
  assert.equal(Array.from(cleanName('😀'.repeat(100))).length, 64, 'sixty-four characters, counted as characters');
  assert.equal(cleanName('😀'.repeat(100)).length, 128, 'none of them cut in half');
  assert.equal(cleanName('x'.repeat(100000)).length, 64, 'however long it was');
  assert.equal(cleanName('a\uD83Db'), 'ab', 'half a character is none');
  assert.equal(cleanName('a \uD83D b'), 'a b', 'and leaves one space behind, not two');
  assert.equal(cleanName('A\u00AD\u034F\u115F\u1160\u180E\u2800\u3164\uFFA0\u206A\u206F\uFFF9\uFFFB\uFE0F\u{E0001}\u{E007F}\u{E0100}B'), 'AB', 'more of what cannot be seen is taken out');
  // Cleaning what was cleaned changes nothing.
  let state = 12345;
  const random = (): number => {
    state = (Math.imul(state, 1103515245) + 12345) >>> 0;
    return state / 4294967296;
  };
  const alphabet = ['a', 'B', ' ', '  ', '\t', '\u0000', '\u00AD', '\u200B', '\u202E', '\u2066', '\uFEFF', '\uD83D', '\uDE00', '😀', '\u{E0041}', '\uFE0F', '\u3164', 'é', '\u0301', '<', '\u2028'];
  for (let i = 0; i < 20000; i++) {
    let text = '';
    for (let k = Math.floor(random() * 90); k > 0; k--) text += alphabet[Math.floor(random() * alphabet.length)];
    const once = cleanName(text);
    if (cleanName(once) !== once) assert.fail(`cleaning twice is not cleaning once: ${JSON.stringify(text)} -> ${JSON.stringify(once)} -> ${JSON.stringify(cleanName(once))}`);
    if (/^\s|\s$|\s\s/.test(once) || Array.from(once).length > 64) assert.fail(`not clean: ${JSON.stringify(once)}`);
  }
  assert.equal(cleanName('‮​ '), '', 'a name of nothing is no name');
  assert.deepEqual(cleanUser({ name: '‮', color: '#abcdef' }), { color: '#ABCDEF' }, 'and its participant is a guest');
  assert.deepEqual([initials('Ada Lovelace'), initials('ben'), initials('😀 x'), initials('')], ['AL', 'B', '😀X', '?'], 'the letters in the small circle');
  // What the relay says of this participant in "hello" is cleaned like any other name.
  const hello = parseServerMessage(JSON.stringify({ t: 'hello', rev: 0, from: 0, ep: 'e', you: 'c1', ro: false, v: [], peers: [], user: { name: '‮Named by the relay  ', color: 'nope' } }));
  assert.deepEqual((hello as { user?: unknown }).user, { name: 'Named by the relay' }, 'cleaned');
});
