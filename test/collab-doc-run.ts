// A document that is edited together against the reference model (collab-doc-model.ts): two to
// five participants type, delete, replace, press Enter, colour text and indent paragraphs, with
// long queues, connections that go away and messages delivered in any order. After every
// message and every change each participant has to see exactly what the model says it sees;
// after every step the relay placed, a participant that only follows has to hold exactly the
// text of that step; nothing is dropped that the model does not expect to be dropped, and
// what is dropped is said.
//
// Every character that is typed is one that was never typed before, so that the text itself
// says which character is which.

import type { DocWire, WirePos } from '../src/collab/doc/rebase';
import { DocNet, bodyText, canonDoc, charColors, docFile, firstDifference, joinDoc, seeded, type DocMember } from './collab-doc-fixtures';
import { Model, type Change, type Item, type Seat } from './collab-doc-model';

export interface Result {
  violations: string[];
  made: number;
  placed: number;
  dropped: number;
  quiet: number;
}

export interface Config {
  seed: number;
  members: number;
  ticks: number;
  /** How often, of all that happens, a message is delivered: the less, the longer the queues. */
  deliver: number;
  cuts: boolean;
  /** Only typing and deleting, as two people writing do. */
  plain?: boolean;
  trace?(line: string): void;
}

/** Where the place before the `index`-th item a participant sees is in its document. */
function placeOf(view: Item[], index: number): WirePos {
  let row = 0;
  let start = 0;
  for (let i = 0; i < index; i++) {
    if (view[i].ch === '\n') {
      row++;
      start = i + 1;
    }
  }
  return { s: 'body', p: [row], o: index - start };
}

export async function run(config: Config): Promise<Result> {
  const random = seeded(config.seed);
  const int = (low: number, high: number): number => low + Math.floor(random() * (high - low + 1));
  let nextChar = 0x4e00;
  let unique = 0;
  const chars = (count: number): string[] => {
    const out: string[] = [];
    for (let i = 0; i < count; i++) out.push(String.fromCharCode(nextChar++));
    return out;
  };
  const start = [chars(6).join(''), chars(8).join(''), chars(5).join('')];
  const bytes = await docFile(start);
  const model = new Model(start.join('\n'));
  const net = new DocNet();
  const violations: string[] = [];
  const fail = (text: string): void => {
    if (violations.length < 4) violations.push(text);
  };
  const seats = new Map<string, Seat>();
  const ids = new Map<string, Seat>();

  const members: DocMember[] = [];
  // The one that only follows: it gets every message at once, and so holds the text of every step right after the relay placed it.
  let witness: DocMember | undefined;
  witness = await joinDoc(net, 'witness', bytes, {
    instant: true,
    readOnly: true,
    before: (message) => {
      if (message.t !== 'op' || typeof message.rev !== 'number' || message.rev <= model.rev) return;
      // A connection that was made anew has a new name.
      if (!ids.has(String(message.from))) for (const member of members) if (member.session.user.id !== '') ids.set(member.session.user.id, seats.get(member.name) as Seat);
      const author = ids.get(String(message.from));
      if (!author) return fail(`step ${message.rev} is of ${message.from}, whom nobody knows`);
      if (config.trace) config.trace(`step ${message.rev} is of ${author.name}: ${JSON.stringify((message as { ops?: unknown }).ops)} on ${JSON.stringify(model.textAt(model.rev))}`);
      if (config.trace) config.trace(`  model: ${JSON.stringify(author.waiting[0], (key, value) => (key === 'items' ? value.map((item: { ch: string }) => item.ch).join('') : value))} view ${JSON.stringify(model.text(author))}`);
      try {
        model.placed(author, message.rev);
      } catch (error) {
        fail(String(error));
      }
    },
    after: (message) => {
      if (message.t !== 'op' || !witness || message.rev !== model.rev) return;
      const text = bodyText(witness.doc);
      const expected = model.textAt(model.rev);
      if (text !== expected) fail(`after step ${model.rev} the text is ${JSON.stringify(text)}, and the model says ${JSON.stringify(expected)}`);
      const indent = model.indents.get(model.rev);
      if (indent) {
        const block = witness.doc.body.blocks[indent.row];
        const left = block && block.type === 'paragraph' && block.props.indent ? block.props.indent.left : undefined;
        if (left !== indent.value) fail(`step ${model.rev} was to indent paragraph ${indent.row} by ${indent.value}; it has ${left}`);
      }
    },
  });

  for (let i = 0; i < config.members; i++) {
    const name = `m${i}`;
    const seat = model.seat(name);
    seats.set(name, seat);
    const member: DocMember = await joinDoc(net, name, bytes, {
      before: (message) => {
        if (message.t !== 'op' || message.rev !== seat.rev + 1 || member.session.rev !== seat.rev) return;
        if (config.trace) config.trace(`${name} reads step ${message.rev}`);
        try {
          model.read(seat, message.rev);
        } catch (error) {
          fail(String(error));
        }
      },
    });
    members.push(member);
  }

  const check = (when: string): void => {
    for (const member of members) {
      if (member.session.user.id !== '') ids.set(member.session.user.id, seats.get(member.name) as Seat);
      const seat = seats.get(member.name) as Seat;
      if (member.session.rev !== seat.rev) fail(`${when}: ${member.name} holds ${member.session.rev} steps, the model ${seat.rev}`);
      const text = bodyText(member.doc);
      const expected = model.text(seat);
      if (text !== expected) fail(`${when}: ${member.name} sees ${JSON.stringify(text)}, and the model says ${JSON.stringify(expected)}`);
      if (member.session.stats.quiet !== seat.quiet) fail(`${when}: ${member.name} let ${member.session.stats.quiet} changes go quietly, the model ${seat.quiet}`);
      if (member.session.stats.dropped !== seat.dropped) fail(`${when}: ${member.name} dropped ${member.session.stats.dropped} changes, the model ${seat.dropped}`);
      const said = member.notices.filter((notice) => notice.code === 'dropped').reduce((sum, notice) => sum + (notice.count || 0), 0);
      if (said !== member.session.stats.dropped) fail(`${when}: ${member.name} dropped ${member.session.stats.dropped} changes and said ${said}`);
      const cut = member.notices.filter((notice) => notice.code === 'partial').reduce((sum, notice) => sum + (notice.count || 0), 0);
      if (cut !== seat.partial) fail(`${when}: ${member.name} was told ${cut} times that part of a change went with a deletion of someone else, the model says ${seat.partial}`);
      const odd = member.notices.find((notice) => notice.code !== 'dropped' && notice.code !== 'backlog' && notice.code !== 'partial');
      if (odd) fail(`${when}: ${member.name} was told "${odd.code}"`);
    }
  };

  const act = (member: DocMember): void => {
    const seat = seats.get(member.name) as Seat;
    const view = model.seen(seat);
    const roll = random();
    const gap = int(0, view.length);
    let wires: DocWire[];
    let change: Change;
    const some = (): { from: number; to: number } | undefined => {
      if (view.length === 0) return undefined;
      const from = int(0, view.length - 1);
      return { from, to: Math.min(view.length, from + int(1, 4)) };
    };
    const typed = (): string[] => {
      const text = chars(int(1, 3));
      // Now and then the end of a paragraph goes in with the text, as a paste does it.
      if (random() < 0.15) text.splice(int(0, text.length), 0, '\n');
      return text;
    };
    if (config.plain ? roll < 0.72 : roll < 0.36) {
      const text = typed();
      wires = [{ t: 'text', at: placeOf(view, gap), x: text.join('') }];
      change = { kind: 'ins', index: gap, items: model.fresh(text) };
    } else if (config.plain ? roll < 0.8 : roll < 0.46) {
      wires = [{ t: 'split', at: placeOf(view, gap) }];
      change = { kind: 'ins', index: gap, items: model.fresh(['\n']) };
    } else if (config.plain || roll < 0.7) {
      const range = some();
      if (!range) return;
      wires = [{ t: 'del', at: placeOf(view, range.from), to: placeOf(view, range.to) }];
      change = { kind: 'del', victims: view.slice(range.from, range.to).map((item) => item.id) };
    } else if (roll < 0.78) {
      const range = some();
      if (!range) return;
      const text = typed();
      wires = [{ t: 'text', at: placeOf(view, range.from), to: placeOf(view, range.to), x: text.join('') }];
      change = { kind: 'ins', index: range.from, items: model.fresh(text), victims: view.slice(range.from, range.to).map((item) => item.id) };
    } else if (roll < 0.9) {
      const range = some();
      if (!range) return;
      // A colour nobody used before: a change that sets what is set already is no change, and no step.
      const color = (0x100000 + ++unique * 7).toString(16).toUpperCase();
      wires = [{ t: 'run', at: placeOf(view, range.from), to: placeOf(view, range.to), f: { color } }];
      change = { kind: 'color', victims: view.slice(range.from, range.to).map((item) => item.id), color };
    } else {
      const value = ++unique;
      const place = placeOf(view, gap);
      wires = [{ t: 'para', at: place, to: place, f: { indent: { left: value } } }];
      change = { kind: 'indent', index: gap, value };
    }
    if (config.trace) config.trace(`${member.name} on ${JSON.stringify(model.text(seat))}: ${JSON.stringify(wires)}`);
    try {
      member.session.submit(wires);
    } catch (error) {
      return fail(`${member.name} could not make ${JSON.stringify(wires)}: ${String(error)}`);
    }
    model.made(seat, change);
  };

  net.drain(random);
  check('at the start');
  for (let tick = 0; tick < config.ticks && violations.length === 0; tick++) {
    const roll = random();
    if (roll < config.deliver) net.step(random);
    else if (config.cuts && roll < config.deliver + 0.03) members[int(0, members.length - 1)].transport.suspend();
    else if (config.cuts && roll < config.deliver + 0.07) members[int(0, members.length - 1)].transport.resume();
    else {
      const member = members[int(0, members.length - 1)];
      // A burst now and then: several changes before anything arrives.
      const burst = random() < 0.2 ? int(2, 6) : 1;
      for (let i = 0; i < burst && violations.length === 0; i++) act(member);
    }
    check(`tick ${tick}`);
  }
  if (violations.length === 0) {
    for (const member of members) member.transport.resume();
    for (let round = 0; round < 50; round++) {
      net.drain(random);
      await new Promise<void>((resolve) => setImmediate(resolve));
      if (net.waiting === 0) break;
    }
    check('at the end');
    const all = members.concat([witness]);
    const text = model.textAt(model.rev);
    for (const member of all) {
      if (member !== witness && member.session.pending !== 0) fail(`${member.name} still has ${member.session.pending} changes waiting`);
      if (bodyText(member.doc) !== text) fail(`at the end ${member.name} holds ${JSON.stringify(bodyText(member.doc))}, the model ${JSON.stringify(text)}`);
    }
    const canon = all.map((member) => canonDoc(member.doc));
    for (let i = 1; i < canon.length; i++) if (canon[i] !== canon[0]) fail(`${all[i].name} and ${all[0].name} hold different documents ${firstDifference(canon[0], canon[i])}`);
    const order = model.inOrder();
    if (order) fail(order);
    const colors = charColors(witness.doc);
    const expected = model.colors();
    if (colors.length !== expected.length) fail(`${colors.length} characters have a colour to look at, the model has ${expected.length}`);
    else for (let i = 0; i < expected.length; i++) if (expected[i] !== '' && colors[i] !== expected[i]) fail(`character ${i} was to be ${expected[i]} and is ${colors[i]}`);
  }
  const result: Result = { violations, made: 0, placed: 0, dropped: 0, quiet: 0 };
  for (const member of members) {
    const stats = member.session.stats;
    result.made += stats.made;
    result.placed += stats.placed;
    result.dropped += stats.dropped;
    result.quiet += stats.quiet;
    if (violations.length === 0 && stats.made !== stats.placed + stats.dropped + stats.quiet) violations.push(`${member.name} made ${stats.made} changes: ${stats.placed} placed, ${stats.dropped} dropped, ${stats.quiet} let go`);
    member.session.leave();
  }
  witness.session.leave();
  return result;
}

