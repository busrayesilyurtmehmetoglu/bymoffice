// A reference model of a text that is edited together, written without looking at how the editor
// reads a waiting change through one that arrived (src/collab/doc) and without using the
// commands of the document (src/ops/doc): it imports nothing of either.
//
// Every character and every paragraph end that was ever put in is an item with a number of its
// own, and the text is one line of items (no paragraphs, no paths). A change is said as its
// author MEANT it: "these items before the n-th item I see", "these items go", "these items
// get this colour", "the paragraph of the place before the n-th item gets this indent". What
// a participant sees at any moment, and what the text has to be after each step the relay
// placed, follows from the order of the relay and three rules:
//
//   - Only the items a deletion or a colouring names are deleted or coloured, whatever came
//     to stand between them since.
//   - Items that were put where others were deleted since stand where those were.
//   - Of two changes that put items at the same place, the one that got its place in the
//     order LATER stands first.
//   - A waiting change that loses items it names to a deletion that got its place first is
//     told so, every time (the other one may think again, and the items come back unchanged:
//     a deletion of them did not happen here, a colouring of them did not).
//
// Beside that, whatever the rules, no two items may ever change places: that is checked at
// the end against what every author saw when it made its change (inOrder()).

export interface Item {
  id: number;
  /** One character; "\n" for the end of a paragraph. */
  ch: string;
  /** The step it came with (0: it was in the file), and the step it went with. */
  born: number;
  died: number;
  /** The colour it was given last, by a step that named it. */
  color?: string;
}

export type Change =
  /** Items before the `index`-th item its author sees (after `victims` went, when it replaces something). */
  | { kind: 'ins'; index: number; items: Item[]; victims?: number[] }
  | { kind: 'del'; victims: number[] }
  | { kind: 'color'; victims: number[]; color: string }
  /** The paragraph the place before the `index`-th item is in gets an indent. */
  | { kind: 'indent'; index: number; value: number };

export interface Seat {
  name: string;
  /** How many steps of the order this participant holds. */
  rev: number;
  waiting: Change[];
  /** What the model expects of the counters of its session. */
  quiet: number;
  dropped: number;
  /** How often a waiting deletion, colouring or replacement lost part (or all) of what it named to a deletion that got its place first: each time is said. */
  partial: number;
}

const NEVER = Number.MAX_SAFE_INTEGER;

export class Model {
  /** Every item there ever was, in its place; the ones that went stay as marks of where they were. */
  private all: Item[] = [];
  private nextId = 1;
  rev = 0;
  /** Who made each step, and what it was (as it got its place). */
  readonly steps: { seat: Seat; change: Change }[] = [];
  readonly seats: Seat[] = [];
  /** Per step that set an indent: the row (paragraph) it has to be on in the text of that step, and the value. */
  readonly indents = new Map<number, { row: number; value: number }>();
  /** What every author saw right after it made a change: the order of the items in it is never to change. */
  private sights: number[][] = [];

  constructor(text: string) {
    for (const ch of Array.from(text)) this.all.push({ id: this.nextId++, ch, born: 0, died: NEVER });
  }

  seat(name: string): Seat {
    const seat: Seat = { name, rev: this.rev, waiting: [], quiet: 0, dropped: 0, partial: 0 };
    this.seats.push(seat);
    return seat;
  }

  fresh(chars: string[]): Item[] {
    return chars.map((ch) => ({ id: this.nextId++, ch, born: NEVER, died: NEVER }));
  }

  /** The items of the text after the first `rev` steps of the order. */
  private at(rev: number): Item[] {
    return this.all.filter((item) => item.born <= rev && item.died > rev);
  }

  private static apply(list: Item[], change: Change): void {
    if (change.kind === 'ins') {
      if (change.victims) Model.remove(list, change.victims);
      list.splice(change.index, 0, ...change.items);
    } else if (change.kind === 'del') Model.remove(list, change.victims);
  }

  private static remove(list: Item[], victims: number[]): void {
    for (const id of victims) {
      const at = list.findIndex((item) => item.id === id);
      if (at >= 0) list.splice(at, 1);
    }
  }

  /** The items a participant sees, in order: the text of the steps it holds with its own waiting changes laid over it. */
  seen(seat: Seat): Item[] {
    const list = this.at(seat.rev);
    for (const change of seat.waiting) Model.apply(list, change);
    return list;
  }

  text(seat: Seat): string {
    return this.seen(seat)
      .map((item) => item.ch)
      .join('');
  }

  textAt(rev: number): string {
    return this.at(rev)
      .map((item) => item.ch)
      .join('');
  }

  /** A participant made a change: it waits for its place. */
  made(seat: Seat, change: Change): void {
    seat.waiting.push(change);
    this.sights.push(this.seen(seat).map((item) => item.id));
  }

  /** The relay gave the first waiting change of a participant its place: step `rev`. */
  placed(seat: Seat, rev: number): void {
    if (rev !== this.rev + 1) throw new Error(`the model expected step ${this.rev + 1}, the relay placed ${rev}`);
    if (seat.rev !== this.rev) throw new Error(`${seat.name} got a place for a change it made on step ${seat.rev}, and the room was at ${this.rev}`);
    const change = seat.waiting[0];
    if (!change) throw new Error(`${seat.name} got a place for a change the model does not know`);
    const now = this.at(this.rev);
    const kill = (victims: number[]): void => {
      for (const id of victims) {
        const item = now.find((one) => one.id === id);
        if (item) item.died = rev;
      }
    };
    if (change.kind === 'del') kill(change.victims);
    else if (change.kind === 'color') {
      for (const id of change.victims) {
        const item = now.find((one) => one.id === id);
        if (item) item.color = change.color;
      }
    } else if (change.kind === 'indent') {
      let row = 0;
      for (let i = 0; i < change.index; i++) if (now[i].ch === '\n') row++;
      this.indents.set(rev, { row, value: change.value });
    } else {
      if (change.victims) kill(change.victims);
      const left = this.at(rev);
      // Before the item that stands at that place now, wherever the ones that went are marked.
      const before = left[change.index];
      const place = before ? this.all.indexOf(before) : this.all.length;
      for (const item of change.items) item.born = rev;
      this.all.splice(place, 0, ...change.items);
    }
    this.steps.push({ seat, change });
    this.rev = rev;
  }

  /** A participant read step `rev` of the order in. */
  read(seat: Seat, rev: number): void {
    if (rev !== seat.rev + 1) throw new Error(`${seat.name} read step ${rev} after step ${seat.rev}`);
    const step = this.steps[rev - 1];
    if (step.seat === seat) {
      seat.waiting.shift();
      seat.rev = rev;
      return;
    }
    const theirs = step.change;
    // What the participant has in front of it while its waiting changes are gone through: the text before the step, with
    // the waiting changes that were gone through laid over it.
    const list = this.at(rev - 1);
    const gone = new Set<number>(theirs.kind === 'del' || (theirs.kind === 'ins' && theirs.victims) ? (theirs as { victims: number[] }).victims.filter((id) => list.some((item) => item.id === id)) : []);
    /** Where the items of the step come to stand in what the participant has in front of it. */
    let place = 0;
    if (theirs.kind === 'ins') {
      // In the text before the step, without what the step took away: before the item that stands there.
      const left = list.filter((item) => !gone.has(item.id));
      const before = left[theirs.index];
      place = before ? list.indexOf(before) : list.length;
    }
    const before = (index: number): number => {
      let count = 0;
      for (let i = 0; i < index; i++) if (gone.has(list[i].id)) count++;
      return count;
    };
    const kept: Change[] = [];
    for (const change of seat.waiting) {
      if (change.kind === 'del' || change.kind === 'color') {
        // Is anything left of what it names, in what its author has in front of it?
        const left = change.victims.filter((id) => !gone.has(id) && list.some((item) => item.id === id));
        if (change.victims.some((id) => gone.has(id) && list.some((item) => item.id === id))) seat.partial++;
        if (change.kind === 'del') {
          for (const id of change.victims) {
            const at = list.findIndex((item) => item.id === id);
            if (at < 0) continue;
            if (at < place) place--;
            list.splice(at, 1);
          }
        }
        // Everything it was about went with what someone else did: nothing is left to do, and nothing to say.
        if (left.length === 0) seat.quiet++;
        else kept.push(change);
        continue;
      }
      if (change.kind === 'indent') {
        const inside = change.index > 0 && change.index < list.length && gone.has(list[change.index - 1].id) && gone.has(list[change.index].id);
        if (inside) {
          // The place was inside what someone took away. When paragraphs were joined there, the paragraph that was meant is no more.
          let joined = false;
          for (let i = change.index - 1; i >= 0 && gone.has(list[i].id); i--) if (list[i].ch === '\n') joined = true;
          for (let i = change.index; i < list.length && gone.has(list[i].id); i++) if (list[i].ch === '\n') joined = true;
          if (joined) {
            seat.dropped++;
            continue;
          }
        }
        let index = change.index - before(change.index);
        if (theirs.kind === 'ins' && place < change.index) index += theirs.items.length;
        change.index = index;
        kept.push(change);
        continue;
      }
      if (change.victims) {
        if (change.victims.some((id) => gone.has(id) && list.some((item) => item.id === id))) seat.partial++;
        for (const id of change.victims) {
          const at = list.findIndex((item) => item.id === id);
          if (at < 0) continue;
          if (at < place) place--;
          list.splice(at, 1);
        }
      }
      // Where it stands once the step is in: what the step took away before it no longer counts, what the step put in
      // before it does. At the very same place the waiting change stands first: it gets its place in the order later.
      let index = change.index - before(change.index);
      if (theirs.kind === 'ins' && place < change.index) index += theirs.items.length;
      if (change.index <= place) place += change.items.length;
      list.splice(change.index, 0, ...change.items);
      change.index = index;
      kept.push(change);
    }
    seat.waiting = kept;
    seat.rev = rev;
  }

  /** The colours that are known for the text after every step: one per character, "" where no step named it. */
  colors(): string[] {
    return this.at(this.rev)
      .filter((item) => item.ch !== '\n')
      .map((item) => item.color || '');
  }

  /** What is wrong with the order of the items at the end: every author saw them in the order they are in now. */
  inOrder(): string | undefined {
    const place = new Map<number, number>();
    this.at(this.rev).forEach((item, index) => place.set(item.id, index));
    for (const sight of this.sights) {
      let last = -1;
      for (const id of sight) {
        const at = place.get(id);
        if (at === undefined) continue;
        if (at < last) return `item ${id} stands before an item its author saw it after`;
        last = at;
      }
    }
    return undefined;
  }
}
