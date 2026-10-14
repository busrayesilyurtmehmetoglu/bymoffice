// What the view of a document needs beside itself while the document is edited together:
// where the caret of this participant is after a change of someone else, who the others are
// and where, and the carets of the others over the pages. Names are text of other people:
// they are only ever set as textContent.

import { mapPos, type Mark, type WirePos } from '../collab/doc/rebase';
import type { DocPeer } from '../collab/doc/session';
import type { DocModel } from '../model/doc/types';
import { containerBlocks, paragraphLength, resolvePosition, storyStart, type DocPosition, type DocRange } from '../ops/doc/position';

const toWire = (position: DocPosition): WirePos => (position.before ? { s: position.story, p: position.path, o: position.offset, b: true } : { s: position.story, p: position.path, o: position.offset });
const fromWire = (pos: WirePos): DocPosition => (pos.b ? { story: pos.s, path: pos.p.slice(), offset: pos.o, before: true } : { story: pos.s, path: pos.p.slice(), offset: pos.o });

/**
 * A selection read through what a change of someone else did: the caret stays before text
 * that came in exactly where it is (what this participant types next goes before it, as its
 * waiting changes do), and falls onto the place where text it stood in went. Where the blocks
 * it was in were replaced as a whole it stays in the block that stands there now (the view
 * brings it to a place of that block: placeNear). Undefined when nothing is known of its story.
 */
export function movedSelection(selection: DocRange, marks: readonly Mark[]): DocRange | undefined {
  let anchor = toWire(selection.anchor);
  let focus = toWire(selection.focus);
  for (const mark of marks) {
    if (mark.k === 'opq') {
      if (mark.s === '*' || mark.s === anchor.s || mark.s === focus.s) return undefined;
      continue;
    }
    if (mark.k === 'tch') continue;
    anchor = through(anchor, mark);
    focus = through(focus, mark);
  }
  return { anchor: fromWire(anchor), focus: fromWire(focus) };
}

/** How long what the others do is held back for an input method that composes. */
export const COLLAB_HOLD_MS = 30000;

/**
 * Holds what the others do back while text is composed (a change of the text around a
 * composition would end it), and lets go: a turn after the composition ended, at once when
 * the keyboard went away, and after COLLAB_HOLD_MS whatever happens - a composition whose end
 * is never told does not keep the others out for ever.
 */
export class CompositionHold {
  private held = false;
  private releasing = false;
  private stops: (() => void)[] = [];

  constructor(
    private readonly session: { pause(): void; resume(): void },
    private readonly later: (run: () => void, ms: number) => () => void = (run, ms) => {
      const timer = setTimeout(run, ms);
      return () => clearTimeout(timer);
    },
    private readonly limit: number = COLLAB_HOLD_MS,
  ) {}

  get holding(): boolean {
    return this.held;
  }

  begin(): void {
    if (this.held) {
      // A composition that starts while the last one waits to be let go: the hold goes on.
      if (this.releasing) this.arm();
      return;
    }
    this.held = true;
    this.session.pause();
    this.arm();
  }

  /** `soon`: a turn later (what was composed is put into the document first). */
  end(soon: boolean): void {
    if (!this.held) return;
    this.clear();
    if (!soon) return this.release();
    this.releasing = true;
    this.stops.push(this.later(() => this.release(), 0));
  }

  private arm(): void {
    this.clear();
    this.stops.push(this.later(() => this.release(), this.limit));
  }

  private clear(): void {
    for (const stop of this.stops.splice(0)) stop();
    this.releasing = false;
  }

  private release(): void {
    if (!this.held) return;
    this.clear();
    this.held = false;
    this.session.resume();
  }
}

/** A place read through a mark; in blocks that were replaced as a whole it stays in the block that stands where its own stood. */
function through(pos: WirePos, mark: Mark): WirePos {
  const moved = mapPos(pos, mark, false);
  if (!moved.gone || mark.k !== 'blk') return moved.pos;
  const depth = mark.c.length;
  // In a paragraph whose text is known before and after: the caret follows the text that stayed.
  if (mark.was && mark.now && pos.p.length === depth + 1 && mark.m > 0) {
    const read = followText(mark.was, mark.now, pos.p[depth] - mark.at, pos.o);
    if (read) return { s: pos.s, p: pos.p.slice(0, depth).concat([mark.at + read[0]]), o: read[1] };
  }
  const path = pos.p.slice(0, depth + 1);
  path[depth] = mark.at + Math.min(pos.p[depth] - mark.at, Math.max(mark.m, 1) - 1);
  // Inside a block that was replaced (a table) nothing is known: the place is the block.
  return { s: pos.s, p: path, o: pos.p.length === depth + 1 ? pos.o : 0 };
}

/**
 * A place in blocks whose text went from `was` to `now` (the blocks one after the other, as one
 * text with a line between two), read by what stayed: in the text before the first difference
 * it stays where it is, in the text after the last one it keeps its distance from the end, and
 * in what changed it stands at the end of the new text. [block, offset] in `now`.
 */
export function followText(was: readonly string[], now: readonly string[], block: number, offset: number): [number, number] | undefined {
  if (block < 0 || block >= was.length) return undefined;
  const before = was.join('\n');
  const after = now.join('\n');
  let at = offset;
  for (let i = 0; i < block; i++) at += was[i].length + 1;
  if (at > before.length) at = before.length;
  let head = 0;
  const most = Math.min(before.length, after.length);
  while (head < most && before.charCodeAt(head) === after.charCodeAt(head)) head++;
  let tail = 0;
  while (tail < most - head && before.charCodeAt(before.length - 1 - tail) === after.charCodeAt(after.length - 1 - tail)) tail++;
  const moved = at <= head ? at : at >= before.length - tail ? after.length - (before.length - at) : after.length - tail;
  let left = moved;
  for (let i = 0; i < now.length; i++) {
    if (left <= now[i].length) return [i, left];
    left -= now[i].length + 1;
  }
  return [now.length - 1, now[now.length - 1].length];
}

/**
 * The place of the document that is nearest to a place that may be none any more: the same
 * paragraph with the offset inside it, else the nearest paragraph of the same list, else the
 * start of the story. A caret that someone else's change left nowhere goes on from there, and
 * not from the start of the document.
 */
export function placeNear(doc: DocModel, position: DocPosition): DocPosition | undefined {
  try {
    resolvePosition(doc, position);
    return position;
  } catch {
    // Not a place: the nearest one is looked for.
  }
  try {
    const list = position.path.slice(0, -1);
    const blocks = containerBlocks(doc, position.story, list);
    const row = Math.max(0, Math.min(blocks.length - 1, position.path[position.path.length - 1]));
    for (let far = 0; far < blocks.length; far++) {
      for (const at of far === 0 ? [row] : [row - far, row + far]) {
        const block = blocks[at];
        if (!block || block.type !== 'paragraph') continue;
        const length = paragraphLength(block);
        // In its own paragraph the caret keeps its offset as far as the paragraph goes; in a neighbour it stands at the near end.
        return { story: position.story, path: list.concat([at]), offset: far === 0 ? Math.min(position.offset, length) : at < row ? length : 0 };
      }
    }
  } catch {
    // The list is no more.
  }
  try {
    return storyStart(doc, position.story) || storyStart(doc, 'body') || undefined;
  } catch {
    try {
      return storyStart(doc, 'body') || undefined;
    } catch {
      return undefined;
    }
  }
}

/** Another participant, as the pages show it. */
export interface PeerPlace {
  id: string;
  /** "#RRGGBB". */
  color: string;
  label: string;
  anchor: DocPosition;
  focus: DocPosition;
}

const COLOR = /^#[0-9A-Fa-f]{6}$/;

/** The others who said where they are. */
export function peerPlaces(peers: readonly DocPeer[], guest: (n: number) => string): PeerPlace[] {
  const out: PeerPlace[] = [];
  for (const peer of peers) {
    const presence = peer.presence;
    if (!presence) continue;
    out.push({
      id: peer.id,
      color: COLOR.test(peer.color) ? peer.color : '#5F6368',
      label: peer.name !== undefined && peer.name !== '' ? peer.name : guest(peer.guest),
      anchor: { story: presence.story, path: presence.anchor.path.slice(), offset: presence.anchor.offset },
      focus: { story: presence.story, path: presence.focus.path.slice(), offset: presence.focus.offset },
    });
  }
  return out;
}

/** A colour of a participant as the fill of its selection. */
export function peerFill(color: string): string {
  const value = COLOR.test(color) ? color : '#5F6368';
  return `rgba(${parseInt(value.slice(1, 3), 16)}, ${parseInt(value.slice(3, 5), 16)}, ${parseInt(value.slice(5, 7), 16)}, 0.22)`;
}

/** Where a caret is on a page: the element of the page, and the place in parts of its width and height. */
export interface CaretSpot {
  host: HTMLElement;
  left: number;
  top: number;
  height: number;
}

/** The carets of the others: a line in the colour of each with its name above it. */
export class PeerCarets {
  private shown = new Map<string, HTMLDivElement>();

  /** Puts every caret where `locate` finds it; the ones it does not find (their page is not there) are taken away. */
  place(peers: readonly PeerPlace[], locate: (position: DocPosition) => CaretSpot | undefined): void {
    const seen = new Set<string>();
    for (const peer of peers) {
      let spot: CaretSpot | undefined;
      try {
        spot = locate(peer.focus);
      } catch {
        // A place the document does not have (the other one is a step ahead): no caret until it says where it is again.
        spot = undefined;
      }
      if (!spot || seen.has(peer.id)) continue;
      seen.add(peer.id);
      let caret = this.shown.get(peer.id);
      if (!caret) {
        caret = document.createElement('div');
        caret.className = 'bym-doc-peer';
        caret.setAttribute('aria-hidden', 'true');
        const label = document.createElement('div');
        label.className = 'bym-doc-peer-name';
        label.style.cssText =
          'position:absolute;left:-1px;bottom:100%;max-width:160px;padding:0 4px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;font:10px/14px system-ui,sans-serif;color:#fff;border-radius:3px 3px 3px 0;pointer-events:none';
        caret.appendChild(label);
        this.shown.set(peer.id, caret);
      }
      caret.setAttribute('data-peer', peer.id);
      caret.style.cssText = `position:absolute;z-index:1;width:2px;margin:0 0 0 -1px;padding:0;border:0;pointer-events:none;background:${peer.color};left:${spot.left * 100}%;top:${spot.top * 100}%;height:${spot.height * 100}%`;
      const label = caret.firstChild as HTMLDivElement;
      label.style.background = peer.color;
      if (label.textContent !== peer.label) label.textContent = peer.label;
      if (caret.parentNode !== spot.host) spot.host.appendChild(caret);
    }
    this.shown.forEach((caret, id) => {
      if (seen.has(id)) return;
      if (caret.parentNode) caret.parentNode.removeChild(caret);
      this.shown.delete(id);
    });
  }

  destroy(): void {
    this.shown.forEach((caret) => {
      if (caret.parentNode) caret.parentNode.removeChild(caret);
    });
    this.shown.clear();
  }
}
