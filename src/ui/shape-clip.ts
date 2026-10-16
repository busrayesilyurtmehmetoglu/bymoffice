// A shape on the clipboard. The copy (ops/shapes.ts: its model, size and kind of anchor) is held
// here, for the page: every view of the page pastes it, into whichever workbook it shows. The
// clipboard gets a word in its place and HTML that carries a mark; a paste is the shape that is
// held only when its HTML carries that very mark (text is never taken for a shape). Only one shape
// is held at a time.

import type { ShapeCopy } from '../ops/shapes';

const MARK = /<[^<>]*\sdata-bym-shape\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

let held: { mark: string; copy: ShapeCopy } | undefined;

/** Holds a copy of a shape for the views of the page and gives what the clipboard gets for it: `word` stands for it, and the HTML carries the mark of the copy. */
export function holdShape(copy: ShapeCopy, word: string): { text: string; html: string; mark: string } {
  const mark = Math.random().toString(36).slice(2) + Date.now().toString(36);
  held = { mark, copy };
  return { text: word, html: `<span data-bym-shape="${mark}">${word.replace(/[&<>"']/g, '')}</span>`, mark };
}

/** The copy that is held, whatever its mark; undefined when none is. */
export function heldShapeCopy(): ShapeCopy | undefined {
  return held ? held.copy : undefined;
}

/** Lets go of the copy that is held (the clipboard has moved on). */
export function releaseShape(): void {
  held = undefined;
}

/**
 * What the HTML of a paste is about: "shape" (with the copy) when it carries the mark of the shape
 * that is held; "stale" when it carries the mark of another copy (of another page, or of this one
 * before it was loaded again); "other" for anything else, no HTML at all included. At most 64
 * marks are looked at.
 */
export function shapePasteKind(html: string | undefined): { kind: 'shape'; copy: ShapeCopy } | { kind: 'stale' | 'other' } {
  if (typeof html !== 'string' || html.indexOf('data-bym-shape') < 0) return { kind: 'other' };
  let found = false;
  MARK.lastIndex = 0;
  for (let match = MARK.exec(html), n = 0; match && n < 64; match = MARK.exec(html), n++) {
    found = true;
    const mark = match[1] !== undefined ? match[1] : match[2];
    if (held !== undefined && mark !== undefined && mark !== '' && held.mark === mark) return { kind: 'shape', copy: held.copy };
  }
  return { kind: found ? 'stale' : 'other' };
}
