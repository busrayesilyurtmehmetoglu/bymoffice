// A picture on the clipboard. The copy (ops/pictures.ts: its bytes, size, turn, crop and text)
// is held here, for the page: every view of the page pastes it, into whichever workbook it shows.
// The clipboard gets a word in its place and HTML that carries a mark; a paste is the picture
// that is held only when its HTML carries that very mark (text is never taken for a picture).
// Where the browser lets a page write pictures to the clipboard, a PNG is put there as well, for
// other applications (see writeSystemClipboard()). Only one picture is held at a time.

import type { PictureCopy } from '../ops/pictures';

const MARK = /<[^<>]*\sdata-bym-picture\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

let held: { mark: string; copy: Promise<PictureCopy> } | undefined;
/** Goes up with every copy the page makes (of a picture, of cells, of anything): a write of the clipboard of the system that waits for its bytes is for the copy it was started for only. */
let generation = 0;

/** Says that something else was put on the clipboard (cells, a chart, text): a write of a picture that is still on its way is given up. */
export function clipboardMovedOn(): void {
  generation++;
}

/**
 * Holds a copy of a picture for the views of the page and gives what the clipboard gets for it:
 * `word` stands for it, and the HTML carries the mark of the copy. Nothing of a file is in either.
 */
export function holdPicture(copy: Promise<PictureCopy>, word: string): { text: string; html: string; mark: string } {
  // A copy that cannot be read is said when it is pasted; until then nothing waits for it.
  copy.then(undefined, () => undefined);
  const mark = Math.random().toString(36).slice(2) + Date.now().toString(36);
  generation++;
  held = { mark, copy };
  return { text: word, html: `<span data-bym-picture="${mark}">${word.replace(/[&<>"']/g, '')}</span>`, mark };
}

/** The copy that is held, when `mark` is its mark. */
export function heldPicture(mark: string | undefined): Promise<PictureCopy> | undefined {
  return held !== undefined && mark !== undefined && mark !== '' && held.mark === mark ? held.copy : undefined;
}

/** The copy that is held, whatever its mark; undefined when none is. */
export function heldPictureCopy(): Promise<PictureCopy> | undefined {
  return held ? held.copy : undefined;
}

/** Whether a picture is held for the views of the page. */
export function holdsPicture(): boolean {
  return held !== undefined;
}

/** Lets go of the copy that is held (the clipboard has moved on). */
export function releasePicture(): void {
  held = undefined;
}

/**
 * What the HTML of a paste is about: "picture" (with the copy) when it carries the mark of the
 * picture that is held; "stale" when it carries the mark of another copy (of this page before it
 * was loaded again, of another page): that copy is gone; "other" for anything else, no HTML at all
 * included. At most 64 marks are looked at.
 */
export function picturePasteKind(html: string | undefined): { kind: 'picture'; copy: Promise<PictureCopy> } | { kind: 'stale' | 'other' } {
  if (typeof html !== 'string' || html.indexOf('data-bym-picture') < 0) return { kind: 'other' };
  let found = false;
  MARK.lastIndex = 0;
  for (let match = MARK.exec(html), n = 0; match && n < 64; match = MARK.exec(html), n++) {
    found = true;
    const copy = heldPicture(match[1] !== undefined ? match[1] : match[2]);
    if (copy) return { kind: 'picture', copy };
  }
  return { kind: found ? 'stale' : 'other' };
}

/**
 * Puts the picture on the clipboard of the system as a PNG, with the word and the HTML of the
 * copy beside it, where the browser lets a page do that (the asynchronous clipboard with
 * ClipboardItem, inside the user's own copy or cut). A picture of another kind is not converted:
 * only the word and the mark stay on the clipboard. Whatever the browser refuses is let be: the
 * copy is held all the same. The write is given up when anything else was copied before the bytes
 * came: another picture, cells or a chart of a view (clipboardMovedOn()), or anything the window
 * hears a copy or a cut of (a text field, another editor) - a late write would put the picture
 * over what was copied since.
 */
export function writeSystemClipboard(put: { text: string; html: string; mark: string }, copy: Promise<PictureCopy>): void {
  const clipboard = typeof navigator !== 'undefined' ? navigator.clipboard : undefined;
  const host = typeof window !== 'undefined' ? (window as unknown as { ClipboardItem?: typeof ClipboardItem; addEventListener?: Window['addEventListener']; removeEventListener?: Window['removeEventListener'] }) : undefined;
  const Item = host ? host.ClipboardItem : undefined;
  if (!clipboard || typeof clipboard.write !== 'function' || typeof Item !== 'function' || !host) return;
  const mine = generation;
  // Heard in the capture phase of the window from now on: the copy that started this write is past it already.
  const moved = (): void => clipboardMovedOn();
  const listen = typeof host.addEventListener === 'function' && typeof host.removeEventListener === 'function';
  if (listen) {
    host.addEventListener!('copy', moved, true);
    host.addEventListener!('cut', moved, true);
  }
  const stop = (): void => {
    if (!listen) return;
    host.removeEventListener!('copy', moved, true);
    host.removeEventListener!('cut', moved, true);
  };
  const png = copy.then(
    (made) => {
      stop();
      if (made.type !== 'image/png' || !held || held.mark !== put.mark || generation !== mine) throw new Error('Not written');
      return new Blob([made.bytes as BlobPart], { type: 'image/png' });
    },
    (error: unknown) => {
      stop();
      throw error;
    },
  );
  try {
    const item = new Item({ 'image/png': png, 'text/html': new Blob([put.html], { type: 'text/html' }), 'text/plain': new Blob([put.text], { type: 'text/plain' }) });
    clipboard.write([item]).then(undefined, () => undefined);
  } catch {
    // ClipboardItem without promises, or without one of these kinds: the event's own data stays.
  }
  png.then(undefined, () => undefined);
}
