// The match in the text of a comment Find went to last. The selection of the document cannot be
// there, so the view keeps it beside the selection: it is where the next search starts and what
// Replace replaces while the document, its selection and the comment are what they were then.

import type { DocModel } from '../model/doc/types';
import { commentOfStory, commentTextOffset, samePosition, type DocRange } from '../ops/doc/index';

/** A match in a comment that was gone to: its range, and the count of changes and the selection of the document then. */
export interface FoundAt {
  range: DocRange;
  changes: number;
  selection: DocRange | undefined;
}

const sameRange = (a: DocRange | undefined, b: DocRange | undefined): boolean =>
  a === undefined || b === undefined ? a === b : samePosition(a.anchor, b.anchor) && samePosition(a.focus, b.focus);

/**
 * The range of the match that was gone to, while it still holds: the document had no change since
 * (`changes`), its selection is the same, and the comment still has both ends of the match.
 * Undefined otherwise.
 */
export function foundInComment(at: FoundAt | undefined, doc: DocModel, changes: number, selection: DocRange | undefined): DocRange | undefined {
  if (!at || at.changes !== changes || !sameRange(at.selection, selection)) return undefined;
  const range = at.range;
  const id = commentOfStory(range.anchor.story);
  const comment = id !== undefined && doc.comments ? doc.comments.items.get(id) : undefined;
  if (!comment || range.focus.story !== range.anchor.story) return undefined;
  if (commentTextOffset(comment, range.anchor.path, range.anchor.offset) === undefined || commentTextOffset(comment, range.focus.path, range.focus.offset) === undefined) return undefined;
  return range;
}
