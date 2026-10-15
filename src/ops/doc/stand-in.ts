// Stand-in documents. A command in a text box runs on a stand-in: a document that is the real one
// except for its body, which is the box (see text-box.ts). What the commands keep beside a
// document, keyed by it - the counter of revision ids, the ids given in this session, how changes
// are tracked, the paragraph ids given, what stands for the document in a copy - belongs to the
// real document: a stand-in that counted for itself would give ids the document has already given.
// Those keys go through ownerOf().

import type { DocModel } from '../../model/doc/types';

const OWNERS = new WeakMap<DocModel, DocModel>();

/** Notes that `standIn` stands for `doc` (for the document `doc` stands for, when it is a stand-in itself). */
export function noteStandIn(standIn: DocModel, doc: DocModel): void {
  OWNERS.set(standIn, ownerOf(doc));
}

/** The document a stand-in stands for; the document itself for any other. */
export function ownerOf(doc: DocModel): DocModel {
  return OWNERS.get(doc) || doc;
}
