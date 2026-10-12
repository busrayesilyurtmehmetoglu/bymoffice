// Editing of documents: positions, operations, commands and undo.

export * from './position';
export { applyDocOp, applyDocOps, type DocOp } from './ops';
export { DocHistory } from './history';
export { buildEdit, Draft, type DocEdit, type DocEditOptions } from './draft';
export { deleteBackward, deleteForward, deleteRange, expandRangeForTables, insertBlock, insertPageBreak, insertText, splitParagraph } from './text';
export {
  indentList,
  patchParaProps,
  patchRunProps,
  rangeFormat,
  setParaFormat,
  setRunFormat,
  toggleList,
  toggleRunFormat,
  type ParaFormatPatch,
  type RangeFormat,
  type RunFormatPatch,
  type ToggleKey,
} from './format';
export { deleteColumn, deleteRow, insertColumn, insertHyperlink, insertImage, insertRow, insertTable, type InsertImageOptions, type NewImage } from './structure';
export { FRAGMENT_LIMITS, MAX_TABLE_DEPTH, fragmentHasContent, fragmentNeeds, fragmentOf, insertFragment, originOf, type DocFragment, type FragmentMedia, type InsertFragmentOptions } from './fragment';
export { DEFAULT_AUTHOR, TRACK_UNSUPPORTED, forcesTracking, isTracking, nextRevisionId, revisionDate, setTracking, type TrackOptions, type TrackingSetup } from './track';
export {
  REVISIONS_PROTECTED,
  REVISION_GONE,
  REVISION_LIST_LIMIT,
  REVISION_UNSUPPORTED,
  TRACK_ENFORCED,
  acceptAll,
  acceptRevision,
  listRevisions,
  placeOrder,
  rejectAll,
  rejectRevision,
  revisionCounters,
  revisionStep,
  revisionsOf,
  revisionsTouching,
  setTrackRevisions,
  type DocRevision,
  type DocRevisions,
  type ListRevisionsOptions,
  type RevisionEdit,
  type RevisionKind,
  type RevisionsEdit,
} from './revisions';
export {
  COMMENT_FIXED,
  COMMENT_GONE,
  COMMENT_NOT_HERE,
  MAX_COMMENTS,
  MAX_COMMENT_TEXT,
  addComment,
  commentAtPosition,
  commentText,
  editComment,
  initialsOf,
  listComments,
  removeComment,
  replyToComment,
  resolveComment,
  type CommentEdit,
  type CommentOptions,
  type DocComment,
} from './comments';
