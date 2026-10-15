// Editing of documents: positions, operations, commands and undo.

export * from './position';
export { MAX_DEFAULT_TAB, MIN_DEFAULT_TAB, applyDocOp, applyDocOps, type DocOp } from './ops';
export { DocHistory } from './history';
export { buildEdit, Draft, type DocEdit, type DocEditOptions } from './draft';
export { deleteBackward, deleteForward, deleteRange, expandRangeForTables, insertBlock, insertPageBreak, insertText, splitParagraph } from './text';
export {
  MAX_TAB_STOPS,
  indentList,
  paragraphTabStops,
  patchParaProps,
  patchRunProps,
  rangeFormat,
  setDefaultTabStop,
  setParaFormat,
  setRunFormat,
  toggleList,
  toggleRunFormat,
  type ParaFormatPatch,
  type TabStopSpec,
  type TabStopsPatch,
  type RangeFormat,
  type RunFormatPatch,
  type ToggleKey,
} from './format';
export {
  DEFAULT_SECTION_MARGINS,
  DEFAULT_SECTION_PAGE,
  MAX_PAGE_TWIPS,
  MIN_PAGE_TWIPS,
  SECTION_PAPERS,
  documentSections,
  insertSectionBreak,
  patchedSection,
  sectionSetup,
  sectionsAt,
  setSectionSetup,
  type SectionBreakKind,
  type SectionMarginName,
  type SectionPatch,
  type SectionRef,
  type SectionScope,
  type SectionSetup,
} from './section';
export { deleteColumn, deleteRow, insertColumn, insertHyperlink, insertImage, insertRow, insertTable, type InsertImageOptions, type NewImage } from './structure';
export { FRAGMENT_LIMITS, MAX_TABLE_DEPTH, fragmentHasContent, fragmentNeeds, fragmentOf, insertFragment, originOf, type DocFragment, type FragmentMedia, type InsertFragmentOptions } from './fragment';
export { DocMatchScan, DocSearch, MAX_FIND_PARAGRAPHS, compileDocFind, docParagraphs, findAllInDocument, findInDocument, matchRange, paragraphMatches, replaceMatchEdit, replaceMatches, docFindCounters, type ReplaceMatchesOptions, type ReplaceMatchesResult, type DocFixed, type DocFindOptions, type DocMatch, type DocParagraphRef } from './find';
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
export { TEXT_BOX_OUTSIDE, editInTextBox, expandRangeInTextBox, insertTextInTextBox, readInTextBox, replaceInTextBox, textBoxExit, textBoxKey, textBoxProblem, type BoxCommand } from './text-box';
export { textLocked } from './find-edit';
export { replaceText } from './find-edit';
