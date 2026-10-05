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
