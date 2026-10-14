// The commands of a document as its view calls them: the ones of ../../ops/doc, with those that
// can be shared noted while the view is in a room (see trace.ts). Outside of a room they are the
// commands themselves.

import * as raw from '../../ops/doc/index';
import { traced } from './trace';

export * from '../../ops/doc/index';

const shared = traced(raw, false);

export const insertText = shared.insertText;
export const deleteRange = shared.deleteRange;
export const splitParagraph = shared.splitParagraph;
export const deleteBackward = shared.deleteBackward;
export const deleteForward = shared.deleteForward;
export const setRunFormat = shared.setRunFormat;
export const toggleRunFormat = shared.toggleRunFormat;
export const setParaFormat = shared.setParaFormat;
export const toggleList = shared.toggleList;
export const indentList = shared.indentList;
export const insertPageBreak = shared.insertPageBreak;
export const insertTable = shared.insertTable;
export const insertRow = shared.insertRow;
export const deleteRow = shared.deleteRow;
export const insertColumn = shared.insertColumn;
export const deleteColumn = shared.deleteColumn;
export const insertHyperlink = shared.insertHyperlink;
