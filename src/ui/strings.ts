// The few words the views themselves say: labels for assistive technology, the
// header menu, notes, the document view's bar. Content of files is never translated.

import type { TabStopsStrings } from './doc-tabs-dialog';
import type { PrintPreviewStrings } from './print-preview';

export interface UiStrings {
  spreadsheet: string;
  cellEditor: string;
  actionFailed: string;
  /** Label of the grid: "Sheet <name>". */
  sheet(name: string): string;
  chartSheetNote: string;
  /** The name of a series of a chart that has none: `n` counts from 1. */
  chartSeries(n: number): string;
  /** The words a chart shows of its own: the entries of the legend of a waterfall. */
  chartWord(word: 'increase' | 'decrease' | 'total'): string;
  otherSheetNote: string;
  /** Said after the active cell when more than one cell is selected. */
  selected(range: string): string;
  insertRowsAbove(count: number): string;
  insertRowsBelow(count: number): string;
  deleteRows(count: number): string;
  insertColumnsLeft(count: number): string;
  insertColumnsRight(count: number): string;
  deleteColumns(count: number): string;
  cut: string;
  copy: string;
  paste: string;
  clearContents: string;
  /** Said when the menu cannot read the clipboard and the keyboard has to be used. */
  pasteWithKeyboard: string;
  // What the operations refuse to do, in the user's language. `axis` is 'row' or 'col'.
  tooManyCells: string;
  tooManyRows(action: 'format' | 'resize' | 'change'): string;
  cannotMergeWhole: string;
  cannotMergeArray(ref: string): string;
  cannotMergeTable: string;
  cannotMergePivot(name: string): string;
  cannotMergeFiltered: string;
  /** The question before a merge discards what `count` cells hold: its title, what it explains, and the answer that merges. */
  mergeLossTitle: string;
  mergeLossQuestion(count: number): string;
  mergeConfirm: string;
  /** Said for assistive technology after the panes were frozen ("" rows/columns) or unfrozen. */
  panesFrozen(rows: number, cols: number): string;
  /** The dialog of the tab stops of a document. */
  tabs: TabStopsStrings;
  insertPushesOff(count: number, axis: string): string;
  arrayFormulaBlocks(kind: 'insert' | 'delete', axis: string, ref: string): string;
  /** `what` names the content in the way: comments, drawings, tables. */
  notSupportedWith(kind: 'insert' | 'delete', axis: string, what: string): string;
  /** Rows or columns that would cut through the pivot table `name`. */
  pivotCut(kind: 'insert' | 'delete', axis: string, name: string): string;
  /** A deletion that would take the header row or a column of the source of the pivot table `name`. */
  pivotSourceCut(axis: string, name: string): string;
  /** Rows or columns that would push the pivot table `name` off the sheet. */
  pivotPushedOff(kind: 'insert' | 'delete', axis: string, name: string): string;
  /** The source of the pivot table `name` is not, or would not stay, one range of cells. */
  pivotSourceNotRange(kind: 'insert' | 'delete', axis: string, name: string): string;
  /** Rows or columns cannot move for a pivot table the model does not hold; `reason` as the engine says it. */
  pivotsFixed(kind: 'insert' | 'delete', axis: string, reason: string): string;
  pasteTooLarge(max: string): string;
  tooManyFormats: string;
  badFormula: string;
  copyTooLarge: string;
  // Sorting and filtering.
  /** Label of the menu of a filter button: "Sort and filter column B". */
  filterMenu(column: string): string;
  sortAscending: string;
  sortDescending: string;
  clearFilter: string;
  searchValues: string;
  selectAll: string;
  blanks: string;
  apply: string;
  cancel: string;
  noMatches: string;
  valueList: string;
  /** Under a value list that does not hold every value of its column. */
  tooManyValues(max: number): string;
  /** What a filter with criteria shows: "12 of 340 records shown". */
  recordsShown(shown: number, total: number): string;
  /** Said when choosing values took the place of a criterion that cannot be shown as values (a colour, an icon). */
  filterReplaced: string;
  nothingToFilter: string;
  cannotSortMerged: string;
  cannotSortArray(ref: string): string;
  /** `what` names the content in the way: comments, tables, pivot tables. */
  sortNotSupportedWith(action: 'sort' | 'filter', what: string): string;
  // Data validation.
  /** Label of the list a cell offers its values in. */
  chooseFromList: string;
  /** What a rule says when it has no message of its own. */
  validationFailed: string;
  /** Said when a value breaks a rule that only warns: the value is kept. */
  validationWarning(message: string): string;
  // Formulas this engine has something to say about (the cell has a mark in its corner).
  formulaNotCalculated: string;
  formulaCircular: string;
  formulaStale: string;
  /** Said before the comment of a cell, and the label of the box that shows it. */
  comment: string;
  // Notes on cells: the commands of the menu of a cell, and the field a note is typed in.
  addNote: string;
  editNote: string;
  deleteNote: string;
  /** Label of the field a note is typed in; the address of its cell follows it. */
  noteField: string;
  /** Said beside the field: how what was typed is applied and given up. */
  noteHint: string;
  noteDeleted: string;
  /** Why the comment of a cell cannot be changed: it is a conversation. `ref` is the address of the cell. */
  noteThreaded(ref: string): string;
  noteTooLong(max: number): string;
  tooManyNotes: string;
  /** Why no note of a sheet can be changed: `kind` is "drawing", "shared" or "parts" (Sheet.notesBlocked). */
  notesBlocked(kind: string): string;
  /** Said once when a sheet is shown of which some drawings or comments could not be read. */
  drawingsNotShown: string;
  // The document view.
  document: string;
  documentTools: string;
  /** Label of one page: "Page 3". */
  page(number: number): string;
  /** The page indicator: "Page 3 / 12". */
  pageOf(page: number, count: number): string;
  zoomIn: string;
  zoomOut: string;
  fitWidth: string;
  /** Label of the button that shows the zoom and resets it to 100%. */
  zoomLevel(percent: number): string;
  /** Shown while the pages are still being laid out. */
  preparingPages: string;
  documentTooLong: string;
  documentLayoutFailed: string;
  /** What stands in a placeholder: `kind` is "picture", "chart", "diagram", "shape", "vml", "object", ... */
  placeholder(kind: string): string;
  // Editing a document.
  /** Label of the field that takes what is typed. */
  documentEditor: string;
  cannotEditHere: string;
  /**
   * Said when a change in a text box is refused: `why` is "locked" (the shape or a content control locks its text),
   * "unwritable" (the file keeps another copy of the box that could not take the change), "tracked" (changes are
   * tracked), "shared" (the document is edited together) or "command" (this command does not work in a text box).
   */
  textBoxRefused(why: 'locked' | 'unwritable' | 'tracked' | 'shared' | 'command'): string;
  /** Said when the caret goes into a text box (`fixed`: one whose text cannot be changed) and out of it. */
  textBoxEntered(fixed: boolean): string;
  textBoxLeft: string;
  /** Said when a change is refused because the document tracks its changes and this kind of change cannot be tracked yet. */
  cannotTrackChange: string;
  /** Said once for a document whose protection keeps it read-only; `kind` is "readOnly", "comments" or "forms". */
  protectedDocument(kind: string): string;
  // Tracked changes.
  /**
   * What a tracked change is: `kind` is a kind of DocRevision ("insert", "delete", "moveFrom", "moveTo", "format",
   * "paragraphFormat", "tableFormat", "sectionFormat", "other", and "insertColumn" and "deleteColumn" for a column of
   * a table); the author, the date and the text come from the file.
   */
  revisionDescription(kind: string, author: string | undefined, date: string | undefined, text: string): string;
  /** Said when a change is stepped to: its number, how many there are, what it is, and whether the step went round to the first or the last one. */
  revisionStep(number: number, count: string, description: string, wrapped?: 'first' | 'last'): string;
  noRevisions: string;
  /** Accept or reject was asked for and the selection touches no change. */
  noRevisionHere: string;
  /**
   * Said after changes were accepted or rejected: how many, how many the document still has (a number as text; "2000+"
   * when there are more than are counted), how many could not be resolved, and how many section breaks went with
   * the paragraph marks that were removed.
   */
  revisionsResolved(verdict: 'accept' | 'reject', resolved: number, left: string, skipped: number, sections: number): string;
  cannotResolveRevision: string;
  revisionGone: string;
  revisionsProtected: string;
  trackingTurned(on: boolean): string;
  cannotStopTracking: string;
  // Comments of a document.
  /** The name of the box that shows a comment with its answers, and of its mark beside the page; `author` comes from the file. */
  commentBy(author: string): string;
  /** Said when a comment is stepped to: its number, how many threads there are, who wrote it and what it says. */
  commentStep(number: number, count: number, author: string, text: string, wrapped?: 'first' | 'last'): string;
  /** How many answers a comment has, beside its mark. */
  commentReplies(count: number): string;
  commentResolved: string;
  commentEdit: string;
  commentReply: string;
  commentResolve: string;
  commentReopen: string;
  commentDelete: string;
  commentSave: string;
  commentCancel: string;
  commentClose: string;
  /** Label of the field a comment is written in, and of the one for an answer. */
  commentText: string;
  commentReplyText: string;
  newComment: string;
  noComments: string;
  /** Said after a command of comments was done. */
  commentAdded: string;
  commentChanged: string;
  commentReplied: string;
  commentRemoved: string;
  commentResolvedSaid(done: boolean): string;
  /** A comment was asked for without a caret or a selection to put it at. */
  commentNeedsPlace: string;
  /** No comment can be put where the selection is (a header, a footer, a text box). */
  cannotCommentHere: string;
  /** The comment cannot be changed: its marks sit in content the editor keeps whole. */
  commentFixed: string;
  commentGone: string;
  /** The text of a comment is too long or is not there. */
  commentTextNeeded: string;
  /** The panel beside the pages that lists the comments of a document, and the button that shows and hides it. */
  commentsPanel: string;
  /** Label of the box that shows the resolved threads in the panel too. */
  commentsShowResolved: string;
  /** What the head of the panel says: how many threads it lists, and how many resolved ones it hides. */
  commentsPanelCount(shown: number, hidden: number): string;
  /** The panel has nothing to list: there are no comments, or every one is resolved and hidden. */
  commentsPanelEmpty(hidden: number): string;
  /** The button at the end of the list that shows the threads after the ones listed. */
  commentsShowMore(count: number): string;
  /** Said in the panel of a comment no text refers to. */
  commentNotAnchored: string;
  /** Spoken before the text a thread is about. */
  commentAbout: string;
  /** What a thread of the panel is read out as; every value but the counts comes from the file. */
  commentThreadLabel(author: string, date: string, text: string, replies: number, done: boolean): string;
  /** Said after a copy was pasted as text alone because changes are tracked. */
  cannotCutObjects: string;
  cannotWriteClipboard: string;
  /** A copy made in an editor of the page does not fit where it is pasted. */
  cannotPasteHere: string;
  /** A paste waits for the pictures of the copy to be read from the document they are in. */
  pasteWaiting: string;
  /** The caret moved or the document changed while a paste was waiting: it was not made. */
  pasteInterrupted: string;
  /** Said after Escape: the next Tab moves the focus out of the editor instead of typing a tab. */
  tabLeaves: string;
  documentPasteTooLarge: string;
  /** A file was converted to be opened for editing; the formats as they are written ("doc", "docx"). */
  openedConverted(from: string, to: string): string;
  /** A file could not be converted to be edited and is shown as it is, read-only. */
  notConverted(from: string): string;
  // Saving the file to where it came from (MountOptions.saveTo, the bridge's `save`).
  saved: string;
  /** `detail` is short and never from the file: "HTTP 500", or what the application said; may be empty. */
  saveFailed(detail: string): string;
  /** The server holds another file than the one that was opened (HTTP 412). */
  saveConflict: string;
  /** A save nobody asked for met the saves of others in the room again and again. */
  saveRoomBusy: string;
  saveInterrupted: string;
  /** The question a conflict leads to (see SaveToRequest.onConflict): its title, what it explains, and its two answers. */
  saveConflictTitle: string;
  saveConflictQuestion: string;
  saveOverwrite: string;
  saveCancel: string;
  /** The question for the password of a password-protected file: title, explanation, the label of the field, its two answers. */
  passwordTitle: string;
  passwordQuestion: string;
  passwordLabel: string;
  /** The password that was tried is not the file's; `attempts` passwords were tried so far. */
  passwordWrong(attempts: number): string;
  /** Shown while a password is being checked. */
  passwordChecking: string;
  passwordOpen: string;
  passwordCancel: string;
  /** The encrypted file that was opened has no integrity check. */
  noIntegrityCheck: string;
  /** The server answered the save with a redirect, which is not followed. */
  saveRedirected: string;
  /** A converted file is not saved over its original without the page's consent; `to` is "docx" or "xlsx". */
  saveConvertedNotAllowed(to: string): string;
  saveDenied: string;
  saveLocked: string;
  saveTimedOut: string;
  saveNoConnection: string;
  /** What is being typed could not be applied, so nothing was saved. */
  savePending: string;
  /** Auto-save sends nothing because the server named no version of the file to guard the save with. */
  autoSaveNoVersion: string;
  /** A font of MountOptions.fonts that could not be loaded; `late` when it was still loading when the file was shown. */
  fontNotLoaded(family: string, late: boolean): string;
  /** The font of the package that stands in for `family` was needed and could not be loaded. */
  embeddedFontNotLoaded(family: string): string;
  pictureNotSupported: string;
  linkNotAllowed: string;
  /** Said when the caret enters another part of the document: `kind` is "body", "header", "footer", "footnote" or "endnote". */
  editingStory(kind: string): string;
  /** Said for an empty line and for a space the caret passes. */
  blankLine: string;
  space: string;
  endOfDocument: string;
  /** The bubble over a selection made by touch, and its button that selects the whole text. */
  selectionTools: string;
  selectWholeText: string;
  // Printing a document.
  print: string;
  /** Said when the pages are being prepared for the print dialog. */
  printPreparing: string;
  /** Said when the print dialog was closed; replaces `printPreparing`. */
  printClosed: string;
  /** Shown in the bar while they are: "Preparing to print… 3 / 12". */
  printProgress(done: number, total: number): string;
  printNotAvailable: string;
  printTooLong: string;
  printFailed: string;
  // Printing a workbook.
  printSheetTooLong: string;
  printSheetFailed: string;
  /** Said when the sheets that were to be printed hold nothing. */
  printNothing: string;
  // Making a PDF file of a document or a workbook.
  /** The label of the button that downloads the PDF. */
  pdfDownload: string;
  /** Said when the pages are being painted for the file. */
  pdfPreparing: string;
  /** Said when the file is there; replaces `pdfPreparing`. */
  pdfReady: string;
  /** Shown while the pages are painted: "Making the PDF… 3 / 12". */
  pdfProgress(done: number, total: number): string;
  pdfNotAvailable: string;
  pdfTooLong: string;
  pdfFailed: string;
  /** Said when there is nothing to make a PDF of. */
  pdfNothing: string;
  // The sheet tabs.
  /** Label of the button beside the tabs that adds a sheet after the last one. */
  addSheet: string;
  /** Items of the menu of a tab. */
  insertSheet: string;
  renameSheet: string;
  deleteSheet: string;
  /** More items of the menu of a tab. */
  duplicateSheet: string;
  moveSheetLeft: string;
  moveSheetRight: string;
  hideSheet: string;
  /** Opens the list of the hidden sheets. */
  unhideSheet: string;
  /** Label of the list of the hidden sheets. */
  hiddenSheets: string;
  tabColor: string;
  noTabColor: string;
  /** Said when a tab that is dragged is let go: "Sheet Bir moved to place 2 of 3". */
  sheetMoved(name: string, place: number, count: number): string;
  /** Said after a sheet was copied without some of its content; `what` is the list copyLoss() made. */
  sheetCopiedWithout(name: string, what: string): string;
  /** One kind of content a copy of a sheet does not get, with its number: "2 conditional formats". */
  copyLoss(what: string, count: number): string;
  /** Said after a paste that left notes out (their cells have a conversation, or the notes of the sheet cannot be changed). */
  notesNotPasted(count: number): string;
  /**
   * Said after a paste that could not do rules (data validations, conditional formats) as they are where the cells
   * come from: rules of the copy that did not come, and rules that stayed on the cells though the copy has none there.
   */
  rulesNotPasted(count: number): string;
  cannotHideLastSheet: string;
  cannotSortThreaded: string;
  /** Label of the field a sheet is renamed in. */
  sheetName: string;
  /** Said after a sheet was deleted. */
  /** What is said while a workbook is edited together with others. */
  collab: {
    /** A participant who gave no name. */
    guest(n: number): string;
    state: { connecting: string; online: string; offline: string; detached: string; closed: string };
    /** The title of the list of participants: how many others are there. */
    others(n: number): string;
    /** What the button of the participants says: how many are in the file, this one too. */
    peopleCount(n: number): string;
    morePeople(n: number): string;
    you: string;
    viewOnly: string;
    onSheet(name: string): string;
    dropped(n: number): string;
    partial: string;
    undoGone: string;
    undoNone: string;
    undoLossy: string;
    skipped: string;
    peer: string;
    stale: string;
    reloaded(n: number): string;
    refused: string;
    /** The room is edited with another version of the package. */
    version: string;
    backlog(n: number): string;
    badStep: string;
    readOnly: string;
    detached: string;
    waiting: string;
    undoWait: string;
    slow: string;
    updating: string;
    typedKept(text: string): string;
    autoWaits: { room: string; offline: string };
    /** The question when the stored file is another one than the workbook that was kept apart from its room was opened as. */
    fork: { title: string; question: string; load: string; save: string; download: string; downloaded: string; downloadFailed: string };
    noStamp: string;
    copySheet: string;
    /** A chart is not cut while the workbook is edited together: it was copied, and stays where it is. */
    chartCut: string;
    /** The chart that was copied was changed or deleted since: edited together, it is not pasted. */
    chartCopyGone: string;
    tooLarge: string;
    unsupported: string;
    /** Pictures are refused while the workbook is edited together. */
    pictures: string;
    sheetDeleted(name: string): string;
    editGone: string;
    formulaGone(text: string): string;
    dialogClosed: string;
    /** The same for a document, where it is said in other words. */
    doc: {
      dropped(n: number): string;
      undoGone: string;
      undoNone: string;
      peer: string;
      badStep: string;
      readOnly: string;
      detached: string;
      /** Part of what was deleted or formatted here was deleted by someone else at the same time. */
      partial: string;
      /** A change that does not travel: a picture, a comment, a revision that is accepted, a paste with everything in it. */
      notShared: string;
      /** The document tracks its changes: it is not edited together. */
      tracked: string;
      /** The file is no docx (or was converted here): everybody in a room has to read the very same file. */
      unsupported: string;
      forkQuestion: string;
    };
  };
  sheetDeleted(name: string): string;
  /** Said when a reference was put into the formula being typed by pointing at cells: "Reference B5:D9". */
  referenceAdded(reference: string): string;
  /** Said when what was typed could not be written because the sheet of the cell is no longer in the workbook. */
  editSheetGone: string;
  // What the operations on sheets refuse to do.
  badSheetName: string;
  sheetNameTaken(name: string): string;
  cannotDeleteLastSheet: string;
  /** A chart sheet or a sheet of another kind that is only kept. */
  sheetKindFixed(action: 'rename' | 'delete' | 'copy' | 'color'): string;
  /** `what` names the content in the way: chart, pivotTable, table, form controls. */
  sheetBlockedBy(action: 'rename' | 'delete', what: string): string;
  /** `where` is the cell or the name whose formula cannot be read. */
  sheetFormulaUnread(action: 'rename' | 'delete', where: string): string;
  /** Two sheets whose names differ only by the case of their letters: the sheets of such a workbook are not changed. */
  sheetNamesClash(first: string, second: string): string;
  /** The dialogs that set rules: data validation and conditional formats. */
  rules: RuleStrings;
  /** Charts: the menu that adds one, the dialog that changes one, and what is said about them. */
  charts: ChartStrings;
  /** Pivot tables: the commands of the menu of a cell of one, and what is said about a refresh. */
  pivots: PivotStrings;
  /** Pictures of a sheet: the button that adds one, the menu of one, and what is said about them. */
  pictures: PictureStrings;
  /** The bar of Find and Replace of both views. */
  find: FindStrings;
  /** The dialog of the page setup of both views. */
  pageSetup: PageSetupStrings;
  /** The print preview of a workbook. */
  preview: PrintPreviewStrings;
}

/** What is said about pictures of a sheet. */
export interface PictureStrings {
  /** The button and the command that add a picture. */
  insert: string;
  added: string;
  moved: string;
  removed: string;
  /** Said when a picture is selected; `name` may be empty. */
  selected(name: string): string;
  /** The command of the menu of a picture that deletes it. */
  remove: string;
  /** The bytes are not a PNG, JPEG, GIF or BMP picture, or cannot be read. */
  notSupported: string;
  tooLarge: string;
  /** Too many pictures, or their bytes together. */
  tooMany: string;
  /** The sheet is protected, or its drawing in the file cannot be changed. */
  protectedSheet: string;
  placesFixed: string;
  /** A picture that shares its place in the drawing with another object. */
  wrapped: string;
  /** Copy and cut while a picture is selected. */
  noCopy: string;
  /** Paste of cells while a picture is selected. */
  noCellPaste: string;
  /** A file dropped on the grid while nothing can be inserted (a read-only view, a cell being edited). */
  dropRefused: string;
}

/** What the dialog of the page setup says (File -> Page Setup), in a workbook and in a document. */
export interface PageSetupStrings {
  title: string;
  ok: string;
  cancel: string;
  /** The row of buttons that act at once (page breaks, section breaks). */
  actions: string;
  /** "cm" or "in": the unit margins are given in. */
  unit: 'cm' | 'in';
  /** The decimal separator of lengths typed and shown. */
  decimal: string;
  page: string;
  paper: string;
  /** A paper the list does not have (a size of the file): it stays as it is. */
  otherPaper: string;
  orientation: string;
  portrait: string;
  landscape: string;
  scaling: string;
  adjustTo: string;
  fitTo: string;
  /** Percent of the normal size. */
  scale: string;
  fitWide: string;
  fitTall: string;
  margins: string;
  top: string;
  bottom: string;
  left: string;
  right: string;
  header: string;
  footer: string;
  centerOnPage: string;
  horizontally: string;
  vertically: string;
  sheet: string;
  /** The print area as references ("A1:D20"); empty for the used range. */
  printArea: string;
  titleRows: string;
  titleCols: string;
  gridLines: string;
  headings: string;
  /** At the active cell: a break above it and left of it. */
  insertBreak: string;
  removeBreak: string;
  resetBreaks: string;
  applyTo: string;
  thisSection: string;
  wholeDocument: string;
  /** Section breaks at the caret: how the next section starts. */
  sectionBreak(kind: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'): string;
  badLength: string;
  badScale: string;
  badFit: string;
  badArea: string;
  badRows: string;
  badCols: string;
  noRoom: string;
  /** Under the dialog where nothing can be changed: the file is edited together, where the page setup does not travel. */
  shared: string;
  /** The same in a document whose changes are tracked: a change of the page is not tracked. */
  tracked: string;
}

/** The names of papers: the same in both languages (Excel's codes and the names of the document's papers). */
export const PAPER_NAMES: { readonly [key: string]: string } = {
  '1': 'Letter',
  '3': 'Tabloid',
  '5': 'Legal',
  '7': 'Executive',
  '8': 'A3',
  '9': 'A4',
  '11': 'A5',
  '12': 'B4 (JIS)',
  '13': 'B5 (JIS)',
  letter: 'Letter',
  legal: 'Legal',
  executive: 'Executive',
  a3: 'A3',
  a4: 'A4',
  a5: 'A5',
  b5: 'B5 (JIS)',
};

/** What the bar of Find and Replace says. */
export interface FindStrings {
  /** The label of the bar. */
  title: string;
  findLabel: string;
  replaceLabel: string;
  previous: string;
  next: string;
  replace: string;
  replaceAll: string;
  /** Shows or hides the field of the replacement. */
  toggleReplace: string;
  close: string;
  matchCase: string;
  /** Of a document: the text is a word of its own. */
  wholeWord: string;
  /** Of a workbook: the text is the whole content of the cell. */
  entireCell: string;
  /** Of a document: only the body and its tables, not the headers, footers, notes, text boxes and comments. */
  mainOnly: string;
  lookIn: string;
  lookInFormulas: string;
  lookInValues: string;
  within: string;
  withinSheet: string;
  withinWorkbook: string;
  /** Said while a search goes through a large document or workbook. */
  searching: string;
  noMatches: string;
  /** How many matches there are; `more`: at least that many. */
  matches(count: number, more: boolean): string;
  /** Said after replacing: how many were replaced, and how many could not be. */
  replaced(count: number, skipped: number): string;
  /**
   * Added to what replaced() says: of those that could not be, how many are in text boxes, and why: 'shared' the file
   * is edited together, 'tracked' changes are tracked, 'kept' the file holds the box in a form that would lose
   * something if it were changed (its other version holds other text, say).
   */
  inTextBoxes(count: number, why: 'shared' | 'tracked' | 'kept'): string;
  /** Added to what replaced() says: of those that could not be, how many are in comments, which are not replaced while the file is edited together. */
  inComments(count: number): string;
  /** Added to what replaced() says: of those that could not be, how many are outside comments, in a file that takes comments and no other change. */
  outsideComments(count: number): string;
  /** A replacement in a room that had to be made in several steps. */
  replacedInSteps(steps: number): string;
  /** Said when the search was given up because the file changed while it ran. */
  changedMeanwhile: string;
  /** Said when there are more matches to replace than one command takes. */
  tooMany(limit: number): string;
  /** Nothing is replaced in a view that is read-only. */
  readOnly: string;
}

/** What the commands of pivot tables say. */
export interface PivotStrings {
  refresh: string;
  refreshAll: string;
  /** Said after a refresh: how many tables were refreshed. */
  refreshed(count: number): string;
  /** Why a table was not refreshed, from the text of the engine ('Cannot refresh the pivot table "<name>": <reason>'). */
  refused(message: string): string;
  /** Said when the workbook has no pivot table to refresh. */
  none: string;
  /** The command that makes a pivot table (cell menu), and the title of its dialog. */
  insert: string;
  createTitle: string;
  /** The fields of the dialog: the cells the table is made from, where it goes. */
  source: string;
  place: string;
  newSheet: string;
  thisSheet: string;
  cell: string;
  ok: string;
  cancel: string;
  /** A range or a cell that cannot be read. */
  badSource: string;
  badCell: string;
  /** Said once a table was made. */
  made: string;
  /** The command that opens the panel of the fields of a table (cell menu), and the title of the panel. */
  fields: string;
  panelTitle: string;
  close: string;
  fieldList: string;
  boxes: { filters: string; cols: string; rows: string; values: string };
  dropHere: string;
  /** Labels of the controls of an entry, by what it shows. */
  moveUp(name: string): string;
  moveDown(name: string): string;
  moveTo(name: string): string;
  remove(name: string): string;
  summarize(name: string): string;
  numberFormat(name: string): string;
  generalFormat: string;
  aggregates: Record<string, string>;
  subtotal: string;
  subtotals(name: string): string;
  sort(name: string): string;
  sorts: { manual: string; ascending: string; descending: string };
  shownItem(name: string): string;
  allItems: string;
  filterItems(name: string): string;
  moreItems(count: number): string;
  rowGrandTotals: string;
  colGrandTotals: string;
  /** Said once the fields of a table were changed. */
  changed: string;
  /** Said when the panel is asked for away from a pivot table. */
  noTable: string;
}

/** What the commands and the dialog of charts say. */
export interface ChartStrings {
  /** The label of the toolbar button and the title of its menu. */
  insert: string;
  /** By the types the commands make: column, bar, line, area, pie, doughnut, scatter; and "stackedColumn". */
  types: Record<string, string>;
  /** Items of the menu of a chart; with an ellipsis where a dialog follows. */
  edit: string;
  remove: string;
  dialogTitle: string;
  type: string;
  /** A chart of a type the dialog does not offer: what its list shows in the place of one. */
  otherType: string;
  grouping: string;
  /** By standard, stacked, percentStacked. */
  groupings: Record<string, string>;
  title: string;
  range: string;
  seriesIn: string;
  /** By auto, columns, rows. */
  directions: Record<string, string>;
  legend: string;
  /** By r, l, t, b, tr, none. */
  legends: Record<string, string>;
  dataLabels: string;
  axes: string;
  categoryTitle: string;
  valueTitle: string;
  minimum: string;
  maximum: string;
  series: string;
  categories: string;
  seriesName(n: number): string;
  seriesValues(n: number): string;
  seriesColor(n: number): string;
  ownColor(n: number): string;
  moveUp(n: number): string;
  moveDown(n: number): string;
  removeSeries(n: number): string;
  addSeries: string;
  ok: string;
  cancel: string;
  /** Read out when a chart is selected (`name`: its title, or the name it has in the file), added, changed, moved and deleted. */
  selected(name: string): string;
  added: string;
  changed: string;
  moved: string;
  removed: string;
  /** What is wrong with what was typed. */
  needRange: string;
  badRange: string;
  badNumber: string;
  needValues: string;
  /** Why the type of a chart of the file cannot be chosen, and why a chart cannot be changed at all. */
  typeFixed: string;
  notEditable: string;
  noChart: string;
  /** Why the series of a chart cannot be added, taken away or put in another order. */
  seriesFixed: string;
  groupingFixed: string;
  noAxis: string;
  tooManySeries(most: number): string;
  tooManyCharts: string;
  textTooLong(most: number): string;
  badCharacter: string;
  /** A sheet cannot be renamed or deleted: a chart names it in a reference that cannot be read. */
  refUnread: string;
  /** No chart is added to the sheet, moved or deleted on it: its drawing in the file takes no change. */
  placesFixed: string;
  /** The chart shares a wrapper of the file with another object and is not deleted. */
  wrapped: string;
  /** The item of the menu of a chart that puts a copy beside it. (Cut, copy and paste are named as they are for cells.) */
  duplicate: string;
  /** Read out when a chart was copied, cut, pasted and duplicated. */
  copied: string;
  cutOut: string;
  pasted: string;
  duplicated: string;
  /** Why a chart is not copied: its file holds references that are not known, and a copy could not be kept on the cells it shows. */
  notCopyable: string;
  /** The chart that was copied is not pasted: rows, columns or sheets changed since. */
  copyOutdated: string;
  /** What a copy was made without: `what` names the elements of the file ("userShapes", "externalData"). */
  copyLeftOut(what: readonly string[]): string;
  /** References of a copy on another sheet that still name no sheet. */
  copyUnbound(count: number): string;
  /** Said when cells are pasted while a chart is selected: nothing is pasted. */
  noCellPaste: string;
  /** The clipboard holds a chart that was copied in another view (or before this page was loaded again): nothing is pasted. */
  copyGone: string;
}

/** What the dialogs of data validation say. Kinds and operators are named by the words the file uses for them. */
export interface RuleStrings {
  /** The item of the cell menu and the label of the toolbar button; with an ellipsis, since a dialog follows. */
  validationCommand: string;
  validationTitle: string;
  ok: string;
  cancel: string;
  close: string;
  clearAll: string;
  /** Headings of the three parts of the dialog. */
  settings: string;
  inputMessage: string;
  errorAlert: string;
  allow: string;
  /** By ST_DataValidationType: none, whole, decimal, list, date, time, textLength, custom. */
  types: Record<string, string>;
  data: string;
  /** By ST_DataValidationOperator. */
  operators: Record<string, string>;
  /** Labels of the value fields, by what the kind of rule asks for. */
  minimum: string;
  maximum: string;
  value: string;
  startDate: string;
  endDate: string;
  date: string;
  startTime: string;
  endTime: string;
  time: string;
  source: string;
  formula: string;
  /** Under the source of a list: how items are typed in this locale. */
  sourceHint(separator: string): string;
  ignoreBlank: string;
  inCellDropdown: string;
  showInputMessage: string;
  showErrorAlert: string;
  style: string;
  /** By the style of the alert: stop, warning, information. */
  styles: Record<string, string>;
  title: string;
  message: string;
  /** Said in the dialog when the selection has cells of several rules, or some cells with a rule and some without. */
  mixedSelection: string;
  /** What the dialog refuses. */
  needNumber(field: string): string;
  needWholeNumber(field: string): string;
  needDate(field: string): string;
  needTime(field: string): string;
  needValue(field: string): string;
  needItems: string;
  itemWithComma: string;
  listTooLong(limit: number): string;
  textTooLong(field: string, limit: number): string;
  badRuleFormula(field: string): string;
  /** What the operations refuse. */
  extensionRule: string;
  tooManyValidations: string;
  tooManyRanges: string;
  formulaBeingTyped: string;
  /** Said when something needs the workbook as it stands while a dialog of rules is open with changes in it. */
  dialogOpen: string;
  /** Said in the dialog when cells of another sheet are pointed at for the cells a rule applies to. */
  pointOwnSheet: string;
  /** Said once a rule was set or cleared. */
  validationSet: string;
  validationCleared: string;

  // Conditional formats.
  /** The label of the toolbar button, and of the menu it opens. */
  conditionalCommand: string;
  /** Items of that menu; those that open a dialog end with an ellipsis. */
  quickRules: { greaterThan: string; lessThan: string; between: string; equal: string; containsText: string; duplicateValues: string; top10: string; colorScale: string; dataBar: string };
  clearFromSelection: string;
  clearFromSheet: string;
  manageRules: string;
  conditionalTitle: string;
  managerTitle: string;
  ruleType: string;
  /** By ST_CfType. */
  conditionalTypes: Record<string, string>;
  text: string;
  rank: string;
  percentOfRange: string;
  lowestValues: string;
  averageSide: string;
  averages: { above: string; below: string };
  includeAverage: string;
  stopIfTrue: string;
  format: string;
  formatWith: string;
  presets: Record<string, string>;
  textColor: string;
  fillColor: string;
  bold: string;
  italic: string;
  underline: string;
  strike: string;
  scaleKind: string;
  scales: Record<string, string>;
  lowest: string;
  midpoint: string;
  highest: string;
  barColor: string;
  /** Icon sets: the label of the list of sets, the sets by their name in the file (those of the 2010 extension included), and what the picture beside the list is. */
  iconStyle: string;
  iconSets: Record<string, string>;
  iconPreview: string;
  reverseIcons: string;
  iconOnly: string;
  /** The label of the list of periods of a "dates occurring" rule. */
  period: string;
  showRulesFor: string;
  scopeSelection: string;
  scopeSheet: string;
  appliesTo: string;
  moveUp: string;
  moveDown: string;
  edit: string;
  deleteRule: string;
  newRule: string;
  noRules: string;
  /** Under a list that does not show every rule. */
  /** A list of rules that shows a part of them: the places of the first and the last shown, and how many there are. */
  rulesShown(first: number, last: number, total: number): string;
  /** The buttons that show the rules after and before the ones shown; `count` is how many that are. */
  nextRules(count: number): string;
  previousRules(count: number): string;
  /** The name of one row of the list for assistive technology: "Rule 2: Cell value greater than 5". */
  ruleRow(place: number, description: string): string;
  /** How rules are named in the list. */
  cellValue: string;
  and: string;
  textRules: Record<string, string>;
  topItems(count: number, percent: boolean, bottom: boolean): string;
  average(above: boolean, equal: boolean): string;
  formulaIs: string;
  datesOccurring: string;
  periods: Record<string, string>;
  /** What the dialogs and the operations refuse. */
  needRange: string;
  badRange: string;
  needWholeUpTo(field: string, limit: number): string;
  /** A rule only a newer Excel states: nothing of it is changed. */
  extensionRuleKept: string;
  /** A rule a newer Excel states more about (its data bars): what it says is not changed; its cells and its place are, and it can be deleted. */
  extensionRuleSays: string;
  orderFixed: string;
  ruleKindKept: string;
  tooManyConditional: string;
  /** Said once rules changed. */
  ruleAdded: string;
  ruleChanged: string;
  ruleDeleted: string;
  rulesCleared: string;
}

const rows = (count: number) => (count === 1 ? 'row' : `${count} rows`);
const columns = (count: number) => (count === 1 ? 'column' : `${count} columns`);

export const STRINGS_EN: UiStrings = {
  spreadsheet: 'Spreadsheet',
  cellEditor: 'Cell editor',
  actionFailed: 'The action could not be completed',
  sheet: (name) => `Sheet ${name}`,
  chartSheetNote: 'This sheet is a chart sheet. Its chart is of a kind that is not shown yet.',
  chartSeries: (n) => `Series${n}`,
  chartWord: (word) => (word === 'increase' ? 'Increase' : word === 'decrease' ? 'Decrease' : 'Total'),
  otherSheetNote: 'This kind of sheet is not shown yet.',
  selected: (range) => `${range} selected`,
  insertRowsAbove: (count) => `Insert ${rows(count)} above`,
  insertRowsBelow: (count) => `Insert ${rows(count)} below`,
  deleteRows: (count) => `Delete ${rows(count)}`,
  insertColumnsLeft: (count) => `Insert ${columns(count)} left`,
  insertColumnsRight: (count) => `Insert ${columns(count)} right`,
  deleteColumns: (count) => `Delete ${columns(count)}`,
  cut: 'Cut',
  copy: 'Copy',
  paste: 'Paste',
  clearContents: 'Clear contents',
  pasteWithKeyboard: 'The clipboard cannot be read here. Press Ctrl+V to paste.',
  tooManyCells: 'Too many cells to format at once',
  tooManyRows: (action) => `Too many rows to ${action} at once`,
  cannotMergeWhole: 'Whole rows or columns cannot be merged',
  cannotMergeArray: (ref) => `Cells cannot be merged through the array formula in ${ref}`,
  cannotMergeTable: 'Cells cannot be merged in a sheet with a table yet',
  cannotMergePivot: (name) => `Cells of the pivot table "${name}" cannot be merged`,
  cannotMergeFiltered: 'Cells cannot be merged across rows the filter hides',
  mergeLossTitle: 'Merge cells',
  mergeLossQuestion: (count) => `Merging keeps only the value of the upper-left cell. ${count === 1 ? 'The value of 1 other cell is' : `The values of ${count} other cells are`} discarded.`,
  mergeConfirm: 'Merge',
  panesFrozen: (rows, cols) => (rows === 0 && cols === 0 ? 'Panes unfrozen' : `Frozen: ${rows} ${rows === 1 ? 'row' : 'rows'}, ${cols} ${cols === 1 ? 'column' : 'columns'}`),
  tabs: {
    title: 'Tabs',
    stops: 'Tab stops',
    none: 'No tab stops: the default ones apply',
    position: 'Position',
    alignment: 'Alignment',
    leader: 'Leader',
    aligns: { left: 'Left', center: 'Centre', right: 'Right', decimal: 'Decimal', bar: 'Bar' },
    leaders: { none: 'None', dot: 'Dots', hyphen: 'Dashes', underscore: 'Line', heavy: 'Heavy line', middleDot: 'Middle dots' },
    set: 'Set',
    clear: (stop) => `Clear the tab stop at ${stop}`,
    clearAll: 'Clear all',
    defaultStops: 'Default tab stops',
    ok: 'OK',
    cancel: 'Cancel',
    unit: 'in',
    decimal: '.',
    badPosition: 'Type a position between -22 and 22 inches',
    badDefault: 'Type a distance greater than 0 and at most 22 inches',
    defaultShared: 'Not changed while the document is edited together',
  },
  insertPushesOff: (count, axis) => `Cannot insert ${count} ${axis === 'row' ? 'row' : 'column'}(s): cells would be pushed off the sheet`,
  arrayFormulaBlocks: (kind, axis, ref) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} through the array formula in ${ref}`,
  notSupportedWith: (kind, axis, what) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} in a sheet with ${what}: not supported yet`,
  pivotCut: (kind, axis, name) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} through the pivot table "${name}"`,
  pivotSourceCut: (axis, name) =>
    axis === 'row' ? `The header row of the source of the pivot table "${name}" cannot be deleted` : `Columns of the fields of the source of the pivot table "${name}" cannot be deleted`,
  pivotPushedOff: (kind, axis, name) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'}: the pivot table "${name}" would be pushed off the sheet`,
  pivotSourceNotRange: (kind, axis, name) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'}: the source of the pivot table "${name}" is not one range of cells`,
  pivotsFixed: (kind, axis, reason) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} in this workbook because of its pivot tables: ${reason}`,
  pasteTooLarge: (max) => `Cannot paste more than ${max} cells at once`,
  tooManyFormats: 'The workbook has too many cell formats',
  badFormula: 'This formula cannot be read. Use English function names and commas between arguments.',
  copyTooLarge: 'The range is too large to copy',
  filterMenu: (column) => `Sort and filter column ${column}`,
  sortAscending: 'Sort A to Z',
  sortDescending: 'Sort Z to A',
  clearFilter: 'Clear filter',
  searchValues: 'Search',
  selectAll: '(Select all)',
  blanks: '(Blanks)',
  apply: 'Apply',
  cancel: 'Cancel',
  noMatches: 'Nothing matches',
  valueList: 'Values',
  tooManyValues: (max) => `Only the first ${max} values are listed`,
  recordsShown: (shown, total) => `${shown} of ${total} records shown`,
  filterReplaced: 'The colour or icon filter of this column was replaced by the chosen values.',
  nothingToFilter: 'There is no data to filter here',
  cannotSortMerged: 'A range with merged cells cannot be sorted',
  cannotSortArray: (ref) => `The array formula in ${ref} cannot be sorted through`,
  sortNotSupportedWith: (action, what) => `${action === 'sort' ? 'Sorting' : 'Filtering'} is not supported yet in a sheet with ${what}`,
  chooseFromList: 'Choose from the list',
  validationFailed: 'This value does not match the data validation rule of the cell.',
  validationWarning: (message) => `${message} (the value was entered anyway)`,
  formulaNotCalculated: 'This formula cannot be calculated here; the last value in the file is shown',
  formulaCircular: 'Circular reference',
  formulaStale: 'The calculation is not finished; the value may be out of date',
  comment: 'Comment',
  addNote: 'Add note',
  editNote: 'Edit note',
  deleteNote: 'Delete note',
  noteField: 'Note',
  noteHint: 'Ctrl+Enter applies, Escape cancels',
  noteDeleted: 'Note deleted',
  noteThreaded: (ref) => `The comment in ${ref} is a threaded comment: it cannot be edited here`,
  noteTooLong: (max) => `A note cannot be longer than ${max} characters`,
  tooManyNotes: 'The sheet has too many notes',
  notesBlocked: (kind) =>
    `The notes of this sheet cannot be changed: ${kind === 'shared' ? 'another sheet shows the same notes' : kind === 'parts' ? 'the file keeps them in more than one part' : 'its drawing of note boxes cannot be read'}`,
  drawingsNotShown: 'Some drawings or comments of this sheet could not be read and are not shown',
  document: 'Document',
  documentTools: 'Document tools',
  page: (number) => `Page ${number}`,
  pageOf: (page, count) => `Page ${page} / ${count}`,
  zoomIn: 'Zoom in',
  zoomOut: 'Zoom out',
  fitWidth: 'Fit to width',
  zoomLevel: (percent) => `Zoom ${percent}%, reset to 100%`,
  preparingPages: 'Preparing pages\u2026',
  documentTooLong: 'The document is too long: only its first pages are shown',
  documentLayoutFailed: 'The document could not be laid out completely',
  placeholder: (kind) => (kind === 'chart' ? 'Chart' : kind === 'diagram' ? 'Diagram' : kind === 'object' ? 'Embedded object' : kind === 'picture' ? 'Picture' : 'Drawing'),
  documentEditor: 'Document editor',
  cannotEditHere: 'This part of the document cannot be edited',
  textBoxRefused: (why) =>
    why === 'locked'
      ? 'The text of this text box is locked; nothing was changed'
      : why === 'unwritable'
        ? 'This text box cannot be changed: the file keeps another copy of it that could not take the change; nothing was changed'
        : why === 'tracked'
          ? 'Text boxes cannot be edited while changes are tracked; nothing was changed'
          : why === 'shared'
            ? 'Text boxes cannot be edited while the document is edited together; nothing was changed'
            : 'This cannot be done in a text box; nothing was changed',
  textBoxEntered: (fixed) => (fixed ? 'In a text box whose text cannot be changed. Escape leaves it' : 'In a text box. Escape leaves it'),
  textBoxLeft: 'Out of the text box',
  cannotTrackChange: 'Changes are tracked in this document, and this kind of change cannot be tracked yet; nothing was changed',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'This document is protected: only comments can be added and changed; its text cannot be edited'
      : kind === 'forms'
        ? 'This document is protected: only its form fields may be filled in. It is shown read-only'
        : 'This document is protected against changes. It is shown read-only',
  revisionDescription: (kind, author, date, text) => {
    const what: { [kind: string]: string } = {
      insert: 'Inserted',
      delete: 'Deleted',
      insertColumn: 'Column inserted',
      deleteColumn: 'Column deleted',
      moveFrom: 'Moved from here',
      moveTo: 'Moved here',
      format: 'Formatting changed',
      paragraphFormat: 'Paragraph formatting changed',
      tableFormat: 'Table formatting changed',
      sectionFormat: 'Page layout changed',
    };
    return `${what[kind] || 'A change that can only be kept'}${author !== undefined && author !== '' ? ` by ${author}` : ''}${date ? `, ${date}` : ''}${text !== '' ? `: ${text}` : ''}`;
  },
  revisionStep: (number, count, description, wrapped) => `${wrapped === 'first' ? 'Back at the first change. ' : wrapped === 'last' ? 'At the last change. ' : ''}Change ${number} of ${count}. ${description}`,
  noRevisions: 'The document has no tracked changes',
  noRevisionHere: 'There is no tracked change here',
  revisionsResolved: (verdict, resolved, left, skipped, sections) =>
    `${verdict === 'accept' ? 'Accepted' : 'Rejected'} ${resolved} ${resolved === 1 ? 'change' : 'changes'}. ${left === '0' ? 'No changes left.' : `${left} ${left === '1' ? 'change' : 'changes'} left.`}` +
    (skipped > 0 ? ` ${skipped} ${skipped === 1 ? 'change' : 'changes'} could not be ${verdict === 'accept' ? 'accepted' : 'rejected'} and ${skipped === 1 ? 'was' : 'were'} left as ${skipped === 1 ? 'it is' : 'they are'}.` : '') +
    (sections > 0 ? ` ${sections} section ${sections === 1 ? 'break was' : 'breaks were'} removed with ${sections === 1 ? 'its paragraph mark' : 'their paragraph marks'}.` : ''),
  cannotResolveRevision: 'This change cannot be accepted or rejected yet; nothing was changed',
  revisionGone: 'The change is no longer in the document',
  revisionsProtected: 'The protection of this document does not allow changes to be accepted or rejected',
  trackingTurned: (on) => (on ? 'Changes are tracked from now on' : 'Changes are no longer tracked'),
  cannotStopTracking: 'Tracking cannot be turned off in this document',
  commentBy: (author) => (author !== '' ? `Comment by ${author}` : 'Comment'),
  commentStep: (number, count, author, text, wrapped) =>
    `${wrapped === 'first' ? 'Back at the first comment. ' : wrapped === 'last' ? 'At the last comment. ' : ''}Comment ${number} of ${count}${author !== '' ? `, by ${author}` : ''}${text !== '' ? `: ${text}` : ''}`,
  commentReplies: (count) => (count === 1 ? '1 reply' : `${count} replies`),
  commentResolved: 'Resolved',
  commentEdit: 'Edit',
  commentReply: 'Reply',
  commentResolve: 'Resolve',
  commentReopen: 'Reopen',
  commentDelete: 'Delete',
  commentSave: 'Save',
  commentCancel: 'Cancel',
  commentClose: 'Close',
  commentText: 'Comment text',
  commentReplyText: 'Reply text',
  newComment: 'New comment',
  noComments: 'The document has no comments',
  commentAdded: 'Comment added',
  commentChanged: 'Comment changed',
  commentReplied: 'Reply added',
  commentRemoved: 'Comment deleted',
  commentResolvedSaid: (done) => (done ? 'Comment resolved' : 'Comment reopened'),
  commentNeedsPlace: 'Put the caret in the text, or select the text, the comment is about',
  cannotCommentHere: 'A comment cannot be added here',
  commentFixed: 'This comment cannot be changed here: its marks sit in a part of the document that is kept as it is',
  commentGone: 'The comment is no longer in the document',
  commentTextNeeded: 'A comment needs text, of at most 32,767 characters',
  commentsPanel: 'Comments',
  commentsShowResolved: 'Show resolved',
  commentsPanelCount: (shown, hidden) => `${shown === 1 ? '1 thread' : `${shown} threads`}${hidden > 0 ? `, ${hidden} resolved hidden` : ''}`,
  commentsPanelEmpty: (hidden) => (hidden > 0 ? `Every comment is resolved (${hidden} hidden)` : 'The document has no comments'),
  commentsShowMore: (count) => `Show ${count} more`,
  commentNotAnchored: 'Not tied to any text',
  commentAbout: 'About',
  commentThreadLabel: (author, date, text, replies, done) =>
    `${author !== '' ? `Comment by ${author}` : 'Comment'}${date !== '' ? `, ${date}` : ''}${done ? ', resolved' : ''}${replies > 0 ? `, ${replies === 1 ? '1 reply' : `${replies} replies`}` : ''}${text !== '' ? `: ${text}` : ''}`,
  cannotCutObjects: 'The selection holds something that cannot be cut (a note, an equation, a text box, a content control, a tracked change); it was left in place',
  cannotWriteClipboard: 'The clipboard cannot be written here; nothing was cut',
  cannotPasteHere: 'What was copied cannot be pasted here',
  pasteWaiting: 'Reading the pictures of the copy; it is pasted when they are there',
  pasteInterrupted: 'The selection or the document changed while the pictures of the copy were read; nothing was pasted',
  tabLeaves: 'Press Tab to leave the editor',
  documentPasteTooLarge: 'The text on the clipboard is too long to paste',
  openedConverted: (from, to) => `This ${from} file was converted to ${to} to be edited; it is saved as ${to}`,
  notConverted: (from) => `This ${from} file could not be converted to be edited; it is shown read-only`,
  saved: 'Saved',
  saveFailed: (detail) => (detail ? `The file could not be saved (${detail})` : 'The file could not be saved'),
  saveInterrupted: 'Not saved yet: the file was being edited the whole time it was written. It is saved when the editing pauses.',
  saveRoomBusy: 'Not saved yet: several people are saving this file at once. Auto-save waits and tries again.',
  saveConflict: 'The file was not saved: the file on the server is not the one you opened (someone else changed it, or an earlier save of yours was written although its answer did not arrive)',
  saveConflictTitle: 'The file was changed on the server',
  saveConflictQuestion:
    'The file was changed on the server after you opened it (someone else changed it, or an earlier save of yours was written although its answer did not arrive). If you save, the changes there are overwritten.',
  saveOverwrite: 'Overwrite',
  saveCancel: 'Cancel',
  passwordTitle: 'Password required',
  passwordQuestion: 'This file is protected with a password.',
  passwordLabel: 'Password',
  passwordWrong: (attempts) => (attempts > 1 ? `Wrong password (${attempts} attempts). Try again.` : 'Wrong password. Try again.'),
  passwordChecking: 'Checking the password…',
  passwordOpen: 'Open',
  passwordCancel: 'Cancel',
  noIntegrityCheck: 'This file has no integrity check: its content may have been changed by someone who does not know the password',
  saveRedirected: 'The file was not saved: the server sent the request to another address; you may have been signed out',
  saveConvertedNotAllowed: (to) => `Changes to this file cannot be saved here: it was converted to ${to}, and this page has not allowed saving that in place of the original file`,
  saveDenied: 'The file was not saved: you are not allowed to save it',
  saveLocked: 'The file was not saved: it is locked',
  saveTimedOut: 'The file could not be saved: no answer came in time',
  saveNoConnection: 'The file could not be saved: the server could not be reached',
  savePending: 'The file was not saved: what is being typed could not be applied',
  autoSaveNoVersion: 'Auto-save is not working: the server did not say which version of the file it holds. Save with the Save button.',
  embeddedFontNotLoaded: (family) => `This device does not have the font "${family}", and the font that stands in for it could not be loaded; the text is shown in another font. Give the address of the font folder with "fontsUrl", or turn the font off with "fonts: false"`,
  fontNotLoaded: (family, late) => (late ? `The font "${family}" is still loading; the file is shown without it until it arrives` : `The font "${family}" could not be loaded; the file is shown without it`),
  pictureNotSupported: 'This picture cannot be inserted',
  linkNotAllowed: 'This address cannot be used as a link',
  editingStory: (kind) => (kind === 'header' ? 'Editing the header' : kind === 'footer' ? 'Editing the footer' : kind === 'footnote' ? 'Editing a footnote' : kind === 'endnote' ? 'Editing an endnote' : 'Editing the text'),
  blankLine: 'blank',
  space: 'space',
  endOfDocument: 'End of the document',
  selectionTools: 'Selection',
  selectWholeText: 'Select all',
  print: 'Print',
  printPreparing: 'Preparing to print\u2026',
  printClosed: 'The print dialog was closed',
  printProgress: (done, total) => `Preparing to print\u2026 ${done} / ${total}`,
  printNotAvailable: 'Printing is not available here',
  printTooLong: 'The document is too long to print at once',
  printFailed: 'The document could not be printed',
  printSheetTooLong: 'The sheet is too long to print at once',
  printSheetFailed: 'The sheet could not be printed',
  printNothing: 'There is nothing to print',
  pdfDownload: 'Download as PDF',
  pdfPreparing: 'Making the PDF\u2026',
  pdfReady: 'The PDF is ready',
  pdfProgress: (done, total) => `Making the PDF\u2026 ${done} / ${total}`,
  pdfNotAvailable: 'A PDF cannot be made here',
  pdfTooLong: 'This is too long to be made into one PDF',
  pdfFailed: 'The PDF could not be made',
  pdfNothing: 'There is nothing to make a PDF of',
  addSheet: 'Add sheet',
  insertSheet: 'Insert sheet',
  renameSheet: 'Rename',
  deleteSheet: 'Delete',
  duplicateSheet: 'Duplicate',
  moveSheetLeft: 'Move left',
  moveSheetRight: 'Move right',
  hideSheet: 'Hide',
  unhideSheet: 'Unhide\u2026',
  hiddenSheets: 'Hidden sheets',
  tabColor: 'Tab colour\u2026',
  noTabColor: 'No colour',
  sheetMoved: (name, place, count) => `Sheet ${name} moved to place ${place} of ${count}`,
  sheetCopiedWithout: (name, what) => `Sheet ${name} was copied without: ${what}`,
  copyLoss: (what, count) => {
    const names: Record<string, [string, string]> = {
      conditionalFormatExtensions: ['conditional format setting of a newer Excel', 'conditional format settings of a newer Excel'],
      ruleContent: ['rule with content that only its own sheet can hold', 'rules with content that only their own sheet can hold'],
      drawings: ['picture or shape', 'pictures and shapes'],
      threadedComments: ['threaded comment', 'threaded comments'],
      externalLinks: ['link to another place', 'links to other places'],
      pageBreaks: ['page break', 'page breaks'],
      table: ['table', 'tables'],
      pivotTable: ['pivot table', 'pivot tables'],
      rowAttributes: ['row with settings of its own', 'rows with settings of their own'],
      sparklineGroups: ['group of sparklines', 'groups of sparklines'],
      customSheetViews: ['list of custom views', 'lists of custom views'],
      ignoredErrors: ['list of ignored errors', 'lists of ignored errors'],
      cellWatches: ['list of watched cells', 'lists of watched cells'],
      sortState: ['sort setting', 'sort settings'],
      phoneticPr: ['phonetic setting', 'phonetic settings'],
      outlinePr: ['outline setting', 'outline settings'],
      outlineLevels: ['setting of outline levels', 'settings of outline levels'],
      codeName: ['code name', 'code names'],
      legacyDrawingHF: ['set of header and footer pictures', 'sets of header and footer pictures'],
      picture: ['background picture', 'background pictures'],
    };
    const known = Object.prototype.hasOwnProperty.call(names, what) ? names[what] : undefined;
    return known ? `${count} ${count === 1 ? known[0] : known[1]}` : what;
  },
  notesNotPasted: (count) => (count === 1 ? 'One note was not pasted: its cell cannot take a note' : `${count} notes were not pasted: their cells cannot take notes`),
  rulesNotPasted: (count) =>
    count === 1
      ? 'One data validation or conditional format did not come with the cells as it is: a rule that is kept as the file has it stays where it is and is not copied, and extension content this sheet cannot hold is left out'
      : `${count} data validations or conditional formats did not come with the cells as they are: rules that are kept as the file has them stay where they are and are not copied, and extension content this sheet cannot hold is left out`,
  cannotHideLastSheet: 'The last visible sheet cannot be hidden',
  cannotSortThreaded: 'A range with a threaded comment cannot be sorted yet',
  sheetName: 'Sheet name',
  collab: {
    guest: (n) => `Guest ${n}`,
    state: { connecting: 'Connecting…', online: 'Editing together', offline: 'Offline: changes are sent when the connection is back', detached: 'Detached: your changes are here only. Save the file', closed: 'Not connected: changes stay here' },
    others: (n) => (n === 1 ? '1 other person is here' : `${n} other people are here`),
    peopleCount: (n) => (n === 1 ? '1 person in this file' : `${n} people in this file`),
    morePeople: (n) => `and ${n} more`,
    you: '(you)',
    viewOnly: 'viewing only',
    onSheet: (name) => `on ${name}`,
    dropped: (n) => (n === 1 ? 'One of your changes could not be applied because someone else deleted or changed what it was about.' : `${n} of your changes could not be applied because someone else deleted or changed what they were about.`),
    partial: 'Part of one of your changes was not applied: someone else deleted the rows, columns or sheet it was in.',
    undoGone: 'This step cannot be undone any more: someone else changed or deleted what it was about.',
    undoNone: 'This step cannot be undone while the workbook is edited together.',
    undoLossy: 'The cells are back; notes, links and rules that went with them are not.',
    skipped: 'Rules and charts do not travel with cells while the workbook is edited together; the cells were changed.',
    peer: 'Something that is no change of this workbook arrived; editing together has ended. Your changes stay here.',
    stale: 'The file was changed elsewhere and could not be loaded again; editing together has ended.',
    reloaded: (n) => (n === 1 ? 'The file was loaded again; one change that was not sent yet was lost.' : `The file was loaded again; ${n} changes that were not sent yet were lost.`),
    refused: 'The server does not take changes from here; editing together has ended.',
    version: 'This file is being edited with another version; reload the page.',
    backlog: (n) => `${n} changes wait to be sent.`,
    badStep: 'Someone sent a change that is none for this workbook. It was left out for everybody.',
    readOnly: 'The server takes no changes from you: the workbook is read-only here.',
    detached: 'The file was changed elsewhere, or the room was started anew, while this workbook holds changes that are in no file. Nothing was thrown away: your changes are here, and nobody else sees them. Save the file to go on together.',
    waiting: 'The room cannot be joined yet: someone who is in it has to save the file first. Trying again…',
    copySheet: 'A sheet cannot be copied while the workbook is edited together.',
    chartCut: 'A chart cannot be cut while the workbook is edited together: it was copied. Paste it, then delete the chart.',
    chartCopyGone: 'The chart was changed or deleted since it was copied; while the workbook is edited together it cannot be pasted. Copy the chart again.',
    tooLarge: 'This change is too large to be shared in one step.',
    unsupported: 'This cannot be done while the workbook is edited together.',
    pictures: 'Pictures cannot be added, moved or deleted while the workbook is edited together.',
    sheetDeleted: (name) => `Sheet ${name} was deleted. This cannot be undone while the workbook is edited together.`,
    editGone: 'The cell you were typing in was deleted by someone else.',
    formulaGone: (text) => `Someone else moved cells or renamed a sheet while you were typing a formula: what you typed may have meant other cells, and was not applied. You typed: ${text}`,
    undoWait: 'Undo has to wait until your last changes have arrived, and they have not yet: nothing was undone. Try again.',
    slow: 'A change has been on its way for a long time; the connection is made anew.',
    updating: 'The file was changed elsewhere and is loaded again; it cannot be edited until that is done.',
    noStamp: 'The custom properties of this file cannot be read: it is saved without the mark of the room, and whoever opens it joins when someone in the room has saved again.',
    fork: {
      title: 'The stored file is not the one you opened',
      question: 'Your workbook is apart from the others, and the file was saved by someone else meanwhile. Load the stored file (your unsaved changes are lost; save a copy first if you need them), or save yours over it (what the others saved is lost).',
      load: 'Load the stored file',
      save: 'Save mine over it',
      download: 'Download my copy',
      downloaded: 'The download of your copy was started. Check that the file is there before you load the stored one.',
      downloadFailed: 'Your copy could not be downloaded here. Nothing was changed: your workbook is as it was.',
    },
    autoWaits: { room: 'Not saved yet: another participant saves for all', offline: 'Not saved: auto-save waits for the connection. Save to store the file now' },
    typedKept: (text) => `The file is loaded again, and what you were typing was not applied. You typed: ${text}`,
    dialogClosed: 'Someone else changed the sheet or its rules: the dialog was closed without applying anything.',
    doc: {
      dropped: (n) => (n === 1 ? 'One of your changes could not be applied because someone else changed the same place at the same time.' : `${n} of your changes could not be applied because someone else changed the same place at the same time.`),
      undoGone: 'This step cannot be undone any more: someone else changed the same paragraphs since.',
      undoNone: 'This step cannot be undone while the document is edited together.',
      partial: 'Someone else deleted, at the same time, part of the text you deleted or formatted. If they undo their deletion, that text comes back without your change.',
      peer: 'Something that is no change of this document arrived; editing together has ended. Your changes stay here.',
      badStep: 'Someone sent a change that is none for this document. It was left out for everybody.',
      readOnly: 'The server takes no changes from you: the document is read-only here.',
      detached: 'The file was changed elsewhere, or the room was started anew, while this document holds changes that are in no file. Nothing was thrown away: your changes are here, and nobody else sees them. Save the file to go on together.',
      notShared: 'This cannot be done while the document is edited together.',
      tracked: 'This document tracks its changes, or holds tracked changes: it cannot be edited together yet. It is edited here alone.',
      unsupported: 'Only a docx file can be edited together: this one is edited here alone.',
      forkQuestion: 'Your document is apart from the others, and the file was saved by someone else meanwhile. Load the stored file (your unsaved changes are lost; save a copy first if you need them), or save yours over it (what the others saved is lost).',
    },
  },
  sheetDeleted: (name) => `Sheet ${name} was deleted. Undo brings it back.`,
  referenceAdded: (reference) => `Reference ${reference}`,
  editSheetGone: 'The sheet of the cell is gone; what was typed was not written.',
  badSheetName: 'A sheet name has 1 to 31 characters, none of [ ] : * ? / \\, no apostrophe at its start or end, and is not "History"',
  sheetNameTaken: (name) => `Another sheet is already named "${name}"`,
  cannotDeleteLastSheet: 'The last visible sheet cannot be deleted',
  sheetKindFixed: (action) => `This kind of sheet cannot be ${action === 'rename' ? 'renamed' : action === 'copy' ? 'copied' : action === 'color' ? 'given a tab colour' : 'deleted'} yet`,
  sheetBlockedBy: (action, what) =>
    what === 'table'
      ? `A sheet with a table cannot be ${action === 'rename' ? 'renamed' : 'deleted'} yet`
      : `Sheets cannot be ${action === 'rename' ? 'renamed' : 'deleted'} yet in a workbook with ${what === 'pivotTable' ? 'a pivot table' : what === 'chart' ? 'a chart' : what}`,
  sheetFormulaUnread: (action, where) => `The sheet cannot be ${action === 'rename' ? 'renamed' : 'deleted'}: the formula of ${where} cannot be read`,
  sheetNamesClash: (first, second) => `The sheets of this workbook cannot be changed: "${first}" and "${second}" differ only by case`,
  rules: {
    validationCommand: 'Data validation\u2026',
    validationTitle: 'Data validation',
    ok: 'OK',
    cancel: 'Cancel',
    close: 'Close',
    clearAll: 'Clear all',
    settings: 'Settings',
    inputMessage: 'Input message',
    errorAlert: 'Error alert',
    allow: 'Allow',
    types: { none: 'Any value', whole: 'Whole number', decimal: 'Decimal', list: 'List', date: 'Date', time: 'Time', textLength: 'Text length', custom: 'Custom formula' },
    data: 'Data',
    operators: {
      between: 'between',
      notBetween: 'not between',
      equal: 'equal to',
      notEqual: 'not equal to',
      greaterThan: 'greater than',
      lessThan: 'less than',
      greaterThanOrEqual: 'greater than or equal to',
      lessThanOrEqual: 'less than or equal to',
    },
    minimum: 'Minimum',
    maximum: 'Maximum',
    value: 'Value',
    startDate: 'Start date',
    endDate: 'End date',
    date: 'Date',
    startTime: 'Start time',
    endTime: 'End time',
    time: 'Time',
    source: 'Source',
    formula: 'Formula',
    sourceHint: (separator) => `Items separated by "${separator}", or cells: =$A$1:$A$5`,
    ignoreBlank: 'Ignore blank',
    inCellDropdown: 'In-cell dropdown',
    showInputMessage: 'Show input message when the cell is selected',
    showErrorAlert: 'Show error alert after invalid data is entered',
    style: 'Style',
    styles: { stop: 'Stop', warning: 'Warning', information: 'Information' },
    title: 'Title',
    message: 'Message',
    mixedSelection: 'The selection contains more than one rule, or cells without one. OK replaces them all with this rule.',
    needNumber: (field) => `${field}: enter a number or a formula`,
    needWholeNumber: (field) => `${field}: enter a whole number or a formula`,
    needDate: (field) => `${field}: enter a date or a formula`,
    needTime: (field) => `${field}: enter a time or a formula`,
    needValue: (field) => `${field} is required`,
    needItems: 'Enter the items of the list, or the cells that hold them',
    itemWithComma: 'An item of a list cannot contain a comma',
    listTooLong: (limit) => `The items of a list cannot be longer than ${limit} characters together; use cells for a longer list`,
    textTooLong: (field, limit) => `${field} cannot be longer than ${limit} characters`,
    badRuleFormula: (field) => `${field}: the formula cannot be read`,
    extensionRule: 'A rule kept as it is applies to these cells (its list is on another sheet); it cannot be changed here',
    tooManyValidations: 'The sheet has too many data validations',
    tooManyRanges: 'The rule would be left with too many separate ranges',
    formulaBeingTyped: 'Finish the formula first',
    dialogOpen: 'Confirm or close the open rule dialog first',
    pointOwnSheet: 'A rule applies to cells of its own sheet',
    validationSet: 'Data validation set',
    validationCleared: 'Data validation cleared',
    conditionalCommand: 'Conditional formatting',
    quickRules: {
      greaterThan: 'Greater than\u2026',
      lessThan: 'Less than\u2026',
      between: 'Between\u2026',
      equal: 'Equal to\u2026',
      containsText: 'Text that contains\u2026',
      duplicateValues: 'Duplicate values\u2026',
      top10: 'Top 10 items\u2026',
      colorScale: 'Colour scales\u2026',
      dataBar: 'Data bars\u2026',
    },
    clearFromSelection: 'Clear rules from selected cells',
    clearFromSheet: 'Clear rules from this sheet',
    manageRules: 'Manage rules\u2026',
    conditionalTitle: 'Conditional formatting rule',
    managerTitle: 'Conditional formatting rules',
    ruleType: 'Rule type',
    conditionalTypes: {
      cellIs: 'Cell value',
      containsText: 'Text contains',
      notContainsText: 'Text does not contain',
      beginsWith: 'Text begins with',
      endsWith: 'Text ends with',
      duplicateValues: 'Duplicate values',
      uniqueValues: 'Unique values',
      top10: 'Top or bottom values',
      aboveAverage: 'Above or below average',
      containsBlanks: 'Blank cells',
      notContainsBlanks: 'Cells that are not blank',
      containsErrors: 'Cells with an error',
      notContainsErrors: 'Cells without an error',
      expression: 'Formula',
      colorScale: 'Colour scale',
      dataBar: 'Data bar',
      iconSet: 'Icon set',
      timePeriod: 'Dates occurring',
    },
    text: 'Text',
    rank: 'Number',
    percentOfRange: 'Percent of the range',
    lowestValues: 'Lowest values',
    averageSide: 'Values',
    averages: { above: 'above the average', below: 'below the average' },
    includeAverage: 'Include the average itself',
    stopIfTrue: 'Stop if true',
    format: 'Format',
    formatWith: 'Format with',
    presets: {
      lightRed: 'Light red fill with dark red text',
      yellow: 'Yellow fill with dark yellow text',
      green: 'Green fill with dark green text',
      keep: 'The format it has',
      custom: 'Custom format',
      none: 'No format',
    },
    textColor: 'Text colour',
    fillColor: 'Fill colour',
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    strike: 'Strikethrough',
    scaleKind: 'Colours',
    scales: { '2': 'Two colours', '3': 'Three colours' },
    lowest: 'Lowest',
    midpoint: 'Midpoint',
    highest: 'Highest',
    barColor: 'Bar colour',
    iconStyle: 'Icon style',
    iconSets: {
      '3Arrows': '3 arrows (coloured)',
      '3ArrowsGray': '3 arrows (grey)',
      '3Flags': '3 flags',
      '3TrafficLights1': '3 traffic lights',
      '3TrafficLights2': '3 traffic lights (rimmed)',
      '3Signs': '3 signs',
      '3Symbols': '3 symbols (circled)',
      '3Symbols2': '3 symbols',
      '3Stars': '3 stars',
      '3Triangles': '3 triangles',
      '4Arrows': '4 arrows (coloured)',
      '4ArrowsGray': '4 arrows (grey)',
      '4RedToBlack': 'Red to black',
      '4Rating': '4 ratings',
      '4TrafficLights': '4 traffic lights',
      '5Arrows': '5 arrows (coloured)',
      '5ArrowsGray': '5 arrows (grey)',
      '5Rating': '5 ratings',
      '5Quarters': '5 quarters',
      '5Boxes': '5 boxes',
    },
    iconPreview: 'The icons of the set, from the lowest values to the highest',
    reverseIcons: 'Reverse the order of the icons',
    iconOnly: 'Show the icon only',
    period: 'Period',
    showRulesFor: 'Show rules for',
    scopeSelection: 'Current selection',
    scopeSheet: 'This sheet',
    appliesTo: 'Applies to',
    moveUp: 'Move up',
    moveDown: 'Move down',
    edit: 'Edit',
    deleteRule: 'Delete',
    newRule: 'New rule\u2026',
    noRules: 'No rules',
    rulesShown: (first, last, total) => `Rules ${first} to ${last} of ${total}`,
    nextRules: (count) => `Next ${count}`,
    previousRules: (count) => `Previous ${count}`,
    ruleRow: (place, description) => `Rule ${place}: ${description}`,
    cellValue: 'Cell value',
    and: 'and',
    textRules: { containsText: 'Text contains', notContainsText: 'Text does not contain', beginsWith: 'Text begins with', endsWith: 'Text ends with' },
    topItems: (count, percent, bottom) => `${bottom ? 'Bottom' : 'Top'} ${count}${percent ? '%' : ''}`,
    average: (above, equal) => `${above ? 'Above' : 'Below'}${equal ? ' or equal to' : ''} average`,
    formulaIs: 'Formula',
    datesOccurring: 'Dates occurring',
    periods: {
      today: 'today',
      yesterday: 'yesterday',
      tomorrow: 'tomorrow',
      last7Days: 'in the last 7 days',
      thisWeek: 'this week',
      lastWeek: 'last week',
      nextWeek: 'next week',
      thisMonth: 'this month',
      lastMonth: 'last month',
      nextMonth: 'next month',
    },
    needRange: 'Enter the cells the rule applies to',
    badRange: 'These are not ranges of the sheet; write them like A1:B10',
    needWholeUpTo: (field, limit) => `${field}: enter a whole number from 1 to ${limit}`,
    extensionRuleKept: 'This rule has settings of a newer Excel, which are kept as they are; it cannot be changed or deleted here',
    extensionRuleSays: 'What this rule says has settings of a newer Excel and is kept as it is; its cells and its place can be changed, and it can be deleted',
    orderFixed: 'The order of the rules of this sheet cannot be changed: one of them is kept as it is, with its place',
    ruleKindKept: 'A rule of this kind is kept as it is; it can be moved and deleted, not changed',
    tooManyConditional: 'The sheet has too many conditional formats',
    ruleAdded: 'Conditional format added',
    ruleChanged: 'Conditional format changed',
    ruleDeleted: 'Conditional format deleted',
    rulesCleared: 'Conditional formats cleared',
  },
  charts: {
    insert: 'Insert chart',
    types: { column: 'Column', stackedColumn: 'Stacked column', bar: 'Bar', line: 'Line', area: 'Area', pie: 'Pie', doughnut: 'Doughnut', scatter: 'Scatter' },
    edit: 'Edit chart\u2026',
    remove: 'Delete chart',
    dialogTitle: 'Chart',
    type: 'Type',
    otherType: 'Another type',
    grouping: 'Series',
    groupings: { standard: 'Side by side', stacked: 'Stacked', percentStacked: 'Stacked to 100%' },
    title: 'Title',
    range: 'Data range',
    seriesIn: 'Series in',
    directions: { auto: 'Automatic', columns: 'Columns', rows: 'Rows' },
    legend: 'Legend',
    legends: { r: 'Right', l: 'Left', t: 'Top', b: 'Bottom', tr: 'Top right', none: 'None' },
    dataLabels: 'Show the values as labels',
    axes: 'Axes',
    categoryTitle: 'Title of the category axis',
    valueTitle: 'Title of the value axis',
    minimum: 'Smallest value',
    maximum: 'Largest value',
    series: 'Series',
    categories: 'Category labels',
    seriesName: (n) => `Name of series ${n}`,
    seriesValues: (n) => `Values of series ${n}`,
    seriesColor: (n) => `Colour of series ${n}`,
    ownColor: (n) => `Series ${n} has a colour of its own`,
    moveUp: (n) => `Move series ${n} up`,
    moveDown: (n) => `Move series ${n} down`,
    removeSeries: (n) => `Remove series ${n}`,
    addSeries: 'Add series',
    ok: 'OK',
    cancel: 'Cancel',
    selected: (name) => (name === '' ? 'Chart selected' : `Chart selected: ${name}`),
    added: 'Chart added',
    changed: 'Chart changed',
    moved: 'Chart moved',
    removed: 'Chart deleted',
    needRange: 'Enter the range of cells the chart shows',
    badRange: 'This is not a range of cells',
    badNumber: 'Enter a number, or leave the field empty',
    needValues: 'A series needs a range of values',
    typeFixed: 'The type of this chart cannot be changed: its file holds formatting another type cannot keep',
    notEditable: 'This chart can be moved and deleted, not changed: its file holds what cannot be kept',
    noChart: 'No chart is selected',
    seriesFixed: 'The series of this chart cannot be added, removed or reordered: its type or its file does not allow it',
    groupingFixed: 'The series of this type of chart cannot be stacked',
    noAxis: 'This chart has no such axis',
    tooManySeries: (most) => `A chart has at most ${most} series`,
    tooManyCharts: 'No more charts can be added here',
    textTooLong: (most) => `This text cannot be longer than ${most} characters`,
    badCharacter: 'The text holds a character a file cannot hold',
    refUnread: 'The sheet cannot be changed: a chart names it in a reference that cannot be read',
    placesFixed: 'Charts cannot be added, moved or deleted on this sheet: its drawing in the file cannot be changed',
    wrapped: 'This chart cannot be deleted: in the file it is bound to another object',
    duplicate: 'Duplicate',
    copied: 'Chart copied',
    cutOut: 'Chart cut',
    pasted: 'Chart pasted',
    duplicated: 'Chart duplicated',
    notCopyable: 'This chart cannot be copied: what its file holds cannot be kept in a copy',
    copyOutdated: 'The chart cannot be pasted: rows, columns or sheets changed since it was copied. Copy it again.',
    copyLeftOut: (what) => `The copy was made without ${what.map((name) => (name === 'userShapes' ? 'the shapes drawn over the chart' : name === 'externalData' ? 'the data kept outside the workbook' : name)).join(' and ')}`,
    copyUnbound: (count) => (count === 1 ? 'One reference of the copy still names no sheet' : `${count} references of the copy still name no sheet`),
    noCellPaste: 'A chart is selected: select a cell to paste cells',
    copyGone: 'The chart on the clipboard was copied elsewhere and cannot be pasted here: copy the chart again',
  },
  pictures: {
    insert: 'Insert picture',
    added: 'Picture inserted',
    moved: 'Picture moved',
    removed: 'Picture deleted',
    selected: (name) => (name ? `Picture selected: ${name}` : 'Picture selected'),
    remove: 'Delete picture',
    notSupported: 'This picture cannot be inserted: only PNG, JPEG, GIF and BMP pictures can',
    tooLarge: 'This picture is too large to be inserted',
    tooMany: 'No more pictures can be inserted into this workbook',
    protectedSheet: 'The sheet is protected: its pictures cannot be changed',
    placesFixed: 'Pictures cannot be added, moved or deleted on this sheet: its drawing in the file cannot be changed',
    wrapped: 'This picture cannot be deleted: it shares its place in the file with another object',
    noCopy: 'Pictures cannot be copied yet',
    noCellPaste: 'A picture is selected: select a cell to paste cells',
    dropRefused: 'A picture cannot be dropped here now',
  },
  pivots: {
    refresh: 'Refresh pivot table',
    refreshAll: 'Refresh all pivot tables',
    refreshed: (count) => (count === 1 ? 'Pivot table refreshed' : `${count} pivot tables refreshed`),
    refused: (message) => message,
    none: 'The workbook has no pivot table to refresh',
    insert: 'Insert pivot table…',
    createTitle: 'Insert pivot table',
    source: 'Data',
    place: 'Place',
    newSheet: 'New sheet',
    thisSheet: 'This sheet, at a cell',
    cell: 'Cell',
    ok: 'OK',
    cancel: 'Cancel',
    badSource: 'Type a range such as A1:D20, or the name of one',
    badCell: 'Type a cell such as F2',
    made: 'Pivot table inserted',
    fields: 'Pivot table fields…',
    panelTitle: 'Pivot table fields',
    close: 'Close',
    fieldList: 'Fields',
    boxes: { filters: 'Filters', cols: 'Columns', rows: 'Rows', values: 'Values' },
    dropHere: 'Drag fields here',
    moveUp: (name) => `Move ${name} up`,
    moveDown: (name) => `Move ${name} down`,
    moveTo: (name) => `Where ${name} stands`,
    remove: (name) => `Remove ${name}`,
    summarize: (name) => `Summarize ${name} by`,
    numberFormat: (name) => `Number format of ${name}`,
    generalFormat: 'General',
    aggregates: { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count numbers', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' },
    subtotal: 'Subtotals',
    subtotals: (name) => `Subtotals of ${name}`,
    sort: (name) => `Order of ${name}`,
    sorts: { manual: 'As in the data', ascending: 'A to Z', descending: 'Z to A' },
    shownItem: (name) => `Item of ${name} shown`,
    allItems: '(All)',
    filterItems: (name) => `Items of ${name}`,
    moreItems: (count) => `${count} more items are not listed`,
    rowGrandTotals: 'Grand totals of rows',
    colGrandTotals: 'Grand totals of columns',
    changed: 'Pivot table changed',
    noTable: 'Select a cell of a pivot table',
  },
  find: {
    title: 'Find and replace',
    findLabel: 'Find',
    replaceLabel: 'Replace with',
    previous: 'Previous match (Shift+Enter)',
    next: 'Next match (Enter)',
    replace: 'Replace',
    replaceAll: 'Replace all',
    toggleReplace: 'Show replace',
    close: 'Close (Esc)',
    matchCase: 'Match case',
    wholeWord: 'Whole words only',
    entireCell: 'Entire cell contents',
    mainOnly: 'Main text only',
    lookIn: 'Look in',
    lookInFormulas: 'Formulas',
    lookInValues: 'Values',
    within: 'Within',
    withinSheet: 'Sheet',
    withinWorkbook: 'Workbook',
    searching: 'Searching…',
    noMatches: 'No matches',
    matches: (count, more) => (more ? `More than ${count} matches` : count === 1 ? '1 match' : `${count} matches`),
    replaced: (count, skipped) => (count === 1 ? '1 replacement' : `${count} replacements`) + (skipped > 0 ? `; ${skipped} could not be replaced` : ''),
    inTextBoxes: (count, why) =>
      ` (${count} in text boxes: ${why === 'shared' ? 'they cannot be changed while the file is edited together' : why === 'tracked' ? 'changes there cannot be tracked' : 'the file keeps them in a form that cannot be changed without loss'})`,
    inComments: (count) => ` (${count} in comments: they cannot be replaced while the file is edited together)`,
    outsideComments: (count) => ` (${count} outside comments: only comments can be changed in this file)`,
    replacedInSteps: (steps) => `Replaced in ${steps} steps, each undone on its own, because the file is edited together`,
    changedMeanwhile: 'The file changed during the search: try again',
    tooMany: (limit) => `There are more than ${limit} matches to replace at once: narrow the search`,
    readOnly: 'Nothing can be replaced: the file is read-only',
  },
  pageSetup: {
    title: 'Page setup',
    ok: 'OK',
    cancel: 'Cancel',
    actions: 'Breaks',
    unit: 'in',
    decimal: '.',
    page: 'Page',
    paper: 'Paper size',
    otherPaper: 'Other (as it is)',
    orientation: 'Orientation',
    portrait: 'Portrait',
    landscape: 'Landscape',
    scaling: 'Scaling',
    adjustTo: 'Adjust to',
    fitTo: 'Fit to pages',
    scale: '% of normal size',
    fitWide: 'Pages wide',
    fitTall: 'Pages tall (0: as many as needed)',
    margins: 'Margins',
    top: 'Top',
    bottom: 'Bottom',
    left: 'Left',
    right: 'Right',
    header: 'Header',
    footer: 'Footer',
    centerOnPage: 'Centre on page',
    horizontally: 'Horizontally',
    vertically: 'Vertically',
    sheet: 'Sheet',
    printArea: 'Print area',
    titleRows: 'Rows to repeat at top',
    titleCols: 'Columns to repeat at left',
    gridLines: 'Gridlines',
    headings: 'Row and column headings',
    insertBreak: 'Insert page break',
    removeBreak: 'Remove page break',
    resetBreaks: 'Reset all page breaks',
    applyTo: 'Apply to',
    thisSection: 'This section',
    wholeDocument: 'Whole document',
    sectionBreak: (kind) => ({ nextPage: 'Section break: next page', continuous: 'Section break: continuous', evenPage: 'Section break: even page', oddPage: 'Section break: odd page' })[kind],
    badLength: 'Type a length of 0 or more',
    badScale: 'Type a whole percent from 10 to 400',
    badFit: 'Type a whole number of pages from 0 to 32767',
    badArea: 'Type the print area as references, such as A1:D20',
    badRows: 'Type the rows to repeat, such as $1:$2',
    badCols: 'Type the columns to repeat, such as $A:$B',
    noRoom: 'The margins leave no room on the paper',
    shared: 'The page setup is not changed while the file is edited together',
    tracked: 'The page setup is not changed while changes are tracked: turn tracking off first',
  },
  preview: {
    title: 'Print preview',
    pages: 'Pages',
    page: (number, total) => `Page ${number} of ${total}`,
    counter: (number, total) => `${number} / ${total}`,
    previous: 'Previous page',
    next: 'Next page',
    zoomIn: 'Zoom in',
    zoomOut: 'Zoom out',
    fit: 'Fit page',
    actualSize: 'Actual size (100%)',
    zoom: (percent) => `Zoom ${percent}%`,
    percent: (value) => `${value}%`,
    print: 'Print',
    pdf: 'Download as PDF',
    pageSetup: 'Page setup…',
    close: 'Close print preview',
    which: 'What to print',
    active: 'Active sheet',
    all: 'Entire workbook',
    planning: 'Laying out the pages…',
  },
};

const satir = (count: number) => (count === 1 ? 'satır' : `${count} satır`);
const sutun = (count: number) => (count === 1 ? 'sütun' : `${count} sütun`);

export const STRINGS_TR: UiStrings = {
  spreadsheet: 'Hesap tablosu',
  cellEditor: 'Hücre düzenleyici',
  actionFailed: 'İşlem tamamlanamadı',
  sheet: (name) => `Sayfa ${name}`,
  chartSheetNote: 'Bu sayfa bir grafik sayfası. Grafiği henüz gösterilmeyen bir türde.',
  chartSeries: (n) => `Seri${n}`,
  chartWord: (word) => (word === 'increase' ? 'Artış' : word === 'decrease' ? 'Azalış' : 'Toplam'),
  otherSheetNote: 'Bu tür sayfalar henüz gösterilmiyor.',
  selected: (range) => `${range} seçili`,
  insertRowsAbove: (count) => `Üste ${satir(count)} ekle`,
  insertRowsBelow: (count) => `Alta ${satir(count)} ekle`,
  deleteRows: (count) => (count === 1 ? 'Satırı sil' : `${count} satırı sil`),
  insertColumnsLeft: (count) => `Sola ${sutun(count)} ekle`,
  insertColumnsRight: (count) => `Sağa ${sutun(count)} ekle`,
  deleteColumns: (count) => (count === 1 ? 'Sütunu sil' : `${count} sütunu sil`),
  cut: 'Kes',
  copy: 'Kopyala',
  paste: 'Yapıştır',
  clearContents: 'İçeriği temizle',
  pasteWithKeyboard: 'Pano buradan okunamıyor. Yapıştırmak için Ctrl+V kullanın.',
  tooManyCells: 'Bir kerede biçimlendirilemeyecek kadar çok hücre seçili',
  tooManyRows: (action) =>
    action === 'format'
      ? 'Bir kerede biçimlendirilemeyecek kadar çok satır seçili'
      : action === 'resize'
        ? 'Bir kerede boyutlandırılamayacak kadar çok satır seçili'
        : 'Bir kerede değiştirilemeyecek kadar çok satır seçili',
  cannotMergeWhole: 'Tam satırlar ya da sütunlar birleştirilemez',
  cannotMergeArray: (ref) => `${ref} aralığındaki dizi formülünün içinden hücre birleştirilemez`,
  cannotMergeTable: 'Tablo içeren bir sayfada hücre birleştirme henüz desteklenmiyor',
  cannotMergePivot: (name) => `"${name}" özet tablosunun hücreleri birleştirilemez`,
  cannotMergeFiltered: 'Süzgecin gizlediği satırların üzerinden hücre birleştirilemez',
  mergeLossTitle: 'Hücreleri birleştir',
  mergeLossQuestion: (count) => `Birleştirme yalnız sol üst hücrenin değerini korur. Diğer ${count} hücrenin değeri silinir.`,
  mergeConfirm: 'Birleştir',
  panesFrozen: (rows, cols) => (rows === 0 && cols === 0 ? 'Bölmelerin dondurulması kaldırıldı' : `Dondurulan: ${rows} satır, ${cols} sütun`),
  tabs: {
    title: 'Sekmeler',
    stops: 'Sekme durakları',
    none: 'Sekme durağı yok: varsayılanlar geçerli',
    position: 'Konum',
    alignment: 'Hizalama',
    leader: 'Dolgu',
    aligns: { left: 'Sol', center: 'Orta', right: 'Sağ', decimal: 'Ondalık', bar: 'Çubuk' },
    leaders: { none: 'Yok', dot: 'Noktalar', hyphen: 'Tireler', underscore: 'Çizgi', heavy: 'Kalın çizgi', middleDot: 'Orta noktalar' },
    set: 'Ekle',
    clear: (stop) => `${stop} konumundaki sekme durağını kaldır`,
    clearAll: 'Tümünü temizle',
    defaultStops: 'Varsayılan sekme aralığı',
    ok: 'Tamam',
    cancel: 'İptal',
    unit: 'cm',
    decimal: ',',
    badPosition: '-55,88 ile 55,88 cm arasında bir konum yazın',
    badDefault: '0 ile 55,88 cm arasında bir aralık yazın',
    defaultShared: 'Belge birlikte düzenlenirken değiştirilemez',
  },
  insertPushesOff: (count, axis) => `${count} ${axis === 'row' ? 'satır' : 'sütun'} eklenemiyor: hücreler sayfanın dışına itilirdi`,
  arrayFormulaBlocks: (kind, axis, ref) =>
    `${ref} aralığındaki dizi formülünün içinden ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}`,
  notSupportedWith: (kind, axis, what) =>
    `Bu sayfada ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'ekleme' : 'silme'} henüz desteklenmiyor (sayfada bulunan: ${what})`,
  pivotCut: (kind, axis, name) => `"${name}" özet tablosunun içinden ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}`,
  pivotSourceCut: (axis, name) =>
    axis === 'row' ? `"${name}" özet tablosunun kaynağındaki başlık satırı silinemez` : `"${name}" özet tablosunun kaynağındaki alan sütunları silinemez`,
  pivotPushedOff: (kind, axis, name) => `${axis === 'row' ? 'Satır' : 'Sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}: "${name}" özet tablosu sayfanın dışına itilirdi`,
  pivotSourceNotRange: (kind, axis, name) => `${axis === 'row' ? 'Satır' : 'Sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}: "${name}" özet tablosunun kaynağı tek bir hücre aralığı değil`,
  pivotsFixed: (kind, axis, reason) => `Bu çalışma kitabında özet tabloları yüzünden ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}: ${pivotReasonTr(reason)}`,
  pasteTooLarge: (max) => `Bir kerede ${max} hücreden fazlası yapıştırılamaz`,
  tooManyFormats: 'Çalışma kitabında çok fazla hücre biçimi var',
  badFormula: 'Bu formül okunamadı. Argümanları noktalı virgülle (;) ayırın, ondalıkları virgülle yazın; tanınmayan fonksiyonlarda İngilizce adı kullanın.',
  copyTooLarge: 'Aralık kopyalanamayacak kadar büyük',
  filterMenu: (column) => `${column} sütununu sırala ve süz`,
  sortAscending: "A'dan Z'ye sırala",
  sortDescending: "Z'den A'ya sırala",
  clearFilter: 'Süzgeci temizle',
  searchValues: 'Ara',
  selectAll: '(Tümünü seç)',
  blanks: '(Boşlar)',
  apply: 'Uygula',
  cancel: 'İptal',
  noMatches: 'Eşleşen değer yok',
  valueList: 'Değerler',
  tooManyValues: (max) => `Yalnız ilk ${max} değer listeleniyor`,
  recordsShown: (shown, total) => `${total} kayıttan ${shown} tanesi gösteriliyor`,
  filterReplaced: 'Bu sütunun renk ya da simge süzgecinin yerini seçilen değerler aldı.',
  nothingToFilter: 'Burada süzülecek veri yok',
  cannotSortMerged: 'Birleştirilmiş hücre içeren aralık sıralanamaz',
  cannotSortArray: (ref) => `${ref} aralığındaki dizi formülü bölünerek sıralanamaz`,
  sortNotSupportedWith: (action, what) => `Bu sayfada ${action === 'sort' ? 'sıralama' : 'süzme'} henüz desteklenmiyor (sayfada bulunan: ${what})`,
  chooseFromList: 'Listeden seçin',
  validationFailed: 'Bu değer hücrenin veri doğrulama kuralına uymuyor.',
  validationWarning: (message) => `${message} (değer yine de yazıldı)`,
  formulaNotCalculated: 'Bu formül burada hesaplanamıyor; dosyadaki son değer gösteriliyor',
  formulaCircular: 'Döngüsel başvuru',
  formulaStale: 'Hesap bitmedi; değer eski olabilir',
  comment: 'Yorum',
  addNote: 'Not ekle',
  editNote: 'Notu düzenle',
  deleteNote: 'Notu sil',
  noteField: 'Not',
  noteHint: 'Ctrl+Enter uygular, Esc vazgeçer',
  noteDeleted: 'Not silindi',
  noteThreaded: (ref) => `${ref} hücresindeki yorum bir yazışma: burada düzenlenemez`,
  noteTooLong: (max) => `Bir not ${max} karakterden uzun olamaz`,
  tooManyNotes: 'Sayfada çok fazla not var',
  notesBlocked: (kind) =>
    `Bu sayfanın notları değiştirilemez: ${kind === 'shared' ? 'aynı notları başka bir sayfa da gösteriyor' : kind === 'parts' ? 'dosya onları birden çok parçada tutuyor' : 'not kutularının çizimi okunamıyor'}`,
  drawingsNotShown: 'Bu sayfadaki bazı çizimler ya da yorumlar okunamadı; gösterilmiyor',
  document: 'Belge',
  documentTools: 'Belge araçları',
  page: (number) => `Sayfa ${number}`,
  pageOf: (page, count) => `Sayfa ${page} / ${count}`,
  zoomIn: 'Yakınlaştır',
  zoomOut: 'Uzaklaştır',
  fitWidth: 'Genişliğe sığdır',
  zoomLevel: (percent) => `Yakınlaştırma %${percent}, %100'e döndür`,
  preparingPages: 'Sayfalar hazırlanıyor\u2026',
  documentTooLong: 'Belge çok uzun: yalnız ilk sayfaları gösteriliyor',
  documentLayoutFailed: 'Belgenin tamamı sayfalara yerleştirilemedi',
  placeholder: (kind) => (kind === 'chart' ? 'Grafik' : kind === 'diagram' ? 'Diyagram' : kind === 'object' ? 'Gömülü nesne' : kind === 'picture' ? 'Görsel' : 'Çizim'),
  documentEditor: 'Belge düzenleyici',
  cannotEditHere: 'Belgenin bu bölümü düzenlenemez',
  textBoxRefused: (why) =>
    why === 'locked'
      ? 'Bu metin kutusunun metni kilitli; hiçbir şey değiştirilmedi'
      : why === 'unwritable'
        ? 'Bu metin kutusu değiştirilemez: dosya kutunun değişikliği alamayan bir kopyasını daha tutuyor; hiçbir şey değiştirilmedi'
        : why === 'tracked'
          ? 'Değişiklikler izlenirken metin kutuları düzenlenemez; hiçbir şey değiştirilmedi'
          : why === 'shared'
            ? 'Belge birlikte düzenlenirken metin kutuları düzenlenemez; hiçbir şey değiştirilmedi'
            : 'Bu işlem metin kutusunda yapılamaz; hiçbir şey değiştirilmedi',
  textBoxEntered: (fixed) => (fixed ? 'Metni değiştirilemeyen bir metin kutusundasınız. Escape ile çıkılır' : 'Metin kutusundasınız. Escape ile çıkılır'),
  textBoxLeft: 'Metin kutusundan çıkıldı',
  cannotTrackChange: 'Bu belgede değişiklikler izleniyor ve bu tür değişiklik henüz izlenemiyor; hiçbir şey değiştirilmedi',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'Bu belge korumalı: yalnız yorum eklenip değiştirilebilir; metni düzenlenemez'
      : kind === 'forms'
        ? 'Bu belge korumalı: yalnız form alanları doldurulabilir. Salt okunur gösteriliyor'
        : 'Bu belge değişikliğe karşı korumalı. Salt okunur gösteriliyor',
  revisionDescription: (kind, author, date, text) => {
    const what: { [kind: string]: string } = {
      insert: 'Eklendi',
      delete: 'Silindi',
      insertColumn: 'Sütun eklendi',
      deleteColumn: 'Sütun silindi',
      moveFrom: 'Buradan taşındı',
      moveTo: 'Buraya taşındı',
      format: 'Biçim değişti',
      paragraphFormat: 'Paragraf biçimi değişti',
      tableFormat: 'Tablo biçimi değişti',
      sectionFormat: 'Sayfa düzeni değişti',
    };
    return `${what[kind] || 'Yalnız korunabilen bir değişiklik'}${author !== undefined && author !== '' ? ` (${author})` : ''}${date ? `, ${date}` : ''}${text !== '' ? `: ${text}` : ''}`;
  },
  revisionStep: (number, count, description, wrapped) => `${wrapped === 'first' ? 'İlk değişikliğe dönüldü. ' : wrapped === 'last' ? 'Son değişikliğe gelindi. ' : ''}Değişiklik ${number} / ${count}. ${description}`,
  noRevisions: 'Belgede izlenen değişiklik yok',
  noRevisionHere: 'Burada izlenen bir değişiklik yok',
  revisionsResolved: (verdict, resolved, left, skipped, sections) =>
    `${resolved} değişiklik ${verdict === 'accept' ? 'kabul edildi' : 'reddedildi'}. ${left === '0' ? 'Değişiklik kalmadı.' : `${left} değişiklik kaldı.`}` +
    (skipped > 0 ? ` ${skipped} değişiklik ${verdict === 'accept' ? 'kabul edilemedi' : 'reddedilemedi'}, olduğu gibi bırakıldı.` : '') +
    (sections > 0 ? ` ${sections} bölüm sonu, paragraf işaretiyle birlikte kaldırıldı.` : ''),
  cannotResolveRevision: 'Bu değişiklik henüz kabul ya da reddedilemiyor; hiçbir şey değiştirilmedi',
  revisionGone: 'Değişiklik artık belgede yok',
  revisionsProtected: 'Bu belgenin koruması değişikliklerin kabul ya da reddedilmesine izin vermiyor',
  trackingTurned: (on) => (on ? 'Değişiklikler artık izleniyor' : 'Değişiklikler artık izlenmiyor'),
  cannotStopTracking: 'Bu belgede izleme kapatılamaz',
  commentBy: (author) => (author !== '' ? `Yorum: ${author}` : 'Yorum'),
  commentStep: (number, count, author, text, wrapped) =>
    `${wrapped === 'first' ? 'İlk yoruma dönüldü. ' : wrapped === 'last' ? 'Son yoruma gelindi. ' : ''}Yorum ${number} / ${count}${author !== '' ? `, ${author}` : ''}${text !== '' ? `: ${text}` : ''}`,
  commentReplies: (count) => `${count} yanıt`,
  commentResolved: 'Çözüldü',
  commentEdit: 'Düzenle',
  commentReply: 'Yanıtla',
  commentResolve: 'Çöz',
  commentReopen: 'Yeniden aç',
  commentDelete: 'Sil',
  commentSave: 'Kaydet',
  commentCancel: 'Vazgeç',
  commentClose: 'Kapat',
  commentText: 'Yorum metni',
  commentReplyText: 'Yanıt metni',
  newComment: 'Yeni yorum',
  noComments: 'Belgede yorum yok',
  commentAdded: 'Yorum eklendi',
  commentChanged: 'Yorum değiştirildi',
  commentReplied: 'Yanıt eklendi',
  commentRemoved: 'Yorum silindi',
  commentResolvedSaid: (done) => (done ? 'Yorum çözüldü' : 'Yorum yeniden açıldı'),
  commentNeedsPlace: 'İmleci yorumun ilgili olduğu metne koyun ya da o metni seçin',
  cannotCommentHere: 'Buraya yorum eklenemez',
  commentFixed: 'Bu yorum burada değiştirilemez: işaretleri belgenin olduğu gibi korunan bir bölümünde',
  commentGone: 'Yorum artık belgede yok',
  commentTextNeeded: 'Yorum için en çok 32.767 karakterlik bir metin gerekir',
  commentsPanel: 'Yorumlar',
  commentsShowResolved: 'Çözülenleri göster',
  commentsPanelCount: (shown, hidden) => `${shown} yorum${hidden > 0 ? `, ${hidden} çözülmüş gizli` : ''}`,
  commentsPanelEmpty: (hidden) => (hidden > 0 ? `Bütün yorumlar çözüldü (${hidden} gizli)` : 'Belgede yorum yok'),
  commentsShowMore: (count) => `${count} tane daha göster`,
  commentNotAnchored: 'Hiçbir metne bağlı değil',
  commentAbout: 'İlgili metin',
  commentThreadLabel: (author, date, text, replies, done) =>
    `${author !== '' ? `Yorum: ${author}` : 'Yorum'}${date !== '' ? `, ${date}` : ''}${done ? ', çözüldü' : ''}${replies > 0 ? `, ${replies} yanıt` : ''}${text !== '' ? `: ${text}` : ''}`,
  cannotCutObjects: 'Seçimde kesilemeyen bir öğe var (dipnot, denklem, metin kutusu, içerik denetimi, izlenen değişiklik); yerinde bırakıldı',
  cannotWriteClipboard: 'Pano buradan yazılamıyor; hiçbir şey kesilmedi',
  cannotPasteHere: 'Kopyalanan içerik buraya yapıştırılamıyor',
  pasteWaiting: 'Kopyanın görselleri okunuyor; okununca yapıştırılacak',
  pasteInterrupted: 'Kopyanın görselleri okunurken seçim ya da belge değişti; hiçbir şey yapıştırılmadı',
  tabLeaves: 'Düzenleyiciden çıkmak için Sekme tuşuna basın',
  documentPasteTooLarge: 'Panodaki metin yapıştırılamayacak kadar uzun',
  openedConverted: (from, to) => `Bu ${from} dosyası düzenlenebilmesi için ${to} biçimine çevrildi; ${to} olarak kaydedilir`,
  notConverted: (from) => `Bu ${from} dosyası düzenlenebilecek biçime çevrilemedi; salt okunur gösteriliyor`,
  saved: 'Kaydedildi',
  saveFailed: (detail) => (detail ? `Dosya kaydedilemedi (${detail})` : 'Dosya kaydedilemedi'),
  saveInterrupted: 'Henüz kaydedilmedi: dosya yazılırken düzenleme hiç durmadı. Düzenleme durunca kaydedilecek.',
  saveRoomBusy: 'Henüz kaydedilmedi: bu dosyayı aynı anda birkaç kişi kaydediyor. Kendiliğinden kaydetme bekleyip yeniden deneyecek.',
  saveConflict: 'Dosya kaydedilmedi: sunucudaki dosya sizin açtığınızdan farklı (başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir)',
  saveConflictTitle: 'Dosya sunucuda değişmiş',
  saveConflictQuestion:
    'Dosya, siz açtıktan sonra sunucuda değişmiş (başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir). Kaydederseniz oradaki değişikliklerin üzerine yazılır.',
  saveOverwrite: 'Üzerine yaz',
  saveCancel: 'Vazgeç',
  passwordTitle: 'Parola gerekli',
  passwordQuestion: 'Bu dosya parola ile korunuyor.',
  passwordLabel: 'Parola',
  passwordWrong: (attempts) => (attempts > 1 ? `Parola yanlış (${attempts} deneme). Yeniden deneyin.` : 'Parola yanlış. Yeniden deneyin.'),
  passwordChecking: 'Parola denetleniyor…',
  passwordOpen: 'Aç',
  passwordCancel: 'Vazgeç',
  noIntegrityCheck: 'Bu dosyanın bütünlük denetimi yok: içeriğini parolayı bilmeyen biri değiştirmiş olabilir',
  saveRedirected: 'Dosya kaydedilmedi: sunucu isteği başka bir adrese yönlendirdi; oturumunuzun süresi dolmuş olabilir',
  saveConvertedNotAllowed: (to) => `Bu dosyadaki değişiklikler buraya kaydedilemez: dosya ${to} biçimine çevrildi ve sayfa bunun özgün dosyanın yerine kaydedilmesine izin vermedi`,
  saveDenied: 'Dosya kaydedilmedi: kaydetme yetkiniz yok',
  saveLocked: 'Dosya kaydedilmedi: kilitli',
  saveTimedOut: 'Dosya kaydedilemedi: zamanında yanıt gelmedi',
  saveNoConnection: 'Dosya kaydedilemedi: sunucuya ulaşılamadı',
  savePending: 'Dosya kaydedilmedi: yazılmakta olan uygulanamadı',
  autoSaveNoVersion: 'Kendiliğinden kaydetme çalışmıyor: sunucu dosyanın sürümünü bildirmedi. Kaydet düğmesiyle kaydedin.',
  embeddedFontNotLoaded: (family) => `Bu cihazda "${family}" fontu yok, yerine geçen font da yüklenemedi; metin başka bir fontla gösteriliyor. Font klasörünün adresini "fontsUrl" ile verin ya da fontu "fonts: false" ile kapatın`,
  fontNotLoaded: (family, late) => (late ? `"${family}" fontu hâlâ yükleniyor; gelene dek dosya onsuz gösteriliyor` : `"${family}" fontu yüklenemedi; dosya onsuz gösteriliyor`),
  pictureNotSupported: 'Bu görsel eklenemiyor',
  linkNotAllowed: 'Bu adres bağlantı olarak kullanılamaz',
  editingStory: (kind) => (kind === 'header' ? 'Üstbilgi düzenleniyor' : kind === 'footer' ? 'Altbilgi düzenleniyor' : kind === 'footnote' ? 'Dipnot düzenleniyor' : kind === 'endnote' ? 'Sonnot düzenleniyor' : 'Metin düzenleniyor'),
  blankLine: 'boş',
  space: 'boşluk',
  endOfDocument: 'Belgenin sonu',
  selectionTools: 'Seçim',
  selectWholeText: 'Tümünü seç',
  print: 'Yazdır',
  printPreparing: 'Yazdırmaya hazırlanıyor\u2026',
  printClosed: 'Yazdırma penceresi kapandı',
  printProgress: (done, total) => `Yazdırmaya hazırlanıyor\u2026 ${done} / ${total}`,
  printNotAvailable: 'Burada yazdırma kullanılamıyor',
  printTooLong: 'Belge bir kerede yazdırılamayacak kadar uzun',
  printFailed: 'Belge yazdırılamadı',
  printSheetTooLong: 'Sayfa bir kerede yazdırılamayacak kadar uzun',
  printSheetFailed: 'Sayfa yazdırılamadı',
  printNothing: 'Yazdırılacak bir şey yok',
  pdfDownload: 'PDF olarak indir',
  pdfPreparing: 'PDF hazırlanıyor\u2026',
  pdfReady: 'PDF hazır',
  pdfProgress: (done, total) => `PDF hazırlanıyor\u2026 ${done} / ${total}`,
  pdfNotAvailable: 'Burada PDF oluşturulamıyor',
  pdfTooLong: 'Bu, tek bir PDF yapılamayacak kadar uzun',
  pdfFailed: 'PDF oluşturulamadı',
  pdfNothing: 'PDF yapılacak bir şey yok',
  addSheet: 'Sayfa ekle',
  insertSheet: 'Sayfa ekle',
  renameSheet: 'Yeniden adlandır',
  deleteSheet: 'Sil',
  duplicateSheet: 'Kopyasını oluştur',
  moveSheetLeft: 'Sola taşı',
  moveSheetRight: 'Sağa taşı',
  hideSheet: 'Gizle',
  unhideSheet: 'Göster\u2026',
  hiddenSheets: 'Gizli sayfalar',
  tabColor: 'Sekme rengi\u2026',
  noTabColor: 'Renk yok',
  sheetMoved: (name, place, count) => `${name} sayfası ${count} sayfanın ${place}. sırasına taşındı`,
  sheetCopiedWithout: (name, what) => `${name} sayfası kopyalandı; şunlar kopyalanmadı: ${what}`,
  copyLoss: (what, count) => {
    const names: Record<string, string> = {
      conditionalFormatExtensions: 'yeni Excel sürümüne özgü koşullu biçim ayarı',
      ruleContent: 'içeriği yalnız kendi sayfasında durabilen kural',
      drawings: 'görsel ya da şekil',
      threadedComments: 'yorum zinciri',
      externalLinks: 'dış bağlantı',
      pageBreaks: 'sayfa sonu',
      table: 'tablo',
      pivotTable: 'pivot tablo',
      rowAttributes: 'kendi ayarı olan satır',
      sparklineGroups: 'mini grafik grubu',
      customSheetViews: 'özel görünüm listesi',
      ignoredErrors: 'yok sayılan hata listesi',
      cellWatches: 'izlenen hücre listesi',
      sortState: 'sıralama ayarı',
      phoneticPr: 'fonetik ayarı',
      outlinePr: 'anahat ayarı',
      outlineLevels: 'anahat düzeyi ayarı',
      codeName: 'kod adı',
      legacyDrawingHF: 'üstbilgi/altbilgi görsel takımı',
      picture: 'arka plan görseli',
    };
    return Object.prototype.hasOwnProperty.call(names, what) ? `${count} ${names[what]}` : what;
  },
  notesNotPasted: (count) => `${count} not yapıştırılmadı: hücresi not alamıyor`,
  rulesNotPasted: (count) => `${count} veri doğrulama ya da koşullu biçim kuralı hücrelerle birlikte olduğu gibi gelmedi: dosyadaki haliyle korunan kural yerinde bırakılır ve kopyalanmaz, bu sayfanın tutamadığı uzantı içeriği ise kuraldan çıkarılır`,
  cannotHideLastSheet: 'Görünen son sayfa gizlenemez',
  cannotSortThreaded: 'Yorum zinciri içeren aralık henüz sıralanamıyor',
  sheetName: 'Sayfa adı',
  collab: {
    guest: (n) => `Misafir ${n}`,
    state: { connecting: 'Bağlanıyor…', online: 'Birlikte düzenleniyor', offline: 'Çevrimdışı: değişiklikler bağlantı gelince gönderilir', detached: 'Kopuk: değişiklikleriniz yalnız burada. Dosyayı kaydedin', closed: 'Bağlı değil: değişiklikler burada kalır' },
    others: (n) => `${n} kişi daha burada`,
    peopleCount: (n) => `Bu dosyada ${n} kişi`,
    morePeople: (n) => `ve ${n} kişi daha`,
    you: '(siz)',
    viewOnly: 'yalnız görüntülüyor',
    onSheet: (name) => `${name} sayfasında`,
    dropped: (n) => (n === 1 ? 'Bir değişikliğiniz, ilgili yeri başkası sildiği ya da değiştirdiği için uygulanamadı.' : `${n} değişikliğiniz, ilgili yeri başkası sildiği ya da değiştirdiği için uygulanamadı.`),
    partial: 'Bir değişikliğinizin bir bölümü uygulanmadı: içinde olduğu satırları, sütunları ya da sayfayı başkası sildi.',
    undoGone: 'Bu adım artık geri alınamaz: ilgili yeri başkası değiştirdi ya da sildi.',
    undoNone: 'Bu adım, kitap birlikte düzenlenirken geri alınamaz.',
    undoLossy: 'Hücreler geri geldi; onlarla giden notlar, bağlantılar ve kurallar gelmedi.',
    skipped: 'Kitap birlikte düzenlenirken kurallar ve grafikler hücrelerle birlikte taşınmaz; hücreler değişti.',
    peer: 'Bu kitabın değişikliği olmayan bir ileti geldi; birlikte düzenleme sona erdi. Değişiklikleriniz burada duruyor.',
    stale: 'Dosya başka yerde değişti ve yeniden yüklenemedi; birlikte düzenleme sona erdi.',
    reloaded: (n) => (n === 1 ? 'Dosya yeniden yüklendi; henüz gönderilmemiş bir değişiklik kayboldu.' : `Dosya yeniden yüklendi; henüz gönderilmemiş ${n} değişiklik kayboldu.`),
    refused: 'Sunucu buradan değişiklik almıyor; birlikte düzenleme sona erdi.',
    version: 'Bu dosya başka bir sürümle düzenleniyor; sayfayı yenileyin.',
    backlog: (n) => `${n} değişiklik gönderilmeyi bekliyor.`,
    badStep: 'Biri bu kitabın değişikliği olmayan bir şey gönderdi. Herkeste atlandı.',
    readOnly: 'Sunucu sizden değişiklik almıyor: kitap burada salt okunur.',
    detached: 'Dosya başka yerde değişti ya da oda yeniden başladı; bu kitapta ise hiçbir dosyada olmayan değişiklikler var. Hiçbir şey atılmadı: değişiklikleriniz burada ve başkası görmüyor. Birlikte devam etmek için dosyayı kaydedin.',
    waiting: 'Odaya şu an katılınamıyor: odadakilerden birinin önce dosyayı kaydetmesi gerekiyor. Yeniden deneniyor…',
    copySheet: 'Kitap birlikte düzenlenirken sayfa kopyalanamaz.',
    chartCut: 'Kitap birlikte düzenlenirken grafik kesilemez: kopyalandı. Yapıştırın, sonra grafiği silin.',
    chartCopyGone: 'Grafik kopyalandıktan sonra değişti ya da silindi; kitap birlikte düzenlenirken yapıştırılamaz. Grafiği yeniden kopyalayın.',
    tooLarge: 'Bu değişiklik tek adımda paylaşılamayacak kadar büyük.',
    unsupported: 'Bu, kitap birlikte düzenlenirken yapılamaz.',
    pictures: 'Kitap birlikte düzenlenirken görsel eklenemez, taşınamaz ve silinemez.',
    sheetDeleted: (name) => `${name} sayfası silindi. Kitap birlikte düzenlenirken bu geri alınamaz.`,
    editGone: 'Yazdığınız hücreyi başkası sildi.',
    formulaGone: (text) => `Siz formül yazarken başkası hücreleri kaydırdı ya da bir sayfayı yeniden adlandırdı: yazdığınız başka hücreleri kastediyor olabilirdi, uygulanmadı. Yazdığınız: ${text}`,
    undoWait: 'Geri alma, son değişiklikleriniz yerine ulaşana dek bekler; henüz ulaşmadı: hiçbir şey geri alınmadı. Yeniden deneyin.',
    slow: 'Bir değişiklik uzun süredir yolda; bağlantı yeniden kuruluyor.',
    updating: 'Dosya başka yerde değişti ve yeniden yükleniyor; bitene dek düzenlenemez.',
    noStamp: 'Bu dosyanın özel özellikleri okunamıyor: dosya oda damgası olmadan kaydedildi; açan kişi odadaki biri yeniden kaydedince katılır.',
    fork: {
      title: 'Saklanan dosya sizin açtığınız değil',
      question: 'Kitabınız öbürlerinden ayrı düştü ve dosyayı bu arada başkası kaydetti. Saklanan dosyayı yükleyin (kaydedilmemiş değişiklikleriniz gider; gerekiyorsa önce kopyanızı indirin) ya da kendinizinkini üzerine kaydedin (öbürlerinin kaydettiği gider).',
      load: 'Saklanan dosyayı yükle',
      save: 'Benimkini üzerine kaydet',
      download: 'Kopyamı indir',
      downloaded: 'Kopyanızın indirilmesi başlatıldı. Saklanan dosyayı yüklemeden önce dosyanın indiğini denetleyin.',
      downloadFailed: 'Kopyanız burada indirilemedi. Hiçbir şey değişmedi: kitabınız olduğu gibi duruyor.',
    },
    autoWaits: { room: 'Henüz kaydedilmedi: herkes için başka bir katılımcı kaydediyor', offline: 'Kaydedilmedi: kendiliğinden kaydetme bağlantıyı bekliyor. Şimdi saklamak için Kaydet' },
    typedKept: (text) => `Dosya yeniden yükleniyor; yazmakta olduğunuz uygulanmadı. Yazdığınız: ${text}`,
    dialogClosed: 'Başkası sayfayı ya da kurallarını değiştirdi: kutu hiçbir şey uygulanmadan kapatıldı.',
    doc: {
      dropped: (n) => (n === 1 ? 'Bir değişikliğiniz, aynı yeri aynı anda başkası değiştirdiği için uygulanamadı.' : `${n} değişikliğiniz, aynı yeri aynı anda başkası değiştirdiği için uygulanamadı.`),
      undoGone: 'Bu adım artık geri alınamaz: aynı paragrafı sonradan başkası değiştirdi.',
      undoNone: 'Bu adım, belge birlikte düzenlenirken geri alınamaz.',
      partial: 'Sildiğiniz ya da biçimlediğiniz metnin bir bölümünü aynı anda başkası sildi. O silmesini geri alırsa metin sizin değişikliğiniz olmadan geri gelir.',
      peer: 'Bu belgenin değişikliği olmayan bir ileti geldi; birlikte düzenleme sona erdi. Değişiklikleriniz burada duruyor.',
      badStep: 'Biri bu belgenin değişikliği olmayan bir şey gönderdi. Herkeste atlandı.',
      readOnly: 'Sunucu sizden değişiklik almıyor: belge burada salt okunur.',
      detached: 'Dosya başka yerde değişti ya da oda yeniden başladı; bu belgede ise hiçbir dosyada olmayan değişiklikler var. Hiçbir şey atılmadı: değişiklikleriniz burada ve başkası görmüyor. Birlikte devam etmek için dosyayı kaydedin.',
      notShared: 'Bu, belge birlikte düzenlenirken yapılamaz.',
      tracked: 'Bu belge değişikliklerini izliyor ya da izlenen değişiklik taşıyor: henüz birlikte düzenlenemez. Burada tek başına düzenleniyor.',
      unsupported: 'Yalnız docx dosyası birlikte düzenlenebilir: bu dosya burada tek başına düzenleniyor.',
      forkQuestion: 'Belgeniz öbürlerinden ayrı düştü ve dosyayı bu arada başkası kaydetti. Saklanan dosyayı yükleyin (kaydedilmemiş değişiklikleriniz gider; gerekiyorsa önce kopyanızı indirin) ya da kendinizinkini üzerine kaydedin (öbürlerinin kaydettiği gider).',
    },
  },
  sheetDeleted: (name) => `${name} sayfası silindi. Geri al ile geri gelir.`,
  referenceAdded: (reference) => `Başvuru ${reference}`,
  editSheetGone: 'Hücrenin sayfası artık yok; yazılan kaydedilmedi.',
  badSheetName: 'Sayfa adı 1-31 karakter olmalı; [ ] : * ? / \\ içeremez, kesme işaretiyle başlayıp bitemez ve "History" olamaz',
  sheetNameTaken: (name) => `"${name}" adında başka bir sayfa var`,
  cannotDeleteLastSheet: 'Görünen son sayfa silinemez',
  sheetKindFixed: (action) => `Bu tür sayfalar henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : action === 'copy' ? 'kopyalanamıyor' : action === 'color' ? 'renklendirilemiyor' : 'silinemiyor'}`,
  sheetBlockedBy: (action, what) =>
    what === 'table'
      ? `Tablo içeren sayfa henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : 'silinemiyor'}`
      : `${what === 'pivotTable' ? 'Pivot tablo' : what === 'chart' ? 'Grafik' : what === 'form controls' ? 'Form denetimi' : what} içeren çalışma kitabında sayfalar henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : 'silinemiyor'}`,
  sheetFormulaUnread: (action, where) => `Sayfa ${action === 'rename' ? 'yeniden adlandırılamadı' : 'silinemedi'}: ${where} formülü okunamıyor`,
  sheetNamesClash: (first, second) => `Bu çalışma kitabının sayfaları değiştirilemiyor: "${first}" ve "${second}" adları yalnız harf büyüklüğüyle ayrılıyor`,
  rules: {
    validationCommand: 'Veri doğrulama\u2026',
    validationTitle: 'Veri doğrulama',
    ok: 'Tamam',
    cancel: 'İptal',
    close: 'Kapat',
    clearAll: 'Tümünü temizle',
    settings: 'Ayarlar',
    inputMessage: 'Giriş iletisi',
    errorAlert: 'Hata uyarısı',
    allow: 'İzin verilen',
    types: { none: 'Herhangi bir değer', whole: 'Tam sayı', decimal: 'Ondalık', list: 'Liste', date: 'Tarih', time: 'Saat', textLength: 'Metin uzunluğu', custom: 'Özel formül' },
    data: 'Veri',
    operators: {
      between: 'arasında',
      notBetween: 'arasında değil',
      equal: 'eşit',
      notEqual: 'eşit değil',
      greaterThan: 'büyük',
      lessThan: 'küçük',
      greaterThanOrEqual: 'büyük ya da eşit',
      lessThanOrEqual: 'küçük ya da eşit',
    },
    minimum: 'En az',
    maximum: 'En çok',
    value: 'Değer',
    startDate: 'Başlangıç tarihi',
    endDate: 'Bitiş tarihi',
    date: 'Tarih',
    startTime: 'Başlangıç saati',
    endTime: 'Bitiş saati',
    time: 'Saat',
    source: 'Kaynak',
    formula: 'Formül',
    sourceHint: (separator) => `Öğeler "${separator}" ile ayrılır; ya da hücreler: =$A$1:$A$5`,
    ignoreBlank: 'Boşu yok say',
    inCellDropdown: 'Hücre içi açılır liste',
    showInputMessage: 'Hücre seçildiğinde giriş iletisini göster',
    showErrorAlert: 'Geçersiz veri girildiğinde hata uyarısı göster',
    style: 'Stil',
    styles: { stop: 'Dur', warning: 'Uyarı', information: 'Bilgi' },
    title: 'Başlık',
    message: 'İleti',
    mixedSelection: 'Seçim birden çok kural ya da kuralsız hücreler içeriyor. Tamam, hepsinin yerine bu kuralı koyar.',
    needNumber: (field) => `${field}: bir sayı ya da formül girin`,
    needWholeNumber: (field) => `${field}: bir tam sayı ya da formül girin`,
    needDate: (field) => `${field}: bir tarih ya da formül girin`,
    needTime: (field) => `${field}: bir saat ya da formül girin`,
    needValue: (field) => `${field} gerekli`,
    needItems: 'Listenin öğelerini ya da öğeleri tutan hücreleri girin',
    itemWithComma: 'Liste öğesi virgül içeremez',
    listTooLong: (limit) => `Liste öğeleri toplam ${limit} karakteri aşamaz; daha uzun liste için hücre kullanın`,
    textTooLong: (field, limit) => `${field} ${limit} karakterden uzun olamaz`,
    badRuleFormula: (field) => `${field}: formül okunamıyor`,
    extensionRule: 'Bu hücrelere olduğu gibi korunan bir kural uygulanıyor (listesi başka sayfada); burada değiştirilemez',
    tooManyValidations: 'Sayfada çok fazla veri doğrulama kuralı var',
    tooManyRanges: 'Kural çok fazla ayrı aralığa bölünürdü',
    formulaBeingTyped: 'Önce formülü bitirin',
    dialogOpen: 'Önce açık kural kutusunu onaylayın ya da kapatın',
    pointOwnSheet: 'Kural kendi sayfasının hücrelerine uygulanır',
    validationSet: 'Veri doğrulama uygulandı',
    validationCleared: 'Veri doğrulama temizlendi',
    conditionalCommand: 'Koşullu biçimlendirme',
    quickRules: {
      greaterThan: 'Büyüktür\u2026',
      lessThan: 'Küçüktür\u2026',
      between: 'Arasında\u2026',
      equal: 'Eşittir\u2026',
      containsText: 'Metin içerir\u2026',
      duplicateValues: 'Yinelenen değerler\u2026',
      top10: 'İlk 10 öğe\u2026',
      colorScale: 'Renk ölçekleri\u2026',
      dataBar: 'Veri çubukları\u2026',
    },
    clearFromSelection: 'Seçili hücrelerden kuralları temizle',
    clearFromSheet: 'Bu sayfadan kuralları temizle',
    manageRules: 'Kuralları yönet\u2026',
    conditionalTitle: 'Koşullu biçimlendirme kuralı',
    managerTitle: 'Koşullu biçimlendirme kuralları',
    ruleType: 'Kural türü',
    conditionalTypes: {
      cellIs: 'Hücre değeri',
      containsText: 'Metin içerir',
      notContainsText: 'Metin içermez',
      beginsWith: 'Metin şununla başlar',
      endsWith: 'Metin şununla biter',
      duplicateValues: 'Yinelenen değerler',
      uniqueValues: 'Benzersiz değerler',
      top10: 'İlk ya da son değerler',
      aboveAverage: 'Ortalamanın üstü ya da altı',
      containsBlanks: 'Boş hücreler',
      notContainsBlanks: 'Boş olmayan hücreler',
      containsErrors: 'Hata içeren hücreler',
      notContainsErrors: 'Hata içermeyen hücreler',
      expression: 'Formül',
      colorScale: 'Renk ölçeği',
      dataBar: 'Veri çubuğu',
      iconSet: 'Simge kümesi',
      timePeriod: 'Tarih dönemi',
    },
    text: 'Metin',
    rank: 'Sayı',
    percentOfRange: 'Aralığın yüzdesi',
    lowestValues: 'En düşük değerler',
    averageSide: 'Değerler',
    averages: { above: 'ortalamanın üstünde', below: 'ortalamanın altında' },
    includeAverage: 'Ortalamanın kendisi de',
    stopIfTrue: 'Doğruysa durdur',
    format: 'Biçim',
    formatWith: 'Biçim',
    presets: {
      lightRed: 'Açık kırmızı dolgu, koyu kırmızı metin',
      yellow: 'Sarı dolgu, koyu sarı metin',
      green: 'Yeşil dolgu, koyu yeşil metin',
      keep: 'Var olan biçim',
      custom: 'Özel biçim',
      none: 'Biçim yok',
    },
    textColor: 'Yazı rengi',
    fillColor: 'Dolgu rengi',
    bold: 'Kalın',
    italic: 'İtalik',
    underline: 'Altı çizili',
    strike: 'Üstü çizili',
    scaleKind: 'Renkler',
    scales: { '2': 'İki renk', '3': 'Üç renk' },
    lowest: 'En düşük',
    midpoint: 'Orta nokta',
    highest: 'En yüksek',
    barColor: 'Çubuk rengi',
    iconStyle: 'Simge stili',
    iconSets: {
      '3Arrows': '3 ok (renkli)',
      '3ArrowsGray': '3 ok (gri)',
      '3Flags': '3 bayrak',
      '3TrafficLights1': '3 trafik ışığı',
      '3TrafficLights2': '3 trafik ışığı (çerçeveli)',
      '3Signs': '3 işaret',
      '3Symbols': '3 simge (daire içinde)',
      '3Symbols2': '3 simge',
      '3Stars': '3 yıldız',
      '3Triangles': '3 üçgen',
      '4Arrows': '4 ok (renkli)',
      '4ArrowsGray': '4 ok (gri)',
      '4RedToBlack': 'Kırmızıdan siyaha',
      '4Rating': '4 derecelendirme',
      '4TrafficLights': '4 trafik ışığı',
      '5Arrows': '5 ok (renkli)',
      '5ArrowsGray': '5 ok (gri)',
      '5Rating': '5 derecelendirme',
      '5Quarters': '5 çeyrek',
      '5Boxes': '5 kutu',
    },
    iconPreview: 'Kümenin simgeleri, en düşük değerlerden en yükseğe',
    reverseIcons: 'Simge sırasını ters çevir',
    iconOnly: 'Yalnız simgeyi göster',
    period: 'Dönem',
    showRulesFor: 'Kuralları göster',
    scopeSelection: 'Geçerli seçim',
    scopeSheet: 'Bu sayfa',
    appliesTo: 'Uygulandığı yer',
    moveUp: 'Yukarı taşı',
    moveDown: 'Aşağı taşı',
    edit: 'Düzenle',
    deleteRule: 'Sil',
    newRule: 'Yeni kural\u2026',
    noRules: 'Kural yok',
    rulesShown: (first, last, total) => `${total} kuralın ${first}\u2013${last}. olanları`,
    nextRules: (count) => `Sonraki ${count}`,
    previousRules: (count) => `Önceki ${count}`,
    ruleRow: (place, description) => `${place}. kural: ${description}`,
    cellValue: 'Hücre değeri',
    and: 've',
    textRules: { containsText: 'Metin içerir', notContainsText: 'Metin içermez', beginsWith: 'Metin şununla başlar', endsWith: 'Metin şununla biter' },
    topItems: (count, percent, bottom) => `${bottom ? 'Son' : 'İlk'} ${percent ? '%' : ''}${count}`,
    average: (above, equal) => `Ortalamanın ${above ? 'üstü' : 'altı'}${equal ? ' ya da ortalama' : ''}`,
    formulaIs: 'Formül',
    datesOccurring: 'Tarih',
    periods: {
      today: 'bugün',
      yesterday: 'dün',
      tomorrow: 'yarın',
      last7Days: 'son 7 gün',
      thisWeek: 'bu hafta',
      lastWeek: 'geçen hafta',
      nextWeek: 'gelecek hafta',
      thisMonth: 'bu ay',
      lastMonth: 'geçen ay',
      nextMonth: 'gelecek ay',
    },
    needRange: 'Kuralın uygulanacağı hücreleri girin',
    badRange: 'Bunlar sayfanın aralıkları değil; A1:B10 gibi yazın',
    needWholeUpTo: (field, limit) => `${field}: 1 ile ${limit} arasında bir tam sayı girin`,
    extensionRuleKept: 'Bu kuralın yeni Excel sürümüne özgü ayarları var ve olduğu gibi korunuyor; burada değiştirilemez ya da silinemez',
    extensionRuleSays: 'Bu kuralın içeriği yeni Excel sürümüne özgü ayarlar taşıyor ve olduğu gibi korunuyor; uygulandığı yer ve sırası değiştirilebilir, kural silinebilir',
    orderFixed: 'Bu sayfadaki kuralların sırası değiştirilemiyor: biri yeriyle birlikte olduğu gibi korunuyor',
    ruleKindKept: 'Bu tür kural olduğu gibi korunur; taşınabilir ve silinebilir, değiştirilemez',
    tooManyConditional: 'Sayfada çok fazla koşullu biçim var',
    ruleAdded: 'Koşullu biçim eklendi',
    ruleChanged: 'Koşullu biçim değiştirildi',
    ruleDeleted: 'Koşullu biçim silindi',
    rulesCleared: 'Koşullu biçimler temizlendi',
  },
  charts: {
    insert: 'Grafik ekle',
    types: { column: 'Sütun', stackedColumn: 'Yığılmış sütun', bar: 'Çubuk', line: 'Çizgi', area: 'Alan', pie: 'Pasta', doughnut: 'Halka', scatter: 'Dağılım' },
    edit: 'Grafiği düzenle\u2026',
    remove: 'Grafiği sil',
    dialogTitle: 'Grafik',
    type: 'Tür',
    otherType: 'Başka bir tür',
    grouping: 'Seriler',
    groupings: { standard: 'Yan yana', stacked: 'Yığılmış', percentStacked: '%100 yığılmış' },
    title: 'Başlık',
    range: 'Veri aralığı',
    seriesIn: 'Seriler',
    directions: { auto: 'Otomatik', columns: 'Sütunlarda', rows: 'Satırlarda' },
    legend: 'Gösterge',
    legends: { r: 'Sağda', l: 'Solda', t: 'Üstte', b: 'Altta', tr: 'Sağ üstte', none: 'Yok' },
    dataLabels: 'Değerleri etiket olarak göster',
    axes: 'Eksenler',
    categoryTitle: 'Kategori ekseninin başlığı',
    valueTitle: 'Değer ekseninin başlığı',
    minimum: 'En küçük değer',
    maximum: 'En büyük değer',
    series: 'Seriler',
    categories: 'Kategori etiketleri',
    seriesName: (n) => `${n}. serinin adı`,
    seriesValues: (n) => `${n}. serinin değerleri`,
    seriesColor: (n) => `${n}. serinin rengi`,
    ownColor: (n) => `${n}. serinin kendi rengi var`,
    moveUp: (n) => `${n}. seriyi yukarı taşı`,
    moveDown: (n) => `${n}. seriyi aşağı taşı`,
    removeSeries: (n) => `${n}. seriyi kaldır`,
    addSeries: 'Seri ekle',
    ok: 'Tamam',
    cancel: 'Vazgeç',
    selected: (name) => (name === '' ? 'Grafik seçildi' : `Grafik seçildi: ${name}`),
    added: 'Grafik eklendi',
    changed: 'Grafik değiştirildi',
    moved: 'Grafik taşındı',
    removed: 'Grafik silindi',
    needRange: 'Grafiğin göstereceği hücre aralığını girin',
    badRange: 'Bu bir hücre aralığı değil',
    badNumber: 'Bir sayı girin ya da alanı boş bırakın',
    needValues: 'Bir serinin değer aralığı olmalı',
    typeFixed: 'Bu grafiğin türü değiştirilemez: dosyası başka bir türün koruyamayacağı biçimlendirme taşıyor',
    notEditable: 'Bu grafik taşınabilir ve silinebilir, değiştirilemez: dosyası korunamayacak içerik taşıyor',
    noChart: 'Seçili grafik yok',
    seriesFixed: 'Bu grafiğin serileri eklenemez, kaldırılamaz ya da yeniden sıralanamaz: türü ya da dosyası buna izin vermiyor',
    groupingFixed: 'Bu grafik türünün serileri yığılamaz',
    noAxis: 'Bu grafikte böyle bir eksen yok',
    tooManySeries: (most) => `Bir grafikte en çok ${most} seri olabilir`,
    tooManyCharts: 'Buraya daha fazla grafik eklenemez',
    textTooLong: (most) => `Bu metin ${most} karakterden uzun olamaz`,
    badCharacter: 'Metin, dosyaya yazılamayan bir karakter içeriyor',
    refUnread: 'Sayfa değiştirilemiyor: bir grafik onu okunamayan bir başvuruda anıyor',
    placesFixed: 'Bu sayfada grafik eklenemez, taşınamaz ve silinemez: dosyadaki çizimi değiştirilemiyor',
    wrapped: 'Bu grafik silinemez: dosyada başka bir nesneye bağlı',
    duplicate: 'Çoğalt',
    copied: 'Grafik kopyalandı',
    cutOut: 'Grafik kesildi',
    pasted: 'Grafik yapıştırıldı',
    duplicated: 'Grafik çoğaltıldı',
    notCopyable: 'Bu grafik kopyalanamaz: dosyasındaki içerik kopyada korunamıyor',
    copyOutdated: 'Grafik yapıştırılamıyor: kopyalandıktan sonra satırlar, sütunlar ya da sayfalar değişti. Yeniden kopyalayın.',
    copyLeftOut: (what) => `Kopya şunlar olmadan yapıldı: ${what.map((name) => (name === 'userShapes' ? 'grafiğin üzerine çizilmiş şekiller' : name === 'externalData' ? 'kitabın dışında tutulan veri' : name)).join(', ')}`,
    copyUnbound: (count) => `Kopyanın ${count} başvurusu hâlâ sayfa adı taşımıyor`,
    noCellPaste: 'Bir grafik seçili: hücre yapıştırmak için bir hücre seçin',
    copyGone: 'Panodaki grafik başka bir yerde kopyalanmış, buraya yapıştırılamıyor: grafiği yeniden kopyalayın',
  },
  pictures: {
    insert: 'Görsel ekle',
    added: 'Görsel eklendi',
    moved: 'Görsel taşındı',
    removed: 'Görsel silindi',
    selected: (name) => (name ? `Görsel seçildi: ${name}` : 'Görsel seçildi'),
    remove: 'Görseli sil',
    notSupported: 'Bu görsel eklenemiyor: yalnız PNG, JPEG, GIF ve BMP görseller eklenebilir',
    tooLarge: 'Bu görsel eklenemeyecek kadar büyük',
    tooMany: 'Bu çalışma kitabına daha fazla görsel eklenemez',
    protectedSheet: 'Sayfa korumalı: görselleri değiştirilemez',
    placesFixed: 'Bu sayfada görsel eklenemez, taşınamaz ve silinemez: dosyadaki çizimi değiştirilemiyor',
    wrapped: 'Bu görsel silinemez: dosyada yerini başka bir nesneyle paylaşıyor',
    noCopy: 'Görseller henüz kopyalanamıyor',
    noCellPaste: 'Bir görsel seçili: hücre yapıştırmak için bir hücre seçin',
    dropRefused: 'Görsel şu anda buraya bırakılamaz',
  },
  pivots: {
    refresh: 'Pivot tabloyu yenile',
    refreshAll: 'Tüm pivot tabloları yenile',
    refreshed: (count) => (count === 1 ? 'Pivot tablo yenilendi' : `${count} pivot tablo yenilendi`),
    refused: pivotRefusalTr,
    none: 'Çalışma kitabında yenilenecek pivot tablo yok',
    insert: 'Pivot tablo ekle…',
    createTitle: 'Pivot tablo ekle',
    source: 'Veri',
    place: 'Yeri',
    newSheet: 'Yeni sayfa',
    thisSheet: 'Bu sayfada, bir hücrede',
    cell: 'Hücre',
    ok: 'Tamam',
    cancel: 'İptal',
    badSource: 'A1:D20 gibi bir aralık ya da bir aralığın adını yazın',
    badCell: 'F2 gibi bir hücre yazın',
    made: 'Pivot tablo eklendi',
    fields: 'Pivot tablo alanları…',
    panelTitle: 'Pivot tablo alanları',
    close: 'Kapat',
    fieldList: 'Alanlar',
    boxes: { filters: 'Süzgeçler', cols: 'Sütunlar', rows: 'Satırlar', values: 'Değerler' },
    dropHere: 'Alanları buraya sürükleyin',
    moveUp: (name) => `${name} yukarı`,
    moveDown: (name) => `${name} aşağı`,
    moveTo: (name) => `${name} nerede dursun`,
    remove: (name) => `${name} çıkar`,
    summarize: (name) => `${name} özetleme işlevi`,
    numberFormat: (name) => `${name} sayı biçimi`,
    generalFormat: 'Genel',
    aggregates: { sum: 'Toplam', count: 'Say', average: 'Ortalama', max: 'En büyük', min: 'En küçük', product: 'Çarpım', countNums: 'Sayıları say', stdDev: 'Std sapma', stdDevp: 'Std sapma (popülasyon)', var: 'Varyans', varp: 'Varyans (popülasyon)' },
    subtotal: 'Ara toplamlar',
    subtotals: (name) => `${name} ara toplamları`,
    sort: (name) => `${name} sırası`,
    sorts: { manual: 'Verideki sırayla', ascending: 'A’dan Z’ye', descending: 'Z’den A’ya' },
    shownItem: (name) => `${name} için gösterilen öğe`,
    allItems: '(Tümü)',
    filterItems: (name) => `${name} öğeleri`,
    moreItems: (count) => `${count} öğe daha listelenmedi`,
    rowGrandTotals: 'Satırların genel toplamları',
    colGrandTotals: 'Sütunların genel toplamları',
    changed: 'Pivot tablo değişti',
    noTable: 'Bir pivot tablo hücresi seçin',
  },
  find: {
    title: 'Bul ve değiştir',
    findLabel: 'Bul',
    replaceLabel: 'Yeni metin',
    previous: 'Önceki (Shift+Enter)',
    next: 'Sonraki (Enter)',
    replace: 'Değiştir',
    replaceAll: 'Tümünü değiştir',
    toggleReplace: 'Değiştirmeyi göster',
    close: 'Kapat (Esc)',
    matchCase: 'Büyük/küçük harf duyarlı',
    wholeWord: 'Yalnız tam sözcükler',
    entireCell: 'Hücrenin tamamı',
    mainOnly: 'Yalnız ana metin',
    lookIn: 'Aranan yer',
    lookInFormulas: 'Formüller',
    lookInValues: 'Değerler',
    within: 'Kapsam',
    withinSheet: 'Sayfa',
    withinWorkbook: 'Çalışma kitabı',
    searching: 'Aranıyor…',
    noMatches: 'Sonuç yok',
    matches: (count, more) => (more ? `${count} sonuçtan fazla` : `${count} sonuç`),
    replaced: (count, skipped) => `${count} yer değiştirildi` + (skipped > 0 ? `; ${skipped} yer değiştirilemedi` : ''),
    inTextBoxes: (count, why) =>
      ` (${count} tanesi metin kutusunda: ${why === 'shared' ? 'dosya birlikte düzenlenirken değiştirilemez' : why === 'tracked' ? 'orada değişiklik izlenemez' : 'dosya kutuyu kayıpsız değiştirilemeyecek bir biçimde tutuyor'})`,
    inComments: (count) => ` (${count} tanesi yorumlarda: dosya birlikte düzenlenirken değiştirilemez)`,
    outsideComments: (count) => ` (${count} tanesi yorumların dışında: bu dosyada yalnız yorumlar değiştirilebilir)`,
    replacedInSteps: (steps) => `Dosya birlikte düzenlendiği için ${steps} adımda değiştirildi; her adım ayrı geri alınır`,
    changedMeanwhile: 'Arama sürerken dosya değişti: yeniden deneyin',
    tooMany: (limit) => `Bir seferde değiştirilecek ${limit} sonuçtan fazlası var: aramayı daraltın`,
    readOnly: 'Hiçbir şey değiştirilemez: dosya salt okunur',
  },
  pageSetup: {
    title: 'Sayfa yapısı',
    ok: 'Tamam',
    cancel: 'İptal',
    actions: 'Sonlar',
    unit: 'cm',
    decimal: ',',
    page: 'Sayfa',
    paper: 'Kâğıt boyutu',
    otherPaper: 'Diğer (olduğu gibi)',
    orientation: 'Yön',
    portrait: 'Dikey',
    landscape: 'Yatay',
    scaling: 'Ölçek',
    adjustTo: 'Ayarla',
    fitTo: 'Sayfaya sığdır',
    scale: '% normal boyut',
    fitWide: 'Sayfa genişliği',
    fitTall: 'Sayfa yüksekliği (0: gerektiği kadar)',
    margins: 'Kenar boşlukları',
    top: 'Üst',
    bottom: 'Alt',
    left: 'Sol',
    right: 'Sağ',
    header: 'Üstbilgi',
    footer: 'Altbilgi',
    centerOnPage: 'Sayfada ortala',
    horizontally: 'Yatay',
    vertically: 'Dikey',
    sheet: 'Sayfa ayarları',
    printArea: 'Yazdırma alanı',
    titleRows: 'Üstte yinelenecek satırlar',
    titleCols: 'Solda yinelenecek sütunlar',
    gridLines: 'Kılavuz çizgileri',
    headings: 'Satır ve sütun başlıkları',
    insertBreak: 'Sayfa sonu ekle',
    removeBreak: 'Sayfa sonunu kaldır',
    resetBreaks: 'Tüm sayfa sonlarını sıfırla',
    applyTo: 'Uygulama yeri',
    thisSection: 'Bu bölüm',
    wholeDocument: 'Tüm belge',
    sectionBreak: (kind) => ({ nextPage: 'Bölüm sonu: sonraki sayfa', continuous: 'Bölüm sonu: sürekli', evenPage: 'Bölüm sonu: çift sayfa', oddPage: 'Bölüm sonu: tek sayfa' })[kind],
    badLength: '0 ya da daha büyük bir uzunluk yazın',
    badScale: '10 ile 400 arasında tam bir yüzde yazın',
    badFit: '0 ile 32767 arasında tam bir sayfa sayısı yazın',
    badArea: 'Yazdırma alanını A1:D20 gibi başvurularla yazın',
    badRows: 'Yinelenecek satırları $1:$2 gibi yazın',
    badCols: 'Yinelenecek sütunları $A:$B gibi yazın',
    noRoom: 'Kenar boşlukları kâğıtta yer bırakmıyor',
    shared: 'Dosya birlikte düzenlenirken sayfa yapısı değiştirilemez',
    tracked: 'Değişiklikler izlenirken sayfa yapısı değiştirilemez: önce izlemeyi kapatın',
  },
  preview: {
    title: 'Baskı önizleme',
    pages: 'Sayfalar',
    page: (number, total) => `Sayfa ${number} / ${total}`,
    counter: (number, total) => `${number} / ${total}`,
    previous: 'Önceki sayfa',
    next: 'Sonraki sayfa',
    zoomIn: 'Yakınlaştır',
    zoomOut: 'Uzaklaştır',
    fit: 'Sayfaya sığdır',
    actualSize: 'Gerçek boyut (%100)',
    zoom: (percent) => `Yakınlaştırma %${percent}`,
    percent: (value) => `%${value}`,
    print: 'Yazdır',
    pdf: 'PDF olarak indir',
    pageSetup: 'Sayfa yapısı…',
    close: 'Baskı önizlemeyi kapat',
    which: 'Yazdırılacak',
    active: 'Etkin sayfa',
    all: 'Tüm çalışma kitabı',
    planning: 'Sayfalar hazırlanıyor…',
  },
};

/** Why a pivot table cannot be refreshed, in Turkish: by the reason the engine gives (see src/ops/pivots.ts, src/formats/xlsx/pivot.ts). */
const PIVOT_REASONS_TR: [RegExp, (m: RegExpExecArray) => string][] = [
  [/^its data does not come from a range of the workbook/, () => 'verisi çalışma kitabındaki bir aralıktan gelmiyor (dış veri, OLAP, birleştirme ya da senaryo)'],
  [/^its source is not stated$/, () => 'kaynağı belirtilmemiş'],
  [/^its data comes from another workbook$/, () => 'verisi başka bir çalışma kitabından geliyor'],
  [/^its data is a consolidation of several ranges$/, () => 'verisi birden çok aralığın birleştirilmesi'],
  [/^its data comes from an OLAP source$/, () => 'verisi bir OLAP kaynağından geliyor'],
  [/^it has a calculated field$/, () => 'hesaplanmış alanı var'],
  [/^it groups items/, () => 'öğeleri gruplanmış (tarih ya da sayı aralıkları, elle birleştirilmiş öğeler)'],
  [/^its fields cannot be read$/, () => 'alanları okunamıyor'],
  [/^the workbook does not list its cache$/, () => 'çalışma kitabı önbelleğini listelemiyor'],
  [/^its values stand in rows$/, () => 'değerleri satırlarda duruyor'],
  [/^it shows no headers$/, () => 'başlıkları gösterilmiyor'],
  [/^it has the classic layout$/, () => 'klasik düzende'],
  [/^it merges the cells of its labels$/, () => 'etiket hücrelerini birleştiriyor'],
  [/^a field is sorted in a way that is not known here$/, () => 'bir alanı burada bilinmeyen bir biçimde sıralanmış'],
  [/^a field has subtotals by functions of their own$/, () => 'bir alanın kendi işlevleriyle ara toplamları var'],
  [/^a field puts blank rows after its items$/, () => 'bir alan öğelerinden sonra boş satır bırakıyor'],
  [/^a field is sorted by values$/, () => 'bir alan değerlere göre sıralanmış'],
  [/^a field repeats the labels of its items$/, () => 'bir alan öğe etiketlerini yineliyor'],
  [/^it has calculated items$/, () => 'hesaplanmış öğeleri var'],
  [/^an item of a field is missing from its source$/, () => 'bir alanın bir öğesi kaynağında yok'],
  [/^a value is summarized by a function that is not known here \("(.*)"\)$/, (m) => `bir değer burada bilinmeyen bir işlevle özetleniyor ("${m[1]}")`],
  [/^a value is shown as a calculation/, () => 'bir değer hesaplama olarak gösteriliyor (oran, fark, değişen toplam …)'],
  [/^it formats some of its items$/, () => 'bazı öğelerine biçim verilmiş'],
  [/^a pivot chart is made from it$/, () => 'ondan bir pivot grafik yapılmış'],
  [/^it filters by labels or by values$/, () => 'etikete ya da değere göre süzülüyor'],
  [/^its fields are not those of its cache$/, () => 'alanları önbelleğinin alanlarıyla uyuşmuyor'],
  [/^it names a cache the workbook does not have under that number$/, () => 'çalışma kitabında o numarayla bulunmayan bir önbelleği anıyor'],
  [/^two sheets show the same table$/, () => 'iki sayfa aynı tabloyu gösteriyor'],
  [/^its source "(.*)" is a table or a name that is not known here$/, (m) => `kaynağı "${m[1]}" burada bilinmeyen bir tablo ya da ad`],
  [/^its source "(.*)" is not one range of cells$/, (m) => `kaynağı "${m[1]}" tek bir hücre aralığı değil`],
  [/^the sheet of its source \("(.*)"\) is not in the workbook$/, (m) => `kaynağının sayfası ("${m[1]}") çalışma kitabında yok`],
  [/^its source \("(.*)"\) is not a range of cells$/, (m) => `kaynağı ("${m[1]}") bir hücre aralığı değil`],
  [/^its source has (\d+) columns and its cache (\d+) fields$/, (m) => `kaynağında ${m[1]} sütun, önbelleğinde ${m[2]} alan var`],
  [/^its source has (\d+) columns and its cache (\d+) fields: columns were inserted into it, and a refresh here adds no fields$/, (m) => `kaynağında ${m[1]} sütun, önbelleğinde ${m[2]} alan var: kaynağa sütun eklenmiş, buradaki yenileme önbelleğe yeni alan eklemez`],
  // Why rows and columns cannot move in a workbook with pivot tables (see src/ops/pivot-places.ts).
  [/^the workbook has more pivot tables than are read$/, () => 'çalışma kitabında okunabilenden çok pivot tablo var'],
  [/^a sheet names a pivot table that is not in the package$/, () => 'bir sayfa pakette olmayan bir pivot tabloyu anıyor'],
  [/^a pivot table has no cache$/, () => 'bir pivot tablonun önbelleği yok'],
  [/^the cache of a pivot table cannot be read$/, () => 'bir pivot tablonun önbelleği okunamıyor'],
  [/^two sheets show the same pivot table$/, () => 'iki sayfa aynı pivot tabloyu gösteriyor'],
  [/^a pivot table cannot be read$/, () => 'bir pivot tablo okunamıyor'],
  [/^the package has a pivot table no sheet shows$/, () => 'pakette hiçbir sayfanın göstermediği bir pivot tablo var'],
  [/^the package has a pivot cache no table is made from$/, () => 'pakette hiçbir tablonun kullanmadığı bir pivot önbelleği var'],
  [/^its pivot tables cannot be read$/, () => 'pivot tabloları okunamıyor'],
  [/^its source is not a range or a name of the workbook$/, () => 'bir önbelleğin kaynağı çalışma kitabında bir aralık ya da ad değil'],
  [/^its data is a consolidation of several ranges$/, () => 'bir önbelleğin verisi birden çok aralığın birleştirmesi'],
  [/^its source states both a name and a range$/, () => 'bir önbelleğin kaynağı hem ad hem aralık belirtiyor'],
  [/^its source \("(.*)"\) is not one range of cells$/, (m) => `bir önbelleğin kaynağı ("${m[1]}") tek bir hücre aralığı değil`],
  [/^its source has no rows below its header$/, () => 'kaynağında başlığın altında satır yok'],
  [/^its source has more than (\d+) cells$/, (m) => `kaynağında ${m[1]} hücreden fazlası var`],
  [/^the header of column (\S+) of its source is empty$/, (m) => `kaynağındaki ${m[1]} başlığı boş`],
  [/^the header of (\S+) of its source is "(.*)", the field of the cache is "(.*)"$/, (m) => `kaynağındaki ${m[1]} başlığı "${m[2]}", önbellekteki alan "${m[3]}"`],
  [/^it has no values$/, () => 'değer alanı yok'],
  [/^its values have no place in its columns$/, () => 'değerlerinin sütunlarda yeri yok'],
  [/^its values do not stand after the fields of its columns$/, () => 'değerleri sütun alanlarından sonra durmuyor'],
  [/^its values stand in its columns although there is one of them$/, () => 'tek değeri olduğu hâlde değerler sütunlarda duruyor'],
  [/^it has several values as well as fields in its columns$/, () => 'hem birden çok değeri hem sütun alanları var'],
  [/^it has no fields in its rows$/, () => 'satırlarında alan yok'],
  [/^it names a field it does not have$/, () => 'olmayan bir alanı anıyor'],
  [/^its fields in rows are laid out in different forms$/, () => 'satır alanları farklı düzenlerde'],
  [/^it has fields in its columns and is not in compact form$/, () => 'sütun alanları var ve sıkıştırılmış düzende değil'],
  [/^the field "(.*)" lists no items$/, (m) => `"${m[1]}" alanı öğe listelemiyor`],
  [/^the items of a field do not match its cache$/, () => 'bir alanın öğeleri önbelleğiyle uyuşmuyor'],
  [/^the item the page shows for the field "(.*)" is not in its source any more$/, (m) => `sayfa alanı "${m[1]}" için gösterilen öğe artık kaynakta yok`],
  [/^it would have more than (\d+) cells$/, (m) => `${m[1]} hücreden büyük olurdu`],
  [/^it has more than (\d+) cells$/, (m) => `${m[1]} hücreden büyük`],
  [/^it would reach past the end of the sheet$/, () => 'sayfanın sonunu aşardı'],
  [/^it would grow over the pivot table "(.*)"$/, (m) => `"${m[1]}" pivot tablosunun üzerine taşardı`],
  [/^it would grow over the merged cells (\S+)$/, (m) => `${m[1]} birleştirilmiş hücrelerinin üzerine taşardı`],
  [/^it would grow over (\S+), which is not empty$/, (m) => `boş olmayan ${m[1]} hücresinin üzerine taşardı`],
  [/^its cache is not known$/, () => 'önbelleği bilinmiyor'],
  [/^every item of the field "(.*)" would be hidden$/, (m) => `"${m[1]}" alanının bütün öğeleri gizli kalırdı`],
  [/^it cannot tell what refers to its cache$/, () => 'önbelleğini neyin andığı anlaşılamıyor (çalışma kitabı taranamayacak kadar büyük ya da okunamayan bir parçası var)'],
  [/^its source has formulas, and another pivot table refreshed with it changed cells: refresh it on its own$/, () => 'kaynağında formül var ve onunla birlikte yenilenen başka bir pivot tablo hücreleri değiştirdi: bu tabloyu ayrıca yenileyin'],
  [/^another part of the workbook refers to its cache \((.*)\)$/, (m) => `çalışma kitabının başka bir parçası önbelleğini anıyor (${m[1]})`],
  [/^another part of the workbook refers to its records \((.*)\)$/, (m) => `çalışma kitabının başka bir parçası önbellek kayıtlarını anıyor (${m[1]})`],
  [/^a slicer or a timeline reads the caches of the workbook$/, () => 'bir dilimleyici ya da zaman çizelgesi çalışma kitabının önbelleklerini okuyor'],
  [/^a pivot chart is made from a table of its cache$/, () => 'önbelleğindeki bir tablodan pivot grafik yapılmış'],
  [/^a pivot chart is made from a table that is not known here$/, () => 'burada bilinmeyen bir tablodan pivot grafik yapılmış'],
  [/^the charts of the workbook are too large to look through for pivot charts$/, () => 'çalışma kitabının grafikleri pivot grafik aramak için fazla büyük'],
  [/^a chart of the workbook cannot be read to tell whether it is a pivot chart$/, () => 'bir grafiğin pivot grafik olup olmadığı okunamıyor'],
  [/^the workbook has too many parts to look through$/, () => 'çalışma kitabında taranamayacak kadar çok parça var'],
  [/^relationships of the workbook cannot be read$/, () => 'çalışma kitabının ilişkileri okunamıyor'],
  [/^it has a setting that is not known here \((.*)\)$/, (m) => `burada bilinmeyen bir ayarı var (${m[1]})`],
  [/^it has a setting that is not applied here \((.*)\)$/, (m) => `burada uygulanmayan bir ayarı var (${m[1]})`],
  [/^it has a part that is not known here \((.*)\)$/, (m) => `burada bilinmeyen bir öğesi var (${m[1]})`],
  [/^a field has a part that is not known here \((.*)\)$/, (m) => `bir alanının burada bilinmeyen bir öğesi var (${m[1]})`],
  [/^a value has a part that is not known here \((.*)\)$/, (m) => `bir değerinin burada bilinmeyen bir öğesi var (${m[1]})`],
  [/^it has an extension that is not known here \((.*)\)$/, (m) => `burada bilinmeyen bir uzantısı var (${m[1]})`],
  [/^a field has a caption of its own for its subtotals/, () => 'bir alanın ara toplamları için kendi başlığı var (subtotalCaption)'],
  [/^an item of a field is collapsed/, () => 'bir alanın bir öğesi daraltılmış (ayrıntıları gizli)'],
  [/^a field has an item of a kind that is not built here \("(.*)"\)$/, (m) => `bir alanda burada kurulmayan türde bir öğe var ("${m[1]}")`],
  [/^it has more than (\d+) fields in its rows and columns$/, (m) => `satır ve sütunlarında ${m[1]} alandan fazlası var`],
  [/^it sorts a field whose items are names of months or days/, () => 'öğeleri ay ya da gün adları olan bir alanı sıralıyor (Excel bunları takvim sırasıyla sıralar)'],
  [/^it would take too long to add up/, () => 'toplanması çok uzun sürer (alanları ve değerleri için kayıt sayısı çok fazla)'],
  [/^the labels of the field "(.*)" have a number format the style sheet has no style with$/, (m) => `"${m[1]}" alanının etiketlerinin sayı biçimi stil sayfasındaki hiçbir stilde yok`],
  [/^its sheet is not in the workbook$/, () => 'sayfası çalışma kitabında yok'],
  [/^it does not stand on a worksheet$/, () => 'bir çalışma sayfasında durmuyor'],
  // Why the fields of a table are not changed, or a table is not made (see src/ops/pivot-fields.ts, pivot-create.ts).
  [/^the fields are lists$/, () => 'alanlar liste olmalı'],
  [/^it has too many fields$/, () => 'çok fazla alanı var'],
  [/^there is no field (.*)$/, (m) => `${m[1]} numaralı alan yok`],
  [/^a field stands in two places \(rows, columns, filters\)$/, () => 'bir alan iki yerde duruyor (satırlar, sütunlar, süzgeçler)'],
  [/^it would have more than (\d+) fields in its rows and columns$/, (m) => `satır ve sütunlarında ${m[1]} alandan fazlası olurdu`],
  [/^it needs a field in its rows here$/, () => 'burada satırlarında en az bir alan olmalı'],
  [/^it needs a value here$/, () => 'burada en az bir değeri olmalı'],
  [/^it would have more than (\d+) values$/, (m) => `${m[1]} değerden fazlası olurdu`],
  [/^several values and fields in the columns together are not built here$/, () => 'birden çok değer ile sütun alanları burada birlikte kurulmuyor'],
  [/^a value names a field the (?:table|source) does not have$/, () => 'bir değer olmayan bir alanı anıyor'],
  [/^a value has no name$/, () => 'bir değerin adı yok'],
  [/^the name of a value is longer than (\d+) characters$/, (m) => `bir değerin adı ${m[1]} karakterden uzun`],
  [/^a value cannot have the name of a field \("(.*)"\)$/, (m) => `bir değer bir alanın adını alamaz ("${m[1]}")`],
  [/^two values have the name "(.*)"$/, (m) => `iki değerin adı aynı: "${m[1]}"`],
  [/^a number format is not one$/, () => 'sayı biçimi geçersiz'],
  [/^hidden items name no field$/, () => 'gizlenen öğeler bir alan anmıyor'],
  [/^a subtotal names no field$/, () => 'ara toplam bir alan anmıyor'],
  [/^a sort names no field or no order$/, () => 'sıralama bir alan ya da sıra anmıyor'],
  [/^the field "(.*)" lists no items yet: put it into the rows, columns or filters first$/, (m) => `"${m[1]}" alanının henüz öğe listesi yok: önce satırlara, sütunlara ya da süzgeçlere koyun`],
  [/^the field "(.*)" has no item (\S+)(?: to show)?$/, (m) => `"${m[1]}" alanında ${m[2]} numaralı öğe yok`],
  [/^the field "(.*)" has (\d+) different values: more than (\d+) cannot be laid out$/, (m) => `"${m[1]}" alanında ${m[2]} farklı değer var: ${m[3]} taneden fazlası yerleştirilemez`],
  [/^its fields of the page stand in several columns$/, () => 'süzgeç alanları birden çok sütunda duruyor'],
  [/^there is no room above it for its fields of the page$/, () => 'üstünde süzgeç alanlarına yer yok'],
  [/^it has formats of parts of the table, which name its fields$/, () => 'tablonun parçalarına verilmiş biçimler var ve bunlar alanlarını anıyor'],
  [/^its fields of the page are laid out in a way that is not built here$/, () => 'süzgeç alanları burada kurulmayan bir düzende'],
  [/^it has an extension that names its fields$/, () => 'alanlarını anan bir uzantısı var'],
  [/^nothing is given$/, () => 'hiçbir şey verilmedi'],
  [/^its source is not given$/, () => 'kaynağı verilmedi'],
  [/^there is no sheet (.*)$/, (m) => `${m[1]} numaralı sayfa yok`],
  [/^its source is not a range of cells$/, () => 'kaynağı bir hücre aralığı değil'],
  [/^its source is not on a worksheet$/, () => 'kaynağı bir çalışma sayfasında değil'],
  [/^two columns of its source have the header "(.*)"$/, (m) => `kaynağında iki sütunun başlığı aynı: "${m[1]}"`],
  [/^the pivot tables of the workbook cannot be read$/, () => 'çalışma kitabının pivot tabloları okunamıyor'],
  [/^its name is not one$/, () => 'adı geçersiz'],
  [/^a pivot table of the workbook is called "(.*)" already$/, (m) => `çalışma kitabında "${m[1]}" adlı bir pivot tablo zaten var`],
  [/^a sheet cannot be put at (.*)$/, (m) => `${m[1]} konumuna sayfa eklenemez`],
  [/^its place is not on a worksheet$/, () => 'yeri bir çalışma sayfasında değil'],
  [/^its place is not a cell$/, () => 'yeri bir hücre değil'],
  [/^its place holds something \((\S+)\)$/, (m) => `yerinde dolu bir hücre var (${m[1]})`],
  [/^it would stand over its source$/, () => 'kaynağının üzerine gelirdi'],
  [/^it would stand over the merged cells (\S+)$/, (m) => `${m[1]} birleştirilmiş hücrelerinin üzerine gelirdi`],
  [/^it could not be built$/, () => 'kurulamadı'],
];

/** A reason of the engine about pivot tables, in Turkish; as it is where no translation fits. */
function pivotReasonTr(reason: string): string {
  for (const [pattern, say] of PIVOT_REASONS_TR) {
    const found = pattern.exec(reason);
    if (found) return say(found);
  }
  return reason;
}

/** The reason a pivot table was not refreshed, in Turkish; the text of the engine where no translation fits. */
function pivotRefusalTr(message: string): string {
  // A table that is not made, or whose fields are not changed.
  const made = /^Cannot make the pivot table: (.*)$/.exec(message);
  if (made) {
    for (const [pattern, say] of PIVOT_REASONS_TR) {
      const found = pattern.exec(made[1]);
      if (found) return `Pivot tablo eklenemiyor: ${say(found)}`;
    }
    return message;
  }
  const changed = /^Cannot change the fields of the pivot table "(.*?)": (.*)$/.exec(message);
  if (changed) {
    for (const [pattern, say] of PIVOT_REASONS_TR) {
      const found = pattern.exec(changed[2]);
      if (found) return `"${changed[1]}" pivot tablosunun alanları değiştirilemiyor: ${say(found)}`;
    }
    return message;
  }
  const match = /^Cannot refresh the pivot table(?: "(.*?)")?: (.*)$/.exec(message);
  if (!match) return message;
  for (const [pattern, say] of PIVOT_REASONS_TR) {
    const found = pattern.exec(match[2]);
    if (found) return `${match[1] !== undefined ? `"${match[1]}" pivot tablosu` : 'Pivot tablo'} yenilenemiyor: ${say(found)}`;
  }
  return message;
}

/** Strings for a language tag such as "tr" or "en-US"; English for anything unknown. */
export function getUiStrings(tag: string | undefined): UiStrings {
  return typeof tag === 'string' && /^tr(?:[-_]|$)/i.test(tag) ? STRINGS_TR : STRINGS_EN;
}

/**
 * The strings a view speaks for the `locale` it was given: those of a language
 * tag, and for anything else (no locale, a locale object, which only changes
 * formats) those of the browser's language.
 */
export function resolveUiStrings(locale: string | object | undefined): UiStrings {
  return getUiStrings(typeof locale === 'string' ? locale : typeof navigator !== 'undefined' ? navigator.language : undefined);
}
