// The few words the views themselves say: labels for assistive technology, the
// header menu, notes, the document view's bar. Content of files is never translated.

export interface UiStrings {
  spreadsheet: string;
  cellEditor: string;
  actionFailed: string;
  /** Label of the grid: "Sheet <name>". */
  sheet(name: string): string;
  chartSheetNote: string;
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
  insertPushesOff(count: number, axis: string): string;
  arrayFormulaBlocks(kind: 'insert' | 'delete', axis: string, ref: string): string;
  /** `what` names the content in the way: comments, drawings, tables. */
  notSupportedWith(kind: 'insert' | 'delete', axis: string, what: string): string;
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
  /** Said when a change is refused because the document tracks its changes and this kind of change cannot be tracked yet. */
  cannotTrackChange: string;
  /** Said once for a document whose protection keeps it read-only; `kind` is "readOnly", "comments" or "forms". */
  protectedDocument(kind: string): string;
  // Tracked changes.
  /**
   * What a tracked change is: `kind` is a kind of DocRevision ("insert", "delete", "moveFrom", "moveTo", "format",
   * "paragraphFormat", "tableFormat", "sectionFormat", "other"); the author, the date and the text come from the file.
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
  /** Said after a copy was pasted as text alone because changes are tracked. */
  pastedAsText: string;
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
  cannotHideLastSheet: string;
  cannotSortThreaded: string;
  /** Label of the field a sheet is renamed in. */
  sheetName: string;
  /** Said after a sheet was deleted. */
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
}

const rows = (count: number) => (count === 1 ? 'row' : `${count} rows`);
const columns = (count: number) => (count === 1 ? 'column' : `${count} columns`);

export const STRINGS_EN: UiStrings = {
  spreadsheet: 'Spreadsheet',
  cellEditor: 'Cell editor',
  actionFailed: 'The action could not be completed',
  sheet: (name) => `Sheet ${name}`,
  chartSheetNote: 'This sheet is a chart sheet. Charts are not shown yet.',
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
  insertPushesOff: (count, axis) => `Cannot insert ${count} ${axis === 'row' ? 'row' : 'column'}(s): cells would be pushed off the sheet`,
  arrayFormulaBlocks: (kind, axis, ref) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} through the array formula in ${ref}`,
  notSupportedWith: (kind, axis, what) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} in a sheet with ${what}: not supported yet`,
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
  cannotTrackChange: 'Changes are tracked in this document, and this kind of change cannot be tracked yet; nothing was changed',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'This document is protected: only comments are allowed. It is shown read-only'
      : kind === 'forms'
        ? 'This document is protected: only its form fields may be filled in. It is shown read-only'
        : 'This document is protected against changes. It is shown read-only',
  revisionDescription: (kind, author, date, text) => {
    const what: { [kind: string]: string } = {
      insert: 'Inserted',
      delete: 'Deleted',
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
  pastedAsText: 'Changes are tracked in this document: the copy was pasted as text, without its formatting and pictures',
  cannotCutObjects: 'The selection holds something that cannot be cut (a note, an equation, a text box, a content control, a tracked change); it was left in place',
  cannotWriteClipboard: 'The clipboard cannot be written here; nothing was cut',
  cannotPasteHere: 'What was copied cannot be pasted here',
  pasteWaiting: 'Reading the pictures of the copy; it is pasted when they are there',
  pasteInterrupted: 'The selection or the document changed while the pictures of the copy were read; nothing was pasted',
  tabLeaves: 'Press Tab to leave the editor',
  documentPasteTooLarge: 'The text on the clipboard is too long to paste',
  openedConverted: (from, to) => `This ${from} file was converted to ${to} to be edited; it is saved as ${to}`,
  notConverted: (from) => `This ${from} file could not be converted to be edited; it is shown read-only`,
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
      conditionalFormats: ['conditional format', 'conditional formats'],
      dataValidations: ['data validation', 'data validations'],
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
    };
    const known = Object.prototype.hasOwnProperty.call(names, what) ? names[what] : undefined;
    return known ? `${count} ${count === 1 ? known[0] : known[1]}` : what;
  },
  notesNotPasted: (count) => (count === 1 ? 'One note was not pasted: its cell cannot take a note' : `${count} notes were not pasted: their cells cannot take notes`),
  cannotHideLastSheet: 'The last visible sheet cannot be hidden',
  cannotSortThreaded: 'A range with a threaded comment cannot be sorted yet',
  sheetName: 'Sheet name',
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
};

const satir = (count: number) => (count === 1 ? 'satır' : `${count} satır`);
const sutun = (count: number) => (count === 1 ? 'sütun' : `${count} sütun`);

export const STRINGS_TR: UiStrings = {
  spreadsheet: 'Hesap tablosu',
  cellEditor: 'Hücre düzenleyici',
  actionFailed: 'İşlem tamamlanamadı',
  sheet: (name) => `Sayfa ${name}`,
  chartSheetNote: 'Bu sayfa bir grafik sayfası. Grafikler henüz gösterilmiyor.',
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
  insertPushesOff: (count, axis) => `${count} ${axis === 'row' ? 'satır' : 'sütun'} eklenemiyor: hücreler sayfanın dışına itilirdi`,
  arrayFormulaBlocks: (kind, axis, ref) =>
    `${ref} aralığındaki dizi formülünün içinden ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}`,
  notSupportedWith: (kind, axis, what) =>
    `Bu sayfada ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'ekleme' : 'silme'} henüz desteklenmiyor (sayfada bulunan: ${what})`,
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
  cannotTrackChange: 'Bu belgede değişiklikler izleniyor ve bu tür değişiklik henüz izlenemiyor; hiçbir şey değiştirilmedi',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'Bu belge korumalı: yalnız yorum eklenebilir. Salt okunur gösteriliyor'
      : kind === 'forms'
        ? 'Bu belge korumalı: yalnız form alanları doldurulabilir. Salt okunur gösteriliyor'
        : 'Bu belge değişikliğe karşı korumalı. Salt okunur gösteriliyor',
  revisionDescription: (kind, author, date, text) => {
    const what: { [kind: string]: string } = {
      insert: 'Eklendi',
      delete: 'Silindi',
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
  pastedAsText: 'Bu belgede değişiklikler izleniyor: kopya biçimi ve görselleri olmadan, metin olarak yapıştırıldı',
  cannotCutObjects: 'Seçimde kesilemeyen bir öğe var (dipnot, denklem, metin kutusu, içerik denetimi, izlenen değişiklik); yerinde bırakıldı',
  cannotWriteClipboard: 'Pano buradan yazılamıyor; hiçbir şey kesilmedi',
  cannotPasteHere: 'Kopyalanan içerik buraya yapıştırılamıyor',
  pasteWaiting: 'Kopyanın görselleri okunuyor; okununca yapıştırılacak',
  pasteInterrupted: 'Kopyanın görselleri okunurken seçim ya da belge değişti; hiçbir şey yapıştırılmadı',
  tabLeaves: 'Düzenleyiciden çıkmak için Sekme tuşuna basın',
  documentPasteTooLarge: 'Panodaki metin yapıştırılamayacak kadar uzun',
  openedConverted: (from, to) => `Bu ${from} dosyası düzenlenebilmesi için ${to} biçimine çevrildi; ${to} olarak kaydedilir`,
  notConverted: (from) => `Bu ${from} dosyası düzenlenebilecek biçime çevrilemedi; salt okunur gösteriliyor`,
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
      conditionalFormats: 'koşullu biçim',
      dataValidations: 'veri doğrulama',
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
    };
    return Object.prototype.hasOwnProperty.call(names, what) ? `${count} ${names[what]}` : what;
  },
  notesNotPasted: (count) => `${count} not yapıştırılmadı: hücresi not alamıyor`,
  cannotHideLastSheet: 'Görünen son sayfa gizlenemez',
  cannotSortThreaded: 'Yorum zinciri içeren aralık henüz sıralanamıyor',
  sheetName: 'Sayfa adı',
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
};

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
