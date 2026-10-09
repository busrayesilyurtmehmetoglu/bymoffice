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
  // Saving the file to where it came from (MountOptions.saveTo, the bridge's `save`).
  saved: string;
  /** `detail` is short and never from the file: "HTTP 500", or what the application said; may be empty. */
  saveFailed(detail: string): string;
  /** The server holds another file than the one that was opened (HTTP 412). */
  saveConflict: string;
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
  saved: 'Saved',
  saveFailed: (detail) => (detail ? `The file could not be saved (${detail})` : 'The file could not be saved'),
  saveConflict: 'The file was not saved: the file on the server is not the one you opened (someone else changed it, or an earlier save of yours was written although its answer did not arrive)',
  saveRedirected: 'The file was not saved: the server sent the request to another address; you may have been signed out',
  saveConvertedNotAllowed: (to) => `Changes to this file cannot be saved here: it was converted to ${to}, and this page has not allowed saving that in place of the original file`,
  saveDenied: 'The file was not saved: you are not allowed to save it',
  saveLocked: 'The file was not saved: it is locked',
  saveTimedOut: 'The file could not be saved: no answer came in time',
  saveNoConnection: 'The file could not be saved: the server could not be reached',
  savePending: 'The file was not saved: what is being typed could not be applied',
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
      conditionalFormatExtensions: ['conditional format setting of a newer Excel', 'conditional format settings of a newer Excel'],
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
      ? 'One data validation or conditional format could not be pasted as it is where the cells come from: it is of a kind that is kept as the file has it'
      : `${count} data validations or conditional formats could not be pasted as they are where the cells come from: they are of a kind that is kept as the file has it`,
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
  saved: 'Kaydedildi',
  saveFailed: (detail) => (detail ? `Dosya kaydedilemedi (${detail})` : 'Dosya kaydedilemedi'),
  saveConflict: 'Dosya kaydedilmedi: sunucudaki dosya sizin açtığınızdan farklı (başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir)',
  saveRedirected: 'Dosya kaydedilmedi: sunucu isteği başka bir adrese yönlendirdi; oturumunuzun süresi dolmuş olabilir',
  saveConvertedNotAllowed: (to) => `Bu dosyadaki değişiklikler buraya kaydedilemez: dosya ${to} biçimine çevrildi ve sayfa bunun özgün dosyanın yerine kaydedilmesine izin vermedi`,
  saveDenied: 'Dosya kaydedilmedi: kaydetme yetkiniz yok',
  saveLocked: 'Dosya kaydedilmedi: kilitli',
  saveTimedOut: 'Dosya kaydedilemedi: zamanında yanıt gelmedi',
  saveNoConnection: 'Dosya kaydedilemedi: sunucuya ulaşılamadı',
  savePending: 'Dosya kaydedilmedi: yazılmakta olan uygulanamadı',
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
      conditionalFormatExtensions: 'yeni Excel sürümüne özgü koşullu biçim ayarı',
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
  rulesNotPasted: (count) => `${count} veri doğrulama ya da koşullu biçim kuralı kaynaktaki gibi yapıştırılamadı: bu tür kurallar dosyadaki haliyle korunur`,
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
