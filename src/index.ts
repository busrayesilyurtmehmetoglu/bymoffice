import { CfbFile } from './cfb';
import { MAX_PASSWORD_CHARS } from './crypto/agile';
import { PasswordError, type EncryptedFormat } from './crypto/errors';
import { decryptDocStreams, decryptXlsStream, docEncryptionFlags, findFilePass } from './crypto/legacy';
import { DEFAULT_PASSWORD, decryptPackage, encryptPackage, isEncryptedPackage, packageEncryption, type FileEncryption } from './crypto/ooxml';
import { PasswordSession, bindProtection, sessionOf, type Protection } from './crypto/session';
import { PasswordPrompt, type PasswordRequest } from './ui/password-dialog';
import { openAsking } from './ui/password-flow';
import { readDoc } from './formats/doc/document';
import { readRtf } from './formats/rtf/document';
import { readDocx } from './formats/docx/document';
import { buildDocx, writeDocx } from './formats/docx/write';
import { needsConversion, openConverted } from './formats/convert';
import { documentPictures } from './model/doc/pictures';
import { hasPermissionRanges, type DocModel } from './model/doc/types';
import { readXls, readXlsStream } from './formats/xls/workbook';
import { readXlsx } from './formats/xlsx/workbook';
import { buildXlsx } from './formats/xlsx/build';
import { resolveUiStrings } from './ui/strings';
import { registerFont, registeredFonts, settleFonts, unregisterFont, type FontFailure, type FontRegistration } from './render/fonts';
import { configureEmbeddedFonts, embeddedFontsState } from './render/embedded-fonts';
import { writeXlsx } from './formats/xlsx/write';
import type { Workbook } from './model/sheet/types';
import { SpreadsheetEditor, type SpreadsheetEditorOptions } from './ui/editor';
import type { WorkbookView } from './ui/sheet-view';
import type { DocumentView, DocumentZoom } from './ui/doc-view';
import { DocumentEditor } from './ui/doc-editor';
import type { DocumentPrintOptions, DocumentPrintResult } from './ui/doc-print';
import type { SheetPrintOptions } from './ui/sheet-print';
import type { PdfExportOptions } from './pdf/export';
import type { SheetPdfExportOptions } from './ui/sheet-view';
import { OpcPackage } from './opc';
import { askOverwrite } from './ui/save-confirm';
import { listPeople, type CollabPerson, type People } from './collab/people';
import { createSaveSync, type SaveSync } from './collab/save-sync';
import { keepStamp, readCollabStamp, samePackage, writeSharedFile } from './formats/xlsx/collab-stamp';
import { CollabSession, contentVersion, isSocketUrl, socketTransport, type CollabPeer, type CollabState, type CollabTransport, type CollabUser } from './collab/index';
import {
  SaveController,
  bindSaveToPage,
  fileETag,
  locksSaving,
  nameFromUrl,
  resolveAutoSave,
  resolveSaveTarget,
  strongETag,
  saveInfoOf,
  type AutoSave,
  type SaveNowOptions,
  type SaveState,
  type SaveTo,
} from './embed/save-to';

export { CfbFile, type CfbEntry } from './cfb';
export { allocateCfb, writeCfb, type CfbLayout, type CfbStreamSize } from './cfb-write';
export { PasswordError, type EncryptedFormat, type PasswordErrorCode } from './crypto/errors';
export { xorPasswordCandidates } from './crypto/legacy';
export { DEFAULT_PASSWORD, type FileEncryption } from './crypto/ooxml';
export { MAX_PASSWORD_CHARS } from './crypto/agile';
export { MAX_SPIN_COUNT } from './crypto/work';
export type { PasswordRequest } from './ui/password-dialog';
export { OpcPackage, type Relationship } from './opc';
export { ZipArchive, type ZipEntry } from './zip';
export * from './xml';
export * from './model/sheet/types';
export * from './model/sheet/color';
export type {
  Chart,
  ChartAxis,
  ChartCache,
  ChartData,
  ChartDataLabels,
  ChartErrorBars,
  ChartGroup,
  ChartKind,
  ChartLabelPosition,
  ChartLegend,
  ChartLine,
  ChartManualLayout,
  ChartMarker,
  ChartPoint,
  ChartPointLabel,
  ChartSeries,
  ChartShape,
  ChartTextStyle,
  ChartTitle,
  ChartTrendline,
  ChartUpDownBars,
} from './model/sheet/chart';
export { readXlsx } from './formats/xlsx/workbook';
export { readXls, readXlsWithReport, type XlsReadReport } from './formats/xls/workbook';
export { readDocx, type DocxReadOptions } from './formats/docx/document';
export { MAX_DOC_CHARTS, MAX_DOC_CHART_BOOKS, MAX_DOC_CHART_BOOKS_BYTES, MAX_DOC_CHART_BOOK_BYTES, MAX_DOC_CHART_BYTES, MAX_DOC_CHART_POINTS, type DocChartLimits } from './formats/docx/charts';
export { XLS_MAX_CHART_CACHE_BYTES, XLS_MAX_CHART_DEPTH, XLS_MAX_CHART_RECORDS } from './formats/xls/chart';
export { buildDocx, writeDocx, type DocxWriteReport } from './formats/docx/write';
export { createBlank, type BlankOptions } from './formats/blank';
export { convertForEditing, openConverted, type ConvertibleFile, type EditableConversion, type OpenedForEditing } from './formats/convert';
export * as docOps from './ops/doc/index';
export { readDoc } from './formats/doc/document';
export { readRtf } from './formats/rtf/document';
export { documentPictures } from './model/doc/pictures';
export { enforcedProtection, hasPermissionRanges, inFieldCode, paragraphText, rawOf, safeHyperlinkTarget, tracksChanges, type DocModel, type DocProtection } from './model/doc/types';
export { NumberingCounters } from './formats/docx/numbering';
export { resolveNumberRunProps, resolveParaProps, resolveRunProps } from './formats/docx/resolve';
export { CalcEngine, type CalcEngineOptions } from './calc/engine';
export { ConditionalFormats, type CellConditional, type ConditionalSource } from './format/conditional';
export { DataValidations, MAX_LIST_ITEMS, ValidationError, type ValidationList, type ValidationOptions, type ValidationVerdict } from './format/validation';
export { FormulaParseError, parseFormula } from './calc/parse';
export { FUNCTIONS } from './calc/functions/index';
export { calcError, isError, type CalcError, type ErrorCode, type Scalar, type Value } from './calc/values';
export { writeXlsx } from './formats/xlsx/write';
export { buildXlsx, type BuildXlsxOptions, type XlsxBuildReport } from './formats/xlsx/build';
export { History, applyOp, ensureCellFormat, ensureNumberFormat, type CellChange, type Op, type SheetsSnapshot } from './ops/ops';
export {
  canCopySheet,
  canDeleteSheet,
  canHideSheet,
  canInsertSheet,
  canMoveSheet,
  canRenameSheet,
  copySheet,
  copySheetLosses,
  copySheetName,
  defaultSheetName,
  deleteSheet,
  insertSheet,
  moveSheet,
  renameSheet,
  setSheetState,
  setTabColor,
  type CopySheetLoss,
  type CopySheetOptions,
  type InsertSheetOptions,
} from './ops/sheets';
export { DEFAULT_NOTE_AUTHOR, MAX_NOTE_TEXT, clearNotes, commentAtCell, hasLiveComments, noteProblem, setComment, type NoteOptions } from './ops/notes';
export {
  CHART_PART_CHANGES,
  CHART_TYPES,
  DEFAULT_CHART_HEIGHT,
  DEFAULT_CHART_WIDTH,
  addChart,
  anchorOfRect,
  chartProblem,
  chartRangeRef,
  moveChart,
  removeChart,
  seriesOfRange,
  sheetCharts,
  updateChart,
  type ChartGrouping,
  type ChartInfo,
  type ChartLegendPlace,
  type ChartPatch,
  type ChartSeriesEdit,
  type ChartSeriesSpec,
  type ChartSpec,
  type ChartType,
  type RangeSeries,
} from './ops/charts';
export {
  CONDITIONAL_TYPES,
  ICON_SET_NAMES,
  MAX_SHEET_CONDITIONAL_RULES,
  MAX_SHEET_VALIDATIONS,
  MAX_VALIDATION_LIST,
  MAX_VALIDATION_MESSAGE,
  MAX_VALIDATION_TITLE,
  TIME_PERIODS,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
  addConditionalRule,
  clearConditionalRules,
  conditionalFormatNow,
  conditionalRules,
  dataValidationNow,
  keptFor,
  moveConditionalRule,
  removeConditionalRule,
  ruleFormatOf,
  ruleKeptNow,
  ruleSpace,
  textRuleCell,
  MAX_DIFFERENTIAL_FORMATS,
  setValidation,
  updateConditionalRule,
  validationKeptNow,
  validationProblem,
  validationsIn,
  type ConditionalPatch,
  type ConditionalRuleInfo,
  type ConditionalSettings,
  type ConditionalType,
  type CopiedConditional,
  type CopiedLeft,
  type CopiedRules,
  type CopiedValidation,
  type IconSetName,
  type KeptView,
  type RuleFormat,
  type RuleSpace,
  type TimePeriod,
  type ValidationOperator,
  type ValidationSettings,
  type ValidationType,
} from './ops/rules';
export { clearRange, copyRange, cutRange, editTextAt, pasteAt, setCellInput, type CopiedNote, type CopiedRange, type EditOptions, type PasteResult } from './ops/commands';
export { MAX_FILTER_VALUES, clearFilter, filterSummary, filterValues, setAutoFilter, setFilterColumn, type FilterCriterion, type FilterValue, type FilterValueList } from './ops/filter';
export { collationKey, currentRegion, sortRange, type SortKey, type SortOptions } from './ops/sort';
export { parseCellInput, cellEditText, type ParsedInput, type InputOptions } from './edit/input';
export { formulaFromLocal, formulaToLocal } from './edit/formula-locale';
export { parseTsv, toTsv } from './edit/clipboard';
export { adjustFormula, adjustRange, adjustReferenceText, formulaNamesSheetsUnread, quoteSheetName, renameSheetsInFormula, shiftFormula, type StructuralEdit } from './calc/refs';
export { DEFAULT_GRID_THEME, cellButtonRect, renderGrid, renderGridPrint, type GridBand, type GridFilter, type GridPrintScene, type GridScene, type GridSelection, type GridTheme } from './render/grid';
export type { GridCommentSource, GridDrawingSource } from './render/grid';
export { SheetDrawings, adjustCommentCell, adjustDrawingAnchor, commentCellNow, commentEdits, drawingAnchorNow, anchorRect, type DrawingTarget, type PlacedComment, type SheetDrawingLabels, type SheetDrawingsOptions } from './render/sheet-drawings';
export type { DecodedImage } from './render/images';
export { cellText, formatGeneral } from './render/cell-text';
export { cellDisplayText, formatCell, isoToSerial } from './format/cell-format';
export { LOCALE_EN, LOCALE_TR, getLocale, normalizeLocale, type FormatLocale } from './format/locale';
export { formatValue, isDateFormat, serialToDate, type FormatOptions, type FormattedValue } from './format/number-format';
export { WorkbookView, type HistoryMark, type WorkbookViewOptions } from './ui/sheet-view';
export { SpreadsheetEditor, type SpreadsheetEditorOptions } from './ui/editor';
export { EDITOR_STRINGS_EN, EDITOR_STRINGS_TR, getEditorStrings, type EditorStrings } from './ui/editor-strings';
export { ensureStyle, formatAt, formatRange, mergeRange, setColumnWidth, setColumnsHidden, setRowHeight, setRowsHidden, unmergeRange, type BorderPreset, type StylePatch } from './ops/format';
export { STRINGS_EN, STRINGS_TR, getUiStrings, type UiStrings } from './ui/strings';
export * from './layout/grid-geometry'; export * from './layout/sheet-metrics';
export { base64Decode, base64Encode, createBridge, installBridge, type Bridge, type BridgeDocumentView, type BridgeErrorCode, type BridgeHost, type BridgeState, type InstallBridgeOptions } from './embed/bridge';

export { DocLayout } from './layout/doc/layout';
export { revisionColor } from './layout/doc/prep';
export type { DocLayoutOptions, FontMetrics, ImageItem, LayoutBudget, LayoutFont, LayoutPage, LayoutProgress, LineItem, PageItem, RectItem, ShapeItem, TextItem, TextMeasurer } from './layout/doc/types';
export { renderPage, type PageImage, type RenderPageOptions } from './render/doc';
export { DocChartPainter, type DocChartPainterOptions } from './render/doc-chart';
export { baseFontFamily, registerFont, registeredFonts, unregisterFont, type FontRegistration, type FontSource, type RegisteredFont } from './render/fonts';
export { EMBEDDED_FONT_FACES, EMBEDDED_FONT_FAMILY, configureEmbeddedFonts, embeddedFontsState, type EmbeddedFontsOptions, type EmbeddedFontsState } from './render/embedded-fonts';
export { CanvasTextMeasurer, createCanvasMeasurer, docCssFont, type CanvasMeasurerOptions, type FontHint, type MeasureContext } from './render/doc-measure';
export { DocumentView, documentText, type DocumentViewOptions, type DocumentViewState, type DocumentZoom } from './ui/doc-view';
export { DocumentEditor, type DocumentEditorOptions } from './ui/doc-editor';
export type { CommentEdit, CommentOptions, DocComment } from './ops/doc/index';
export type { SaveCommand } from './ui/save-command';
export {
  SaveError,
  type AutoSave,
  type ConflictQuestion,
  type SaveErrorCode,
  type SaveInfo,
  type SaveNowOptions,
  type SaveState,
  type SaveTo,
  type SaveToFunction,
  type SaveToRequest,
} from './embed/save-to';
export { DocumentPrintError, planPrint, printCss, type DocumentPrintErrorCode, type DocumentPrintOptions, type DocumentPrintResult, type PrintPlan } from './ui/doc-print';
export { PdfExportError, type PdfAbortSignal, type PdfDownload, type PdfExportErrorCode, type PdfExportOptions } from './pdf/export';
export type { SheetPdfExportOptions } from './ui/sheet-view';
export { defaultPaperFor, headerFooterFor, paintSheetPage, planSheetPages, type SheetJobPage, type SheetJobPlan, type SheetPagePaint, type SheetPlanOptions, type SheetPrintOptions } from './ui/sheet-print';
export {
  DEFAULT_PAGE_MARGINS,
  MAX_SHEET_PAGES,
  SheetPagesError,
  paginateSheet,
  paperSize,
  printRanges,
  usedRange,
  type PrintBand,
  type SheetPagesOptions,
  type SheetPrintLayout,
  type SheetPrintPage,
} from './layout/sheet-pages';
export {
  headerFooterLines,
  parseHeaderFooter,
  type HeaderFooterField,
  type HeaderFooterFormat,
  type HeaderFooterPiece,
  type HeaderFooterRun,
  type HeaderFooterSections,
  type HeaderFooterValues,
} from './layout/header-footer';
export { DocToolbar, type DocShellHost } from './ui/doc-toolbar';
export { DocLines, type Affinity, type CaretBox, type PageLine, type Place } from './ui/doc-caret';
export { sliceHtml, sliceOf, type DocSlice } from './ui/doc-clipboard';
export type { LineMark } from './layout/doc/types';
export type { EditProgress } from './layout/doc/layout';

/**
 * The fonts of the application: registerFont() adds a font to the ones the
 * device has (also as a stand-in for font names of files), registeredFonts()
 * lists them and unregisterFont() takes one out again. The same functions
 * are exported on their own.
 *
 * configure() sets where the font the package brings (TeX Gyre Bonum, shown
 * for Bookman Old Style on devices that lack it) is fetched from, or turns
 * it off; embedded() tells its state. See EmbeddedFontsOptions.
 */
export { collab } from './collab/index';
export {
  CollabSession,
  CollabUnsupportedError,
  MemoryServer,
  broadcastTransport,
  contentVersion,
  socketTransport,
  type CollabNotice,
  type CollabPeer,
  type CollabPerson,
  type CollabPresence,
  type CollabState,
  type CollabTransport,
  type CollabUser,
  type MemoryServerOptions,
  type PausableTransport,
  type SessionChange,
} from './collab/index';

/** Editing a workbook together with others: where the relay is, and who this participant is. See MountOptions.collab. */
export interface CollabOptions {
  /** The address of the relay ("ws://…" or "wss://…"); everybody who opens the same address (and room) edits together. */
  url?: string;
  /** A channel of the application in place of `url`: see CollabTransport. */
  transport?: CollabTransport;
  /** A room at the address, when one address serves several files. */
  room?: string;
  /** The name and the colour ("#RRGGBB") the others see. */
  user?: CollabUser;
  /** Sent to the relay in the first message, for it to decide who may join. Never shown, never logged here. */
  token?: string;
  onState?(state: CollabState): void;
  onPeers?(peers: CollabPeer[]): void;
  /** Who is in the file changed (someone came, left, was renamed, went to another sheet, can only read now): the list, this participant first. */
  onPeople?(people: CollabPerson[]): void;
  /**
   * With `autoSave`: how long, in milliseconds, changes wait here for the one participant who saves for all,
   * before this side saves them itself. At least 1000; by default five delays of auto-save and at least 15 seconds
   * (a quarter more or less, so that two who wait do not save at the same moment).
   */
  saveWatchMs?: number;
}

/** What mount() answers for a workbook that is edited together. */
export interface MountedCollab {
  readonly state: CollabState;
  /** How many steps of the common order the workbook holds. */
  readonly rev: number;
  /** How many changes made here wait for the relay. */
  readonly pending: number;
  readonly peers: CollabPeer[];
  /** This participant as the others see it. */
  readonly user: CollabPeer;
  onState?: (state: CollabState) => void;
  onPeers?: (peers: CollabPeer[]) => void;
  /**
   * Everybody in the file, this participant first (`self`), then the others in the order the relay named them:
   * `{ id, name, color, self, readOnly, sheet? }`. `name` is what the participant said (or the relay gave it), or
   * "Guest N" in the language of the view: what someone says of itself, no proof of who it is. Out of the room
   * (offline, detached, closed) the list is this participant alone. At most 200 are listed.
   */
  readonly people: CollabPerson[];
  onPeople?: (people: CollabPerson[]) => void;
  /**
   * For an application that saves by itself (with save(), without `saveTo`): tells the others that
   * the bytes save() last answered with were saved, and under which ETag when the server names one.
   */
  saved(bytes: Uint8Array, etag?: string): void;
  /**
   * Gives the workbook up for the file as it is stored and joins the room with that. For a session
   * that is `detached` (the room is at another version of the file, and this workbook holds changes
   * no file has): the other way out is to save, after which the room is joined by itself. Changes
   * that are in no file are lost.
   */
  reload(): void;
  /** Leaves the room. The workbook stays; what is changed from here on is told to nobody. */
  leave(): void;
}

interface CollabPlan {
  transport(): CollabTransport;
  options: CollabOptions;
}

/** What `collab` asks for; undefined when it asks for nothing. Throws TypeError for what cannot be a relay. */
function resolveCollab(value: unknown): CollabPlan | undefined {
  if (value === undefined || value === null || value === false) return undefined;
  if (typeof value === 'string') {
    if (!isSocketUrl(value)) throw new TypeError('"collab" is the address of a relay: "ws://…" or "wss://…"');
    return { transport: () => socketTransport(value), options: {} };
  }
  if (typeof value !== 'object') throw new TypeError('"collab" is an address or an object with "url" or "transport"');
  const given = value as CollabOptions;
  const transport = given.transport;
  if (transport !== undefined && transport !== null) {
    if (typeof transport !== 'object' || typeof transport.send !== 'function' || typeof transport.onMessage !== 'function' || typeof transport.close !== 'function') {
      throw new TypeError('"collab.transport" has the methods send, onMessage and close');
    }
    return { transport: () => transport, options: given };
  }
  const url = given.url;
  if (!isSocketUrl(url)) throw new TypeError('"collab.url" is the address of a relay: "ws://…" or "wss://…"');
  return { transport: () => socketTransport(url), options: given };
}

/** What a file is known by among those who edit it together: the ETag its address named, and its bytes. */
/** Hands the user a file to keep: a link that is pressed for them. False where a page cannot do that. */
function downloadFile(doc: Document, bytes: Uint8Array, name: string): boolean {
  const win = doc.defaultView;
  const urls = win ? win.URL : undefined;
  if (!win || !urls || typeof urls.createObjectURL !== 'function' || typeof Blob === 'undefined') return false;
  const url = urls.createObjectURL(new Blob([bytes as unknown as BlobPart], { type: 'application/octet-stream' }));
  const link = doc.createElement('a');
  link.href = url;
  // The name is the application's or the file's: a name of a file, never a path.
  link.download = name.replace(/[\\/:*?"<>|\u0000-\u001F]+/g, '_').slice(0, 200) || 'workbook.xlsx';
  link.style.display = 'none';
  doc.body.appendChild(link);
  try {
    link.click();
  } finally {
    doc.body.removeChild(link);
    win.setTimeout(() => urls.revokeObjectURL(url), 60000);
  }
  return true;
}

function fileVersions(etag: string | null | undefined, bytes: Uint8Array): string[] {
  const versions: string[] = [];
  if (typeof etag === 'string' && etag !== '' && etag.length <= 190) versions.push(`e:${etag}`);
  versions.push(contentVersion(bytes));
  return versions;
}

export const fonts = { registerFont, registeredFonts, unregisterFont, configure: configureEmbeddedFonts, embedded: embeddedFontsState };

export type OfficeFormat = 'docx' | 'xlsx' | 'doc' | 'xls' | 'rtf';

/** What open() takes besides the file. */
export interface OpenOptions {
  /**
   * The password of a password-protected file. Without one such a file
   * rejects with PasswordError "required" (unless it is encrypted with the
   * password Office uses when the user gave none: such a file just opens).
   * Has no effect on a file that is not encrypted.
   */
  password?: string;
  /**
   * An open file whose password opens this one as well: for another version of the same file (one that was
   * saved meanwhile and is loaded again), without the application having to keep the password. Used when
   * `password` is not given. A file that has no password, or was closed, counts as none given.
   */
  passwordOf?: OfficeFile;
  /** Calls the opening of an encrypted file off: an AbortSignal fits. The call then rejects with PasswordError "cancelled". */
  signal?: { readonly aborted: boolean };
  /**
   * Refuses an encrypted docx or xlsx that carries no integrity check (agile encryption without `dataIntegrity`)
   * with PasswordError "damaged". Without it such a file is opened, and says so in `warnings`: its content
   * could have been changed by someone who does not know the password, and nothing would show.
   */
  requireIntegrity?: boolean;
}

/** Something a file that was opened has to say about itself. */
export interface FileWarning {
  /**
   * "no-integrity": the file is encrypted (agile encryption) and has no integrity check, so that its content
   * could have been changed without the password. Files the engine writes always carry the check.
   */
  code: 'no-integrity';
  /** In English, for developers; mount() tells the user in the language of the view. */
  message: string;
}

/** What save(), saveAsDocx() and saveAsXlsx() take. */
export interface SaveOptions {
  /**
   * What this one save is encrypted with: a password, or `null` for a file
   * without one. Left out, the file is saved the way it is kept: with the
   * password it was opened with, or the one setPassword() last named. A
   * password that is set here is 255 characters at most.
   */
  password?: string | null;
  /** Calls the encryption off: the save then rejects with PasswordError "cancelled". */
  signal?: { readonly aborted: boolean };
}

export interface OfficeFile {
  format: OfficeFormat;
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb' | 'rtf';
  /** Present for docx/xlsx. */
  package?: OpcPackage;
  /** Present for doc/xls. */
  cfb?: CfbFile;
  /** Set when the file that was opened was encrypted with a password: how. */
  encryption?: FileEncryption;
  /** Set when there is something to be told about the file (see FileWarning); mount() passes it on through onNotice. */
  warnings?: FileWarning[];
  /**
   * Whether the file is saved encrypted: it was opened with a password, or
   * setPassword() named one. Such a file is never saved without its
   * password unless that is asked for (`save({ password: null })`,
   * `setPassword(null)`).
   */
  readonly passwordProtected: boolean;
  /**
   * Changes what the file is saved with from now on: another password, or
   * none (`null`). Throws TypeError for an empty password, and RangeError
   * for one of more than 255 characters. For doc, xls and RTF, which are
   * not written, it changes what saveAsDocx() and saveAsXlsx() give.
   */
  setPassword(password: string | null): void;
  /**
   * Zeroes the password the engine keeps for saving. Afterwards a save of a
   * password-protected file rejects, rather than write it without one.
   * mount() calls this when its view is destroyed.
   */
  close(): void;
  /**
   * The spreadsheet model: sheets, cells, styles. Read on first use and then
   * kept. Rejects with UnsupportedFileError for documents and, until the BIFF
   * reader exists, for xls.
   */
  workbook(): Promise<Workbook>;
  /**
   * The document model: paragraphs, tables, styles, numbering. Read on first use
   * and then kept. Present for docx only, until the other readers exist.
   */
  document?(): Promise<DocModel>;
  /**
   * The document as a docx file. For a docx this is save(); for doc and RTF,
   * which cannot be written, the model is written out as a new docx (shapes
   * lose their geometry, theme colours become plain ones). A password-protected
   * file gives an encrypted docx (see SaveOptions).
   */
  saveAsDocx?(options?: SaveOptions): Promise<Uint8Array>;
  /**
   * The spreadsheet as an xlsx file. For an xlsx this is save(); for xls, which
   * cannot be written, the model is written out as a new xlsx: what the model
   * holds, not what the xls reader leaves out (charts, macros, print settings).
   * A password-protected file gives an encrypted xlsx (see SaveOptions).
   */
  saveAsXlsx?(options?: SaveOptions): Promise<Uint8Array>;
  /**
   * The file as bytes. For xlsx and docx, changes made to the model are
   * written into it; everything that was not changed keeps its original
   * bytes. A file that is password-protected is encrypted again (agile
   * encryption: AES-256, SHA-512), with new salts and keys every time; where
   * that cannot be done the call rejects and gives no bytes. doc, xls and
   * RTF give the bytes they were opened from, still encrypted if they were;
   * asked for another password, or for none, they reject.
   */
  save(options?: SaveOptions): Promise<Uint8Array>;
}

/** The input is not a Word or Excel file this engine can open. */
export class UnsupportedFileError extends Error {
  override name = 'UnsupportedFileError';
}

/** "{\\rtf" at the start, after a byte order mark and white space some writers put before it. */
function isRtf(bytes: Uint8Array): boolean {
  let at = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const end = Math.min(bytes.length, at + 64);
  while (at < end && (bytes[at] === 0x20 || bytes[at] === 0x09 || bytes[at] === 0x0a || bytes[at] === 0x0d)) at++;
  return bytes[at] === 0x7b && bytes[at + 1] === 0x5c && bytes[at + 2] === 0x72 && bytes[at + 3] === 0x74 && bytes[at + 4] === 0x66;
}

const notASpreadsheet = (): Promise<Workbook> =>
  Promise.reject(new UnsupportedFileError('This file is a document, not a spreadsheet'));

/** What goes on inside open() and is not for applications. */
interface OpenInternals {
  /** The password the file is saved with although it was not opened with one: a file that was converted from an encrypted one. */
  session?: PasswordSession;
  /** How the original of this file was encrypted. */
  encryption?: FileEncryption;
  /** What the encrypted file this one was inside has to say (see FileWarning). */
  warnings?: FileWarning[];
}

function checkPassword(password: string): void {
  if (typeof password !== 'string') throw new TypeError('A password is a string, or null for a file without one');
  if (password === '') throw new TypeError('A file is not encrypted with an empty password; use null to save it without one');
  if (password.length > MAX_PASSWORD_CHARS) throw new RangeError(`A file is encrypted with a password of at most ${MAX_PASSWORD_CHARS} characters`);
}

/** A session of its own with the password an open file is kept with; undefined for a file without one, or one that was closed. */
function lentSession(file: OfficeFile | undefined): PasswordSession | undefined {
  const session = file ? sessionOf(file) : undefined;
  return session && !session.forgotten ? session.clone() : undefined;
}

/**
 * Opens an encrypted file with the password of the options, or, when none
 * was given, with the one Office encrypts with when the user gave none. A
 * password that was given and is wrong is wrong: nothing else is tried
 * behind it. Gives what `attempt` gave, the session of the password that
 * opened the file, and whether that was the password everyone knows.
 */
async function unlock<T>(
  attempt: (password: Uint8Array) => T | Promise<T>,
  options: OpenOptions,
  format: EncryptedFormat,
): Promise<{ value: T; session: PasswordSession; byDefault: boolean }> {
  // The password of another open file (see OpenOptions.passwordOf), when no password was given.
  const lent = options.password === undefined || options.password === null ? lentSession(options.passwordOf) : undefined;
  const given = lent !== undefined || (options.password !== undefined && options.password !== null);
  const password = lent ? undefined : given ? String(options.password) : DEFAULT_PASSWORD;
  const session = lent || PasswordSession.of(password as string);
  try {
    try {
      return { value: await attempt(session.password), session, byDefault: password === DEFAULT_PASSWORD };
    } catch (error) {
      session.forget();
      if (!(error instanceof PasswordError) || error.code !== 'wrong') throw error;
    }
    throw given ? new PasswordError('wrong', 'The password is not the password of this file') : new PasswordError('required', 'The file is protected with a password');
  } catch (error) {
    if (error instanceof PasswordError) error.format = format;
    throw error;
  }
}

/**
 * How much of a stream is looked at to tell whether a file is encrypted: the start of the header of a Word
 * document (its flags), and the first records of a workbook (FILEPASS is the second or the third, and small).
 * A file that is not encrypted has none of its streams read before its model is asked for.
 */
const DOC_HEAD = 32;
const XLS_HEAD = 1024;

/** The content of a stream, or nothing for an entry that is no stream (a storage of that name) or cannot be read: the reader says what is wrong with it. */
function streamOf(cfb: CfbFile, path: string): Uint8Array {
  try {
    return cfb.stream(path);
  } catch (error) {
    return new Uint8Array(0);
  }
}

/**
 * Whether a file is encrypted with a password, and what kind of file it is
 * as far as that shows without the password ("ooxml": a docx or an xlsx).
 * Undefined for a file that is not encrypted, and for one that is no Word or
 * Excel file. Nothing is decrypted: this is quick.
 */
export function encryptionOf(input: Uint8Array | ArrayBuffer): { format: EncryptedFormat; encryption?: FileEncryption } | undefined {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (!CfbFile.isCfb(bytes)) return undefined;
  let cfb: CfbFile;
  try {
    cfb = CfbFile.read(bytes);
  } catch (error) {
    return undefined;
  }
  if (cfb.has('WordDocument')) return docEncryptionFlags(cfb.head('WordDocument', DOC_HEAD)) ? { format: 'doc' } : undefined;
  const book = cfb.has('Workbook') ? 'Workbook' : cfb.has('Book') ? 'Book' : undefined;
  if (book) return findFilePass(cfb.head(book, XLS_HEAD)) ? { format: 'xls' } : undefined;
  if (!isEncryptedPackage(cfb)) return undefined;
  try {
    return { format: 'ooxml', encryption: packageEncryption(cfb) };
  } catch (error) {
    // Encrypted in a way that is not read: open() says how.
    return { format: 'ooxml' };
  }
}

/**
 * Opens a Word or Excel file. The format is detected from the content, never
 * from the file name. A password-protected file (docx and xlsx: agile and
 * standard encryption; doc and xls: RC4, RC4 CryptoAPI and XOR obfuscation)
 * is opened with `options.password`; without one, or with a wrong one, the
 * call rejects with PasswordError ("required", "wrong"), and with
 * "unsupported" or "damaged" for a file that cannot be decrypted at all.
 */
export async function open(input: Uint8Array | ArrayBuffer | Blob, options: OpenOptions = {}): Promise<OfficeFile> {
  return openWith(input, options, {});
}

async function openWith(input: Uint8Array | ArrayBuffer | Blob, options: OpenOptions, internals: OpenInternals): Promise<OfficeFile> {
  const bytes =
    input instanceof Uint8Array
      ? input
      : new Uint8Array(input instanceof ArrayBuffer ? input : await input.arrayBuffer());

  // What the file is saved with; the same object for as long as the file is open.
  const protection: Protection = { session: internals.session };
  /** The package as it leaves: encrypted when the file is, or when the save asks for it. */
  const protect = async (plain: Uint8Array, saveOptions?: SaveOptions): Promise<Uint8Array> => {
    const asked = saveOptions ? saveOptions.password : undefined;
    const signal = saveOptions ? saveOptions.signal : undefined;
    if (asked === null) return plain;
    if (asked !== undefined) {
      checkPassword(asked);
      const once = PasswordSession.of(asked);
      try {
        return await encryptPackage(plain, once, { signal });
      } finally {
        once.forget();
      }
    }
    return protection.session ? encryptPackage(plain, protection.session, { signal }) : plain;
  };
  /** A format that is not written gives the bytes it was opened from: it cannot be asked for other ones. */
  const asOpened =
    (format: string) =>
    async (saveOptions?: SaveOptions): Promise<Uint8Array> => {
      if (saveOptions && saveOptions.password !== undefined) {
        const instead = format === 'xls' ? 'xlsx (saveAsXlsx)' : 'docx (saveAsDocx)';
        throw new Error(`A ${format} file is saved as the bytes it was opened from: its password cannot be changed, set or taken away. Save it as ${instead} instead`);
      }
      return bytes;
    };
  const finish = (file: Omit<OfficeFile, 'passwordProtected' | 'setPassword' | 'close'>): OfficeFile => {
    const out = file as OfficeFile;
    Object.defineProperty(out, 'passwordProtected', { get: () => protection.session !== undefined, enumerable: true });
    out.setPassword = (password: string | null): void => {
      if (password !== null) checkPassword(password);
      if (protection.session) protection.session.forget();
      protection.session = password === null ? undefined : PasswordSession.of(password);
    };
    out.close = (): void => {
      // The session stays, forgotten: a save now rejects instead of writing the file without its password.
      if (protection.session) protection.session.forget();
    };
    if (internals.encryption && !out.encryption) out.encryption = internals.encryption;
    if (internals.warnings && internals.warnings.length > 0) out.warnings = internals.warnings;
    bindProtection(out, protection);
    return out;
  };

  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    const pkg = await OpcPackage.open(bytes);
    const main = await pkg.mainPart();
    const contentType = (main && pkg.contentTypeOf(main)) || '';
    const base = { container: 'ooxml' as const, package: pkg };
    if (/wordprocessingml|ms-word/.test(contentType)) {
      let model: Promise<DocModel> | undefined;
      // Saving writes the edited stories of the model into the package first; a model that
      // was never read, or read and not changed, leaves the package as it is.
      const save = async (saveOptions?: SaveOptions) => {
        if (model) {
          const doc = await model.catch(() => undefined);
          if (doc) await writeDocx(pkg, doc);
        }
        return protect(await pkg.save(), saveOptions);
      };
      return finish({ format: 'docx', kind: 'document', ...base, save, workbook: notASpreadsheet, document: () => (model ??= readDocx(pkg)), saveAsDocx: save });
    }
    if (/spreadsheetml|ms-excel/.test(contentType)) {
      let workbook: Promise<Workbook> | undefined;
      // Saving writes the changes made to the model into the package first; a model
      // that was never read, or read and not changed, leaves the package as it is.
      const save = async (saveOptions?: SaveOptions) => {
        if (workbook) {
          // A file whose model could not be read was not changed either: it is saved as it is.
          const model = await workbook.catch(() => undefined);
          if (model) await writeXlsx(pkg, model);
        }
        return protect(await pkg.save(), saveOptions);
      };
      return finish({ format: 'xlsx', kind: 'spreadsheet', ...base, save, workbook: () => (workbook ??= readXlsx(pkg)), saveAsXlsx: save });
    }
    throw new UnsupportedFileError(`Unsupported OOXML document type: ${contentType || 'unknown'}`);
  }

  if (CfbFile.isCfb(bytes)) {
    const cfb = CfbFile.read(bytes);
    // No binary writer, and nothing can be modified, so saving returns the original bytes.
    const base = { container: 'cfb' as const, cfb };
    if (cfb.has('WordDocument')) {
      // An encrypted document is decrypted here, so that a missing or wrong password is told by open() and not by the first read.
      let source: Pick<CfbFile, 'has' | 'stream'> = cfb;
      let encryption: FileEncryption | undefined;
      // Only the start of the stream is looked at: a document without a password has nothing read here.
      const flags = docEncryptionFlags(cfb.head('WordDocument', DOC_HEAD));
      if (flags) {
        const word = streamOf(cfb, 'WordDocument');
        const other = flags.table === '1Table' ? '0Table' : '1Table';
        const tableName = cfb.has(flags.table) ? flags.table : cfb.has(other) ? other : undefined;
        if (tableName === undefined) throw new Error('Invalid doc: the file has no table stream');
        const streams = { word, table: cfb.stream(tableName), data: cfb.has('Data') ? cfb.stream('Data') : new Uint8Array(0) };
        const opened = await unlock((password) => decryptDocStreams(streams, password), options, 'doc');
        protection.session = opened.session;
        encryption = opened.value.encryption;
        if (opened.byDefault) encryption.defaultPassword = true;
        const plain: Record<string, Uint8Array> = { WordDocument: opened.value.word, [tableName]: opened.value.table };
        if (cfb.has('Data')) plain.Data = opened.value.data;
        source = {
          has: (path: string) => Object.prototype.hasOwnProperty.call(plain, path),
          stream: (path: string) => {
            if (!Object.prototype.hasOwnProperty.call(plain, path)) throw new Error(`Stream not found: ${path}`);
            return plain[path];
          },
        };
      }
      let model: Promise<DocModel> | undefined;
      const document = () =>
        (model ??= Promise.resolve().then(() => {
          try {
            return readDoc(source);
          } catch (error) {
            if (error instanceof Error && error.name === 'UnsupportedFileError') throw new UnsupportedFileError(error.message);
            throw error;
          }
        }));
      const file = finish({
        format: 'doc',
        kind: 'document',
        ...base,
        save: asOpened('doc'),
        workbook: notASpreadsheet,
        document,
        saveAsDocx: async (saveOptions?: SaveOptions) => protect(await buildDocx(await document()), saveOptions),
      });
      if (encryption) file.encryption = encryption;
      return file;
    }
    if (cfb.has('Workbook') || cfb.has('Book')) {
      // A file saved for both Excel 97 and Excel 95 has both streams; the newer one says more.
      const name = cfb.has('Workbook') ? 'Workbook' : 'Book';
      let decrypted: Uint8Array | undefined;
      let encryption: FileEncryption | undefined;
      // Only the first records are looked at: a workbook without a password has nothing read here.
      if (findFilePass(cfb.head(name, XLS_HEAD))) {
        const stream = streamOf(cfb, name);
        const opened = await unlock((password) => decryptXlsStream(stream, password), options, 'xls');
        protection.session = opened.session;
        decrypted = opened.value.bytes;
        encryption = opened.value.encryption;
        if (opened.byDefault) encryption.defaultPassword = true;
      }
      let model: Promise<Workbook> | undefined;
      // Read on first use. What the reader cannot open (older than BIFF5) is reported as unsupported.
      const workbook = () =>
        (model ??= Promise.resolve().then(() => {
          try {
            return decrypted ? readXlsStream(decrypted, name, true).workbook : readXls(cfb);
          } catch (error) {
            if (error instanceof Error && error.name === 'UnsupportedFileError') throw new UnsupportedFileError(error.message);
            throw error;
          }
        }));
      const file = finish({
        format: 'xls',
        kind: 'spreadsheet',
        ...base,
        save: asOpened('xls'),
        workbook,
        saveAsXlsx: async (saveOptions?: SaveOptions) => protect(await buildXlsx(await workbook()), saveOptions),
      });
      if (encryption) file.encryption = encryption;
      return file;
    }
    if (isEncryptedPackage(cfb)) {
      // A docx or an xlsx with a password: the package is inside, encrypted. Which of the two it is, the package says.
      const opened = await unlock((password) => decryptPackage(cfb, password, { signal: options.signal, requireIntegrity: options.requireIntegrity }), options, 'ooxml');
      const plain = opened.value.bytes;
      const encryption = opened.value.encryption;
      if (opened.byDefault) encryption.defaultPassword = true;
      // Agile encryption without its HMAC: the file opens, and whoever opens it is told.
      const warnings: FileWarning[] = [];
      if (encryption.scheme === 'agile' && !encryption.integrity) {
        warnings.push({ code: 'no-integrity', message: 'This encrypted file has no integrity check: its content could have been changed by someone who does not know the password' });
      }
      if (!(plain[0] === 0x50 && plain[1] === 0x4b)) {
        opened.session.forget();
        const error = new PasswordError('damaged', 'The encrypted file is damaged: what was decrypted is not a docx or xlsx package');
        error.format = 'ooxml';
        throw error;
      }
      try {
        return await openWith(plain, {}, { session: opened.session, encryption, warnings });
      } catch (error) {
        opened.session.forget();
        // Nothing vouched for what was decrypted: that it does not read as a package is most likely what that lets through.
        if (warnings.length > 0 && !(error instanceof PasswordError)) {
          const damaged = new PasswordError(
            'damaged',
            `The encrypted file is damaged: it has no integrity check, and what was decrypted cannot be read (${error instanceof Error ? error.message.slice(0, 160) : 'unreadable'})`,
          );
          damaged.format = 'ooxml';
          throw damaged;
        }
        throw error;
      }
    }
    throw new UnsupportedFileError('Compound file is not a Word or Excel document');
  }

  // "{\rtf": Word documents are often RTF saved with a .doc extension.
  if (isRtf(bytes)) {
    let model: Promise<DocModel> | undefined;
    // Nothing writes RTF and nothing can be changed, so saving returns the original bytes.
    return finish({
      format: 'rtf',
      kind: 'document',
      container: 'rtf',
      save: asOpened('RTF'),
      workbook: notASpreadsheet,
      document: () => (model ??= Promise.resolve().then(() => readRtf(bytes))),
      saveAsDocx: async (saveOptions?: SaveOptions) => protect(await buildDocx(await (model ??= Promise.resolve().then(() => readRtf(bytes)))), saveOptions),
    });
  }

  throw new UnsupportedFileError('Not a Word or Excel file');
}

/**
 * Opens a file for a view, asking for the password of a password-protected
 * one: `onPassword` when there is one, otherwise a dialog inside
 * `container`. Rejects with PasswordError "cancelled" when the question is
 * given up, and with "required" or "wrong" where there is nobody to ask.
 */
async function openForView(
  input: Uint8Array | ArrayBuffer | Blob,
  options: Pick<MountOptions, 'password' | 'onPassword' | 'locale' | 'requireIntegrity'>,
  container: HTMLElement | undefined,
): Promise<OfficeFile> {
  // Read once: every password that is tried reads the same bytes.
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input instanceof ArrayBuffer ? input : await input.arrayBuffer());
  return openAsking(
    (password, signal) => openWith(bytes, { password, signal, requireIntegrity: options.requireIntegrity }, {}),
    options,
    container && typeof document !== 'undefined' ? (onCancel) => new PasswordPrompt({ container, strings: resolveUiStrings(options.locale), onCancel }) : undefined,
  );
}

export interface MountOptions extends Omit<SpreadsheetEditorOptions, 'saveCommand'> {
  /** The file itself, or a URL to fetch it from. The format is detected from the content. */
  file: Uint8Array | ArrayBuffer | Blob | string;
  /**
   * The password of a password-protected file (see OpenOptions.password).
   * When it is missing or wrong the user is asked: see `onPassword`. The
   * file is then saved with the password that opened it, by save(), by
   * saveNow(), by auto-save and by the Save button alike; `file.setPassword()`
   * of the result changes that, and `save({ password })` does for one save.
   */
  password?: string;
  /**
   * Called when a password-protected file needs a password: none was given,
   * or the one that was tried is wrong. `attempts` counts the passwords that
   * were tried and failed. Return the password, or null to give up (mount()
   * then rejects with PasswordError "cancelled"); a promise of either is
   * waited for. It is called again for as long as the password is wrong.
   *
   * Without it the editor asks by itself, in a dialog inside the container
   * (a field of type "password"; Escape gives up). The password is not kept
   * in the page: the field is emptied when its content is taken.
   */
  onPassword?(request: PasswordRequest): string | null | undefined | Promise<string | null | undefined>;
  /**
   * Refuses an encrypted file that has no integrity check (see OpenOptions.requireIntegrity). Without it such a
   * file is shown, and the user is told through the notice of the view and onNotice that its content could have
   * been changed by someone who does not know the password.
   */
  requireIntegrity?: boolean;
  /**
   * Where the file is saved to, by the Save button of the toolbar, by Ctrl+S
   * or Cmd+S in the view and by saveNow():
   *
   * - `true`: the address the file was fetched from (`file` must be an address);
   * - an address: http: or https:, or one relative to the page;
   * - a request, `{ url, method, headers, credentials, timeoutMs, allowConverted }`:
   *   `url` defaults to the address the file was fetched from, `method` to
   *   "PUT", `credentials` to "same-origin" (what the file is fetched with),
   *   `timeoutMs` to two minutes;
   * - a function `(bytes, info)`: the application saves by itself, and the
   *   save has failed when it throws or rejects (or takes more than ten minutes).
   *
   * A redirect is not followed: a save the server answers with one has
   * failed. A file that was converted to be edited (`convert`) is not saved
   * to the address its original was fetched from unless the request says
   * `allowConverted: true`: onNotice says so when the file is opened, and a
   * save fails with SaveError "converted". Another address and a function
   * take the converted file as it is.
   *
   * The request carries the bytes with the media type of their format,
   * `X-Bym-Format` ("docx" or "xlsx"), `X-Bym-Converted-From` for a file that
   * was converted to be edited, and `If-Match` with the ETag the file came
   * with when it is saved to where it was fetched from (not when the file was
   * fetched through a redirect). Any 2xx answer means
   * saved; its body is not read. What goes wrong is told to onNotice, to
   * onSaveState and, as a SaveError, to onError.
   *
   * A save to an address that the server refuses with 412 (the file there was
   * changed since it was opened) never overwrites by itself: after onError
   * was told, the editor asks the user inside the container ("Overwrite" /
   * "Cancel"), and "Overwrite" is saveNow({ overwrite: true }). A page that
   * asks by itself turns the question off with `{ onConflict: 'fail' }`. A
   * function is not asked about.
   *
   * Nothing is saved, and there is no button, for a file that cannot be
   * written (doc, RTF and xls that were not converted, a document whose
   * protection allows no editing); the button is not shown while the view is
   * read-only. `false` is the same as leaving the option out. mount() rejects
   * with TypeError for a value that cannot be saved to: `true` for a file
   * given as bytes, an address of another scheme, another method.
   */
  saveTo?: SaveTo | false;
  /**
   * With `saveTo`: saves by itself, without the button. Off unless given. `true` saves two
   * seconds after the last change (every change starts the two seconds again) and, for someone
   * who never stops, thirty seconds after the first change at the latest; `{ delayMs, maxWaitMs }`
   * sets the two times (`delayMs` 250 to 600 000, `maxWaitMs` not below it). mount() rejects with
   * TypeError for other values, and for auto-save without `saveTo`.
   *
   * Such a save goes the way the save of the button goes (one request at a time, `If-Match`,
   * the same state and the same errors), with these differences:
   *
   * - What is being typed is neither applied nor cut short: while a cell, a note or the name of
   *   a sheet is being typed, a dialog of rules or the question about a conflict is open, an
   *   input method composes or something is dragged, nothing is sent; it is looked at again one
   *   delay later.
   * - It only goes with a condition: a save to an address is sent by itself only with `If-Match`
   *   (the ETag the server last named for the address the file came from, or a header of the
   *   page's own). Where there is none - the file was given as bytes or is saved to another
   *   address, the server sends no strong ETag - nothing is sent by itself: onNotice says so once
   *   and onSaveState says `autoPaused`; the Save button works as ever, and a save whose answer
   *   brings an ETag sets auto-save going again. `{ unconditional: true }` lifts this: auto-save
   *   then saves without a condition and can silently overwrite what someone else saved. A
   *   function given as `saveTo` guards its own target.
   * - A conflict (412) is told (onNotice, onSaveState, onError) and not asked about, and
   *   auto-save stops: the question opens when the user saves next.
   * - A request that did not get through, one without an answer in time and a 5xx answer are
   *   tried again after 5, 10, 20… seconds (five minutes at most), and told once. Any other
   *   failure stops auto-save until the user saves. A save that succeeds starts it over.
   * - A converted file that may not replace its original (`allowConverted`) and a view that is
   *   read-only are never sent by it.
   * - When the page is hidden (the tab is left or closed) what waits is sent at once.
   * - A save that succeeds is not told to onNotice ("Saved" every few seconds), only to
   *   onSaveState; the first one after something went wrong is.
   *
   * A function given as `saveTo` is told `info.auto: true` for these saves.
   */
  autoSave?: AutoSave;
  /**
   * Edits the workbook together with everybody who opens the same address: the address of a
   * relay ("wss://example.com/collab/42"), or an object (see CollabOptions) with `url` or a
   * `transport` of the application, and `user`, `room`, `token`. An xlsx only: for a document, a
   * file of another format or one that was converted, a notice says that it is not edited
   * together, and the file is shown as without the option. Throws TypeError for anything that
   * is neither an address nor such an object.
   */
  collab?: string | CollabOptions;
  /** With `saveTo`: called whenever there is something else to say about saving. */
  onSaveState?(state: SaveState): void;
  /**
   * With `saveTo`: leaving the page with changes that were not saved asks the
   * user first (the browser's own question). Default true.
   */
  warnOnLeave?: boolean;
  /**
   * Without `saveTo`: called with the file's bytes when the user saves with
   * Ctrl+S or Cmd+S in the view; without either the shortcut is left to the
   * browser. With `saveTo`: called with the bytes after every save that succeeded.
   */
  onSave?(bytes: Uint8Array): void;
  /** Called when saving fails: with a SaveError when `saveTo` is given, otherwise with what the writer threw. */
  onError?(error: unknown): void;
  /**
   * Opens the formats that cannot be written back (doc, RTF, xls) for
   * editing by converting them first: doc and RTF to docx, xls to xlsx. What
   * is mounted, edited and saved is then the converted file, and the result
   * says so in `conversion`; onNotice is told as well. Without it, and with
   * `readOnly`, these files are shown as they are, read-only, and nothing
   * turns editing on. A file whose conversion fails is shown the same way,
   * and onNotice says so. Has no effect on docx and xlsx.
   */
  convert?: boolean;
  /**
   * Fonts to register before the file is shown: see registerFont(). They are
   * waited for, for three seconds at most, so that the file is laid out with
   * them from the start. A font that cannot be loaded does not keep the file
   * from being shown: onNotice is told, and so it is of a font that is still
   * loading when the wait ends (the view lays its text out again when it
   * arrives). The fonts stay registered when the view is destroyed.
   *
   * `false` turns the font the package brings off, as `fontsUrl: null` does.
   */
  fonts?: FontRegistration[] | false;
  /**
   * Where the fonts of the package are: the address of its `dist/fonts/`
   * folder, as fonts.configure({ baseUrl }) takes it. Needed where the
   * engine is inside a bundle (the folder is then not next to the script).
   * `null` turns the embedded font off. Like the fonts themselves the
   * setting is the page's, not this view's: it stays when the view is destroyed.
   */
  fontsUrl?: string | null;
  /** Documents: the zoom, 1 = 100%, or a mode that follows the container. Default: 100%, less when the page is wider than the container. */
  zoom?: DocumentZoom;
  /** Documents: called when the page in view or the number of pages changes. `page` counts from 1. */
  onPageChange?(page: number, pageCount: number): void;
  /** Documents: called when a hyperlink is activated; return false to keep the view from opening it in a new window. */
  onLinkClick?(target: string): boolean | void;
  /** Documents: called after every change made in the editor, undo and redo included. */
  onDocumentChange?(): void;
  /**
   * Documents: tracks the changes made in the editor although the document
   * does not ask for it. A document that asks for it (its "track changes"
   * setting, or a protection that allows only tracked changes) is always
   * edited with tracked changes; false does not turn that off.
   */
  trackChanges?: boolean;
  /**
   * Who is editing: the author that tracked changes of a document name, and the author of
   * the notes added to the cells of a spreadsheet. Default "Author".
   */
  author?: string;
  /** Documents: the clock tracked changes take their date from (tests). Default: the system's. */
  now?(): Date;
}

/** What the protection of a mounted document enforces. */
export interface MountProtection {
  /**
   * "readOnly" or "forms": the document is shown read-only and nothing
   * turns editing on (form fields cannot be filled in here either).
   * "comments": the text is read-only, and comments can be added, edited,
   * answered, resolved and removed (`view.canComment`); the file is saved
   * with them. "trackedChanges": the document can be edited, and every
   * change is tracked.
   */
  edit: 'readOnly' | 'comments' | 'forms' | 'trackedChanges';
  /**
   * The document marks ranges that its protection leaves open for editing
   * (w:permStart). They are not honoured: the whole document is read-only.
   */
  exceptions: boolean;
}

/** A file that mount() converted to open it for editing. */
export interface MountConversion {
  /** The format of the file that was given. */
  from: 'doc' | 'rtf' | 'xls';
  /** The format of the file that is mounted and that save() gives. */
  to: 'docx' | 'xlsx';
  /**
   * What the model held and the new file does not carry over, each with its
   * count, in English: for a document, shapes written as rectangles,
   * groups, pictures without data, characters XML cannot hold, links that
   * are not followed and theme references (the last two only for models
   * that hold them, which those read from doc and RTF do not); for a
   * workbook, what buildXlsx() reports (chart and macro sheets, pictures it
   * cannot read, ...). Empty when nothing was lost. What the reader never
   * put into the model is not listed.
   */
  warnings: string[];
}

export interface MountedSpreadsheet {
  kind: 'spreadsheet';
  /** The file that is shown: the converted one when `conversion` is set. */
  file: OfficeFile;
  /** Set when the file was converted to be opened for editing (see MountOptions.convert). */
  conversion?: MountConversion;
  workbook: Workbook;
  /** The grid: selection, editing, undo and redo, formatting. */
  view: WorkbookView;
  /** The editor around the grid: toolbar, formula bar and status bar. */
  editor: SpreadsheetEditor;
  /**
   * Set when the workbook is edited together with others (MountOptions.collab). When the file
   * was changed elsewhere it is loaded again in place: `file`, `workbook`, `view` and `editor`
   * are then other objects, on this same result.
   */
  collab?: MountedCollab;
  /**
   * The file with every change made in the view, as bytes. A cell that is
   * being edited is written first, as leaving it would do. A file that was
   * opened with a password is encrypted with it again (see SaveOptions); once
   * the view is destroyed the password is forgotten, and such a save rejects.
   */
  save(options?: SaveOptions): Promise<Uint8Array>;
  /**
   * Prints the sheet that is shown, or every visible sheet, through the
   * print dialog of the browser (which can also save a PDF): see
   * WorkbookView.print(). Resolves when the dialog was closed.
   */
  print(options?: SheetPrintOptions): Promise<DocumentPrintResult>;
  /**
   * The sheet that is shown, or the sheets `sheets` names, as a PDF file:
   * the pages print() makes, each one a picture with its text lying
   * invisible over it (the file can be searched and its text copied). What
   * is being typed is applied first. See WorkbookView.exportPdf().
   */
  exportPdf(options?: SheetPdfExportOptions): Promise<Uint8Array>;
  /**
   * Saves the file to MountOptions.saveTo, as the Save button does: what is
   * being typed is applied first. Resolves to true when it was saved, also
   * when there was nothing to save; to false when it could not be (onNotice
   * and onSaveState say why), and always to false without `saveTo`. Never rejects.
   * `{ overwrite: true }` saves over whatever the server holds (no `If-Match`):
   * the user's answer to a SaveError "conflict". While the editor asks the user
   * about a conflict a call sends nothing, and resolves as the question ends.
   */
  saveNow(options?: SaveNowOptions): Promise<boolean>;
  /**
   * Whether there is work that saveNow() has not saved. With `saveTo`: changes that no save holds,
   * or text that is typed and not applied yet (a cell, a note or the name of a sheet that is being
   * typed, a dialog of rules with changes) - what leaving the page now would lose, and what the
   * question before leaving goes by. The second is no change of the file yet: onSaveState's `dirty`
   * does not count it, and a save (not auto-save) applies it first.
   */
  readonly hasUnsavedChanges: boolean;
  /**
   * Turns auto-save on, off or to other times while the view is shown (see MountOptions.autoSave).
   * Turned on over unsaved changes, it saves them one delay later; it does not start again what a
   * conflict or a refusal stopped (the user's next save does). Throws TypeError for a value that
   * cannot be, and for turning it on without `saveTo`.
   */
  setAutoSave(value: AutoSave): void;
  /** Removes the view from the container; the same as `editor.destroy()`. */
  destroy(): void;
}

export interface MountedDocument {
  kind: 'document';
  /** The file that is shown: the converted one when `conversion` is set. */
  file: OfficeFile;
  /** Set when the file was converted to be opened for editing (see MountOptions.convert). */
  conversion?: MountConversion;
  /** Set when the document enforces a protection; onNotice is told once when that keeps it read-only. */
  protection?: MountProtection;
  /** Set (true) when the changes made in the editor are tracked: the document asks for it, or MountOptions.trackChanges did. */
  trackChanges?: boolean;
  document: DocModel;
  /** The pages: zoom, navigation, the caret and the selection, editing, undo and redo. */
  view: DocumentView;
  /** The editor around the pages: the formatting toolbar. */
  editor: DocumentEditor;
  /**
   * The file as bytes: a docx with the changes made in the editor; doc and RTF, which are read-only, as the bytes they
   * were opened from. A file that was opened with a password is encrypted with it again (see SaveOptions); once the
   * view is destroyed the password is forgotten, and such a save rejects.
   */
  save(options?: SaveOptions): Promise<Uint8Array>;
  /**
   * Prints the document through the print dialog of the browser (which can
   * also save it as a PDF): see DocumentView.print(). Resolves when the dialog was closed.
   */
  print(options?: DocumentPrintOptions): Promise<DocumentPrintResult>;
  /**
   * The document as a PDF file: every page a picture of what is shown, with
   * its text lying invisible over it (the file can be searched and its text
   * copied). Also for doc and RTF files that are shown read-only. See
   * DocumentView.exportPdf().
   */
  exportPdf(options?: PdfExportOptions): Promise<Uint8Array>;
  /**
   * Saves the file to MountOptions.saveTo, as the Save button does: what is
   * being typed is applied first. Resolves to true when it was saved, also
   * when there was nothing to save; to false when it could not be (onNotice
   * and onSaveState say why), and always to false without `saveTo`. Never rejects.
   * `{ overwrite: true }` saves over whatever the server holds (no `If-Match`):
   * the user's answer to a SaveError "conflict". While the editor asks the user
   * about a conflict a call sends nothing, and resolves as the question ends.
   */
  saveNow(options?: SaveNowOptions): Promise<boolean>;
  /** Whether changes were made that saveNow() has not saved: with `saveTo`, what the Save button shows. */
  readonly hasUnsavedChanges: boolean;
  /**
   * Turns auto-save on, off or to other times while the view is shown (see MountOptions.autoSave).
   * Turned on over unsaved changes, it saves them one delay later; it does not start again what a
   * conflict or a refusal stopped (the user's next save does). Throws TypeError for a value that
   * cannot be, and for turning it on without `saveTo`.
   */
  setAutoSave(value: AutoSave): void;
  /** Removes the view from the container; the same as `editor.destroy()`. */
  destroy(): void;
}

/** What mount() resolves to; `kind` tells which of the two it is. */
export type MountedView = MountedSpreadsheet | MountedDocument;

/** How long mount() waits for its fonts before it shows the file without them. */
const FONT_WAIT_MS = 3000;

/**
 * Opens a file and shows it inside `container`, which must have a size: a
 * spreadsheet (xlsx) in the editor, with a toolbar, a formula bar and a status
 * bar unless the options turn them off; a document as pages, a docx in the
 * document editor (unless `readOnly` is set), doc and RTF read-only.
 * Other formats reject with UnsupportedFileError.
 */
export async function mount(container: HTMLElement, options: MountOptions): Promise<MountedView> {
  // A file that was opened and is not shown after all (its conversion or its view failed) lets go of its password.
  const opened: OfficeFile[] = [];
  try {
    return await mountFile(container, options, opened);
  } catch (error) {
    for (let i = 0; i < opened.length; i++) opened[i].close();
    throw error;
  }
}

async function mountFile(container: HTMLElement, options: MountOptions, held: OfficeFile[]): Promise<MountedView> {
  // Before anything is measured: the first text in a family the embedded font stands in for is what fetches it.
  if (options.fonts === false || options.fontsUrl === null) configureEmbeddedFonts({ baseUrl: null });
  else if (options.fontsUrl !== undefined) configureEmbeddedFonts({ baseUrl: options.fontsUrl });
  // The fonts load while the file is fetched and read. This never rejects: a font is no reason not to show a file.
  const fontsSettled: Promise<FontFailure[]> = Array.isArray(options.fonts) ? settleFonts(options.fonts, FONT_WAIT_MS) : Promise.resolve([]);
  let input = options.file;
  // Where the page is: what a relative address is resolved against.
  const base = typeof document !== 'undefined' && typeof document.baseURI === 'string' ? document.baseURI : undefined;
  // Before anything is fetched: a `saveTo` that cannot be saved to is a mistake of the page, and is told at once.
  const saveTarget = options.saveTo === undefined || options.saveTo === false || options.saveTo === null ? undefined : resolveSaveTarget(options.saveTo, typeof input === 'string' ? input : undefined, base);
  // The same for auto-save: its times are read, and it has nowhere to save to without `saveTo`.
  // And for editing together: an address that is none is told before anything is fetched.
  const collabPlan = resolveCollab(options.collab);
  const autoSave = resolveAutoSave(options.autoSave);
  if (autoSave && !saveTarget) throw new TypeError('"autoSave" needs "saveTo": there is nowhere to save to');
  /** The ETag the file came with: what a save to the same address is sent with. */
  let etag: string | null = null;
  let readETag: string | null = null;
  let fileName: string | undefined;
  if (typeof input === 'string') {
    fileName = nameFromUrl(input, base);
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
    // Not the ETag of an answer that came from another address than the one that is saved to.
    etag = fileETag(response);
    // For those who edit together the file is also known by the ETag of an answer that was redirected.
    readETag = response.headers.get('ETag');
    input = await response.arrayBuffer();
  } else if (typeof Blob !== 'undefined' && input instanceof Blob && typeof (input as { name?: unknown }).name === 'string') {
    fileName = (input as unknown as { name: string }).name || undefined;
  }
  /**
   * The bytes of the file as it came: what it is known by when it is edited together. For a password-protected
   * file these are its encrypted bytes, the ones everybody in the room fetched: a file that was encrypted again
   * is another sequence of bytes every time (new salts), so a version never stands for "this content", only
   * for "these bytes at the address", which is what the versions of a save name as well (see sayShared()).
   */
  let readBytes: Uint8Array | undefined;
  if (collabPlan) {
    readBytes = input instanceof Uint8Array ? input : input instanceof ArrayBuffer ? new Uint8Array(input) : new Uint8Array(await (input as Blob).arrayBuffer());
    input = readBytes;
  }
  let file = await openForView(input, options, container);
  held.push(file);
  /** What the file that was given has to say; a converted file is another object. */
  const fileWarnings = file.warnings || [];
  let conversion: MountConversion | undefined;
  /** Why a file that had to be converted was not: it is then shown as it is, read-only. */
  let notConverted: string | undefined;
  if (options.convert === true && options.readOnly !== true && needsConversion(file.format)) {
    if (file.kind === 'document' && !file.document) throw new UnsupportedFileError('Showing doc documents is not supported yet');
    // The model is written out as a new file of the format that can be saved, and that file is what is opened.
    const from = file.format;
    const opened = await openConverted(file, (bytes: Uint8Array) => open(bytes));
    if (opened.conversion) {
      // The converted file took the password of its original over.
      file.close();
      file = opened.file;
      held.push(file);
      conversion = { from: opened.conversion.from, to: file.format as 'docx' | 'xlsx', warnings: opened.conversion.warnings };
    } else if (opened.failed !== undefined) notConverted = from;
  }
  const fontFailures = await fontsSettled;
  const converted = (): void => {
    said();
    // What the file has to say of itself is said through the view, and last: it is heard (in the live region of the
    // view) also where the page does not listen, and it is what stays there.
    for (let i = 0; i < fileWarnings.length; i++) {
      if (fileWarnings[i].code === 'no-integrity') notice(resolveUiStrings(options.locale).noIntegrityCheck, 'error');
    }
  };
  const said = (): void => {
    if (!options.onNotice) return;
    // In the language the views speak for the same option.
    const strings = resolveUiStrings(options.locale);
    for (let i = 0; i < fontFailures.length; i++) {
      options.onNotice({ message: strings.fontNotLoaded(fontFailures[i].family, fontFailures[i].pending), kind: fontFailures[i].pending ? 'info' : 'error' });
    }
    // Last of what is said when the file is shown is what matters most: the view shows the notice that came last.
    if (conversion) options.onNotice({ message: strings.openedConverted(conversion.from, conversion.to), kind: 'info' });
    else if (notConverted !== undefined) options.onNotice({ message: strings.notConverted(notConverted), kind: 'error' });
    // Said before anything is typed: what is changed in this file has nowhere to go (see MountOptions.saveTo).
    if (conversion && controller && controller.blockedConversion) options.onNotice({ message: strings.saveConvertedNotAllowed(conversion.to), kind: 'error' });
  };
  let save: (saveOptions?: SaveOptions) => Promise<Uint8Array>;
  // What the two kinds of view answer differently; set once the view is there.
  let commit: () => boolean = () => true;
  /** Whether the user is in the middle of something, and the file as the model is: what auto-save asks and takes. */
  let pendingInput: () => boolean = () => false;
  let saveAsIs: () => Promise<Uint8Array> = () => save();
  let locked: () => boolean = () => false;
  let notice: (message: string, kind: 'error' | 'info') => void = () => undefined;
  let focusView: () => void = () => undefined;
  /** The root element of the editor, and whether its view is only read: what the question about a conflict needs. */
  let editorRoot: (() => HTMLElement) | undefined;
  let readOnlyNow: () => boolean = () => false;
  const info = () => saveInfoOf(file.format, conversion, fileName);
  const controller = saveTarget
    ? new SaveController({
        target: saveTarget,
        writable: () => (file.format === 'docx' || file.format === 'xlsx') && !locked(),
        commit: () => commit(),
        bytes: () => save(),
        bytesAsIs: () => saveAsIs(),
        pending: () => pendingInput(),
        readOnly: () => readOnlyNow(),
        // Edited together: one participant saves for all, and a conflict with a save of someone in the room is none.
        autoAllowed: () => (saveSync ? saveSync.autoAllowed() : true),
        autoWaiting: () => (saveSync ? saveSync.waiting() : undefined),
        conflictWait: () => (saveSync ? saveSync.conflictWait() : Promise.resolve(false)),
        autoSave,
        info,
        strings: () => resolveUiStrings(options.locale),
        notice: (message, kind) => notice(message, kind),
        onState: (state) => {
          // What the bar of the room says of auto-save follows.
          showRoom();
          if (options.onSaveState) options.onSaveState(state);
        },
        onSaved: (bytes) => {
          if (saveSync) saveSync.ownSaved();
          sayShared(bytes);
          if (options.onSave) options.onSave(bytes);
        },
        onError: options.onError,
        // A save the server refused with 412: the user is asked over the editor, unless the page said
        // `onConflict: 'fail'` (the controller looks at that). A function of the application knows its own conflicts.
        askOverwrite:
          saveTarget.kind === 'request' && typeof document !== 'undefined'
            ? (question) => {
                // A view that is only read is not asked to overwrite anything: the conflict was told, and that is all.
                if (readOnlyNow() || !editorRoot) return question.answer(false);
                const words = resolveUiStrings(options.locale);
                // A workbook that was kept apart from its room meets a file someone else saved: a fork, and three ways out of it.
                const apart = session;
                if (apart && apart.state === 'detached' && typeof options.file === 'string') {
                  const fork = words.collab.fork;
                  return askOverwrite({
                    root: editorRoot(),
                    kind: 'collab-fork',
                    strings: { saveConflictTitle: fork.title, saveConflictQuestion: fork.question, saveOverwrite: fork.save, saveCancel: words.saveCancel },
                    detail: question.detail,
                    answer: question.answer,
                    focusView: () => focusView(),
                    also: [
                      {
                        name: 'download',
                        label: fork.download,
                        stays: true,
                        // The workbook as it is here, as a file of the user's: nothing is lost by what is chosen next.
                        run: () => {
                          void saveAsIs().then(
                            (bytes) => {
                              if (downloadFile(document, bytes, options.fileName !== undefined ? options.fileName : fileName || 'workbook.xlsx')) notice(fork.downloaded, 'info');
                            },
                            () => undefined,
                          );
                        },
                      },
                      {
                        name: 'load',
                        label: fork.load,
                        destructive: true,
                        run: () => {
                          // The save is given up, and the workbook with it: the stored file is loaded, and the room joined with it.
                          question.answer(false);
                          apart.reload();
                        },
                      },
                    ],
                  });
                }
                return askOverwrite({
                  root: editorRoot(),
                  strings: words,
                  detail: question.detail,
                  answer: question.answer,
                  focusView: () => focusView(),
                });
              }
            : undefined,
        etag,
      })
    : undefined;
  /** The session, while the workbook is edited together; and the file that was written last, with the step it holds. */
  let session: CollabSession | undefined;
  let saveSync: SaveSync | undefined;
  let showRoom: () => void = () => undefined;
  let unstampedSaid = false;
  let written: { bytes: Uint8Array; rev: number | undefined } | undefined;
  /** Tells the others of a save that succeeded, when the bytes are the ones that were written last. */
  const sayShared = (bytes: Uint8Array, tag?: string): void => {
    if (!session || !written || written.bytes !== bytes) return;
    const versions = fileVersions(tag !== undefined ? tag : controller ? controller.currentETag : undefined, bytes);
    // A workbook that was detached from the room is in a file now: with that file it joins again.
    if (session.state === 'detached') session.rejoin(versions);
    else session.announceSaved(versions, written.rev);
  };
  /**
   * Writes the file while the workbook is edited together: once what was changed here has its
   * place in the common order (or after a short wait), and with what arrives held back until
   * the file is written, so that it is the workbook of one moment.
   */
  const sharedBytes = async (take: () => Promise<Uint8Array>): Promise<Uint8Array> => {
    const shared = session;
    if (!shared) return take();
    const done = await writeSharedFile(shared, file.package, take);
    // The file could not be made to say which step it holds: saved all the same, and said once.
    if (done.unstamped && !unstampedSaid) {
      unstampedSaid = true;
      notice(resolveUiStrings(options.locale).collab.noStamp, 'info');
    }
    // A closed session tells nobody of its saves.
    if (shared.state !== 'closed') written = done;
    return done.bytes;
  };
  /** What ties the controller to the page: the question before leaving, the hidden page, the view that turns read-only. */
  let bound: { unsaved(): boolean; release(): void } | undefined;
  /** `unapplied`: whether text is typed in the view that is in no change yet (see SavePage). */
  const watchReadOnly = (view: { setReadOnly(readOnly: boolean): void }, unapplied: () => boolean): void => {
    if (!controller) return;
    bound = bindSaveToPage(controller, {
      view,
      unapplied,
      window: typeof window !== 'undefined' ? window : undefined,
      document: typeof document !== 'undefined' ? document : undefined,
      warnOnLeave: options.warnOnLeave !== false,
    });
  };
  /** Without `saveTo`: whether anything was changed at all. */
  let touched = false;
  const changed = () => {
    touched = true;
    if (controller) controller.changed();
    // In a room this side's work waits for the one who saves for all, and not for ever.
    if (saveSync) saveSync.localChanged();
  };
  const saveNow = (saveOptions?: SaveNowOptions): Promise<boolean> => (controller ? controller.saveNow(saveOptions) : Promise.resolve(false));
  // With `saveTo`: also text that is typed and in no change yet, which a user who leaves would lose.
  const unsaved = (): boolean => (bound ? bound.unsaved() : controller ? controller.dirty : touched);
  const setAutoSave = (value: AutoSave): void => {
    if (controller) return controller.setAutoSave(value);
    if (resolveAutoSave(value)) throw new TypeError('"autoSave" needs "saveTo": there is nowhere to save to');
  };
  /**
   * Ties what mount() holds to the life of the editor: its destroy() (which is also the result's) takes the key
   * listener and the question before leaving away and gives up a save that is under way, whoever calls it.
   */
  const tie = (editor: { destroy(): void }): (() => void) => {
    // While the event goes down, before any field inside the editor can keep it to itself (a dialog of
    // rules, the field of a note or of a link do): otherwise the browser would open its "save page" there.
    container.addEventListener('keydown', onKeyDown, true);
    const destroyEditor = editor.destroy.bind(editor);
    let over = false;
    editor.destroy = () => {
      if (!over) {
        over = true;
        container.removeEventListener('keydown', onKeyDown, true);
        if (bound) bound.release();
        if (saveSync) saveSync.stop();
        if (controller) controller.destroy();
        // The password of a protected file is not kept longer than its view.
        file.close();
      }
      destroyEditor();
    };
    // For a file that is loaded again in place: the editor goes, what mount() holds stays.
    detach = () => {
      over = true;
      container.removeEventListener('keydown', onKeyDown, true);
      destroyEditor();
    };
    return () => editor.destroy();
  };
  let detach: () => void = () => undefined;
  const onKeyDown = (event: KeyboardEvent) => {
    if ((!controller && !options.onSave) || !(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey || (event.key !== 's' && event.key !== 'S')) return;
    event.preventDefault();
    if (controller) return void controller.saveNow();
    save().then(options.onSave, (error) => {
      if (options.onError) options.onError(error);
      else throw error;
    });
  };

  if (file.kind === 'document') {
    if (!file.document) throw new UnsupportedFileError('Showing doc documents is not supported yet');
    const model = await file.document();
    const editor = new DocumentEditor(container, model, {
      package: documentPictures(model, file.package),
      locale: options.locale,
      zoom: options.zoom,
      toolbar: options.toolbar,
      // Only a docx is written back with its changes; the other formats are shown read-only until they can be saved.
      readOnly: file.format !== 'docx' || options.readOnly === true,
      partTaken: file.package ? (part: string) => (file.package as OpcPackage).hasPart(part) : undefined,
      onPageChange: options.onPageChange,
      onNotice: options.onNotice,
      onLinkClick: options.onLinkClick,
      onChange: () => {
        changed();
        if (options.onDocumentChange) options.onDocumentChange();
      },
      saveCommand: controller,
      trackChanges: options.trackChanges,
      author: options.author,
      now: options.now,
      // The name the PDF is downloaded under: the one the page states, or the one the file came with.
      fileName: options.fileName !== undefined ? options.fileName : fileName,
    });
    const view = editor.view;
    if (file.format !== 'docx') {
      // Nothing may turn editing on: saving gives the original bytes and the edits would be lost.
      const lockView = view.setReadOnly.bind(view);
      const lockEditor = editor.setReadOnly.bind(editor);
      view.setReadOnly = () => lockView(true);
      editor.setReadOnly = () => lockEditor(true);
    }
    save = async (saveOptions?: SaveOptions) => {
      // A composition still under way belongs to the document that is written.
      view.commitEditing();
      const changes = view.changeCount;
      const bytes = await file.save(saveOptions);
      // A change made while the file was written is not in these bytes.
      if (view.changeCount === changes) view.markSaved();
      return bytes;
    };
    commit = () => {
      view.commitEditing();
      return true;
    };
    pendingInput = () => view.hasPendingInput();
    saveAsIs = async () => {
      // As save(), without ending anything: auto-save only runs while nothing is being composed.
      const changes = view.changeCount;
      const bytes = await file.save();
      if (view.changeCount === changes) view.markSaved();
      return bytes;
    };
    // A protection that allows no editing: nothing is written back. One that allows comments is written back with them.
    locked = () => view.protection !== 'comments' && locksSaving(view.protection);
    notice = (message, kind) => view.showNotice(message, kind);
    focusView = () => view.focus();
    editorRoot = () => editor.element;
    // A view that takes comments (and nothing else) has changes to save, although its text is read-only.
    readOnlyNow = () => view.readOnly && !view.canComment;
    // What is typed in a document is in the document: nothing waits in a field to be applied.
    watchReadOnly(view, () => false);
    const destroy = tie(editor);
    converted();
    // Said, not thrown: the document is shown and edited as without the option.
    if (collabPlan) view.showNotice(resolveUiStrings(options.locale).collab.document, 'error');
    const print = (printOptions?: DocumentPrintOptions) => view.print(printOptions);
    const exportPdf = (pdfOptions?: PdfExportOptions) => view.exportPdf(pdfOptions);
    const mounted: MountedDocument = {
      kind: 'document',
      file,
      document: model,
      view,
      editor,
      save,
      print,
      exportPdf,
      saveNow,
      get hasUnsavedChanges() {
        return unsaved();
      },
      setAutoSave,
      destroy,
    };
    if (conversion) mounted.conversion = conversion;
    if (view.trackingChanges) mounted.trackChanges = true;
    const enforced = view.protection;
    if (enforced) mounted.protection = { edit: enforced, exceptions: hasPermissionRanges(model) };
    return mounted;
  }

  let workbook = await file.workbook();
  // An xls file cannot be written yet: it is shown read-only, so that no edit is lost on saving.
  const sheetOptions: SpreadsheetEditorOptions = {
    ...options,
    onChange: (change) => {
      // What someone else changed is there to be saved: by the one participant who saves for all.
      if (change.remote && controller) {
        touched = true;
        if (saveSync) saveSync.remoteChanged();
        else controller.changed(true);
      } else changed();
      if (options.onChange) options.onChange(change);
    },
    saveCommand: controller,
    fileName: options.fileName !== undefined ? options.fileName : fileName,
  };
  // The editor has no use for the password, and does not keep it.
  delete (sheetOptions as { password?: unknown }).password;
  delete (sheetOptions as { onPassword?: unknown }).onPassword;
  if (file.format === 'xls') sheetOptions.readOnly = true;
  let editor = new SpreadsheetEditor(container, workbook, sheetOptions);
  let view = editor.view;
  if (file.format === 'xls') {
    // Nothing may turn editing on: saving would give the original bytes and the edits would be lost.
    const lockView = view.setReadOnly.bind(view);
    const lockEditor = editor.setReadOnly.bind(editor);
    view.setReadOnly = () => lockView(true);
    editor.setReadOnly = () => lockEditor(true);
  }
  save = async (saveOptions?: SaveOptions) => {
    // What is being typed (in a cell, in the field of a note, in a tab) is written first, as leaving it would do.
    if (!view.commitPending()) throw new Error('Cannot save: a note or a sheet name that is being typed cannot be applied, or a dialog of rules is open with changes (confirm or close it first)');
    return sharedBytes(() => file.save(saveOptions));
  };
  commit = () => {
    try {
      return view.commitPending();
    } catch (error) {
      // A cell whose text cannot be written: the editor stays open, and the view says why.
      view.showNotice(view.localizeError(error), 'error');
      return false;
    }
  };
  pendingInput = () => view.hasPendingInput();
  saveAsIs = () => sharedBytes(() => file.save());
  notice = (message, kind) => view.showNotice(message, kind);
  focusView = () => view.focus();
  editorRoot = () => editor.element;
  readOnlyNow = () => view.readOnly;
  watchReadOnly(view, () => view.hasUnappliedText());
  const destroyTied = tie(editor);
  converted();
  const print = (printOptions?: SheetPrintOptions) => view.print(printOptions);
  const exportPdf = (pdfOptions?: SheetPdfExportOptions) => view.exportPdf(pdfOptions);
  const mounted: MountedSpreadsheet = {
    kind: 'spreadsheet',
    file,
    workbook,
    view,
    editor,
    save: (saveOptions?: SaveOptions) => save(saveOptions),
    print,
    exportPdf,
    saveNow,
    get hasUnsavedChanges() {
      return unsaved();
    },
    setAutoSave,
    destroy: () => {
      if (session) session.leave();
      mounted.editor.destroy();
    },
  };
  void destroyTied;

  if (collabPlan && (file.format !== 'xlsx' || conversion)) {
    // Everybody in a room has to read the very same file: one that was converted here is not it.
    view.showNotice(resolveUiStrings(options.locale).collab.unsupported, 'error');
  } else if (collabPlan && readBytes) {
    const plan = collabPlan;
    let peopleTold = '';
    const peopleNow = (): People => (session ? listPeople(session.user, session.peers, session.state, resolveUiStrings(options.locale).collab.guest) : { list: [], more: 0, count: 0 });
    const show = (): void => {
      if (!session) return;
      const people = peopleNow();
      mounted.editor.setCollab({ state: session.state, peers: session.peers, pending: session.pending, autoWaits: controller ? controller.state.autoWaits : undefined, people });
      // Whoever listens hears of the list when it is another one.
      const key = JSON.stringify(people.list);
      if (key === peopleTold) return;
      peopleTold = key;
      if (plan.options.onPeople) plan.options.onPeople(people.list);
      const out = mounted.collab;
      if (out && out.onPeople) out.onPeople(people.list);
    };
    showRoom = show;
    /** The file is loaded again from its address and shown in place (the room is at a version this side does not have). */
    const reload = async (): Promise<{ versions: string[]; stamp?: { ep: string; rev: number } }> => {
      if (typeof options.file !== 'string') throw new Error('The file has no address to load it from again');
      // Until the file is there nothing is typed into the workbook that is about to go.
      const editable = !view.readOnly;
      // What is being typed can be applied to nothing (the workbook goes): it is said back when the file is there.
      const typed = editable ? view.takeTypedText() : undefined;
      if (editable) editor.setReadOnly(true);
      view.showNotice(resolveUiStrings(options.locale).collab.updating, 'info');
      try {
        const loaded = await loadAgain();
        if (typed !== undefined) view.showNotice(resolveUiStrings(options.locale).collab.typedKept(typed.length > 500 ? `${typed.slice(0, 500)}…` : typed), 'error');
        return loaded;
      } catch (error) {
        if (editable) mounted.editor.setReadOnly(false);
        throw error;
      }
    };
    const loadAgain = async (): Promise<{ versions: string[]; stamp?: { ep: string; rev: number } }> => {
      const address = options.file as string;
      const response = await fetch(address, { cache: 'no-store' });
      if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
      const tag = fileETag(response);
      const named = response.headers.get('ETag');
      const bytes = new Uint8Array(await response.arrayBuffer());
      // A password-protected file: its next version is opened with the password this one is kept with.
      const next = await open(bytes, { passwordOf: file, requireIntegrity: options.requireIntegrity });
      if (next.format !== 'xlsx') {
        next.close();
        throw new Error('The file is no longer a workbook that can be edited together');
      }
      const nextBook = await next.workbook();
      const shownSheet = view.activeSheet;
      if (bound) bound.release();
      detach();
      // The file that was shown lets go of its password; the new one holds its own.
      file.close();
      file = next;
      workbook = nextBook;
      editor = new SpreadsheetEditor(container, workbook, { ...sheetOptions, sheet: shownSheet < workbook.sheets.length ? shownSheet : undefined });
      view = editor.view;
      watchReadOnly(view, () => view.hasUnappliedText());
      tie(editor);
      mounted.file = file;
      mounted.workbook = workbook;
      mounted.view = view;
      mounted.editor = editor;
      view.attachCollab(session);
      // The file as it is on the server is what is shown: there is nothing to save, and the next save names it.
      if (controller) controller.adopt(tag === null ? undefined : tag, true);
      touched = false;
      show();
      return { versions: fileVersions(named, bytes), stamp: file.package ? await readCollabStamp(file.package) : undefined };
    };
    const handle: MountedCollab = {
      get state() {
        return session ? session.state : 'closed';
      },
      get rev() {
        return session ? session.rev : 0;
      },
      get pending() {
        return session ? session.pending : 0;
      },
      get peers() {
        return session ? session.peers : [];
      },
      get people() {
        return peopleNow().list;
      },
      get user() {
        return (session as CollabSession).user;
      },
      saved: (bytes, tag) => sayShared(bytes, typeof tag === 'string' ? tag : undefined),
      reload: () => {
        if (session) session.reload();
      },
      leave: () => {
        if (session) session.leave();
      },
    };
    saveSync = controller
      ? createSaveSync({
          controller,
          session: () => session,
          target: saveTarget && saveTarget.kind === 'request' ? { url: saveTarget.url, credentials: saveTarget.credentials, headers: saveTarget.headers } : undefined,
          people: () => (session ? session.peers.length + 1 : 1),
          strong: strongETag,
          readOnly: () => mounted.view.readOnly,
          // What this workbook would be written as now: with it a save of someone else is proven to hold the work done here.
          watchMs: typeof plan.options.saveWatchMs === 'number' && plan.options.saveWatchMs >= 1000 ? plan.options.saveWatchMs : undefined,
          ownBytes: async () => {
            if (!session) return { bytes: new Uint8Array(0), rev: undefined };
            // A file that is written to be looked at leaves the package as it was: what it says of the room is put back.
            const restore = file.package ? await keepStamp(file.package) : () => undefined;
            try {
              return await writeSharedFile(session, file.package, () => file.save());
            } finally {
              restore();
            }
          },
          // Part by part, unpacked: two browsers do not compress alike.
          sameFile: samePackage,
        })
      : undefined;
    session = new CollabSession(
      {
        workbook: () => workbook,
        changed: (change) => view.collabChanged(change, change.local !== true),
        notice: (said) => view.collabNotice(said),
        reload: typeof options.file === 'string' ? reload : undefined,
        // With `saveTo`: changes that no save holds. Without it: anything changed at all.
        unsaved: () => (controller ? controller.dirty : touched),
      },
      {
        transport: plan.transport(),
        versions: fileVersions(readETag, readBytes),
        stamp: file.package ? await readCollabStamp(file.package) : undefined,
        user: plan.options.user,
        room: plan.options.room,
        token: plan.options.token,
        readOnly: view.readOnly,
        onState: (state) => {
          if (saveSync) saveSync.roomChanged();
          show();
          if (plan.options.onState) plan.options.onState(state);
          if (handle.onState) handle.onState(state);
        },
        onPeers: (peers) => {
          if (saveSync) saveSync.roomChanged();
          view.collabPeersChanged(peers);
          show();
          if (plan.options.onPeers) plan.options.onPeers(peers);
          if (handle.onPeers) handle.onPeers(peers);
        },
        onSaved: (saved) => void (saveSync ? saveSync.saved(saved) : undefined),
        onReadOnly: () => mounted.editor.setReadOnly(true),
        // Someone cannot join because no file holds the steps of the room: the one participant who saves for the others does.
        onSync: () => void (saveSync ? saveSync.sync() : undefined),
      },
    );
    view.attachCollab(session);
    mounted.collab = handle;
    show();
  }
  if (conversion) mounted.conversion = conversion;
  return mounted;
}
