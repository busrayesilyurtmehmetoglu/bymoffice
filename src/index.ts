import { CfbFile } from './cfb';
import { readDoc } from './formats/doc/document';
import { readRtf } from './formats/rtf/document';
import { readDocx } from './formats/docx/document';
import { buildDocx, writeDocx } from './formats/docx/write';
import { needsConversion, openConverted } from './formats/convert';
import { documentPictures } from './model/doc/pictures';
import { hasPermissionRanges, type DocModel } from './model/doc/types';
import { readXls } from './formats/xls/workbook';
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
import { OpcPackage } from './opc';
import { askOverwrite } from './ui/save-confirm';
import {
  SaveController,
  bindSaveToPage,
  fileETag,
  locksSaving,
  nameFromUrl,
  resolveAutoSave,
  resolveSaveTarget,
  saveInfoOf,
  type AutoSave,
  type SaveNowOptions,
  type SaveState,
  type SaveTo,
} from './embed/save-to';

export { CfbFile, type CfbEntry } from './cfb';
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
export { readDocx } from './formats/docx/document';
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
export { SheetDrawings, adjustCommentCell, adjustDrawingAnchor, commentCellNow, commentEdits, anchorRect, type DrawingTarget, type PlacedComment, type SheetDrawingLabels, type SheetDrawingsOptions } from './render/sheet-drawings';
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
export const fonts = { registerFont, registeredFonts, unregisterFont, configure: configureEmbeddedFonts, embedded: embeddedFontsState };

export type OfficeFormat = 'docx' | 'xlsx' | 'doc' | 'xls' | 'rtf';

export interface OfficeFile {
  format: OfficeFormat;
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb' | 'rtf';
  /** Present for docx/xlsx. */
  package?: OpcPackage;
  /** Present for doc/xls. */
  cfb?: CfbFile;
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
   * lose their geometry, theme colours become plain ones).
   */
  saveAsDocx?(): Promise<Uint8Array>;
  /**
   * The spreadsheet as an xlsx file. For an xlsx this is save(); for xls, which
   * cannot be written, the model is written out as a new xlsx: what the model
   * holds, not what the xls reader leaves out (charts, macros, print settings).
   */
  saveAsXlsx?(): Promise<Uint8Array>;
  /**
   * The file as bytes. For xlsx, changes made to the model are written into
   * it; everything that was not changed keeps its original bytes.
   */
  save(): Promise<Uint8Array>;
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

/**
 * Opens a Word or Excel file. The format is detected from the content, never
 * from the file name.
 */
export async function open(input: Uint8Array | ArrayBuffer | Blob): Promise<OfficeFile> {
  const bytes =
    input instanceof Uint8Array
      ? input
      : new Uint8Array(input instanceof ArrayBuffer ? input : await input.arrayBuffer());

  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    const pkg = await OpcPackage.open(bytes);
    const main = await pkg.mainPart();
    const contentType = (main && pkg.contentTypeOf(main)) || '';
    const base = { container: 'ooxml' as const, package: pkg, save: () => pkg.save() };
    if (/wordprocessingml|ms-word/.test(contentType)) {
      let model: Promise<DocModel> | undefined;
      // Saving writes the edited stories of the model into the package first; a model that
      // was never read, or read and not changed, leaves the package as it is.
      const save = async () => {
        if (model) {
          const doc = await model.catch(() => undefined);
          if (doc) await writeDocx(pkg, doc);
        }
        return pkg.save();
      };
      return { format: 'docx', kind: 'document', ...base, save, workbook: notASpreadsheet, document: () => (model ??= readDocx(pkg)), saveAsDocx: save };
    }
    if (/spreadsheetml|ms-excel/.test(contentType)) {
      let workbook: Promise<Workbook> | undefined;
      // Saving writes the changes made to the model into the package first; a model
      // that was never read, or read and not changed, leaves the package as it is.
      const save = async () => {
        if (workbook) {
          // A file whose model could not be read was not changed either: it is saved as it is.
          const model = await workbook.catch(() => undefined);
          if (model) await writeXlsx(pkg, model);
        }
        return pkg.save();
      };
      return { format: 'xlsx', kind: 'spreadsheet', ...base, save, workbook: () => (workbook ??= readXlsx(pkg)), saveAsXlsx: save };
    }
    throw new UnsupportedFileError(`Unsupported OOXML document type: ${contentType || 'unknown'}`);
  }

  if (CfbFile.isCfb(bytes)) {
    const cfb = CfbFile.read(bytes);
    // No binary writer yet, and nothing can be modified, so saving returns the original bytes.
    const base = { container: 'cfb' as const, cfb, save: async () => bytes };
    if (cfb.has('WordDocument')) {
      let model: Promise<DocModel> | undefined;
      const document = () =>
        (model ??= Promise.resolve().then(() => {
          try {
            return readDoc(cfb);
          } catch (error) {
            if (error instanceof Error && error.name === 'UnsupportedFileError') throw new UnsupportedFileError(error.message);
            throw error;
          }
        }));
      return { format: 'doc', kind: 'document', ...base, workbook: notASpreadsheet, document, saveAsDocx: async () => buildDocx(await document()) };
    }
    if (cfb.has('Workbook') || cfb.has('Book')) {
      let model: Promise<Workbook> | undefined;
      // Read on first use. What the reader cannot open (encrypted, older than BIFF5) is reported as unsupported.
      const workbook = () =>
        (model ??= Promise.resolve().then(() => {
          try {
            return readXls(cfb);
          } catch (error) {
            if (error instanceof Error && error.name === 'UnsupportedFileError') throw new UnsupportedFileError(error.message);
            throw error;
          }
        }));
      return { format: 'xls', kind: 'spreadsheet', ...base, workbook, saveAsXlsx: async () => buildXlsx(await workbook()) };
    }
    if (cfb.has('EncryptedPackage')) throw new UnsupportedFileError('Password-protected files are not supported yet');
    throw new UnsupportedFileError('Compound file is not a Word or Excel document');
  }

  // "{\rtf": Word documents are often RTF saved with a .doc extension.
  if (isRtf(bytes)) {
    let model: Promise<DocModel> | undefined;
    // Nothing writes RTF and nothing can be changed, so saving returns the original bytes.
    return {
      format: 'rtf',
      kind: 'document',
      container: 'rtf',
      save: async () => bytes,
      workbook: notASpreadsheet,
      document: () => (model ??= Promise.resolve().then(() => readRtf(bytes))),
      saveAsDocx: async () => buildDocx(await (model ??= Promise.resolve().then(() => readRtf(bytes)))),
    };
  }

  throw new UnsupportedFileError('Not a Word or Excel file');
}

export interface MountOptions extends Omit<SpreadsheetEditorOptions, 'saveCommand'> {
  /** The file itself, or a URL to fetch it from. The format is detected from the content. */
  file: Uint8Array | ArrayBuffer | Blob | string;
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
   * The file with every change made in the view, as bytes. A cell that is
   * being edited is written first, as leaving it would do.
   */
  save(): Promise<Uint8Array>;
  /**
   * Prints the sheet that is shown, or every visible sheet, through the
   * print dialog of the browser (which can also save a PDF): see
   * WorkbookView.print(). Resolves when the dialog was closed.
   */
  print(options?: SheetPrintOptions): Promise<DocumentPrintResult>;
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
  /** The file as bytes: a docx with the changes made in the editor; doc and RTF, which are read-only, as the bytes they were opened from. */
  save(): Promise<Uint8Array>;
  /**
   * Prints the document through the print dialog of the browser (which can
   * also save it as a PDF): see DocumentView.print(). Resolves when the dialog was closed.
   */
  print(options?: DocumentPrintOptions): Promise<DocumentPrintResult>;
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
  const autoSave = resolveAutoSave(options.autoSave);
  if (autoSave && !saveTarget) throw new TypeError('"autoSave" needs "saveTo": there is nowhere to save to');
  /** The ETag the file came with: what a save to the same address is sent with. */
  let etag: string | null = null;
  let fileName: string | undefined;
  if (typeof input === 'string') {
    fileName = nameFromUrl(input, base);
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
    // Not the ETag of an answer that came from another address than the one that is saved to.
    etag = fileETag(response);
    input = await response.arrayBuffer();
  } else if (typeof Blob !== 'undefined' && input instanceof Blob && typeof (input as { name?: unknown }).name === 'string') {
    fileName = (input as unknown as { name: string }).name || undefined;
  }
  let file = await open(input);
  let conversion: MountConversion | undefined;
  /** Why a file that had to be converted was not: it is then shown as it is, read-only. */
  let notConverted: string | undefined;
  if (options.convert === true && options.readOnly !== true && needsConversion(file.format)) {
    if (file.kind === 'document' && !file.document) throw new UnsupportedFileError('Showing doc documents is not supported yet');
    // The model is written out as a new file of the format that can be saved, and that file is what is opened.
    const from = file.format;
    const opened = await openConverted(file, open);
    if (opened.conversion) {
      file = opened.file;
      conversion = { from: opened.conversion.from, to: file.format as 'docx' | 'xlsx', warnings: opened.conversion.warnings };
    } else if (opened.failed !== undefined) notConverted = from;
  }
  const fontFailures = await fontsSettled;
  const converted = (): void => {
    if (!options.onNotice) return;
    // In the language the views speak for the same option.
    const strings = resolveUiStrings(options.locale);
    for (let i = 0; i < fontFailures.length; i++) {
      options.onNotice({ message: strings.fontNotLoaded(fontFailures[i].family, fontFailures[i].pending), kind: fontFailures[i].pending ? 'info' : 'error' });
    }
    if (conversion) options.onNotice({ message: strings.openedConverted(conversion.from, conversion.to), kind: 'info' });
    else if (notConverted !== undefined) options.onNotice({ message: strings.notConverted(notConverted), kind: 'error' });
    // Said before anything is typed: what is changed in this file has nowhere to go (see MountOptions.saveTo).
    if (conversion && controller && controller.blockedConversion) options.onNotice({ message: strings.saveConvertedNotAllowed(conversion.to), kind: 'error' });
  };
  let save: () => Promise<Uint8Array>;
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
        autoSave,
        info,
        strings: () => resolveUiStrings(options.locale),
        notice: (message, kind) => notice(message, kind),
        onState: options.onSaveState,
        onSaved: options.onSave,
        onError: options.onError,
        // A save the server refused with 412: the user is asked over the editor, unless the page said
        // `onConflict: 'fail'` (the controller looks at that). A function of the application knows its own conflicts.
        askOverwrite:
          saveTarget.kind === 'request' && typeof document !== 'undefined'
            ? (question) => {
                // A view that is only read is not asked to overwrite anything: the conflict was told, and that is all.
                if (readOnlyNow() || !editorRoot) return question.answer(false);
                return askOverwrite({
                  root: editorRoot(),
                  strings: resolveUiStrings(options.locale),
                  detail: question.detail,
                  answer: question.answer,
                  focusView: () => focusView(),
                });
              }
            : undefined,
        etag,
      })
    : undefined;
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
        if (controller) controller.destroy();
      }
      destroyEditor();
    };
    return () => editor.destroy();
  };
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
    });
    const view = editor.view;
    if (file.format !== 'docx') {
      // Nothing may turn editing on: saving gives the original bytes and the edits would be lost.
      const lockView = view.setReadOnly.bind(view);
      const lockEditor = editor.setReadOnly.bind(editor);
      view.setReadOnly = () => lockView(true);
      editor.setReadOnly = () => lockEditor(true);
    }
    save = async () => {
      // A composition still under way belongs to the document that is written.
      view.commitEditing();
      const changes = view.changeCount;
      const bytes = await file.save();
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
    const print = (printOptions?: DocumentPrintOptions) => view.print(printOptions);
    const mounted: MountedDocument = {
      kind: 'document',
      file,
      document: model,
      view,
      editor,
      save,
      print,
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

  const workbook = await file.workbook();
  // An xls file cannot be written yet: it is shown read-only, so that no edit is lost on saving.
  const sheetOptions: SpreadsheetEditorOptions = {
    ...options,
    onChange: (change) => {
      changed();
      if (options.onChange) options.onChange(change);
    },
    saveCommand: controller,
  };
  if (file.format === 'xls') sheetOptions.readOnly = true;
  const editor = new SpreadsheetEditor(container, workbook, sheetOptions);
  const view = editor.view;
  if (file.format === 'xls') {
    // Nothing may turn editing on: saving would give the original bytes and the edits would be lost.
    const lockView = view.setReadOnly.bind(view);
    const lockEditor = editor.setReadOnly.bind(editor);
    view.setReadOnly = () => lockView(true);
    editor.setReadOnly = () => lockEditor(true);
  }
  save = async () => {
    // What is being typed (in a cell, in the field of a note, in a tab) is written first, as leaving it would do.
    if (!view.commitPending()) throw new Error('Cannot save: a note or a sheet name that is being typed cannot be applied, or a dialog of rules is open with changes (confirm or close it first)');
    return file.save();
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
  saveAsIs = () => file.save();
  notice = (message, kind) => view.showNotice(message, kind);
  focusView = () => view.focus();
  editorRoot = () => editor.element;
  readOnlyNow = () => view.readOnly;
  watchReadOnly(view, () => view.hasUnappliedText());
  const destroy = tie(editor);
  converted();
  const print = (printOptions?: SheetPrintOptions) => view.print(printOptions);
  const mounted: MountedSpreadsheet = {
    kind: 'spreadsheet',
    file,
    workbook,
    view,
    editor,
    save,
    print,
    saveNow,
    get hasUnsavedChanges() {
      return unsaved();
    },
    setAutoSave,
    destroy,
  };
  if (conversion) mounted.conversion = conversion;
  return mounted;
}
