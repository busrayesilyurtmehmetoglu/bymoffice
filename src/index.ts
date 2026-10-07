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
import { writeXlsx } from './formats/xlsx/write';
import type { Workbook } from './model/sheet/types';
import { SpreadsheetEditor, type SpreadsheetEditorOptions } from './ui/editor';
import type { WorkbookView } from './ui/sheet-view';
import type { DocumentView, DocumentZoom } from './ui/doc-view';
import { DocumentEditor } from './ui/doc-editor';
import type { DocumentPrintOptions, DocumentPrintResult } from './ui/doc-print';
import type { SheetPrintOptions } from './ui/sheet-print';
import { OpcPackage } from './opc';

export { CfbFile, type CfbEntry } from './cfb';
export { OpcPackage, type Relationship } from './opc';
export { ZipArchive, type ZipEntry } from './zip';
export * from './xml';
export * from './model/sheet/types';
export * from './model/sheet/color';
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
export { canDeleteSheet, canInsertSheet, canRenameSheet, defaultSheetName, deleteSheet, insertSheet, renameSheet, type InsertSheetOptions } from './ops/sheets';
export { DEFAULT_NOTE_AUTHOR, MAX_NOTE_TEXT, commentAtCell, hasLiveComments, noteProblem, setComment, type NoteOptions } from './ops/notes';
export { clearRange, copyRange, editTextAt, pasteAt, setCellInput, type CopiedRange, type EditOptions, type PasteResult } from './ops/commands';
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
export { CanvasTextMeasurer, createCanvasMeasurer, docCssFont, type CanvasMeasurerOptions, type FontHint, type MeasureContext } from './render/doc-measure';
export { DocumentView, documentText, type DocumentViewOptions, type DocumentViewState, type DocumentZoom } from './ui/doc-view';
export { DocumentEditor, type DocumentEditorOptions } from './ui/doc-editor';
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
 */
export const fonts = { registerFont, registeredFonts, unregisterFont };

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

export interface MountOptions extends SpreadsheetEditorOptions {
  /** The file itself, or a URL to fetch it from. The format is detected from the content. */
  file: Uint8Array | ArrayBuffer | Blob | string;
  /**
   * Called with the file's bytes when the user saves with Ctrl+S or Cmd+S in
   * the view. Without it the shortcut is left to the browser.
   */
  onSave?(bytes: Uint8Array): void;
  /** Called when saving from the shortcut fails. */
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
   */
  fonts?: FontRegistration[];
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
   * "readOnly", "comments" or "forms": the document is shown read-only and
   * nothing turns editing on (comments and form fields cannot be edited
   * here either). "trackedChanges": the document can be edited, and every
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
  /** Removes the view from the container. */
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
  /** Removes the view from the container. */
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
  // The fonts load while the file is fetched and read. This never rejects: a font is no reason not to show a file.
  const fontsSettled: Promise<FontFailure[]> = Array.isArray(options.fonts) ? settleFonts(options.fonts, FONT_WAIT_MS) : Promise.resolve([]);
  let input = options.file;
  if (typeof input === 'string') {
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
    input = await response.arrayBuffer();
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
  };
  let save: () => Promise<Uint8Array>;
  const onKeyDown = (event: KeyboardEvent) => {
    if (!options.onSave || !(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey || (event.key !== 's' && event.key !== 'S')) return;
    event.preventDefault();
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
      onChange: options.onDocumentChange,
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
    container.addEventListener('keydown', onKeyDown);
    const destroy = () => {
      container.removeEventListener('keydown', onKeyDown);
      editor.destroy();
    };
    converted();
    const print = (printOptions?: DocumentPrintOptions) => view.print(printOptions);
    const mounted: MountedDocument = { kind: 'document', file, document: model, view, editor, save, print, destroy };
    if (conversion) mounted.conversion = conversion;
    if (view.trackingChanges) mounted.trackChanges = true;
    const enforced = view.protection;
    if (enforced) mounted.protection = { edit: enforced, exceptions: hasPermissionRanges(model) };
    return mounted;
  }

  const workbook = await file.workbook();
  // An xls file cannot be written yet: it is shown read-only, so that no edit is lost on saving.
  const editor = new SpreadsheetEditor(container, workbook, file.format === 'xls' ? { ...options, readOnly: true } : options);
  const view = editor.view;
  if (file.format === 'xls') {
    // Nothing may turn editing on: saving would give the original bytes and the edits would be lost.
    const lockView = view.setReadOnly.bind(view);
    const lockEditor = editor.setReadOnly.bind(editor);
    view.setReadOnly = () => lockView(true);
    editor.setReadOnly = () => lockEditor(true);
  }
  save = async () => {
    if (view.isEditing) view.commitEditing();
    return file.save();
  };
  container.addEventListener('keydown', onKeyDown);
  const destroy = () => {
    container.removeEventListener('keydown', onKeyDown);
    editor.destroy();
  };
  converted();
  const print = (printOptions?: SheetPrintOptions) => view.print(printOptions);
  return conversion
    ? { kind: 'spreadsheet', file, conversion, workbook, view, editor, save, print, destroy }
    : { kind: 'spreadsheet', file, workbook, view, editor, save, print, destroy };
}
