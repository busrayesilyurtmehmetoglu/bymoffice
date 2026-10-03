import { CfbFile } from './cfb';
import { readDoc } from './formats/doc/document';
import { readRtf } from './formats/rtf/document';
import { readDocx } from './formats/docx/document';
import { buildDocx, writeDocx } from './formats/docx/write';
import { documentPictures } from './model/doc/pictures';
import type { DocModel } from './model/doc/types';
import { readXls } from './formats/xls/workbook';
import { readXlsx } from './formats/xlsx/workbook';
import { writeXlsx } from './formats/xlsx/write';
import type { Workbook } from './model/sheet/types';
import { SpreadsheetEditor, type SpreadsheetEditorOptions } from './ui/editor';
import type { WorkbookView } from './ui/sheet-view';
import type { DocumentView, DocumentZoom } from './ui/doc-view';
import { DocumentEditor } from './ui/doc-editor';
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
export * as docOps from './ops/doc/index';
export { readDoc } from './formats/doc/document';
export { readRtf } from './formats/rtf/document';
export { documentPictures } from './model/doc/pictures';
export { inFieldCode, paragraphText, rawOf, safeHyperlinkTarget, type DocModel } from './model/doc/types';
export { NumberingCounters } from './formats/docx/numbering';
export { resolveNumberRunProps, resolveParaProps, resolveRunProps } from './formats/docx/resolve';
export { CalcEngine, type CalcEngineOptions } from './calc/engine';
export { ConditionalFormats, type CellConditional, type ConditionalSource } from './format/conditional';
export { DataValidations, MAX_LIST_ITEMS, ValidationError, type ValidationList, type ValidationOptions, type ValidationVerdict } from './format/validation';
export { FormulaParseError, parseFormula } from './calc/parse';
export { FUNCTIONS } from './calc/functions/index';
export { calcError, isError, type CalcError, type ErrorCode, type Scalar, type Value } from './calc/values';
export { writeXlsx } from './formats/xlsx/write';
export { History, applyOp, ensureCellFormat, ensureNumberFormat, type CellChange, type Op } from './ops/ops';
export { clearRange, copyRange, editTextAt, pasteAt, setCellInput, type CopiedRange, type EditOptions, type PasteResult } from './ops/commands';
export { MAX_FILTER_VALUES, clearFilter, filterSummary, filterValues, setAutoFilter, setFilterColumn, type FilterCriterion, type FilterValue, type FilterValueList } from './ops/filter';
export { collationKey, currentRegion, sortRange, type SortKey, type SortOptions } from './ops/sort';
export { parseCellInput, cellEditText, type ParsedInput, type InputOptions } from './edit/input';
export { formulaFromLocal, formulaToLocal } from './edit/formula-locale';
export { parseTsv, toTsv } from './edit/clipboard';
export { adjustFormula, adjustRange, adjustReferenceText, shiftFormula, type StructuralEdit } from './calc/refs';
export { DEFAULT_GRID_THEME, cellButtonRect, renderGrid, type GridFilter, type GridScene, type GridSelection, type GridTheme } from './render/grid';
export type { GridCommentSource, GridDrawingSource } from './render/grid';
export { SheetDrawings, adjustCommentCell, adjustDrawingAnchor, anchorRect, type DrawingTarget, type PlacedComment, type SheetDrawingLabels, type SheetDrawingsOptions } from './render/sheet-drawings';
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
export type { DocLayoutOptions, FontMetrics, ImageItem, LayoutBudget, LayoutFont, LayoutPage, LayoutProgress, LineItem, PageItem, RectItem, TextItem, TextMeasurer } from './layout/doc/types';
export { renderPage, type PageImage, type RenderPageOptions } from './render/doc';
export { CanvasTextMeasurer, createCanvasMeasurer, docCssFont, type CanvasMeasurerOptions, type FontHint, type MeasureContext } from './render/doc-measure';
export { DocumentView, documentText, type DocumentViewOptions, type DocumentViewState, type DocumentZoom } from './ui/doc-view';
export { DocumentEditor, type DocumentEditorOptions } from './ui/doc-editor';
export { DocToolbar, type DocShellHost } from './ui/doc-toolbar';
export { DocLines, type Affinity, type CaretBox, type PageLine, type Place } from './ui/doc-caret';
export { sliceHtml, sliceOf, type DocSlice } from './ui/doc-clipboard';
export type { LineMark } from './layout/doc/types';
export type { EditProgress } from './layout/doc/layout';

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
      return { format: 'xlsx', kind: 'spreadsheet', ...base, save, workbook: () => (workbook ??= readXlsx(pkg)) };
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
      return { format: 'xls', kind: 'spreadsheet', ...base, workbook };
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
  /** Documents: the zoom, 1 = 100%, or a mode that follows the container. Default: 100%, less when the page is wider than the container. */
  zoom?: DocumentZoom;
  /** Documents: called when the page in view or the number of pages changes. `page` counts from 1. */
  onPageChange?(page: number, pageCount: number): void;
  /** Documents: called when a hyperlink is activated; return false to keep the view from opening it in a new window. */
  onLinkClick?(target: string): boolean | void;
  /** Documents: called after every change made in the editor, undo and redo included. */
  onDocumentChange?(): void;
}

export interface MountedSpreadsheet {
  kind: 'spreadsheet';
  file: OfficeFile;
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
  /** Removes the view from the container. */
  destroy(): void;
}

export interface MountedDocument {
  kind: 'document';
  file: OfficeFile;
  document: DocModel;
  /** The pages: zoom, navigation, the caret and the selection, editing, undo and redo. */
  view: DocumentView;
  /** The editor around the pages: the formatting toolbar. */
  editor: DocumentEditor;
  /** The file as bytes; a document cannot be changed yet, so these are the bytes it was opened from. */
  save(): Promise<Uint8Array>;
  /** Removes the view from the container. */
  destroy(): void;
}

/** What mount() resolves to; `kind` tells which of the two it is. */
export type MountedView = MountedSpreadsheet | MountedDocument;

/**
 * Opens a file and shows it inside `container`, which must have a size: a
 * spreadsheet (xlsx) in the editor, with a toolbar, a formula bar and a status
 * bar unless the options turn them off; a document as pages, a docx in the
 * document editor (unless `readOnly` is set), doc and RTF read-only.
 * Other formats reject with UnsupportedFileError.
 */
export async function mount(container: HTMLElement, options: MountOptions): Promise<MountedView> {
  let input = options.file;
  if (typeof input === 'string') {
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
    input = await response.arrayBuffer();
  }
  const file = await open(input);
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
    });
    const view = editor.view;
    save = async () => {
      // A composition still under way belongs to the document that is written.
      view.commitEditing();
      const bytes = await file.save();
      view.markSaved();
      return bytes;
    };
    container.addEventListener('keydown', onKeyDown);
    const destroy = () => {
      container.removeEventListener('keydown', onKeyDown);
      editor.destroy();
    };
    return { kind: 'document', file, document: model, view, editor, save, destroy };
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
  return { kind: 'spreadsheet', file, workbook, view, editor, save, destroy };
}
