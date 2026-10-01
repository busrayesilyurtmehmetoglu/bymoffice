import { CfbFile } from './cfb';
import { readXlsx } from './formats/xlsx/workbook';
import type { Workbook } from './model/sheet/types';
import { WorkbookView, type WorkbookViewOptions } from './ui/sheet-view';
import { OpcPackage } from './opc';

export { CfbFile, type CfbEntry } from './cfb';
export { OpcPackage, type Relationship } from './opc';
export { ZipArchive, type ZipEntry } from './zip';
export * from './xml';
export * from './model/sheet/types';
export * from './model/sheet/color';
export { readXlsx } from './formats/xlsx/workbook';
export { DEFAULT_GRID_THEME, renderGrid, type GridScene, type GridSelection, type GridTheme } from './render/grid';
export { cellText, formatGeneral } from './render/cell-text';
export { WorkbookView, type WorkbookViewOptions } from './ui/sheet-view';
export * from './layout/grid-geometry'; export * from './layout/sheet-metrics';

export type OfficeFormat = 'docx' | 'xlsx' | 'doc' | 'xls';

export interface OfficeFile {
  format: OfficeFormat;
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb';
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
  save(): Promise<Uint8Array>;
}

/** The input is not a Word or Excel file this engine can open. */
export class UnsupportedFileError extends Error {
  override name = 'UnsupportedFileError';
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
      return { format: 'docx', kind: 'document', ...base, workbook: notASpreadsheet };
    }
    if (/spreadsheetml|ms-excel/.test(contentType)) {
      let workbook: Promise<Workbook> | undefined;
      return { format: 'xlsx', kind: 'spreadsheet', ...base, workbook: () => (workbook ??= readXlsx(pkg)) };
    }
    throw new UnsupportedFileError(`Unsupported OOXML document type: ${contentType || 'unknown'}`);
  }

  if (CfbFile.isCfb(bytes)) {
    const cfb = CfbFile.read(bytes);
    // No binary writer yet, and nothing can be modified, so saving returns the original bytes.
    const base = { container: 'cfb' as const, cfb, save: async () => bytes };
    if (cfb.has('WordDocument')) return { format: 'doc', kind: 'document', ...base, workbook: notASpreadsheet };
    if (cfb.has('Workbook') || cfb.has('Book')) {
      const workbook = () => Promise.reject(new UnsupportedFileError('Reading xls cell data is not supported yet'));
      return { format: 'xls', kind: 'spreadsheet', ...base, workbook };
    }
    if (cfb.has('EncryptedPackage')) throw new UnsupportedFileError('Password-protected files are not supported yet');
    throw new UnsupportedFileError('Compound file is not a Word or Excel document');
  }

  // "{\rtf": Word documents are often RTF saved with a .doc extension.
  if (bytes[0] === 0x7b && bytes[1] === 0x5c && bytes[2] === 0x72 && bytes[3] === 0x74 && bytes[4] === 0x66) {
    throw new UnsupportedFileError('RTF documents are not supported yet');
  }

  throw new UnsupportedFileError('Not a Word or Excel file');
}

export interface MountOptions extends WorkbookViewOptions {
  /** The file itself, or a URL to fetch it from. The format is detected from the content. */
  file: Uint8Array | ArrayBuffer | Blob | string;
}

export interface MountedView {
  file: OfficeFile;
  workbook: Workbook;
  view: WorkbookView;
  /** Removes the view from the container. */
  destroy(): void;
}

/**
 * Opens a spreadsheet and shows it inside `container`, which must have a size.
 * Only xlsx can be shown so far; other formats reject with UnsupportedFileError.
 */
export async function mount(container: HTMLElement, options: MountOptions): Promise<MountedView> {
  let input = options.file;
  if (typeof input === 'string') {
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Could not load the file: HTTP ${response.status}`);
    input = await response.arrayBuffer();
  }
  const file = await open(input);
  const workbook = await file.workbook();
  const view = new WorkbookView(container, workbook, options);
  return { file, workbook, view, destroy: () => view.destroy() };
}
