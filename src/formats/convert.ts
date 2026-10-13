// Opening the formats that cannot be written back (doc, RTF, xls) for
// editing: the model is written out as a new file of the format that can be
// saved, doc and RTF as docx, xls as xlsx, and that file is opened. mount()
// and the WebView bridge both open their files this way.

import type { DocModel } from '../model/doc/types';
import type { Workbook } from '../model/sheet/types';
import { encryptPackage } from '../crypto/ooxml';
import { assertProtectionAlive, inheritProtection, sessionOf } from '../crypto/session';
import { buildDocx, emptyWriteReport } from './docx/write';
import { buildXlsx } from './xlsx/build';

/** What the conversion needs of an opened file. */
export interface ConvertibleFile {
  format: string;
  kind: 'document' | 'spreadsheet';
  workbook(): Promise<Workbook>;
  document?(): Promise<DocModel>;
}

/** A file written out in the format it can be edited and saved in. */
export interface EditableConversion {
  /** The format of the file that was given. */
  from: 'doc' | 'rtf' | 'xls';
  /** The format of `bytes`. */
  to: 'docx' | 'xlsx';
  /**
   * The new file; opening it gives the file that is edited and saved. For a file that has a password
   * (it was opened with one, or setPassword() gave it one) these bytes are encrypted with that password,
   * and open() needs it: nothing that is made from a protected file is ever given out without its password.
   */
  bytes: Uint8Array;
  /** Set (true) when `bytes` is encrypted. */
  encrypted?: true;
  /** What the conversion could not carry over, as the writers report it; empty when nothing the model held was lost. */
  warnings: string[];
}

/** What openConverted() resolves to. */
export interface OpenedForEditing<F> {
  /** The file to show: the converted one, or the one that was given when nothing was or could be converted. */
  file: F;
  /** Set when `file` is the converted file. */
  conversion?: { from: 'doc' | 'rtf' | 'xls'; to: 'docx' | 'xlsx'; warnings: string[] };
  /** Set when the file needed a conversion and it failed: why, in English. `file` is then the file as it was given. */
  failed?: string;
}

/** Whether a file of this format has to be converted before it can be edited. */
export function needsConversion(format: string): format is 'doc' | 'rtf' | 'xls' {
  return format === 'doc' || format === 'rtf' || format === 'xls';
}

/**
 * The file as a docx (doc, RTF) or an xlsx (xls), or undefined for a file
 * that can be edited as it is. A document format without a reader rejects
 * with an error named UnsupportedFileError.
 */
export async function convertForEditing(file: ConvertibleFile): Promise<EditableConversion | undefined> {
  if (!needsConversion(file.format)) return undefined;
  // A protected file that was closed has forgotten its password: it is not converted at all.
  assertProtectionAlive(file);
  const written = await convertPlain(file);
  const session = sessionOf(file);
  if (written && session) {
    written.bytes = await encryptPackage(written.bytes, session);
    written.encrypted = true;
  }
  return written;
}

/** The conversion itself: the new package, never encrypted. Stays inside the engine. */
async function convertPlain(file: ConvertibleFile): Promise<EditableConversion | undefined> {
  const from = file.format;
  if (!needsConversion(from)) return undefined;
  if (file.kind === 'document') {
    if (!file.document) throw Object.assign(new Error('Showing doc documents is not supported yet'), { name: 'UnsupportedFileError' });
    const report = emptyWriteReport();
    const bytes = await buildDocx(await file.document(), { report });
    const warnings = report.warnings.slice();
    // Characters an XML file cannot hold were left out of the text.
    if (report.droppedCharacters > 0) warnings.push(`${report.droppedCharacters} character(s) that a docx file cannot hold were left out`);
    return { from, to: 'docx', bytes, warnings };
  }
  const report = { parts: [], warnings: [] as string[] };
  const bytes = await buildXlsx(await file.workbook(), { report });
  return { from, to: 'xlsx', bytes, warnings: report.warnings };
}

/**
 * Converts a doc, RTF or xls file and opens the result with `open`; a file
 * of another format is given back as it is. A file whose own model cannot be
 * read rejects, as showing it would. When the model can be read but the new
 * file cannot be written or read back, the file is given back as it was,
 * with the reason in `failed`: it can still be shown read-only.
 */
export async function openConverted<F extends ConvertibleFile>(file: F, open: (bytes: Uint8Array) => Promise<F>): Promise<OpenedForEditing<F>> {
  if (!needsConversion(file.format)) return { file };
  // What cannot be read cannot be shown either: this is the caller's error to report.
  if (file.kind === 'document') {
    if (!file.document) throw Object.assign(new Error('Showing doc documents is not supported yet'), { name: 'UnsupportedFileError' });
    await file.document();
  } else await file.workbook();
  try {
    // Before anything is written: the new file could not be given the password of a file that has forgotten its own.
    assertProtectionAlive(file);
    const written = await convertPlain(file);
    if (!written) return { file };
    const converted = await open(written.bytes);
    if (converted.format !== written.to) throw new Error(`The converted file is not a ${written.to} file`);
    // A password-protected original: the new file is saved with the same password, and says how its original was encrypted.
    inheritProtection(file, converted);
    const encryption = (file as { encryption?: unknown }).encryption;
    if (encryption !== undefined) (converted as { encryption?: unknown }).encryption = encryption;
    // The model of the new file is what will be shown: it has to be readable.
    if (converted.kind === 'document') {
      if (!converted.document) throw new Error('The converted file cannot be shown');
      await converted.document();
    } else {
      const workbook = await converted.workbook();
      if (workbook.sheets.length === 0) throw new Error('The converted workbook has no sheets');
    }
    return { file: converted, conversion: { from: written.from, to: written.to, warnings: written.warnings } };
  } catch (error) {
    return { file, failed: error instanceof Error && error.message ? error.message : String(error) };
  }
}
