// What the tests of frozen panes, merged cells and tab stops share: a workbook loaded with its
// history, saved through the writer, and files handed to python3 (openpyxl, zipfile and
// xml.etree, test/tools/package_check.py) for a reading that owes nothing to the engine.

import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open } from '../src/index';
import type { OfficeFile, Workbook } from '../src/index';
import { History, type Op } from '../src/ops/ops';

export interface Loaded {
  bytes: Uint8Array;
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

export async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { bytes, file, wb, history: new History(wb) };
}

export async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

export const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

const decoder = new TextDecoder();
export const partText = async (bytes: Uint8Array, name: string): Promise<string> => decoder.decode(await ZipArchive.read(bytes).read(name));

/** Names of the parts whose bytes differ between two packages (or that only one has). */
export async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = ZipArchive.read(before);
  const b = ZipArchive.read(after);
  const names = new Set<string>(a.names().concat(b.names()));
  const out: string[] = [];
  for (const name of names) {
    const x = a.has(name) ? await a.read(name) : undefined;
    const y = b.has(name) ? await b.read(name) : undefined;
    if (!x || !y || x.length !== y.length || x.some((value, i) => value !== y[i])) out.push(name);
  }
  return out.sort();
}

export const sameBytes = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((value, i) => value === b[i]);

/** Whether python3 can import these modules. */
export function hasPython(modules: string[]): boolean {
  const probe = spawnSync('python3', ['-c', modules.map((name) => `import ${name}`).join('; ')], { encoding: 'utf8' });
  return probe.status === 0;
}

/**
 * Writes the files into a folder of the repository (removed afterwards), and runs a python
 * script with the folder as its argument; what it prints is read as JSON.
 */
export function runPython<T>(files: { [name: string]: Uint8Array }, script: string): T {
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-edit-commands-${process.pid}-${Math.floor(Math.random() * 1e9)}`);
  mkdirSync(folder, { recursive: true });
  try {
    for (const name of Object.keys(files)) writeFileSync(join(folder, name), files[name]);
    const run = spawnSync('python3', ['-c', script, folder], { encoding: 'utf8', timeout: 120000 });
    if (run.status !== 0) throw new Error(`python3 failed: ${run.stderr.slice(-2000)}`);
    return JSON.parse(run.stdout) as T;
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

/** What test/tools/package_check.py finds wrong with the files: an empty list when all is well. */
export function packageProblems(files: { [name: string]: Uint8Array }): unknown[] {
  const tool = join(dirname(fileURLToPath(import.meta.url)), 'tools', 'package_check.py');
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-edit-commands-${process.pid}-${Math.floor(Math.random() * 1e9)}`);
  mkdirSync(folder, { recursive: true });
  try {
    for (const name of Object.keys(files)) writeFileSync(join(folder, name), files[name]);
    const run = spawnSync('python3', ['-I', tool, folder], { encoding: 'utf8', timeout: 120000 });
    if (run.status !== 0) throw new Error(`package_check.py failed: ${run.stderr.slice(-2000)}`);
    return JSON.parse(run.stdout) as unknown[];
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}
