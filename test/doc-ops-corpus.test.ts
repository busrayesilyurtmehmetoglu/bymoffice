// Random edits on real documents: none may throw anything but a refusal, and
// undoing all of them must give back the document as it was read, down to the
// very block objects.
//
// The documents are those of the local corpus (.corpus.local, see
// scripts/corpus.ts). They are only read. Without that file the tests are
// skipped, so the suite runs on any machine. Nothing of a document (name,
// path, text) is printed: a failure names the format and a number.

import assert from 'node:assert/strict';
import { closeSync, existsSync, openSync, readFileSync, readSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { readDoc } from '../src/formats/doc/document';
import { readDocx } from '../src/formats/docx/document';
import { readRtf } from '../src/formats/rtf/document';
import type { DocModel } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, resolvePosition, type DocRange } from '../src/ops/doc/index';
import { assertSameIdentities, checkInvariants, checkLayout, dice, docPicture, identities, randomEdit } from './doc-ops-fixtures';

const SKIP_DIRS = new Set(['node_modules', 'vendor', '.git']);
const EDITS = 50;

function corpusDirs(): string[] {
  if (!existsSync('.corpus.local')) return [];
  return readFileSync('.corpus.local', 'utf8')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '' && line.charAt(0) !== '#' && line.indexOf('--') !== 0 && existsSync(line));
}

type Format = 'docx' | 'doc' | 'rtf';

/** What a file is, by how it begins: the extension says little in this corpus (many ".doc" files are RTF). */
function sniff(path: string): Format | undefined {
  const head = Buffer.alloc(8);
  let fd: number | undefined;
  try {
    fd = openSync(path, 'r');
    readSync(fd, head, 0, 8, 0);
  } catch {
    return undefined;
  } finally {
    if (fd !== undefined) closeSync(fd);
  }
  if (head[0] === 0x50 && head[1] === 0x4b) return 'docx';
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) return 'doc';
  if (head.toString('latin1', 0, 5) === '{\\rtf') return 'rtf';
  return undefined;
}

let found: Record<Format, string[]> | undefined;

/** The word processing files of the corpus by format. */
function corpus(): Record<Format, string[]> {
  if (found) return found;
  const out: Record<Format, string[]> = { docx: [], doc: [], rtf: [] };
  const walk = (dir: string): void => {
    for (const item of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, item.name);
      if (item.isDirectory()) {
        if (!SKIP_DIRS.has(item.name)) walk(path);
      } else if (item.isFile() && /\.(docx|docm|dotx|doc|rtf)$/i.test(item.name) && item.name.indexOf('~$') !== 0) {
        const format = sniff(path);
        if (format) out[format].push(path);
      }
    }
  };
  for (const dir of corpusDirs()) walk(dir);
  return (found = out);
}

/** Up to `count` of the files, spread evenly over the sorted list. */
function sample(files: string[], count: number): string[] {
  files.sort();
  if (files.length <= count) return files;
  const out: string[] = [];
  for (let i = 0; i < count; i++) out.push(files[Math.floor((i * files.length) / count)]);
  return out;
}

async function check(format: Format, count: number, read: (bytes: Uint8Array) => Promise<DocModel> | DocModel, diagnostic: (message: string) => void): Promise<void> {
  const picked = sample(corpus()[format].slice(), count);
  let documents = 0;
  let unread = 0;
  let applied = 0;
  let refused = 0;
  let laid = 0;
  let laidExact = 0;
  const failures: string[] = [];
  for (let n = 0; n < picked.length; n++) {
    let doc: DocModel;
    try {
      doc = await read(new Uint8Array(readFileSync(picked[n])));
    } catch {
      // Not a document of this format after all, or one the reader refuses: nothing to edit.
      unread++;
      continue;
    }
    documents++;
    try {
      // Does the document show its text as the model holds it (no capitals, no symbol fonts)? Then it must after editing too.
      let exact = n % 4 === 0;
      if (exact) {
        try {
          checkLayout(doc);
        } catch (error) {
          if (!(error instanceof assert.AssertionError)) throw error;
          exact = false;
        }
      }
      const before = docPicture(doc);
      const objects = identities(doc);
      const history = new DocHistory<DocRange | undefined>(doc, 1000);
      const d = dice(n + 1);
      for (let step = 0; step < EDITS; step++) {
        let made;
        try {
          made = randomEdit(doc, d);
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
          refused++;
          continue;
        }
        if (!made || made.edit.ops.length === 0) continue;
        history.applyEdit(made.edit, undefined);
        history.seal();
        applied++;
        resolvePosition(doc, made.edit.selection.anchor);
        resolvePosition(doc, made.edit.selection.focus);
      }
      checkInvariants(doc, false);
      // The edited model goes through the layout engine; large documents only now and then, for the time it takes.
      if (n % 4 === 0) {
        checkLayout(doc, exact);
        laid++;
        if (exact) laidExact++;
      }
      while (history.undo());
      if (docPicture(doc) !== before) throw new Error('the document is not as it was read after undoing everything');
      assertSameIdentities(identities(doc), objects);
    } catch (error) {
      const message = error instanceof Error ? `${error.name}: ${error.message.slice(0, 200)}` : String(error);
      // Only the kind of failure: no text of the document.
      failures.push(`${format} #${n}: ${message.replace(/"[^"]*"/g, '"…"').replace(/'[^']*'/g, "'…'")}`);
    }
  }
  diagnostic(`${format}: ${documents} documents edited (${unread} could not be read), ${applied} edits applied, ${refused} refused, ${laid} laid out after editing (${laidExact} with every text item compared to the model)`);
  assert.deepEqual(failures, []);
}

const skip = corpusDirs().length === 0 ? 'no local corpus' : false;

test('real docx documents: 50 random edits each, then undo', { skip, timeout: 900000 }, async (t) => {
  await check('docx', 100, async (bytes) => readDocx(await OpcPackage.open(bytes)), (m) => t.diagnostic(m));
});

test('real doc documents: 50 random edits each, then undo', { skip, timeout: 900000 }, async (t) => {
  await check('doc', 200, (bytes) => readDoc(CfbFile.read(bytes)), (m) => t.diagnostic(m));
});

test('real rtf documents: 50 random edits each, then undo', { skip, timeout: 900000 }, async (t) => {
  await check('rtf', 200, (bytes) => readRtf(bytes), (m) => t.diagnostic(m));
});
