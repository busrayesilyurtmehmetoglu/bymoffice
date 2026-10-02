// The docx writer on real documents: those of the local corpus
// (.corpus.local, see scripts/corpus.ts). They are only read. Without that
// file the tests are skipped, so the suite runs on any machine. Nothing of a
// document (name, path, text) is printed: a failure names a format and a
// number.

import assert from 'node:assert/strict';
import { closeSync, existsSync, openSync, readFileSync, readSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { readDoc } from '../src/formats/doc/document';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { readRtf } from '../src/formats/rtf/document';
import { rawOf, type Block, type DocModel, type Wrapper } from '../src/model/doc/types';
import { DocHistory, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { dice, randomEdit } from './doc-ops-fixtures';
import { convertedPicture, differenceKind, documentPicture, firstDifference, schemaOrderViolations } from './docx-write-helpers.test';

const SKIP_DIRS = new Set(['node_modules', 'vendor', '.git']);

function corpusDirs(): string[] {
  if (!existsSync('.corpus.local')) return [];
  return readFileSync('.corpus.local', 'utf8')
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line !== '' && line.charAt(0) !== '#' && line.indexOf('--') !== 0 && existsSync(line));
}

type Format = 'docx' | 'doc' | 'rtf';

/** What a file is, by how it begins: the extension says little in this corpus. */
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
  for (const format of ['docx', 'doc', 'rtf'] as const) out[format].sort();
  return (found = out);
}

/** Up to `count` of the files, spread evenly over the sorted list. */
function sample(files: string[], count: number): string[] {
  if (files.length <= count) return files;
  const out: string[] = [];
  for (let i = 0; i < count; i++) out.push(files[Math.floor((i * files.length) / count)]);
  return out;
}

const sameBytes = (a: Uint8Array | undefined, b: Uint8Array): boolean => !!a && a.length === b.length && a.every((value, i) => value === b[i]);

async function partsOf(pkg: OpcPackage): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) out.set(part, await pkg.readPart(part));
  return out;
}

const storiesOf = (doc: DocModel): { part: string; modified?: boolean }[] =>
  [doc.body, ...new Set([...doc.headers.values(), ...doc.footers.values()]), doc.footnotes, doc.endnotes, doc.comments].filter((story) => story !== undefined) as { part: string; modified?: boolean }[];

const rootOf = (block: Block): Wrapper | undefined => {
  let root = block.wrap;
  while (root && root.parent) root = root.parent;
  return root;
};

/** The source text a top-level block is copied from: its own, or that of the wrapper around it. */
const sourceOf = (doc: DocModel, block: Block): string | undefined => {
  const own = rawOf(doc.body, block);
  if (own !== undefined) return own;
  const root = rootOf(block);
  return root ? rawOf(doc.body, root) : undefined;
};

const skip = corpusDirs().length === 0 ? 'no local corpus (.corpus.local)' : false;

test('real docx files: nothing modified leaves the package alone; marked as modified without an edit gives the bytes of the file', { skip }, async (t) => {
  const failures: string[] = [];
  let documents = 0;
  let copied = 0;
  const files = corpus().docx;
  for (let n = 0; n < files.length; n++) {
    const bytes = new Uint8Array(readFileSync(files[n]));
    let pkg: OpcPackage;
    let doc: DocModel;
    try {
      pkg = await OpcPackage.open(bytes);
      doc = await readDocx(pkg);
    } catch {
      continue;
    }
    documents++;
    try {
      const before = await partsOf(pkg);
      const report = emptyWriteReport();
      await writeDocx(pkg, doc, { report });
      assert.deepEqual(report.parts, []);
      for (const story of storiesOf(doc)) story.modified = true;
      await writeDocx(pkg, doc, { report });
      assert.equal(report.writtenBlocks, 0);
      copied += report.copiedBlocks;
      const after = await partsOf(await OpcPackage.open(await pkg.save()));
      assert.equal(after.size, before.size);
      for (const [part, data] of before) assert.ok(sameBytes(after.get(part), data), 'a part changed');
    } catch (error) {
      failures.push(`docx #${n}: ${(error as Error).message.slice(0, 120)}`);
    }
  }
  t.diagnostic(`${documents} documents, ${copied} blocks copied`);
  assert.deepEqual(failures, []);
});

test('real docx files: every block written from the model reads back as the same model, in the order of the schema', { skip }, async (t) => {
  const failures: string[] = [];
  let documents = 0;
  let written = 0;
  const files = corpus().docx;
  for (let n = 0; n < files.length; n++) {
    const bytes = new Uint8Array(readFileSync(files[n]));
    let pkg: OpcPackage;
    let doc: DocModel;
    try {
      pkg = await OpcPackage.open(bytes);
      doc = await readDocx(pkg);
    } catch {
      continue;
    }
    documents++;
    try {
      const before = await partsOf(pkg);
      const picture = documentPicture(doc);
      const report = emptyWriteReport();
      await writeDocx(pkg, doc, { regenerate: true, report });
      written += report.writtenBlocks;
      const saved = await OpcPackage.open(await pkg.save());
      const difference = firstDifference(picture, documentPicture(await readDocx(saved)));
      assert.equal(difference === undefined ? undefined : differenceKind(difference), undefined);
      for (const part of report.parts) {
        // An order the file itself got wrong in a kept tree is not the writer's.
        const was = before.get(part);
        const known = was ? schemaOrderViolations(was) : [];
        assert.deepEqual(schemaOrderViolations(await saved.readPart(part)).filter((line) => known.indexOf(line) < 0), []);
      }
    } catch (error) {
      failures.push(`docx #${n}: ${(error as Error).message.slice(0, 160)}`);
    }
  }
  t.diagnostic(`${documents} documents, ${written} blocks written from the model`);
  assert.deepEqual(failures, []);
});

test('real docx files: random edits are in the file that is written; untouched blocks keep their source; undoing all gives the file back', { skip }, async (t) => {
  const failures: string[] = [];
  let documents = 0;
  let applied = 0;
  let kept = 0;
  const files = corpus().docx;
  for (let n = 0; n < files.length; n++) {
    const bytes = new Uint8Array(readFileSync(files[n]));
    let pkg: OpcPackage;
    let doc: DocModel;
    try {
      pkg = await OpcPackage.open(bytes);
      doc = await readDocx(pkg);
    } catch {
      continue;
    }
    documents++;
    try {
      const original = await partsOf(pkg);
      const history = new DocHistory<DocRange | undefined>(doc, 1000);
      const d = dice(n + 1);
      let done = 0;
      for (let step = 0; step < 60 && done < 20; step++) {
        let made;
        try {
          made = randomEdit(doc, d);
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
          continue;
        }
        if (!made || made.edit.ops.length === 0) continue;
        history.applyEdit(made.edit, undefined);
        history.seal();
        done++;
      }
      applied += done;
      // Wrappers by what they say: one that an edit split is two equal ones in the file.
      const edited = convertedPicture(doc, false);
      const sources = doc.body.blocks.map((block) => sourceOf(doc, block));
      await writeDocx(pkg, doc);
      const again = await readDocx(await OpcPackage.open(await pkg.save()));
      const difference = firstDifference(edited, convertedPicture(again, false));
      assert.equal(difference === undefined ? undefined : differenceKind(difference), undefined);
      assert.equal(again.body.blocks.length, sources.length);
      sources.forEach((source, i) => {
        if (source === undefined) return;
        kept++;
        assert.equal(sourceOf(again, again.body.blocks[i]) === source, true, 'an untouched block was not copied');
      });

      while (history.canUndo) history.undo();
      await writeDocx(pkg, doc);
      const after = await partsOf(await OpcPackage.open(await pkg.save()));
      assert.equal(after.size, original.size);
      for (const [part, data] of original) assert.ok(sameBytes(after.get(part), data), 'a part differs after undoing everything');
    } catch (error) {
      failures.push(`docx #${n}: ${(error as Error).message.slice(0, 160)}`);
    }
  }
  t.diagnostic(`${documents} documents, ${applied} edits, ${kept} untouched blocks compared`);
  assert.deepEqual(failures, []);
});

async function convert(format: 'doc' | 'rtf', count: number, t: { diagnostic(message: string): void }): Promise<void> {
  const failures: string[] = [];
  let documents = 0;
  const picked = sample(corpus()[format], count);
  for (let n = 0; n < picked.length; n++) {
    const bytes = new Uint8Array(readFileSync(picked[n]));
    let source: DocModel;
    try {
      source = format === 'doc' ? readDoc(CfbFile.read(bytes)) : readRtf(bytes);
    } catch {
      // Not a document the source reader takes: nothing to convert.
      continue;
    }
    documents++;
    try {
      const pkg = await OpcPackage.open(await buildDocx(source));
      const result = await readDocx(pkg);
      assert.deepEqual(result.warnings, []);
      // The separators made for notes that had none are not in the source.
      for (const list of ['footnotes', 'endnotes'] as const) {
        const made = result[list];
        if (!made) continue;
        for (const [id, note] of Array.from(made.items)) if (note.kind !== undefined && !(source[list] && source[list].items.has(id))) made.items.delete(id);
      }
      const difference = firstDifference(convertedPicture(source), convertedPicture(result));
      assert.equal(difference === undefined ? undefined : differenceKind(difference), undefined);
      for (const part of pkg.partNames()) if (/\.xml$/.test(part)) assert.deepEqual(schemaOrderViolations(await pkg.readPart(part)), []);
      for (const [target, media] of source.media || []) assert.ok(sameBytes(await pkg.readPart('/word/media/' + target.slice(target.lastIndexOf('/') + 1)), media.bytes), 'a picture changed');
    } catch (error) {
      failures.push(`${format} #${n}: ${(error as Error).message.slice(0, 160)}`);
    }
  }
  t.diagnostic(`${documents} documents converted`);
  assert.deepEqual(failures, []);
}

test('real binary doc files: the docx built from the model reads back as that model', { skip }, async (t) => {
  await convert('doc', 60, t);
});

test('real RTF files: the docx built from the model reads back as that model', { skip }, async (t) => {
  await convert('rtf', 60, t);
});
