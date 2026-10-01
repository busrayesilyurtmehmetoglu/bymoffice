// Runs the engine over a directory of real Office files and reports what
// survives open -> save -> reopen.
// Usage: pnpm corpus <dir> [<dir> ...] [--allow-failures=N] [--min-passed=N]
// The two options turn a known baseline into the pass criterion.
//
// `pnpm corpus --local` takes its arguments from the file .corpus.local (one
// directory or option per line, "#" starts a comment). That file is not checked
// in: where a team keeps its real documents is nobody else's business. Without
// the file the run is skipped, so the full test run works on any machine.

import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { OpcPackage, UnsupportedFileError, open, parseXml, serializeXml, type OfficeFile } from '../src/index';

const EXTENSIONS = /\.(docx|docm|dotx|xlsx|xlsm|xltx|doc|xls)$/i;
const SKIP_DIRS = new Set(['node_modules', 'vendor', '.git']);

function* walk(dir: string): Generator<string> {
  for (const item of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, item.name);
    if (item.isDirectory()) {
      if (!SKIP_DIRS.has(item.name)) yield* walk(path);
    } else if (item.isFile() && EXTENSIONS.test(item.name) && !item.name.startsWith('~$')) yield path;
  }
}

let sheets = 0;
let cells = 0;

async function checkOoxml(file: OfficeFile, original: Uint8Array): Promise<void> {
  const pkg = file.package!;
  const parts = pkg.partNames();

  // Every XML part must survive our parser and serializer unchanged.
  for (const part of parts) {
    if (!/\.(xml|rels)$/i.test(part)) continue;
    const doc = await pkg.readXml(part);
    if (!isDeepStrictEqual(parseXml(serializeXml(doc)), doc)) throw new Error(`XML round-trip changed ${part}`);
  }

  // Plain save: same parts, same order, same bytes.
  const savedBytes = await file.save();
  const saved = await OpcPackage.open(savedBytes);
  if (!isDeepStrictEqual(saved.partNames(), parts)) throw new Error('part list changed after save');
  const reference = await OpcPackage.open(original);
  for (const part of parts) {
    if (!isDeepStrictEqual(await saved.readPart(part), await reference.readPart(part))) {
      throw new Error(`content changed after save: ${part}`);
    }
  }

  // Spreadsheets must read into the model; reading must not disturb the package.
  if (file.format === 'xlsx') {
    const workbook = await file.workbook();
    for (const sheet of workbook.sheets) {
      cells += [...sheet.rows.values()].reduce((sum, row) => sum + row.cells.filter(() => true).length, 0);
    }
    sheets += workbook.sheets.length;
    if (!isDeepStrictEqual(await file.save(), savedBytes)) throw new Error('saved bytes changed after reading the model');
  }

  // Rewriting the main part through the XML layer must keep its meaning.
  const main = await pkg.mainPart();
  if (!main) throw new Error('no main document part');
  const before = await pkg.readXml(main);
  pkg.writePart(main, before);
  const rewritten = await OpcPackage.open(await pkg.save());
  if (!isDeepStrictEqual(await rewritten.readXml(main), before)) throw new Error('main part changed after rewrite');
}

function checkCfb(file: OfficeFile): void {
  for (const entry of file.cfb!.entries) {
    if (entry.type !== 'stream') continue;
    const size = file.cfb!.stream(entry.path).length;
    if (size !== entry.size) throw new Error(`stream ${entry.path}: read ${size} of ${entry.size} bytes`);
  }
}

// BYM_NO_NATIVE=1 hides the built-in compression streams, so the whole corpus
// goes through our own inflate/deflate as it would on an old WebView.
if (process.env.BYM_NO_NATIVE) {
  const scope = globalThis as Record<string, unknown>;
  scope.CompressionStream = undefined;
  scope.DecompressionStream = undefined;
  console.log('built-in compression streams disabled');
}

const LOCAL_CONFIG = '.corpus.local';
let args = process.argv.slice(2);
if (args.includes('--local')) {
  if (!existsSync(LOCAL_CONFIG)) {
    console.log(`corpus check skipped: ${LOCAL_CONFIG} not found (see docs/kullanim-kilavuzu.md)`);
    process.exit(0);
  }
  const lines = readFileSync(LOCAL_CONFIG, 'utf8').split('\n');
  args = lines.map((line) => line.trim()).filter((line) => line !== '' && !line.startsWith('#'));
}
const roots = args.filter((arg) => !arg.startsWith('--'));
const option = (name: string, fallback: number) => {
  const arg = args.find((candidate) => candidate.startsWith(`--${name}=`));
  return arg ? Number(arg.slice(name.length + 3)) : fallback;
};
const allowedFailures = option('allow-failures', 0);
const minPassed = option('min-passed', 0);
if (roots.length === 0) {
  console.error('usage: pnpm corpus <dir> [<dir> ...]');
  process.exit(2);
}

const passed: Record<string, number> = {};
const failures: string[] = [];
let unsupported = 0;
const started = performance.now();

for (const root of roots) {
  for (const path of walk(root)) {
    try {
      const bytes = new Uint8Array(readFileSync(path));
      const file = await open(bytes);
      if (file.container === 'ooxml') await checkOoxml(file, bytes);
      else checkCfb(file);
      passed[file.format] = (passed[file.format] ?? 0) + 1;
    } catch (error) {
      if (error instanceof UnsupportedFileError) unsupported++;
      else failures.push(`${path}\n    ${(error as Error).message}`);
    }
  }
}

const total = Object.values(passed).reduce((sum, n) => sum + n, 0);
console.log(`passed: ${total}`, passed);
console.log(`xlsx model: ${sheets} sheets, ${cells} cells`);
console.log(`not a Word/Excel file: ${unsupported}`);
console.log(`failed: ${failures.length}`);
for (const failure of failures.slice(0, 40)) console.log(`  ${failure}`);
if (failures.length > 40) console.log(`  ... and ${failures.length - 40} more`);
console.log(`time: ${((performance.now() - started) / 1000).toFixed(1)}s`);
if (total < minPassed) console.log(`REGRESSION: ${total} passed, at least ${minPassed} expected`);
process.exit(failures.length > allowedFailures || total < minPassed ? 1 : 0);
