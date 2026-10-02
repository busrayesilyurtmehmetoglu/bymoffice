// Runs the engine over a directory of real Office files and reports what
// survives open -> save -> reopen.
// Usage: pnpm corpus <dir> [<dir> ...] [--allow-failures=N] [--min-passed=N]
// The two options turn a known baseline into the pass criterion.
//
// `pnpm corpus --local` takes its arguments from the file .corpus.local (one
// directory or option per line, "#" starts a comment). That file is not checked
// in: where a team keeps its real documents is nobody else's business. Without
// the file the run is skipped, so the full test run works on any machine.

import { readDoc } from '../src/formats/doc/document';
import { readXls } from '../src/formats/xls/workbook';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { OpcPackage, UnsupportedFileError, applyOp, open, parseXml, serializeXml, type OfficeFile, type Workbook } from '../src/index';

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
let docFiles = 0;
let docBlocksBinary = 0;
let xlsBooks = 0;
let xlsSheets = 0;
let xlsCells = 0;
let documents = 0;
let docBlocks = 0;
let edited = 0;
let restructured = 0;

/** What a workbook holds, as plain data that survives a save: used to compare a model with its reopened copy. */
function picture(workbook: Workbook): string {
  return JSON.stringify({
    names: workbook.definedNames,
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      state: sheet.state,
      merges: sheet.merges,
      links: sheet.hyperlinks.map((link) => [link.range, link.target, link.location, link.tooltip, link.display]),
      columns: sheet.columns,
      rows: [...sheet.rows.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([index, row]) => [
          index,
          row.height,
          row.hidden,
          row.style,
          row.attrs,
          row.cells
            .map((cell, col) => [col, cell.kind, cell.value, cell.style, cell.formula && { ...cell.formula, attrs: cell.formula.attrs && [...cell.formula.attrs] }, cell.runs, cell.cm, cell.vm])
            // A blank cell without a format or a formula says nothing; the writer leaves it out.
            .filter((entry) => !(entry[1] === 'blank' && entry[3] === 0 && !entry[4])),
        ])
        // Likewise a row with neither cells nor properties.
        .filter((entry) => (entry[5] as unknown[]).length > 0 || entry[1] !== undefined || entry[2] || entry[3] !== undefined || entry[4] !== undefined),
    })),
  });
}

/** Edits a real workbook the way a user would and checks that the saved file reads back as the edited model. */
async function checkEditing(original: Uint8Array): Promise<void> {
  const file = await open(original);
  const workbook = await file.workbook();
  const index = workbook.sheets.findIndex((sheet) => sheet.kind === 'worksheet');
  if (index < 0) return;
  const sheet = workbook.sheets[index];

  // A new cell below the data, and the first number in the sheet changed.
  const changes = [{ row: sheet.rowCount + 2, col: 0, cell: { kind: 'number' as const, value: 12345.678, style: 0 } }];
  outer: for (const [rowIndex, row] of sheet.rows) {
    for (let col = 0; col < row.cells.length; col++) {
      const cell = row.cells[col];
      if (cell && cell.kind === 'number' && !cell.formula) {
        changes.push({ row: rowIndex, col, cell: { ...cell, value: (cell.value as number) + 1 } });
        break outer;
      }
    }
  }
  applyOp(workbook, { type: 'setCells', sheet: index, cells: changes });
  const saved = await file.save();
  const reopened = await open(saved);
  if (picture(await reopened.workbook()) !== picture(workbook)) throw new Error('edited workbook reads back differently');

  // Parts the edit has no business with keep their bytes.
  const reference = await OpcPackage.open(original);
  const touched = new Set([sheet.part.toLowerCase(), '/xl/workbook.xml', '/xl/_rels/workbook.xml.rels', '/xl/calcchain.xml']);
  for (const part of reference.partNames()) {
    if (touched.has(part.toLowerCase())) continue;
    if (!reopened.package!.hasPart(part)) throw new Error(`part lost after an edit: ${part}`);
    if (!isDeepStrictEqual(await reopened.package!.readPart(part), await reference.readPart(part))) throw new Error(`untouched part changed after an edit: ${part}`);
  }
  edited++;

  // Rows inserted and deleted, where the sheet allows it (tables and pivot tables do not yet).
  try {
    applyOp(workbook, { type: 'insert', sheet: index, axis: 'row', at: 1, count: 2 });
    applyOp(workbook, { type: 'delete', sheet: index, axis: 'col', at: 1, count: 1 });
  } catch (error) {
    // Array formulas and full sheets refuse structural edits; that is not a failure of the writer.
    if (error instanceof RangeError) return;
    throw error;
  }
  const again = await open(await file.save());
  if (picture(await again.workbook()) !== picture(workbook)) throw new Error('restructured workbook reads back differently');
  restructured++;
}

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
    await checkEditing(original);
  }

  // Documents must read into the model; reading must not disturb the package.
  if (file.format === 'docx' && file.document) {
    const model = await file.document();
    documents++;
    docBlocks += model.body.blocks.length;
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
  // Old documents must read into the model.
  if (file.format === 'doc') {
    const model = readDoc(file.cfb!);
    docFiles++;
    docBlocksBinary += model.body.blocks.length;
  }
  // Old spreadsheets must read into the model.
  if (file.format === 'xls') {
    const workbook = readXls(file.cfb!);
    xlsBooks++;
    xlsSheets += workbook.sheets.length;
    for (const sheet of workbook.sheets) for (const row of sheet.rows.values()) xlsCells += row.cells.filter(() => true).length;
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
console.log(`doc model: ${docFiles} documents, ${docBlocksBinary} blocks`);
console.log(`xls model: ${xlsBooks} workbooks, ${xlsSheets} sheets, ${xlsCells} cells`);
console.log(`docx model: ${documents} documents, ${docBlocks} blocks`);
console.log(`xlsx model: ${sheets} sheets, ${cells} cells; edited and saved: ${edited}; restructured and saved: ${restructured}`);
console.log(`not a Word/Excel file: ${unsupported}`);
console.log(`failed: ${failures.length}`);
for (const failure of failures.slice(0, 40)) console.log(`  ${failure}`);
if (failures.length > 40) console.log(`  ... and ${failures.length - 40} more`);
console.log(`time: ${((performance.now() - started) / 1000).toFixed(1)}s`);
if (total < minPassed) console.log(`REGRESSION: ${total} passed, at least ${minPassed} expected`);
process.exit(failures.length > allowedFailures || total < minPassed ? 1 : 0);
