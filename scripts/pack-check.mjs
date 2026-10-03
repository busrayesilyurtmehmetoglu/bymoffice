// Checks the package as someone who installs it gets it: packs it, installs
// the archive into an empty project (tmp-pack/, not checked in), and uses it
// from there as an ES module in Node, through its types with both ways
// TypeScript resolves modules, and as the script a page loads from a CDN.
// Needs `pnpm build` first. Nothing is fetched: the package has no
// dependencies.

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { Script, createContext } from 'node:vm';

const root = resolve('.');
const work = join(root, 'tmp-pack');
rmSync(work, { recursive: true, force: true });
mkdirSync(join(work, 'consumer'), { recursive: true });

const run = (command, args, cwd) => execFileSync(command, args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
const fail = (message) => {
  console.error(`pack-check: ${message}`);
  process.exit(1);
};

for (const file of ['dist/bymoffice.min.js', 'dist/bymoffice.esm.js', 'dist/types/index.d.ts', 'dist/bymoffice.webview.html']) {
  if (!existsSync(join(root, file))) fail(`${file} is missing; run pnpm build first`);
}

// 1. What goes into the archive.
const packed = JSON.parse(run('npm', ['pack', '--json', '--ignore-scripts', '--pack-destination', work], root))[0];
const names = packed.files.map((file) => file.path);
for (const wanted of ['package.json', 'README.md', 'CHANGELOG.md', 'dist/bymoffice.min.js', 'dist/bymoffice.esm.js', 'dist/bymoffice.webview.html', 'dist/types/index.d.ts']) {
  if (!names.includes(wanted)) fail(`the archive lacks ${wanted}`);
}
const stray = names.filter((name) => !/^(package\.json|README\.md|CHANGELOG\.md|LICENSE[^/]*|dist\/.+)$/.test(name));
if (stray.length) fail(`the archive holds files that are not part of the package: ${stray.slice(0, 5).join(', ')}`);
const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
if (manifest.dependencies && Object.keys(manifest.dependencies).length) fail('the package must have no runtime dependencies');

// 2. Installed into an empty project.
const consumer = join(work, 'consumer');
writeFileSync(join(consumer, 'package.json'), JSON.stringify({ name: 'consumer', private: true, type: 'module' }));
run('npm', ['install', '--offline', '--no-audit', '--no-fund', '--ignore-scripts', '--no-package-lock', join(work, packed.filename)], consumer);

// 3. As an ES module in Node: a workbook is built, written, opened again and read.
writeFileSync(
  join(consumer, 'use.mjs'),
  `import * as api from 'bymoffice';
import { open } from 'bymoffice';
for (const name of ['open', 'mount', 'buildXlsx', 'buildDocx', 'writeXlsx', 'writeDocx', 'DocumentEditor', 'SpreadsheetEditor', 'UnsupportedFileError']) {
  if (typeof api[name] !== 'function') throw new Error('missing export: ' + name);
}
const rtf = await open(new TextEncoder().encode('{\\\\rtf1\\\\ansi Hello from the package\\\\par}'));
if (rtf.format !== 'rtf') throw new Error('format: ' + rtf.format);
const docx = await open(await rtf.saveAsDocx());
const doc = await docx.document();
const text = api.paragraphText(doc.body.blocks[0]);
if (text !== 'Hello from the package') throw new Error('text: ' + text);
console.log('esm ok');
`,
);
if (!run('node', ['use.mjs'], consumer).includes('esm ok')) fail('the ES module did not run');

// 4. The types, resolved the way a bundler does and the way Node does.
writeFileSync(
  join(consumer, 'use.ts'),
  `import { open, mount, type OfficeFile, type MountedView, type Workbook, type DocModel } from 'bymoffice';
export async function read(bytes: Uint8Array): Promise<string> {
  const file: OfficeFile = await open(bytes);
  if (file.kind === 'spreadsheet') {
    const workbook: Workbook = await file.workbook();
    return workbook.sheets.map((sheet) => sheet.name).join(',');
  }
  const doc: DocModel = await file.document!();
  return String(doc.body.blocks.length);
}
export async function show(host: HTMLElement, bytes: Uint8Array): Promise<MountedView> {
  const mounted = await mount(host, { file: bytes, locale: 'tr' });
  if (mounted.kind === 'document') mounted.view.goToPage(1);
  else mounted.view.focus();
  return mounted;
}
`,
);
const tsc = join(root, 'node_modules', '.bin', 'tsc');
for (const [module, resolution] of [
  ['esnext', 'bundler'],
  ['node16', 'node16'],
  ['nodenext', 'nodenext'],
]) {
  try {
    run(tsc, ['--noEmit', '--strict', '--target', 'es2022', '--lib', 'es2022,dom', '--module', module, '--moduleResolution', resolution, 'use.ts'], consumer);
  } catch (error) {
    fail(`the types do not resolve with moduleResolution ${resolution}:\n${error.stdout || error.message}`);
  }
}

// 5. As the script of a page: one global, nothing else needed to load it.
{
  // What every browser has and the script uses while it loads.
  const context = createContext({ TextDecoder, TextEncoder, setTimeout, clearTimeout });
  try {
    new Script(readFileSync(join(consumer, 'node_modules', 'bymoffice', manifest.unpkg), 'utf8'), { filename: manifest.unpkg }).runInContext(context);
  } catch (error) {
    fail(`the script does not load: ${error && error.message}`);
  }
  const api = context.BymOffice;
  if (!api || typeof api.open !== 'function' || typeof api.mount !== 'function') fail('the script does not define the global BymOffice');
}
if (manifest.unpkg !== manifest.jsdelivr || !names.includes(manifest.unpkg)) fail('unpkg and jsdelivr must name the script that is in the archive');

console.log(`pack-check: ${packed.filename}, ${names.length} files, ${(packed.size / 1024).toFixed(0)} KB packed, ${(packed.unpackedSize / 1024).toFixed(0)} KB unpacked; ESM, types (bundler, node16, nodenext) and the CDN script work`);
