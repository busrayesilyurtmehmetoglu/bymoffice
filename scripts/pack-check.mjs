// Checks the package as someone who installs it gets it: packs it, installs
// the archive into an empty project (tmp-pack/, not checked in), and uses it
// from there as an ES module in Node, through its types with both ways
// TypeScript resolves modules, and as the script a page loads from a CDN.
// Needs `pnpm build` first. Nothing is fetched: the package has no
// dependencies.

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
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
// server/: the example relay for collaborative editing (two files, no dependencies).
const stray = names.filter((name) => !/^(package\.json|README\.md|CHANGELOG\.md|THIRD-PARTY-NOTICES\.md|LICENSE[^/]*|dist\/.+|server\/collab-server\.(mjs|d\.mts))$/.test(name));
if (stray.length) fail(`the archive holds files that are not part of the package: ${stray.slice(0, 5).join(', ')}`);
const manifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
// No source maps: neither as files nor named by the scripts (a CDN would answer that name with a 404).
const maps = names.filter((name) => /\.map$/.test(name));
if (maps.length) fail(`the archive holds source maps: ${maps.join(', ')}`);
for (const file of ['dist/bymoffice.min.js', 'dist/bymoffice.esm.js', 'dist/bymoffice.webview.html']) {
  if (/sourceMappingURL=/.test(readFileSync(join(root, file), 'utf8').slice(-4096))) fail(`${file} names a source map`);
}
if (manifest.private) fail('package.json is private: it cannot be published');

// The font of the package: its four files and their licence, each as large as its source, and the notice that
// names them. (That they are the same bytes is looked at below, in the installed copy.)
const fontSource = join(root, 'assets', 'fonts', 'tex-gyre-bonum');
const fontFiles = readdirSync(fontSource).sort();
if (fontFiles.filter((name) => /\.otf$/.test(name)).length !== 4 || !fontFiles.includes('GUST-FONT-LICENSE.txt')) fail('assets/fonts/tex-gyre-bonum must hold four font files and their licence');
let fontBytes = 0;
for (const name of fontFiles) {
  const entry = packed.files.find((file) => file.path === `dist/fonts/tex-gyre-bonum/${name}`);
  if (!entry) fail(`the archive lacks dist/fonts/tex-gyre-bonum/${name}`);
  const size = readFileSync(join(fontSource, name)).length;
  if (entry.size !== size) fail(`dist/fonts/tex-gyre-bonum/${name} is ${entry.size} bytes in the archive and ${size} in assets/`);
  fontBytes += size;
}
const fontEntries = names.filter((name) => name.startsWith('dist/fonts/'));
if (fontEntries.length !== fontFiles.length) fail(`dist/fonts holds files that are not in assets/: ${fontEntries.join(', ')}`);
if (!names.includes('THIRD-PARTY-NOTICES.md')) fail('the archive lacks THIRD-PARTY-NOTICES.md');
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
for (const name of ['open', 'mount', 'createBlank', 'buildXlsx', 'buildDocx', 'writeXlsx', 'writeDocx', 'DocumentEditor', 'SpreadsheetEditor', 'UnsupportedFileError']) {
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

// 3b. The fonts as they were installed: the same bytes as in assets/, reachable under "bymoffice/fonts/…" and next to
// the scripts, where the engine looks for them.
writeFileSync(
  join(consumer, 'fonts.mjs'),
  `import { fileURLToPath } from 'node:url';
import * as api from 'bymoffice';
const out = {};
for (const face of api.EMBEDDED_FONT_FACES) out[face.file] = fileURLToPath(import.meta.resolve('bymoffice/fonts/tex-gyre-bonum/' + face.file));
out['GUST-FONT-LICENSE.txt'] = fileURLToPath(import.meta.resolve('bymoffice/fonts/tex-gyre-bonum/GUST-FONT-LICENSE.txt'));
if (typeof api.fonts.configure !== 'function' || api.fonts.embedded().family !== 'TeX Gyre Bonum') throw new Error('fonts.configure / fonts.embedded');
console.log(JSON.stringify(out));
`,
);
{
  const installed = JSON.parse(run('node', ['fonts.mjs'], consumer));
  if (Object.keys(installed).sort().join() !== fontFiles.join()) fail(`the package names other font files than assets/ holds: ${Object.keys(installed).join(', ')}`);
  for (const name of fontFiles) {
    if (!readFileSync(installed[name]).equals(readFileSync(join(fontSource, name)))) fail(`${name} of the installed package differs from assets/`);
    const beside = join(consumer, 'node_modules', 'bymoffice', 'dist', 'fonts', 'tex-gyre-bonum', name);
    if (resolve(installed[name]) !== resolve(beside)) fail(`bymoffice/fonts/tex-gyre-bonum/${name} is not the file next to the scripts`);
  }
}

// 4. The types, resolved the way a bundler does and the way Node does.
writeFileSync(
  join(consumer, 'use.ts'),
  `import { open, mount, collab, readCollabStamp, writeSharedFile, CollabSession, contentVersion, type CollabSessionHost, type CollabSessionOptions, type CollabStamp, type SharedFile, type OfficeFile, type MountedView, type Workbook, type DocModel } from 'bymoffice';
// A session without mount(): the file is read with the step of its room and written with it.
export async function shared(bytes: Uint8Array): Promise<SharedFile> {
  const file: OfficeFile = await open(bytes);
  const workbook: Workbook = await file.workbook();
  const stamp: CollabStamp | undefined = file.package ? await readCollabStamp(file.package) : undefined;
  const host: CollabSessionHost = { workbook: () => workbook, changed: () => undefined, notice: () => undefined };
  const options: CollabSessionOptions = { transport: collab.memoryServer().connect(), versions: [contentVersion(bytes)], stamp };
  const session = new CollabSession(host, options);
  const saved = await collab.writeSharedFile(session, file.package, () => file.save());
  session.announceSaved([contentVersion(saved.bytes)], saved.rev);
  return writeSharedFile(session, file.package, () => file.save());
}
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

const kb = (bytes) => `${(bytes / 1024).toFixed(0)} KB`;
console.log(
  `pack-check: ${packed.filename}, ${names.length} files, ${kb(packed.size)} packed, ${kb(packed.unpackedSize)} unpacked (of which the fonts: ${fontFiles.length} files, ${kb(fontBytes)}); ` +
    'ESM, types (bundler, node16, nodenext), the CDN script and the font files work',
);
