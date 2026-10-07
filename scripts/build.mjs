import { createHash } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { Script } from 'node:vm';
import { build } from 'esbuild';

const shared = {
  entryPoints: ['src/index.ts'],
  bundle: true,
  // Old WebViews (iOS before 16.4, Android 8-era) must parse the bundle; newer syntax is lowered.
  target: 'es2019',
  sourcemap: true,
};

// BYM_OUT_DIR: where the scripts, the bridge page and the fonts go instead of dist/ (the tests of the build use it).
const OUT = process.env.BYM_OUT_DIR || 'dist';

await Promise.all([
  // <script src=".../bymoffice.min.js"> exposes a global `BymOffice`.
  build({ ...shared, format: 'iife', globalName: 'BymOffice', minify: true, outfile: `${OUT}/bymoffice.min.js` }),
  build({ ...shared, format: 'esm', outfile: `${OUT}/bymoffice.esm.js` }),
]);

// The font of the package: TeX Gyre Bonum, as GUST e-foundry distributes it (CTAN, fonts/tex-gyre). The files are
// shipped unchanged and under their own names, with their licence: they are only ever copied here, and the copy is
// compared with what the files are known to be.
const FONT_SOURCE = 'assets/fonts/tex-gyre-bonum';
const FONT_TARGET = `${OUT}/fonts/tex-gyre-bonum`;
const FONT_LICENSE = 'GUST-FONT-LICENSE.txt';
const FONT_FILES = {
  'texgyrebonum-regular.otf': '8ad5a53ae61d710fe1773f4ecc7fdf6dde7df1df59faa4492a24def95077d6b3',
  'texgyrebonum-bold.otf': '173710158270252ec1ddd5e69a07fc71853680c7188d67a84d794a465a4dff81',
  'texgyrebonum-italic.otf': '518836a0e5bd71f644ae131185be53b04fadb30ef7dbdc483deb5f69bd1fe47d',
  'texgyrebonum-bolditalic.otf': 'bd037941d21149ea0d650286537a3266fadfb03f91b4a5a78848ea5781befac9',
  [FONT_LICENSE]: '2bd69affc3da00715116f713f57eab9707e96daf3562ad0215987b15b9c16f73',
};
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
/** The bytes of the font files, by name, as they were copied. */
const fontBytes = new Map();
await mkdir(FONT_TARGET, { recursive: true });
for (const [name, expected] of Object.entries(FONT_FILES)) {
  const bytes = await readFile(`${FONT_SOURCE}/${name}`);
  if (sha256(bytes) !== expected) throw new Error(`${FONT_SOURCE}/${name} is not the file it is known to be: the fonts are shipped unchanged`);
  await copyFile(`${FONT_SOURCE}/${name}`, `${FONT_TARGET}/${name}`);
  if (sha256(await readFile(`${FONT_TARGET}/${name}`)) !== expected) throw new Error(`${FONT_TARGET}/${name} differs from its source`);
  fontBytes.set(name, bytes);
}

// dist/bymoffice.webview.html: the bridge page for WebViews, with the engine
// inside it, so that an application ships one file and needs no network.
{
  const template = await readFile('src/embed/host.html', 'utf8');
  const { version } = JSON.parse(await readFile('package.json', 'utf8'));
  const engine = (await readFile(`${OUT}/bymoffice.min.js`, 'utf8'))
    // The source map is not shipped with the page.
    .replace(/\n\/\/# sourceMappingURL=.*\s*$/, '\n')
    // Inside an inline script these would end the element or open an HTML comment.
    .replace(/<\/script/gi, '<\\/script')
    .replace(/<!--/g, '<\\!--');
  // The escapes are harmless inside strings; anywhere else they would break the code, and this notices.
  new Script(engine, { filename: 'bymoffice.webview.html (engine)' });

  // Replacer functions: the engine's text contains "$&" and the like.
  const fill = (text, marker, value) => {
    if (text.split(marker).length !== 2) throw new Error(`src/embed/host.html must contain ${marker} exactly once`);
    return text.replace(marker, () => value);
  };
  let page = fill(fill(template, '__BYM_ENGINE__', engine), '__BYM_VERSION__', version);

  // BYM_PARENT_ORIGINS="https://app.example.com https://admin.example.com": the pages that may
  // embed this one in an iframe. Not set: none, which is right for a page shipped inside an application.
  const origins = (process.env.BYM_PARENT_ORIGINS ?? '').split(/\s+/).filter(Boolean).map((text) => {
    let url;
    try {
      url = new URL(text);
    } catch {
      throw new Error(`BYM_PARENT_ORIGINS: "${text}" is not an origin`);
    }
    if ((url.protocol !== 'http:' && url.protocol !== 'https:') || url.origin === 'null') throw new Error(`BYM_PARENT_ORIGINS: "${text}" is not an http(s) origin`);
    return url.origin;
  });
  page = fill(page, '__BYM_PARENT_ORIGINS__', origins.join(' '));

  // The fonts of the package inside the page, which works without a network: each file as base64 in a block of
  // data (a script element of a type no browser runs; the policy below has nothing to say about it), which the
  // script of the page hands to the engine when a file asks for the font. With them goes their licence, which asks
  // for a copy of itself, or a pointer to one, wherever the files go. BYM_NO_FONTS=1 builds the page without them
  // (about 770 KB less): Bookman Old Style is then shown with what the device has.
  let fonts = '';
  if (process.env.BYM_NO_FONTS !== '1') {
    const license = fontBytes.get(FONT_LICENSE).toString('utf8').trim();
    if (license.includes('--')) throw new Error('The licence of the fonts cannot be put into an HTML comment');
    fonts =
      '<!--\n      The font files below are TeX Gyre Bonum (regular, bold, italic, bold italic), by B. Jackowski and J. M. Nowacki,\n' +
      '      GUST e-foundry, https://www.gust.org.pl/projects/e-foundry/tex-gyre (CTAN: fonts/tex-gyre). They are included\n' +
      '      unchanged (each block decodes to the very bytes of the file of that name), under the GUST Font License. They\n' +
      '      are a part of the TeX Gyre Bonum distribution; the complete distribution is at the address above. Its authors\n' +
      '      have no part in this page and do not support it. The licence:\n\n' +
      license +
      '\n    -->\n';
    for (const name of Object.keys(FONT_FILES)) {
      if (name === FONT_LICENSE) continue;
      fonts += `    <script type="text/plain" data-bym-font="${name}">${fontBytes.get(name).toString('base64')}</script>\n`;
    }
    fonts = fonts.trimEnd();
  }
  page = fill(page, '__BYM_FONTS__', fonts);

  // Only the two scripts of the page itself may run: each is allowed by the hash of its text.
  const hashes = [];
  for (const match of page.matchAll(/<script>([\s\S]*?)<\/script>/g)) {
    hashes.push(`'sha256-${createHash('sha256').update(match[1], 'utf8').digest('base64')}'`);
  }
  if (hashes.length !== 2) throw new Error(`Expected 2 inline scripts in the bridge page, found ${hashes.length}`);
  const policy = [
    "default-src 'none'",
    `script-src ${hashes.join(' ')}`,
    // The view adds its own style element, whose text is not known here.
    "style-src 'unsafe-inline'",
    // blob: the pictures of a document, made by the page itself from the bytes of the file.
    'img-src data: blob:',
    // No remote hosts: a `load` with a URL reaches the page's own origin, blob: and data: only.
    "connect-src 'self' blob: data:",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; ');
  page = fill(page, '__BYM_CSP__', policy);
  await writeFile(`${OUT}/bymoffice.webview.html`, page);
}
