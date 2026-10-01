import assert from 'node:assert/strict';
import { test } from 'node:test';
import vm from 'node:vm';
import { build } from 'esbuild';

// Old WebViews have neither `globalThis` nor compression streams. The bundle is
// built the way scripts/build.mjs builds it and run in a context shaped like one.
test('bundle: reads and writes compressed archives in an engine without globalThis', async () => {
  const result = await build({
    entryPoints: ['src/index.ts'],
    bundle: true,
    target: 'es2019',
    format: 'iife',
    globalName: 'BymOffice',
    write: false,
  });
  const context = vm.createContext({ TextEncoder, TextDecoder, Uint8Array });
  vm.runInContext('delete globalThis.globalThis; var self = this;', context);
  assert.equal(vm.runInContext('typeof globalThis', context), 'undefined');
  vm.runInContext(result.outputFiles[0].text, context);

  const text = await vm.runInContext(
    `(async () => {
      var data = new TextEncoder().encode('şğü repeated text '.repeat(500));
      var zip = new BymOffice.ZipArchive();
      zip.set('a.txt', data);
      var again = BymOffice.ZipArchive.read(await zip.write());
      if (again.entry('a.txt').method !== 8) throw new Error('entry was not deflated');
      return new TextDecoder().decode(await again.read('a.txt'));
    })()`,
    context,
  );
  assert.equal(text, 'şğü repeated text '.repeat(500));
});
