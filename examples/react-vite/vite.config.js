import { cpSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

/**
 * The font the package brings (TeX Gyre Bonum, for documents set in Bookman Old Style) is not part of the
 * script a bundler makes: it is a folder of the package, node_modules/bymoffice/dist/fonts/. This copies it
 * into public/bymoffice-fonts/ whenever Vite starts, so that the dev server serves it and the build copies it
 * into dist/. The editor is told where it is with `fontsUrl` (see src/OfficeEditor.jsx).
 */
function bymofficeFonts() {
  const fonts = join(dirname(createRequire(import.meta.url).resolve('bymoffice/package.json')), 'dist', 'fonts');
  return {
    name: 'bymoffice-fonts',
    configResolved(config) {
      if (config.publicDir) cpSync(fonts, join(config.publicDir, 'bymoffice-fonts'), { recursive: true });
    },
  };
}

export default defineConfig({
  plugins: [react(), bymofficeFonts()],
});
