import { build } from 'esbuild';

const shared = {
  entryPoints: ['src/index.ts'],
  bundle: true,
  // Old WebViews (iOS before 16.4, Android 8-era) must parse the bundle; newer syntax is lowered.
  target: 'es2019',
  sourcemap: true,
};

await Promise.all([
  // <script src=".../bymoffice.min.js"> exposes a global `BymOffice`.
  build({ ...shared, format: 'iife', globalName: 'BymOffice', minify: true, outfile: 'dist/bymoffice.min.js' }),
  build({ ...shared, format: 'esm', outfile: 'dist/bymoffice.esm.js' }),
]);
