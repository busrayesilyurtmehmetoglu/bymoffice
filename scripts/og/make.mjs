// Makes the share picture of the demo site (site/assets/og.png, 1200 x 630) from card.html.
// Run by hand when the card changes: node scripts/og/make.mjs (the PNG is kept in the repository).
import { chromium } from '@playwright/test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const card = pathToFileURL(fileURLToPath(new URL('card.html', import.meta.url))).href;
const out = fileURLToPath(new URL('../../site/assets/og.png', import.meta.url));
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 1200, height: 630 }, deviceScaleFactor: 1 });
  await page.goto(card);
  await page.screenshot({ path: out, type: 'png' });
} finally {
  await browser.close();
}
console.log('site/assets/og.png');
