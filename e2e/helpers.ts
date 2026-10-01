import { test as base, expect, type Page } from '@playwright/test';

/**
 * Every test gets a page whose console errors and uncaught exceptions are
 * recorded; the test fails afterwards if there were any.
 */
export const test = base.extend<{ pageProblems: string[] }>({
  pageProblems: [
    async ({ page }, use) => {
      const problems: string[] = [];
      page.on('console', (message) => {
        if (message.type() === 'error') problems.push(`console.error: ${message.text()} (${message.location().url})`);
      });
      page.on('pageerror', (error) => problems.push(`pageerror: ${error.name}: ${error.message}`));
      await use(problems);
      expect(problems, 'console errors and uncaught exceptions').toEqual([]);
    },
    { auto: true },
  ],
});

export { expect };

export type Flavour = 'script' | 'esm';

/** Loads a harness page and waits until the engine is available as window.__bym. */
export async function gotoHarness(page: Page, flavour: Flavour): Promise<void> {
  await page.goto(`/e2e/pages/${flavour}.html`);
  await page.waitForFunction(() => (window as any).__bym !== undefined);
  expect(await page.evaluate(() => (window as any).__bymFlavour)).toBe(flavour);
}

export interface RoundTrip {
  format: string;
  kind: string;
  container: string;
  parts: string[];
  before: Record<string, number[]>;
  saved: number[];
  reopenedFormat: string;
  reopenedParts: string[];
  after: Record<string, number[]>;
}

/** In the page: open -> read every part -> save -> open again -> read every part. */
export function roundTripInPage(page: Page, bytes: Uint8Array): Promise<RoundTrip> {
  return page.evaluate(async (input) => {
    const api = (window as any).__bym;
    const readAll = async (file: any) => {
      const out: Record<string, number[]> = {};
      for (const part of file.package.partNames()) out[part] = Array.from(await file.package.readPart(part));
      return out;
    };
    const file = await api.open(new Uint8Array(input));
    const before = await readAll(file);
    const saved: Uint8Array = await file.save();
    const reopened = await api.open(saved);
    return {
      format: file.format,
      kind: file.kind,
      container: file.container,
      parts: file.package.partNames(),
      before,
      saved: Array.from(saved),
      reopenedFormat: reopened.format,
      reopenedParts: reopened.package.partNames(),
      after: await readAll(reopened),
    };
  }, Array.from(bytes));
}

/** The page still runs script and still paints (a frame callback fires). */
export async function expectResponsive(page: Page): Promise<void> {
  const alive = await page.evaluate(
    () => new Promise<string>((resolve) => requestAnimationFrame(() => setTimeout(() => resolve('alive'), 0))),
  );
  expect(alive).toBe('alive');
}
