import { corruptDataXlsx, corruptDirectoryXlsx, emptyFile, textFile, truncatedXlsx } from './fixtures/files';
import { expect, expectResponsive, gotoHarness, test } from './helpers';

/** open() must settle well inside this budget, even for garbage. */
const OPEN_BUDGET_MS = 5_000;

const CASES: Array<[string, () => Uint8Array | Promise<Uint8Array>]> = [
  ['truncated xlsx (first half)', truncatedXlsx],
  ['xlsx with corrupted compressed data', corruptDataXlsx],
  ['xlsx with corrupted central directory record', corruptDirectoryXlsx],
  ['empty file', emptyFile],
  ['plain text file', textFile],
];

for (const [name, make] of CASES) {
  test(`robustness: ${name} is rejected with an error and the page stays responsive`, async ({ page }) => {
    const input = await make();
    await gotoHarness(page, 'script');

    const outcome = await page.evaluate(
      async ({ bytes, budget }) => {
        const started = performance.now();
        const attempt = (window as any).__bym.open(new Uint8Array(bytes)).then(
          () => ({ state: 'opened' as const, name: '', message: '' }),
          (error: unknown) => ({
            state: 'rejected' as const,
            name: error instanceof Error ? error.name : typeof error,
            message: error instanceof Error ? error.message : String(error),
          }),
        );
        const timeout = new Promise<{ state: 'timeout'; name: string; message: string }>((resolve) =>
          setTimeout(() => resolve({ state: 'timeout', name: '', message: '' }), budget),
        );
        const result = await Promise.race([attempt, timeout]);
        return { ...result, elapsed: performance.now() - started };
      },
      { bytes: Array.from(input), budget: OPEN_BUDGET_MS },
    );

    expect(outcome.state, `open() outcome after ${Math.round(outcome.elapsed)} ms`).toBe('rejected');
    expect(outcome.elapsed).toBeLessThan(OPEN_BUDGET_MS);
    expect(outcome.message.length, 'error carries a message').toBeGreaterThan(0);
    await expectResponsive(page);
  });
}
