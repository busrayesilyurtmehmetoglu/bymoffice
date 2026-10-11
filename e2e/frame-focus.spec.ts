import type { Frame, Page } from '@playwright/test';
import { createBlank } from '../src/index';
import { expect, test } from './helpers';

// The editor inside a frame of a page that holds the focus itself: the first
// thing the user does is press a cell (or the page of a document) and type.
// Firefox does not move the focus into a frame that is not the active one
// while the press is handled; what was typed went nowhere.

const BRIDGE = '/dist/bymoffice.webview.html';

const base64 = (bytes: Uint8Array): string => Buffer.from(bytes).toString('base64');

/** A page of the same origin with the bridge page in a frame and a button of its own, and the file loaded for editing. */
async function framed(page: Page, bytes: Uint8Array): Promise<Frame> {
  await page.goto('/e2e/pages/blank.html');
  await page.setContent(
    `<!doctype html><body style="margin:0"><button id="outside">outside</button><iframe id="frame" src="${BRIDGE}" style="display:block;width:100%;height:520px;border:0"></iframe></body>`,
  );
  await page.waitForFunction(() => {
    const frame = document.getElementById('frame') as HTMLIFrameElement | null;
    return !!frame && !!frame.contentWindow && !!(frame.contentWindow as unknown as { BymOfficeBridge?: unknown }).BymOfficeBridge;
  });
  await page.evaluate(
    (data) =>
      new Promise<void>((resolve, reject) => {
        const inner = (document.getElementById('frame') as HTMLIFrameElement).contentWindow as unknown as {
          BymOfficeHost: { postMessage(json: string): void };
          BymOfficeBridge: { receive(json: string): void };
        };
        const sent: { type: string; dirty?: boolean; code?: string }[] = [];
        (window as unknown as { sent: typeof sent }).sent = sent;
        inner.BymOfficeHost = {
          postMessage(json) {
            const message = JSON.parse(json);
            sent.push(message);
            if (message.type === 'loaded') resolve();
            if (message.type === 'error') reject(new Error(message.code));
          },
        };
        inner.BymOfficeBridge.receive(JSON.stringify({ type: 'load', id: 'file', data, save: true, locale: 'en' }));
      }),
    base64(bytes),
  );
  const frame = page.frames().find((candidate) => candidate.url().endsWith(BRIDGE));
  if (!frame) throw new Error('the frame of the bridge page is not there');
  await expect(frame.locator('canvas').first()).toBeVisible();
  return frame;
}

const changes = (page: Page): Promise<number> =>
  page.evaluate(() => (window as unknown as { sent: { type: string }[] }).sent.filter((message) => message.type === 'changed').length);

for (const format of ['xlsx', 'docx'] as const) {
  test(`${format}: a press into the frame while the page outside has the focus, then typing: what is typed is in the file`, async ({ page, hasTouch }) => {
    test.skip(hasTouch, 'a finger does not type; the press of a mouse is what loses the focus');
    const frame = await framed(page, await createBlank(format, { locale: 'en' }));
    // The focus is with the page outside, as it is with an application that has just opened its viewer.
    await page.focus('#outside');
    expect(await page.evaluate(() => document.activeElement && document.activeElement.id)).toBe('outside');

    const box = await frame.locator('canvas').first().boundingBox();
    if (!box) throw new Error('the canvas has no box');
    await page.mouse.click(box.x + Math.min(240, box.width / 2), box.y + Math.min(120, box.height / 2));
    await page.keyboard.type('first');
    if (format === 'xlsx') await page.keyboard.press('Enter');

    await expect.poll(() => changes(page), { message: 'the workbook or document was changed by what was typed' }).toBeGreaterThan(0);
    // The input of the editor has the focus, inside the frame; the press took it from the page outside.
    expect(await frame.evaluate(() => (document.activeElement ? document.activeElement.tagName : ''))).toBe('TEXTAREA');
    expect(await page.evaluate(() => document.activeElement && document.activeElement.id)).toBe('frame');
    const dirty = await page.evaluate(() => {
      const states = (window as unknown as { sent: { type: string; dirty?: boolean }[] }).sent.filter((message) => message.type === 'saveState');
      return states.length ? states[states.length - 1].dirty : undefined;
    });
    expect(dirty, 'the page says there is something to save').toBe(true);
  });
}

test('a press into the frame does not take the focus back from a field the application focused meanwhile', async ({ page, hasTouch }) => {
  test.skip(hasTouch, 'as above');
  const frame = await framed(page, await createBlank('xlsx', { locale: 'en' }));
  await page.focus('#outside');
  const box = await frame.locator('canvas').first().boundingBox();
  if (!box) throw new Error('the canvas has no box');
  // The application moves the focus to a field of its own while the button is down (a dialog that opens on the press).
  await page.mouse.move(box.x + 200, box.y + 100);
  await page.mouse.down();
  await frame.evaluate(() => {
    const field = document.createElement('input');
    field.id = 'own';
    document.body.appendChild(field);
    field.focus();
  });
  await page.mouse.up();
  expect(await frame.evaluate(() => (document.activeElement ? document.activeElement.id : ''))).toBe('own');
});
