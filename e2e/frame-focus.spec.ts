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
    `<!doctype html><meta name="viewport" content="width=device-width, initial-scale=1" /><body style="margin:0"><button id="outside">outside</button><input id="field" /><p id="plain" style="margin:0;padding:8px">plain text of the page</p><iframe id="frame" src="${BRIDGE}" style="display:block;width:100%;height:520px;border:0"></iframe></body>`,
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

test('the page outside gives the focus to a field of its own while the press into the frame is held: the release leaves it there', async ({ page, hasTouch }) => {
  test.skip(hasTouch, 'as above');
  const frame = await framed(page, await createBlank('xlsx', { locale: 'en' }));
  await page.focus('#outside');
  const box = await frame.locator('canvas').first().boundingBox();
  if (!box) throw new Error('the canvas has no box');
  await page.mouse.move(box.x + 200, box.y + 100);
  await page.mouse.down();
  // A field of the page outside, not of the frame: inside the frame nothing has the focus, before and after.
  await page.evaluate(() => (document.getElementById('field') as HTMLInputElement).focus());
  await page.mouse.up();
  expect(await page.evaluate(() => document.activeElement && document.activeElement.id), 'the editor in the frame did not take the focus at the release').toBe('field');
  await page.keyboard.type('abc');
  expect(await page.evaluate(() => (document.getElementById('field') as HTMLInputElement).value)).toBe('abc');
  expect(await changes(page), 'nothing was typed into the file').toBe(0);
});

test('focus() of a view that is hidden does not wait for the next release: a press elsewhere in the page leaves the focus where it is', async ({ page, hasTouch }) => {
  await page.goto('/e2e/pages/mount.html');
  await page.evaluate(async () => {
    const w = window as unknown as { BymOffice: { createBlank(format: string): Promise<Uint8Array>; mount(el: HTMLElement, options: object): Promise<{ view: { focus(): void } }> }; mounted: { view: { focus(): void } } };
    const editor = document.getElementById('editor') as HTMLElement;
    const plain = document.createElement('p');
    plain.id = 'plain';
    plain.textContent = 'plain text of the page';
    plain.style.cssText = 'margin:0;padding:12px';
    document.body.insertBefore(plain, editor);
    w.mounted = await w.BymOffice.mount(editor, { file: await w.BymOffice.createBlank('xlsx'), locale: 'en' });
    // The application hides the editor (a tab of its own that is not shown) and asks for the focus all the same.
    editor.style.display = 'none';
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    for (let i = 0; i < 5; i++) w.mounted.view.focus();
    editor.style.display = '';
  });
  expect(await page.evaluate(() => (document.activeElement ? document.activeElement.tagName : '')), 'a hidden input takes no focus').toBe('BODY');
  const plain = page.locator('#plain');
  if (hasTouch) await plain.tap();
  else await plain.click();
  expect(await page.evaluate(() => (document.activeElement ? document.activeElement.tagName : '')), 'the press on plain text did not hand the focus to the editor').toBe('BODY');
  // And the editor takes it when it is pressed itself.
  const canvas = page.locator('#editor canvas').first();
  if (hasTouch) await canvas.tap({ position: { x: 150, y: 100 } });
  else await canvas.click({ position: { x: 150, y: 100 } });
  expect(await page.evaluate(() => (document.activeElement ? document.activeElement.tagName : ''))).toBe('TEXTAREA');
});

test('a finger scrolls the grid in the frame and then taps the page outside: the editor does not take the focus', async ({ page, hasTouch, browserName }) => {
  test.skip(!hasTouch, 'a finger is needed');
  const frame = await framed(page, await createBlank('xlsx', { locale: 'en' }));
  await page.focus('#outside');
  const box = await frame.locator('canvas').first().boundingBox();
  if (!box) throw new Error('the canvas has no box');
  const x = box.x + Math.min(150, box.width / 2);
  const y = box.y + Math.min(260, box.height - 40);
  let real = false;
  if (browserName === 'chromium') {
    // A real gesture: the browser scrolls and tells the page that the pointer was taken from it.
    const session = await page.context().newCDPSession(page);
    const point = (dy: number) => [{ x, y: y - dy, id: 1 }];
    await session.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: point(0) });
    for (let dy = 20; dy <= 160; dy += 20) await session.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: point(dy) });
    await session.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    real = true;
  } else {
    // What the page sees of such a gesture: the finger goes down, moves, and the browser takes it over for the scroll.
    await frame.evaluate(() => {
      const canvas = document.querySelector('canvas') as HTMLCanvasElement;
      const rect = canvas.getBoundingClientRect();
      const fire = (type: string, dy: number): void => {
        canvas.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 11, pointerType: 'touch', isPrimary: true, clientX: rect.left + 150, clientY: rect.top + 200 - dy }));
      };
      fire('pointerdown', 0);
      for (let dy = 20; dy <= 160; dy += 20) fire('pointermove', dy);
      fire('pointercancel', 160);
    });
  }
  expect(await frame.evaluate(() => (document.activeElement ? document.activeElement.tagName : '')), `after the scroll (${real ? 'a real gesture' : 'events of one'})`).toBe('BODY');
  await page.locator('#plain').tap();
  expect(await frame.evaluate(() => (document.activeElement ? document.activeElement.tagName : '')), 'the tap outside the frame did not hand the focus to the editor').toBe('BODY');
  expect(await page.evaluate(() => (document.activeElement ? document.activeElement.id : '')), 'the focus of the page outside is not in the frame').not.toBe('frame');
});
