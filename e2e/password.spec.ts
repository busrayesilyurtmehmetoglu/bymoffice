import type { Page } from '@playwright/test';
import { encryptionOf, open } from '../src/index';
import { simpleDoc } from '../test/doc-binary-fixtures';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { REFTOOLS, REFTOOLS_SKIPPED, passwordArg, refOk } from '../test/reftools';
import { cat, number, workbookStream } from '../test/xls-fixtures';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';
import { docx } from './fixtures/docx';
import { expect, test } from './helpers';

// Password-protected files in the browser: the question for the password inside the container,
// the `password` and `onPassword` options of mount(), saving with the password the file was opened
// with, and the same through the WebView bridge. The encrypted files are made here, by the engine
// (xlsx, docx) and by the reference script of the unit tests (xls, which the engine does not write).

const PASSWORD = 'Gizli-Parola-42';
const BRIDGE = '/dist/bymoffice.webview.html';
const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');

/** One sheet with a number in A1 and nothing else, so that typing lands in A1. */
const smallWorkbook = () => buildXlsx({ sheets: [{ name: 'Sayfa1', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') }] });

let xlsxBytes: number[] | undefined;
let docxBytes: number[] | undefined;
/** The small workbook, encrypted with PASSWORD (once: key derivation takes its time). */
async function encryptedXlsx(): Promise<number[]> {
  if (!xlsxBytes) xlsxBytes = Array.from(await (await open(await smallWorkbook())).save({ password: PASSWORD }));
  return xlsxBytes;
}
async function encryptedDocx(): Promise<number[]> {
  if (!docxBytes) docxBytes = Array.from(await (await open(await docx('<w:p><w:r><w:t>Gizli belge</w:t></w:r></w:p>'))).save({ password: PASSWORD }));
  return docxBytes;
}

/** Starts mount() in the page without waiting for it: the question is answered by the test. What comes of it is kept in the page. */
async function startMount(page: Page, bytes: number[], options: Record<string, unknown> = {}): Promise<void> {
  await page.evaluate(
    ({ input, more }) => {
      const state: any = ((window as any).__state = { done: false });
      (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), ...more }).then(
        (mounted: any) => {
          (window as any).__mounted = mounted;
          state.done = true;
          state.format = mounted.file.format;
          state.protected = mounted.file.passwordProtected;
        },
        (error: any) => {
          state.done = true;
          state.error = { name: error && error.name, code: error && error.code, message: String(error && error.message) };
        },
      );
    },
    { input: bytes, more: options },
  );
}

const state = (page: Page): Promise<any> => page.evaluate(() => (window as any).__state);
const settled = async (page: Page): Promise<any> => {
  await expect.poll(async () => (await state(page)).done, { message: 'mount() settles' }).toBe(true);
  return state(page);
};

const dialog = (page: Page) => page.locator('#editor [data-bym-dialog="password"]');
const field = (page: Page) => page.locator('#editor [data-bym-field="password"]');

test.describe('password-protected files', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
  });

  test('an encrypted xlsx asks for its password inside the container: wrong, then right, and the file is shown', async ({ page }) => {
    await startMount(page, await encryptedXlsx());
    await expect(dialog(page)).toBeVisible();
    await expect(dialog(page)).toHaveAttribute('role', 'dialog');
    await expect(dialog(page)).toHaveAttribute('aria-modal', 'true');
    await expect(field(page)).toBeFocused();
    await expect(field(page)).toHaveAttribute('type', 'password');
    // Not filled in from the passwords the browser has stored for the site, and nothing to submit or to save.
    await expect(field(page)).toHaveAttribute('autocomplete', 'new-password');
    expect(await field(page).evaluate((element) => [element.getAttribute('name'), element.closest('form') === null])).toEqual([null, true]);
    await expect(page.locator('#editor .bym-view')).toHaveCount(0);

    await field(page).fill('yanlış parola');
    await page.keyboard.press('Enter');
    // Checked, found wrong, asked again: the field is empty and says so.
    await expect(page.locator('#editor .bym-password-message')).toHaveText('Wrong password. Try again.');
    await expect(field(page)).toHaveValue('');
    await expect(field(page)).toHaveAttribute('aria-invalid', 'true');
    await expect(field(page)).toBeEnabled();
    expect((await state(page)).done).toBe(false);

    // The field is watched from here on: whenever the page is looked at after the password was sent, the field is empty.
    await page.evaluate(() => {
      const input = document.querySelector('#editor [data-bym-field="password"]') as HTMLInputElement;
      const seen: string[] = ((window as any).__fieldValues = []);
      (window as any).__fieldTimer = setInterval(() => seen.push(input.value), 5);
    });
    await field(page).fill(PASSWORD);
    await dialog(page).locator('[data-bym-action="open"]').click();
    // At once, while the password is still being checked: the dialog is there and its field holds nothing.
    expect(await page.evaluate(() => Array.from(document.querySelectorAll('#editor input[type="password"]')).map((input) => (input as HTMLInputElement).value))).toEqual(['']);
    expect(await settled(page)).toMatchObject({ done: true, format: 'xlsx', protected: true });
    await expect(dialog(page)).toHaveCount(0);
    await expect(page.locator('#editor .bym-view')).toBeVisible();
    // Nothing of the question is left: no password field anywhere in the document, and the one there was held the
    // password only between the typing and the press of the button (it was read at least once while it was typed).
    await expect(page.locator('input[type="password"]')).toHaveCount(0);
    const seen: string[] = await page.evaluate(() => {
      clearInterval((window as any).__fieldTimer);
      return (window as any).__fieldValues;
    });
    expect(seen.length).toBeGreaterThan(0);
    expect(seen[seen.length - 1]).toBe('');
    const lastHeld = seen.lastIndexOf(PASSWORD);
    expect(seen.slice(lastHeld + 1).every((value) => value === ''), 'once the password was sent the field stays empty').toBe(true);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('the question speaks Turkish with locale "tr", and Escape gives up with PasswordError "cancelled"', async ({ page }) => {
    await startMount(page, await encryptedXlsx(), { locale: 'tr' });
    await expect(dialog(page)).toBeVisible();
    await expect(page.locator('#editor .bym-password-title')).toHaveText('Parola gerekli');
    await expect(dialog(page).locator('[data-bym-action="open"]')).toHaveText('Aç');
    await field(page).fill('yarım kalan');
    await page.keyboard.press('Escape');
    expect(await settled(page)).toMatchObject({ error: { name: 'PasswordError', code: 'cancelled' } });
    await expect(dialog(page)).toHaveCount(0);
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('the Cancel button gives up as well, and a second mount asks afresh', async ({ page }) => {
    await startMount(page, await encryptedXlsx());
    await dialog(page).locator('[data-bym-action="cancel"]').click();
    expect(await settled(page)).toMatchObject({ error: { code: 'cancelled' } });
    await startMount(page, await encryptedXlsx());
    await expect(field(page)).toBeFocused();
    await field(page).fill(PASSWORD);
    await page.keyboard.press('Enter');
    expect(await settled(page)).toMatchObject({ format: 'xlsx' });
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('mount({ password }) opens the file without a question; a wrong one asks, and says it was wrong', async ({ page }) => {
    await startMount(page, await encryptedXlsx(), { password: PASSWORD });
    expect(await settled(page)).toMatchObject({ format: 'xlsx', protected: true });
    await expect(dialog(page)).toHaveCount(0);
    await page.evaluate(() => (window as any).__mounted.destroy());

    await startMount(page, await encryptedXlsx(), { password: 'eski parola' });
    await expect(page.locator('#editor .bym-password-message')).toHaveText('Wrong password. Try again.');
    await field(page).fill(PASSWORD);
    await page.keyboard.press('Enter');
    expect(await settled(page)).toMatchObject({ format: 'xlsx' });
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('onPassword is asked instead of the dialog, again for a wrong password, and null gives up', async ({ page }) => {
    const bytes = await encryptedXlsx();
    const asked = await page.evaluate(
      async ({ input, password }) => {
        const requests: any[] = [];
        const answers = ['yanlış', password];
        const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), {
          file: new Uint8Array(input),
          onPassword: async (request: any) => {
            requests.push({ ...request, dialogs: document.querySelectorAll('[data-bym-dialog="password"]').length });
            return answers.shift();
          },
        });
        const format = mounted.file.format;
        mounted.destroy();
        return { requests, format };
      },
      { input: bytes, password: PASSWORD },
    );
    expect(asked).toEqual({
      format: 'xlsx',
      requests: [
        { attempts: 0, wrong: false, dialogs: 0 },
        { attempts: 1, wrong: true, dialogs: 0 },
      ],
    });
    const gaveUp = await page.evaluate(async (input) => {
      try {
        await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), onPassword: () => null });
        return 'opened';
      } catch (error: any) {
        return `${error.name} ${error.code}`;
      }
    }, bytes);
    expect(gaveUp).toBe('PasswordError cancelled');
  });

  test('a file that is edited is saved as a compound file that opens with the same password, and with no other', async ({ page }) => {
    await startMount(page, await encryptedXlsx(), { password: PASSWORD });
    await settled(page);
    const input = page.locator('#editor .bym-input');
    await input.focus();
    await page.keyboard.type('4711');
    await page.keyboard.press('Enter');
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    const bytes = new Uint8Array(saved);
    expect(Array.from(bytes.subarray(0, 8))).toEqual([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(encryptionOf(bytes)).toMatchObject({ format: 'ooxml', encryption: { scheme: 'agile', keyBits: 256, hash: 'SHA512', spinCount: 100000 } });
    // Read back outside the browser.
    const again = await open(bytes, { password: PASSWORD });
    const sheet = (await again.workbook()).sheets[0];
    expect(sheet.rows.size).toBeGreaterThan(0);
    await expect(open(bytes)).rejects.toMatchObject({ name: 'PasswordError', code: 'required' });
    await expect(open(bytes, { password: 'başka' })).rejects.toMatchObject({ name: 'PasswordError', code: 'wrong' });
    // And inside it.
    const reopened = await page.evaluate(
      async ({ input: data, password }) => {
        const api = (window as any).BymOffice;
        const file = await api.open(new Uint8Array(data), { password });
        const workbook = await file.workbook();
        const cell = api.getCell(workbook.sheets[0], 0, 0);
        return { format: file.format, value: cell && cell.value };
      },
      { input: saved, password: PASSWORD },
    );
    expect(reopened).toEqual({ format: 'xlsx', value: 4711 });
    // Without a password only when that is asked for.
    const plain = await page.evaluate(async () => Array.from(((await (window as any).__mounted.save({ password: null })) as Uint8Array).subarray(0, 2)));
    expect(plain).toEqual([0x50, 0x4b]);
    // Once the view is gone the password is forgotten, and nothing is saved without it.
    const after = await page.evaluate(async () => {
      const mounted = (window as any).__mounted;
      mounted.destroy();
      try {
        await mounted.save();
        return 'saved';
      } catch (error: any) {
        return String(error.message);
      }
    });
    expect(after).toContain('forgotten');
  });

  test('saveTo sends the encrypted file: the server receives a compound file, never the package', async ({ page }) => {
    const bytes = await encryptedXlsx();
    const sent = await page.evaluate(
      async ({ input, password }) => {
        const received: number[][] = [];
        const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), {
          file: new Uint8Array(input),
          password,
          saveTo: async (saved: Uint8Array) => void received.push(Array.from(saved.subarray(0, 8))),
        });
        mounted.view.select({ top: 0, left: 0, bottom: 0, right: 0 });
        (window as any).__mounted = mounted;
        (window as any).__received = received;
        return received.length;
      },
      { input: bytes, password: PASSWORD },
    );
    expect(sent).toBe(0);
    await page.locator('#editor .bym-input').focus();
    await page.keyboard.type('9');
    await page.keyboard.press('Enter');
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
    expect(await page.evaluate(() => (window as any).__received)).toEqual([[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]]);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('an encrypted docx is asked for, opened and shown as a document', async ({ page }) => {
    await startMount(page, await encryptedDocx());
    await expect(field(page)).toBeFocused();
    await field(page).fill(PASSWORD);
    await page.keyboard.press('Enter');
    expect(await settled(page)).toMatchObject({ format: 'docx', protected: true });
    const text = await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document));
    expect(text).toContain('Gizli belge');
    const saved = await page.evaluate(async () => Array.from(((await (window as any).__mounted.save()) as Uint8Array).subarray(0, 4)));
    expect(saved).toEqual([0xd0, 0xcf, 0x11, 0xe0]);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('a file that is not encrypted asks nothing, and is saved as it always was', async ({ page }) => {
    await startMount(page, Array.from(await smallWorkbook()), { password: 'not needed' });
    expect(await settled(page)).toMatchObject({ format: 'xlsx', protected: false });
    await expect(dialog(page)).toHaveCount(0);
    const saved = await page.evaluate(async () => Array.from(((await (window as any).__mounted.save()) as Uint8Array).subarray(0, 2)));
    expect(saved).toEqual([0x50, 0x4b]);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('a protected xls is opened read-only with its password; converted, it is saved as an encrypted xlsx', async ({ page }) => {
    test.skip(!REFTOOLS, REFTOOLS_SKIPPED);
    const stream = workbookStream({ sheets: [{ name: 'Eski', records: cat(number(0, 0, 12.5), number(1, 0, 3)) }] });
    const xls = Array.from(writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'capi', bits: 128 }, stream) }]));
    await startMount(page, xls);
    await expect(field(page)).toBeFocused();
    await field(page).fill(PASSWORD);
    await page.keyboard.press('Enter');
    expect(await settled(page)).toMatchObject({ format: 'xls', protected: true });
    expect(await page.evaluate(() => (window as any).__mounted.view.readOnly)).toBe(true);
    expect(await page.evaluate(() => (window as any).BymOffice.getCell((window as any).__mounted.workbook.sheets[0], 0, 0).value)).toBe(12.5);
    await page.evaluate(() => (window as any).__mounted.destroy());

    await startMount(page, xls, { password: PASSWORD, convert: true });
    expect(await settled(page)).toMatchObject({ format: 'xlsx', protected: true });
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    const again = await open(new Uint8Array(saved), { password: PASSWORD });
    expect(again.format).toBe('xlsx');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('a protected doc is opened read-only with its password', async ({ page }) => {
    test.skip(!REFTOOLS, REFTOOLS_SKIPPED);
    const doc = Array.from(refOk('doc-encrypt', { password: passwordArg(PASSWORD), method: 'rc4' }, simpleDoc('Gizli metin\r').build()));
    await startMount(page, doc, { password: PASSWORD });
    expect(await settled(page)).toMatchObject({ format: 'doc', protected: true });
    expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document))).toContain('Gizli metin');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('an encryption that is not read, and a damaged encrypted file, reject without asking again', async ({ page }) => {
    const extensible = Array.from(
      writeCfb([
        { path: 'EncryptionInfo', data: new Uint8Array([4, 0, 3, 0, 0x10, 0, 0, 0, 1, 2, 3, 4]) },
        { path: 'EncryptedPackage', data: new Uint8Array(64) },
      ]),
    );
    await startMount(page, extensible);
    expect(await settled(page)).toMatchObject({ error: { name: 'PasswordError', code: 'unsupported' } });
    await expect(dialog(page)).toHaveCount(0);
    await expect(page.locator('#editor')).toBeEmpty();

    // A file whose encrypted package was changed: the password is right, the integrity check fails, and the user is
    // not asked a second time (the password was not the problem).
    const damaged = (await encryptedXlsx()).slice();
    const cfb = CfbFile.read(new Uint8Array(damaged));
    const pkg = cfb.stream('EncryptedPackage').slice();
    pkg[40] ^= 1;
    const changed = Array.from(
      writeCfb([
        { path: 'EncryptionInfo', data: cfb.stream('EncryptionInfo') },
        { path: 'EncryptedPackage', data: pkg },
      ]),
    );
    await startMount(page, changed);
    await expect(field(page)).toBeFocused();
    await field(page).fill(PASSWORD);
    await page.keyboard.press('Enter');
    expect(await settled(page)).toMatchObject({ error: { name: 'PasswordError', code: 'damaged' } });
    await expect(dialog(page)).toHaveCount(0);
    await expect(page.locator('#editor')).toBeEmpty();
  });

  /** The encrypted workbook with its integrity check taken out of the descriptor (which needs no password). */
  async function withoutIntegrity(): Promise<number[]> {
    const cfb = CfbFile.read(new Uint8Array(await encryptedXlsx()));
    const info = cfb.stream('EncryptionInfo');
    const body = new TextEncoder().encode(new TextDecoder().decode(info.subarray(8)).replace(/<dataIntegrity [^>]*\/>/, ''));
    const stripped = new Uint8Array(8 + body.length);
    stripped.set(info.subarray(0, 8));
    stripped.set(body, 8);
    return Array.from(
      writeCfb([
        { path: 'EncryptionInfo', data: stripped },
        { path: 'EncryptedPackage', data: cfb.stream('EncryptedPackage') },
      ]),
    );
  }

  test('a file without an integrity check is shown with a warning, in the view and to onNotice; requireIntegrity refuses it', async ({ page }) => {
    const bytes = await withoutIntegrity();
    const told = await page.evaluate(
      async ({ input, password }) => {
        const notices: any[] = [];
        const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), {
          file: new Uint8Array(input),
          password,
          locale: 'tr',
          onNotice: (notice: any) => notices.push(notice),
        });
        (window as any).__mounted = mounted;
        return { notices, integrity: mounted.file.encryption.integrity, warnings: mounted.file.warnings.map((warning: any) => warning.code) };
      },
      { input: bytes, password: PASSWORD },
    );
    expect(told.integrity).toBe(false);
    expect(told.warnings).toEqual(['no-integrity']);
    expect(told.notices).toContainEqual({ message: 'Bu dosyanın bütünlük denetimi yok: içeriğini parolayı bilmeyen biri değiştirmiş olabilir', kind: 'error' });
    // Also where the page does not listen: the view says it in its live region.
    await expect(page.locator('#editor [aria-live]').first()).toHaveText(/bütünlük denetimi yok/);
    await page.evaluate(() => (window as any).__mounted.destroy());

    await startMount(page, bytes, { password: PASSWORD, requireIntegrity: true });
    expect(await settled(page)).toMatchObject({ error: { name: 'PasswordError', code: 'damaged' } });
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('a file encrypted with the password everyone knows opens without a question and says so', async ({ page }) => {
    const bytes = Array.from(await (await open(await smallWorkbook())).save({ password: 'VelvetSweatshop' }));
    await startMount(page, bytes);
    expect(await settled(page)).toMatchObject({ format: 'xlsx', protected: true });
    await expect(dialog(page)).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__mounted.file.encryption.defaultPassword)).toBe(true);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('Cancel while a password is being checked ends the check: mount() rejects with "cancelled"', async ({ page }) => {
    await startMount(page, await encryptedXlsx());
    await field(page).fill('yanlış ama uzun süren');
    // Both in one turn of the page: the check has started and cannot have ended.
    await page.evaluate(() => {
      const root = document.getElementById('editor') as HTMLElement;
      (root.querySelector('[data-bym-action="open"]') as HTMLElement).click();
      (root.querySelector('[data-bym-action="cancel"]') as HTMLElement).click();
    });
    expect(await settled(page)).toMatchObject({ error: { name: 'PasswordError', code: 'cancelled' } });
    await expect(dialog(page)).toHaveCount(0);
  });
});

test.describe('password-protected files through the bridge', () => {
  /** The page as a WebView hosts it: a channel object that keeps what the page posts. */
  async function start(page: Page): Promise<{ send(message: unknown): Promise<void>; waitFor(type: string, count?: number): Promise<any>; all(): Promise<any[]> }> {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto(BRIDGE);
    const all = () => page.evaluate(() => ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)));
    return {
      all,
      send: (message) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message)),
      async waitFor(type, count = 1) {
        let found: any[] = [];
        await expect.poll(async () => (found = (await all()).filter((message) => message.type === type)).length, { message: `${count} x ${type}` }).toBeGreaterThanOrEqual(count);
        return found[count - 1];
      },
    };
  }

  test('load with a password opens the file; saved hands back an encrypted file and no message holds the password', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'load', id: 'a', data: base64(new Uint8Array(await encryptedXlsx())), password: PASSWORD });
    expect(await app.waitFor('loaded')).toMatchObject({ id: 'a', format: 'xlsx', encrypted: true });
    await app.send({ type: 'save', id: 's' });
    const saved = await app.waitFor('saved');
    expect(saved.encrypted).toBe(true);
    const bytes = new Uint8Array(Buffer.from(saved.data, 'base64'));
    expect((await open(bytes, { password: PASSWORD })).format).toBe('xlsx');
    expect(JSON.stringify(await app.all()).includes(PASSWORD)).toBe(false);
  });

  test('without a password the application is told, and the page asks the user: the answer of either opens the file', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'load', id: 'a', data: base64(new Uint8Array(await encryptedXlsx())) });
    expect(await app.waitFor('passwordRequired')).toEqual({ type: 'passwordRequired', id: 'a', wrong: false, attempts: 0, format: 'ooxml' });
    // The page shows its own question over the editor's element, not the "opening" notice.
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toBeVisible();
    await expect(page.locator('#status')).toBeHidden();
    // The user types a wrong one in the page.
    await page.locator('#editor [data-bym-field="password"]').fill('yanlış');
    await page.keyboard.press('Enter');
    expect(await app.waitFor('passwordRequired', 2)).toMatchObject({ wrong: true, attempts: 1 });
    // The application answers.
    await app.send({ type: 'password', id: 'p', value: PASSWORD });
    expect(await app.waitFor('loaded')).toMatchObject({ id: 'a', encrypted: true });
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toHaveCount(0);
    await expect(page.locator('#editor .bym-view')).toBeVisible();
    const said = JSON.stringify(await app.all());
    expect(said.includes(PASSWORD) || said.includes('yanlış')).toBe(false);
  });

  test('the password of the application takes the question of the page away', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'load', id: 'a', data: base64(new Uint8Array(await encryptedXlsx())) });
    await app.waitFor('passwordRequired');
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toBeVisible();
    await app.send({ type: 'password', value: 'uygulamanın yanlışı' });
    // Wrong: asked again, in a question that is new (and empty).
    expect(await app.waitFor('passwordRequired', 2)).toMatchObject({ wrong: true, attempts: 1 });
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toHaveCount(1);
    await expect(page.locator('#editor [data-bym-field="password"]')).toHaveValue('');
    await expect(page.locator('#editor [data-bym-field="password"]')).toBeEnabled();
    await app.send({ type: 'password', value: PASSWORD });
    expect(await app.waitFor('loaded')).toMatchObject({ id: 'a', encrypted: true });
    await expect(page.locator('input[type="password"]')).toHaveCount(0);
  });

  test('a password nobody asked for is refused with "password-not-expected"', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'password', id: 'x', value: 'kimse sormadı' });
    expect(await app.waitFor('error')).toMatchObject({ id: 'x', code: 'password-not-expected' });
    expect(JSON.stringify(await app.all()).includes('kimse sormadı')).toBe(false);
  });

  test('password with a value of null gives up: the load fails with "password-cancelled"', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'load', id: 'a', data: base64(new Uint8Array(await encryptedXlsx())), passwordPrompt: false });
    await app.waitFor('passwordRequired');
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toHaveCount(0);
    await app.send({ type: 'password', value: null });
    expect(await app.waitFor('error')).toMatchObject({ id: 'a', code: 'password-cancelled' });
    await expect(page.locator('#status')).toBeVisible();
  });

  test('Escape in the question of the page gives up as well', async ({ page }) => {
    const app = await start(page);
    await app.waitFor('ready');
    await app.send({ type: 'load', id: 'a', data: base64(new Uint8Array(await encryptedXlsx())) });
    await app.waitFor('passwordRequired');
    await page.locator('#editor [data-bym-field="password"]').focus();
    await page.keyboard.press('Escape');
    expect(await app.waitFor('error')).toMatchObject({ id: 'a', code: 'password-cancelled' });
    await expect(page.locator('#editor [data-bym-dialog="password"]')).toHaveCount(0);
  });
});
