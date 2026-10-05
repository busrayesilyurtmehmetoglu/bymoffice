import { simpleDoc } from '../test/doc-binary-fixtures';
import { minimalCfb, sampleWorkbook } from '../test/fixtures';
import { cat, number, workbookStream, xlsFile } from '../test/xls-fixtures';
import { expect, test } from './helpers';

// The embedding entry point as a host page would use it.
test.describe('BymOffice.mount', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
  });

  test('shows a workbook from bytes and reports selection changes', async ({ page }) => {
    const bytes = Array.from(await sampleWorkbook());
    const result = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const events: string[] = [];
      const mounted = await api.mount(document.getElementById('editor'), {
        file: new Uint8Array(input),
        onSelectionChange: (selection: any) => events.push(api.rangeName(selection.range)),
      });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 0, left: 0, bottom: 1, right: 1 });
      return { format: mounted.file.format, sheets: mounted.workbook.sheets.length, events };
    }, bytes);
    expect(result).toEqual({ format: 'xlsx', sheets: 3, events: ['A1:B2'] });
    await expect(page.locator('#editor .bym-tabs button')).toHaveText(['Veri', 'Grafik']);

    await page.evaluate(() => (window as any).__mounted.destroy());
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('the locale option and setLocale change what cells show', async ({ page }) => {
    const bytes = Array.from(await sampleWorkbook());
    await page.evaluate(async (input) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), locale: 'tr' });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 1, left: 1, bottom: 1, right: 1 });
    }, bytes);
    const live = page.locator('#editor [aria-live]');
    await expect(live).toHaveText('B2 12,50 ₺');

    await page.evaluate(() => (window as any).__mounted.view.setLocale('en'));
    await expect(live).toHaveText('B2 12.50 ₺');

    // An incomplete locale object falls back to en-US for what it leaves out.
    await page.evaluate(() => (window as any).__mounted.view.setLocale({ decimal: ';' }));
    await expect(live).toHaveText('B2 12;50 ₺');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('the view speaks the language of its locale', async ({ page, isMobile }) => {
    const bytes = Array.from(await sampleWorkbook());
    await page.evaluate(async (input) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), locale: 'tr' });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 0, left: 0, bottom: 1, right: 1 });
    }, bytes);
    await expect(page.locator('#editor .bym-view')).toHaveAttribute('aria-label', 'Hesap tablosu');
    await expect(page.locator('#editor .bym-input')).toHaveAttribute('aria-label', 'Hücre düzenleyici');
    await expect(page.locator('#editor .bym-scroller')).toHaveAttribute('aria-label', 'Sayfa Veri');
    await expect(page.locator('#editor [aria-live]')).toHaveText('A1 Ürün, A1:B2 seçili');

    if (!isMobile) {
      // Header of row 5: the header strip is 72 px wide; rows 1-2 are 32 and 20 px, row 3 is hidden.
      await page.mouse.click(30, 22 + 32 + 20 + 20 + 10, { button: 'right' });
      await expect(page.locator('#editor .bym-menu button')).toHaveText(['Üste satır ekle', 'Alta satır ekle', 'Satırı sil']);
      await page.keyboard.press('Escape');
    }

    await page.evaluate(() => (window as any).__mounted.view.setLocale('en'));
    await expect(page.locator('#editor .bym-view')).toHaveAttribute('aria-label', 'Spreadsheet');
    await expect(page.locator('#editor .bym-scroller')).toHaveAttribute('aria-label', 'Sheet Veri');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('loads the file from a URL', async ({ page, pageProblems }) => {
    const bytes = Buffer.from(await sampleWorkbook());
    await page.route('**/files/sample.xlsx', (route) => route.fulfill({ body: bytes, contentType: 'application/octet-stream' }));
    await page.route('**/files/missing.xlsx', (route) => route.fulfill({ status: 404, body: 'nope' }));

    const names = await page.evaluate(async () => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: '/files/sample.xlsx' });
      const sheets = mounted.workbook.sheets.map((sheet: any) => sheet.name);
      mounted.destroy();
      return sheets;
    });
    expect(names).toEqual(['Veri', 'Gizli', 'Grafik']);

    const message = await page.evaluate(() =>
      (window as any).BymOffice.mount(document.getElementById('editor'), { file: '/files/missing.xlsx' }).then(
        () => 'resolved',
        (error: Error) => error.message,
      ),
    );
    expect(message).toBe('Could not load the file: HTTP 404');
    // Browsers log the failed request themselves; that one message is expected here.
    expect(pageProblems.every((problem) => problem.includes('missing.xlsx'))).toBe(true);
    pageProblems.length = 0;
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('an xls file is shown read-only and a doc file as a document', async ({ page }) => {
    const sheet = cat(number(0, 0, 41), number(1, 0, 1.5));
    const xlsBytes = Array.from(xlsFile(workbookStream({ sheets: [{ name: 'Old', records: sheet }] })));
    const docBytes = Array.from(simpleDoc('First paragraph\rSecond paragraph\r').build());
    const outcome = await page.evaluate(
      async ([xls, doc]) => {
        const api = (window as any).BymOffice;
        const host = document.getElementById('editor')!;
        const first = await api.mount(host, { file: new Uint8Array(xls) });
        const sheetView = first.view;
        const before = sheetView.workbook.sheets[0].rows.get(0).cells[0].value;
        // Editing cannot be turned on for a file that cannot be written.
        first.editor.setReadOnly(false);
        sheetView.setReadOnly(false);
        sheetView.setCellInput(0, 0, '7');
        const result = {
          xlsKind: first.kind,
          format: first.file.format,
          sheet: sheetView.workbook.sheets[0].name,
          before,
          after: sheetView.workbook.sheets[0].rows.get(0).cells[0].value,
          readOnly: sheetView.readOnly,
          docKind: '',
          pages: 0,
          text: '',
        };
        first.destroy();
        const third = await api.mount(host, { file: new TextEncoder().encode('{\\rtf1\\ansi Hello {\\b rich} text\\par Second\\par}') });
        const rtf = [third.kind, third.file.format, third.view.pageCount, api.paragraphText(third.document.body.blocks[0])].join('|');
        third.destroy();
        if (rtf !== 'document|rtf|1|Hello rich text') throw new Error('rtf: ' + rtf);
        const second = await api.mount(host, { file: new Uint8Array(doc) });
        result.docKind = second.kind;
        result.pages = second.view.pageCount;
        result.text = api.paragraphText(second.document.body.blocks[0]);
        second.destroy();
        return result;
      },
      [xlsBytes, docBytes],
    );
    expect(outcome).toEqual({ xlsKind: 'spreadsheet', format: 'xls', sheet: 'Old', before: 41, after: 41, readOnly: true, docKind: 'document', pages: 1, text: 'First paragraph' });
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('convert: doc, RTF and xls are opened for editing as docx and xlsx, and saved as such', async ({ page }) => {
    const sheet = cat(number(0, 0, 41), number(1, 0, 1.5));
    const xlsBytes = Array.from(xlsFile(workbookStream({ sheets: [{ name: 'Old', records: sheet }] })));
    const docBytes = Array.from(simpleDoc('First paragraph\rSecond paragraph\r').build());
    const outcome = await page.evaluate(
      async ([xls, doc]) => {
        const api = (window as any).BymOffice;
        const host = document.getElementById('editor')!;
        const out: Record<string, unknown> = {};
        const notices: string[] = [];
        const onNotice = (notice: { message: string }) => notices.push(notice.message);

        // A doc file: converted, edited, saved as docx.
        let mounted = await api.mount(host, { file: new Uint8Array(doc), convert: true, onNotice });
        out.doc = [mounted.kind, mounted.file.format, mounted.conversion.from, mounted.conversion.to, mounted.view.readOnly];
        mounted.view.setSelection({ story: 'body', path: [0], offset: 0 });
        mounted.view.insertText('Edited ');
        let again = await api.open(await mounted.save());
        out.docSaved = [again.format, api.documentText(await again.document())];
        mounted.destroy();

        // RTF, in Turkish.
        mounted = await api.mount(host, { file: new TextEncoder().encode('{\\rtf1\\ansi Hello {\\b rich} text\\par}'), convert: true, locale: 'tr', onNotice });
        out.rtf = [mounted.file.format, mounted.conversion.from, mounted.view.readOnly, mounted.conversion.warnings.length];
        mounted.destroy();

        // An xls file: a cell is changed and the xlsx holds it.
        mounted = await api.mount(host, { file: new Uint8Array(xls), convert: true, onNotice });
        out.xls = [mounted.kind, mounted.file.format, mounted.conversion.from, mounted.view.readOnly];
        mounted.view.setCellInput(0, 0, '7');
        again = await api.open(await mounted.save());
        const book = await again.workbook();
        out.xlsSaved = [again.format, book.sheets[0].name, api.getCell(book.sheets[0], 0, 0).value, api.getCell(book.sheets[0], 1, 0).value];
        mounted.destroy();

        // Read-only, and files that need no conversion: nothing is converted.
        mounted = await api.mount(host, { file: new Uint8Array(doc), convert: true, readOnly: true, onNotice });
        out.readOnly = [mounted.file.format, mounted.conversion === undefined, mounted.view.readOnly];
        mounted.destroy();
        mounted = await api.mount(host, { file: await api.createBlank('docx'), convert: true, onNotice });
        out.docx = [mounted.file.format, mounted.conversion === undefined, mounted.view.readOnly];
        mounted.destroy();
        out.notices = notices;
        return out;
      },
      [xlsBytes, docBytes],
    );
    expect(outcome).toEqual({
      doc: ['document', 'docx', 'doc', 'docx', false],
      docSaved: ['docx', 'Edited First paragraph\nSecond paragraph'],
      rtf: ['docx', 'rtf', false, 0],
      xls: ['spreadsheet', 'xlsx', 'xls', false],
      xlsSaved: ['xlsx', 'Old', 7, 1.5],
      readOnly: ['doc', true, true],
      docx: ['docx', true, false],
      notices: [
        'This doc file was converted to docx to be edited; it is saved as docx',
        'Bu rtf dosyası düzenlenebilmesi için docx biçimine çevrildi; docx olarak kaydedilir',
        'This xls file was converted to xlsx to be edited; it is saved as xlsx',
      ],
    });
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('doc and RTF that are not converted cannot be made editable: nothing unlocks the view and the model stays as it is', async ({ page }) => {
    const docBytes = Array.from(simpleDoc('First paragraph\rSecond paragraph\r').build());
    const outcome = await page.evaluate(async (doc) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const out: unknown[] = [];
      for (const file of [new Uint8Array(doc), new TextEncoder().encode('{\\rtf1\\ansi Hello text\\par}')]) {
        const mounted = await api.mount(host, { file, readOnly: false });
        const before = api.documentText(mounted.document);
        mounted.editor.setReadOnly(false);
        mounted.view.setReadOnly(false);
        mounted.view.setSelection({ story: 'body', path: [0], offset: 0 });
        const typed = mounted.view.insertText('Lost ');
        const saved = await mounted.save();
        out.push([mounted.file.format, mounted.view.readOnly, typed, api.documentText(mounted.document) === before, mounted.view.isModified, saved.length === file.length]);
        mounted.destroy();
      }
      // A docx is not locked.
      const mounted = await api.mount(host, { file: await api.createBlank('docx'), readOnly: true });
      mounted.editor.setReadOnly(false);
      out.push([mounted.file.format, mounted.view.readOnly]);
      mounted.destroy();
      return out;
    }, docBytes);
    expect(outcome).toEqual([
      ['doc', true, false, true, false, true],
      ['rtf', true, false, true, false, true],
      ['docx', false],
    ]);
  });

  test('convert: what the new file does not carry over is in conversion.warnings; a file that loses nothing has none', async ({ page }) => {
    const sheet = cat(number(0, 0, 41));
    const charted = Array.from(xlsFile(workbookStream({ sheets: [{ name: 'Data', records: sheet }, { name: 'Chart1', type: 2 }] })));
    const plain = Array.from(xlsFile(workbookStream({ sheets: [{ name: 'Data', records: sheet }] })));
    const outcome = await page.evaluate(
      async ([withChart, without]) => {
        const api = (window as any).BymOffice;
        const host = document.getElementById('editor')!;
        const rtf = (body: string) => new TextEncoder().encode('{\\rtf1\\ansi ' + body + '\\par}');
        // An ellipse is carried over with its outline and fill; an arrow (type 13) has a geometry the model does not hold.
        const shapeOf = (type: number) => `{\\shp{\\*\\shpinst\\shpleft0\\shptop0\\shpright2000\\shpbottom1000{\\sp{\\sn shapeType}{\\sv ${type}}}}}`;
        const shape = shapeOf(3);
        const out: Record<string, unknown> = {};
        for (const [key, file] of [
          ['chart', new Uint8Array(withChart)],
          ['plain', new Uint8Array(without)],
          ['shape', rtf('Before ' + shape + ' after')],
          ['arrow', rtf('Before ' + shapeOf(13) + ' after')],
          ['text', rtf('Only text')],
        ] as [string, Uint8Array][]) {
          const mounted = await api.mount(host, { file, convert: true });
          out[key] = [mounted.conversion.to, mounted.conversion.warnings];
          mounted.destroy();
        }
        return out;
      },
      [charted, plain],
    );
    expect(outcome).toEqual({
      chart: ['xlsx', ['Sheet: the content of a chart sheet is not in the model; it is written as an empty worksheet']],
      plain: ['xlsx', []],
      shape: ['docx', []],
      arrow: ['docx', ['1 shape(s) were written as rectangles: their geometry was not carried over']],
      text: ['docx', []],
    });
  });

  test('convert: a picture goes into the docx and stays through an edit; xls with readOnly and xlsx are left alone; a converted xlsx locks and unlocks', async ({ page }) => {
    const xlsBytes = Array.from(xlsFile(workbookStream({ sheets: [{ name: 'Old', records: cat(number(0, 0, 41)) }] })));
    const outcome = await page.evaluate(async (xls) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const out: Record<string, unknown> = {};
      const picture =
        '{\\pict\\pngblip\\picw10\\pich10\\picwgoal150\\pichgoal150 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000b49444154789c6360000200000500017a5eab3f0000000049454e44ae426082}';
      let mounted = await api.mount(host, { file: new TextEncoder().encode('{\\rtf1\\ansi Text ' + picture + ' here\\par}'), convert: true });
      mounted.view.setSelection({ story: 'body', path: [0], offset: 0 });
      mounted.view.insertText('Edited ');
      const saved = await api.open(await mounted.save());
      const media = saved.package.partNames().filter((name: string) => name.indexOf('/word/media/') === 0);
      const bytes = media.length ? await saved.package.readPart(media[0]) : new Uint8Array(0);
      out.picture = [saved.format, mounted.conversion.warnings, media.length, bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47, api.documentText(await saved.document()).indexOf('Edited Text') === 0];
      mounted.destroy();

      mounted = await api.mount(host, { file: new Uint8Array(xls), convert: true, readOnly: true });
      out.xlsReadOnly = [mounted.file.format, mounted.conversion === undefined, mounted.view.readOnly];
      mounted.editor.setReadOnly(false);
      out.xlsStaysLocked = mounted.view.readOnly;
      mounted.destroy();

      const xlsx = await api.createBlank('xlsx');
      mounted = await api.mount(host, { file: xlsx, convert: true });
      out.xlsx = [mounted.file.format, mounted.conversion === undefined, mounted.view.readOnly];
      mounted.destroy();

      mounted = await api.mount(host, { file: new Uint8Array(xls), convert: true });
      const states = [mounted.view.readOnly];
      mounted.editor.setReadOnly(true);
      states.push(mounted.view.readOnly);
      mounted.view.setCellInput(0, 0, '5');
      const locked = mounted.workbook.sheets[0].rows.get(0).cells[0].value;
      mounted.editor.setReadOnly(false);
      states.push(mounted.view.readOnly);
      mounted.view.setCellInput(0, 0, '6');
      out.converted = [states, locked, mounted.workbook.sheets[0].rows.get(0).cells[0].value];
      mounted.destroy();
      return out;
    }, xlsBytes);
    expect(outcome).toEqual({
      picture: ['docx', [], 1, true, true],
      xlsReadOnly: ['xls', true, true],
      xlsStaysLocked: true,
      xlsx: ['xlsx', true, false],
      converted: [[false, true, false], 41, 6],
    });
  });

  test('convert: the notice follows the locale as the views do: a locale object leaves the language to the browser', async ({ page }) => {
    const notices = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const out: string[] = [];
      const file = () => new TextEncoder().encode('{\\rtf1\\ansi Text\\par}');
      for (const locale of [undefined, api.LOCALE_TR, 'tr']) {
        const mounted = await api.mount(host, { file: file(), convert: true, locale, onNotice: (notice: { message: string }) => out.push(notice.message) });
        mounted.destroy();
      }
      return out;
    });
    // The browser of the tests speaks English: only the language tag makes the notice Turkish.
    expect(notices).toEqual([
      'This rtf file was converted to docx to be edited; it is saved as docx',
      'This rtf file was converted to docx to be edited; it is saved as docx',
      'Bu rtf dosyası düzenlenebilmesi için docx biçimine çevrildi; docx olarak kaydedilir',
    ]);
  });

  test('rejects files it cannot show and leaves the container empty', async ({ page }) => {
    const xls = Array.from(minimalCfb(new Uint8Array(9 * 512), new Uint8Array(10)));
    const outcome = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor');
      const names: string[] = [];
      for (const file of [new Uint8Array(input), new TextEncoder().encode('plain text')]) {
        await api.mount(host, { file }).catch((error: Error) => names.push(error.name));
      }
      return names;
    }, xls);
    // The compound file names a workbook stream that is not one: damaged, not unsupported.
    expect(outcome).toEqual(['Error', 'UnsupportedFileError']);
    await expect(page.locator('#editor')).toBeEmpty();
  });
});

test.describe('BymOffice.mount in a browser that speaks Turkish', () => {
  test.use({ locale: 'tr-TR' });

  test('convert: without a locale, and with a locale object, the notice is in the language of the browser', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    const notices = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const out: string[] = [navigator.language];
      const file = () => new TextEncoder().encode('{\\rtf1\\ansi Text\\par}');
      for (const locale of [undefined, api.LOCALE_EN, 'en']) {
        const mounted = await api.mount(host, { file: file(), convert: true, locale, onNotice: (notice: { message: string }) => out.push(notice.message) });
        // The view speaks the same language as the notice.
        out.push(host.querySelector('.bym-doc')!.getAttribute('aria-label')!);
        mounted.destroy();
      }
      return out;
    });
    const turkish = 'Bu rtf dosyası düzenlenebilmesi için docx biçimine çevrildi; docx olarak kaydedilir';
    expect(notices).toEqual(['tr-TR', turkish, 'Belge', turkish, 'Belge', 'This rtf file was converted to docx to be edited; it is saved as docx', 'Document']);
  });
});
