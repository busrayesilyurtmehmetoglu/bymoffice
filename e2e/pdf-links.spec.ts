// The links and the outline of a PDF made in the browser: exportPdf() of a
// document and of a workbook, and the command of the WebView bridge. The
// file is taken out of the browser and read by the reader of the tests
// (test/pdf-fixtures), which follows the specification: the link
// annotations of its pages, their addresses and places, and its outline.

import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { Ref as RefClass, checkPdf, poppler, popplerOrSay, type CheckedPdf, type Dict, type Ref, type Str } from '../test/pdf-fixtures';
import { DOC_STYLES } from './fixtures/docx';
import { expect, test } from './helpers';

const SECTION = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/></w:sectPr>';
const STYLES =
  DOC_STYLES +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>';

const heading = (level: number, text: string) => p(r(text), `<w:pStyle w:val="Heading${level}"/>`);

/** Two pages: headings, a link out of the document, one to a bookmark on the second page, one that may not be a link. */
function linkedDocx(): Promise<Uint8Array> {
  const body =
    heading(1, 'Giriş') +
    p(r('Bakınız ') + `<w:hyperlink r:id="rIdWeb">${r('örnek site')}</w:hyperlink>` + r(' ve ') + `<w:hyperlink w:anchor="Sonuc">${r('sonuç bölümü')}</w:hyperlink>` + r(' ve ') + `<w:hyperlink r:id="rIdBad">${r('kötü')}</w:hyperlink>`) +
    heading(2, 'Ayrıntı') +
    p(r('Sonraki sayfa'), '<w:pageBreakBefore/>') +
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="0" w:name="Sonuc"/>${r('Sonuç')}<w:bookmarkEnd w:id="0"/></w:p>`;
  return buildDocx({
    styles: STYLES,
    body: body + SECTION,
    rels: [
      ['rIdWeb', 'hyperlink', 'https://example.com/örnek', true],
      ['rIdBad', 'hyperlink', 'javascript:alert(1)', true],
    ],
  });
}

const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const SETUP = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9"/>';

/** Two sheets: on the first a link out of the workbook and one to a cell of the second. */
function linkedXlsx(): Promise<Uint8Array> {
  const first = worksheetXml(
    `<cols><col min="1" max="1" width="30" customWidth="1"/></cols><sheetData><row r="1">${str('A1', 'Web sitesi')}</row><row r="2">${str('A2', 'Özete git')}</row></sheetData>` +
      '<hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" location="\'Özet\'!B3"/></hyperlinks>' +
      SETUP,
  );
  const second = worksheetXml(`<sheetData><row r="3">${str('B3', 'Hedef')}</row></sheetData>${SETUP}`);
  return buildXlsx({ sheets: [{ name: 'Veri', xml: first, links: { rId1: 'https://example.com/tablo' } }, { name: 'Özet', xml: second }] });
}

interface Made {
  data?: string;
  error?: string;
  report?: Record<string, number>;
}

/** exportPdf() of window.__view, with the report of its links. */
function exportPdf(page: Page, options: object = {}): Promise<Made> {
  return page.evaluate(async (opts) => {
    let report: Record<string, number> | undefined;
    try {
      const bytes: Uint8Array = await (window as any).__view.exportPdf({ text: 'image', dpi: 72, ...(opts as object), onLinks: (one: Record<string, number>) => (report = one) });
      let binary = '';
      for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
      return { data: btoa(binary), report };
    } catch (error: any) {
      return { error: String(error && error.message) };
    }
  }, options);
}

function read(data: string | undefined): CheckedPdf {
  expect(data, 'a PDF was made').toBeTruthy();
  const bytes = new Uint8Array(Buffer.from(data as string, 'base64'));
  const checked = checkPdf(bytes);
  if (popplerOrSay()) expect(poppler(bytes).complaints, 'Poppler has nothing to say against the file').toBe('');
  return checked;
}

const text = (value: Str) => {
  const bytes = value.bytes;
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    let out = '';
    for (let i = 2; i + 1 < bytes.length; i += 2) out += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
    return out;
  }
  return String.fromCharCode(...Array.from(bytes));
};

/** The links of every page: the address, or the index of the page they go to. */
function linksOf(pdf: CheckedPdf): (string | number | null)[][] {
  const object = (ref: Ref) => pdf.objects.get(ref.id)!.value as Dict;
  const root = object(pdf.trailer.Root as Ref);
  const kids = (object(root.Pages as Ref).Kids as Ref[]).map((kid) => kid.id);
  return kids.map((id) => {
    const page = pdf.objects.get(id)!.value as Dict;
    // /Annots is an array, or a reference to one (a page with links into the file).
    const annots = page.Annots instanceof RefClass ? (object(page.Annots) as unknown as Ref[]) : (page.Annots as Ref[] | undefined);
    return (annots || []).map((ref) => {
      const annot = object(ref);
      if (annot.A) return text((annot.A as Dict).URI as Str);
      if (annot.Dest) return kids.indexOf(((annot.Dest as unknown[])[0] as Ref).id);
      return null;
    });
  });
}

/** The titles of the outline, deeper ones indented by two spaces. */
function outlineOf(pdf: CheckedPdf): string[] {
  const object = (ref: Ref) => pdf.objects.get(ref.id)!.value as Dict;
  const root = object(pdf.trailer.Root as Ref);
  if (!root.Outlines) return [];
  const out: string[] = [];
  const walk = (first: Ref | undefined, depth: number) => {
    for (let ref = first; ref; ref = object(ref).Next as Ref | undefined) {
      out.push('  '.repeat(depth) + text(object(ref).Title as Str));
      walk(object(ref).First as Ref | undefined, depth + 1);
    }
  };
  walk(object(root.Outlines as Ref).First as Ref, 0);
  return out;
}

async function mountDoc(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate((input) => (window as any).__mountDoc(input, {}), Array.from(bytes));
  await page.evaluate(() => ((window as any).__view = (window as any).__docView));
  await expect.poll(() => page.evaluate(() => (window as any).__view.layoutDone)).toBe(true);
}

test.describe('links and outline of a PDF', () => {
  test('a document: its hyperlinks become links, its headings the outline; what may not be a link is counted', async ({ page }) => {
    await mountDoc(page, await linkedDocx());
    const made = await exportPdf(page);
    expect(made.error).toBeUndefined();
    const pdf = read(made.data);
    expect(pdf.pages.length).toBe(2);
    expect(linksOf(pdf)).toEqual([['https://example.com/%C3%B6rnek', 1], []]);
    expect(outlineOf(pdf)).toEqual(['Giriş', '  Ayrıntı', 'Sonuç']);
    expect(made.report).toEqual({ external: 1, internal: 1, refused: 1, unresolved: 0, dropped: 0, rotated: 0, outline: 3, outlineSkipped: 0, unread: 0 });
    // Without them.
    const bare = await exportPdf(page, { links: false, outline: false });
    const plain = read(bare.data);
    expect(linksOf(plain)).toEqual([[], []]);
    expect(outlineOf(plain)).toEqual([]);
  });

  test('a workbook: the links of its cells, into it and out of it, and its sheets in the outline', async ({ page }) => {
    await page.goto('/e2e/pages/view.html');
    await expect(page.locator('#state')).toHaveText('ready');
    await page.evaluate((input) => (window as any).__mountView(input, { locale: 'en-US' }), Array.from(await linkedXlsx()));
    await expect(page.locator('.bym-canvas')).toBeVisible();
    const made = await exportPdf(page, { sheets: 'all' });
    expect(made.error).toBeUndefined();
    const pdf = read(made.data);
    expect(pdf.pages.length).toBe(2);
    expect(linksOf(pdf)).toEqual([['https://example.com/tablo', 1], []]);
    expect(outlineOf(pdf)).toEqual(['Veri', 'Özet']);
    expect(made.report).toMatchObject({ external: 1, internal: 1, outline: 2 });
  });

  test('the bridge passes the options on and says what became of the links', async ({ page }) => {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await page.goto('/dist/bymoffice.webview.html');
    const send = (message: object) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
    const answer = async (id: string) => {
      let found: any;
      await expect.poll(async () => (found = await page.evaluate((wanted) => (window as any).__messages.find((message: any) => message.id === wanted), id)), { timeout: 60_000 }).toBeTruthy();
      return found;
    };
    await send({ type: 'load', id: 'l', name: 'Bağlantılar.docx', data: Buffer.from(await linkedDocx()).toString('base64') });
    expect((await answer('l')).type).toBe('loaded');
    await send({ type: 'exportPdf', id: 'p', dpi: 72 });
    const reply = await answer('p');
    expect(reply.type).toBe('pdf');
    expect(reply.links).toEqual({ external: 1, internal: 1, refused: 1, unresolved: 0, dropped: 0, rotated: 0, outline: 3, outlineSkipped: 0, unread: 0 });
    expect(linksOf(read(reply.data))).toEqual([['https://example.com/%C3%B6rnek', 1], []]);
    await send({ type: 'exportPdf', id: 'q', dpi: 72, outline: false });
    const without = await answer('q');
    expect(outlineOf(read(without.data))).toEqual([]);
    expect(without.links.outline).toBe(0);
    await send({ type: 'exportPdf', id: 'b', links: 'yes' });
    expect(await answer('b')).toMatchObject({ type: 'error', code: 'bad-message' });
  });
});
