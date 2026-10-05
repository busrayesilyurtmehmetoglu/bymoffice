import type { Page } from '@playwright/test';
import { p, r, tbl, tc, tr } from '../test/doc-layout-fixtures';
import { docx, field, picture, png, RED, BLUE } from './fixtures/docx';
import { expect, test } from './helpers';

// Two editors on one page: what is copied or cut in one is pasted in the same
// one and in the other, with its pictures, links, fields and tables.

type Editor = 'a' | 'b';

interface Position {
  story: string;
  path: number[];
  offset: number;
  before?: true;
}

const PICTURE = png(40, 20, RED, BLUE);

const source = () =>
  docx(
    p(r('ab') + picture(120, 60) + r('cd')) + // 0
      p(r('see ') + '<w:hyperlink r:id="rIdLink">' + r('the link', '<w:color w:val="0000FF"/>') + '</w:hyperlink>' + r(' and page ') + field(' PAGE ', '1') + r('.')) + // 1
      tbl(tr(tc('A1', '<w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="FFFF00"/>') + tc('B1', '<w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/>')) + tr(tc('A2') + tc('B2') + tc('C2')), [2000, 2000, 2000]) + // 2
      p('end'), // 3
    {
      rels: [
        ['rIdImg', 'image', 'media/image1.png'],
        ['rIdLink', 'hyperlink', 'https://example.com/', true],
      ],
      parts: [['word/media/image1.png', PICTURE]],
    },
  );

const target = () => docx(p('target') + tbl(tr(tc('X1') + tc('Y1')), [3000, 3000]) + p('last'));

async function mount(page: Page, a: Uint8Array, b: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/doc-clipboard.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([first, second]) => Promise.all([(window as any).__mountEditor('a', first), (window as any).__mountEditor('b', second)]), [Array.from(a), Array.from(b)]);
  await page.waitForFunction(() => (window as any).__views.a.layoutDone === true && (window as any).__views.b.layoutDone === true, undefined, { timeout: 120_000 });
}

const at = (path: number[] | number, offset: number, before?: boolean): Position => {
  const position: Position = { story: 'body', path: typeof path === 'number' ? [path] : path, offset };
  if (before) position.before = true;
  return position;
};

const select = (page: Page, editor: Editor, anchor: Position, focus: Position = anchor) =>
  page.evaluate(([name, from, to]) => (window as any).__views[name].setSelection({ anchor: from, focus: to }), [editor, anchor, focus] as [Editor, Position, Position]);

/** Fires a clipboard event at the field for typing of an editor, the way the browser does for the keys, and returns what is on the clipboard afterwards. */
function clipboardEvent(page: Page, editor: Editor, type: 'copy' | 'cut' | 'paste', give: { text?: string; html?: string } = {}): Promise<{ prevented: boolean; text: string; html: string }> {
  return page.evaluate(
    ([name, kind, content]) => {
      const target = document.querySelector(`#host-${name} .bym-doc-input`)!;
      const data = new DataTransfer();
      if (content.text !== undefined) data.setData('text/plain', content.text);
      if (content.html !== undefined) data.setData('text/html', content.html);
      let event: Event | undefined;
      try {
        event = new ClipboardEvent(kind, { clipboardData: data, bubbles: true, cancelable: true });
      } catch {
        event = undefined;
      }
      if (!event || (event as ClipboardEvent).clipboardData !== data) {
        event = new Event(kind, { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
      }
      const allowed = target.dispatchEvent(event);
      return { prevented: !allowed, text: data.getData('text/plain'), html: data.getData('text/html') };
    },
    [editor, type, give] as [Editor, 'copy' | 'cut' | 'paste', { text?: string; html?: string }],
  );
}

/** The body of an editor's document as text: runs item by item, links and field code marked, tables with their rows and cells. */
const shape = (page: Page, editor: Editor): Promise<string> =>
  page.evaluate((name) => {
    const walk = (blocks: any[]): string => {
      const out: string[] = [];
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          let text = '';
          for (const inline of block.children) {
            if (inline.type !== 'run') continue;
            let wrap = '';
            for (let w = inline.wrap; w; w = w.parent) wrap = w.kind + ':' + wrap;
            inline.content.forEach((item: any, k: number) => {
              const code = inline.fieldCode === true || (inline.fieldCode !== undefined && inline.fieldCode[k]);
              let piece = item.type === 'text' || item.type === 'instrText' ? item.text : item.type === 'fieldChar' ? `<${item.charType}>` : `<${item.type}>`;
              if (code) piece = `(${piece})`;
              text += wrap ? `[${wrap}${piece}]` : piece;
            });
          }
          out.push(text);
        } else if (block.type === 'table') out.push('TABLE(' + block.rows.map((row: any) => row.cells.map((cell: any) => walk(cell.blocks)).join(' | ')).join(' / ') + ')');
      }
      return out.join('¶');
    };
    return walk((window as any).__views[name].document.body.blocks);
  }, editor);

const notices = (page: Page) => page.evaluate(() => (window as any).__notices.map((n: any) => n.message) as string[]);

/** Saves the document of an editor, opens the bytes again and says what they hold. */
const reopened = (page: Page, editor: Editor) =>
  page.evaluate(async (name) => {
    const api = (window as any).BymOffice;
    const bytes = await (window as any).__mounted[name].save();
    const file = await api.open(bytes);
    const doc = await file.document();
    const pictures: { target: string | undefined; bytes: number[]; id: string | undefined }[] = [];
    const links: (string | undefined)[] = [];
    const tables: { grid: number[]; fill: string | undefined; span: number | undefined }[] = [];
    const walk = async (blocks: any[]) => {
      for (const block of blocks) {
        if (block.type === 'paragraph') {
          for (const inline of block.children) {
            if (inline.type !== 'run') continue;
            for (let w = inline.wrap; w; w = w.parent) if (w.kind === 'hyperlink' && links.indexOf(w.target) < 0) links.push(w.target);
            for (const item of inline.content) {
              if (item.type !== 'drawing') continue;
              const frame = item.xml.children.find((node: any) => node.type === 'element');
              const docPr = frame.children.find((node: any) => node.type === 'element' && /docPr$/.test(node.name));
              const part = item.image && item.image.target;
              pictures.push({ target: part, bytes: part ? Array.from(await file.package.readPart(part)) : [], id: docPr && docPr.attrs.get('id') });
            }
          }
        } else if (block.type === 'table') {
          const cell = block.rows[0].cells[0];
          tables.push({ grid: block.grid, fill: cell.props.shading && cell.props.shading.fill && cell.props.shading.fill.rgb, span: block.rows[0].cells[1] && block.rows[0].cells[1].props.gridSpan });
          for (const row of block.rows) for (const inner of row.cells) await walk(inner.blocks);
        }
      }
    };
    await walk(doc.body.blocks);
    return { text: api.documentText(doc) as string, pictures, links, tables };
  }, editor);

const START = 'ab<drawing>cd¶see [hyperlink:the link] and page <begin>( PAGE )(<separate>)1<end>.¶TABLE(A1 | B1 / A2 | B2 | C2)¶end';

test.describe('document editors: the clipboard carries more than text', () => {
  test('a picture: copied into the other editor with its bytes, cut and pasted as a move, saved and opened again', async ({ page }) => {
    await mount(page, await source(), await target());
    expect(await shape(page, 'a')).toBe(START);
    // The picture among text: the clipboard gets the text, and HTML made of the text with the mark of the copy.
    await select(page, 'a', at(0, 1), at(0, 3));
    const copied = await clipboardEvent(page, 'a', 'copy');
    expect(copied.prevented).toBe(true);
    expect(copied.text).toBe('bc');
    expect(copied.html).toMatch(/^<div data-bym-doc-copy="[a-z0-9]+"><p>(<span style="[^"<>]*">)?bc(<\/span>)?<\/p><\/div>$/);
    // Into the other editor: the picture is read from the package of the first one, then pasted.
    await select(page, 'b', at(0, 6));
    const pasted = await clipboardEvent(page, 'b', 'paste', copied);
    expect(pasted.prevented).toBe(true);
    await expect.poll(() => shape(page, 'b')).toBe('targetb<drawing>c¶TABLE(X1 | Y1)¶last');
    // It is shown: the view decoded the picture from the bytes that came with the copy.
    await expect.poll(() => page.evaluate(() => (window as any).__views.b.getViewState().imageBytes)).toBeGreaterThan(0);
    // One step: a single undo takes the paste back, with the picture and its relationship.
    const state = await page.evaluate(() => {
      const view = (window as any).__views.b;
      const media = view.document.media ? Array.from(view.document.media.keys()) : [];
      const relationships = view.document.body.relationships.size;
      view.undo();
      const after = { media: view.document.media ? view.document.media.size : 0, relationships: view.document.body.relationships.size };
      view.redo();
      return { media, relationships, after };
    });
    expect(state.media).toEqual(['/word/media/image1.png']);
    expect(state.after).toEqual({ media: 0, relationships: state.relationships - 1 });
    let saved = await reopened(page, 'b');
    expect(saved.text.split('\n')[0]).toBe('targetbc');
    expect(saved.pictures.map((picture) => [picture.target, picture.bytes])).toEqual([['/word/media/image1.png', Array.from(PICTURE)]]);

    // The picture alone, as a press on it selects it: a cut takes it out, a paste puts it back.
    await select(page, 'a', at(0, 2, true), at(0, 2));
    const cut = await clipboardEvent(page, 'a', 'cut');
    expect(cut.text).toBe('');
    expect(cut.html).toMatch(/data-bym-doc-copy="[a-z0-9]+"/);
    expect(await shape(page, 'a')).toBe(START.replace('ab<drawing>cd', 'abcd'));
    expect(await notices(page)).toEqual([]);
    await select(page, 'a', at(3, 3));
    await clipboardEvent(page, 'a', 'paste', cut);
    expect(await shape(page, 'a')).toBe(START.replace('ab<drawing>cd', 'abcd') + '<drawing>');
    // And the same copy once more, into the other editor: two pictures there, one part, two ids.
    await select(page, 'b', at(2, 0));
    await clipboardEvent(page, 'b', 'paste', cut);
    await expect.poll(() => shape(page, 'b')).toBe('targetb<drawing>c¶TABLE(X1 | Y1)¶<drawing>last');
    saved = await reopened(page, 'b');
    expect(saved.pictures.map((picture) => picture.target)).toEqual(['/word/media/image1.png', '/word/media/image1.png']);
    expect(new Set(saved.pictures.map((picture) => picture.id)).size).toBe(2);
    const moved = await reopened(page, 'a');
    expect(moved.text.split('\n')[0]).toBe('abcd');
    expect(moved.pictures.map((picture) => [picture.target, picture.bytes.length])).toEqual([['/word/media/image1.png', PICTURE.length]]);
    // Two steps back and the first editor has its document as it was.
    await page.evaluate(() => {
      (window as any).__views.a.undo();
      (window as any).__views.a.undo();
    });
    expect(await shape(page, 'a')).toBe(START);
  });

  test('a link and a field go with their text; a copy is known by its mark and its text, anything else is plain text', async ({ page }) => {
    await mount(page, await source(), await target());
    await select(page, 'a', at(1, 2), at(1, 23));
    const cut = await clipboardEvent(page, 'a', 'cut');
    expect(cut.text).toBe('e the link and page 1');
    // The HTML is the text, escaped, and nothing of the address or the field code.
    expect(cut.html).not.toContain('example.com');
    expect(cut.html).not.toContain('PAGE');
    expect((await shape(page, 'a')).split('¶')[1]).toBe('se.');
    await select(page, 'b', at(0, 3));
    await clipboardEvent(page, 'b', 'paste', cut);
    expect((await shape(page, 'b')).split('¶')[0]).toBe('tare [hyperlink:the link] and page <begin>( PAGE )(<separate>)1<end>get');
    const saved = await reopened(page, 'b');
    expect(saved.links).toEqual(['https://example.com/']);
    expect(saved.text.split('\n')[0]).toBe('tare the link and page 1get');
    // Back where it was: a move.
    await select(page, 'a', at(1, 2));
    await clipboardEvent(page, 'a', 'paste', cut);
    expect(await shape(page, 'a')).toBe(START);

    // The mark with another text, and the text with HTML from elsewhere: the text of the clipboard is what is pasted.
    await select(page, 'b', at(2, 0));
    await clipboardEvent(page, 'b', 'paste', { text: 'other text', html: cut.html });
    expect((await shape(page, 'b')).split('¶')[2]).toBe('other textlast');
    await select(page, 'b', at(2, 0));
    await clipboardEvent(page, 'b', 'paste', { text: cut.text, html: '<a href="https://example.org/" onclick="window.__hacked=1">' + cut.text + '</a><img src=x onerror="window.__hacked=1">' });
    expect((await shape(page, 'b')).split('¶')[2]).toBe('e the link and page 1other textlast');
    expect(await page.evaluate(() => (window as any).__hacked)).toBeUndefined();
  });

  test('a table goes as a table: between paragraphs of the other editor and into a cell of its table', async ({ page }) => {
    await mount(page, await source(), await target());
    // From the end of the paragraph before the table to the start of the one after it.
    await select(page, 'a', at(1, 24), at(3, 0));
    const copied = await clipboardEvent(page, 'a', 'copy');
    expect(copied.text).toBe('\nA1\tB1\nA2\tB2\tC2\n');
    await select(page, 'b', at(0, 6));
    await clipboardEvent(page, 'b', 'paste', copied);
    expect(await shape(page, 'b')).toBe('target¶TABLE(A1 | B1 / A2 | B2 | C2)¶¶TABLE(X1 | Y1)¶last');
    // Into a cell: a table inside the cell.
    await select(page, 'b', at([3, 0, 1, 0], 2));
    await clipboardEvent(page, 'b', 'paste', copied);
    expect(await shape(page, 'b')).toBe('target¶TABLE(A1 | B1 / A2 | B2 | C2)¶¶TABLE(X1 | Y1¶TABLE(A1 | B1 / A2 | B2 | C2)¶)¶last');
    const saved = await reopened(page, 'b');
    expect(saved.tables.map((table) => [table.grid, table.fill, table.span])).toEqual([
      [[2000, 2000, 2000], 'FFFF00', 2],
      [[3000, 3000], undefined, undefined],
      [[2000, 2000, 2000], 'FFFF00', 2],
    ]);
    // The pages show it: the text of the pasted cells is on a page of the second editor.
    await expect.poll(() => page.evaluate(() => (document.querySelector('#host-b .bym-doc-text') || { textContent: '' }).textContent)).toContain('A2');

    // Whole rows selected inside the table: the rows, as a table.
    await select(page, 'a', at([2, 1, 0, 0], 0), at([2, 1, 2, 0], 2));
    const row = await clipboardEvent(page, 'a', 'cut');
    expect(row.text).toBe('A2\tB2\tC2');
    expect((await shape(page, 'a')).split('¶')[2]).toBe('TABLE(A1 | B1 /  |  | )');
    await select(page, 'a', at(3, 3));
    await clipboardEvent(page, 'a', 'paste', row);
    expect(await shape(page, 'a')).toBe(START.replace('A2 | B2 | C2', ' |  | ') + '¶TABLE(A2 | B2 | C2)¶');
  });

  test('what a copy does not carry is not cut: a note reference stays, with a word about it', async ({ page }) => {
    const bytes = await docx(p(r('note') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r(' after')) + p('next'), {
      footnotes: '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> the note</w:t></w:r></w:p></w:footnote>',
    });
    await mount(page, bytes, await target());
    await select(page, 'a', at(0, 2), at(0, 7));
    const cut = await clipboardEvent(page, 'a', 'cut', { text: 'what was there' });
    expect(cut.prevented).toBe(true);
    expect(cut.text).toBe('what was there');
    expect((await shape(page, 'a')).split('¶')[0]).toBe('note<footnoteReference> after');
    expect(await notices(page)).toContain('The selection holds something that cannot be cut (a note, an equation, a text box, a content control, a tracked change); it was left in place');
    expect(await page.evaluate(() => (window as any).__views.a.isModified)).toBe(false);
    // A copy of it is its text.
    const copied = await clipboardEvent(page, 'a', 'copy');
    expect(copied.text).toBe('te af');
    await select(page, 'b', at(0, 0));
    await clipboardEvent(page, 'b', 'paste', copied);
    expect((await shape(page, 'b')).split('¶')[0]).toBe('te aftarget');
  });

  test('review 4: a paste that waits for the pictures of the copy is made where it was asked for, or not at all', async ({ page }) => {
    await mount(page, await source(), await target());
    // The pictures of the first editor are read only when the test lets them.
    await page.evaluate(() => {
      const pictures = (window as any).__views.a.pictures;
      const read = pictures.readPart.bind(pictures);
      (window as any).__held = [];
      pictures.readPart = (part: string) => new Promise((resolve, reject) => (window as any).__held.push(() => read(part).then(resolve, reject)));
    });
    const release = () =>
      page.evaluate(async () => {
        for (const go of (window as any).__held.splice(0)) go();
        await new Promise((resolve) => setTimeout(resolve, 50));
      });
    const START_B = 'target¶TABLE(X1 | Y1)¶last';
    await select(page, 'a', at(0, 1), at(0, 3));
    let copied = await clipboardEvent(page, 'a', 'copy');
    await select(page, 'b', at(0, 6));
    await clipboardEvent(page, 'b', 'paste', copied);
    expect(await shape(page, 'b')).toBe(START_B);
    expect(await notices(page)).toEqual(['Reading the pictures of the copy; it is pasted when they are there']);
    // The caret goes elsewhere before the pictures are there: nothing is pasted, and that is said.
    await select(page, 'b', at(2, 0));
    await release();
    expect(await shape(page, 'b')).toBe(START_B);
    expect((await notices(page))[1]).toBe('The selection or the document changed while the pictures of the copy were read; nothing was pasted');

    // The document changes before the pictures are there (the caret is where the typing left it): nothing is pasted either.
    await select(page, 'a', at(0, 0), at(0, 3));
    copied = await clipboardEvent(page, 'a', 'copy');
    await select(page, 'b', at(2, 4));
    await clipboardEvent(page, 'b', 'paste', copied);
    await page.evaluate(() => (window as any).__views.b.insertText('!'));
    await release();
    expect(await shape(page, 'b')).toBe(START_B + '!');
    expect((await notices(page)).length).toBe(4);

    // Nothing changes while it waits: the paste is made at the caret.
    await select(page, 'a', at(0, 1), at(0, 4));
    copied = await clipboardEvent(page, 'a', 'copy');
    await select(page, 'b', at(0, 6));
    await clipboardEvent(page, 'b', 'paste', copied);
    expect(await shape(page, 'b')).toBe(START_B + '!');
    await release();
    await expect.poll(() => shape(page, 'b')).toBe('targetb<drawing>cd¶TABLE(X1 | Y1)¶last!');
    expect((await notices(page)).length).toBe(5);
  });

  test('review 6a: a cut whose copy cannot be written to the clipboard as more than text removes nothing that is more than text', async ({ page }) => {
    await mount(page, await source(), await target());
    // A clipboard that takes text and refuses HTML.
    const cut = (editor: Editor) =>
      page.evaluate((name) => {
        const held: { [type: string]: string } = {};
        const data = {
          setData: (type: string, value: string) => {
            if (type === 'text/html') throw new Error('not here');
            held[type] = value;
          },
          getData: (type: string) => held[type] || '',
          items: [],
          files: [],
        };
        const event = new Event('cut', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
        document.querySelector(`#host-${name} .bym-doc-input`)!.dispatchEvent(event);
        return held['text/plain'];
      }, editor);
    await select(page, 'a', at(0, 1), at(0, 3));
    expect(await cut('a')).toBe('bc');
    expect(await shape(page, 'a')).toBe(START);
    expect(await notices(page)).toEqual(['The clipboard cannot be written here; nothing was cut']);
    // Text that is nothing but text is cut: the clipboard holds all of it.
    await select(page, 'a', at(3, 0), at(3, 2));
    expect(await cut('a')).toBe('en');
    expect((await shape(page, 'a')).split('¶')[3]).toBe('d');
  });

  test('review 2: a cut over the end of a paragraph, pasted back with the keys of the editor, leaves the paragraphs formatted as they were', async ({ page }) => {
    const bytes = await docx(p('abc', '<w:jc w:val="center"/>') + p('def', '<w:jc w:val="right"/><w:ind w:left="720"/>') + p('tail'));
    await mount(page, bytes, await target());
    const formats = () => page.evaluate(() => (window as any).__views.a.document.body.blocks.map((block: any) => `${(window as any).BymOffice.paragraphText(block)}:${block.props.align || ''}:${block.props.indent ? block.props.indent.left : ''}`));
    expect(await formats()).toEqual(['abc:center:', 'def:right:720', 'tail::']);
    await select(page, 'a', at(0, 1), at(1, 1));
    const cut = await clipboardEvent(page, 'a', 'cut');
    expect(cut.text).toBe('bc\nd');
    expect(await formats()).toEqual(['aef:center:', 'tail::']);
    await clipboardEvent(page, 'a', 'paste', cut);
    expect(await formats()).toEqual(['abc:center:', 'def:right:720', 'tail::']);
  });

  test('the Paste button of the touch bubble knows a copy of the page by the HTML of the clipboard', async ({ page, browserName, isMobile }) => {
    test.skip(!isMobile || browserName !== 'chromium', 'touch events can only be made up in Chromium, on the phone profiles');
    await mount(page, await source(), await target());
    await select(page, 'a', at(0, 1), at(0, 3));
    const copied = await clipboardEvent(page, 'a', 'copy');
    // The clipboard of the system, as the button reads it: the text and the HTML of the copy.
    await page.evaluate((content) => {
      const item = { types: ['text/plain', 'text/html'], getType: (type: string) => Promise.resolve(new Blob([type === 'text/html' ? content.html : content.text], { type })) };
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { read: () => Promise.resolve([item]), readText: () => Promise.resolve(content.text) } });
    }, copied);
    // A long press on the word of the first paragraph of the second editor selects it and brings the bubble.
    const word = await page.evaluate(() => {
      const view = (window as any).__views.b;
      const box = view.lines.caret({ story: 'body', path: [0], offset: 3 }, 'start', 0);
      const element = document.querySelector(`#host-b .bym-doc-page[data-page="${box.page + 1}"]`)!;
      const rect = element.getBoundingClientRect();
      const laid = view.pages[box.page];
      return { x: rect.left + (box.x * rect.width) / laid.width, y: rect.top + ((box.top + box.height / 2) * rect.height) / laid.height };
    });
    await page.evaluate(
      ([x, y]) => {
        const target = document.elementFromPoint(x, y)!;
        const touch = new Touch({ identifier: 1, target, clientX: x, clientY: y });
        target.dispatchEvent(new TouchEvent('touchstart', { touches: [touch], targetTouches: [touch], changedTouches: [touch], bubbles: true, cancelable: true }));
        (window as any).__touch = { target, touch };
      },
      [word.x, word.y],
    );
    await page.waitForFunction(() => (window as any).__views.b.getSelectedText() === 'target', undefined, { timeout: 3000 });
    await page.evaluate(() => {
      const { target, touch } = (window as any).__touch;
      target.dispatchEvent(new TouchEvent('touchend', { touches: [], targetTouches: [], changedTouches: [touch], bubbles: true, cancelable: true }));
    });
    await page.locator('#host-b .bym-doc-bubble-paste').tap();
    await expect.poll(() => shape(page, 'b')).toBe('b<drawing>c¶TABLE(X1 | Y1)¶last');
    expect(await notices(page)).toEqual([]);
  });

  test('the keys: Ctrl/Cmd+C, X and V carry a picture and a link through the real clipboard', async ({ page, context, browserName, isMobile }) => {
    test.skip(browserName !== 'chromium' || isMobile, 'only Chromium lets a test use the system clipboard');
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await mount(page, await source(), await target());
    const key = await page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));
    const focus = async (editor: Editor) => {
      await page.evaluate((name) => (window as any).__views[name].focus(), editor);
      await expect(page.locator(`#host-${editor} .bym-doc-input`)).toBeFocused();
    };
    await focus('a');
    await select(page, 'a', at(0, 1), at(0, 3));
    await page.keyboard.press(`${key}+c`);
    await focus('b');
    await select(page, 'b', at(0, 6));
    await page.keyboard.press(`${key}+v`);
    await expect.poll(() => shape(page, 'b')).toBe('targetb<drawing>c¶TABLE(X1 | Y1)¶last');
    // A cut of the link in the first editor, pasted at the end of the second.
    await focus('a');
    await select(page, 'a', at(1, 4), at(1, 12));
    await page.keyboard.press(`${key}+x`);
    expect((await shape(page, 'a')).split('¶')[1]).toBe('see  and page <begin>( PAGE )(<separate>)1<end>.');
    await focus('b');
    await select(page, 'b', at(2, 4));
    await page.keyboard.press(`${key}+v`);
    await expect.poll(() => shape(page, 'b')).toBe('targetb<drawing>c¶TABLE(X1 | Y1)¶last[hyperlink:the link]');
    const saved = await reopened(page, 'b');
    expect(saved.links).toEqual(['https://example.com/']);
    expect(saved.pictures.map((picture) => picture.bytes.length)).toEqual([PICTURE.length]);
    // A paste that came before the bytes of the picture were read says that it waits; nothing else is said.
    expect((await notices(page)).filter((message: string) => message.indexOf('Reading the pictures of the copy') !== 0)).toEqual([]);
  });
});
