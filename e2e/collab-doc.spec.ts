import type { APIRequestContext, BrowserContext, Page } from '@playwright/test';
import { documentText, open } from '../src/index';
import { buildDocx } from '../test/docx-fixtures';
import { expect, test } from './helpers';

// MountOptions.collab for a document (docx): two pages of one browser open the same file and
// edit it together. The room is held by the first page (a relay of the page, reached over a
// BroadcastChannel) or by the relay of the test server (a WebSocket). No screenshots: what is
// shown is read from the model, from the elements of the carets and from the bar of the room.

const PAGE = '/e2e/pages/mount.html';
const RED = '#D93025';
const BLUE = '#1A73E8';

const paragraph = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';

/** A document of three paragraphs. */
const sample = (settings?: string): Promise<Uint8Array> => buildDocx({ body: paragraph('First paragraph.') + paragraph('Second paragraph.') + paragraph('Third paragraph.'), styles: STYLES, settings });

/** A name no other test uses: for the room and for the file. */
function unique(): string {
  const info = test.info();
  return `${info.project.name}-${info.testId}-${info.repeatEachIndex}`.replace(/[^A-Za-z0-9_-]/g, '_');
}

async function seed(request: APIRequestContext, url: string, bytes: Uint8Array): Promise<void> {
  const answer = await request.put(url, { data: Buffer.from(bytes), headers: { 'content-type': 'application/octet-stream' } });
  expect(answer.status(), `seeding ${url}`).toBe(204);
}

interface Spec {
  file: string;
  room: string;
  /** "host": this page holds the room; "tab": it reaches the room of another page; "socket": the relay of the server. */
  via: 'host' | 'tab' | 'socket';
  name?: string;
  color?: string;
  readOnly?: boolean;
  saveTo?: boolean;
  autoSave?: boolean;
  trackChanges?: boolean;
}

function mount(page: Page, spec: Spec): Promise<string> {
  return page.evaluate(async (given) => {
    const w = window as any;
    w.__notices = [];
    w.__states = [];
    const api = w.BymOffice;
    let collab: any;
    if (given.via === 'socket') collab = { url: `ws://${location.host}/e2e/collab/${given.room}` };
    else {
      if (given.via === 'host') {
        w.__relay = api.collab.memoryServer();
        w.__relay.serve(given.room);
      }
      w.__transport = api.collab.broadcastTransport(given.room);
      collab = { transport: w.__transport };
    }
    collab.user = { name: given.name, color: given.color };
    collab.onState = (state: string) => w.__states.push(state);
    const options: any = {
      file: given.file,
      locale: 'en',
      collab,
      onNotice: (notice: { kind: string; message: string }) => w.__notices.push(`${notice.kind}: ${notice.message}`),
    };
    if (given.readOnly) options.readOnly = true;
    if (given.saveTo) options.saveTo = true;
    if (given.autoSave) options.autoSave = { delayMs: 300, maxWaitMs: 2000 };
    if (given.trackChanges) options.trackChanges = true;
    try {
      w.__mounted = await api.mount(document.getElementById('editor'), options);
      return w.__mounted.kind;
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  }, spec);
}

const online = (page: Page) => expect.poll(() => page.evaluate(() => (window as any).__mounted.collab.state), { message: 'the page is online' }).toBe('online');
const peers = (page: Page): Promise<number> => page.evaluate(() => (window as any).__mounted.collab.peers.length);
const text = (page: Page): Promise<string> => page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document));
const notices = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__notices);
const settled = (page: Page) => expect.poll(() => page.evaluate(() => (window as any).__mounted.collab.pending), { message: 'nothing waits for the relay' }).toBe(0);
/** Puts the caret at a place of the body and types there, as the public API of the view does it. */
const typeAt = (page: Page, row: number, offset: number, typed: string): Promise<boolean> =>
  page.evaluate(
    ([r, o, t]) => {
      const view = (window as any).__mounted.view;
      view.setSelection({ story: 'body', path: [r], offset: o });
      return view.insertText(t);
    },
    [row, offset, typed] as [number, number, string],
  );

/** Two pages on one file and one room; the first one holds the room unless the relay of the server does. */
async function pair(context: BrowserContext, page: Page, request: APIRequestContext, options: { socket?: boolean; saveTo?: boolean; autoSave?: boolean; bytes?: Uint8Array; first?: Partial<Spec>; second?: Partial<Spec> } = {}): Promise<{ a: Page; b: Page; url: string; room: string }> {
  const room = unique();
  const url = `/e2e/saved/${room}.docx`;
  await seed(request, url, options.bytes || (await sample()));
  const b = await context.newPage();
  const problems: string[] = [];
  b.on('console', (message) => {
    if (message.type() === 'error') problems.push(message.text());
  });
  b.on('pageerror', (error) => problems.push(`${error.name}: ${error.message}`));
  (b as any).__problems = problems;
  await page.goto(PAGE);
  await b.goto(PAGE);
  expect(await mount(page, { file: url, room, via: options.socket ? 'socket' : 'host', name: 'Ada', color: RED, saveTo: options.saveTo, autoSave: options.autoSave, ...options.first })).toBe('document');
  expect(await mount(b, { file: url, room, via: options.socket ? 'socket' : 'tab', name: 'Ben', color: BLUE, saveTo: options.saveTo, autoSave: options.autoSave, ...options.second })).toBe('document');
  await online(page);
  await online(b);
  await expect.poll(() => peers(page), { message: 'the first page sees the second' }).toBe(1);
  await expect.poll(() => peers(b), { message: 'the second page sees the first' }).toBe(1);
  return { a: page, b, url, room };
}

const clean = (b: Page) => expect((b as any).__problems, 'console errors and uncaught exceptions of the second page').toEqual([]);

for (const socket of [false, true]) {
  test(`what one page types in a document, the other one shows (${socket ? 'the relay of the server' : 'a room in the first page'})`, async ({ page, context, request }) => {
    const { a, b } = await pair(context, page, request, { socket });
    expect(await typeAt(a, 0, 5, ' typed')).toBe(true);
    await expect.poll(() => text(b), { message: 'the second page has the text' }).toBe('First typed paragraph.\nSecond paragraph.\nThird paragraph.');
    expect(await typeAt(b, 2, 0, 'B: ')).toBe(true);
    await expect.poll(() => text(a), { message: 'the first page has the text of the second' }).toBe('First typed paragraph.\nSecond paragraph.\nB: Third paragraph.');
    await settled(a);
    await settled(b);
    expect(await notices(a), 'nothing was said to the first page').toEqual([]);
    expect(await notices(b), 'nothing was said to the second page').toEqual([]);
    clean(b);
  });
}

test('keys typed into the pages of one page arrive in the other, Enter and Backspace too', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ story: 'body', path: [1], offset: 6 });
    view.focus();
  });
  await a.keyboard.type(' new');
  await a.keyboard.press('Enter');
  await a.keyboard.type('x');
  await a.keyboard.press('Backspace');
  await expect.poll(() => text(b), { message: 'the second page has what was typed' }).toBe('First paragraph.\nSecond new\n paragraph.\nThird paragraph.');
  await settled(a);
  expect(await text(a)).toBe(await text(b));
  clean(b);
});

test('two who type at the same place: both texts stay, and both pages show the same', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  // Both at once: neither has heard of the other when it types.
  await Promise.all([typeAt(a, 0, 5, 'AAA'), typeAt(b, 0, 5, 'bbb')]);
  await settled(a);
  await settled(b);
  await expect.poll(async () => (await text(a)) === (await text(b)), { message: 'the pages agree' }).toBe(true);
  const first = (await text(a)).split('\n')[0];
  expect(['FirstAAAbbb paragraph.', 'FirstbbbAAA paragraph.'], `"${first}" holds both texts whole`).toContain(first);
  clean(b);
});

test('the caret of the other one is shown with its name, in its colour, and follows text that comes in before it', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await b.evaluate(() => (window as any).__mounted.view.setSelection({ story: 'body', path: [1], offset: 6 }));
  const caret = a.locator('#editor .bym-doc-peer');
  await expect(caret).toHaveCount(1);
  await expect(caret.locator('.bym-doc-peer-name')).toHaveText('Ben');
  expect(await caret.evaluate((element) => getComputedStyle(element).backgroundColor)).toBe('rgb(26, 115, 232)');
  const before = await caret.evaluate((element) => element.getBoundingClientRect().left);
  // Text typed before it in its line: the caret of the other moves on with the text it stands in.
  expect(await typeAt(a, 1, 0, 'Wider wider ')).toBe(true);
  await expect.poll(() => caret.evaluate((element) => element.getBoundingClientRect().left), { message: 'the caret moved to the right' }).toBeGreaterThan(before + 20);
  // What the first page shows of the room: its state and how many are in the file.
  const bar = a.locator('#editor .bym-collab');
  await expect(bar).toHaveAttribute('data-state', 'online');
  await expect(bar).toHaveAttribute('data-people', '2');
  clean(b);
});

test('the caret of this page stays with its text when the other one types before it', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await a.evaluate(() => (window as any).__mounted.view.setSelection({ story: 'body', path: [0], offset: 10 }));
  expect(await typeAt(b, 0, 0, '12345')).toBe(true);
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.view.getSelection().focus.offset), { message: 'the caret moved on with the text' }).toBe(15);
  // And a paragraph that came in above it: the caret is one paragraph further down.
  await b.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ story: 'body', path: [0], offset: 0 });
    view.insertText('above\n');
  });
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.view.getSelection().focus.path[0]), { message: 'the caret is in the paragraph it was in' }).toBe(1);
  clean(b);
});

test('the caret of this page stays in its paragraph when the other one puts that paragraph back by undo', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  // The other one writes over a word of the second paragraph, and takes that back: the paragraph is put back as a whole.
  await b.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ anchor: { story: 'body', path: [1], offset: 0 }, focus: { story: 'body', path: [1], offset: 7 } });
    view.insertText('New ');
  });
  await settled(b);
  await expect.poll(() => text(a)).toBe('First paragraph.\nNew paragraph.\nThird paragraph.');
  await a.evaluate(() => (window as any).__mounted.view.setSelection({ story: 'body', path: [1], offset: 4 }));
  expect(await b.evaluate(() => (window as any).__mounted.view.undo())).toBe(true);
  await expect.poll(() => text(a)).toBe('First paragraph.\nSecond paragraph.\nThird paragraph.');
  const focus = await a.evaluate(() => (window as any).__mounted.view.getSelection().focus);
  expect(focus.path).toEqual([1]);
  // The caret follows the text: it stood before "paragraph" and stands there still.
  expect(focus.offset).toBe(7);
  clean(b);
});

test('undo takes back the own step and leaves what the other one typed elsewhere', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  expect(await typeAt(a, 0, 0, 'mine ')).toBe(true);
  await settled(a);
  expect(await typeAt(b, 2, 0, 'theirs ')).toBe(true);
  await expect.poll(() => text(a)).toBe('mine First paragraph.\nSecond paragraph.\ntheirs Third paragraph.');
  expect(await a.evaluate(() => (window as any).__mounted.view.undo())).toBe(true);
  await expect.poll(() => text(b), { message: 'the second page lost what the first undid, and kept its own' }).toBe('First paragraph.\nSecond paragraph.\ntheirs Third paragraph.');
  // The other one cannot undo what is not its own: its undo takes back its own step.
  expect(await b.evaluate(() => (window as any).__mounted.view.undo())).toBe(true);
  await expect.poll(() => text(a)).toBe('First paragraph.\nSecond paragraph.\nThird paragraph.');
  expect(await notices(a)).toEqual([]);
  clean(b);
});

test('what cannot be shared is refused and said, and changes nothing', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  const before = await text(a);
  const done = await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 5 } });
    return view.setTrackChanges(true);
  });
  expect(done, 'tracking was not turned on').toBe(false);
  expect((await notices(a)).filter((line) => /edited together/.test(line)).length, 'that was said').toBe(1);
  expect(await text(a)).toBe(before);
  expect(await text(b)).toBe(before);
  // What can be shared still is.
  expect(await a.evaluate(() => (window as any).__mounted.view.toggleFormat('bold'))).toBe(true);
  await settled(a);
  clean(b);
});

/** A document whose second paragraph anchors a text box Word writes (a DrawingML Choice and its VML Fallback) of two paragraphs. */
const boxSample = (): Promise<Uint8Array> => {
  const inner = paragraph('Box first line.') + paragraph('Box second line.');
  const head =
    '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>2540000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="2286000" cy="914400"/><wp:wrapNone/><wp:docPr id="7" name="Box"/>';
  const box =
    '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
    head +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
    `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:shape style="width:180pt;height:72pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  return buildDocx({ body: paragraph('First paragraph.') + `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${box}</w:p>` + paragraph('Third paragraph.'), styles: STYLES });
};

/** The text of the text box of boxSample(), a line a paragraph. */
const boxText = (page: Page): Promise<string> =>
  page.evaluate(() => {
    const api = (window as any).BymOffice;
    const doc = (window as any).__mounted.document;
    return api.docOps.containerBlocks(doc, 'body', [1, -1, 0, 0]).map((block: any) => api.paragraphText(block)).join('\n');
  });

/** Puts the caret at a place of the text box of boxSample() and types there. */
const typeInBox = (page: Page, row: number, offset: number, typed: string): Promise<boolean> =>
  page.evaluate(
    ([r, o, t]) => {
      const view = (window as any).__mounted.view;
      view.setSelection({ story: 'body', path: [1, -1, 0, 0, r], offset: o });
      return view.insertText(t);
    },
    [row, offset, typed] as [number, number, string],
  );

test('what one page types in a text box, the other one shows; both type in it at once and both texts stay; undo takes back the own', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { bytes: await boxSample() });
  expect(await typeInBox(a, 0, 3, ' typed')).toBe(true);
  await expect.poll(() => boxText(b), { message: 'the second page has the text in the box' }).toBe('Box typed first line.\nBox second line.');
  // Both at once, in the same box.
  await Promise.all([typeInBox(a, 1, 0, 'A: '), typeInBox(b, 0, 0, 'B: ')]);
  await settled(a);
  await settled(b);
  await expect.poll(async () => (await boxText(a)) === (await boxText(b)), { message: 'the pages agree' }).toBe(true);
  expect(await boxText(a)).toBe('B: Box typed first line.\nA: Box second line.');
  // The text around the box is as it was.
  expect((await text(a)).split('\n')[0]).toBe('First paragraph.');
  // Undo takes back what this page typed, in the box, and leaves what the other one typed.
  expect(await b.evaluate(() => (window as any).__mounted.view.undo())).toBe(true);
  await expect.poll(() => boxText(a), { message: 'the first page lost what the second undid' }).toBe('Box typed first line.\nA: Box second line.');
  expect(await notices(a), 'nothing was said to the first page').toEqual([]);
  expect(await notices(b), 'nothing was said to the second page').toEqual([]);
  clean(b);
});

test('keys typed in a text box arrive in the other page, Enter and Backspace too; the caret of the other one in the box moves with what was deleted before it', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { bytes: await boxSample() });
  await b.evaluate(() => (window as any).__mounted.view.setSelection({ story: 'body', path: [1, -1, 0, 0, 1], offset: 10 }));
  await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ story: 'body', path: [1, -1, 0, 0, 1], offset: 0 });
    view.focus();
  });
  await a.keyboard.type('xy');
  await a.keyboard.press('Enter');
  await a.keyboard.press('Backspace');
  await a.keyboard.press('Backspace');
  await expect.poll(() => boxText(b), { message: 'the second page has what was typed in the box' }).toBe('Box first line.\nxBox second line.');
  await settled(a);
  // The caret of the second page stood right after "second" in the box: it still does.
  const focus = await b.evaluate(() => (window as any).__mounted.view.getSelection().focus);
  expect(focus).toEqual({ story: 'body', path: [1, -1, 0, 0, 1], offset: 11 });
  expect(await boxText(a)).toBe(await boxText(b));
  clean(b);
});

test('saved to the same place after both typed in a text box: the file holds the box in both of its branches, and who opens it holds the same', async ({ page, context, request }) => {
  const { a, b, url, room } = await pair(context, page, request, { saveTo: true, bytes: await boxSample() });
  expect(await typeInBox(a, 0, 0, 'A ')).toBe(true);
  expect(await typeInBox(b, 1, 0, 'B ')).toBe(true);
  await settled(a);
  await settled(b);
  await expect.poll(async () => (await boxText(a)) === (await boxText(b))).toBe(true);
  expect(await a.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
  const stored = new Uint8Array(await (await request.get(url)).body());
  const xml = Buffer.from(await (await open(stored)).package!.readPart('/word/document.xml')).toString('utf8');
  expect(xml.split('A Box first line.').length - 1, 'the Choice and the Fallback').toBe(2);
  const c = await context.newPage();
  await c.goto(PAGE);
  expect(await mount(c, { file: url, room, via: 'tab', name: 'Cem' })).toBe('document');
  await online(c);
  expect(await boxText(c)).toBe(await boxText(a));
  expect(await typeInBox(c, 0, 0, 'C ')).toBe(true);
  await expect.poll(() => boxText(a), { message: 'the first page has what the third typed in the box' }).toBe('C A Box first line.\nB Box second line.');
  clean(b);
});

/** The comments a page lists: id, author, text, the comment answered, resolved. */
const commentsOf = (page: Page): Promise<[number, string, string, number | null, boolean][]> =>
  page.evaluate(() => (window as any).__mounted.view.comments().map((c: any) => [c.id, c.author, c.text, c.parent === undefined ? null : c.parent, c.done]));

test('a comment added in one page is in the other, under its name; answered, edited and resolved there, and the panel follows', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  // The panel of the second page is open before anything is said.
  expect(await b.evaluate(() => (window as any).__mounted.view.setCommentsPanel(true))).toBe(true);
  const id = await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 5 } });
    return view.addComment('Why first?');
  });
  expect(typeof id, 'a comment was added').toBe('number');
  await expect.poll(() => commentsOf(b), { message: 'the second page has the comment, by the name of the first in the room' }).toEqual([[id, 'Ada', 'Why first?', null, false]]);
  // The panel of the second page lists it, with what can be done to it.
  const panel = b.locator('#editor .bym-doc-cpanel');
  await expect(panel.locator('.bym-doc-cpanel-item')).toHaveCount(1);
  await expect(panel.locator('.bym-doc-cpanel-author').first()).toHaveText('Ada');
  // The buttons of the panel are those of the thread that is selected.
  expect(await b.evaluate((given) => (window as any).__mounted.view.showComment(given), id)).toBe(true);
  await expect(panel.locator('button[data-action="reply"]')).toHaveCount(1);
  // Answered there, edited and resolved here: both pages hold the same.
  expect(typeof (await b.evaluate((given) => (window as any).__mounted.view.replyToComment(given, 'Because.'), id))).toBe('number');
  await expect.poll(() => commentsOf(a)).toHaveLength(2);
  expect(await a.evaluate((given) => (window as any).__mounted.view.editComment(given, 'Why first, really?'), id)).toBe(true);
  expect(await a.evaluate((given) => (window as any).__mounted.view.resolveComment(given, true), id)).toBe(true);
  await settled(a);
  await settled(b);
  await expect.poll(async () => JSON.stringify(await commentsOf(a)) === JSON.stringify(await commentsOf(b)), { message: 'the pages agree' }).toBe(true);
  const list = await commentsOf(b);
  expect(list.map((c) => [c[1], c[2], c[4]])).toEqual([
    ['Ada', 'Why first, really?', true],
    ['Ben', 'Because.', true],
  ]);
  // Removed in the second page: gone from the first, and from its panel.
  expect(await a.evaluate(() => (window as any).__mounted.view.setCommentsPanel(true))).toBe(true);
  expect(await b.evaluate((given) => (window as any).__mounted.view.removeComment(given), id)).toBe(true);
  await expect.poll(() => commentsOf(a)).toEqual([]);
  await expect(a.locator('#editor .bym-doc-cpanel .bym-doc-cpanel-item')).toHaveCount(0);
  expect(await notices(a)).toEqual(expect.not.arrayContaining([expect.stringMatching(/edited together/)]));
  clean(b);
});

test('two pages that add a comment at the same time both keep theirs, under numbers both pages agree on', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  const add = (who: Page, row: number, said: string) =>
    who.evaluate(
      ([r, t]) => {
        const view = (window as any).__mounted.view;
        view.setSelection({ anchor: { story: 'body', path: [r], offset: 0 }, focus: { story: 'body', path: [r], offset: 5 } });
        return view.addComment(t);
      },
      [row, said] as [number, string],
    );
  const made = await Promise.all([add(a, 0, 'from Ada'), add(b, 2, 'from Ben')]);
  expect(made.every((one) => typeof one === 'number')).toBe(true);
  await settled(a);
  await settled(b);
  await expect.poll(async () => JSON.stringify(await commentsOf(a)) === JSON.stringify(await commentsOf(b)), { message: 'the pages agree' }).toBe(true);
  const list = await commentsOf(a);
  expect(list.map((c) => c[2]).sort()).toEqual(['from Ada', 'from Ben']);
  expect(new Set(list.map((c) => c[0])).size, 'two numbers').toBe(2);
  clean(b);
});

test('a page that only reads lists the comments and offers nothing to change them', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { second: { readOnly: true } });
  await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ anchor: { story: 'body', path: [1], offset: 0 }, focus: { story: 'body', path: [1], offset: 6 } });
    return view.addComment('seen by both');
  });
  await expect.poll(() => commentsOf(b)).toHaveLength(1);
  expect(await b.evaluate(() => (window as any).__mounted.view.canComment)).toBe(false);
  expect(await b.evaluate(() => (window as any).__mounted.view.setCommentsPanel(true))).toBe(true);
  const panel = b.locator('#editor .bym-doc-cpanel');
  await expect(panel.locator('.bym-doc-cpanel-item')).toHaveCount(1);
  // Selected, the thread shows no buttons either.
  const id = await b.evaluate(() => (window as any).__mounted.view.comments()[0].id);
  expect(await b.evaluate((given) => (window as any).__mounted.view.showComment(given), id)).toBe(true);
  await expect(panel.locator('button[data-action="reply"]')).toHaveCount(0);
  await expect(panel.locator('button[data-action="delete"]')).toHaveCount(0);
  clean(b);
});

test('a document that tracks its changes is edited alone, and that is said', async ({ page, request }) => {
  const room = unique();
  const url = `/e2e/saved/${room}.docx`;
  await seed(request, url, await sample('<w:trackRevisions/>'));
  await page.goto(PAGE);
  expect(await mount(page, { file: url, room, via: 'host', name: 'Ada', color: RED })).toBe('document');
  expect(await page.evaluate(() => (window as any).__mounted.collab)).toBeUndefined();
  expect((await notices(page)).some((line) => /tracks its changes/.test(line)), 'said').toBe(true);
  await expect(page.locator('#editor .bym-collab')).toHaveCount(0);
});

test('a page that only reads follows, and sends nothing', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { second: { readOnly: true } });
  expect(await typeAt(a, 0, 0, 'seen ')).toBe(true);
  await expect.poll(() => text(b)).toBe('seen First paragraph.\nSecond paragraph.\nThird paragraph.');
  expect(await typeAt(b, 0, 0, 'nope'), 'nothing is typed in a view that only reads').toBe(false);
  expect(await a.evaluate(() => (window as any).__mounted.collab.people.filter((who: { readOnly: boolean }) => who.readOnly).length)).toBe(1);
  clean(b);
});

test('saved to the same place: the file holds what both typed and says where the room was, and who opens it joins there', async ({ page, context, request }) => {
  const { a, b, url, room } = await pair(context, page, request, { saveTo: true });
  expect(await typeAt(a, 0, 0, 'A ')).toBe(true);
  expect(await typeAt(b, 2, 0, 'B ')).toBe(true);
  await settled(a);
  await settled(b);
  await expect.poll(async () => (await text(a)) === (await text(b))).toBe(true);
  expect(await a.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
  const stored = new Uint8Array(await (await request.get(url)).body());
  const file = await open(stored);
  expect(documentText(await file.document!())).toBe(await text(a));
  // Someone who opens the stored file joins the room at its step, and gets what is typed from there on.
  const c = await context.newPage();
  await c.goto(PAGE);
  expect(await mount(c, { file: url, room, via: 'tab', name: 'Cem' })).toBe('document');
  await online(c);
  expect(await typeAt(a, 1, 0, 'later ')).toBe(true);
  await expect.poll(() => text(c), { message: 'the third page has what came after the save' }).toBe(await text(a));
  clean(b);
});

test('saved to the same place with comments: the file holds them, and who opens it joins with them and goes on commenting', async ({ page, context, request }) => {
  const { a, url, room } = await pair(context, page, request, { saveTo: true });
  const id = await a.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 5 } });
    return view.addComment('kept in the file');
  });
  expect(typeof id).toBe('number');
  await settled(a);
  expect(await a.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
  const stored = new Uint8Array(await (await request.get(url)).body());
  const doc = await (await open(stored)).document!();
  expect(Array.from(doc.comments ? doc.comments.items.values() : []).map((comment) => comment.author)).toEqual(['Ada']);
  const c = await context.newPage();
  await c.goto(PAGE);
  expect(await mount(c, { file: url, room, via: 'tab', name: 'Cem' })).toBe('document');
  await online(c);
  expect(await commentsOf(c)).toEqual(await commentsOf(a));
  expect(typeof (await c.evaluate((given) => (window as any).__mounted.view.replyToComment(given, 'from Cem'), id))).toBe('number');
  await expect.poll(() => commentsOf(a)).toHaveLength(2);
  expect(await commentsOf(a)).toEqual(await commentsOf(c));
});
