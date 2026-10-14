// Links and the outline of a PDF: the addresses that may be written, the
// annotations and the outline the file holds, the links and headings found
// on the laid out pages of a document and on the printed pages of a
// workbook. What the engine writes is read back by the reader of the tests
// (pdf-fixtures.ts), by a reader in Python written from the specification
// (tools/pdf_links_ref.py) and, where the machine has it, by Poppler.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { test } from 'node:test';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { docPdfLinks } from '../src/pdf/doc-links';
import { PdfDocument, buildPdf, type PdfPage } from '../src/pdf/document';
import { pdfSettings, reportLinks } from '../src/pdf/export';
import { MAX_FILE_LINKS, MAX_LINK_RECTS, MAX_OUTLINE_ENTRIES, MAX_PAGE_LINKS, MAX_URI_BYTES, outlineTitle, outlineTree, pdfLinkTarget, pdfLinkUri, type PdfLink, type PdfOutlineEntry } from '../src/pdf/links';
import { parseLocation, sheetPdfLinks } from '../src/pdf/sheet-links';
import { joinParts } from '../src/pdf/writer';
import { createSheet, emptyStyleSheet, type Cell, type Sheet, type Workbook } from '../src/model/sheet/types';
import { planSheetPages } from '../src/ui/sheet-print';
import { STYLES, lay, p, r, sectPr } from './doc-layout-fixtures';
import { Ref as RefClass, checkPdf, hasPoppler, popplerTest, requirePoppler, type Dict, type Ref, type Str } from './pdf-fixtures';
import type { Block } from '../src/model/doc/types';
import { readRtf } from '../src/formats/rtf/document';
import { DocLayout } from '../src/layout/doc/layout';
import { fakeMeasurer } from './doc-layout-fixtures';

const FAKE_METRICS = (font: { size: number }) => ({ ascent: font.size * 0.8, descent: font.size * 0.2 });

// ---------------------------------------------------------------- helpers

const DATE = new Date(0);

/** The file as tools/pdf_links_ref.py reads it. */
interface RefLinks {
  pages: { annots: { rect: number[]; uri?: string; dest?: [number, number | null, number | null]; flags: number; border: number[] }[] }[];
  outline: RefEntry[];
  outlineCount: number | null;
  pageMode: string | null;
}
interface RefEntry {
  title: string;
  dest: [number, number | null, number | null] | null;
  count: number | null;
  kids: RefEntry[];
}

let pythonFound: boolean | undefined;
function hasPython(): boolean {
  if (pythonFound === undefined) {
    try {
      execFileSync('python3', ['-I', '-c', 'import json'], { stdio: 'pipe' });
      pythonFound = true;
    } catch {
      pythonFound = false;
    }
  }
  return pythonFound;
}

let serial = 0;
function withFile<T>(file: Uint8Array, run: (path: string) => T): T {
  const folder = join(process.cwd(), `tmp-pdf-links-${process.pid}-${++serial}`);
  mkdirSync(folder, { recursive: true });
  try {
    const path = join(folder, 'file.pdf');
    writeFileSync(path, file);
    return run(path);
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

/** The links and the outline as the reader in Python finds them. */
function refLinks(file: Uint8Array): RefLinks {
  const script = join(process.cwd(), 'test', 'tools', 'pdf_links_ref.py');
  return withFile(file, (path) => JSON.parse(execFileSync('python3', ['-I', script, path], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 })) as RefLinks);
}

/** A test that reads the file with the reader in Python; skipped, with the reason, where there is no Python (fails with BYM_REQUIRE_POPPLER=1). */
function pythonTest(name: string, run: () => Promise<void> | void): void {
  if (hasPython() || requirePoppler()) {
    test(name, async () => {
      if (!hasPython()) throw new Error('BYM_REQUIRE_POPPLER=1, and python3 is not installed: the independent PDF checks cannot run');
      await run();
    });
    return;
  }
  test(name, { skip: 'python3 not found: the independent reading of links skipped' }, () => undefined);
}

/** What Poppler makes of the file as XML: the outline, and the text of the pages with the links over it. */
function popplerXml(file: Uint8Array): string {
  return withFile(file, (path) => execFileSync('pdftohtml', ['-xml', '-i', '-hidden', '-stdout', '-q', path], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }));
}

/** The words of the pages and where they are, as `pdftotext -bbox` reads them: points from the top left corner. */
function popplerBoxes(file: Uint8Array): { page: number; word: string; x1: number; y1: number; x2: number; y2: number }[] {
  const html = withFile(file, (path) => execFileSync('pdftotext', ['-bbox', '-enc', 'UTF-8', path, '-'], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }));
  const out: { page: number; word: string; x1: number; y1: number; x2: number; y2: number }[] = [];
  let page = -1;
  for (const line of html.split('\n')) {
    if (/<page /.test(line)) page++;
    const word = /<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*)<\/word>/.exec(line);
    if (word) out.push({ page, word: word[5].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'), x1: Number(word[1]), y1: Number(word[2]), x2: Number(word[3]), y2: Number(word[4]) });
  }
  return out;
}

const decode = (value: Str): string => {
  let out = '';
  for (let i = 0; i < value.bytes.length; i++) out += String.fromCharCode(value.bytes[i]);
  return out;
};

/** The annotations of the pages, as the reader of the tests finds them. */
function annotations(file: Uint8Array): Dict[][] {
  const checked = checkPdf(file);
  const root = checked.objects.get((checked.trailer.Root as Ref).id)!.value as Dict;
  const pages = checked.objects.get((root.Pages as Ref).id)!.value as Dict;
  return (pages.Kids as Ref[]).map((kid) => {
    const page = checked.objects.get(kid.id)!.value as Dict;
    // /Annots is an array, or a reference to one (a page with links into the file).
    const annots = page.Annots instanceof RefClass ? (checked.objects.get(page.Annots.id)!.value as Ref[]) : (page.Annots as Ref[] | undefined);
    return (annots || []).map((ref) => checked.objects.get(ref.id)!.value as Dict);
  });
}

const page = (links?: PdfLink[], more: Partial<PdfPage> = {}): PdfPage => ({ width: 600, height: 800, fill: [255, 255, 255], links, ...more });
const rect = (x: number, y: number, width = 50, height = 20) => ({ x, y, width, height });

async function fileOf(pages: PdfPage[], outline?: PdfOutlineEntry[]): Promise<{ file: Uint8Array; pdf: PdfDocument }> {
  const pdf = new PdfDocument({ date: DATE });
  const parts: Uint8Array[] = [];
  for (const one of pages) {
    await pdf.addPage(one);
    parts.push(...pdf.takeParts());
  }
  if (outline) pdf.setOutline(outline);
  await pdf.finish();
  parts.push(...pdf.takeParts());
  return { file: joinParts(parts), pdf };
}

// ---------------------------------------------------------------- addresses

test('pdf links: web and mail addresses pass, as printable ASCII', () => {
  assert.equal(pdfLinkUri('https://example.com/a?b=1#c'), 'https://example.com/a?b=1#c');
  assert.equal(pdfLinkUri('HTTP://Example.com'), 'http://Example.com');
  assert.equal(pdfLinkUri('mailto:someone@example.com?subject=Merhaba'), 'mailto:someone@example.com?subject=Merhaba');
  // What is around the address and the line ends in it are not part of it.
  assert.equal(pdfLinkUri('  https://exa\nmple.com/x\t '), 'https://example.com/x');
  // Characters beyond ASCII and those an address does not hold as they are: %XX of their UTF-8.
  assert.equal(pdfLinkUri('https://örnek.com.tr/ğüşıöç İ'), 'https://%C3%B6rnek.com.tr/%C4%9F%C3%BC%C5%9F%C4%B1%C3%B6%C3%A7%20%C4%B0');
  assert.equal(pdfLinkUri('https://example.com/a"b<c>d{f}|^`?q\\r'), 'https://example.com/a%22b%3Cc%3Ed%7Bf%7D%7C%5E%60?q%5Cr');
  assert.equal(pdfLinkUri('https://example.com/😀'), 'https://example.com/%F0%9F%98%80');
  assert.equal(pdfLinkUri('https://example.com/\ud800x'), 'https://example.com/%EF%BF%BDx', 'half a pair of surrogates is U+FFFD');
  // A % that starts %XX stays; one that does not is itself written as %25.
  assert.equal(pdfLinkUri('https://example.com/%41%zz%'), 'https://example.com/%41%25zz%25');
  // Parentheses are ASCII: the writer escapes them in the string.
  assert.equal(pdfLinkUri('https://example.com/(a)'), 'https://example.com/(a)');
  // A backslash before the query is a slash, as readers of addresses take it: the host is the one they read.
  assert.equal(pdfLinkUri('https:\\\\example.com\\x'), 'https://example.com/x');
  assert.equal(pdfLinkUri('https://example.com\\docs\\a.html'), 'https://example.com/docs/a.html');
  assert.equal(pdfLinkUri('http://evil.com\\@good.com/'), 'http://evil.com/@good.com/');
  assert.deepEqual(pdfLinkTarget('http://a\\@'), { uri: 'http://a/@', host: 'a' });
  assert.deepEqual(pdfLinkTarget('https://@host'), { uri: 'https://@host', host: 'host' }, 'an empty user is no user; the host is there');
  // The user is not the host: the host is what follows the last "@", the port is not part of it.
  assert.deepEqual(pdfLinkTarget('http://good.com@evil.com/x'), { uri: 'http://good.com@evil.com/x', host: 'evil.com' });
  assert.deepEqual(pdfLinkTarget('https://u:p@a@b.com:8080/x'), { uri: 'https://u:p@a@b.com:8080/x', host: 'b.com' });
  assert.deepEqual(pdfLinkTarget('http://[::1]:80/'), { uri: 'http://[::1]:80/', host: '[::1]' });
  assert.deepEqual(pdfLinkTarget('mailto:a@b.com'), { uri: 'mailto:a@b.com' });
});

test('pdf links: an http address the reader would find no host in, or another host, is refused', () => {
  for (const target of ['http://host:65536/', 'http://[1.2.3.4]/', 'http://[1::2::3]/', 'http://[::1]:99999/', 'http://xn--zz/', 'http://xn--a.com/', 'http://\u2488.com/', 'http://evil.com\uff20good.com/', 'http://..80/', 'http://1.2.3/', 'http://0x7f.1/', 'http://a.1/', 'http://a\u200d.com/', 'http://\u0301a.com/', 'http://a@', 'http://:80/', 'http://a@:80', 'http://evil.com:x/', 'http://ev il.com/', 'http://evil.com%40good.com/', 'http://a<b.com/', 'http://[::1/', 'http://[zz]/', 'http://a.com:80:80/', 'http://\\\\\\a.com/'])
    assert.equal(pdfLinkUri(target), undefined, JSON.stringify(target));
});

test('pdf links: the host a reader of the written address finds is the host that was checked (random addresses)', () => {
  // Pieces that change where a host starts and ends, put together at random from a seed: the same addresses every run.
  const pieces = ['a', 'b.com', 'evil.com', '@', ':', '80', '65536', '/', '\\', '?', '#', '[', ']', '::1', '1.2.3.4', '0x7f', '%40', '%', ' ', 'ö', 'İ', '.', '-', 'x\\@y', '\t', '..', 'xn--', 'zz', 'rnek-4qa', '\uff20', '\u2488', '\uff41', '\u0301', '\u00ad', '\u200d', '\u05d0', '\u2160', '\ud83d\ude00'];
  let seed = 20261014;
  const next = () => {
    seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
    return seed >>> 8;
  };
  let written = 0;
  for (let round = 0; round < 20000; round++) {
    let target = next() % 2 ? 'http://' : 'https:\\\\';
    const count = 1 + (next() % 8);
    for (let i = 0; i < count; i++) target += pieces[next() % pieces.length];
    const read = pdfLinkTarget(target);
    if (!read) continue;
    written++;
    const parse = (text: string): URL | undefined => {
      try {
        return new URL(text);
      } catch {
        return undefined;
      }
    };
    const url = parse(read.uri);
    // A browser that is given the address as the file has it.
    const raw = parse(target);
    // Every address that is written can be read: one no reader can read ("http://..80", "http://xn--zz") is not written.
    assert.ok(url, `${JSON.stringify(target)} -> ${read.uri}, which a reader of addresses cannot read`);
    if (!url) continue;
    assert.equal(url.hostname, new URL('http://' + read.host).hostname, `${JSON.stringify(target)} -> ${read.uri}`);
    if (raw) assert.equal(raw.hostname, url.hostname, `${JSON.stringify(target)} is read as ${raw.hostname}`);
  }
  assert.ok(written > 1000, `only ${written} addresses were written`);
});

test('pdf links: every other kind of address is refused', () => {
  for (const target of [
    'javascript:alert(1)',
    'JavaScript:alert(1)',
    ' javascript:alert(1)',
    'java\tscript:alert(1)',
    'vbscript:x',
    'file:///etc/passwd',
    'file://server/share',
    'data:text/html,<script>alert(1)</script>',
    'tel:+905551234567',
    'ftp://example.com',
    'smb://server/x',
    'relative/path.html',
    '/absolute/path',
    '//example.com/no-scheme',
    '#bookmark',
    'https:example.com',
    'https:/example.com',
    'https://',
    'https:///path',
    'mailto:',
    '',
    '   ',
    'launch:calc.exe',
  ])
    assert.equal(pdfLinkUri(target), undefined, JSON.stringify(target));
  assert.equal(pdfLinkUri(undefined), undefined);
  assert.equal(pdfLinkUri(42 as unknown as string), undefined);
  // Too long once it is written: refused, not cut.
  assert.equal(pdfLinkUri('https://example.com/' + 'a'.repeat(MAX_URI_BYTES)), undefined);
  assert.equal(pdfLinkUri('https://example.com/' + 'ğ'.repeat(MAX_URI_BYTES / 6)), undefined);
  assert.ok(pdfLinkUri('https://example.com/' + 'a'.repeat(MAX_URI_BYTES - 100)));
});

test('pdf links: the title of an entry of the outline is one line of text, cut to its limit', () => {
  assert.equal(outlineTitle('  Giriş\tve\n\namaç  '), 'Giriş ve amaç');
  assert.equal(outlineTitle('a\u0000b\u0007c\u00adD\ufeff'), 'abcD');
  assert.equal(outlineTitle(''), '');
  assert.equal(outlineTitle(' \t '), '');
  assert.equal(outlineTitle('x'.repeat(1000)).length, 256);
  // Not cut inside a pair of surrogates.
  const cut = outlineTitle('x'.repeat(255) + '😀');
  assert.equal(cut, 'x'.repeat(255));
});

test('pdf links: the outline is a tree by the levels of its entries, made without recursion', () => {
  const entry = (title: string, level: number): PdfOutlineEntry => ({ title, level, dest: { page: 0 } });
  const { nodes, top } = outlineTree([entry('A', 0), entry('A.1', 1), entry('A.1.a', 3), entry('A.2', 1), entry('B', 0), entry('B.1', 2), entry('C', 5), entry('D', -4), entry('E', NaN)]);
  const titles = (list: number[]) => list.map((index) => nodes[index].entry.title);
  assert.deepEqual(titles(top), ['A', 'B', 'D', 'E']);
  assert.deepEqual(titles(nodes[0].children), ['A.1', 'A.2']);
  assert.deepEqual(titles(nodes[1].children), ['A.1.a']);
  assert.deepEqual(titles(nodes[4].children), ['B.1']);
  assert.deepEqual(titles(nodes[5].children), ['C'], 'a level that skips some is one deeper');
  assert.equal(nodes[0].descendants, 3);
  assert.equal(nodes[4].descendants, 2);
  assert.deepEqual(
    nodes.map((node) => node.index),
    [0, 0, 0, 1, 1, 0, 0, 2, 3],
  );
  // Ten thousand levels one under the other: a chain of nine at most.
  const deep: PdfOutlineEntry[] = [];
  for (let i = 0; i < 10000; i++) deep.push(entry('x' + i, i));
  const chain = outlineTree(deep);
  assert.equal(chain.top.length, 1);
  assert.equal(chain.nodes[0].descendants, 9999);
});

// ---------------------------------------------------------------- the file

test('pdf links: a link out of the file is a link annotation with a URI action, over its rectangle', async () => {
  const { file, pdf } = await fileOf([page([{ rects: [rect(100, 50, 80, 20)], uri: 'https://example.com/(a)?\\b' }])]);
  const [annots] = annotations(file);
  assert.equal(annots.length, 1);
  const annot = annots[0];
  assert.equal((annot.Type as { name: string }).name, 'Annot');
  assert.equal((annot.Subtype as { name: string }).name, 'Link');
  // From the top left corner of the page to [left bottom right top] from the bottom.
  assert.deepEqual(annot.Rect, [100, 730, 180, 750]);
  assert.deepEqual(annot.Border, [0, 0, 0]);
  assert.equal(annot.F, 4);
  const action = annot.A as Dict;
  assert.equal((action.S as { name: string }).name, 'URI');
  assert.equal(decode(action.URI as Str), 'https://example.com/(a)?%5Cb');
  assert.deepEqual(pdf.links, { external: 1, internal: 0, refused: 0, unresolved: 0, dropped: 0, rotated: 0, outline: 0, outlineSkipped: 0, unread: 0 });
  // The parentheses of the address are escaped in the file, not left to close the string.
  assert.ok(Buffer.from(file).toString('latin1').indexOf('/URI (https://example.com/\\(a\\)?%5Cb)') > 0);
});

test('pdf links: a link into the file goes to its page and place, also to a page that comes later', async () => {
  const { file, pdf } = await fileOf([
    page([{ rects: [rect(10, 10)], dest: { page: 2, x: 72, y: 100 } }]),
    page([{ rects: [rect(10, 10)], dest: { page: 0 } }]),
    page(undefined, { height: 500 }),
  ]);
  const checked = checkPdf(file);
  const root = checked.objects.get((checked.trailer.Root as Ref).id)!.value as Dict;
  const kids = (checked.objects.get((root.Pages as Ref).id)!.value as Dict).Kids as Ref[];
  const all = annotations(file);
  const first = all[0][0].Dest as unknown[];
  assert.equal((first[0] as Ref).id, kids[2].id);
  assert.equal((first[1] as { name: string }).name, 'XYZ');
  // The place is measured from the bottom of its page, which is 500 points high.
  assert.deepEqual(first.slice(2), [72, 400, null]);
  const second = all[1][0].Dest as unknown[];
  assert.equal((second[0] as Ref).id, kids[0].id);
  assert.deepEqual(second.slice(2), [null, null, null]);
  assert.equal(pdf.links.internal, 2);
});

test('pdf links: what cannot be a link is left out and counted', async () => {
  const { file, pdf } = await fileOf([
    page([
      { rects: [rect(10, 10)], uri: 'javascript:alert(1)' },
      { rects: [rect(10, 40)], uri: 'file:///etc/passwd' },
      { rects: [rect(10, 70)], uri: 'data:text/html,x' },
      { rects: [rect(10, 100)], uri: 'tel:123' },
      { rects: [rect(10, 130)], dest: { page: 7 } },
      { rects: [rect(10, 160)], dest: { page: -1 } },
      { rects: [rect(10, 190)], dest: { page: 0.5 } },
      { rects: [rect(-100, -100, 10, 10)], uri: 'https://example.com' },
      { rects: [rect(10, 220, 0, 10)], uri: 'https://example.com' },
      { rects: [{ x: NaN, y: 0, width: 10, height: 10 }], uri: 'https://example.com' },
      { rects: [rect(590, 790, 100, 100)], uri: 'https://example.com/edge' },
      null as unknown as PdfLink,
    ]),
  ]);
  assert.deepEqual(pdf.links, { external: 1, internal: 0, refused: 4, unresolved: 3, dropped: 3, rotated: 0, outline: 0, outlineSkipped: 0, unread: 0 });
  const [annots] = annotations(file);
  // A link to a page the file does not have is not in the file at all (no annotation that goes nowhere); the one at the edge is cut to the page.
  assert.equal(annots.length, 1);
  assert.equal(typeof (annots[0].A as Dict).URI, 'object');
  assert.deepEqual(annots[0].Rect, [590, 0, 600, 10]);
  for (const one of annotations(file).flat()) assert.ok(one.A || one.Dest, 'every annotation goes somewhere');
  const text = Buffer.from(file).toString('latin1');
  for (const bad of ['javascript', 'file:', 'data:', 'tel:']) assert.equal(text.indexOf(bad), -1, bad);
});

test('pdf links: a link of several lines has one annotation per line, and is counted once', async () => {
  const rects = [rect(300, 100, 200, 16), rect(50, 120, 100, 16)];
  const { file, pdf } = await fileOf([page([{ rects, uri: 'https://example.com/long' }, { rects, dest: { page: 0, y: 0 } }])]);
  assert.equal(annotations(file)[0].length, 4);
  assert.equal(pdf.links.external, 1);
  assert.equal(pdf.links.internal, 1);
  const many: { x: number; y: number; width: number; height: number }[] = [];
  for (let i = 0; i < 500; i++) many.push(rect(10, i, 10, 1));
  const big = await fileOf([page([{ rects: many, uri: 'https://example.com' }])]);
  assert.equal(annotations(big.file)[0].length, MAX_LINK_RECTS);
});

test('pdf links: links beyond the limits of a page and of the file are left out and counted', async () => {
  const links: PdfLink[] = [];
  for (let i = 0; i < MAX_PAGE_LINKS + 25; i++) links.push({ rects: [rect(10, 10, 5, 5)], uri: 'https://example.com/' + i });
  const { file, pdf } = await fileOf([page(links)]);
  assert.equal(annotations(file)[0].length, MAX_PAGE_LINKS);
  assert.equal(pdf.links.external, MAX_PAGE_LINKS);
  assert.equal(pdf.links.dropped, 25);
  // The file: as many pages as it takes to pass its limit.
  const pages: PdfPage[] = [];
  for (let n = 0; n * MAX_PAGE_LINKS < MAX_FILE_LINKS + 10; n++) pages.push(page(links.slice(0, MAX_PAGE_LINKS)));
  const whole = new PdfDocument({ date: DATE });
  for (const one of pages) {
    await whole.addPage(one);
    whole.takeParts();
  }
  await whole.finish();
  assert.equal(whole.links.external, MAX_FILE_LINKS);
  assert.equal(whole.links.dropped, pages.length * MAX_PAGE_LINKS - MAX_FILE_LINKS);
});

test('pdf links: the outline is written with its tree, its counts and its places', async () => {
  const outline: PdfOutlineEntry[] = [
    { title: 'Giriş', level: 0, dest: { page: 0, x: 0, y: 72 } },
    { title: 'Amaç', level: 1, dest: { page: 0, y: 300 } },
    { title: 'Kapsam', level: 1, dest: { page: 1, y: 10 } },
    { title: 'Ayrıntı', level: 2, dest: { page: 1, y: 400 } },
    { title: 'Sonuç (özet) \\', level: 0, dest: { page: 2 } },
    { title: '   ', level: 0, dest: { page: 1 } },
    { title: 'Olmayan sayfa', level: 1, dest: { page: 9 } },
  ];
  const { file, pdf } = await fileOf([page(), page(), page()], outline);
  assert.equal(pdf.links.outline, 5);
  assert.equal(pdf.links.outlineSkipped, 2, 'the entry without a title and the one whose page is not in the file');
  const checked = checkPdf(file);
  const root = checked.objects.get((checked.trailer.Root as Ref).id)!.value as Dict;
  assert.equal((root.PageMode as { name: string }).name, 'UseOutlines');
  const top = checked.objects.get((root.Outlines as Ref).id)!.value as Dict;
  assert.equal((top.Type as { name: string }).name, 'Outlines');
  assert.equal(top.Count, 5);
  const entry = (ref: Ref) => checked.objects.get(ref.id)!.value as Dict;
  const first = entry(top.First as Ref);
  const last = entry(top.Last as Ref);
  assert.equal(first.Count, 3);
  assert.equal((first.Next as Ref).id, (top.Last as Ref).id);
  assert.equal(last.Prev && (last.Prev as Ref).id, (top.First as Ref).id);
  assert.equal(last.Next, undefined);
  assert.equal(last.Count, undefined);
  assert.equal((first.Parent as Ref).id, (root.Outlines as Ref).id);
  const amac = entry(first.First as Ref);
  assert.equal((amac.Parent as Ref).id, (top.First as Ref).id);
  assert.equal(amac.Prev, undefined);
  const kapsam = entry(first.Last as Ref);
  assert.equal(kapsam.Count, 1);
  assert.deepEqual((kapsam.Dest as unknown[]).slice(2), [null, 790, null]);
});

test('pdf links: an outline of more entries than its limit keeps the first ones', async () => {
  const outline: PdfOutlineEntry[] = [];
  for (let i = 0; i < MAX_OUTLINE_ENTRIES + 7; i++) outline.push({ title: 'e' + i, level: i % 3, dest: { page: 0 } });
  const { pdf } = await fileOf([page()], outline);
  assert.equal(pdf.links.outline, MAX_OUTLINE_ENTRIES);
  assert.equal(pdf.links.dropped, 7);
  // No outline, no entry in the catalog.
  const none = await fileOf([page()], []);
  const checked = checkPdf(none.file);
  const root = checked.objects.get((checked.trailer.Root as Ref).id)!.value as Dict;
  assert.equal(root.Outlines, undefined);
  assert.equal(root.PageMode, undefined);
});

pythonTest('pdf links: the reader in Python finds the links and the outline the file was given', async () => {
  const { file } = await fileOf(
    [
      page([
        { rects: [rect(100, 50, 80, 20)], uri: 'https://örnek.com.tr/a b' },
        { rects: [rect(100, 100, 80, 20), rect(20, 120, 40, 20)], dest: { page: 1, x: 10, y: 20 } },
      ]),
      page([{ rects: [rect(30, 30)], uri: 'mailto:a@example.com' }]),
    ],
    [
      { title: 'Bölüm 1', level: 0, dest: { page: 0, y: 0 } },
      { title: 'Alt', level: 1, dest: { page: 1, y: 100 } },
      { title: 'Bölüm 2', level: 0, dest: { page: 1, y: 300 } },
    ],
  );
  const read = refLinks(file);
  assert.deepEqual(
    read.pages.map((one) => one.annots.map((annot) => ({ rect: annot.rect, uri: annot.uri, dest: annot.dest }))),
    [
      [
        { rect: [100, 730, 180, 750], uri: 'https://%C3%B6rnek.com.tr/a%20b', dest: undefined },
        { rect: [100, 680, 180, 700], uri: undefined, dest: [1, 10, 780] },
        { rect: [20, 660, 60, 680], uri: undefined, dest: [1, 10, 780] },
      ],
      [{ rect: [30, 750, 80, 770], uri: 'mailto:a@example.com', dest: undefined }],
    ],
  );
  assert.equal(read.outlineCount, 3);
  assert.equal(read.pageMode, 'UseOutlines');
  assert.deepEqual(read.outline, [
    { title: 'Bölüm 1', dest: [0, null, 800], count: 1, kids: [{ title: 'Alt', dest: [1, null, 700], count: null, kids: [] }] },
    { title: 'Bölüm 2', dest: [1, null, 500], count: null, kids: [] },
  ]);
});

popplerTest('pdf links: Poppler reads the outline, and the link lies over the words of its text', async () => {
  const runs = [{ text: 'Siteye bakın: örnek sayfa', x: 100, y: 100, width: 200, size: 16 }];
  // "örnek sayfa" begins after 14 characters of 8 points.
  const links: PdfLink[] = [{ rects: [{ x: 100 + 14 * 8, y: 100 - 13, width: 11 * 8, height: 17 }], uri: 'https://example.com/örnek' }];
  const { file } = await fileOf([page(links, { text: { runs, width: 600 } }), page()], [
    { title: 'Başlık', level: 0, dest: { page: 0 } },
    { title: 'Alt başlık', level: 1, dest: { page: 1 } },
  ]);
  const xml = popplerXml(file);
  assert.match(xml, /<outline>\s*<item page="1">Başlık<\/item>\s*<outline>\s*<item page="2">Alt başlık<\/item>\s*<\/outline>\s*<\/outline>/);
  assert.match(xml, /<a href="https:\/\/example.com\/%C3%B6rnek">/);
  const boxes = popplerBoxes(file);
  const words = boxes.filter((box) => box.word === 'örnek' || box.word === 'sayfa');
  assert.equal(words.length, 2);
  const [annot] = refLinksOrOwn(file)[0];
  // [left bottom right top] from the bottom: the words, read from the top, lie inside it.
  for (const word of words) {
    assert.ok(word.x1 >= annot[0] - 0.5 && word.x2 <= annot[2] + 0.5, `${word.word} lies across the link`);
    assert.ok(800 - word.y2 >= annot[1] - 1 && 800 - word.y1 <= annot[3] + 1, `${word.word} lies above or below the link`);
  }
  const other = boxes.find((box) => box.word === 'Siteye')!;
  assert.ok(other.x2 <= annot[0], 'the words before the link are not under it');
});

/** The rectangles of the link annotations, by page: by the reader of the tests. */
function refLinksOrOwn(file: Uint8Array): number[][][] {
  return annotations(file).map((annots) => annots.map((annot) => annot.Rect as number[]));
}

test('pdf links: a file without links and outline is as it was', async () => {
  const plain = await buildPdf([page()], { date: DATE });
  const text = Buffer.from(plain).toString('latin1');
  assert.equal(text.indexOf('/Annots'), -1);
  assert.equal(text.indexOf('/Outlines'), -1);
});

// ---------------------------------------------------------------- export options

test('pdf links: the options of links and outline are checked, and the report is told once', () => {
  assert.deepEqual([pdfSettings({}).links, pdfSettings({}).outline], [true, true]);
  assert.deepEqual([pdfSettings({ links: false, outline: false }).links, pdfSettings({ links: false, outline: false }).outline], [false, false]);
  assert.throws(() => pdfSettings({ links: 'yes' as unknown as boolean }), TypeError);
  assert.throws(() => pdfSettings({ outline: 1 as unknown as boolean }), TypeError);
  assert.throws(() => pdfSettings({ onLinks: 'x' as unknown as () => void }), TypeError);
  const told: unknown[] = [];
  const report = reportLinks({ onLinks: (one) => told.push(one) }, { external: 2, internal: 1, refused: 1, unresolved: 0, dropped: 0, rotated: 0, outline: 4, outlineSkipped: 1, unread: 0 }, { refused: 2, unresolved: 3, rotated: 1, outlineSkipped: 2, unread: 5 });
  assert.deepEqual(report, { external: 2, internal: 1, refused: 3, unresolved: 3, dropped: 0, rotated: 1, outline: 4, outlineSkipped: 3, unread: 5 });
  assert.deepEqual(told, [report]);
  // A listener that throws does not stop the file.
  assert.doesNotThrow(() =>
    reportLinks(
      {
        onLinks: () => {
          throw new Error('x');
        },
      },
      report,
      {},
    ),
  );
});

// ---------------------------------------------------------------- documents

const HEADING_STYLES =
  STYLES +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  // A heading style that does not state its level: Word's own heading styles are headings of their level.
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>' +
  // A style of its own with a level, and one based on it.
  '<w:style w:type="paragraph" w:styleId="Ozel"><w:name w:val="Özel başlık"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="OzelAlt"><w:name w:val="Özel alt"/><w:basedOn w:val="Ozel"/></w:style>' +
  // Body text: level 9.
  '<w:style w:type="paragraph" w:styleId="Govde"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
  // A chain that goes round.
  '<w:style w:type="paragraph" w:styleId="Loop1"><w:name w:val="Loop1"/><w:basedOn w:val="Loop2"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Loop2"><w:name w:val="Loop2"/><w:basedOn w:val="Loop1"/></w:style>';

const styled = (style: string, text: string) => p(r(text), `<w:pStyle w:val="${style}"/>`);
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const field = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('pdf links of a document: the headings make the outline, by the levels Word gives them', async () => {
  const body =
    styled('Heading1', 'Giriş') +
    styled('Heading2', 'Amaç') +
    p(r('Doğrudan düzey'), '<w:outlineLvl w:val="1"/>') +
    styled('OzelAlt', 'Özel alt başlık') +
    styled('Govde', 'Gövde metni') +
    styled('Loop1', 'Döngü') +
    p('Sıradan paragraf') +
    styled('Heading1', '   ') +
    pageBreak +
    `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${styled('Heading1', 'Tablodaki başlık')}</w:tc></w:tr></w:tbl>` +
    styled('Heading1', 'Sonuç');
  const { doc, pages } = await lay({ body, styles: HEADING_STYLES });
  assert.equal(pages.length, 2);
  const found = docPdfLinks(pages, doc, { scale: 0.75, metrics: FAKE_METRICS });
  assert.deepEqual(
    found.outline.map((entry) => [entry.title, entry.level, entry.dest.page]),
    [
      ['Giriş', 0, 0],
      ['Amaç', 1, 0],
      ['Doğrudan düzey', 1, 0],
      ['Özel alt başlık', 2, 0],
      ['Tablodaki başlık', 0, 1],
      ['Sonuç', 0, 1],
    ],
  );
  // At the top of its line: the first line of the page is 40 px from the top, 30 points.
  assert.equal(found.outline[0].dest.y, 30);
  assert.equal(found.outline[1].dest.y, 30 + 16 * 0.75);
  assert.equal(found.outlineSkipped, 1, 'the heading of spaces only');
  assert.equal(docPdfLinks(pages, doc, { scale: 0.75, outline: false }).outline.length, 0);
});

test('pdf links of a document: links out of it, into it, and those that cannot be', async () => {
  const body =
    p(`<w:hyperlink r:id="rIdWeb">${r('site')}</w:hyperlink>${r(' ve ')}<w:hyperlink r:id="rIdBad">${r('kötü')}</w:hyperlink>`) +
    p(`<w:hyperlink w:anchor="hedef">${r('hedefe git')}</w:hyperlink>${r(' ')}<w:hyperlink w:anchor="yok">${r('kayıp')}</w:hyperlink>`) +
    p(field(' HYPERLINK \\l "HEDEF" ', 'alan bağlantısı') + r(' ') + field(' HYPERLINK "file:///c:/x.doc" ', 'dosya')) +
    p(`<w:hyperlink r:id="rIdTel">${r('telefon')}</w:hyperlink>`) +
    pageBreak +
    p(r('Önce ') + '<w:bookmarkStart w:id="1" w:name="hedef"/>' + r('hedef burada') + '<w:bookmarkEnd w:id="1"/>');
  const { doc, pages } = await lay({
    body,
    styles: HEADING_STYLES,
    rels: [
      ['rIdWeb', 'hyperlink', 'https://example.com/a', true],
      ['rIdBad', 'hyperlink', 'javascript:alert(1)', true],
      ['rIdTel', 'hyperlink', 'tel:123', true],
    ],
  });
  const found = docPdfLinks(pages, doc, { scale: 0.75, metrics: FAKE_METRICS });
  // Refused before the pages: javascript: and file: (tel: is refused when it is written).
  assert.equal(found.refused, 2);
  assert.equal(found.unresolved, 1, 'the link to a bookmark that is not there');
  const first = found.pages[0];
  // The bookmark is on the second line of the second page (the first is what is left of the paragraph of the page break):
  // the place is the start of the text item that holds it, at the top of its line.
  assert.deepEqual(
    first.map((link) => link.uri || link.dest),
    ['https://example.com/a', { page: 1, x: 30, y: (40 + 16) * 0.75 }, { page: 1, x: 30, y: (40 + 16) * 0.75 }, 'tel:123'],
  );
  // Over the text of the link: "site" at the start of the first line, 4 characters of 8 px, a line of 16 px.
  assert.deepEqual(first[0].rects, [{ x: 30, y: 30, width: 4 * 8 * 0.75, height: 16 * 0.75 }]);
  const { pdf } = await fileOf(pages.map((one, index) => page(found.pages[index], { width: one.width * 0.75, height: one.height * 0.75 })));
  assert.deepEqual(pdf.links, { external: 1, internal: 2, refused: 1, unresolved: 0, dropped: 0, rotated: 0, outline: 0, outlineSkipped: 0, unread: 0 });
});

test('pdf links of a document: a link that runs over two lines has a rectangle on each', async () => {
  const text = 'bir iki üç dört beş altı yedi sekiz dokuz on';
  const { doc, pages } = await lay({ body: p(r('Başta yazı var: ') + `<w:hyperlink r:id="rIdWeb">${r(text)}</w:hyperlink>`), rels: [['rIdWeb', 'hyperlink', 'https://example.com/uzun', true]] });
  const found = docPdfLinks(pages, doc, { scale: 0.75, metrics: (font) => ({ ascent: font.size * 0.8, descent: font.size * 0.2 }) });
  assert.equal(found.pages[0].length, 1);
  const rects = found.pages[0][0].rects;
  assert.equal(rects.length, 2);
  assert.ok(rects[1].y > rects[0].y, 'the second piece is on the next line');
  assert.equal(rects[0].height, 16 * 0.75);
  // Together as wide as the text of the link, without the space the line broke at.
  const width = rects.reduce((sum, one) => sum + one.width, 0) / 0.75;
  assert.ok(Math.abs(width - text.length * 8) <= 8, String(width));
});

test('pdf links of a document: the links of headers and footers are on every page', async () => {
  const { doc, pages } = await lay({
    body: p(r('bir') + '<w:bookmarkStart w:id="3" w:name="Basa"/><w:bookmarkEnd w:id="3"/>') + pageBreak + p('iki') + sectPr('<w:headerReference w:type="default" r:id="rIdH"/>'),
    headers: { rIdH: p(`<w:hyperlink w:anchor="Basa">${r('başa dön')}</w:hyperlink>`) },
  });
  const found = docPdfLinks(pages, doc, { scale: 0.75 });
  assert.equal(found.pages.length, 2);
  for (let i = 0; i < 2; i++) {
    assert.equal(found.pages[i].length, 1, `page ${i + 1}`);
    assert.equal(found.pages[i][0].dest!.page, 0);
    // In the header: above the body, which starts 40 px from the top.
    assert.ok(found.pages[i][0].rects[0].y < 30);
  }
  assert.equal(found.unresolved, 0);
});

test('pdf links of a document: an RTF file has its links, its bookmarks and its headings too', () => {
  const rtf =
    '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Test;}}' +
    '{\\stylesheet{\\s0 Normal;}{\\s1\\outlinelevel0 heading 1;}}' +
    '\\paperw6000\\paperh6000\\margl600\\margr600\\margt600\\margb600 ' +
    '{\\pard\\s1\\outlinelevel0 Giri\\u351?\\par}' +
    '{\\pard {\\field{\\*\\fldinst HYPERLINK "https://example.com/rtf"}{\\fldrslt site}} ve {\\field{\\*\\fldinst HYPERLINK \\\\l "son"}{\\fldrslt sona}} ve {\\field{\\*\\fldinst HYPERLINK "javascript:x()"}{\\fldrslt k\\u246?t\\u252?}}\\par}' +
    '{\\pard\\page\\par}' +
    '{\\pard {\\*\\bkmkstart son}Son{\\*\\bkmkend son}\\par}}';
  const doc = readRtf(new TextEncoder().encode(rtf));
  const layout = new DocLayout(doc, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const pages = layout.pages;
  const found = docPdfLinks(pages, doc, { scale: 0.75, metrics: FAKE_METRICS });
  assert.deepEqual(
    found.outline.map((entry) => [entry.title, entry.level, entry.dest.page]),
    [['Giriş', 0, 0]],
  );
  const links = found.pages.flat();
  assert.deepEqual(
    links.map((link) => link.uri || (link.dest && link.dest.page)),
    ['https://example.com/rtf', pages.length - 1],
  );
  assert.equal(found.refused, 1);
});

test('pdf links of a document: hidden text is not in the title of a heading, and a heading of hidden text only is not in the outline', async () => {
  const styles = HEADING_STYLES + '<w:style w:type="character" w:styleId="Gizli"><w:name w:val="Gizli"/><w:rPr><w:vanish/></w:rPr></w:style>';
  const body =
    p(r('Gizli ', '<w:vanish/>') + r('Görünen'), '<w:pStyle w:val="Heading1"/>') +
    p(r('Stilden gizli ', '<w:rStyle w:val="Gizli"/>') + r('başlık'), '<w:pStyle w:val="Heading1"/>') +
    p(r('Tümü gizli', '<w:vanish/>'), '<w:pStyle w:val="Heading1"/>') +
    p('son');
  const { doc, pages } = await lay({ body, styles });
  const found = docPdfLinks(pages, doc, { scale: 0.75 });
  assert.deepEqual(
    found.outline.map((entry) => entry.title),
    ['Görünen', 'başlık'],
  );
  assert.equal(found.outlineSkipped, 1);
});

test('pdf links of a document: beyond the limit of what is walked, the blocks left are counted, each once', () => {
  // 2,000,005 paragraphs in the body (the limit is 2,000,000) and two in a header: seven not read.
  const empty = { type: 'paragraph', props: {}, children: [] };
  const model = {
    body: { blocks: new Array(2000005).fill(empty) },
    styles: { styles: new Map(), defaultIds: {} },
    headers: new Map([['h1', { blocks: [empty, empty] }]]),
    footers: new Map(),
  };
  const found = docPdfLinks([{ width: 400, height: 400, section: 0, number: 1, items: [] }], model as never, { scale: 0.75 });
  assert.equal(found.unread, 7);
});

test('pdf links: links into the file that go nowhere take no room from those that go somewhere', async () => {
  // A page of 1,000 links to a page the file does not have and then 2,000 that are fine: all 2,000 are written.
  const links: PdfLink[] = [];
  for (let i = 0; i < 1000; i++) links.push({ rects: [rect(10, 10, 5, 5)], dest: { page: 999 } });
  for (let i = 0; i < MAX_PAGE_LINKS; i++) links.push({ rects: [rect(10, 10, 5, 5)], dest: { page: 1 } });
  // And one more than a page takes: left out when it is written, counted.
  links.push({ rects: [rect(10, 10, 5, 5)], dest: { page: 0 } });
  const { file, pdf } = await fileOf([page(links), page()]);
  assert.equal(pdf.links.internal, MAX_PAGE_LINKS);
  assert.equal(pdf.links.unresolved, 1000);
  assert.equal(pdf.links.dropped, 1);
  assert.equal(annotations(file)[0].length, MAX_PAGE_LINKS);
});

test('pdf links of a document: a link in turned text gets no link, and is counted once', () => {
  const font = { family: 'Test', size: 16, bold: false, italic: false };
  const text = (x: number, y: number, link: string | undefined, rotation?: 90 | 270) => ({ type: 'text', x, y, width: 40, text: 'abc', font, color: '#000000', link, rotation, source: { story: 'body', path: [0], offset: 0 } });
  const page = { width: 400, height: 400, section: 0, number: 1, items: [text(10, 50, 'https://example.com/a', 90), text(10, 90, 'https://example.com/a', 90), text(100, 50, 'https://example.com/b', 270), text(200, 50, 'https://example.com/c')] };
  const model = { body: { blocks: [] }, styles: { styles: new Map(), defaultIds: {} }, headers: new Map(), footers: new Map() };
  const found = docPdfLinks([page as never], model as never, { scale: 0.75 });
  assert.deepEqual(
    found.pages[0].map((link) => link.uri),
    ['https://example.com/c'],
  );
  assert.equal(found.rotated, 2);
});

test('pdf links of a document: nothing is looked for when nothing is asked for', async () => {
  const { doc, pages } = await lay({ body: p(`<w:hyperlink w:anchor="x">${r('a')}</w:hyperlink>`) + styled('Heading1', 'B'), styles: HEADING_STYLES });
  const none = docPdfLinks(pages, doc, { scale: 0.75, links: false, outline: false });
  assert.deepEqual(none, { pages: [[]], outline: [], refused: 0, unresolved: 0, dropped: 0, rotated: 0, outlineSkipped: 0, unread: 0 });
  const onlyLinks = docPdfLinks(pages, doc, { scale: 0.75, outline: false });
  assert.equal(onlyLinks.unresolved, 1);
  assert.equal(onlyLinks.outline.length, 0);
});

test('pdf links of a document: a broken model and broken pages give no links rather than an error', () => {
  const broken = { body: { blocks: [null, { type: 'paragraph', props: {}, children: [null] }, { type: 'table', rows: [null, { cells: [null] }] }] }, styles: { styles: new Map(), defaultIds: {} }, headers: new Map(), footers: new Map() };
  const pages = [{ width: 100, height: 100, section: 0, number: 1, items: [null, { type: 'text', x: 0, y: 0, width: 5, text: 'x', link: '#nowhere', font: { family: 'x', size: NaN, bold: false, italic: false }, color: '#000' }] }];
  const found = docPdfLinks(pages as never, broken as never, { scale: 0.75 });
  assert.equal(found.pages[0].length, 0);
  assert.equal(found.unresolved, 1);
});

test('pdf links of a document: tables in tables deeper than the limit are not walked', async () => {
  const { doc } = await lay({ body: styled('Heading1', 'Derin') + styled('Heading1', 'Sığ'), styles: HEADING_STYLES });
  const heading = doc.body.blocks[0];
  const shallow = doc.body.blocks[1];
  // A model no reader makes (the docx reader refuses more than 64 levels): two hundred tables in each other.
  let blocks: Block[] = [heading];
  for (let i = 0; i < 200; i++) blocks = [{ type: 'table', props: {}, grid: [3000], rows: [{ props: {}, cells: [{ props: {}, blocks }] }] } as unknown as Block];
  const model = { ...doc, body: { ...doc.body, blocks: blocks.concat([shallow]) } };
  const page = { width: 400, height: 400, section: 0, number: 1, items: [{ type: 'text', x: 40, y: 52.8, width: 24, text: 'Sığ', font: { family: 'Test', size: 16, bold: false, italic: false }, color: '#000000', source: { story: 'body', path: [1], offset: 0 } }] };
  const found = docPdfLinks([page as never], model, { scale: 0.75 });
  assert.deepEqual(
    found.outline.map((entry) => entry.title),
    ['Sığ'],
  );
  assert.equal(found.unread, 1, 'the table beyond the depth that is walked');
});

// ---------------------------------------------------------------- workbooks

function sheetOf(name: string, rows: number, cols: number): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/${name}.xml` });
  for (let row = 0; row < rows; row++) {
    const cells: Cell[] = [];
    for (let col = 0; col < cols; col++) cells[col] = { kind: 'string', value: `r${row}c${col}`, style: 0 };
    sheet.rows.set(row, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
  }
  return sheet;
}

const bookOf = (sheets: Sheet[], names: Workbook['definedNames'] = []): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: names, date1904: false, activeSheet: 0 });

test('pdf links of a workbook: a location names a sheet and cells, or a name that does', () => {
  const book = bookOf([sheetOf('Veri', 1, 1), sheetOf("Ali'nin sayfası", 1, 1), sheetOf('A!B', 1, 1)], [
    { name: 'Hedef', value: "'Ali''nin sayfası'!$B$3", hidden: false },
    { name: 'Yerel', value: 'Veri!$C$1', hidden: false, localSheet: 2 },
    { name: 'Formül', value: 'SUM(Veri!A1:A3)', hidden: false },
  ]);
  assert.deepEqual(parseLocation('Veri!B5', book, 2), { sheet: 0, range: { top: 4, left: 1, bottom: 4, right: 1 } });
  assert.deepEqual(parseLocation("#'Ali''nin sayfası'!$A$1:$C$3", book, 0), { sheet: 1, range: { top: 0, left: 0, bottom: 2, right: 2 } });
  assert.deepEqual(parseLocation('VERİ!A1'.replace('İ', 'I'), book, 2), { sheet: 0, range: { top: 0, left: 0, bottom: 0, right: 0 } });
  assert.deepEqual(parseLocation("'A!B'!C2", book, 0), { sheet: 2, range: { top: 1, left: 2, bottom: 1, right: 2 } });
  assert.deepEqual(parseLocation('D4', book, 1), { sheet: 1, range: { top: 3, left: 3, bottom: 3, right: 3 } });
  assert.deepEqual(parseLocation('Hedef', book, 0), { sheet: 1, range: { top: 2, left: 1, bottom: 2, right: 1 } });
  assert.deepEqual(parseLocation('yerel', book, 2), { sheet: 0, range: { top: 0, left: 2, bottom: 0, right: 2 } });
  for (const bad of ['Yok!A1', 'Formül', 'Yerel', "'Açık!A1", 'Veri!', 'Veri!ZZZZ1', '', 'x'.repeat(2000)]) assert.equal(parseLocation(bad, book, 0), undefined, bad);
});

test('pdf links of a workbook: links over their cells on the printed pages, into the workbook and out of it, and the sheets in the outline', () => {
  const data = sheetOf('Veri', 80, 3);
  data.hyperlinks.push(
    { range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/a' },
    { range: { top: 1, left: 0, bottom: 1, right: 0 }, location: 'Özet!B2' },
    { range: { top: 70, left: 1, bottom: 70, right: 1 }, location: 'A2' },
    { range: { top: 2, left: 0, bottom: 2, right: 0 }, location: 'Yok!A1' },
    { range: { top: 3, left: 0, bottom: 3, right: 0 }, rId: 'rId9' },
    { range: { top: 4, left: 0, bottom: 4, right: 0 }, target: 'https://example.com/b', location: 'Sheet1!A1' },
    { range: { top: 5, left: 1, bottom: 5, right: 1 }, target: 'javascript:alert(1)' },
  );
  data.merges.push({ top: 5, left: 1, bottom: 6, right: 2 });
  const summary = sheetOf('Özet', 3, 2);
  const book = bookOf([data, summary]);
  const plan = planSheetPages(book, [0, 1]);
  assert.ok(plan.pages.length >= 3);
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  assert.deepEqual(
    found.outline.map((entry) => [entry.title, entry.dest.page]),
    [
      ['Veri', 0],
      ['Özet', plan.pages.findIndex((one) => one.sheet === 1)],
    ],
  );
  assert.equal(found.unresolved, 2, 'a sheet that is not there, a relationship without a target');
  const first = found.pages[0];
  assert.equal(first[0].uri, 'https://example.com/a');
  const summaryPage = plan.pages.findIndex((one) => one.sheet === 1);
  assert.equal(first[1].dest!.page, summaryPage);
  assert.equal(first[2].uri, 'https://example.com/b#Sheet1!A1');
  assert.equal(first[3].uri, 'javascript:alert(1)', 'refused when it is written');
  // Over A1: at the corner of the printed block, as wide and high as the cell.
  const metrics = createSheetMetrics(data);
  const entry = plan.pages[0];
  const s = entry.layout.scale;
  const a1 = first[0].rects[0];
  assert.ok(Math.abs(a1.x - entry.page.x * 0.75) < 1e-9);
  assert.ok(Math.abs(a1.y - entry.page.y * 0.75) < 1e-9);
  assert.ok(Math.abs(a1.width - metrics.cols.size(0) * s * 0.75) < 1e-9);
  assert.ok(Math.abs(a1.height - metrics.rows.size(0) * s * 0.75) < 1e-9);
  // The merged cell the link starts: the whole of it.
  const merged = first[3].rects[0];
  assert.ok(Math.abs(merged.width - (metrics.cols.size(1) + metrics.cols.size(2)) * s * 0.75) < 1e-9);
  assert.ok(Math.abs(merged.height - (metrics.rows.size(5) + metrics.rows.size(6)) * s * 0.75) < 1e-9);
  // The link of row 71 is on the page that prints it, and goes back to the first page, where A2 is.
  const holder = plan.pages.findIndex((one) => one.sheet === 0 && one.page.rows.first <= 70 && one.page.rows.last >= 70);
  assert.ok(holder > 0);
  const back = found.pages[holder].find((link) => link.dest)!;
  assert.equal(back.dest!.page, 0);
  assert.ok(Math.abs(back.dest!.y! - (entry.page.y + metrics.rows.size(0) * s) * 0.75) < 1e-9);
  const none = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet), links: false, outline: false });
  assert.deepEqual(none.outline, []);
  assert.deepEqual(none.pages.flat(), []);
});

test('pdf links of a workbook: a link in the print titles is on every page that repeats them', () => {
  const data = sheetOf('Veri', 120, 2);
  data.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/t' });
  const book = bookOf([data], [{ name: '_xlnm.Print_Titles', value: 'Veri!$1:$1', hidden: false, localSheet: 0 }]);
  const plan = planSheetPages(book, [0]);
  assert.ok(plan.pages.length >= 2);
  assert.ok(plan.pages[1].page.titleRows, 'the second page repeats the title row');
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  for (let i = 0; i < plan.pages.length; i++) assert.equal(found.pages[i].length, 1, `page ${i + 1}`);
  assert.deepEqual(found.outline, [], 'one sheet has no outline');
});

test('pdf links of a workbook: a link to a cell that is not printed is left out and counted', () => {
  const data = sheetOf('Veri', 3, 2);
  data.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, location: 'Veri!Z900' });
  const book = bookOf([data, sheetOf('Gizli', 2, 2)]);
  data.hyperlinks.push({ range: { top: 1, left: 0, bottom: 1, right: 0 }, location: 'Gizli!A1' });
  const plan = planSheetPages(book, [0]);
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  assert.equal(found.unresolved, 2);
  assert.deepEqual(found.pages[0], []);
});

test('pdf links of a workbook: links beyond the limit of a sheet are counted once', () => {
  const data = sheetOf('Veri', 2, 2);
  for (let i = 0; i < 20010; i++) data.hyperlinks.push({ range: { top: 900000, left: 0, bottom: 900000, right: 0 }, target: 'https://example.com' });
  const book = bookOf([data]);
  const plan = planSheetPages(book, [0]);
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  assert.equal(found.dropped, 10);
});

test('pdf links of a workbook: the links of a page are found by its rows, and one left out is counted once', () => {
  // A sheet of many pages with all its links in A1, which the print titles repeat on every page: each page holds what
  // fits, and every link left out is counted once, not once a page.
  const data = sheetOf('Veri', 3000, 1);
  for (let i = 0; i < 20000; i++) data.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/' + i });
  // A link of a tall range is on every page that prints a row of it; a link to a sheet that is not there is counted once.
  data.hyperlinks.push({ range: { top: 0, left: 0, bottom: 2999, right: 0 }, target: 'https://example.com/tall' });
  const book = bookOf([data], [{ name: '_xlnm.Print_Titles', value: 'Veri!$1:$1', hidden: false, localSheet: 0 }]);
  const plan = planSheetPages(book, [0]);
  assert.ok(plan.pages.length > 20);
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  // A page holds 2,000 links and the file 50,000: the first 25 pages are full, the others get none - no more links are
  // made than the writer takes.
  const full = MAX_FILE_LINKS / MAX_PAGE_LINKS;
  for (let i = 0; i < plan.pages.length; i++) assert.equal(found.pages[i].length, i < full ? MAX_PAGE_LINKS : 0, `page ${i + 1}`);
  // The first 2,000 links of the sheet are on those pages; the other 18,000 and the tall one (which comes after them) on none: each counted once.
  assert.equal(found.dropped, 18001);
  assert.equal(found.pages[0][0].uri, 'https://example.com/0');
  // Without the crowd the tall link is on every page.
  data.hyperlinks.splice(0, 20000);
  data.hyperlinks.push({ range: { top: 5, left: 0, bottom: 5, right: 0 }, location: 'Yok!A1' });
  const tall = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  for (let i = 0; i < plan.pages.length; i++) assert.equal(tall.pages[i].filter((link) => link.uri === 'https://example.com/tall').length, 1, `page ${i + 1}`);
  assert.equal(tall.unresolved, 1);
  assert.equal(tall.dropped, 0);
});

test('pdf links of a workbook: tall links over every page stop at the limit of the file, and links beside the printed cells are not looked at', () => {
  // As the review measured it: 20,000 links of A1:A3000 over every page (before: two million links made and thrown away).
  const data = sheetOf('Veri', 3000, 1);
  // First 10,000 in columns no page prints: on no page, so neither written nor counted, and not looked at page by page.
  for (let i = 0; i < 10000; i++) data.hyperlinks.push({ range: { top: 0, left: 50 + (i % 100), bottom: 2999, right: 50 + (i % 100) }, target: 'https://example.com/' + i });
  for (let i = 0; i < 10000; i++) data.hyperlinks.push({ range: { top: 0, left: 0, bottom: 2999, right: 0 }, location: 'Veri!A' + (1 + (i % 3000)) });
  const book = bookOf([data]);
  const plan = planSheetPages(book, [0]);
  const found = sheetPdfLinks(book, plan.pages, { scale: 0.75, metrics: (sheet) => createSheetMetrics(sheet) });
  const written = found.pages.reduce((sum, list) => sum + list.length, 0);
  assert.equal(written, Math.min(MAX_FILE_LINKS, plan.pages.length * MAX_PAGE_LINKS));
  for (const list of found.pages) assert.ok(list.length <= MAX_PAGE_LINKS);
  // The 8,000 links that are on no page because pages were full: once each.
  assert.equal(found.dropped, 8000);
  // The file takes all of them.
  const pdf = new PdfDocument({ date: DATE });
  return (async () => {
    for (let i = 0; i < plan.pages.length; i++) {
      await pdf.addPage(page(found.pages[i], { width: 595.28, height: 841.89 }));
      pdf.takeParts();
    }
    await pdf.finish();
    assert.equal(pdf.links.internal, written);
    assert.equal(pdf.links.dropped, 0, 'nothing more is left out by the writer');
  })();
});

if (!hasPoppler() && !requirePoppler()) test('pdf links: Poppler is not installed', { skip: 'poppler not found: independent PDF checks skipped' }, () => undefined);
