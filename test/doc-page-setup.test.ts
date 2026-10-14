// The page of a document, section by section: paper, orientation and margins of the section of
// the caret or of every section, and section breaks. The commands, their undo, what the writer
// makes of them (only the section that changed; what the section says beside the page stays), the
// layout that follows, the refusal where changes are tracked, and a reading of the result by
// python's xml.etree and package_check.py where python3 has them.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { DocLayout } from '../src/layout/doc/layout';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, TRACK_UNSUPPORTED, documentSections, insertSectionBreak, sectionSetup, sectionsAt, setSectionSetup, type DocEdit, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { measurer, p, pos, range } from './doc-ops-fixtures';
import { buildDocx } from './docx-fixtures';
import { hasPython, packageProblems, runPython } from './edit-commands-fixtures';

const decoder = new TextDecoder();

/** The section of the body: A4 upright, with a header, columns and an element the model does not know. */
const LAST =
  '<w:sectPr w:rsidR="00AB12CD"><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="11906" w:h="16838" w:code="9"/>' +
  '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
/** A section a paragraph ends: Letter, landscape, a fixed top margin. */
const FIRST = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="-1000" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(body: string, settings = ''): Promise<Opened> {
  const bytes = await buildDocx({ body, headers: { rIdH1: '<w:p><w:r><w:t>Head</w:t></w:r></w:p>' }, settings });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

const take = (o: Opened, edit: DocEdit) => o.history.apply(edit.ops, undefined, undefined);

async function saved(o: Opened): Promise<Uint8Array> {
  await writeDocx(o.pkg, o.doc);
  return o.pkg.save();
}

const part = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));
const sectPrs = (xml: string): string[] => Array.from(xml.matchAll(/<w:sectPr[\s>][^]*?<\/w:sectPr>|<w:sectPr\/>/g)).map((match) => match[0]);

async function changed(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = ZipArchive.read(before);
  const b = ZipArchive.read(after);
  const out: string[] = [];
  for (const name of new Set(a.names().concat(b.names()))) {
    const x = a.has(name) ? await a.read(name) : undefined;
    const y = b.has(name) ? await b.read(name) : undefined;
    if (!x || !y || x.length !== y.length || x.some((v, i) => v !== y[i])) out.push(name);
  }
  return out.sort();
}

const pageSizes = (doc: DocModel): [number, number, number][] => {
  const layout = new DocLayout(doc, measurer, {});
  layout.layout({});
  return layout.pages.map((page) => [Math.round(page.width), Math.round(page.height), page.section]);
};

test('section setup: what a section says, and the papers it is known by', () => {
  const a4 = sectionSetup({ page: { width: 11906, height: 16838 }, margins: { top: -1000, left: 1417 } });
  assert.equal(a4.paper, 'a4');
  assert.equal(a4.orientation, 'portrait');
  assert.deepEqual(a4.margins, { top: 1000, bottom: 1440, left: 1417, right: 1440, header: 720, footer: 720, gutter: 0 });
  const letter = sectionSetup({ page: { width: 15840, height: 12240, orientation: 'landscape' } });
  assert.deepEqual([letter.paper, letter.orientation], ['letter', 'landscape']);
  const none = sectionSetup({});
  assert.deepEqual([none.width, none.height, none.paper], [12240, 15840, 'letter']);
});

test('section setup: the section of the caret changes; the others, the header and what the model does not know stay', async () => {
  const o = await open(p('one', FIRST) + p('two') + LAST);
  assert.equal(documentSections(o.doc).length, 2);
  assert.deepEqual(sectionsAt(o.doc, pos(1, 0)), [1]);
  assert.deepEqual(sectionsAt(o.doc, range(pos(0, 0), pos(1, 1))), [0, 1]);
  const edit = setSectionSetup(o.doc, pos(1, 1), { paper: 'letter', orientation: 'landscape', margins: { left: 720, right: 720 } });
  take(o, edit);
  assert.deepEqual(edit.selection, range(pos(1, 1), pos(1, 1)));
  const now = sectionSetup(o.doc.body.section!);
  assert.deepEqual([now.width, now.height, now.orientation, now.paper, now.margins.left, now.margins.top], [15840, 12240, 'landscape', 'letter', 720, 1417]);
  assert.deepEqual(pageSizes(o.doc), [
    [1056, 816, 0],
    [1056, 816, 1],
  ]);
  const bytes = await saved(o);
  assert.deepEqual(await changed(o.bytes, bytes), ['word/document.xml']);
  const xml = await part(bytes, 'word/document.xml');
  const sections = sectPrs(xml);
  assert.equal(sections.length, 2);
  assert.equal(sections[0], FIRST, 'the first section is written as the file has it');
  assert.match(sections[1], /<w:headerReference w:type="default" r:id="rIdH1"\/>/);
  assert.match(sections[1], /<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/, 'the code of the old paper goes');
  assert.match(sections[1], /<w:pgMar w:top="1417" w:right="720" w:bottom="1417" w:left="720" w:header="708" w:footer="708" w:gutter="0"\/>/);
  assert.match(sections[1], /<w:cols w:space="708"\/>/);
  assert.match(sections[1], /<w:docGrid w:linePitch="360"\/>/);
  assert.match(sections[1], /w:rsidR="00AB12CD"/);
  // Read back: the same.
  const again = await readDocx(await OpcPackage.open(bytes));
  assert.deepEqual(sectionSetup(again.body.section!).margins, sectionSetup(o.doc.body.section!).margins);
  // Undo: the file as it was.
  o.history.undo();
  assert.deepEqual(await changed(o.bytes, await saved(o)), []);
});

test('section setup: a section a paragraph ends; a fixed margin stays fixed; the whole document', async () => {
  const o = await open(p('one', FIRST) + p('two') + LAST);
  take(o, setSectionSetup(o.doc, pos(0, 0), { margins: { top: 2000 }, orientation: 'portrait' }));
  const first = o.doc.body.blocks[0];
  assert.ok(first.type === 'paragraph' && first.props.section);
  if (first.type === 'paragraph' && first.props.section) {
    assert.equal(first.props.section.margins?.top, -2000);
    assert.deepEqual(first.props.section.page, { width: 12240, height: 15840, orientation: 'portrait' });
  }
  assert.deepEqual(pageSizes(o.doc)[0], [816, 1056, 0]);
  const xml = await part(await saved(o), 'word/document.xml');
  assert.match(sectPrs(xml)[0], /<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"\/><w:pgMar w:top="-2000"/);
  assert.equal(sectPrs(xml)[1], LAST);
  o.history.undo();
  take(o, setSectionSetup(o.doc, pos(0, 0), { paper: 'a5' }, 'document'));
  const sizes = documentSections(o.doc).map((section) => [section.props.page?.width, section.props.page?.height, section.props.page?.orientation]);
  assert.deepEqual(sizes, [
    [11906, 8391, 'landscape'],
    [8391, 11906, undefined],
  ]);
});

test('section setup: nothing changes, nothing is done; what cannot be is refused', async () => {
  const o = await open(p('one') + LAST);
  assert.deepEqual(setSectionSetup(o.doc, pos(0, 0), { paper: 'a4', orientation: 'portrait', margins: { top: 1417 } }).ops, []);
  assert.throws(() => setSectionSetup(o.doc, pos(0, 0), { paper: 'tabloid-xxl' }), /paper/);
  assert.throws(() => setSectionSetup(o.doc, pos(0, 0), { margins: { left: 6000, right: 6000 } }), /no room/);
  assert.throws(() => setSectionSetup(o.doc, pos(0, 0), { margins: { top: -5 } }), /margin/);
  assert.throws(() => setSectionSetup(o.doc, pos(0, 0), { paper: { width: 50, height: 50 } }), /paper/);
});

test('section setup: a body without a section gets one, and undo takes it away again', async () => {
  const o = await open(p('one'));
  assert.equal(o.doc.body.section, undefined);
  take(o, setSectionSetup(o.doc, pos(0, 0), { paper: 'a4' }));
  assert.deepEqual(o.doc.body.section, { page: { width: 11906, height: 16838 } });
  assert.match(await part(await saved(o), 'word/document.xml'), /<w:sectPr><w:pgSz w:w="11906" w:h="16838"\/><\/w:sectPr><\/w:body>/);
  o.history.undo();
  assert.equal(o.doc.body.section, undefined);
  assert.deepEqual(await changed(o.bytes, await saved(o)), []);
});

test('section break: the paragraph is cut, the first half ends a copy of its section, the next section starts as asked', async () => {
  const o = await open(p('abcdef') + LAST);
  const edit = insertSectionBreak(o.doc, pos(0, 3), 'continuous');
  take(o, edit);
  assert.deepEqual(edit.selection.focus, pos(1, 0));
  const sections = documentSections(o.doc);
  assert.equal(sections.length, 2);
  assert.equal(sections[0].end, 0);
  assert.equal(sections[1].props.type, 'continuous');
  assert.equal(sections[0].props.type, undefined);
  assert.deepEqual(sections[0].props.headers, { default: 'rIdH1' });
  assert.equal(sections[0].props.attrs, undefined, 'the copy has no rsid of the file');
  const bytes = await saved(o);
  const xml = await part(bytes, 'word/document.xml');
  assert.match(xml, /<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"\/><w:pgSz w:w="11906" w:h="16838" w:code="9"\/>[^]*?<\/w:sectPr><\/w:pPr><w:r><w:t>abc<\/w:t><\/w:r><\/w:p>/);
  assert.match(xml, /<w:p><w:r><w:t>def<\/w:t><\/w:r><\/w:p><w:sectPr w:rsidR="00AB12CD"><w:headerReference[^]*?<w:type w:val="continuous"\/>/);
  // A next-page break where the section has no type: the section of the body stays as the file has it.
  const q = await open(p('abcdef') + LAST);
  take(q, insertSectionBreak(q.doc, pos(0, 6), 'nextPage'));
  assert.equal(sectPrs(await part(await saved(q), 'word/document.xml'))[1], LAST);
  assert.deepEqual(
    pageSizes(q.doc).map((page) => page[2]),
    [0, 1],
  );
  // An odd-page break inside a section a paragraph ends.
  const r = await open(p('ab') + p('cd', FIRST) + LAST);
  take(r, insertSectionBreak(r.doc, pos(0, 1), 'oddPage'));
  const after = documentSections(r.doc);
  assert.equal(after.length, 3);
  assert.equal(after[1].props.type, 'oddPage');
  assert.equal(after[1].end, 2);
  assert.deepEqual(after[0].props.page, after[1].props.page);
  r.history.undo();
  assert.deepEqual(await changed(r.bytes, await saved(r)), []);
});

test('section break: not in a table, a header, or where changes are tracked; page setup refused there too', async () => {
  const table = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const o = await open(table + p('a') + LAST);
  assert.throws(() => insertSectionBreak(o.doc, pos([0, 0, 0, 0], 0), 'nextPage'), /table cell/);
  assert.throws(() => insertSectionBreak(o.doc, pos(0, 0, 'header:rIdH1'), 'nextPage'), /body/);
  assert.throws(() => insertSectionBreak(o.doc, pos(1, 0), 'sideways' as 'nextPage'), /section break/);
  const tracked = await open(p('abc') + LAST, '<w:trackRevisions/>');
  assert.throws(() => insertSectionBreak(tracked.doc, pos(0, 1), 'nextPage'), (error: Error) => error.message === TRACK_UNSUPPORTED);
  assert.throws(() => setSectionSetup(tracked.doc, pos(0, 1), { orientation: 'landscape' }), (error: Error) => error.message === TRACK_UNSUPPORTED);
});

const ETREE = String.raw`
import json, os, sys, zipfile
import xml.etree.ElementTree as E
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
folder = sys.argv[1]
out = {}
for name in sorted(os.listdir(folder)):
    with zipfile.ZipFile(os.path.join(folder, name)) as z:
        root = E.fromstring(z.read('word/document.xml'))
    sections = []
    for sect in root.iter(W + 'sectPr'):
        size = sect.find(W + 'pgSz')
        margins = sect.find(W + 'pgMar')
        kind = sect.find(W + 'type')
        sections.append({
            'w': size.get(W + 'w') if size is not None else None,
            'h': size.get(W + 'h') if size is not None else None,
            'orient': size.get(W + 'orient') if size is not None else None,
            'left': margins.get(W + 'left') if margins is not None else None,
            'type': kind.get(W + 'val') if kind is not None else None,
        })
    out[name] = sections
print(json.dumps(out))
`;

test('section setup and breaks: python reads the sections the engine wrote, and the packages are sound', async (t) => {
  if (!hasPython(['xml.etree.ElementTree', 'zipfile'])) {
    t.skip('python3 is not available');
    return;
  }
  const o = await open(p('abcdef') + LAST);
  take(o, insertSectionBreak(o.doc, pos(0, 3), 'continuous'));
  take(o, setSectionSetup(o.doc, pos(1, 0), { orientation: 'landscape', margins: { left: 1000 } }));
  const files = { 'sections.docx': await saved(o) };
  const read = runPython<{ [name: string]: unknown[] }>(files, ETREE);
  assert.deepEqual(read['sections.docx'], [
    { w: '11906', h: '16838', orient: null, left: '1417', type: null },
    { w: '16838', h: '11906', orient: 'landscape', left: '1000', type: 'continuous' },
  ]);
  assert.deepEqual(packageProblems(files), []);
});

test('section setup: margins a section does not state are written with the one that changed, all seven of them', async () => {
  const o = await open(p('one') + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>');
  take(o, setSectionSetup(o.doc, pos(0, 0), { margins: { top: 2000 } }));
  const bytes = await saved(o);
  assert.match(sectPrs(await part(bytes, 'word/document.xml'))[0], /<w:pgMar w:top="2000" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"\/>/);
  if (hasPython(['xml.etree.ElementTree', 'zipfile'])) assert.deepEqual(packageProblems({ 'margins.docx': bytes }), []);
  // A section that states some of them keeps those.
  const q = await open(p('one') + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:left="900"/></w:sectPr>');
  take(q, setSectionSetup(q.doc, pos(0, 0), { margins: { right: 800 } }));
  assert.match(sectPrs(await part(await saved(q), 'word/document.xml'))[0], /<w:pgMar w:top="1000" w:right="800" w:bottom="1440" w:left="900" w:header="720" w:footer="720" w:gutter="0"\/>/);
});

test('package check: a w:pgMar without all of its attributes is found', async (t) => {
  if (!hasPython(['xml.etree.ElementTree', 'zipfile'])) {
    t.skip('python3 is not available');
    return;
  }
  const bytes = await buildDocx({ body: p('one') + '<w:sectPr><w:pgMar w:top="1000" w:left="900"/></w:sectPr>' });
  const problems = packageProblems({ 'partial.docx': bytes });
  assert.equal(problems.length, 1);
  assert.match(JSON.stringify(problems), /pgMar has no right, bottom, header, footer, gutter/);
});

test('section break: a section with a recorded change is copied without it, and no revision id comes twice', async () => {
  const S =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:sectPrChange w:id="41" w:author="A" w:date="2026-01-01T00:00:00Z"><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>';
  const o = await open(p('abcdef') + S);
  take(o, insertSectionBreak(o.doc, pos(0, 3), 'nextPage'));
  const xml = await part(await saved(o), 'word/document.xml');
  assert.equal((xml.match(/<w:sectPrChange /g) || []).length, 1, 'the record stays with the section it was made for');
  assert.equal((xml.match(/w:id="41"/g) || []).length, 1);
  assert.doesNotMatch(sectPrs(xml)[0], /sectPrChange/);
});
