// Tab stops of paragraphs and the default tab stop of a document: the change worked out
// against the stops of each paragraph's style, recorded where changes are tracked, refused
// where it cannot be, written into the file (only the parts that change), and laid out.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { DocLayout } from '../src/layout/doc/layout';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, TRACK_UNSUPPORTED, applyDocOp, paragraphTabStops, rejectAll, setDefaultTabStop, setParaFormat, type DocEdit, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { SECT, measurer, p, pos, range } from './doc-ops-fixtures';
import { buildDocx, type DocxSpec } from './docx-fixtures';
import { hasPython, packageProblems, runPython } from './edit-commands-fixtures';
import { defaultTabChange, formatLength, tabStopsChange } from '../src/ui/doc-tabs-dialog';

const decoder = new TextDecoder();

/** Normal; "Tabbed" (based on Normal) has a centred stop at 2880 and a right one with dots at 6000. */
const STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Test" w:hAnsi="Test"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Tabbed"><w:name w:val="Tabbed"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="2880"/><w:tab w:val="right" w:leader="dot" w:pos="6000"/></w:tabs></w:pPr></w:style>';

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildDocx({ body: body + SECT, styles: STYLES, settings: '<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>', ...extra });
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

const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((v, i) => v === b[i]);
const all = (o: Opened) => range(pos(0, 0), pos(o.doc.body.blocks.length - 1, 0));
const tabsXml = (xml: string) => (xml.match(/<w:tabs>.*?<\/w:tabs>/g) || []).join(' | ');

test('tab stops: a stop is set and cleared on a plain paragraph, the file gets <w:tabs>, and undo gives the file back', async () => {
  const o = await open(p('one') + p('two'));
  take(o, setParaFormat(o.doc, pos(0, 1), { tabs: { set: [{ pos: 1440, align: 'decimal', leader: 'dot' }, { pos: 720, align: 'left' }] } }));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), [{ pos: 720, align: 'left' }, { pos: 1440, align: 'decimal', leader: 'dot' }]);
  assert.deepEqual(paragraphTabStops(o.doc, pos(1, 0)), []);
  let bytes = await saved(o);
  assert.deepEqual(await changed(o.bytes, bytes), ['word/document.xml']);
  assert.equal(tabsXml(await part(bytes, 'word/document.xml')), '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="decimal" w:leader="dot" w:pos="1440"/></w:tabs>');
  take(o, setParaFormat(o.doc, pos(0, 1), { tabs: { clear: [720] } }));
  bytes = await saved(o);
  assert.equal(tabsXml(await part(bytes, 'word/document.xml')), '<w:tabs><w:tab w:val="decimal" w:leader="dot" w:pos="1440"/></w:tabs>');
  take(o, setParaFormat(o.doc, pos(0, 1), { tabs: null }));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), []);
  while (o.history.canUndo) o.history.undo();
  assert.ok(same(await saved(o), o.bytes), 'set, clear, undo, save: the file as it was');
});

test('tab stops: a stop of the style is cleared with a "clear" entry, and one the style has is not said again', async () => {
  const o = await open(p('one', '<w:pStyle w:val="Tabbed"/>') + p('two', '<w:pStyle w:val="Tabbed"/>'));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), [{ pos: 2880, align: 'center' }, { pos: 6000, align: 'right', leader: 'dot' }]);
  take(o, setParaFormat(o.doc, all(o), { tabs: { clear: [2880], set: [{ pos: 6000, align: 'right', leader: 'dot' }, { pos: 4000, align: 'bar' }] } }));
  for (const at of [0, 1]) {
    assert.deepEqual(paragraphTabStops(o.doc, pos(at, 0)), [{ pos: 4000, align: 'bar' }, { pos: 6000, align: 'right', leader: 'dot' }]);
    assert.deepEqual(o.doc.body.blocks[at].type === 'paragraph' && o.doc.body.blocks[at].props.tabs, [{ pos: 2880, align: 'clear' }, { pos: 4000, align: 'bar' }]);
  }
  const xml = await part(await saved(o), 'word/document.xml');
  assert.match(tabsXml(xml), /<w:tabs><w:tab w:val="clear" w:pos="2880"\/><w:tab w:val="bar" w:pos="4000"\/><\/w:tabs>/);
  // Clearing everything: every stop of the style is cleared, and the paragraph states nothing else.
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clearAll: true } }));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), []);
  assert.deepEqual(o.doc.body.blocks[0].type === 'paragraph' && o.doc.body.blocks[0].props.tabs, [{ pos: 2880, align: 'clear' }, { pos: 6000, align: 'clear' }]);
  // Back to what the style says: the paragraph's own stops go.
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: null }));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)).length, 2);
});

test('tab stops: what a file states beside what the model reads stays, and goes with the stops when they all go', async () => {
  const o = await open(p('one', '<w:tabs w:foo="1"><w:tab w:val="left" w:pos="1000"/></w:tabs>'));
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 2000, align: 'right' }] } }));
  let xml = await part(await saved(o), 'word/document.xml');
  assert.match(xml, /<w:tabs w:foo="1"><w:tab w:val="left" w:pos="1000"\/><w:tab w:val="right" w:pos="2000"\/><\/w:tabs>/);
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clearAll: true } }));
  xml = await part(await saved(o), 'word/document.xml');
  assert.doesNotMatch(xml, /<w:tabs/);
});

test('tab stops: what a change is given is checked', async () => {
  const o = await open(p('one'));
  const bad = (tabs: unknown) => assert.throws(() => setParaFormat(o.doc, pos(0, 0), { tabs: tabs as never }), RangeError, JSON.stringify(tabs));
  bad({ set: [{ pos: 1.5, align: 'left' }] });
  bad({ set: [{ pos: 40000, align: 'left' }] });
  bad({ set: [{ pos: 100, align: 'clear' }] });
  bad({ set: [{ pos: 100, align: 'left', leader: 'stars' }] });
  bad({ set: [{ pos: 100, align: 'left', extra: 1 }] });
  bad({ clear: ['100'] });
  bad({ other: [] });
  bad({ clearAll: 'yes' });
  bad([{ pos: 100, align: 'left' }, { pos: 100, align: 'right' }]);
  bad([{ pos: 100, align: 'nope' }]);
  // A list states the paragraph's own stops as they are, "clear" entries too (how a paragraph travels whole).
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: [{ pos: 500, align: 'right' }, { pos: 200, align: 'clear' }] }));
  assert.deepEqual(o.doc.body.blocks[0].type === 'paragraph' && o.doc.body.blocks[0].props.tabs, [{ pos: 200, align: 'clear' }, { pos: 500, align: 'right' }]);
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: [] }));
  assert.equal(o.doc.body.blocks[0].type === 'paragraph' && o.doc.body.blocks[0].props.tabs, undefined);
  bad({ set: Array.from({ length: 65 }, (_, i) => ({ pos: i * 10, align: 'left' })) });
  // 64 stops are as many as a paragraph takes; one more is refused.
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: Array.from({ length: 64 }, (_, i) => ({ pos: i * 10, align: 'left' })) } }));
  bad({ set: [{ pos: 9999, align: 'left' }] });
});

test('tab stops: where changes are tracked the change is recorded (pPrChange), and rejecting brings the stops back', async () => {
  const o = await open(p('one', '<w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs>'), { settings: '<w:trackRevisions/>' });
  const options = { author: 'Me', date: '2026-10-14T10:00:00Z' };
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 3000, align: 'center' }] } }, options));
  const xml = await part(await saved(o), 'word/document.xml');
  assert.match(xml, /<w:pPr><w:tabs><w:tab w:val="left" w:pos="1000"\/><w:tab w:val="center" w:pos="3000"\/><\/w:tabs><w:pPrChange w:id="\d+" w:author="Me" w:date="2026-10-14T10:00:00Z"><w:pPr><w:tabs><w:tab w:val="left" w:pos="1000"\/><\/w:tabs><\/w:pPr><\/w:pPrChange><\/w:pPr>/);
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), [{ pos: 1000, align: 'left' }, { pos: 3000, align: 'center' }]);
  take(o, rejectAll(o.doc));
  assert.deepEqual(paragraphTabStops(o.doc, pos(0, 0)), [{ pos: 1000, align: 'left' }]);
  // The default tab stop is a setting of the document Word keeps no record of: refused while tracking.
  assert.throws(() => setDefaultTabStop(o.doc, 360, pos(0, 0), options), (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
});

test('default tab stop: the setting is changed in its place, put in where the schema has it, and undone to the bytes of the file', async () => {
  let o = await open(p('one'));
  take(o, setDefaultTabStop(o.doc, 567, pos(0, 0)));
  assert.equal(o.doc.settings.defaultTabStop, 567);
  let bytes = await saved(o);
  assert.deepEqual(await changed(o.bytes, bytes), ['word/settings.xml']);
  assert.match(await part(bytes, 'word/settings.xml'), /<w:zoom w:percent="100"\/><w:defaultTabStop w:val="567"\/><w:characterSpacingControl/);
  assert.equal((await readDocx(await OpcPackage.open(bytes))).settings.defaultTabStop, 567);
  assert.deepEqual(setDefaultTabStop(o.doc, 567, pos(0, 0)).ops, [], 'the distance it is already');
  o.history.undo();
  assert.ok(same(await saved(o), o.bytes));

  // A file that states none: the element is put behind those the schema has before it.
  o = await open(p('one'), { settings: '<w:zoom w:percent="100"/><w:trackRevisions w:val="0"/><w:documentProtection w:edit="none"/><w:compat/>' });
  take(o, setDefaultTabStop(o.doc, 1134, pos(0, 0)));
  bytes = await saved(o);
  assert.match(await part(bytes, 'word/settings.xml'), /<w:documentProtection w:edit="none"\/><w:defaultTabStop w:val="1134"\/><w:compat\/>/);
  // A file without settings gets a part.
  o = await open(p('one'), { settings: undefined });
  take(o, setDefaultTabStop(o.doc, 360, pos(0, 0)));
  bytes = await saved(o);
  assert.match(await part(bytes, 'word/settings.xml'), /<w:settings [^>]*><w:defaultTabStop w:val="360"\/><\/w:settings>/);
  assert.equal((await readDocx(await OpcPackage.open(bytes))).settings.defaultTabStop, 360);
  for (const twips of [0, -5, 1.5, 40000, Number.NaN]) assert.throws(() => setDefaultTabStop(o.doc, twips, pos(0, 0)), RangeError, String(twips));
  // The operation takes whatever a reader takes from a file (its inverse puts that back), but nothing else.
  for (const twips of [-1, Number.NaN, 2000000, '1' as never]) assert.throws(() => applyDocOp(o.doc, { type: 'setDefaultTabStop', twips }), RangeError, String(twips));
  applyDocOp(o.doc, { type: 'setDefaultTabStop', twips: 0 });
});

test('tab stops in the layout: text after a tab starts at a set stop (left), ends at it (right), is centred on it, and the default stops follow the setting', async () => {
  const o = await open(p('<w:r><w:t>a</w:t><w:tab/><w:t>bb</w:t></w:r>'));
  const xOf = (doc: DocModel): number[] => {
    const layout = new DocLayout(doc, measurer, { lineMarks: true });
    layout.layout({});
    const xs: number[] = [];
    for (const page of layout.pages) for (const item of page.items) if (item.type === 'text' && item.text === 'bb') xs.push(Math.round(item.x));
    return xs;
  };
  const start = xOf(o.doc);
  assert.equal(start.length, 1, 'the text after the tab is found');
  take(o, setDefaultTabStop(o.doc, 1440, pos(0, 0)));
  const wider = xOf(o.doc)[0];
  // One inch (96 px) instead of half an inch (48 px) from the margin.
  assert.equal(wider - start[0], 48);
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 2880, align: 'left' }] } }));
  const left = xOf(o.doc)[0];
  assert.equal(left - start[0], 192 - 48);
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 2880, align: 'right' }] } }));
  const right = xOf(o.doc)[0];
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 2880, align: 'center' }] } }));
  const center = xOf(o.doc)[0];
  // "bb" ends at a right stop, and is centred on a centred one.
  assert.ok(right < center && center < left, `${right} < ${center} < ${left}`);
  assert.equal(left - right, 2 * (left - center));
});

const PYTHON = String.raw`
import json, os, sys, zipfile
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
folder = sys.argv[1]
out = {}
for name in sorted(os.listdir(folder)):
    z = zipfile.ZipFile(os.path.join(folder, name))
    doc = ET.fromstring(z.read('word/document.xml'))
    tabs = [[(t.get(W + 'val'), t.get(W + 'pos'), t.get(W + 'leader')) for t in group.iter(W + 'tab')] for group in doc.iter(W + 'tabs')]
    settings = ET.fromstring(z.read('word/settings.xml'))
    order = [c.tag.split('}')[1] for c in settings]
    stop = settings.find(W + 'defaultTabStop')
    out[name] = {'tabs': tabs, 'default': None if stop is None else stop.get(W + 'val'), 'order': order}
print(json.dumps(out))
`;

test('tab stops: an independent reading of the file (zipfile, xml.etree) finds them, and the package is sound', async (t) => {
  if (!hasPython(['xml.etree.ElementTree'])) {
    t.skip('python3 is not available');
    return;
  }
  const o = await open(p('one', '<w:pStyle w:val="Tabbed"/>'));
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clear: [6000], set: [{ pos: 1000, align: 'decimal', leader: 'hyphen' }] } }));
  take(o, setDefaultTabStop(o.doc, 851, pos(0, 0)));
  const files = { 'tabs.docx': await saved(o) };
  const read = runPython<{ [name: string]: { tabs: [string, string, string | null][][]; default: string; order: string[] } }>(files, PYTHON);
  assert.deepEqual(read['tabs.docx'].tabs, [
    [
      ['decimal', '1000', 'hyphen'],
      ['clear', '6000', null],
    ],
  ]);
  assert.equal(read['tabs.docx'].default, '851');
  assert.deepEqual(read['tabs.docx'].order, ['zoom', 'defaultTabStop', 'characterSpacingControl']);
  assert.deepEqual(packageProblems(files), []);
});

test('default tab stop: a distance the file states that no command sets (0, a fraction, more than Word takes, units) is put back by undo, and the file comes back byte for byte', async () => {
  for (const value of ['0', '40000', '720.5', '1in', '-5']) {
    const o = await open(p('one'), { settings: `<w:zoom w:percent="100"/><w:defaultTabStop w:val="${value}"/>` });
    const read = o.doc.settings.defaultTabStop;
    take(o, setDefaultTabStop(o.doc, 360, pos(0, 0)));
    assert.equal(o.doc.settings.defaultTabStop, 360, value);
    assert.match(await part(await saved(o), 'word/settings.xml'), /<w:defaultTabStop w:val="360"\/>/, value);
    assert.ok(o.history.undo(), value);
    assert.equal(o.doc.settings.defaultTabStop, read, value);
    assert.ok(same(await saved(o), o.bytes), `${value}: change, undo, save gives the file back`);
  }
  // The layout takes such a distance within what Word takes.
  const o = await open(p('<w:r><w:t>a</w:t><w:tab/><w:t>bb</w:t></w:r>'), { settings: '<w:defaultTabStop w:val="0"/>' });
  const layout = new DocLayout(o.doc, measurer, {});
  layout.layout({});
  assert.ok(layout.pages.length > 0);
});

test('tab stops: the limit of 64 counts the entries written, those that clear stops of the style too', async () => {
  const stops = Array.from({ length: 40 }, (_, i) => `<w:tab w:val="left" w:pos="${(i + 1) * 100}"/>`).join('');
  const styles = STYLES + `<w:style w:type="paragraph" w:styleId="Many"><w:name w:val="Many"/><w:pPr><w:tabs>${stops}</w:tabs></w:pPr></w:style>`;
  const o = await open(p('one', '<w:pStyle w:val="Many"/>'), { styles });
  // 40 cleared and 30 set: 30 stops apply, but 70 entries would be written.
  const set = Array.from({ length: 30 }, (_, i) => ({ pos: 10000 + i * 100, align: 'right' }));
  assert.throws(() => setParaFormat(o.doc, pos(0, 0), { tabs: { clearAll: true, set } }), /at most 64/);
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clearAll: true, set: set.slice(0, 24) } }));
  assert.equal(o.doc.body.blocks[0].type === 'paragraph' && o.doc.body.blocks[0].props.tabs!.length, 64);
});

test('tab stops: what a stop of the file holds beside what the model reads stays with that stop, and goes when it goes', async () => {
  const o = await open(p('one', '<w:tabs><w:tab w:val="left" w:pos="100"/><w:tab w:val="right" w:pos="900" w:extra="1"/></w:tabs>'));
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { set: [{ pos: 500, align: 'center' }], clear: [100] } }));
  let xml = tabsXml(await part(await saved(o), 'word/document.xml'));
  assert.equal(xml, '<w:tabs><w:tab w:val="center" w:pos="500"/><w:tab w:val="right" w:pos="900" w:extra="1"/></w:tabs>');
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clear: [900] } }));
  xml = tabsXml(await part(await saved(o), 'word/document.xml'));
  assert.equal(xml, '<w:tabs><w:tab w:val="center" w:pos="500"/></w:tabs>');
});

test('tab stops: a stop of the file is known by its position as the reader reads it, not by a "pos" of another namespace', async () => {
  const o = await open(p('one', '<w:tabs><w:tab w:val="left" w:pos="100" xmlns:o="urn:x" o:pos="900"/><w:tab w:val="right" w:pos="900" w:extra="1"/></w:tabs>'));
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clear: [100] } }));
  assert.equal(tabsXml(await part(await saved(o), 'word/document.xml')), '<w:tabs><w:tab w:val="right" w:pos="900" w:extra="1"/></w:tabs>');
  assert.deepEqual(paragraphTabStops(await readDocx(await OpcPackage.open(await saved(o))), pos(0, 0)), [{ pos: 900, align: 'right' }]);
});

test('tab stops: the stops a level of a list gives are worked out like those of a style', async () => {
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/><w:tab w:val="left" w:pos="1500"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const o = await open(p('one', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'), { numbering });
  const before = paragraphTabStops(o.doc, pos(0, 0));
  take(o, setParaFormat(o.doc, pos(0, 0), { tabs: { clear: [1500], set: [{ pos: 4000, align: 'left' }] } }));
  const after = paragraphTabStops(o.doc, pos(0, 0));
  assert.deepEqual(after.map((tab) => tab.pos), before.map((tab) => tab.pos).filter((at) => at !== 1500).concat([4000]).sort((a, b) => a - b));
  // Saved and read again, the paragraph has the same stops.
  const again = await readDocx(await OpcPackage.open(await saved(o)));
  assert.deepEqual(paragraphTabStops(again, pos(0, 0)), after);
});

test('tab stops and the default distance made on one document are one step: both apply, and one undo takes both back', async () => {
  const o = await open(p('bir') + p('iki'));
  const at = pos(0, 0);
  // As the view makes the step of the dialog (applyTabChanges): both commands on the document as it is.
  const stops = setParaFormat(o.doc, at, { tabs: { set: [{ pos: 1440, align: 'right', leader: 'dot' }] } });
  const setting = setDefaultTabStop(o.doc, 567, at);
  o.history.apply(stops.ops.concat(setting.ops), undefined, undefined);
  assert.deepEqual(paragraphTabStops(o.doc, at), [{ pos: 1440, align: 'right', leader: 'dot' }]);
  assert.equal(o.doc.settings.defaultTabStop, 567);
  assert.ok(o.history.undo());
  assert.deepEqual(paragraphTabStops(o.doc, at), []);
  assert.equal(o.doc.settings.defaultTabStop, 720);
  assert.equal(o.history.canUndo, false);
  assert.ok(same(await saved(o), o.bytes));
});

test('tab stops dialog: the default distance is handed over only when its field was changed', () => {
  // 708 twips shows as 1,25 cm, which read back is 709: untouched, nothing is handed over.
  const shown = formatLength(708, 'cm', ',');
  assert.equal(shown, '1,25');
  assert.equal(defaultTabChange(shown, shown, 708, 'cm', ','), undefined);
  assert.equal(defaultTabChange(formatLength(708, 'in', '.'), formatLength(708, 'in', '.'), 708, 'in', '.'), undefined);
  assert.equal(defaultTabChange(shown, ' 1,25 ', 708, 'cm', ','), undefined);
  assert.equal(defaultTabChange(shown, '1', 708, 'cm', ','), 567);
  assert.equal(defaultTabChange(shown, 'x', 708, 'cm', ','), null);
  assert.equal(defaultTabChange(shown, '0', 708, 'cm', ','), null);
  assert.deepEqual(tabStopsChange([{ pos: 100, align: 'left' }], [{ pos: 100, align: 'left' }], false), undefined);
  assert.deepEqual(tabStopsChange([{ pos: 100, align: 'left' }], [{ pos: 200, align: 'right', leader: 'dot' }], false), { set: [{ pos: 200, align: 'right', leader: 'dot' }], clear: [100] });
});
