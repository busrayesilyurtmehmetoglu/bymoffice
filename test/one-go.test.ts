// A file is written from its model in one go: what the writer needs of the package is read first,
// then the model is read and every part produced without waiting for anything. What is changed
// while a file is written neither gets half into it nor makes the write start over.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { lastOneGo, oneGo } from '../src/formats/part-reads';
import { readCollabStamp, writeSharedFile } from '../src/formats/xlsx/collab-stamp';
import { buildXlsx } from '../src/formats/xlsx/build';
import { buildDocx, writeDocx } from '../src/formats/docx/write';
import { readDocx } from '../src/formats/docx/document';
import { writeXlsx } from '../src/formats/xlsx/write';
import { OpcPackage, ZipArchive, open, type Workbook } from '../src/index';
import { SaveInterruptedError, editsOf, readModel, readMomentOf, readingsOf, touchModel, writeSteady } from '../src/model/edits';
import { setCellInput } from '../src/ops/commands';
import { DocHistory, insertText } from '../src/ops/doc/index';
import { applyOp } from '../src/ops/ops';
import { packageTurn } from '../src/package-turn';
import { bookOf } from './chart-edit-fixtures';
import { Net, act, join, seeded, startFile } from './collab-fixtures';
import { p } from './doc-ops-fixtures';
import { buildDocx as buildSpec } from './docx-fixtures';
import { changeSheet, docEdit, docOpen, docRun, docStart, load, sameBytes, sheetRun, sheetStart } from './one-go-fixtures';

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89]);

test('a model is read for a file in one go: the moment is told, and what waits or changes the model meanwhile is refused', async () => {
  const model = {};
  touchModel(model);
  touchModel(model);
  assert.equal(readMomentOf(model), undefined, 'no file was written from it yet');
  assert.equal(readModel(model, () => 7), 7);
  assert.equal(readMomentOf(model), 2, 'the moment is the count of changes the reading began at');
  assert.equal(readingsOf(model), 1);
  touchModel(model);
  assert.equal(readMomentOf(model), 2, 'a change after the reading is not of that moment');
  // What waits is no reading in one go, whether or not anything changed meanwhile.
  assert.throws(() => readModel(model, async () => 1), /may not wait/, 'a reading that answers a promise is refused');
  assert.throws(() => readModel(model, () => ({ then: () => undefined })), /may not wait/);
  // Nor is one during which the model is changed (the writer itself would have to do it: nothing else runs).
  assert.throws(() => readModel(model, () => touchModel(model)), /was changed while it was read/);
  assert.equal(readMomentOf(model), 2, 'a reading that was refused is the moment of no file');
  // A reading inside a reading of the same model is the same moment.
  assert.equal(readModel(model, () => readModel(model, () => editsOf(model))), 4);
  assert.equal(readMomentOf(model), 4);
  // What a reading throws is thrown, and the model is free to be read again.
  assert.throws(() => readModel(model, () => { throw new RangeError('no'); }), RangeError);
  assert.equal(readModel(model, () => 'again'), 'again');
});

test('a writer that reads its model in one go is not asked again, however the model changes around the reading; one that does not still is', async () => {
  const model = {};
  let writes = 0;
  const out = await writeSteady(model, async () => {
    writes++;
    touchModel(model);
    await new Promise((resolve) => setImmediate(resolve));
    const read = readModel(model, () => editsOf(model));
    await new Promise((resolve) => setImmediate(resolve));
    touchModel(model);
    return read;
  });
  assert.deepEqual([writes, out, readMomentOf(model)], [1, 1, 1], 'written once: the file is the model of the moment it was read at');
  // What it throws after it read is thrown, not tried again.
  writes = 0;
  await assert.rejects(
    writeSteady(model, async () => {
      writes++;
      readModel(model, () => undefined);
      touchModel(model);
      throw new Error('could not be packed');
    }),
    /could not be packed/,
  );
  assert.equal(writes, 1);
  // A writer that says nothing of how it read: as before, written again once the changes pause.
  writes = 0;
  const steady = await writeSteady(model, async () => (++writes === 1 ? (touchModel(model), writes) : writes), { sleep: async () => undefined, now: () => 0 });
  assert.deepEqual([steady, writes], [2, 2]);
});

test('the writers ask for the parts they need before they read the model: a usual save reads the model once, and what was not foreseen is found all the same', async () => {
  const e = await load(process.cwd());
  const start = await sheetStart(e);
  // Each kind of change by itself, on the file as it was opened: one reading.
  const more: { [kind: string]: number } = {};
  for (const kind of ['type', 'style', 'note', 'validation', 'rule+', 'chart+', 'insert rows', 'delete columns', 'merge', 'sort', 'sheet+', 'sheet-', 'rename', 'move sheet', 'copy sheet', 'tab', 'filter']) {
    const file = await open(start);
    const wb = await file.workbook();
    let made = false;
    for (let seed = 1; seed < 200 && !made; seed++) made = !!changeSheet(e, wb, seeded(seed), kind);
    assert.ok(made, `${kind}: applies`);
    await file.save();
    const trace = lastOneGo(wb);
    assert.ok(trace, `${kind}: the writer read the workbook in one go`);
    if (trace.readings !== 1) more[kind] = trace.readings;
  }
  assert.deepEqual(more, {}, 'kinds of change whose save read the workbook more than once');
  // A document: text, a comment, a picture.
  for (const kind of ['insertText', 'add', 'insertImage', 'toggleList', 'insertTable']) {
    const o = await docOpen(e, await docStart(e, 3));
    let made = false;
    for (let seed = 1; seed < 600 && !made; seed++) {
      const edit = docEdit(e, o.doc, e.docFixtures.dice(seed), {}, kind === 'add' ? 'add' : undefined);
      if (!edit || edit.name !== kind || edit.edit.ops.length === 0) continue;
      o.history.applyEdit(edit.edit, undefined);
      made = true;
    }
    assert.ok(made, `${kind}: applies`);
    await writeDocx(o.pkg as OpcPackage, o.doc as never);
    assert.equal((lastOneGo(o.doc as object) || { readings: 0 }).readings, 1, `${kind}: the document was read once`);
  }
});

test('random editing with saves in between: the files are the same whether the parts are read ahead or each is asked for by a reading that starts over', async () => {
  const e = await load(process.cwd());
  const start = await sheetStart(e);
  const SEEDS = Math.max(1, Number(process.env.ONE_GO_SEEDS) || 25);
  let files = 0;
  let readings = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    for (const kind of ['xlsx', 'docx']) {
      const run = (): Promise<{ files: Uint8Array[]; trail: string[] }> => (kind === 'xlsx' ? sheetRun(e, seed, start) : docRun(e, seed, seed % 5));
      oneGo.foresee = true;
      const ahead = await run();
      oneGo.foresee = false;
      let again: { files: Uint8Array[]; trail: string[] };
      try {
        again = await run();
      } finally {
        oneGo.foresee = true;
      }
      assert.deepEqual(again.trail, ahead.trail, `${kind}, seed ${seed}: the same run`);
      assert.equal(again.files.length, ahead.files.length);
      ahead.files.forEach((bytes, i) => assert.ok(sameBytes(bytes, again.files[i]), `${kind}, seed ${seed}, file ${i} (${ahead.trail.join(', ')}): the same bytes`));
      files += ahead.files.length;
    }
  }
  // Without reading ahead a save does start over (or the run above tested nothing).
  oneGo.foresee = false;
  try {
    const file = await open(start);
    const wb = await file.workbook();
    setCellInput(wb, 0, 3, 1, '5').forEach((op) => applyOp(wb, op));
    await file.save();
    readings = (lastOneGo(wb) || { readings: 0 }).readings;
  } finally {
    oneGo.foresee = true;
  }
  assert.ok(readings > 1, `a save that reads nothing ahead reads the workbook again for each part it finds missing (${readings} readings)`);
  assert.ok(files > SEEDS * 4, `${files} files were compared`);
});

test('a part that cannot be read fails the save that needs it, and no other', async () => {
  // A package with a damaged part nothing of the save looks at: read ahead with the rest, and in nobody's way.
  const archive = ZipArchive.read(await startFile());
  const names = archive.names();
  const sheet2 = names.find((name) => /sheet2\.xml$/.test(name)) as string;
  const whole = await archive.write();
  // The stored bytes of the second sheet are damaged in the archive (its checksum no longer fits).
  const entry = ZipArchive.read(whole).entry(sheet2)!;
  const damaged = whole.slice();
  const marker = new TextEncoder().encode(sheet2);
  let at = -1;
  for (let i = 0; i + marker.length < damaged.length && at < 0; i++) {
    if (damaged[i] === 0x50 && damaged[i + 1] === 0x4b && damaged[i + 2] === 3 && damaged[i + 3] === 4) {
      let same = true;
      for (let k = 0; k < marker.length && same; k++) same = damaged[i + 30 + k] === marker[k];
      if (same) at = i + 30 + marker.length + (damaged[i + 28] | (damaged[i + 29] << 8));
    }
  }
  assert.ok(at > 0 && entry.compressedSize > 8, 'the entry was found');
  for (let i = 2; i < 8; i++) damaged[at + i] ^= 0xff;
  const pkg = await OpcPackage.open(damaged);
  await assert.rejects(pkg.readPart(`/${sheet2}`), /Corrupt|corrupt/, 'the part does not unpack');
  // The workbook is made by hand for the first sheet alone: what is written does not touch the second.
  const good = await (await open(whole)).workbook();
  const wb: Workbook = { ...good, sheets: [good.sheets[0], { ...good.sheets[1], rows: new Map() }] };
  setCellInput(wb, 0, 2, 1, '77').forEach((op) => applyOp(wb, op));
  await writeXlsx(pkg, wb);
  const out = ZipArchive.read(await pkg.save());
  assert.ok(new TextDecoder().decode(await out.read(names.find((name) => /sheet1\.xml$/.test(name)) as string)).indexOf('<v>77</v>') > 0, 'the first sheet was written');
  // A change of the second sheet needs the part: that save says what is wrong with it.
  setCellInput(wb, 1, 0, 0, '5').forEach((op) => applyOp(wb, op));
  await assert.rejects(writeXlsx(pkg, wb), /Corrupt|corrupt/);
});

test('a workbook that is edited all the time is saved all the same: one write, in the time a write takes', async () => {
  const rows: (number | string | null)[][] = [];
  for (let r = 0; r < 100000; r++) rows.push([`Row ${r}`, r, r * 2.5, `t${r % 977}`]);
  const file = await open(await buildXlsx(bookOf(rows)));
  const wb = await file.workbook();
  const pkg = file.package!;
  const packed = pkg.save.bind(pkg);
  let packs = 0;
  pkg.save = () => (packs++, packed());
  // In peace first: how long a write takes.
  setCellInput(wb, 0, 5, 1, '1').forEach((op) => applyOp(wb, op));
  let started = Date.now();
  await file.save();
  const calm = Date.now() - started;
  // Then with a change every 50 ms, from before the save until after it.
  let edits = 0;
  const typing = setInterval(() => {
    setCellInput(wb, 0, 10 + (edits++ % 50), 1, String(1000000 + edits)).forEach((op) => applyOp(wb, op));
  }, 50);
  try {
    await new Promise((resolve) => setTimeout(resolve, 120));
    for (let round = 0; round < 3; round++) {
      packs = 0;
      const before = edits;
      started = Date.now();
      const bytes = await file.save();
      const took = Date.now() - started;
      assert.equal(packs, 1, `round ${round}: the file was written once`);
      assert.ok(took < calm * 4 + 1500, `round ${round}: in the time a write takes (${took} ms; ${calm} ms in peace)`);
      assert.ok(edits > before || took < 50, `round ${round}: the workbook was changed while the file was written (or the test tests nothing)`);
      // The file is the workbook of the moment it was read at: every change up to one of them, and none after.
      const moment = readMomentOf(wb);
      assert.ok(moment !== undefined && moment <= editsOf(wb));
      const again = await (await open(bytes)).workbook();
      const values: number[] = [];
      for (let i = 0; i < 50; i++) values.push(Number((again.sheets[0].rows.get(10 + i)!.cells[1] || { value: 0 }).value) - 1000000);
      const newest = Math.max(...values);
      assert.ok(newest >= before && newest <= edits, `round ${round}: the file holds the changes made before it was read (${newest}; ${before} before the save, ${edits} after)`);
    }
  } finally {
    clearInterval(typing);
  }
});

test('a document that is typed into all the time is saved all the same: one write', async () => {
  let body = '';
  for (let i = 0; i < 2000; i++) body += p(`Paragraph ${i} with a few words in it, as paragraphs have.`);
  const file = await open(await buildSpec({ body }));
  const doc = await file.document();
  const history = new DocHistory(doc);
  const pkg = file.package!;
  const packed = pkg.save.bind(pkg);
  let packs = 0;
  pkg.save = () => (packs++, packed());
  let typed = 0;
  const type = (): void => {
    history.applyEdit(insertText(doc, { story: 'body', path: [typed % 1500], offset: 0 }, 'x'), undefined);
    history.seal();
    typed++;
  };
  type();
  const typing = setInterval(type, 50);
  try {
    await new Promise((resolve) => setTimeout(resolve, 120));
    for (let round = 0; round < 3; round++) {
      packs = 0;
      const before = typed;
      const bytes = await file.save().catch((error: unknown) => error);
      assert.ok(!(bytes instanceof SaveInterruptedError) && bytes instanceof Uint8Array, `round ${round}: the document was saved (${String(bytes).slice(0, 80)})`);
      assert.equal(packs, 1, `round ${round}: the file was written once`);
      const again = await (await open(bytes)).document();
      const text = JSON.stringify(again.body.blocks.slice(0, 1500));
      const held = text.split('"xParagraph ').length - 1 + (text.split('"xxParagraph ').length - 1);
      assert.ok(held >= Math.min(before, 1500) && held <= typed, `round ${round}: the file holds what was typed before it was read (${held}; ${before} before the save, ${typed} after)`);
    }
  } finally {
    clearInterval(typing);
  }
});

test('a file written while a calculation is under way asks the application to calculate it', async () => {
  const file = await open(await startFile());
  const wb = await file.workbook();
  // A budget that one formula uses up: the calculation stops with the rest left for the next call.
  const engine = new CalcEngine(wb, { budget: { work: 1, hardWork: 1000000 } });
  setCellInput(wb, 0, 1, 1, '55').forEach((op) => applyOp(wb, op));
  engine.recalculate([{ sheet: 0, row: 1, col: 1 }]);
  assert.ok(engine.incomplete, 'the calculation is not done (or the test tests nothing)');
  const bytes = await file.save();
  const pkg = await OpcPackage.open(bytes);
  const book = new TextDecoder().decode(await pkg.readPart((await pkg.mainPart()) as string));
  assert.ok(/fullCalcOnLoad="1"/.test(book), 'the file says that its results are to be calculated when it is opened');
  // The results that arrive afterwards are changes like any other: the workbook is another one than the file that was read.
  const moment = readMomentOf(wb) as number;
  for (let i = 0; i < 100 && engine.incomplete; i++) engine.recalculate([]);
  assert.ok(!engine.incomplete);
  assert.ok(editsOf(wb) > moment, 'what the calculation wrote since counts as changed after the file');
});

test('a file that is built anew reads its model in one go as well: what is changed while its pictures are read is in it', async () => {
  // A workbook with a picture, whose bytes take a moment to arrive; a cell is typed meanwhile.
  const book = bookOf([['a', 1], ['b', 2]]);
  book.sheets[0].drawings = [{ anchor: { kind: 'oneCell', from: { col: 1, colOff: 0, row: 2, rowOff: 0 }, width: 1905000, height: 952500 }, object: { type: 'picture', media: 'one', name: 'Picture' } }] as never;
  let reads = 0;
  book.media = {
    readPart: async () => {
      reads++;
      await new Promise((resolve) => setImmediate(resolve));
      setCellInput(book, 0, 0, 1, '4711').forEach((op) => applyOp(book, op));
      return PNG;
    },
  };
  const before = readingsOf(book);
  const built = await (await open(await buildXlsx(book))).workbook();
  assert.equal(reads, 1);
  assert.equal(built.sheets[0].rows.get(0)!.cells[1]!.value, 4711, 'the workbook as it was when it was read, after the picture had arrived');
  assert.equal(readingsOf(book) - before, 1, 'read once');
  assert.equal(readMomentOf(book), editsOf(book));

  // A document of a package with a picture: the same.
  const bytes = await buildSpec({
    body:
      '<w:p><w:r><w:t>text</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="200"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdI"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    rels: [['rIdI', 'image', 'media/image7.png']],
    parts: [['word/media/image7.png', PNG]],
  });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  let asked = 0;
  const report = { parts: [], copiedBlocks: 0, writtenBlocks: 0, droppedCharacters: 0, warnings: [] as string[] };
  const out = await buildDocx(doc, {
    report,
    pictures: {
      readPart: async (part: string) => {
        asked++;
        await new Promise((resolve) => setImmediate(resolve));
        history.applyEdit(insertText(doc, { story: 'body', path: [0], offset: 0 }, 'typed meanwhile '), undefined);
        return pkg.readPart(part);
      },
    },
  });
  assert.equal(asked, 1);
  const again = await OpcPackage.open(out);
  assert.ok(new TextDecoder().decode(await again.readPart('/word/document.xml')).indexOf('typed meanwhile') > 0, 'the document as it was when it was read');
  assert.deepEqual(await again.readPart('/word/media/image7.png'), PNG);
  assert.equal(readMomentOf(doc), editsOf(doc));
});

test('the file of a room says its step when something is typed after the workbook was read, and is written once; typed before, it is written again', async () => {
  const net = new Net();
  const random = seeded(5);
  const a = await join(net, 'a', await startFile());
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'typed'));
  net.drain(random);
  const pkg = a.file.package!;
  const turn = packageTurn(a.file);
  const packed = pkg.save.bind(pkg);
  let packs = 0;
  let during: (() => void) | undefined;
  // The package is packed right after the workbook was read: what is typed here is typed after the reading.
  pkg.save = () => {
    packs++;
    const run = during;
    during = undefined;
    if (run) run();
    return packed();
  };
  const stamp = a.session.stampNow;
  assert.ok(stamp, 'the workbook is at a step of the room');
  during = () => void act(a, () => setCellInput(a.wb, 0, 13, 0, 'typed after the reading'));
  const late = await writeSharedFile(a.session, pkg, () => turn.save(), turn.locked);
  assert.equal(packs, 1, 'written once');
  assert.equal(late.rev, stamp.rev, 'the file is of the step the workbook was at when it was read');
  const reopened = await open(late.bytes);
  assert.deepEqual(await readCollabStamp(reopened.package!), stamp, 'and says so');
  assert.equal((await reopened.workbook()).sheets[0].rows.get(13), undefined, 'what was typed after the reading is not in it');
  net.drain(random);
  // Typed before the workbook is read (while the stamp is put into the package): the file would hold what its step does not, and is written again.
  packs = 0;
  const stampNow = a.session.stampNow;
  assert.ok(stampNow);
  const early = writeSharedFile(a.session, pkg, () => turn.save(), turn.locked);
  await Promise.resolve();
  act(a, () => setCellInput(a.wb, 0, 14, 0, 'typed before the reading'));
  net.drain(random);
  const done = await early;
  const got = await open(done.bytes);
  const typedIn = (await got.workbook()).sheets[0].rows.get(14) !== undefined;
  const said = await readCollabStamp(got.package!);
  // Whatever the file says, it holds: a step it names is one the typed cell belongs to.
  if (said && typedIn) assert.ok(said.rev > stampNow.rev, 'a file that holds what was typed names a step after it');
  if (said && !typedIn) assert.equal(said.rev, stampNow.rev);
  assert.ok(packs >= 1);
});
