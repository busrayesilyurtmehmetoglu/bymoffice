// Comments of a sheet: the legacy part, the boxes of the legacy drawing,
// threaded comments, and how they follow structural edits. A cross-check
// against openpyxl (an independent implementation) runs where it is installed.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { applyOp, open, type Workbook } from '../src/index';
import { parseVmlNotes } from '../src/formats/xlsx/comments';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { cellName, type SheetStructuralEdit } from '../src/model/sheet/types';
import { SheetDrawings, adjustCommentCell } from '../src/render/sheet-drawings';
import { PX, comment, commentsXml, drawingWorkbook, vmlNote, vmlXml } from './sheet-drawing-fixtures';

async function read(spec: Parameters<typeof drawingWorkbook>[0]): Promise<Workbook> {
  return (await open(await drawingWorkbook(spec))).workbook();
}

const NS_TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

test('comments: text, author and rich runs', async () => {
  const workbook = await read({
    comments: commentsXml(
      ['Ayşe', 'Bora &amp; Can'],
      comment('B2', 0, '<t>Düz metin</t>') +
        comment('C5', 1, '<r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Bora &amp; Can:</t></r><r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\nikinci satır_x000D_</t></r>') +
        comment('D1', 7, '<r><t>tek</t></r><rPh sb="0" eb="1"><t>okunuş</t></rPh>'),
    ),
  });
  const comments = workbook.sheets[0].comments;
  assert.ok(comments);
  assert.deepEqual(comments[0], { row: 1, col: 1, text: 'Düz metin', author: 'Ayşe' });
  assert.equal(comments[1].text, 'Bora & Can:\nikinci satır\r');
  assert.equal(comments[1].author, 'Bora & Can');
  assert.equal(comments[1].runs?.length, 2);
  assert.equal(comments[1].runs?.[0].font?.bold, true);
  assert.equal(comments[1].runs?.[0].font?.name, 'Tahoma');
  assert.equal(comments[1].runs?.[1].font?.bold, false);
  // An author index outside the list, a single run and phonetic text: plain text, no author.
  assert.deepEqual(comments[2], { row: 0, col: 3, text: 'tek' });
  assert.equal(workbook.sheets[0].warnings, undefined);
});

test('comments: a reference that is not a cell and a cell stated twice are left out with a warning', async () => {
  const workbook = await read({
    comments: commentsXml(['a'], comment('A1', 0, '<t>ilk</t>') + comment('A1', 0, '<t>ikinci</t>') + comment('ZZZZ1', 0, '<t>x</t>') + comment('A0', 0, '<t>x</t>') + '<comment authorId="0"><text><t>x</t></text></comment>'),
  });
  assert.deepEqual(workbook.sheets[0].comments, [{ row: 0, col: 0, text: 'ilk', author: 'a' }]);
  assert.deepEqual(workbook.sheets[0].warnings, ['Comments: 4 comment(s) could not be read and are not shown']);
});

test('comments: a comments part that is not XML leaves the sheet readable', async () => {
  const workbook = await read({ sheet: '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>', comments: '<comments><commentList><comment' });
  assert.equal(workbook.sheets[0].comments, undefined);
  assert.equal(workbook.sheets[0].rows.get(0)?.cells[0].value, 7);
  assert.match(workbook.sheets[0].warnings?.[0] ?? '', /^Comments could not be read: /);
});

test('comments: the legacy drawing says which boxes are always shown, and where they are', async () => {
  const workbook = await read({
    comments: commentsXml(['a'], comment('B2', 0, '<t>gizli</t>') + comment('C4', 0, '<t>açık</t>') + comment('E9', 0, '<t>kutusuz</t>')),
    vml: vmlXml(vmlNote(1, 1) + vmlNote(3, 2, { visible: true, anchor: ' 3, 15, 2, 4, 5, 20, 6, 8' })),
  });
  const comments = workbook.sheets[0].comments;
  assert.ok(comments);
  assert.equal(comments[0].visible, undefined);
  assert.deepEqual(comments[0].box, { kind: 'twoCell', from: { col: 2, colOff: 15 * PX, row: 0, rowOff: 2 * PX }, to: { col: 4, colOff: 15 * PX, row: 4, rowOff: 16 * PX } });
  assert.equal(comments[1].visible, true);
  assert.deepEqual(comments[1].box, { kind: 'twoCell', from: { col: 3, colOff: 15 * PX, row: 2, rowOff: 4 * PX }, to: { col: 5, colOff: 20 * PX, row: 6, rowOff: 8 * PX } });
  assert.equal(comments[2].visible, undefined);
  assert.equal(comments[2].box, undefined);
});

test('comments: a legacy drawing that is not well-formed is read as text', () => {
  // An unquoted attribute and an unclosed <br>: what VML written by old tools looks like.
  const text =
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shapetype id=t202></v:shapetype>' +
    '<v:shape id="s1" style=\'visibility:visible\'><v:textbox><div>a<br></div></v:textbox><x:ClientData ObjectType="Note"><x:Anchor>1, 2, 3, 4, 5, 6, 7, 8</x:Anchor><x:Row>3</x:Row><x:Column>0</x:Column><x:Visible/></x:ClientData></v:shape>' +
    '<v:shape id="s2"><x:ClientData ObjectType="Drop"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
    '<v:shape id="s3"><x:ClientData ObjectType="Note"><x:Row>9</x:Row><x:Column>2</x:Column><x:Visible>False</x:Visible></x:ClientData></v:shape></xml>';
  const notes = parseVmlNotes(new TextEncoder().encode(text));
  assert.deepEqual(notes, [
    { row: 3, col: 0, visible: true, box: { kind: 'twoCell', from: { col: 1, colOff: 2 * PX, row: 3, rowOff: 4 * PX }, to: { col: 5, colOff: 6 * PX, row: 7, rowOff: 8 * PX } } },
    { row: 9, col: 2, visible: false },
  ]);
});

test('comments: a box that names a cell outside the sheet, or an anchor that is not eight numbers, is ignored', () => {
  const notes = parseVmlNotes(new TextEncoder().encode(vmlXml(vmlNote(2000000, 1) + vmlNote(1, 1, { anchor: '1, 2, 3' }) + vmlNote(2, 2, { anchor: '1,2,3,4,99999,6,7,8' }))));
  assert.deepEqual(notes, [
    { row: 1, col: 1, visible: false },
    { row: 2, col: 2, visible: false },
  ]);
});

test('comments: threaded comments take the place of the notes written for older applications', async () => {
  const workbook = await read({
    comments: commentsXml(
      ['tc={AAAA}', 'Eski'],
      comment('B2', 0, '<t>[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed.\n\nComment:\n    İlk ileti</t>') +
        comment('D4', 1, '<t>eski not</t>'),
    ),
    vml: vmlXml(vmlNote(1, 1) + vmlNote(3, 3)),
    threaded:
      `<ThreadedComments xmlns="${NS_TC}">` +
      '<threadedComment ref="B2" dT="2024-01-01T10:00:00.00" personId="{P1}" id="{AAAA}"><text>İlk ileti</text></threadedComment>' +
      '<threadedComment ref="B2" dT="2024-01-01T11:00:00.00" personId="{P2}" id="{BBBB}" parentId="{AAAA}"><text>Yanıt</text></threadedComment>' +
      '<threadedComment ref="B2" dT="2024-01-01T12:00:00.00" personId="{P9}" id="{CCCC}" parentId="{AAAA}"><text>Kimden?</text></threadedComment>' +
      '<threadedComment ref="F6" dT="2024-01-01T12:00:00.00" personId="{P2}" id="{DDDD}"><text>Notu olmayan</text></threadedComment>' +
      '<threadedComment ref="B2" personId="{P2}" id="{EEEE}" parentId="{ZZZZ}"><text>Sahipsiz yanıt</text></threadedComment>' +
      '</ThreadedComments>',
    persons: `<personList xmlns="${NS_TC}"><person displayName="Ayşe Yılmaz" id="{P1}" userId="x" providerId="None"/><person displayName="Bora" id="{P2}"/></personList>`,
  });
  const comments = workbook.sheets[0].comments;
  assert.ok(comments);
  assert.equal(comments.length, 3);
  assert.equal(comments[0].threaded, true);
  assert.equal(comments[0].author, 'Ayşe Yılmaz');
  assert.equal(comments[0].text, 'İlk ileti');
  assert.deepEqual(comments[0].replies, [{ author: 'Bora', text: 'Yanıt' }, { text: 'Kimden?' }]);
  // The place of the box is kept from the note it replaces.
  assert.ok(comments[0].box);
  assert.deepEqual(comments[1], { row: 3, col: 3, text: 'eski not', author: 'Eski', box: comments[1].box });
  assert.deepEqual(comments[2], { row: 5, col: 5, text: 'Notu olmayan', threaded: true, author: 'Bora' });
});

test('comments: without the threaded part the fallback text of the note is what is read', async () => {
  const workbook = await read({ comments: commentsXml(['tc={AAAA}'], comment('B2', 0, '<t>[Threaded comment]\n\nComment:\n    İlk ileti</t>')) });
  assert.equal(workbook.sheets[0].comments?.[0].text, '[Threaded comment]\n\nComment:\n    İlk ileti');
  assert.equal(workbook.sheets[0].comments?.[0].threaded, undefined);
});

// --- structural edits ----------------------------------------------------------------------

const EDIT_CASES: [name: string, edits: SheetStructuralEdit[]][] = [
  ['rows inserted', [{ kind: 'insert', axis: 'row', at: 2, count: 3 }]],
  ['columns inserted', [{ kind: 'insert', axis: 'col', at: 1, count: 2 }]],
  ['the row of a comment deleted', [{ kind: 'delete', axis: 'row', at: 3, count: 1 }]],
  ['rows deleted through a box', [{ kind: 'delete', axis: 'row', at: 0, count: 2 }]],
  ['columns deleted', [{ kind: 'delete', axis: 'col', at: 0, count: 1 }]],
  [
    'several edits',
    [
      { kind: 'insert', axis: 'row', at: 0, count: 2 },
      { kind: 'delete', axis: 'col', at: 2, count: 1 },
      { kind: 'insert', axis: 'col', at: 0, count: 1 },
      { kind: 'delete', axis: 'row', at: 8, count: 4 },
    ],
  ],
];

const placed = (workbook: Workbook) => {
  const sheet = workbook.sheets[0];
  const layer = new SheetDrawings(sheet);
  return {
    cells: layer.currentComments().map((item) => `${cellName(item.row, item.col)}=${item.comment.text}`),
    boxes: layer.rects(createSheetMetrics(sheet)).map(({ x, y, width, height }) => ({ x, y, width, height })),
  };
};

for (const [name, edits] of EDIT_CASES) {
  test(`comments: after ${name} the view shows comments and their boxes where the saved file has them`, async () => {
    const file = await open(
      await drawingWorkbook({
        sheet: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
        comments: commentsXml(['a'], comment('B2', 0, '<t>bir</t>') + comment('C4', 0, '<t>iki</t>') + comment('E9', 0, '<t>üç</t>') + comment('A1', 0, '<t>dört</t>')),
        vml: vmlXml(vmlNote(1, 1, { visible: true }) + vmlNote(3, 2, { visible: true, anchor: '3, 15, 2, 4, 5, 20, 6, 8' }) + vmlNote(8, 4, { visible: true }) + vmlNote(0, 0)),
      }),
    );
    const workbook = await file.workbook();
    for (const edit of edits) applyOp(workbook, { type: edit.kind, sheet: 0, axis: edit.axis, at: edit.at, count: edit.count });
    const shown = placed(workbook);
    const reread = await (await open(await file.save())).workbook();
    assert.deepEqual(shown, placed(reread));
  });
}

test('comments: a comment whose cell is deleted is gone; undo brings it back', async () => {
  const file = await open(await drawingWorkbook({ comments: commentsXml(['a'], comment('B2', 0, '<t>bir</t>') + comment('B5', 0, '<t>iki</t>')) }));
  const workbook = await file.workbook();
  const layer = new SheetDrawings(workbook.sheets[0]);
  assert.equal(layer.has(1, 1), true);
  assert.equal(layer.has(4, 1), true);
  assert.equal(layer.has(0, 0), false);
  const inverse = applyOp(workbook, { type: 'delete', sheet: 0, axis: 'row', at: 1, count: 1 });
  assert.equal(layer.has(1, 1), false);
  assert.equal(layer.commentAt(3, 1)?.text, 'iki');
  assert.equal(layer.currentComments().length, 1);
  applyOp(workbook, inverse);
  assert.equal(layer.commentAt(1, 1)?.text, 'bir');
  assert.equal(layer.commentAt(4, 1)?.text, 'iki');
});

test('comments: where a single cell goes', () => {
  assert.deepEqual(adjustCommentCell(4, 2, [{ kind: 'insert', axis: 'row', at: 4, count: 2 }]), { row: 6, col: 2 });
  assert.deepEqual(adjustCommentCell(4, 2, [{ kind: 'insert', axis: 'row', at: 5, count: 2 }]), { row: 4, col: 2 });
  assert.deepEqual(adjustCommentCell(4, 2, [{ kind: 'delete', axis: 'col', at: 0, count: 2 }]), { row: 4, col: 0 });
  assert.equal(adjustCommentCell(4, 2, [{ kind: 'delete', axis: 'col', at: 2, count: 1 }]), undefined);
  assert.deepEqual(adjustCommentCell(4, 2, undefined), { row: 4, col: 2 });
});

// --- openpyxl ------------------------------------------------------------------------------

const here = dirname(fileURLToPath(import.meta.url));
const scratch = join(here, '..', 'tmp-comments-openpyxl');

const hasOpenpyxl = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' }).status === 0;

const WRITE_SCRIPT = `
import sys, json
from openpyxl import Workbook
from openpyxl.comments import Comment
wb = Workbook()
ws = wb.active
ws.title = 'Notlar'
ws['A1'] = 1
ws['B2'].comment = Comment('Birinci not', 'Ayşe')
ws['D7'].comment = Comment('İkinci not\\niki satır', 'Bora', height=120, width=260)
ws['AA100'].comment = Comment('Uzak & <özel> "karakter"', 'Ayşe')
ws2 = wb.create_sheet('Boş')
ws3 = wb.create_sheet('Diğer')
ws3['C3'].comment = Comment('Başka sayfa', 'Can')
wb.save(sys.argv[1])
`;

const READ_SCRIPT = `
import sys, json
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
out = []
for ws in wb.worksheets:
    notes = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.comment is not None:
                notes.append([cell.coordinate, cell.comment.author, cell.comment.text])
    out.append(notes)
print(json.dumps(out))
`;

test('comments: what openpyxl writes is read as openpyxl reads it', { skip: hasOpenpyxl ? false : 'python3 with openpyxl is not installed' }, async () => {
  mkdirSync(scratch, { recursive: true });
  try {
    const path = join(scratch, 'notes.xlsx');
    const written = spawnSync('python3', ['-c', WRITE_SCRIPT, path], { encoding: 'utf8' });
    assert.equal(written.status, 0, written.stderr);
    const read = spawnSync('python3', ['-c', READ_SCRIPT, path], { encoding: 'utf8' });
    assert.equal(read.status, 0, read.stderr);
    const expected = JSON.parse(read.stdout) as [string, string, string][][];
    const workbook = await (await open(new Uint8Array(readFileSync(path)))).workbook();
    const ours = workbook.sheets.map((sheet) => (sheet.comments || []).map((item) => [cellName(item.row, item.col), item.author, item.text]));
    const sort = (list: (string | undefined)[][]) => list.slice().sort((a, b) => String(a[0]).localeCompare(String(b[0])));
    assert.deepEqual(ours.map(sort), expected.map(sort));
    assert.equal(expected[0].length, 3);
    assert.equal(expected[2].length, 1);
    // openpyxl writes a legacy drawing with a box per comment, placed by a style and not by cells; none is always shown.
    for (const sheet of workbook.sheets) for (const item of sheet.comments || []) assert.deepEqual([item.visible, item.box], [undefined, undefined]);
    for (const sheet of workbook.sheets) assert.equal(sheet.warnings, undefined);
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
});
