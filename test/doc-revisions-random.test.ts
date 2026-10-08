// Accepting and rejecting on random documents (seeded): the text that is left
// is what a plain count over the generator's own description says, one by one
// in any order gives what "all" gives, every file written on the way is well
// formed with ids that are unique, and undoing everything gives the file
// back byte for byte.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import {
  DocHistory,
  acceptAll,
  acceptRevision,
  deleteRange,
  insertText,
  listRevisions,
  rejectAll,
  rejectRevision,
  revisionsOf,
  splitParagraph,
  storyIds,
  type DocEdit,
  type DocRange,
  type DocRevision,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { STYLES, checkInvariants, dice, docPicture, paragraphAt, paragraphPaths, run, texts } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

type Dice = ReturnType<typeof dice>;
type Verdicts = (author: string) => boolean;

const decoder = new TextDecoder();
const AUTHORS = ['Ann', 'Bob', 'Cid', 'Dan'];
const DATES = ['2026-10-01T09:00:00Z', '2026-10-02T09:00:00Z'];
const WORDS = ['alpha', 'beta', 'gamma', 'delta', 'epsilon', 'zeta', 'eta', 'theta'];

/** What the generator knows of a piece of text, a mark or a row: who inserted it and who deleted it. */
interface State {
  ins?: string;
  del?: string;
}

/** Does it stay, given who is accepted (true) and who is rejected. */
const stays = (state: State, accepts: Verdicts): boolean => (state.ins === undefined || accepts(state.ins)) && (state.del === undefined || !accepts(state.del));

interface Made {
  spec: DocxSpec;
  /** The text each story is left with, for a set of verdicts; a note whose reference goes is gone. */
  expected(accepts: Verdicts): Map<string, string>;
  revisions: number;
}

export function generate(seed: number): Made {
  const d: Dice = dice(seed);
  let id = 100;
  let serial = 0;
  let revisions = 0;
  const who = (author: string): string => ` w:id="${id++}" w:author="${author}" w:date="${d.pick(DATES)}"`;
  const word = (): string => d.pick(WORDS) + serial++ + ' ';
  const run = (text: string, rPr = ''): string => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const gone = (text: string): string => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;

  type Piece = { text: string; state: State } | { mark: State } | { row: State; open: boolean };
  const notes: { id: number; state: State; pieces: Piece[]; xml: string }[] = [];

  /** The content of one paragraph and what the generator knows of it. */
  const content = (pieces: Piece[], allowNotes: boolean): string => {
    let xml = '';
    const count = d.int(5);
    for (let i = 0; i < count; i++) {
      const author = d.pick(AUTHORS);
      const other = d.pick(AUTHORS);
      const text = word();
      const kind = d.int(allowNotes ? 12 : 10);
      revisions++;
      if (kind === 0 || kind === 1) {
        revisions--;
        xml += run(text);
        pieces.push({ text, state: {} });
      } else if (kind === 2 || kind === 3) {
        xml += `<w:ins${who(author)}>${run(text)}</w:ins>`;
        pieces.push({ text, state: { ins: author } });
      } else if (kind === 4 || kind === 5) {
        // Now and then with a bookmark inside the deletion: it stays when the text goes.
        const mark = kind === 5 && d.chance(0.5) ? 5000 + serial : 0;
        xml += mark ? `<w:del${who(author)}><w:bookmarkStart w:id="${mark}" w:name="b${mark}"/>${gone(text)}<w:bookmarkEnd w:id="${mark}"/></w:del>` : `<w:del${who(author)}>${gone(text)}</w:del>`;
        pieces.push({ text, state: { del: author } });
      } else if (kind === 6) {
        // A deletion of one author inside the insertion of another, with text of the insertion around it.
        const more = word();
        revisions++;
        xml += `<w:ins${who(author)}>${run(more)}<w:del${who(other)}>${gone(text)}</w:del></w:ins>`;
        pieces.push({ text: more, state: { ins: author } }, { text, state: { ins: author, del: other } });
      } else if (kind === 7) {
        // An insertion inside a link that also holds text of its own.
        const more = word();
        xml += `<w:hyperlink w:anchor="here">${run(more)}<w:ins${who(author)}>${run(text)}</w:ins></w:hyperlink>`;
        pieces.push({ text: more, state: {} }, { text, state: { ins: author } });
      } else if (kind === 8) {
        // A whole field that was inserted, or deleted.
        const inserted = d.chance(0.5);
        const field =
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
          `<w:r><w:${inserted ? 'instrText' : 'delInstrText'} xml:space="preserve"> QUOTE x </w:${inserted ? 'instrText' : 'delInstrText'}></w:r>` +
          '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
          (inserted ? run(text) : gone(text)) +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
        xml += inserted ? `<w:ins${who(author)}>${field}</w:ins>` : `<w:del${who(author)}>${field}</w:del>`;
        pieces.push({ text, state: inserted ? { ins: author } : { del: author } });
      } else if (kind === 9) {
        // Formatting that was changed: the text stays either way.
        xml += run(text, `<w:b/><w:rPrChange${who(author)}><w:rPr><w:i/></w:rPr></w:rPrChange>`);
        pieces.push({ text, state: {} });
      } else {
        // A note whose reference was inserted or deleted, with a change of its own inside.
        const noteId = notes.length + 2;
        const state: State = kind === 10 ? { ins: author } : { del: author };
        const ref = `<w:r><w:footnoteReference w:id="${noteId}"/></w:r>`;
        xml += kind === 10 ? `<w:ins${who(author)}>${ref}</w:ins>` : `<w:del${who(author)}>${ref}</w:del>`;
        const inside: Piece[] = [];
        const note = { id: noteId, state, pieces: inside, xml: '' };
        notes.push(note);
        const body = content(inside, false);
        inside.push({ mark: {} });
        note.xml = `<w:footnote w:id="${noteId}"><w:p><w:r><w:footnoteRef/></w:r>${body}</w:p></w:footnote>`;
      }
    }
    return xml;
  };

  const markOf = (plain: boolean): { xml: string; state: State } => {
    if (plain || d.chance(0.55)) return { xml: '', state: {} };
    const author = d.pick(AUTHORS);
    revisions++;
    return d.chance(0.5) ? { xml: `<w:rPr><w:ins${who(author)}/></w:rPr>`, state: { ins: author } } : { xml: `<w:rPr><w:del${who(author)}/></w:rPr>`, state: { del: author } };
  };

  const paragraph = (pieces: Piece[], plainMark: boolean, allowNotes: boolean): string => {
    const inner: Piece[] = [];
    const xml = content(inner, allowNotes);
    const mark = markOf(plainMark);
    for (const piece of inner) pieces.push(piece);
    pieces.push({ mark: mark.state });
    // Paragraphs are formatted differently, so that which of two joined paragraphs gives its formatting shows.
    const jc = d.chance(0.7) ? `<w:jc w:val="${d.pick(['center', 'right', 'both', 'left'])}"/>` : '';
    return `<w:p>${jc || mark.xml ? `<w:pPr>${jc}${mark.xml}</w:pPr>` : ''}${xml}</w:p>`;
  };

  const table = (pieces: Piece[]): string => {
    let rows = '';
    const count = 1 + d.int(3);
    for (let r = 0; r < count; r++) {
      let state: State = {};
      let trPr = '';
      if (r > 0 && d.chance(0.6)) {
        const author = d.pick(AUTHORS);
        revisions++;
        state = d.chance(0.5) ? { ins: author } : { del: author };
        trPr = `<w:trPr><w:${state.ins ? 'ins' : 'del'}${who(author)}/></w:trPr>`;
      }
      pieces.push({ row: state, open: true });
      let cells = '';
      for (let c = 0; c < 2; c++) {
        let blocks = '';
        const paragraphs = 1 + d.int(2);
        for (let k = 0; k < paragraphs; k++) blocks += paragraph(pieces, k === paragraphs - 1, false);
        cells += `<w:tc>${blocks}</w:tc>`;
      }
      pieces.push({ row: state, open: false });
      rows += `<w:tr>${trPr}${cells}</w:tr>`;
    }
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;
  };

  const story = (pieces: Piece[], allowNotes: boolean, tables: boolean): string => {
    let xml = '';
    const count = 2 + d.int(5);
    for (let i = 0; i < count; i++) {
      if (tables && d.chance(0.2)) {
        // A paragraph with a plain mark on either side of a table.
        xml += paragraph(pieces, true, allowNotes) + table(pieces);
        xml += paragraph(pieces, d.chance(0.5), allowNotes);
      } else xml += paragraph(pieces, false, allowNotes);
    }
    pieces.push({ text: 'end', state: {} }, { mark: {} });
    return xml + '<w:p><w:r><w:t>end</w:t></w:r></w:p>';
  };

  const bodyPieces: Piece[] = [];
  const headerPieces: Piece[] = [];
  const body = story(bodyPieces, true, true);
  const header = story(headerPieces, false, false);

  const textOf = (pieces: Piece[], accepts: Verdicts): string => {
    let out = '';
    let hidden = false;
    for (const piece of pieces) {
      if ('row' in piece) hidden = piece.open && !stays(piece.row, accepts);
      else if (hidden) continue;
      else if ('mark' in piece) out += stays(piece.mark, accepts) ? '\n' : '';
      else if (stays(piece.state, accepts)) out += piece.text;
    }
    return out.replace(/\n$/, '');
  };

  const spec: DocxSpec = {
    body: body + '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>',
    styles: STYLES,
    headers: { rId7: header },
    footnotes:
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      notes.map((note) => note.xml).join(''),
  };
  return {
    spec,
    revisions,
    expected: (accepts) => {
      const out = new Map<string, string>([
        ['body', textOf(bodyPieces, accepts)],
        ['header:rId7', textOf(headerPieces, accepts)],
      ]);
      // A note is there as long as its reference is, and the reference is in the body (not in a row here).
      for (const note of notes) if (stays(note.state, accepts)) out.set('footnote:' + note.id, textOf(note.pieces, accepts));
      return out;
    },
  };
}

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(spec: DocxSpec): Promise<Opened> {
  const bytes = await buildSpec(spec);
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

const PARTS = ['/word/document.xml', '/word/header1.xml', '/word/footnotes.xml'];
const REVISION = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b([^>]*)>/g;

/** Writes the document and checks what a reader of the file would: well formed, in schema order, changes with id and author, ids unique, no plain text in a deletion. */
async function checkWritten(o: Opened, label: string): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  const ids: string[] = [];
  let all = '';
  for (const part of PARTS) {
    const text = decoder.decode(await o.pkg.readPart(part));
    parseXml(text);
    assert.deepEqual(schemaOrderViolations(text), [], `${label}: ${part}`);
    for (const match of text.matchAll(REVISION)) {
      const id = /\sw:id="(\d+)"/.exec(match[2]);
      assert.ok(id && /\sw:author="[^"]+"/.test(match[2]), `${label}: ${match[0]}`);
      ids.push(id[1]);
    }
    for (const deletion of text.matchAll(/<w:del [^>]*[^/]>(.*?)<\/w:del>/g)) assert.ok(!/<w:t[ >]|<w:instrText/.test(deletion[1]), `${label}: text in a deletion: ${deletion[0]}`);
    all += text;
  }
  assert.equal(new Set(ids).size, ids.length, `${label}: an id twice`);
  return all;
}

function storyTexts(doc: DocModel): Map<string, string> {
  const out = new Map<string, string>();
  for (const id of storyIds(doc)) out.set(id, texts(doc, id).join('\n'));
  return out;
}

const sorted = (map: Map<string, string>): [string, string][] => Array.from(map.entries()).sort((a, b) => (a[0] < b[0] ? -1 : 1));

const SEEDS = [1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 233];

test('random revisions: "accept all" and "reject all" leave the text a plain count says, and no change', async () => {
  let total = 0;
  for (const seed of SEEDS) {
    const made = generate(seed);
    total += made.revisions;
    for (const accept of [true, false]) {
      const o = await open(made.spec);
      const label = `seed ${seed} ${accept ? 'accept' : 'reject'}`;
      assert.ok(revisionsOf(o.doc).length > 0, label);
      const edit = accept ? acceptAll(o.doc) : rejectAll(o.doc);
      assert.equal(edit.skipped, 0, label);
      run(o.history, edit);
      checkInvariants(o.doc);
      assert.deepEqual(sorted(storyTexts(o.doc)), sorted(made.expected(() => accept)), label);
      assert.deepEqual(revisionsOf(o.doc), [], label);
      const xml = await checkWritten(o, label);
      assert.ok(!/<w:(ins|del|moveFrom|moveTo|rPrChange|delText|delInstrText)\b/.test(xml), label);
      // The file opens again to the same text.
      const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
      assert.deepEqual(sorted(storyTexts(again)), sorted(made.expected(() => accept)), label);
      o.history.undo();
      await writeDocx(o.pkg, o.doc);
      assert.deepEqual(await o.pkg.save(), o.bytes, `${label}: undo`);
    }
  }
  // The documents are not empty of what is tested.
  assert.ok(total > SEEDS.length * 8, String(total));
});

test('random revisions: one by one in a random order gives what "all" gives for the same verdicts', async () => {
  for (const seed of SEEDS) {
    const made = generate(seed);
    const d = dice(seed * 7 + 1);
    // Two authors are accepted, two rejected.
    const accepted = new Set<string>([AUTHORS[d.int(2)], AUTHORS[2 + d.int(2)]]);
    const accepts: Verdicts = (author) => accepted.has(author);
    const byAuthor = (revision: DocRevision): boolean => accepts(revision.author as string);

    const all = await open(made.spec);
    run(all.history, acceptAll(all.doc, byAuthor));
    run(all.history, rejectAll(all.doc, (revision) => !byAuthor(revision)));
    assert.deepEqual(revisionsOf(all.doc), [], `seed ${seed}`);
    assert.deepEqual(sorted(storyTexts(all.doc)), sorted(made.expected(accepts)), `seed ${seed}`);

    const one = await open(made.spec);
    let steps = 0;
    for (let list = revisionsOf(one.doc); list.length > 0; list = revisionsOf(one.doc)) {
      const revision = d.pick(list);
      run(one.history, byAuthor(revision) ? acceptRevision(one.doc, revision) : rejectRevision(one.doc, revision));
      checkInvariants(one.doc);
      if (++steps % 5 === 0) await checkWritten(one, `seed ${seed} step ${steps}`);
      assert.ok(steps < 2000, 'the list of changes does not get shorter');
    }
    assert.deepEqual(sorted(storyTexts(one.doc)), sorted(made.expected(accepts)), `seed ${seed}`);
    assert.equal(await checkWritten(one, `seed ${seed} one`), await checkWritten(all, `seed ${seed} all`), `seed ${seed}: the files differ`);
    assert.equal(docPicture(one.doc), docPicture(all.doc), `seed ${seed}: the models differ`);
    // Every step taken back: the file that was read.
    while (one.history.canUndo) one.history.undo();
    await writeDocx(one.pkg, one.doc);
    assert.deepEqual(await one.pkg.save(), one.bytes, `seed ${seed}: undo`);
  }
});

test('random revisions: editing with tracking on, mixed with accepting and rejecting', async () => {
  for (const seed of SEEDS.slice(0, 8)) {
    const made = generate(seed);
    const o = await open({ ...made.spec, settings: '<w:trackRevisions/>' });
    const d = dice(seed * 13 + 5);
    const me = { author: 'Me', date: '2026-10-08T10:00:00Z' };
    let done = 0;
    let resolved = 0;
    for (let step = 0; step < 60; step++) {
      const paths = paragraphPaths(o.doc);
      const path = d.pick(paths);
      const length = paragraphAt(o.doc, path).children.length ? texts(o.doc)[paths.indexOf(path)].length : 0;
      const at = { story: 'body', path, offset: d.int(length + 1) };
      let edit: DocEdit | undefined;
      try {
        const kind = d.int(6);
        if (kind === 0) edit = insertText(o.doc, at, d.pick(WORDS), undefined, me);
        else if (kind === 1) edit = deleteRange(o.doc, { anchor: at, focus: { story: 'body', path, offset: Math.min(length, at.offset + 1 + d.int(4)) } }, me);
        else if (kind === 2) edit = splitParagraph(o.doc, at, me);
        else {
          const list = listRevisions(o.doc).list;
          if (!list.length) continue;
          const revision = d.pick(list);
          edit = kind % 2 ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision);
          resolved++;
        }
      } catch (error) {
        // A place that cannot be edited, a change that cannot be tracked or resolved: nothing happened.
        if (!(error instanceof RangeError)) throw error;
        continue;
      }
      if (!edit.ops.length) continue;
      run(o.history, edit);
      done++;
      checkInvariants(o.doc);
      await checkWritten(o, `seed ${seed} step ${step}`);
    }
    assert.ok(done > 20 && resolved > 5, `seed ${seed}: ${done} edits, ${resolved} resolved`);
    // What is left can all be accepted.
    const rest = acceptAll(o.doc);
    assert.equal(rest.skipped, 0, `seed ${seed}`);
    run(o.history, rest);
    assert.deepEqual(revisionsOf(o.doc), [], `seed ${seed}`);
    const xml = await checkWritten(o, `seed ${seed} end`);
    assert.ok(!/<w:(ins|del|rPrChange|delText)\b/.test(xml), `seed ${seed}`);
    while (o.history.canUndo) o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, `seed ${seed}: undo`);
  }
});
