// Helpers for the document editing tests: small synthetic documents, a
// picture of the whole model for comparisons, the invariants every edit has
// to keep, and a generator of random documents and random commands.

import assert from 'node:assert/strict';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutFont, TextMeasurer } from '../src/layout/doc/types';
import { paragraphText, type Block, type DocModel, type Paragraph } from '../src/model/doc/types';
import {
  DocHistory,
  adjacentParagraph,
  containerBlocks,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  indentList,
  insertColumn,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertRow,
  insertTable,
  insertText,
  paragraphLength,
  resolvePosition,
  setParaFormat,
  setRunFormat,
  splitParagraph,
  storyIds,
  storyOf,
  storyStart,
  toggleList,
  toggleRunFormat,
  type DocEdit,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { flatten } from '../src/ops/doc/pieces';
import { readSpec, type DocxSpec } from './docx-fixtures';

export const measurer: TextMeasurer = {
  width: (font: LayoutFont, text: string) => text.length * font.size * 0.5,
  metrics: (font: LayoutFont) => ({ ascent: font.size * 0.9, descent: font.size * 0.22, lineGap: 0 }),
};

export const STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Test" w:hAnsi="Test" w:eastAsia="Test" w:cs="Test"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Heading1"/><w:rPr><w:i/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListStyle"><w:name w:val="List Style"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>';

export const NUMBERING =
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="+"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';

export const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
export const p = (content = '', pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.charAt(0) === '<' ? content : content ? r(content) : ''}</w:p>`;
export const tc = (content: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${content.charAt(0) === '<' ? content : p(content)}</w:tc>`;
export const tr = (cells: string) => `<w:tr>${cells}</w:tr>`;
export const tbl = (rows: string, grid: number[]) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;

/** Reads a document with the standard styles, numbering and section; `spec` is the body XML or a full spec. */
export async function open(spec: DocxSpec | string): Promise<DocModel> {
  const full: DocxSpec = typeof spec === 'string' ? { body: spec } : { ...spec };
  if (full.styles === undefined) full.styles = STYLES;
  if (full.numbering === undefined) full.numbering = NUMBERING;
  if (full.body !== undefined && full.body.indexOf('<w:sectPr><w:pgSz') < 0) full.body += SECT;
  return readSpec(full);
}

export const pos = (path: number[] | number, offset: number, story = 'body'): DocPosition => ({ story, path: typeof path === 'number' ? [path] : path, offset });
export const range = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });

/** Paths of all paragraphs of a story, in document order. */
export function paragraphPaths(doc: DocModel, story = 'body'): number[][] {
  const out: number[][] = [];
  const start = storyStart(doc, story);
  for (let path = start && start.path; path; path = adjacentParagraph(doc, story, path, 1)) out.push(path);
  return out;
}

export function paragraphAt(doc: DocModel, path: number[] | number, story = 'body'): Paragraph {
  return resolvePosition(doc, pos(path, 0, story)).paragraph;
}

/** Text of the paragraphs of a story, one per entry. */
export const texts = (doc: DocModel, story = 'body'): string[] => paragraphPaths(doc, story).map((path) => paragraphText(paragraphAt(doc, path, story)));

/** Applies an edit as one step that later typing does not merge into. */
export function run(history: DocHistory, edit: DocEdit, before?: DocRange): DocRange {
  history.applyEdit(edit, before);
  history.seal();
  return edit.selection;
}

// ------------------------------------------------------------------ comparison

/** A text that is equal for two models exactly when they hold the same data. Byte arrays count by length and a checksum. */
export function picture(value: unknown): string {
  const seen: unknown[] = [];
  const walk = (v: unknown): unknown => {
    if (v === null || typeof v !== 'object') return v === undefined ? '\u0000undefined' : v;
    if (v instanceof Uint8Array) {
      let sum = 0;
      for (let i = 0; i < v.length; i++) sum = (sum * 31 + v[i]) >>> 0;
      return { bytes: v.length, sum };
    }
    if (seen.indexOf(v) >= 0) throw new Error('cycle in the model');
    seen.push(v);
    let out: unknown;
    if (v instanceof Map) out = { map: Array.from(v.entries()).map(([key, item]) => [key, walk(item)]) };
    else if (v instanceof Set) out = { set: Array.from(v.values()).map(walk) };
    else if (Array.isArray(v)) out = v.map(walk);
    else {
      const record: Record<string, unknown> = {};
      for (const key of Object.keys(v).sort()) {
        const item = (v as Record<string, unknown>)[key];
        // A missing property and one that is undefined are the same model.
        if (item !== undefined) record[key] = walk(item);
      }
      out = record;
    }
    seen.pop();
    return out;
  };
  return JSON.stringify(walk(value));
}

/** The picture of a document without the source texts of its parts, which no edit touches. */
export function docPicture(doc: DocModel): string {
  const strip = <T extends { source?: string }>(story: T | undefined): T | undefined => (story ? { ...story, source: story.source === undefined ? undefined : story.source.length } : story) as T | undefined;
  const headers = new Map<string, unknown>();
  doc.headers.forEach((story, id) => headers.set(id, strip(story)));
  const footers = new Map<string, unknown>();
  doc.footers.forEach((story, id) => footers.set(id, strip(story)));
  return picture({ ...doc, body: strip(doc.body), headers, footers, footnotes: strip(doc.footnotes), endnotes: strip(doc.endnotes), comments: strip(doc.comments) });
}

/** The objects whose identity undoing has to bring back: the blocks of every story and the notes. */
export function identities(doc: DocModel): Map<string, unknown[]> {
  const out = new Map<string, unknown[]>();
  for (const id of storyIds(doc)) {
    const handle = storyOf(doc, id);
    out.set(id, (handle.note ? [handle.note] : []).concat(handle.blocks as never[]));
  }
  out.set('modified', [doc.body.modified, doc.numbering.modified, doc.footnotes && doc.footnotes.modified, doc.endnotes && doc.endnotes.modified, doc.media]);
  doc.headers.forEach((story, id) => out.set('modified header ' + id, [story.modified, story.relationships]));
  doc.footers.forEach((story, id) => out.set('modified footer ' + id, [story.modified, story.relationships]));
  out.set('relationships', [doc.body.relationships]);
  return out;
}

export function assertSameIdentities(actual: Map<string, unknown[]>, expected: Map<string, unknown[]>): void {
  assert.deepEqual(Array.from(actual.keys()), Array.from(expected.keys()));
  expected.forEach((list, key) => {
    const now = actual.get(key) as unknown[];
    assert.equal(now.length, list.length, `${key}: number of objects`);
    for (let i = 0; i < list.length; i++) assert.ok(now[i] === list[i], `${key}[${i}] is not the object it was`);
  });
}

// ------------------------------------------------------------------ invariants

function checkBlocks(blocks: Block[], where: string, strict: boolean, top: boolean): void {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const at = `${where}/${i}`;
    if (block.type === 'paragraph') {
      let length = 0;
      for (const piece of flatten(block)) length += piece.len;
      assert.equal(length, paragraphText(block).length, `${at}: offsets of the pieces`);
      for (const inline of block.children) {
        if (inline.type !== 'run') continue;
        if (strict) assert.ok(inline.content.length > 0, `${at}: empty run`);
        if (Array.isArray(inline.fieldCode)) assert.equal(inline.fieldCode.length, inline.content.length, `${at}: field code flags`);
        if (strict) for (const item of inline.content) if (item.type === 'text') assert.ok(item.text.length > 0, `${at}: empty text`);
        assert.ok(inline.props !== undefined, `${at}: run without properties`);
      }
    } else if (block.type === 'table') {
      assert.ok(block.rows.length > 0, `${at}: table without rows`);
      let width = -1;
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        assert.ok(row.cells.length > 0, `${at}: row without cells`);
        let span = (row.props.gridBefore || 0) + (row.props.gridAfter || 0);
        for (let c = 0; c < row.cells.length; c++) {
          const cell = row.cells[c];
          span += cell.props.gridSpan && cell.props.gridSpan > 1 ? cell.props.gridSpan : 1;
          if (strict) {
            assert.ok(cell.blocks.length > 0, `${at}: empty cell`);
            assert.equal(cell.blocks[cell.blocks.length - 1].type, 'paragraph', `${at}/${r}/${c}: a cell ends with a paragraph`);
          }
          checkBlocks(cell.blocks, `${at}/${r}/${c}`, strict, false);
        }
        if (strict) {
          if (width < 0) width = span;
          assert.equal(span, width, `${at}: row ${r} covers another number of grid columns`);
          assert.equal(block.grid.length, width, `${at}: grid columns`);
        }
      }
      if (strict && i > 0) assert.notEqual(blocks[i - 1].type, 'table', `${at}: two tables touch`);
    }
    if (top && block.srcStart === undefined) assert.equal(block.srcEnd, undefined);
  }
}

/** What has to hold after every edit. `strict` adds what holds for the generated documents (real ones may break it before any edit). */
export function checkInvariants(doc: DocModel, strict = true): void {
  for (const id of storyIds(doc)) {
    const handle = storyOf(doc, id);
    checkBlocks(handle.blocks, id, strict, true);
    for (const path of paragraphPaths(doc, id)) {
      const at = resolvePosition(doc, { story: id, path, offset: 0 });
      assert.equal(at.length, paragraphText(at.paragraph).length);
    }
    // A block that was replaced makes the story modified.
    if (!handle.note && !handle.host.modified && handle.host.source !== '') {
      for (const block of handle.blocks) {
        let root = block.wrap;
        while (root && root.parent) root = root.parent;
        assert.ok(block.srcStart !== undefined || (root !== undefined && root.srcStart !== undefined), `${id}: a block without a source span in a story that is not modified`);
      }
    }
  }
}

/**
 * Lays the document out and checks that every text item with a source shows the text the model has at that place
 * (`exact`), or at least lies inside its paragraph. Returns the number of items checked.
 */
export function checkLayout(doc: DocModel, exact = true): number {
  const layout = new DocLayout(doc, measurer, { maxPages: 400 });
  layout.layout();
  let checked = 0;
  for (const page of layout.pages) {
    for (const item of page.items) {
      if (item.type !== 'text' || !item.source || item.repeat || (item as { field?: string }).field !== undefined) continue;
      const source = item.source;
      const blocks = containerBlocks(doc, source.story, source.path.slice(0, -1));
      const paragraph = blocks[source.path[source.path.length - 1]];
      assert.equal(paragraph.type, 'paragraph', 'the source of a text item is a paragraph');
      const text = paragraphText(paragraph as Paragraph);
      // Capitals, small capitals and symbol fonts show other characters than the model holds; then only the place is checked.
      if (exact) assert.equal(text.substr(source.offset, item.text.length), item.text, `text at ${source.story} ${source.path.join('.')}+${source.offset}`);
      else if (source.offset < 0 || source.offset + item.text.length > text.length) assert.fail(`a text item lies outside its paragraph at ${source.story} ${source.path.join('.')}+${source.offset}`);
      // Outside text boxes the source is a position the commands accept.
      if (source.path.indexOf(-1) < 0) resolvePosition(doc, source);
      checked++;
    }
  }
  return checked;
}

// ---------------------------------------------------------------------- random

/** A small deterministic generator (mulberry32). */
export function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const WORDS = ['alpha', 'beta', 'gamma', 'delta', 'omega', 'x', 'áb', 'smile😀', 'çığ', '42', 'left-right', 'e.g.'];

export interface Dice {
  next: () => number;
  int: (n: number) => number;
  pick: <T>(list: T[]) => T;
  chance: (p: number) => boolean;
}

export function dice(seed: number): Dice {
  const next = random(seed);
  const int = (n: number) => Math.floor(next() * n);
  return { next, int, pick: (list) => list[int(list.length)], chance: (p) => next() < p };
}

function words(d: Dice, count: number): string {
  const out: string[] = [];
  for (let i = 0; i < count; i++) out.push(d.pick(WORDS));
  return out.join(' ');
}

const DRAWING =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="190500" cy="190500"/><wp:docPr id="7" name="Shape"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>boxed text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';

function randomRuns(d: Dice, notes: number[]): string {
  let out = '';
  const count = d.int(5);
  for (let i = 0; i < count; i++) {
    const kind = d.int(22);
    const rPr = d.pick(['', '', '<w:b/>', '<w:i/>', '<w:sz w:val="32"/>', '<w:color w:val="FF0000"/>', '<w:rStyle w:val="Strong"/>', '<w:u w:val="single"/>', '<w:b/><w:i/>']);
    const text = words(d, 1 + d.int(4)) + (d.chance(0.5) ? ' ' : '');
    if (kind < 9) out += r(text, rPr);
    else if (kind === 9) out += `<w:r><w:t xml:space="preserve">${escape(text)}</w:t><w:tab/><w:t>${escape(words(d, 1))}</w:t></w:r>`;
    else if (kind === 10) out += `<w:r><w:br/><w:t>${escape(words(d, 2))}</w:t></w:r>`;
    else if (kind === 11) out += `<w:hyperlink r:id="rIdLink">${r(text, rPr)}${r(words(d, 1))}</w:hyperlink>`;
    else if (kind === 12) out += `<w:bookmarkStart w:id="${i}" w:name="mark${d.int(1000)}"/>${r(text, rPr)}<w:bookmarkEnd w:id="${i}"/>`;
    else if (kind === 13) out += `<w:fldSimple w:instr=" PAGE ">${r('1', rPr)}</w:fldSimple>`;
    else if (kind === 14) {
      out += `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r('12', rPr)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    } else if (kind === 15) {
      out += `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF mark \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>${r(text, rPr)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    } else if (kind === 16) out += `<w:del w:id="${90 + i}" w:author="A" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins w:id="${190 + i}" w:author="A"><w:r><w:t>came</w:t></w:r></w:ins>`;
    else if (kind === 17) out += `<w:r><w:t>${escape(words(d, 1))}</w:t><w:mystery w:val="1"/></w:r><w:strange/>`;
    else if (kind === 18) out += `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${r(text, rPr)}</w:sdtContent></w:sdt>`;
    else if (kind === 19 && notes.length < 3) {
      notes.push(notes.length + 2);
      out += `${r(text)}<w:r><w:footnoteReference w:id="${notes.length + 1}"/></w:r>`;
    } else if (kind === 20) out += `<w:r>${DRAWING}</w:r>`;
    else out += `<m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>`;
  }
  return out;
}

function randomParagraph(d: Dice, notes: number[]): string {
  const pPr = d.pick(['', '', '', '<w:pStyle w:val="Heading1"/>', '<w:jc w:val="center"/>', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>', '<w:ind w:left="720" w:firstLine="360"/>', '<w:pStyle w:val="ListStyle"/>', '<w:rPr><w:b/></w:rPr>']);
  return p(randomRuns(d, notes), pPr);
}

function randomTable(d: Dice, notes: number[], depth: number): string {
  const columns = 1 + d.int(4);
  const rows = 1 + d.int(4);
  const merge = columns > 1 && d.chance(0.4);
  const vertical = rows > 1 && (!merge || columns > 2) && d.chance(0.4);
  let out = '';
  for (let row = 0; row < rows; row++) {
    let cells = '';
    for (let c = 0; c < columns; c++) {
      let tcPr = '<w:tcW w:w="1500" w:type="dxa"/>';
      if (merge && row === 0 && c === 0) tcPr += '<w:gridSpan w:val="2"/>';
      if (merge && row === 0 && c === 1) continue;
      if (vertical && c === columns - 1 && row < 2) tcPr += row === 0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
      let content = '';
      if (depth < 1 && d.chance(0.1)) content += randomTable(d, notes, depth + 1);
      const paragraphs = 1 + d.int(2);
      for (let k = 0; k < paragraphs; k++) content += randomParagraph(d, notes);
      cells += tc(content, tcPr);
    }
    out += tr(cells);
  }
  const grid: number[] = [];
  for (let c = 0; c < columns; c++) grid.push(1500);
  return tbl(out, grid);
}

/** A random document: paragraphs of all kinds of inline content, tables with merged cells, a header, footnotes. */
export async function randomDoc(seed: number): Promise<DocModel> {
  const d = dice(seed);
  const notes: number[] = [];
  let body = '';
  const count = 3 + d.int(8);
  let lastTable = true;
  for (let i = 0; i < count; i++) {
    const kind = d.int(10);
    if (kind < 2 && !lastTable) {
      body += randomTable(d, notes, 0);
      lastTable = true;
      continue;
    }
    lastTable = false;
    if (kind === 2) body += `<w:sdt><w:sdtPr><w:alias w:val="block"/></w:sdtPr><w:sdtContent>${randomParagraph(d, notes)}${randomParagraph(d, notes)}</w:sdtContent></w:sdt>`;
    else if (kind === 3) body += `<w:bookmarkStart w:id="50${i}" w:name="between${i}"/>${randomParagraph(d, notes)}<w:bookmarkEnd w:id="50${i}"/>`;
    else if (kind === 4) body += p(randomRuns(d, notes), '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/><w:type w:val="continuous"/></w:sectPr>');
    else body += randomParagraph(d, notes);
  }
  body += p(randomRuns(d, notes));
  body += '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="900" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
  let footnotes = '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  for (const id of notes) footnotes += `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r>${r(' note ' + words(d, 3))}</w:p></w:footnote>`;
  return open({
    body,
    footnotes,
    headers: { rIdHeader: p(randomRuns(d, [9, 9, 9])) + p('header end') },
    rels: [['rIdLink', 'hyperlink', 'https://example.com/', true]],
  });
}

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

/** A random position of a story: a random paragraph and a random offset in it. */
export function randomPosition(doc: DocModel, d: Dice, story: string, paths = paragraphPaths(doc, story)): DocPosition {
  const path = d.pick(paths);
  return { story, path, offset: d.int(paragraphLength(paragraphAt(doc, path, story)) + 1) };
}

/** A range of a story with ends not too far apart, so that documents do not shrink to nothing at once. */
function randomRange(doc: DocModel, d: Dice, story: string, paths: number[][]): DocRange {
  const i = d.int(paths.length);
  const j = Math.min(paths.length - 1, i + (d.chance(0.6) ? 0 : d.int(3)));
  const at = (path: number[]): DocPosition => ({ story, path, offset: d.int(paragraphLength(paragraphAt(doc, path, story)) + 1) });
  const a = at(paths[i]);
  const b = at(paths[j]);
  return d.chance(0.5) ? { anchor: a, focus: b } : { anchor: b, focus: a };
}

export const COMMANDS = [
  'insertText',
  'insertLines',
  'typeOver',
  'deleteRange',
  'deleteBackward',
  'deleteForward',
  'splitParagraph',
  'setRunFormat',
  'toggleRunFormat',
  'setParaFormat',
  'toggleList',
  'indentList',
  'insertTable',
  'insertRow',
  'deleteRow',
  'insertColumn',
  'deleteColumn',
  'insertPageBreak',
  'insertHyperlink',
  'insertImage',
];

/** One random command at a random place of a random story. Returns undefined when the command does not apply there (no table at the position). */
export function randomEdit(doc: DocModel, d: Dice, stories = storyIds(doc)): { name: string; edit: DocEdit } | undefined {
  const story = d.pick(stories);
  const paths = paragraphPaths(doc, story).filter((path) => path.indexOf(-1) < 0);
  if (paths.length === 0) return undefined;
  const at = randomPosition(doc, d, story, paths);
  const span = randomRange(doc, d, story, paths);
  const name = d.pick(COMMANDS);
  const inTable = at.path.length >= 4;
  let edit: DocEdit;
  switch (name) {
    case 'insertText':
      edit = insertText(doc, at, d.pick(['a', ' ', 'word', 'two words', '\t', 'é', '😀', 'x\vy']), d.chance(0.2) ? { bold: true } : undefined);
      break;
    case 'insertLines':
      edit = insertText(doc, at, d.pick(['one\ntwo', '\n', 'a\n\nb\tc', 'end\n']));
      break;
    case 'typeOver':
      edit = insertText(doc, span, d.pick(['Z', 'new text', 'p\nq']));
      break;
    case 'deleteRange':
      edit = deleteRange(doc, span);
      break;
    case 'deleteBackward':
      edit = deleteBackward(doc, at, d.chance(0.3) ? 'word' : 'char');
      break;
    case 'deleteForward':
      edit = deleteForward(doc, at, d.chance(0.3) ? 'word' : 'char');
      break;
    case 'splitParagraph':
      edit = splitParagraph(doc, at);
      break;
    case 'setRunFormat':
      edit = setRunFormat(doc, span, d.pick([{ bold: true }, { italic: null }, { size: 40 }, { color: '00FF00' }, { font: 'Other' }, { underline: 'double' }, { vertAlign: 'superscript' as const }, { style: 'Strong' }, { bold: false, highlight: 'yellow' }]));
      break;
    case 'toggleRunFormat':
      edit = toggleRunFormat(doc, span, d.pick(['bold', 'italic', 'underline', 'strike'] as const));
      break;
    case 'setParaFormat':
      edit = setParaFormat(doc, span, d.pick([{ align: 'center' }, { align: null }, { style: 'Heading1' }, { indent: { left: 360 } }, { spacing: { before: 120, after: 120 } }, { numbering: { numId: 1, level: 1 } }, { style: null, indent: null }]));
      break;
    case 'toggleList':
      edit = toggleList(doc, span, d.chance(0.5) ? 'bullet' : 'number');
      break;
    case 'indentList':
      edit = indentList(doc, span, d.chance(0.5) ? 1 : -1);
      break;
    case 'insertTable':
      edit = insertTable(doc, at, 1 + d.int(3), 1 + d.int(3));
      break;
    case 'insertPageBreak':
      edit = insertPageBreak(doc, at);
      break;
    case 'insertHyperlink': {
      const end = { ...at, offset: Math.min(paragraphLength(paragraphAt(doc, at.path, story)), at.offset + d.int(6)) };
      edit = insertHyperlink(doc, { anchor: at, focus: end }, d.chance(0.3) ? '#mark' : 'https://example.org/page');
      break;
    }
    case 'insertImage':
      edit = insertImage(doc, at, { bytes: PNG, type: 'image/png', width: 95250, height: 95250 });
      break;
    default:
      if (!inTable) return undefined;
      if (name === 'insertRow') edit = insertRow(doc, at, d.chance(0.5) ? 'above' : 'below');
      else if (name === 'deleteRow') edit = deleteRow(doc, at);
      else if (name === 'insertColumn') edit = insertColumn(doc, at, d.chance(0.5) ? 'left' : 'right');
      else edit = deleteColumn(doc, at);
  }
  return { name, edit };
}

/**
 * A paragraph as a compact list, for expectations: a run is its formatting
 * flags (b, i, u, a size, a style), ":" and its items (text as it is, other
 * items as "<type>"), then "@" and the kinds of its wrappers, innermost
 * first; another inline is "<type>".
 */
export function shape(paragraph: Paragraph): string[] {
  return paragraph.children.map((inline) => {
    let wraps = '';
    for (let w = inline.wrap; w; w = w.parent) wraps += '@' + w.kind;
    if (inline.type !== 'run') return `<${inline.type}>${wraps}`;
    const props = inline.props;
    let flags = '';
    if (props.bold === true) flags += 'b';
    if (props.bold === false) flags += 'B';
    if (props.italic) flags += 'i';
    if (props.underline) flags += 'u';
    if (props.size !== undefined) flags += props.size;
    if (props.style !== undefined) flags += '{' + props.style + '}';
    const items = inline.content.map((item) => (item.type === 'text' ? item.text : item.type === 'fieldChar' ? `<${item.charType}>` : `<${item.type}>`)).join('');
    return `${flags}:${items}${wraps}`;
  });
}

/** shape() of the paragraph at a path. */
export const shapeAt = (doc: DocModel, path: number[] | number, story = 'body'): string[] => shape(paragraphAt(doc, path, story));
