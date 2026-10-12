// Synthetic documents with comments, and the checks every saved package with
// comments has to pass. Everything is built from XML fragments in code.

import assert from 'node:assert/strict';
import { readDocx } from '../src/formats/docx/document';
import type { DocModel } from '../src/model/doc/types';
import { listComments, type DocComment } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { childElements, localName, parseXml, type XmlElement } from '../src/xml';
import { STYLES } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';

export const EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const IDS_REL = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
export const EXTENSIBLE_REL = 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The two styles Word gives the text of a comment and its mark. */
export const COMMENT_STYLES =
  STYLES +
  '<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>';

export const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const cStart = (id: number) => `<w:commentRangeStart w:id="${id}"/>`;
export const cEnd = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
export const cRef = (id: number) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;

export interface CommentSpec {
  id: number;
  author?: string;
  text?: string;
  /** w14:paraId of the last paragraph; none when undefined. */
  paraId?: string;
  date?: string;
  initials?: string;
  /** The whole content, instead of a paragraph with `text`. */
  content?: string;
  /** More attributes of <w:comment>, as text. */
  attrs?: string;
}

export function commentXml(spec: CommentSpec): string {
  const content =
    spec.content !== undefined
      ? spec.content
      : `<w:p${spec.paraId !== undefined ? ` w14:paraId="${spec.paraId}"` : ''}><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>` +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' +
        `<w:r><w:t xml:space="preserve">${escape(spec.text === undefined ? '' : spec.text)}</w:t></w:r></w:p>`;
  return (
    `<w:comment w:id="${spec.id}" w:author="${escape(spec.author === undefined ? 'Ann' : spec.author)}"` +
    ` w:date="${spec.date === undefined ? '2026-10-01T09:00:00Z' : spec.date}"${spec.initials !== undefined ? ` w:initials="${escape(spec.initials)}"` : ''}${spec.attrs ? ' ' + spec.attrs : ''}>${content}</w:comment>`
  );
}

/** A commentsExtended.xml: [paraId, paraId of the comment it answers, resolved]. */
export const extendedXml = (entries: [paraId: string, parent?: string, done?: boolean][], extra = '') =>
  `${XML}\n<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">` +
  entries.map(([id, parent, done]) => `<w15:commentEx w15:paraId="${id}"${parent ? ` w15:paraIdParent="${parent}"` : ''} w15:done="${done ? 1 : 0}"/>`).join('') +
  `${extra}</w15:commentsEx>`;

export const idsXml = (entries: [paraId: string, durableId: string][]) =>
  `${XML}\n<w16cid:commentsIds xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" mc:Ignorable="w16cid">` +
  entries.map(([id, durable]) => `<w16cid:commentId w16cid:paraId="${id}" w16cid:durableId="${durable}"/>`).join('') +
  '</w16cid:commentsIds>';

export const extensibleXml = (entries: [durableId: string, dateUtc: string][]) =>
  `${XML}\n<w16cex:commentsExtensible xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" mc:Ignorable="w16cex">` +
  entries.map(([durable, date]) => `<w16cex:commentExtensible w16cex:durableId="${durable}" w16cex:dateUtc="${date}"/>`).join('') +
  '</w16cex:commentsExtensible>';

export interface CommentedSpec {
  body: string;
  comments?: CommentSpec[];
  /** Entries of commentsExtended.xml; the part is left out when undefined. */
  extended?: [paraId: string, parent?: string, done?: boolean][];
  ids?: [paraId: string, durableId: string][];
  extensible?: [durableId: string, dateUtc: string][];
  settings?: string;
  styles?: string;
  footnotes?: string;
  headers?: Record<string, string>;
}

/** The spec of a document with comments and the parts beside them, declared in the content types as Word declares them. */
export function commentedSpec(spec: CommentedSpec): DocxSpec {
  const out: DocxSpec = { body: spec.body.indexOf('<w:sectPr') >= 0 ? spec.body : spec.body + SECT, styles: spec.styles === undefined ? COMMENT_STYLES : spec.styles };
  if (spec.comments) out.comments = spec.comments.map(commentXml).join('');
  if (spec.settings !== undefined) out.settings = spec.settings;
  if (spec.footnotes !== undefined) out.footnotes = spec.footnotes;
  if (spec.headers !== undefined) out.headers = spec.headers;
  const rels: NonNullable<DocxSpec['rels']> = [];
  const parts: NonNullable<DocxSpec['parts']> = [];
  if (spec.extended) {
    rels.push(['rIdCex', EXTENDED_REL, 'commentsExtended.xml']);
    parts.push(['word/commentsExtended.xml', extendedXml(spec.extended)]);
  }
  if (spec.ids) {
    rels.push(['rIdCid', IDS_REL, 'commentsIds.xml']);
    parts.push(['word/commentsIds.xml', idsXml(spec.ids)]);
  }
  if (spec.extensible) {
    rels.push(['rIdCxt', EXTENSIBLE_REL, 'commentsExtensible.xml']);
    parts.push(['word/commentsExtensible.xml', extensibleXml(spec.extensible)]);
  }
  if (rels.length) out.rels = rels;
  if (parts.length) out.parts = parts;
  return out;
}

export async function openCommented(spec: CommentedSpec): Promise<{ pkg: OpcPackage; doc: DocModel; bytes: Uint8Array; original: Map<string, Uint8Array> }> {
  const bytes = await buildSpec(commentedSpec(spec));
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const original = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) original.set(part, await pkg.readPart(part));
  return { pkg, doc, bytes, original };
}

/** [id, author, text, parent, done] of every comment, as listComments() gives them. */
export const brief = (list: DocComment[]): [number, string, string, number | undefined, boolean][] => list.map((c) => [c.id, c.author, c.text, c.parent, c.done]);

/** [start path, start offset, end path, end offset] of a comment. */
export const span = (comment: DocComment | undefined): [number[], number, number[], number] | undefined =>
  comment && comment.start && comment.end ? [comment.start.path, comment.start.offset, comment.end.path, comment.end.offset] : undefined;

const decoder = new TextDecoder();

function everyElement(el: XmlElement, visit: (el: XmlElement) => void, depth = 0): void {
  if (depth > 512) return;
  visit(el);
  for (const node of el.children) if (node.type === 'element') everyElement(node, visit, depth + 1);
}

const attr = (el: XmlElement, local: string): string | undefined => {
  for (const [name, value] of el.attrs) if (localName(name) === local) return value;
  return undefined;
};

export interface PackageFacts {
  /** Ids of the comments of comments.xml, in order. */
  ids: number[];
  /** The paragraph id of the last paragraph of each comment, by the id of the comment. */
  paraIds: Map<number, string | undefined>;
  /** Entries of commentsExtended.xml by paragraph id; undefined when there is no such part. */
  extended: Map<string, { parent?: string; done: boolean }> | undefined;
  parts: string[];
}

/**
 * Reads a saved package the way a consumer that knows nothing of the engine
 * would, and requires what holds for every document with comments: each part
 * is well-formed; every mark and reference of a comment names a comment the
 * comments part has, and every comment is referred to exactly once; a range
 * that begins also ends; ids of comments and of paragraphs are unique; every
 * entry of commentsExtended.xml names the last paragraph of a comment, and
 * the comment an answer names exists; the parts the package relates to are
 * there and declared.
 */
export async function checkCommentedPackage(bytes: Uint8Array, options: { orphans?: number[] } = {}): Promise<PackageFacts> {
  const pkg = await OpcPackage.open(bytes);
  const parts = pkg.partNames().sort();
  const trees = new Map<string, XmlElement>();
  for (const part of parts) {
    if (!/\.(xml|rels)$/.test(part)) continue;
    trees.set(part, parseXml(decoder.decode(await pkg.readPart(part))).root);
  }
  const rels = await pkg.relationships('/word/document.xml');
  for (const rel of rels) {
    if (rel.external || rel.part === null) continue;
    assert.ok(pkg.hasPart(rel.part), `the document relates to ${rel.part}, which the package does not have`);
    assert.ok(pkg.contentTypeOf(rel.part) !== undefined, `${rel.part} has no content type`);
  }
  const ids = new Set<string>(rels.map((rel) => rel.id));
  assert.equal(ids.size, rels.length, 'two relationships of the document share an id');
  const commentsRel = rels.filter((rel) => /\/comments$/.test(rel.type));
  const starts = new Map<string, number>();
  const ends = new Map<string, number>();
  const refs = new Map<string, number>();
  const count = (map: Map<string, number>, id: string | undefined) => {
    assert.ok(id !== undefined && /^\d+$/.test(id), `a mark of a comment has the id ${String(id)}`);
    map.set(id as string, (map.get(id as string) || 0) + 1);
  };
  for (const [part, root] of trees) {
    if (!/^\/word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(part)) continue;
    everyElement(root, (el) => {
      const local = localName(el.name);
      if (local === 'commentRangeStart') count(starts, attr(el, 'id'));
      else if (local === 'commentRangeEnd') count(ends, attr(el, 'id'));
      else if (local === 'commentReference') count(refs, attr(el, 'id'));
    });
  }
  const facts: PackageFacts = { ids: [], paraIds: new Map(), extended: undefined, parts };
  if (commentsRel.length === 0) {
    assert.equal(starts.size + ends.size + refs.size, 0, 'the text has marks of comments and the package has no comments part');
    return facts;
  }
  assert.equal(commentsRel.length, 1, 'the document relates to more than one comments part');
  const commentsPart = commentsRel[0].part as string;
  const root = trees.get(commentsPart) as XmlElement;
  const seen = new Set<string>();
  const paraIds = new Set<string>();
  for (const comment of childElements(root, 'comment')) {
    const id = attr(comment, 'id');
    assert.ok(id !== undefined && /^\d+$/.test(id), 'a comment has no usable id');
    assert.ok(!seen.has(id as string), `two comments have the id ${id}`);
    seen.add(id as string);
    assert.ok(attr(comment, 'author') !== undefined, `the comment ${id} has no author`);
    facts.ids.push(Number(id));
    const paragraphs = childElements(comment, 'p');
    let last: string | undefined;
    everyElement(comment, (el) => {
      if (localName(el.name) !== 'p') return;
      const paraId = attr(el, 'paraId');
      if (paraId === undefined) return;
      assert.match(paraId, /^[0-9A-F]{8}$/, `the paragraph id ${paraId} is not 8 hex digits`);
      assert.ok(parseInt(paraId, 16) < 0x80000000 && parseInt(paraId, 16) > 0, `the paragraph id ${paraId} is out of range`);
      assert.ok(!paraIds.has(paraId), `two paragraphs of the comments have the id ${paraId}`);
      paraIds.add(paraId);
      // The prefix of the attribute is declared for the part.
      for (const name of el.attrs.keys()) if (localName(name) === 'paraId') assert.equal(root.attrs.get('xmlns:' + name.slice(0, name.indexOf(':'))), 'http://schemas.microsoft.com/office/word/2010/wordml', 'the prefix of a paragraph id is not declared');
    });
    if (paragraphs.length) last = attr(paragraphs[paragraphs.length - 1], 'paraId');
    facts.paraIds.set(Number(id), last);
  }
  const orphans = new Set((options.orphans || []).map(String));
  for (const id of seen) {
    if (orphans.has(id)) continue;
    assert.equal(refs.get(id), 1, `the comment ${id} is referred to ${refs.get(id) || 0} times`);
    assert.equal(starts.get(id) || 0, ends.get(id) || 0, `the range of the comment ${id} does not begin as often as it ends`);
    assert.ok((starts.get(id) || 0) <= 1, `the range of the comment ${id} begins more than once`);
  }
  for (const map of [starts, ends, refs]) for (const id of map.keys()) assert.ok(seen.has(id), `a mark names the comment ${id}, which the comments part does not have`);
  // The prefixes mc:Ignorable names are declared.
  for (const [part, tree] of trees) {
    for (const [name, value] of tree.attrs) {
      if (localName(name) !== 'Ignorable' || name.indexOf(':') < 0) continue;
      for (const prefix of value.split(/\s+/)) if (prefix) assert.ok(tree.attrs.has('xmlns:' + prefix), `${part} ignores the prefix ${prefix}, which it does not declare`);
    }
  }
  const extendedRel = rels.filter((rel) => rel.type === EXTENDED_REL);
  if (extendedRel.length) {
    assert.equal(extendedRel.length, 1, 'the document relates to more than one commentsExtended part');
    const part = extendedRel[0].part as string;
    const extended = new Map<string, { parent?: string; done: boolean }>();
    const lasts = new Set<string>();
    facts.paraIds.forEach((value) => (value !== undefined ? lasts.add(value) : 0));
    for (const entry of childElements(trees.get(part) as XmlElement, 'commentEx')) {
      const paraId = attr(entry, 'paraId') as string;
      assert.ok(!extended.has(paraId), `commentsExtended.xml has two entries for ${paraId}`);
      assert.ok(lasts.has(paraId), `commentsExtended.xml names ${paraId}, which is the last paragraph of no comment`);
      const parent = attr(entry, 'paraIdParent');
      const done = attr(entry, 'done');
      extended.set(paraId, parent === undefined ? { done: done === '1' } : { parent, done: done === '1' });
    }
    extended.forEach((entry, paraId) => {
      if (entry.parent === undefined) return;
      assert.ok(extended.has(entry.parent), `the comment ${paraId} answers ${entry.parent}, which commentsExtended.xml does not have`);
      assert.notEqual(entry.parent, paraId, 'a comment answers itself');
    });
    facts.extended = extended;
  }
  return facts;
}

/** Requires that reading the bytes gives the comments the model has: the same list, the same places. */
export async function sameCommentsAfterReading(bytes: Uint8Array, doc: DocModel): Promise<DocModel> {
  const again = await readDocx(await OpcPackage.open(bytes));
  const strip = (list: DocComment[]) => list.map((c) => ({ ...c }));
  assert.deepEqual(strip(listComments(again)), strip(listComments(doc)), 'the file does not read back with the comments of the model');
  return again;
}
