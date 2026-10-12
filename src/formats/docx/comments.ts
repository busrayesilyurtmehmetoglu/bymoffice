// The parts beside word/comments.xml that say more about its comments:
//   - commentsExtended.xml (Word 2013): which comment answers which
//     (w15:paraIdParent) and which is resolved (w15:done);
//   - commentsIds.xml (Word 2016) and commentsExtensible.xml (Word 2018): an
//     id of each comment that lasts, and its date in UTC.
// All of them name a comment by the w14:paraId of its last paragraph.
//
// The model keeps what the first of them says in the comments themselves
// (Comment.paraId, .parent, .done). The parts stay in the package: the writer
// brings them up to date with the model by changing their trees (see
// extendedPartText and idPartText), and leaves them alone when they say what
// the model says, so that a comment that was added and taken back leaves the
// bytes of the file.

import type { AttrList, Comment, Paragraph, StoryCollection } from '../../model/doc/types';
import { localName, parseXml, serializeXml, type XmlDocument, type XmlElement } from '../../xml';

export const W14_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NAMESPACE = 'http://schemas.microsoft.com/office/word/2012/wordml';
export const W16CID_NAMESPACE = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
export const W16CEX_NAMESPACE = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
export const MC_NAMESPACE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** Relationship types of the three parts, as Word writes them (the same in strict documents). */
export const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const COMMENTS_IDS_REL = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds';
export const COMMENTS_EXTENSIBLE_REL = 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible';

const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
export const COMMENTS_EXTENDED_TYPE = WML + 'commentsExtended+xml';

/** The w14:textId a paragraph gets beside a new w14:paraId. */
export const DEFAULT_TEXT_ID = '77777777';

/** More entries than this in one of the parts is not a genuine file: the rest is not looked at. */
export const MAX_COMMENT_ENTRIES = 200000;

const PARA_ID = /^[0-9A-Fa-f]{8}$/;

/** The prefix (with its colon) the attributes of a root bind to a namespace; undefined when it declares none. */
export function prefixOf(rootAttrs: Iterable<[string, string]>, namespace: string): string | undefined {
  for (const [name, value] of rootAttrs) if (value === namespace && name.slice(0, 6) === 'xmlns:' && name.length > 6) return name.slice(6) + ':';
  return undefined;
}

/** The last paragraph among the blocks of a comment itself: the one whose id names the comment. */
export function lastParagraphOf(comment: { blocks: Comment['blocks'] }): Paragraph | undefined {
  for (let i = comment.blocks.length - 1; i >= 0; i--) {
    const block = comment.blocks[i];
    if (block.type === 'paragraph') return block;
    // What is neither text nor a table (a bookmark between blocks, a comment of the XML) is passed.
    if (block.type === 'table') return undefined;
  }
  return undefined;
}

/** The w14:paraId a paragraph carries, upper case; undefined when it has none that can be one. */
export function paraIdOf(paragraph: Paragraph | undefined, rootAttrs: AttrList): string | undefined {
  if (!paragraph || !paragraph.attrs) return undefined;
  const prefix = prefixOf(rootAttrs, W14_NAMESPACE);
  for (const [name, value] of paragraph.attrs) {
    // A part that does not declare the namespace is read by the conventional prefix.
    if (name === (prefix !== undefined ? prefix : 'w14:') + 'paraId') return PARA_ID.test(value) ? value.toUpperCase() : undefined;
  }
  return undefined;
}

const attrOf = (el: XmlElement, local: string): string | undefined => {
  let found: string | undefined;
  el.attrs.forEach((value, name) => {
    if (found === undefined && localName(name) === local) found = value;
  });
  return found;
};

const isOn = (value: string | undefined): boolean => value === '1' || value === 'true' || value === 'on';

/** What an entry of commentsExtended.xml says of the comment it names. */
export interface CommentLink {
  paraId: string;
  /** The paraId of the comment it answers. */
  parent?: string;
  done: boolean;
}

/** The entries of a commentsExtended.xml, by the paragraph id they name; of two entries for one id the first counts. */
export function extendedLinks(root: XmlElement): Map<string, CommentLink> {
  const out = new Map<string, CommentLink>();
  let seen = 0;
  for (const node of root.children) {
    if (node.type !== 'element' || localName(node.name) !== 'commentEx') continue;
    if (++seen > MAX_COMMENT_ENTRIES) break;
    const id = attrOf(node, 'paraId');
    if (id === undefined || !PARA_ID.test(id) || out.has(id.toUpperCase())) continue;
    const link: CommentLink = { paraId: id.toUpperCase(), done: isOn(attrOf(node, 'done')) };
    const parent = attrOf(node, 'paraIdParent');
    if (parent !== undefined && PARA_ID.test(parent)) link.parent = parent.toUpperCase();
    out.set(link.paraId, link);
  }
  return out;
}

/**
 * Gives every comment the id of its last paragraph (Comment.paraId) and, with
 * the root of commentsExtended.xml, what that part says of it: the comment it
 * answers and whether it is resolved. An entry that names no comment says
 * nothing here (it stays in the part); an answer to itself, or answers in a
 * ring, are answers to nothing.
 */
export function linkComments(comments: StoryCollection<Comment>, extended: XmlElement | undefined): void {
  const byParaId = new Map<string, Comment>();
  comments.items.forEach((comment) => {
    const id = paraIdOf(lastParagraphOf(comment), comments.rootAttrs);
    // Two comments that claim one id: the first has it, the second is known by none.
    if (id === undefined || byParaId.has(id)) return;
    comment.paraId = id;
    byParaId.set(id, comment);
  });
  if (!extended) return;
  const links = extendedLinks(extended);
  byParaId.forEach((comment, id) => {
    const link = links.get(id);
    if (!link) return;
    if (link.done) comment.done = true;
    const parent = link.parent !== undefined ? byParaId.get(link.parent) : undefined;
    if (parent && parent !== comment) comment.parent = parent.id;
  });
  // No comment is, through its answers, an answer to itself: the comments of a ring answer nothing.
  threadRoots(comments.items).ring.forEach((comment) => {
    delete comment.parent;
  });
}

/**
 * The first comment of the thread of every comment, whatever the depth of
 * the answers: a comment that answers none (or one that is gone) is the
 * first of its own thread. Comments that answer each other in a ring are
 * each the first of a thread (`ring` holds them), and what answers into the
 * ring belongs to the comment of the ring it leads to. Every comment is
 * walked over once.
 */
export function threadRoots<T extends { id: number; parent?: number }>(items: Map<number, T>): { roots: Map<T, T>; ring: Set<T> } {
  const roots = new Map<T, T>();
  const ring = new Set<T>();
  const place = new Map<T, number>();
  items.forEach((start) => {
    if (roots.has(start)) return;
    // The way from this comment towards the first of its thread, until something known is met.
    const path: T[] = [];
    place.clear();
    let at = start;
    let root: T;
    for (;;) {
      const known = roots.get(at);
      if (known) {
        root = known;
        break;
      }
      place.set(at, path.length);
      path.push(at);
      const parent = at.parent !== undefined ? items.get(at.parent) : undefined;
      if (!parent) {
        root = at;
        break;
      }
      const seen = place.get(parent);
      if (seen !== undefined) {
        // A ring: its members are firsts of their own; the way that led into it ends at the member it met.
        for (let i = seen; i < path.length; i++) {
          ring.add(path[i]);
          roots.set(path[i], path[i]);
        }
        path.length = seen;
        root = parent;
        break;
      }
      at = parent;
    }
    for (const comment of path) roots.set(comment, root);
  });
  return { roots, ring };
}

/** What the model says of its comments, for the parts beside them: one entry for each comment that has a paragraph id, in their order. */
export function modelLinks(comments: StoryCollection<Comment> | undefined): CommentLink[] {
  const out: CommentLink[] = [];
  if (!comments) return out;
  comments.items.forEach((comment) => {
    if (comment.paraId === undefined) return;
    const link: CommentLink = { paraId: comment.paraId, done: comment.done === true };
    const parent = comment.parent !== undefined ? comments.items.get(comment.parent) : undefined;
    if (parent && parent.paraId !== undefined) link.parent = parent.paraId;
    out.push(link);
  });
  return out;
}

/** Do the model's comments need a commentsExtended.xml: one of them is an answer or is resolved. */
export function needsExtended(links: CommentLink[]): boolean {
  for (const link of links) if (link.parent !== undefined || link.done) return true;
  return false;
}

/** Every w14:paraId the text of a comments part holds (upper case): the ids of the comments the file had. */
export function paraIdsInSource(source: string): Set<string> {
  const out = new Set<string>();
  if (!source) return out;
  const pattern = /:paraId\s*=\s*(?:"([0-9A-Fa-f]{8})"|'([0-9A-Fa-f]{8})')/g;
  for (let match = pattern.exec(source); match && out.size < MAX_COMMENT_ENTRIES; match = pattern.exec(source)) out.add((match[1] !== undefined ? match[1] : match[2]).toUpperCase());
  return out;
}

const sameLink = (a: CommentLink | undefined, b: CommentLink | undefined): boolean =>
  a !== undefined && b !== undefined && a.paraId === b.paraId && a.parent === b.parent && a.done === b.done;

const rootAttrList = (root: XmlElement): [string, string][] => Array.from(root.attrs);

/**
 * The text of commentsExtended.xml for the comments of the model, or
 * undefined when the part the file has says of them what the model says (and
 * so stays as it is), or when there is no such part and none is needed.
 *
 * `original` is the part as the file had it; `had` the paragraph ids of the
 * comments the file had (see paraIdsInSource). An entry of the file for a
 * comment that is gone is taken out; one that names nothing the file ever had
 * as a comment stays, with everything else the part holds.
 */
export function extendedPartText(original: XmlDocument | undefined, links: CommentLink[], had: Set<string>): string | undefined {
  const wanted = new Map<string, CommentLink>();
  for (const link of links) wanted.set(link.paraId, link);
  if (!original) {
    if (!needsExtended(links)) return undefined;
    let text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w15:commentsEx xmlns:mc="${MC_NAMESPACE}" xmlns:w15="${W15_NAMESPACE}" mc:Ignorable="w15">`;
    for (const link of links) text += `<w15:commentEx w15:paraId="${link.paraId}"${link.parent !== undefined ? ` w15:paraIdParent="${link.parent}"` : ''} w15:done="${link.done ? '1' : '0'}"/>`;
    return text + '</w15:commentsEx>';
  }
  const said = extendedLinks(original.root);
  let same = true;
  // The file says of each comment of the model what the model says. A comment of the file that it has no entry
  // for is one that is no answer and is not resolved, and stays without one; a comment that is new gets its entry,
  // as Word keeps one for each.
  for (const link of links) {
    const entry = said.get(link.paraId);
    if (entry ? !sameLink(entry, link) : link.parent !== undefined || link.done || !had.has(link.paraId)) same = false;
  }
  // And it says nothing of a comment that is gone.
  said.forEach((_entry, id) => {
    if (!wanted.has(id) && had.has(id)) same = false;
  });
  if (same) return undefined;

  const root = original.root;
  const prefix = prefixOf(rootAttrList(root), W15_NAMESPACE) || (root.name.indexOf(':') >= 0 ? root.name.slice(0, root.name.indexOf(':') + 1) : '');
  const children: XmlElement['children'] = [];
  const written = new Set<string>();
  for (const node of root.children) {
    if (node.type !== 'element' || localName(node.name) !== 'commentEx') {
      children.push(node);
      continue;
    }
    const raw = attrOf(node, 'paraId');
    const id = raw !== undefined && PARA_ID.test(raw) ? raw.toUpperCase() : undefined;
    const link = id !== undefined ? wanted.get(id) : undefined;
    if (id === undefined || (!link && !had.has(id)) || written.has(id)) {
      // Not an entry of a comment: kept as it is.
      children.push(node);
      continue;
    }
    if (!link) continue;
    written.add(id);
    // The attributes the entry had, in their order, with the two the model decides.
    const attrs = new Map<string, string>();
    let parentName: string | undefined;
    let doneName: string | undefined;
    node.attrs.forEach((value, name) => {
      const local = localName(name);
      if (local === 'paraIdParent') parentName = name;
      else if (local === 'done') doneName = name;
      if (local === 'paraIdParent' && link.parent === undefined) return;
      attrs.set(name, local === 'paraIdParent' ? (link.parent as string) : local === 'done' ? (link.done ? '1' : '0') : value);
    });
    if (link.parent !== undefined && parentName === undefined) attrs.set(prefix + 'paraIdParent', link.parent);
    if (doneName === undefined) attrs.set(prefix + 'done', link.done ? '1' : '0');
    children.push({ type: 'element', name: node.name, attrs, children: node.children });
  }
  for (const link of links) {
    if (written.has(link.paraId)) continue;
    const attrs = new Map<string, string>([[prefix + 'paraId', link.paraId]]);
    if (link.parent !== undefined) attrs.set(prefix + 'paraIdParent', link.parent);
    attrs.set(prefix + 'done', link.done ? '1' : '0');
    children.push({ type: 'element', name: prefix + 'commentEx', attrs, children: [] });
  }
  return serializeXml({ ...original, root: { type: 'element', name: root.name, attrs: root.attrs, children } });
}

/** An 8 digit hex number that is none of `taken` and is added to them: deterministic, below 0x80000000 and not 0. */
export function freshHexId(taken: Set<string>, seed: number): string {
  let value = (Math.imul(seed | 0, 0x9e3779b1) ^ 0x5bd1e995) >>> 0;
  for (let tries = 0; tries < 0x7fffffff; tries++) {
    value = (Math.imul(value, 1664525) + 1013904223) >>> 0;
    const candidate = value & 0x7fffffff;
    if (candidate === 0) continue;
    let text = candidate.toString(16).toUpperCase();
    while (text.length < 8) text = '0' + text;
    if (taken.has(text)) continue;
    taken.add(text);
    return text;
  }
  throw new RangeError('The document has no room for another comment');
}

/**
 * The text of commentsIds.xml brought up to date with the comments of the
 * model: an entry for each comment it has none for (with a durable id no
 * other entry has), none for a comment that is gone. Undefined when the part
 * needs no change. Returns the durable ids that were taken out, for the
 * commentsExtensible part, which names comments by them.
 */
export function idsPartText(original: XmlDocument, links: CommentLink[], had: Set<string>): { text: string; removed: Set<string>; added: string[] } | undefined {
  const root = original.root;
  const wanted = new Set<string>();
  for (const link of links) wanted.add(link.paraId);
  const prefix = prefixOf(rootAttrList(root), W16CID_NAMESPACE) || (root.name.indexOf(':') >= 0 ? root.name.slice(0, root.name.indexOf(':') + 1) : '');
  const present = new Set<string>();
  const durable = new Set<string>();
  const removed = new Set<string>();
  const children: XmlElement['children'] = [];
  let seen = 0;
  for (const node of root.children) {
    if (node.type !== 'element' || localName(node.name) !== 'commentId' || ++seen > MAX_COMMENT_ENTRIES) {
      children.push(node);
      continue;
    }
    const raw = attrOf(node, 'paraId');
    const id = raw !== undefined && PARA_ID.test(raw) ? raw.toUpperCase() : undefined;
    const lasting = attrOf(node, 'durableId');
    if (id !== undefined && !wanted.has(id) && had.has(id)) {
      if (lasting !== undefined) removed.add(lasting.toUpperCase());
      continue;
    }
    if (id !== undefined) present.add(id);
    if (lasting !== undefined) durable.add(lasting.toUpperCase());
    children.push(node);
  }
  const added: string[] = [];
  let seed = 1;
  for (const link of links) {
    // A comment of the file that the part has no entry for stays without one: only what is new here is added.
    if (present.has(link.paraId) || had.has(link.paraId)) continue;
    const lasting = freshHexId(durable, parseInt(link.paraId, 16) + seed++);
    added.push(lasting);
    children.push({ type: 'element', name: prefix + 'commentId', attrs: new Map([[prefix + 'paraId', link.paraId], [prefix + 'durableId', lasting]]), children: [] });
  }
  if (!removed.size && !added.length) return undefined;
  return { text: serializeXml({ ...original, root: { type: 'element', name: root.name, attrs: root.attrs, children } }), removed, added };
}

/**
 * The text of commentsExtensible.xml without the entries of comments that
 * are gone (by their durable ids); undefined when it has none of them. A new
 * comment gets no entry: the part only adds a date in UTC to what the
 * comment says itself.
 */
export function extensiblePartText(original: XmlDocument, removed: Set<string>): string | undefined {
  if (!removed.size) return undefined;
  const root = original.root;
  let changed = false;
  const children = root.children.filter((node) => {
    if (node.type !== 'element' || localName(node.name) !== 'commentExtensible') return true;
    const lasting = attrOf(node, 'durableId');
    if (lasting === undefined || !removed.has(lasting.toUpperCase())) return true;
    changed = true;
    return false;
  });
  if (!changed) return undefined;
  return serializeXml({ ...original, root: { type: 'element', name: root.name, attrs: root.attrs, children } });
}

/** A part beside the comments as a tree; undefined (with nothing thrown) for bytes that are not XML. */
export function parseSidePart(bytes: Uint8Array | undefined): XmlDocument | undefined {
  if (!bytes) return undefined;
  try {
    return parseXml(bytes);
  } catch {
    return undefined;
  }
}

/**
 * The root attributes a comments part needs so that its paragraphs can carry
 * w14:paraId: the attributes given when they bind a prefix to the namespace
 * already; else a copy with the declaration and, for readers that do not know
 * the namespace, the prefix among the ones to ignore (mc:Ignorable).
 */
export function withParaIdNamespace(rootAttrs: AttrList): { attrs: AttrList; prefix: string } {
  const known = prefixOf(rootAttrs, W14_NAMESPACE);
  if (known !== undefined) return { attrs: rootAttrs, prefix: known };
  const taken = new Set<string>();
  for (const [name] of rootAttrs) if (name.slice(0, 6) === 'xmlns:') taken.add(name.slice(6));
  let prefix = 'w14';
  for (let n = 2; taken.has(prefix) && n < 1000; n++) prefix = 'w14n' + n;
  const attrs: AttrList = [];
  let mc = prefixOf(rootAttrs, MC_NAMESPACE);
  let ignorable = false;
  for (const pair of rootAttrs) {
    if (mc !== undefined && pair[0] === mc + 'Ignorable') {
      ignorable = true;
      attrs.push([pair[0], pair[1].trim() === '' ? prefix : `${pair[1]} ${prefix}`]);
    } else attrs.push([pair[0], pair[1]]);
  }
  // The declarations go before the attribute that names their prefixes.
  const at = attrs.findIndex(([name]) => name.slice(0, 5) !== 'xmlns');
  const declarations: AttrList = [['xmlns:' + prefix, W14_NAMESPACE]];
  if (mc === undefined) {
    let name = 'mc';
    for (let n = 2; taken.has(name) && n < 1000; n++) name = 'mc' + n;
    mc = name + ':';
    declarations.unshift(['xmlns:' + name, MC_NAMESPACE]);
  }
  attrs.splice(at < 0 ? attrs.length : at, 0, ...declarations);
  if (!ignorable) attrs.push([mc + 'Ignorable', prefix]);
  return { attrs, prefix: prefix + ':' };
}

/**
 * The comments as they are written: every comment that is known by a
 * paragraph id which its last paragraph does not carry (one that was made or
 * changed here) gets a copy of that paragraph with the attribute, and the
 * root the declaration the attribute needs. `rootAttrs` are the attributes
 * the root of the part is written with, when they are not those of the
 * collection (a part made from nothing). The collection itself when nothing
 * was missing.
 */
export function commentsForWriting(comments: StoryCollection<Comment>, rootAttrs: AttrList = comments.rootAttrs): StoryCollection<Comment> {
  let items: Map<number, Comment> | undefined;
  let space: { attrs: AttrList; prefix: string } | undefined;
  comments.items.forEach((comment, id) => {
    if (comment.paraId === undefined) return;
    const last = lastParagraphOf(comment);
    if (!last || paraIdOf(last, rootAttrs) === comment.paraId) return;
    if (!space) space = withParaIdNamespace(rootAttrs);
    const name = space.prefix + 'paraId';
    const textName = space.prefix + 'textId';
    const attrs: AttrList = (last.attrs || []).filter((pair) => pair[0] !== name);
    attrs.push([name, comment.paraId]);
    // The two ids go together in the files of Word and of LibreOffice; this value is the one Word writes for a
    // paragraph whose text it has no version of. UNSURE whether Word minds a paragraph id without it.
    if (!attrs.some((pair) => pair[0] === textName)) attrs.push([textName, DEFAULT_TEXT_ID]);
    const paragraph: Paragraph = { ...last, attrs };
    delete paragraph.srcStart;
    delete paragraph.srcEnd;
    const blocks = comment.blocks.slice();
    blocks[blocks.lastIndexOf(last)] = paragraph;
    const copy: Comment = { ...comment, blocks };
    delete copy.srcStart;
    delete copy.srcEnd;
    if (!items) items = new Map(comments.items);
    items.set(id, copy);
  });
  if (!items || !space) return comments;
  // A part that is written with a root of its own making (see StoryWriteOptions.root) has the declaration there.
  return comments.rootName === '' ? { ...comments, items } : { ...comments, items, rootAttrs: space.attrs };
}
