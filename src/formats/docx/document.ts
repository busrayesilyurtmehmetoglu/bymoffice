// Reads a docx package into the document model.
//
// Styles, numbering, theme, settings and the font table are small and go
// through the tree parser; the body, headers, footers, notes and comments are
// scanned as streams (see body.ts), with a cursor that follows the scanner
// through the source text (see source.ts): the model keeps that text once per
// part and its blocks point into it. A supporting part that is missing, or a
// relationship that points nowhere, is noted in `warnings` and read as empty;
// a part that is present but not well-formed XML is an error.

import {
  type AttrList,
  type Block,
  type Body,
  type Comment,
  type DocModel,
  type DocRelationship,
  type Note,
  type SourceSpan,
  type Story,
  type StoryCollection,
} from '../../model/doc/types';
import type { OpcPackage, Relationship } from '../../opc';
import { parseXml, type XmlComment, type XmlDocument, type XmlPI } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { linkFields, scanBlocks } from './body';
import { readDocCharts, type DocChartLimits } from './charts';
import { linkComments, parseSidePart } from './comments';
import { bounded, createContext, createReaderState, kindOf, pushExtra, pushMisc, quote, type PartContext, type ReaderState } from './common';
import { partNames } from '../xlsx/drawings';
import { parseNumbering } from './numbering';
import { parseFontTable, parseSettings, parseTheme, themeColorList } from './parts';
import { SourceCursor, decodePart } from './source';
import { parseDocStyles } from './styles';

const NO_RELATIONSHIPS = new Map<string, DocRelationship>();

/** The part a relationship type leads to, or undefined (with a warning when the target is missing). */
function supportingPart(pkg: OpcPackage, rels: Relationship[], suffix: string, state: ReaderState): string | undefined {
  const rel = rels.find((candidate) => candidate.type.endsWith(suffix) && !candidate.external);
  if (!rel || !rel.part) return undefined;
  if (pkg.hasPart(rel.part)) return rel.part;
  state.warnings.push(`The ${suffix.slice(1)} part ${quote(rel.part)} is missing`);
  return undefined;
}

async function readTree(pkg: OpcPackage, part: string | undefined, state: ReaderState): Promise<{ doc: XmlDocument; ctx: PartContext } | undefined> {
  if (part === undefined) return undefined;
  const doc = await pkg.readXml(part);
  bounded(doc.root);
  return { doc, ctx: createContext(state, part, Array.from(doc.root.attrs), NO_RELATIONSHIPS) };
}

interface OpenStory {
  sc: XmlScanner;
  /** On the root start tag, like the scanner. */
  cursor: SourceCursor;
  ctx: PartContext;
  story: Story;
}

/** Opens a part for scanning, positioned on its root element. */
async function openStory(pkg: OpcPackage, part: string, state: ReaderState): Promise<OpenStory> {
  const source = decodePart(await pkg.readPart(part));
  const sc = new XmlScanner(source);
  // next() throws on an empty part; the first token of a well-formed one is the root.
  sc.next();
  if (sc.kind !== 'open') throw new Error(`Invalid docx: ${quote(part)} has no root element`);
  const rootName = sc.name;
  const rootAttrs = sc.attrs();
  const relationships = new Map<string, DocRelationship>();
  for (const rel of await pkg.relationships(part)) relationships.set(rel.id, rel);
  const ctx = createContext(state, part, rootAttrs, relationships);
  const story: Story = { part, rootName, rootAttrs, blocks: [], relationships, source };
  const cursor = new SourceCursor(source);
  cursor.toTag();
  if (!cursor.lost && cursor.pos > 0) story.prolog = source.slice(0, cursor.pos);
  return { sc, cursor, ctx, story };
}

/**
 * Runs the scanner to the end of the part, which checks what is left for
 * well-formedness, and keeps what follows the root element if it is more
 * than whitespace.
 */
function finish(sc: XmlScanner, cursor: SourceCursor, story: { source: string; epilog?: string }): void {
  while (sc.next()) {
    // Nothing to do with the tokens.
  }
  if (cursor.lost) return;
  const rest = story.source.slice(cursor.pos);
  if (rest.trim()) story.epilog = rest;
}

async function readBody(pkg: OpcPackage, part: string, state: ReaderState): Promise<Body> {
  const { sc, cursor, ctx, story } = await openStory(pkg, part, state);
  if (kindOf(ctx, sc.name) !== 'document') throw new Error('Invalid docx: the main part is not a WordprocessingML document');
  const body: Body = story;
  let seen = false;
  let index = -1;
  cursor.enter();
  if (!sc.selfClosing) {
    const misc = (node: XmlComment | XmlPI) => {
      body.misc = pushMisc(body.misc, index + 1, node);
    };
    while (sc.next()) {
      if (sc.kind === 'text') continue;
      cursor.toTag(misc);
      if (sc.kind === 'close') {
        cursor.leave();
        break;
      }
      index++;
      if (!seen && kindOf(ctx, sc.name) === 'body') {
        seen = true;
        const result = scanBlocks(ctx, sc, cursor, body.blocks);
        if (result.section) {
          body.section = result.section;
          if (result.sectionSpan) body.sectionSpan = result.sectionSpan;
          // The section mark is the last child of a valid body; its place only needs saying when it is not.
          if (result.sectionIndex !== undefined && result.sectionIndex < body.blocks.length) body.sectionIndex = result.sectionIndex;
        }
      } else {
        const raw = sc.readRaw();
        cursor.skip(raw);
        body.extras = pushExtra(ctx, body.extras, index, bounded(parseXml(raw).root), 'document', false);
      }
    }
    finish(sc, cursor, body);
  }
  if (!seen) throw new Error('Invalid docx: the document has no body');
  linkFields(body.blocks);
  return body;
}

async function readHeaderFooter(pkg: OpcPackage, part: string, state: ReaderState): Promise<Story> {
  const { sc, cursor, ctx, story } = await openStory(pkg, part, state);
  scanBlocks(ctx, sc, cursor, story.blocks);
  finish(sc, cursor, story);
  linkFields(story.blocks);
  return story;
}

function attrOf(ctx: PartContext, attrs: AttrList, local: string): string | undefined {
  for (const [name, value] of attrs) if (kindOf(ctx, name) === local) return value;
  return undefined;
}

const ID = /^-?\d{1,10}$/;

/**
 * Reads a part that is a list of small stories (<w:footnote>, <w:endnote>,
 * <w:comment>). An item without a usable id, or with an id seen before, has
 * no place among the items: it is kept whole where it was, like any other
 * child of the root that is not an item, and reported with a warning.
 */
async function readCollection<T extends SourceSpan & { id: number; blocks: Block[]; attrs?: AttrList }>(
  pkg: OpcPackage,
  part: string,
  state: ReaderState,
  itemKind: string,
  fill: (item: T, ctx: PartContext, attrs: AttrList) => void,
): Promise<StoryCollection<T>> {
  const { sc, cursor, ctx, story } = await openStory(pkg, part, state);
  const items = new Map<number, T>();
  const collection: StoryCollection<T> = {
    part: story.part,
    rootName: story.rootName,
    rootAttrs: story.rootAttrs,
    relationships: story.relationships,
    source: story.source,
    items,
  };
  if (story.prolog !== undefined) collection.prolog = story.prolog;
  const context = kindOf(ctx, sc.name);
  cursor.enter();
  if (!sc.selfClosing) {
    let index = -1;
    const misc = (node: XmlComment | XmlPI) => {
      collection.misc = pushMisc(collection.misc, index + 1, node);
    };
    while (sc.next()) {
      if (sc.kind === 'text') continue;
      cursor.toTag(misc);
      if (sc.kind === 'close') {
        cursor.leave();
        break;
      }
      index++;
      const isItem = kindOf(ctx, sc.name) === itemKind;
      const attrs = sc.attrs();
      const idText = isItem ? attrOf(ctx, attrs, 'id') : undefined;
      if (idText === undefined || !ID.test(idText) || items.has(Number(idText))) {
        const raw = sc.readRaw();
        cursor.skip(raw);
        collection.extras = pushExtra(ctx, collection.extras, index, bounded(parseXml(raw).root), context, false);
        if (isItem) state.warnings.push(`${part}: a <${itemKind}> without a usable id was kept as it is`);
        continue;
      }
      const srcStart = cursor.pos;
      const blocks: Block[] = [];
      scanBlocks(ctx, sc, cursor, blocks);
      linkFields(blocks);
      const item = { id: Number(idText), blocks } as T;
      if (!cursor.lost) {
        item.srcStart = srcStart;
        item.srcEnd = cursor.pos;
      }
      const rest = attrs.filter(([name]) => kindOf(ctx, name) !== 'id');
      fill(item, ctx, rest);
      items.set(item.id, item);
    }
    finish(sc, cursor, collection);
  }
  return collection;
}

function fillNote(note: Note, ctx: PartContext, attrs: AttrList): void {
  const rest: AttrList = [];
  for (const pair of attrs) {
    if (kindOf(ctx, pair[0]) === 'type' && note.kind === undefined) note.kind = pair[1];
    else rest.push(pair);
  }
  if (note.kind === 'normal') delete note.kind;
  if (rest.length) note.attrs = rest;
}

function fillComment(comment: Comment, ctx: PartContext, attrs: AttrList): void {
  const rest: AttrList = [];
  for (const pair of attrs) {
    const kind = kindOf(ctx, pair[0]);
    if (kind === 'author' && comment.author === undefined) comment.author = pair[1];
    else if (kind === 'initials' && comment.initials === undefined) comment.initials = pair[1];
    else if (kind === 'date' && comment.date === undefined) comment.date = pair[1];
    else rest.push(pair);
  }
  if (rest.length) comment.attrs = rest;
}

/** Limits of reading a document, for hosts that want tighter ones than the defaults. */
export interface DocxReadOptions extends DocChartLimits {}

export async function readDocx(pkg: OpcPackage, options: DocxReadOptions = {}): Promise<DocModel> {
  const main = await pkg.mainPart();
  if (!main || !pkg.hasPart(main)) throw new Error('Invalid docx: the package has no main document part');
  const state = createReaderState();
  const partName = partNames(pkg);
  state.partName = partName;
  const rels = await pkg.relationships(main);

  const settingsTree = await readTree(pkg, supportingPart(pkg, rels, '/settings', state), state);
  const settings = parseSettings(settingsTree ? settingsTree.ctx : createContext(state, '', [], NO_RELATIONSHIPS), settingsTree?.doc.root);
  const themePart = supportingPart(pkg, rels, '/theme', state);
  const theme = parseTheme(themePart === undefined ? undefined : await pkg.readXml(themePart));
  const fontsTree = await readTree(pkg, supportingPart(pkg, rels, '/fontTable', state), state);
  const fonts = fontsTree ? parseFontTable(fontsTree.ctx, fontsTree.doc.root) : new Map();
  const stylesTree = await readTree(pkg, supportingPart(pkg, rels, '/styles', state), state);
  const styles = parseDocStyles(stylesTree ? stylesTree.ctx : createContext(state, '', [], NO_RELATIONSHIPS), stylesTree?.doc.root);
  const numberingTree = await readTree(pkg, supportingPart(pkg, rels, '/numbering', state), state);
  const numbering = parseNumbering(
    numberingTree ? numberingTree.ctx : createContext(state, '', [], NO_RELATIONSHIPS),
    numberingTree?.doc.root,
  );

  const body = await readBody(pkg, main, state);

  // Every header and footer the document part is related to is read, used by a section or not.
  const headers = new Map<string, Story>();
  const footers = new Map<string, Story>();
  const byPart = new Map<string, Story>();
  for (const rel of rels) {
    const target = rel.type.endsWith('/header') ? headers : rel.type.endsWith('/footer') ? footers : undefined;
    if (!target || rel.external || !rel.part) continue;
    // The part as the package spells it: ids that spell one part differently share its story, read once.
    const part = partName(rel.part);
    if (part === undefined) {
      state.warnings.push(`The ${target === headers ? 'header' : 'footer'} part ${quote(rel.part)} is missing`);
      continue;
    }
    let story = byPart.get(part);
    if (!story) byPart.set(part, (story = await readHeaderFooter(pkg, rel.part, state)));
    target.set(rel.id, story);
  }

  const model: DocModel = {
    body,
    styles,
    numbering,
    theme,
    settings,
    fonts,
    headers,
    footers,
    strict: false,
    warnings: state.warnings,
    unknown: state.unknown,
  };
  const footnotes = supportingPart(pkg, rels, '/footnotes', state);
  if (footnotes !== undefined) model.footnotes = await readCollection<Note>(pkg, footnotes, state, 'footnote', fillNote);
  const endnotes = supportingPart(pkg, rels, '/endnotes', state);
  if (endnotes !== undefined) model.endnotes = await readCollection<Note>(pkg, endnotes, state, 'endnote', fillNote);
  const comments = supportingPart(pkg, rels, '/comments', state);
  if (comments !== undefined) {
    model.comments = await readCollection<Comment>(pkg, comments, state, 'comment', fillComment);
    // Which comment answers which, and which is resolved: the part beside the comments says it.
    let extended: XmlDocument | undefined;
    const part = supportingPart(pkg, rels, '/commentsExtended', state);
    if (part !== undefined) {
      extended = parseSidePart(await pkg.readPart(part));
      if (!extended) state.warnings.push(`${quote(part)} is not well-formed: the replies and the resolved state of the comments were not read`);
      else bounded(extended.root);
    }
    linkComments(model.comments, extended ? extended.root : undefined);
  }
  // The charts the frames of every story name, from the values their parts cached.
  if (state.chartParts.size > 0) {
    const charts = await readDocCharts(pkg, state.chartParts, themeColorList(theme), state.warnings, options);
    if (charts) model.charts = charts;
  }
  model.strict = state.strict;
  return model;
}
