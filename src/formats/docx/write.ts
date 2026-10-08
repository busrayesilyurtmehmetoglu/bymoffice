// Writes the document model to docx.
//
// writeDocx() brings the package of an edited docx up to date with its model.
// Only the parts of stories that were modified are produced again, and of
// those only the blocks that were edited: everything else is copied from the
// source text of the part (see write-package.ts). Parts of stories that were
// not modified are not touched and keep their bytes.
//   - Relationships the model has and the package has not (a new hyperlink, a
//     new picture) are added; none is ever removed.
//   - Pictures the model carries (`DocModel.media`) and the package lacks are
//     added as parts, with their content type.
//   - Numbering definitions and styles the model has and their parts lack
//     are added to the text of those parts; nothing else in them changes. A
//     definition or style that exists in the part is not rewritten. The parts
//     are looked at when a story is written or the numbering says it was
//     modified.
//   - Of the settings only <w:trackRevisions> is written, when the model
//     tracks changes and the file does not or the other way round: the
//     element is put in at its place or taken out, and nothing else in the
//     part changes. A document without a settings part gets one.
//   - Theme, font table and the document properties are left alone. The
//     page and word counts in docProps/app.xml go stale with an edit; Word
//     computes them again when it opens the file.
//   - Writing does not change the model.
//
// The package remembers what was written into it. A part that an earlier
// call wrote is produced again by every later call, whatever the model says
// about being modified: after "edit, save, undo back past the save, save" the
// model is what was read and the package has to be too. Such a part comes out
// with the bytes the file had; a numbering part, relationship or picture that
// only the undone edits needed is taken out of the package again.
//
// An edit can leave one wrapper around content that is no longer in one
// piece (a paragraph split inside an inline content control). Each piece is
// written as an element of its own; for a content control the first keeps the
// id and the others are written without one, which the report says.
//
// buildDocx() makes a new package from a model alone: the way to save a
// document read from a binary doc or RTF file as docx. Everything is written
// from the model; what the model does not hold is not in the result.
//   - There is no theme part, so theme colours and theme fonts are written
//     as the plain values they resolve to.
//   - A hyperlink gets a relationship only when its target passes
//     safeHyperlinkTarget(); the text of any other link is written without
//     the link. (Field codes are text and are written as they are.)
//   - Footnotes and endnotes get the two separator notes Word expects when
//     the model has none.
//   - A shape that is not a picture keeps its place, size and text and is
//     written as a rectangle without outline and fill; the parts of a group
//     lie over its whole area. A picture without data is an empty frame.
//   - The report counts each of these in its warnings; characters XML cannot
//     hold are counted in `droppedCharacters`.

import {
  safeHyperlinkTarget,
  type Block,
  type Body,
  type Comment,
  type DocColor,
  type DocMedia,
  type DocModel,
  type DocRelationship,
  type ImageRef,
  type Note,
  type ParaProps,
  type RevisionInfo,
  type RowProps,
  type RunFonts,
  type RunItem,
  type RunProps,
  type Story,
  type StoryCollection,
  type Wrapper,
} from '../../model/doc/types';
import type { OpcPackage } from '../../opc';
import { parseXml, serializeXml, type XmlDocument, type XmlElement } from '../../xml';
import { ZipArchive } from '../../zip';
import { DEFAULT_AUTHOR, largestAnnotationId } from '../../ops/doc/track';
import { resolveColor, resolveFont } from './resolve';
import { parseOnOff } from './common';
import { decodePart } from './source';
import type { BodyContext } from './write-body';
import {
  STRICT_NS,
  TRANSITIONAL_NS,
  XML_DECLARATION,
  bodyText,
  collectionText,
  createBodyContext,
  docDefaultsText,
  encodePart,
  fontTableText,
  generatedRootAttrs,
  numberingTexts,
  scanPart,
  settingsText,
  storyText,
  styleElement,
  type StoryWriteOptions,
  type WriteCounters,
} from './write-package';
import { treeText } from './write-body';

/** What a write did; filled in when passed with the options. */
export interface DocxWriteReport {
  /** Parts written: replaced, added or put back. */
  parts: string[];
  /** Top-level blocks copied from the source text of their part. */
  copiedBlocks: number;
  /** Top-level blocks written from the model. */
  writtenBlocks: number;
  /** Characters XML cannot hold that were left out of the text. */
  droppedCharacters: number;
  warnings: string[];
}

export function emptyWriteReport(): DocxWriteReport {
  return { parts: [], copiedBlocks: 0, writtenBlocks: 0, droppedCharacters: 0, warnings: [] };
}

export interface WriteDocxOptions {
  /**
   * For tests: every story counts as modified and every block is written
   * from the model, whatever source span it carries.
   */
  regenerate?: boolean;
  report?: DocxWriteReport;
}

export interface BuildDocxOptions {
  /**
   * Where the pictures of a model that was read from a package are (the
   * package itself): such a model does not carry them. Pictures that cannot
   * be read are left out, and their drawings are written without an image.
   */
  pictures?: { readPart(part: string): Promise<Uint8Array> };
  report?: DocxWriteReport;
}

const RELS_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
const CONTENT_TYPES = {
  document: WML + 'document.main+xml',
  styles: WML + 'styles+xml',
  numbering: WML + 'numbering+xml',
  settings: WML + 'settings+xml',
  fontTable: WML + 'fontTable+xml',
  header: WML + 'header+xml',
  footer: WML + 'footer+xml',
  footnotes: WML + 'footnotes+xml',
  endnotes: WML + 'endnotes+xml',
  comments: WML + 'comments+xml',
};

const IMAGE_TYPES: { [extension: string]: string } = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  pict: 'image/pict',
  svg: 'image/svg+xml',
};

const relsPartOf = (part: string): string => `${part.slice(0, part.lastIndexOf('/'))}/_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;

/** The target to write for a part in the relationships of another: relative when it sits below the same folder. */
function relativeTarget(from: string, to: string): string {
  const dir = from.slice(0, from.lastIndexOf('/') + 1);
  return to.slice(0, dir.length) === dir ? to.slice(dir.length) : to;
}

function resolveTarget(source: string, target: string): string {
  const base = target.charAt(0) === '/' ? target : `${source.slice(0, source.lastIndexOf('/'))}/${target}`;
  const segments: string[] = [];
  for (const segment of base.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return `/${segments.join('/')}`;
}

const endsWith = (text: string, suffix: string): boolean => text.slice(-suffix.length) === suffix;

/** The first "rId<n>" that is not taken. */
function freeId(taken: (id: string) => boolean): string {
  for (let n = 1; ; n++) {
    const id = 'rId' + n;
    if (!taken(id)) return id;
  }
}

interface StoryEntry {
  kind: 'body' | 'header' | 'footer' | 'footnotes' | 'endnotes' | 'comments';
  story: Story | StoryCollection<Note> | StoryCollection<Comment>;
}

/** The stories of a document, each part once. */
function storiesOf(doc: DocModel): StoryEntry[] {
  const out: StoryEntry[] = [{ kind: 'body', story: doc.body }];
  const seen = new Set<object>();
  for (const [kind, map] of [
    ['header', doc.headers],
    ['footer', doc.footers],
  ] as const) {
    for (const story of map.values()) {
      if (seen.has(story)) continue;
      seen.add(story);
      out.push({ kind, story });
    }
  }
  if (doc.footnotes) out.push({ kind: 'footnotes', story: doc.footnotes });
  if (doc.endnotes) out.push({ kind: 'endnotes', story: doc.endnotes });
  if (doc.comments) out.push({ kind: 'comments', story: doc.comments });
  return out;
}

function textOf(ctx: BodyContext, entry: StoryEntry, options: StoryWriteOptions): string {
  switch (entry.kind) {
    case 'body':
      return bodyText(ctx, entry.story as Body, options);
    case 'header':
    case 'footer':
      return storyText(ctx, entry.story as Story, options);
    case 'footnotes':
      return collectionText(ctx, entry.story as StoryCollection<Note>, 'footnote', options);
    case 'endnotes':
      return collectionText(ctx, entry.story as StoryCollection<Note>, 'endnote', options);
    case 'comments':
      return collectionText(ctx, entry.story as StoryCollection<Comment>, 'comment', options);
  }
}

// ------------------------------------------------------------------ writeDocx

interface SaveState {
  /** The bytes each rewritten part had when the file was opened; null for a part the file did not have. */
  originals: Map<string, Uint8Array | null>;
  /** Parts the last write put into the package. A write that no longer produces one puts the original back. */
  written: Set<string>;
}

const SAVES = new WeakMap<OpcPackage, SaveState>();

/**
 * The bytes of a part with stretches of its text replaced. The places are
 * those of the text as decodePart() gives it (no byte order mark, line ends
 * as "\n"); the text that is changed is the one the file holds, in its own
 * encoding, so nothing but the stretches differs from the bytes given.
 */
function splicedPart(original: Uint8Array, splices: { start: number; end: number; text: string }[]): Uint8Array {
  const little = original[0] === 0xff && original[1] === 0xfe;
  const big = original[0] === 0xfe && original[1] === 0xff;
  const raw = new TextDecoder(little ? 'utf-16le' : big ? 'utf-16be' : 'utf-8', { ignoreBOM: true }).decode(original);
  // Where each place of the normalized text is in the raw one.
  const sorted = splices.slice().sort((a, b) => a.start - b.start);
  let out = '';
  let rawAt = raw.charCodeAt(0) === 0xfeff ? 1 : 0;
  let at = 0;
  let copied = 0;
  /** Moves on to a place of the normalized text and returns where that is in the raw text. */
  const seek = (target: number): number => {
    while (at < target && rawAt < raw.length) {
      // "\r\n" is one character of the normalized text.
      if (raw.charCodeAt(rawAt) === 13 && raw.charCodeAt(rawAt + 1) === 10) rawAt++;
      rawAt++;
      at++;
    }
    return rawAt;
  };
  for (const splice of sorted) {
    const start = seek(splice.start);
    const end = seek(splice.end);
    out += raw.slice(copied, start) + splice.text;
    copied = end;
  }
  out += raw.slice(copied);
  if (!little && !big) return new TextEncoder().encode(out);
  const bytes = new Uint8Array(out.length * 2);
  for (let i = 0; i < out.length; i++) {
    const c = out.charCodeAt(i);
    bytes[i * 2 + (little ? 0 : 1)] = c & 0xff;
    bytes[i * 2 + (little ? 1 : 0)] = c >> 8;
  }
  return bytes;
}

/** The children of <w:settings> the schema puts before <w:trackRevisions>. */
const BEFORE_TRACK_REVISIONS = new Set<string>([
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'removeDateAndTime',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'printPostScriptOverText',
  'printFractionalCharacterWidth',
  'printFormsData',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'saveFormsData',
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'stylePaneFormatFilter',
  'stylePaneSortMethod',
  'documentType',
  'mailMerge',
  'revisionView',
]);

/** The relationships part of a part, as the file had it, with what is added to it. */
class RelsEdit {
  readonly doc: XmlDocument;
  changed = false;
  private readonly ids = new Set<string>();

  constructor(original: Uint8Array | undefined) {
    this.doc = original
      ? parseXml(original)
      : { declaration: XML_DECLARATION.trim(), leading: [], root: { type: 'element', name: 'Relationships', attrs: new Map([['xmlns', NS_PACKAGE_RELS]]), children: [] }, trailing: [] };
    for (const rel of this.elements()) this.ids.add(rel.attrs.get('Id') || '');
  }

  elements(): XmlElement[] {
    return this.doc.root.children.filter((node): node is XmlElement => node.type === 'element' && node.name.replace(/^.*:/, '') === 'Relationship');
  }

  has(id: string): boolean {
    return this.ids.has(id);
  }

  /** Part an internal relationship of the given kind points to, from `source`. */
  partOf(source: string, suffix: string): string | undefined {
    const rel = this.elements().find((el) => endsWith(el.attrs.get('Type') || '', suffix) && el.attrs.get('TargetMode') !== 'External');
    return rel ? resolveTarget(source, rel.attrs.get('Target') || '') : undefined;
  }

  add(id: string, type: string, target: string, external: boolean): void {
    const root = this.doc.root;
    const prefix = root.name.indexOf(':') >= 0 ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
    const attrs = new Map<string, string>([
      ['Id', id],
      ['Type', type],
      ['Target', target],
    ]);
    if (external) attrs.set('TargetMode', 'External');
    root.children.push({ type: 'element', name: prefix + 'Relationship', attrs, children: [] });
    this.ids.add(id);
    this.changed = true;
  }

  freeId(): string {
    return freeId((id) => this.ids.has(id));
  }
}

/**
 * Brings the package up to date with the model: writes the parts of the
 * stories that were modified, and what the model has that the package lacks
 * (relationships, pictures, numbering definitions, styles). Does nothing when
 * nothing was modified.
 */
export async function writeDocx(pkg: OpcPackage, doc: DocModel, options: WriteDocxOptions = {}): Promise<void> {
  const regenerate = options.regenerate === true;
  const report = options.report;
  const main = doc.body.part;
  if (doc.body.rootName === '' || main.charAt(0) !== '/') {
    throw new Error('Cannot write the document into a package: it was not read from one (use buildDocx)');
  }
  const ns = doc.strict ? STRICT_NS : TRANSITIONAL_NS;
  const saved: SaveState = SAVES.get(pkg) || { originals: new Map(), written: new Set() };
  const { originals } = saved;
  // The first time a part is about to be rewritten, the package still holds it as the file had it.
  const originalOf = async (part: string): Promise<Uint8Array | undefined> => {
    const kept = originals.get(part);
    if (kept !== undefined) return kept === null ? undefined : kept;
    const bytes = pkg.hasPart(part) ? await pkg.readPart(part) : undefined;
    originals.set(part, bytes === undefined ? null : bytes);
    return bytes;
  };

  // Everything is produced before anything is put into the package: a failure leaves the package as it was.
  const writes = new Map<string, { data: Uint8Array; contentType?: string }>();
  const relsEdits = new Map<string, RelsEdit>();
  const relsFor = async (part: string): Promise<RelsEdit> => {
    let edit = relsEdits.get(part);
    if (!edit) relsEdits.set(part, (edit = new RelsEdit(await originalOf(relsPartOf(part)))));
    return edit;
  };
  const shared: WriteCounters = { dropped: { count: 0 }, drawingIds: { next: 100000 } };
  const counts = { copied: 0, written: 0 };
  let splitControls = 0;

  let stories = 0;
  for (const entry of storiesOf(doc)) {
    const story = entry.story;
    // A part that an earlier write put into the package is written again whatever the flag says: the
    // model may have been taken back to what was read (undo past the save), and the package has to follow.
    if (!regenerate && !story.modified && !saved.written.has(story.part)) continue;
    if (story.rootName === '' || story.part.charAt(0) !== '/') throw new Error(`Cannot write the document: the ${entry.kind} has no part in the package`);
    const original = await originalOf(story.part);
    stories++;
    const relationships = story.relationships;
    // Relationships made for this write: the model is not changed by writing it.
    const made = new Map<string, DocRelationship>();
    const ctx = createBodyContext(story.part, story.rootAttrs, relationships, shared);
    // A picture of a drawing that is written from nothing: the relationship the part has for it, or a new one.
    ctx.image = (ref: ImageRef): string | undefined => {
      const known = relationships.get(ref.rId);
      const targetOf = (rel: DocRelationship) => (rel.part === null ? rel.target : rel.part);
      if (known && (ref.target === undefined || targetOf(known) === ref.target)) return ref.rId;
      if (ref.target === undefined) return undefined;
      for (const rel of relationships.values()) if (endsWith(rel.type, '/image') && targetOf(rel) === ref.target) return rel.id;
      for (const rel of made.values()) if (targetOf(rel) === ref.target) return rel.id;
      const id = freeId((candidate) => relationships.has(candidate) || made.has(candidate));
      const external = ref.external === true;
      made.set(id, {
        id,
        type: ns.rel + 'image',
        target: external ? ref.target : relativeTarget(story.part, ref.target),
        external,
        part: external ? null : ref.target,
      });
      return id;
    };
    const text = textOf(ctx, entry, { regenerate, counts });
    if (ctx.splitControls.count) splitControls += ctx.splitControls.count;
    // The bytes of the file when the text is what was read from them (they may differ in a byte order mark or line ends).
    writes.set(story.part, { data: original && text === story.source ? original : encodePart(text) });
    // Relationships the model has and the part lacks.
    const rels = await relsFor(story.part);
    for (const rel of relationships.values()) {
      if (!rels.has(rel.id)) rels.add(rel.id, rel.type, rel.target, rel.external);
    }
    for (const rel of made.values()) rels.add(rel.id, rel.type, rel.target, rel.external);
  }

  // The one setting an edit changes: whether changes are tracked. The settings part is written when the model
  // says otherwise than the file does, and then nothing in it changes but <w:trackRevisions>.
  const dir = main.slice(0, main.lastIndexOf('/') + 1);
  let tracking: { part: string; data: Uint8Array; contentType?: string; relate: boolean } | undefined;
  {
    const wanted = doc.settings.trackRevisions === true;
    const rels = new RelsEdit(await originalOf(relsPartOf(main)));
    const existing = rels.partOf(main, '/settings');
    const part = existing || dir + 'settings.xml';
    const original = await originalOf(part);
    if (!original) {
      if (wanted) {
        const attrs = generatedRootAttrs(doc.strict);
        const ctx = createBodyContext(part, attrs, new Map(), shared);
        let text = XML_DECLARATION + `<${ctx.w}settings`;
        for (const [name, value] of attrs) text += ` ${name}="${value}"`;
        text += `><${ctx.w}trackRevisions/></${ctx.w}settings>`;
        tracking = { part, data: encodePart(text), contentType: CONTENT_TYPES.settings, relate: !existing };
      }
    } else {
      const source = decodePart(original);
      const scan = scanPart(source);
      const ctx = scan ? createBodyContext(part, scan.rootAttrs, new Map(), shared) : undefined;
      // The setting as the reader reads it: the last element that states a value counts.
      let stated = false;
      const marks: { start: number; end: number }[] = [];
      let place = scan ? scan.contentStart : 0;
      if (scan && ctx) {
        for (const child of scan.children) {
          if (child.kind === 'trackRevisions') {
            marks.push(child);
            let val: string | undefined;
            for (const [name, value] of child.attrs) {
              const colon = name.indexOf(':');
              if ((colon < 0 || name.slice(0, colon + 1) === ctx.w) && name.slice(colon + 1) === 'val') val = value;
            }
            const value = parseOnOff(val);
            if (value !== undefined) stated = value;
          } else if (BEFORE_TRACK_REVISIONS.has(child.kind)) place = child.end;
        }
      }
      if (!scan || !ctx) {
        // What the file says cannot be told: it is left as it is.
        if (wanted && report) report.warnings.push(`${part} could not be followed; the tracking of changes was not written to it`);
      } else if (stated !== wanted) {
        const element = `<${ctx.w}trackRevisions/>`;
        // Places in the text as the scanner read it; the bytes are changed at the same places of the text as it
        // was written, so that a byte order mark and the line ends of the file stay what they were.
        const splices: { start: number; end: number; text: string }[] = [];
        if (scan.selfClosing) splices.push({ start: scan.contentStart - 2, end: scan.contentStart, text: `>${element}</${scan.rootName}>` });
        else if (!marks.length) splices.push({ start: place, end: place, text: element });
        // Every statement of the setting goes; the new one takes the place of the first.
        else for (let i = 0; i < marks.length; i++) splices.push({ start: marks[i].start, end: marks[i].end, text: i === 0 && wanted ? element : '' });
        tracking = { part, data: splicedPart(original, splices), relate: false };
      }
    }
  }

  // Nothing to do for a document that was not edited and never written.
  if (!stories && !doc.numbering.modified && !saved.written.size && !tracking) return;

  const mainRels = await relsFor(main);
  if (tracking) {
    writes.set(tracking.part, { data: tracking.data, contentType: tracking.contentType });
    if (tracking.relate) mainRels.add(mainRels.freeId(), ns.rel + 'settings', relativeTarget(main, tracking.part), false);
  }

  const kindOf = (ctx: BodyContext, name: string): string => {
    // Attribute names of the part: the WordprocessingML prefix or none.
    const colon = name.indexOf(':');
    return colon < 0 || name.slice(0, colon + 1) === ctx.w ? name.slice(colon + 1) : name;
  };

  // Numbering definitions and styles the model has and their parts lack are added to the text of those parts.
  const additions = async (
    suffix: string,
    fileName: string,
    contentType: string,
    rootLocal: string,
    produce: (ctx: BodyContext, has: (kind: string, attribute: string, value: string) => boolean) => { first: string; second: string } | undefined,
    firstBefore: string[],
    secondBefore: string[],
  ): Promise<void> => {
    const existing = mainRels.partOf(main, suffix);
    const part = existing || dir + fileName;
    const original = await originalOf(part);
    if (!original) {
      // No such part in the file: a new one, when the model has anything to put into it.
      const attrs = generatedRootAttrs(doc.strict);
      const ctx = createBodyContext(part, attrs, new Map(), shared);
      const made = produce(ctx, () => false);
      if (!made) return;
      let text = XML_DECLARATION + `<${ctx.w}${rootLocal}`;
      for (const [name, value] of attrs) text += ` ${name}="${value}"`;
      text += `>${made.first}${made.second}</${ctx.w}${rootLocal}>`;
      writes.set(part, { data: encodePart(text), contentType });
      if (!existing) mainRels.add(mainRels.freeId(), ns.rel + suffix.slice(1), relativeTarget(main, part), false);
      return;
    }
    const source = decodePart(original);
    const scan = scanPart(source);
    if (!scan) {
      if (report) report.warnings.push(`${part} could not be followed; nothing was added to it`);
      return;
    }
    const ctx = createBodyContext(part, scan.rootAttrs, new Map(), shared);
    const made = produce(ctx, (kind, attribute, value) =>
      scan.children.some((child) => child.kind === kind && child.attrs.some(([name, text]) => text === value && kindOf(ctx, name) === attribute)),
    );
    if (!made) return;
    // The first kind goes before the first child that has to follow it, the second kind before the first that has to follow that.
    const end = scan.closeStart === undefined ? scan.contentStart : scan.closeStart;
    const before = (kinds: string[]): number => {
      for (const child of scan.children) if (kinds.indexOf(child.kind) >= 0) return child.start;
      return end;
    };
    const split = before(firstBefore);
    const tail = Math.max(split, before(secondBefore));
    let text: string;
    if (scan.selfClosing) text = `${source.slice(0, scan.contentStart - 2)}>${made.first}${made.second}</${scan.rootName}>${source.slice(scan.contentStart)}`;
    else text = source.slice(0, split) + made.first + source.slice(split, tail) + made.second + source.slice(tail);
    writes.set(part, { data: encodePart(text), contentType: pkg.contentTypeOf(part) === undefined ? contentType : undefined });
  };

  await additions(
    '/numbering',
    'numbering.xml',
    CONTENT_TYPES.numbering,
    'numbering',
    (ctx, has) => {
      const abstracts = new Set<number>();
      const nums = new Set<number>();
      for (const id of doc.numbering.abstracts.keys()) if (!has('abstractNum', 'abstractNumId', String(id))) abstracts.add(id);
      for (const id of doc.numbering.nums.keys()) if (!has('num', 'numId', String(id))) nums.add(id);
      if (!abstracts.size && !nums.size) return undefined;
      const texts = numberingTexts(ctx, doc.numbering, { abstracts, nums });
      return { first: texts.abstracts, second: texts.nums };
    },
    ['num', 'numIdMacAtCleanup'],
    ['numIdMacAtCleanup'],
  );
  await additions(
    '/styles',
    'styles.xml',
    CONTENT_TYPES.styles,
    'styles',
    (ctx, has) => {
      let text = '';
      for (const style of doc.styles.styles.values()) if (!has('style', 'styleId', style.id)) text += treeText(ctx, styleElement(ctx, style));
      return text ? { first: '', second: text } : undefined;
    },
    [],
    [],
  );

  // Pictures the model carries and the file did not have.
  if (doc.media) {
    for (const [target, media] of doc.media) {
      if (target.charAt(0) !== '/' || (await originalOf(target))) continue;
      writes.set(target, { data: media.bytes, contentType: pkg.contentTypeOf(target) === media.type ? undefined : media.type });
    }
  }

  for (const [part, edit] of relsEdits) {
    if (edit.changed) writes.set(relsPartOf(part), { data: new TextEncoder().encode(serializeXml(edit.doc)), contentType: RELS_TYPE });
  }

  if (!writes.size && !saved.written.size) return;
  // What an earlier write put into the package and this one does not goes back to what the file had.
  for (const part of saved.written) {
    if (writes.has(part)) continue;
    const original = originals.get(part);
    if (original) pkg.writePart(part, original);
    else pkg.deletePart(part);
    if (report) report.parts.push(part);
  }
  for (const [part, write] of writes) {
    pkg.writePart(part, write.data, pkg.hasPart(part) ? undefined : write.contentType);
    if (report) report.parts.push(part);
  }
  saved.written = new Set(writes.keys());
  SAVES.set(pkg, saved);
  if (report) {
    report.copiedBlocks += counts.copied;
    report.writtenBlocks += counts.written;
    report.droppedCharacters += shared.dropped.count;
    if (splitControls) report.warnings.push(`${splitControls} content control(s) were split by an edit; each further piece is written as a control of its own, without the id`);
  }
}

// ------------------------------------------------------------------ buildDocx

/** The relationships of one part of a new package. */
class NewRels {
  private readonly entries: { id: string; type: string; target: string; external: boolean }[] = [];
  private readonly ids = new Set<string>();
  private readonly keys = new Map<string, string>();

  has(id: string): boolean {
    return this.ids.has(id);
  }

  /** Adds a relationship; `id` is used when it is free and can be an XML id, else a new one is made up. */
  add(type: string, target: string, external: boolean, id?: string): string {
    const use = id !== undefined && /^[A-Za-z_][A-Za-z0-9_.-]{0,62}$/.test(id) && !this.ids.has(id) ? id : freeId((candidate) => this.ids.has(candidate));
    this.ids.add(use);
    this.entries.push({ id: use, type, target, external });
    return use;
  }

  /** The relationship made for `key` earlier, or a new one. */
  keyed(key: string, type: string, target: string, external: boolean): string {
    let id = this.keys.get(key);
    if (id === undefined) this.keys.set(key, (id = this.add(type, target, external)));
    return id;
  }

  get size(): number {
    return this.entries.length;
  }

  text(): string {
    const counts = { dropped: { count: 0 } };
    let out = `${XML_DECLARATION}<Relationships xmlns="${NS_PACKAGE_RELS}">`;
    for (const rel of this.entries) {
      out += `<Relationship Id="${rel.id}" Type="${attr(counts, rel.type)}" Target="${attr(counts, rel.target)}"${rel.external ? ' TargetMode="External"' : ''}/>`;
    }
    return out + '</Relationships>';
  }
}

function attr(counts: { dropped: { count: number } }, value: string): string {
  let out = '';
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c === 38) out += '&amp;';
    else if (c === 60) out += '&lt;';
    else if (c === 34) out += '&quot;';
    else if (c < 32 || (c >= 0xd800 && c <= 0xdfff) || c === 0xfffe || c === 0xffff) {
      // A valid pair of surrogates is kept; anything else XML cannot hold is left out.
      const next = value.charCodeAt(i + 1);
      if (c >= 0xd800 && c <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
        out += value.charAt(i) + value.charAt(i + 1);
        i++;
      } else counts.dropped.count++;
    } else out += value.charAt(i);
  }
  return out;
}

/** Relationship types of the parts a new package writes from the model; relationships to them are not carried over. */
const REBUILT = ['/styles', '/numbering', '/settings', '/fontTable', '/theme', '/header', '/footer', '/footnotes', '/endnotes', '/comments', '/webSettings'];

const SEPARATOR_PROPS = Object.freeze({ spacing: Object.freeze({ after: 0, line: 240, lineRule: 'auto' as const }) });

/** The notes with the two separators Word writes into every notes part, when the model has ordinary notes and lacks them. */
function withSeparators(collection: StoryCollection<Note> | undefined): StoryCollection<Note> | undefined {
  if (!collection) return undefined;
  let ordinary = false;
  const kinds = new Set<string>();
  for (const note of collection.items.values()) {
    if (note.kind === undefined) ordinary = true;
    else kinds.add(note.kind);
  }
  if (!ordinary) return collection;
  const items = new Map<number, Note>();
  const make = (id: number, kind: 'separator' | 'continuationSeparator') => {
    if (kinds.has(kind) || collection.items.has(id)) return;
    const block: Block = { type: 'paragraph', props: SEPARATOR_PROPS, children: [{ type: 'run', props: {}, content: [{ type: kind }] }] };
    items.set(id, { id, kind, blocks: [block] });
  };
  make(-1, 'separator');
  make(0, 'continuationSeparator');
  if (!items.size) return collection;
  for (const [id, note] of collection.items) items.set(id, note);
  return { ...collection, items };
}

/** The ids given to tracked changes that had none, and how many were written under the stand-in author. */
interface RevisionFill {
  next: number;
  defaulted: number;
}

/** The largest id a w:id may hold (a signed 32-bit number). */
const MAX_REVISION_ID = 0x7fffffff;

/** What is written for a tracked change the model holds without an id or an author: both, as the schema asks. */
function completeRevision(info: RevisionInfo, fill: RevisionFill): { id: number; author: string } {
  let id = info.id;
  if (id === undefined) {
    if (fill.next > MAX_REVISION_ID) throw new Error('Cannot write the document: it has no room for the ids of its tracked changes');
    id = fill.next++;
  }
  let author = info.author;
  if (author === undefined) {
    author = DEFAULT_AUTHOR;
    fill.defaulted++;
  }
  return { id, author };
}

const MARK_REVISIONS = ['inserted', 'deleted', 'moveFrom', 'moveTo'] as const;

/** Mark or row properties whose tracked changes all have an id and an author; the object itself when they do already. */
function completeProps<T extends RunProps | RowProps>(props: T, fill: RevisionFill): T {
  let out: T | undefined;
  for (const key of MARK_REVISIONS) {
    const info = (props as RunProps)[key];
    if (!info || (info.id !== undefined && info.author !== undefined)) continue;
    if (!out) out = { ...props };
    (out as RunProps)[key] = { ...info, ...completeRevision(info, fill) };
  }
  return out || props;
}

/** Tables in tables in text boxes that are followed; deeper ones are written as they are. */
const MAX_FILL_DEPTH = 64;

/**
 * The blocks with an id and an author on every tracked change of a paragraph
 * mark and of a table row (those of text boxes included). Property bags are
 * shared between paragraphs, and an id is not: each paragraph and row whose
 * change lacks one gets properties of its own. The list itself is returned
 * when nothing was missing.
 */
function completeMarks(blocks: Block[], fill: RevisionFill, depth: number): Block[] {
  if (depth > MAX_FILL_DEPTH) return blocks;
  let out: Block[] | undefined;
  const put = (index: number, block: Block): void => {
    if (!out) out = blocks.slice();
    out[index] = block;
  };
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      let props: ParaProps = block.props;
      const mark = props.markProps;
      if (mark) {
        const filled = completeProps(mark, fill);
        if (filled !== mark) props = { ...props, markProps: filled };
      }
      let children = block.children;
      for (let c = 0; c < block.children.length; c++) {
        const inline = block.children[c];
        if (inline.type !== 'run') continue;
        let content: RunItem[] | undefined;
        for (let k = 0; k < inline.content.length; k++) {
          const item = inline.content[k] as RunItem & { textBoxes?: Block[][] };
          if (!item.textBoxes) continue;
          let boxes: Block[][] | undefined;
          for (let b = 0; b < item.textBoxes.length; b++) {
            const inner = completeMarks(item.textBoxes[b], fill, depth + 1);
            if (inner === item.textBoxes[b]) continue;
            if (!boxes) boxes = item.textBoxes.slice();
            boxes[b] = inner;
          }
          if (!boxes) continue;
          if (!content) content = inline.content.slice();
          content[k] = { ...item, textBoxes: boxes } as RunItem;
        }
        if (!content) continue;
        if (children === block.children) children = block.children.slice();
        children[c] = { ...inline, content };
      }
      if (props !== block.props || children !== block.children) put(i, { ...block, props, children });
    } else if (block.type === 'table') {
      let rows = block.rows;
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        const props = completeProps(row.props, fill);
        let cells = row.cells;
        for (let c = 0; c < row.cells.length; c++) {
          const inner = completeMarks(row.cells[c].blocks, fill, depth + 1);
          if (inner === row.cells[c].blocks) continue;
          if (cells === row.cells) cells = row.cells.slice();
          cells[c] = { ...row.cells[c], blocks: inner };
        }
        if (props === row.props && cells === row.cells) continue;
        if (rows === block.rows) rows = block.rows.slice();
        rows[r] = { ...row, props, cells };
      }
      if (rows !== block.rows) put(i, { ...block, rows });
    }
  }
  return out || blocks;
}

/** A story with completeMarks() applied to its blocks; the story itself when nothing was missing. */
function completeStory<T extends { blocks: Block[] }>(story: T, fill: RevisionFill): T {
  const blocks = completeMarks(story.blocks, fill, 0);
  return blocks === story.blocks ? story : { ...story, blocks };
}

function completeCollection<T extends { blocks: Block[] }>(collection: StoryCollection<T> | undefined, fill: RevisionFill): StoryCollection<T> | undefined {
  if (!collection) return undefined;
  let items: Map<number, T> | undefined;
  for (const [id, item] of collection.items) {
    const filled = completeStory(item, fill);
    if (filled === item) continue;
    if (!items) items = new Map(collection.items);
    items.set(id, filled);
  }
  return items ? { ...collection, items } : collection;
}

/**
 * The model with an id and an author on every tracked change of a paragraph
 * mark and of a table row. A header that several sections share stays one
 * story. The model itself when nothing was missing (as for every model read
 * from a docx that Word wrote).
 */
function withCompleteMarks(doc: DocModel, fill: RevisionFill): DocModel {
  const body = completeStory(doc.body, fill);
  const shared = new Map<Story, Story>();
  const stories = (map: Map<string, Story>): Map<string, Story> => {
    let out: Map<string, Story> | undefined;
    for (const [id, story] of map) {
      let filled = shared.get(story);
      if (!filled) shared.set(story, (filled = completeStory(story, fill)));
      if (filled === story) continue;
      if (!out) out = new Map(map);
      out.set(id, filled);
    }
    return out || map;
  };
  const headers = stories(doc.headers);
  const footers = stories(doc.footers);
  const footnotes = completeCollection(doc.footnotes, fill);
  const endnotes = completeCollection(doc.endnotes, fill);
  const comments = completeCollection(doc.comments, fill);
  if (body === doc.body && headers === doc.headers && footers === doc.footers && footnotes === doc.footnotes && endnotes === doc.endnotes && comments === doc.comments) return doc;
  const out: DocModel = { ...doc, body, headers, footers };
  if (footnotes) out.footnotes = footnotes;
  if (endnotes) out.endnotes = endnotes;
  if (comments) out.comments = comments;
  return out;
}

/**
 * A new docx package with the content of the model. Meant for models read
 * from binary doc and RTF files, which have no package; a model read from a
 * docx file can be given too, with `options.pictures` for its images.
 */
export async function buildDocx(model: DocModel, options: BuildDocxOptions = {}): Promise<Uint8Array> {
  const report = options.report;
  const warn = (text: string) => {
    if (report) report.warnings.push(text);
  };
  // Tracked changes of a doc or RTF file have no ids, and not always an author; the schema asks for both. The
  // ids go on from the largest one the model has, first to the marks of paragraphs and rows, then to the
  // elements around runs as they are written.
  const revisions: RevisionFill = { next: largestAnnotationId(model) + 1, defaulted: 0 };
  const doc = withCompleteMarks(model, revisions);
  const strict = doc.strict;
  const ns = strict ? STRICT_NS : TRANSITIONAL_NS;
  const rootAttrs = generatedRootAttrs(strict);
  const shared: WriteCounters = { dropped: { count: 0 }, drawingIds: { next: 1 } };
  const counts = { copied: 0, written: 0 };
  // What the new file does not carry over; the report says how much of each.
  const lost = { shapes: 0, geometries: 0, groups: 0, pictures: 0, fills: 0, opacities: 0, colors: 0, links: 0, theme: false };
  const parts: [name: string, data: Uint8Array][] = [];
  const defaults = new Map<string, string>([
    ['rels', RELS_TYPE],
    ['xml', 'application/xml'],
  ]);
  const overrides: [part: string, type: string][] = [];

  const footnotes = withSeparators(doc.footnotes);
  const endnotes = withSeparators(doc.endnotes);
  const entries = storiesOf(footnotes === doc.footnotes && endnotes === doc.endnotes ? doc : { ...doc, footnotes, endnotes });

  // The pictures there are: those of the model, and those the source of a packaged model gives.
  const available = new Map<string, DocMedia>(doc.media || []);
  if (options.pictures) {
    for (const entry of entries) {
      for (const rel of entry.story.relationships.values()) {
        if (rel.external || rel.part === null || !endsWith(rel.type, '/image') || available.has(rel.part)) continue;
        try {
          const extension = rel.part.slice(rel.part.lastIndexOf('.') + 1).toLowerCase();
          available.set(rel.part, { bytes: await options.pictures.readPart(rel.part), type: IMAGE_TYPES[extension] || 'application/octet-stream' });
        } catch (error) {
          warn('A picture could not be read and was left out');
        }
      }
    }
  }
  const mediaParts = new Map<string, string>();
  const mediaNames = new Set<string>();
  const mediaPart = (target: string): string | undefined => {
    const known = mediaParts.get(target);
    if (known !== undefined) return known;
    const media = available.get(target);
    if (!media) return undefined;
    // The name the picture had, when it is a plain one that is still free.
    let name = target.slice(target.lastIndexOf('/') + 1);
    let extension = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
    if (!/^[A-Za-z0-9_-]{1,40}\.[A-Za-z0-9]{1,8}$/.test(name) || mediaNames.has(name.toLowerCase())) {
      if (!/^[a-z0-9]{1,8}$/.test(extension) || name.indexOf('.') < 0) extension = 'bin';
      for (let n = mediaNames.size + 1; ; n++) {
        name = `image${n}.${extension}`;
        if (!mediaNames.has(name)) break;
      }
    }
    mediaNames.add(name.toLowerCase());
    const part = '/word/media/' + name;
    mediaParts.set(target, part);
    parts.push([part.slice(1), media.bytes]);
    const declared = defaults.get(extension);
    if (declared === undefined) defaults.set(extension, media.type);
    else if (declared !== media.type) overrides.push([part, media.type]);
    return part;
  };

  // A package without a theme: theme colours and fonts as the plain values they stand for.
  const color = (value: DocColor): DocColor => {
    if (value.theme === undefined) return value;
    lost.theme = true;
    const resolved = resolveColor(doc, value);
    if (resolved !== undefined) return { rgb: resolved.slice(1).toUpperCase() };
    return value.auto ? { auto: true } : value;
  };
  const fonts = (value: RunFonts): RunFonts => {
    if (value.asciiTheme === undefined && value.hAnsiTheme === undefined && value.eastAsiaTheme === undefined && value.csTheme === undefined) return value;
    lost.theme = true;
    const out: RunFonts = {};
    for (const slot of ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const) {
      const name = resolveFont(doc, value, slot);
      if (name !== undefined) out[slot] = name;
    }
    if (value.hint !== undefined) out.hint = value.hint;
    return out;
  };

  const partNames = new Map<StoryEntry['story'], string>();
  const relsOf = new Map<StoryEntry['story'], NewRels>();
  let headers = 0;
  let footers = 0;
  for (const entry of entries) {
    const name =
      entry.kind === 'body' ? '/word/document.xml' : entry.kind === 'header' ? `/word/header${++headers}.xml` : entry.kind === 'footer' ? `/word/footer${++footers}.xml` : `/word/${entry.kind}.xml`;
    partNames.set(entry.story, name);
    const rels = new NewRels();
    relsOf.set(entry.story, rels);
    // A story read from a package: the trees kept from it name relationships by id, so those keep their ids.
    if (entry.story.rootName === '') continue;
    for (const rel of entry.story.relationships.values()) {
      if (rel.external) {
        if (endsWith(rel.type, '/hyperlink') && safeHyperlinkTarget(rel.target) === undefined) continue;
        rels.add(rel.type, rel.target, true, rel.id);
      } else if (endsWith(rel.type, '/image') && rel.part !== null) {
        const part = mediaPart(rel.part);
        if (part !== undefined) rels.add(rel.type, relativeTarget(name, part), false, rel.id);
        else warn('A picture is not available and was left out');
      } else if (!REBUILT.some((suffix) => endsWith(rel.type, suffix))) warn(`A related part of kind "${rel.type.slice(rel.type.lastIndexOf('/') + 1)}" was not carried over`);
    }
  }

  const bodyRels = relsOf.get(doc.body) as NewRels;
  // Headers and footers by the ids section properties use.
  const headerIds = new Map<string, string>();
  for (const [list, map] of [
    ['headers', doc.headers],
    ['footers', doc.footers],
  ] as const) {
    for (const [id, story] of map) {
      const part = partNames.get(story);
      if (part === undefined) continue;
      headerIds.set(list + '\n' + id, bodyRels.add(ns.rel + (list === 'headers' ? 'header' : 'footer'), relativeTarget('/word/document.xml', part), false, id));
    }
  }

  const contextFor = (entry: StoryEntry): BodyContext => {
    const story = entry.story;
    const part = partNames.get(story) as string;
    const rels = relsOf.get(story) as NewRels;
    const ctx = createBodyContext(part, story.rootName === '' ? rootAttrs : story.rootAttrs, story.relationships, shared);
    ctx.color = color;
    ctx.fonts = fonts;
    ctx.generated = lost;
    ctx.revision = (info) => completeRevision(info, revisions);
    ctx.headerRef = (list, id) => headerIds.get(list + '\n' + id);
    ctx.image = (ref: ImageRef): string | undefined => {
      if (story.rootName !== '' && rels.has(ref.rId)) return ref.rId;
      if (ref.target !== undefined && ref.external) return rels.keyed('external\n' + ref.target, ns.rel + 'image', ref.target, true);
      const media = ref.target === undefined ? undefined : mediaPart(ref.target);
      // A picture whose bytes the model does not have: its frame is written, empty.
      if (media === undefined) {
        lost.pictures++;
        return undefined;
      }
      return rels.keyed('image\n' + media, ns.rel + 'image', relativeTarget(part, media), false);
    };
    ctx.link = (wrapper: Extract<Wrapper, { kind: 'hyperlink' }>) => {
      if (story.rootName !== '' && wrapper.rId !== undefined && rels.has(wrapper.rId)) return { rId: wrapper.rId, anchor: wrapper.anchor };
      const safe = safeHyperlinkTarget(wrapper.target);
      if (safe !== undefined && safe.charAt(0) !== '#') return { rId: rels.keyed('link\n' + safe, ns.rel + 'hyperlink', safe, true), anchor: wrapper.anchor };
      if (safe !== undefined) return { anchor: wrapper.anchor !== undefined ? wrapper.anchor : safe.slice(1) };
      if (wrapper.anchor !== undefined) return { anchor: wrapper.anchor };
      // A link that leads nowhere the engine would follow: its text, without the link.
      if (wrapper.target === undefined && wrapper.rId === undefined) return {};
      lost.links++;
      return null;
    };
    return ctx;
  };

  const ROOTS: { [kind: string]: string } = { body: 'document', header: 'hdr', footer: 'ftr', footnotes: 'footnotes', endnotes: 'endnotes', comments: 'comments' };
  const storyParts: [name: string, data: Uint8Array][] = [];
  for (const entry of entries) {
    const ctx = contextFor(entry);
    const part = partNames.get(entry.story) as string;
    const text = textOf(ctx, entry, { regenerate: true, counts, root: { name: ctx.w + ROOTS[entry.kind], attrs: rootAttrs } });
    storyParts.push([part.slice(1), encodePart(text)]);
    const kind = entry.kind === 'body' ? 'document' : entry.kind;
    overrides.push([part, CONTENT_TYPES[kind]]);
    if (entry.kind !== 'body' && entry.kind !== 'header' && entry.kind !== 'footer') bodyRels.add(ns.rel + entry.kind, relativeTarget('/word/document.xml', part), false);
  }

  // The supporting parts, all written from the model.
  const support = (name: string, rootLocal: string, kind: 'styles' | 'numbering' | 'settings' | 'fontTable', inner: (ctx: BodyContext) => string) => {
    const part = '/word/' + name;
    const ctx = createBodyContext(part, rootAttrs, new Map(), shared);
    ctx.color = color;
    ctx.fonts = fonts;
    let text = XML_DECLARATION + `<${ctx.w}${rootLocal}`;
    for (const [key, value] of rootAttrs) text += ` ${key}="${value}"`;
    text += `>${inner(ctx)}</${ctx.w}${rootLocal}>`;
    storyParts.push([part.slice(1), encodePart(text)]);
    overrides.push([part, CONTENT_TYPES[kind]]);
    bodyRels.add(ns.rel + kind, name, false);
  };
  support('styles.xml', 'styles', 'styles', (ctx) => {
    let text = docDefaultsText(ctx, doc.styles);
    for (const style of doc.styles.styles.values()) text += treeText(ctx, styleElement(ctx, style));
    return text;
  });
  // The password of a protection is only carried over with the element a docx was read with.
  if (doc.settings.protection && doc.settings.protection.enforced && !doc.settings.protection.xml) warn('The document protection is carried over without its password');
  support('settings.xml', 'settings', 'settings', (ctx) => settingsText(ctx, footnotes === doc.footnotes && endnotes === doc.endnotes ? doc : { ...doc, footnotes, endnotes }));
  if (doc.fonts.size) support('fontTable.xml', 'fonts', 'fontTable', (ctx) => fontTableText(ctx, doc.fonts));
  if (doc.numbering.abstracts.size || doc.numbering.nums.size) {
    support('numbering.xml', 'numbering', 'numbering', (ctx) => {
      const texts = numberingTexts(ctx, doc.numbering);
      return texts.abstracts + texts.nums;
    });
  }

  let types = `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`;
  for (const [extension, type] of defaults) types += `<Default Extension="${extension}" ContentType="${attr(shared, type)}"/>`;
  for (const [part, type] of overrides) types += `<Override PartName="${attr(shared, part)}" ContentType="${attr(shared, type)}"/>`;
  types += '</Types>';

  const encoder = new TextEncoder();
  const zip = new ZipArchive();
  zip.set('[Content_Types].xml', encoder.encode(types));
  zip.set(
    '_rels/.rels',
    encoder.encode(`${XML_DECLARATION}<Relationships xmlns="${NS_PACKAGE_RELS}"><Relationship Id="rId1" Type="${ns.rel}officeDocument" Target="word/document.xml"/></Relationships>`),
  );
  for (const [name, data] of storyParts) zip.set(name, data);
  for (const entry of entries) {
    const rels = relsOf.get(entry.story) as NewRels;
    if (rels.size) zip.set(relsPartOf(partNames.get(entry.story) as string).slice(1), encoder.encode(rels.text()));
  }
  for (const [name, data] of parts) zip.set(name, data);

  if (report) {
    report.writtenBlocks += counts.written;
    report.droppedCharacters += shared.dropped.count;
    if (lost.shapes) warn(`${lost.shapes} shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over`);
    if (lost.geometries) warn(`${lost.geometries} shape(s) were written as rectangles: their geometry was not carried over`);
    if (lost.fills) warn(`${lost.fills} shape(s) with a fill that is not a plain colour (a pattern, a picture, a gradient) were written without a fill`);
    if (lost.opacities) warn(`${lost.opacities} shape(s) with a partly transparent fill were written with an opaque one`);
    if (lost.colors) warn(`${lost.colors} shape(s) with a colour that is an index into a palette or a system colour were written without that colour: no fill, a black outline`);
    if (lost.groups) warn(`${lost.groups} group(s) of shapes were written with every part over the whole area of the group, without outlines and fills`);
    if (lost.pictures) warn(`${lost.pictures} picture(s) whose data is not available were written as empty frames`);
    if (lost.links) warn(`${lost.links} link(s) with an address that is not followed were written as plain text`);
    if (revisions.defaulted) warn(`${revisions.defaulted} tracked change(s) were carried over without their authors: they are written under the name "${DEFAULT_AUTHOR}"`);
    if (lost.theme) warn('The theme was not carried over: theme colours and theme fonts were written as the plain values they stand for');
    for (const name of zip.names()) report.parts.push('/' + name);
  }
  return zip.write();
}
