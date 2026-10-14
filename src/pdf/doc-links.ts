// The links and the outline of the PDF of a document (docx, doc, RTF), from
// its laid out pages and its model:
//
// - a hyperlink is where its text is: the text items of a page that carry
//   the link (TextItem.link, already passed through safeHyperlinkTarget), one
//   rectangle per piece of a line, the pieces of a link that runs over lines
//   together in one link;
// - a link into the document ("#name": w:hyperlink/@w:anchor, HYPERLINK \l)
//   goes to the place of the bookmark of that name: the line of the body that
//   holds its start. A bookmark that is not there, or not on any page, leaves
//   the link out, counted;
// - the outline is made of the headings of the body, as Word makes the
//   bookmarks of a PDF of them: a paragraph with an outline level (its own,
//   or that of its style or of a style it is based on), or one in a style
//   named "heading 1" to "heading 9" without one. Level 9 is body text.
//
// Links whose target is of a kind no link may have (javascript:, file:,
// data:, relative ones) never reach the pages: the layout leaves them out.
// They are counted here, from the model.

import type { LayoutFont, LayoutPage, LineMark, TextItem } from '../layout/doc/types';
import { parseFieldCode } from '../layout/doc/symbols';
import { paragraphText, safeHyperlinkTarget, type Block, type DocModel, type Inline, type Paragraph, type Wrapper } from '../model/doc/types';
import { resolveRunProps } from '../formats/docx/resolve';
import { MAX_OUTLINE_ENTRIES, MAX_OUTLINE_TITLE, MAX_PAGE_LINKS, outlineTitle, type PdfLink, type PdfLinkRect, type PdfOutlineEntry } from './links';

export interface DocPdfLinkOptions {
  /** Points per px of the pages: 0.75 for pages laid out at 96 px per inch. */
  scale: number;
  /** The links of the pages; true when missing. */
  links?: boolean;
  /** The outline of the headings; true when missing. */
  outline?: boolean;
  /** How far text rises above and falls below its baseline, in px; from the size of the font when missing. */
  metrics?(font: LayoutFont): { ascent: number; descent: number };
}

export interface DocPdfLinks {
  /** The links of every page, by its index: points from the top left corner of the page. */
  pages: PdfLink[][];
  outline: PdfOutlineEntry[];
  /** Links whose target is of a kind a link may not have; they are on no page. */
  refused: number;
  /** Links into the document whose bookmark was not found; they are on no page. */
  unresolved: number;
  /** What was left out beyond the limits. */
  dropped: number;
  /** Links in turned text: they get no link (one per stretch of a line). */
  rotated: number;
  /** Headings that are not in the outline: their title is empty or hidden, or they are on no page. */
  outlineSkipped: number;
  /** Blocks that were not looked at, beyond MAX_WALK_BLOCKS. */
  unread: number;
}

/** Blocks the model is walked through, all stories together. */
const MAX_WALK_BLOCKS = 2000000;
/** Tables in tables. */
const MAX_WALK_DEPTH = 64;
/** Bookmarks that are looked for. */
const MAX_BOOKMARKS = 100000;
/** Steps along a chain of styles: a chain that goes round is ended. */
const MAX_STYLE_CHAIN = 32;
const HEADING_STYLE = /^heading\s*([1-9])$/i;

/** A place in the body the pages are searched for: a paragraph and the offset of a character in it. */
interface Spot {
  offset: number;
  /** Where it is, as far as the pages were searched: the index of its page and the point there, in px. */
  page?: number;
  x?: number;
  y?: number;
  /** The line that holds the offset was found; until then the place is the last line of the paragraph seen. */
  found?: boolean;
  /** The pages are searched for it. */
  asked?: boolean;
}

interface Heading {
  title: string;
  level: number;
  spot: Spot;
}

/** The outline level of a paragraph: its own, of its style or of a style it is based on, or that of a heading style; undefined for body text. */
function outlineLevelOf(model: DocModel, paragraph: Paragraph, cache: Map<string, number | undefined>): number | undefined {
  const own = paragraph.props ? paragraph.props.outlineLevel : undefined;
  if (typeof own === 'number') return own >= 0 && own <= 8 ? Math.floor(own) : undefined;
  const styles = model.styles;
  const first = (paragraph.props && paragraph.props.style) || (styles && styles.defaultIds ? styles.defaultIds.paragraph : undefined);
  if (first === undefined || !styles || !styles.styles) return undefined;
  if (cache.has(first)) return cache.get(first);
  let level: number | undefined;
  let id: string | undefined = first;
  for (let step = 0; id !== undefined && step < MAX_STYLE_CHAIN; step++) {
    const style = styles.styles.get(id);
    if (!style) break;
    const stated = style.para ? style.para.outlineLevel : undefined;
    if (typeof stated === 'number') {
      level = stated >= 0 && stated <= 8 ? Math.floor(stated) : undefined;
      break;
    }
    // A built-in heading style is a heading of its level, whether the file says so or not.
    const named = style.name !== undefined ? HEADING_STYLE.exec(style.name.trim()) : null;
    if (named) {
      level = Number(named[1]) - 1;
      break;
    }
    id = style.basedOn;
  }
  cache.set(first, level);
  return level;
}

/** The target of a field whose code is a HYPERLINK, as it stands in the file; undefined for any other field. */
function hyperlinkFieldTarget(instruction: string | undefined): { target: string | undefined } | undefined {
  if (typeof instruction !== 'string') return undefined;
  const code = parseFieldCode(instruction.slice(0, 4096));
  return code.name === 'HYPERLINK' ? { target: code.target } : undefined;
}

/**
 * The text of a heading as the outline shows it: the text of the paragraph
 * without its hidden runs (w:vanish, of the run or of its styles), which Word
 * leaves out of the title of a heading.
 */
function headingText(model: DocModel, paragraph: Paragraph): string {
  let out = '';
  const children = Array.isArray(paragraph.children) ? paragraph.children : [];
  for (let i = 0; i < children.length; i++) {
    const inline = children[i];
    if (!inline || (inline.type !== 'run' && inline.type !== 'math')) continue;
    if (inline.type === 'run') {
      let hidden = inline.props && inline.props.vanish === true;
      if (!hidden && model.styles) {
        try {
          hidden = resolveRunProps(model, paragraph.props || {}, inline.props || {}).vanish === true;
        } catch {
          // A run whose styles cannot be resolved is taken as it stands.
        }
      }
      if (hidden) continue;
    }
    out += paragraphText({ type: 'paragraph', props: paragraph.props, children: [inline] } as Paragraph);
    if (out.length > MAX_OUTLINE_TITLE * 4) break;
  }
  return out;
}

/**
 * What the model holds that the pages need: the bookmarks of the body (by
 * their name, where they start), its headings, and the number of links whose
 * target no link may have, in every story.
 */
function readModel(model: DocModel, wantOutline: boolean, wantLinks: boolean): { bookmarks: Map<string, { key: string; spot: Spot }>; headings: { key: string; heading: Heading }[]; refused: number; dropped: number; skipped: number; unread: number } {
  const bookmarks = new Map<string, { key: string; spot: Spot }>();
  const headings: { key: string; heading: Heading }[] = [];
  const seenLinks = new Set<unknown>();
  const levels = new Map<string, number | undefined>();
  let refused = 0;
  let dropped = 0;
  let skipped = 0;
  let unread = 0;
  let budget = MAX_WALK_BLOCKS;
  const refuse = (key: unknown, target: string | undefined) => {
    if (seenLinks.has(key)) return;
    seenLinks.add(key);
    if (target === undefined || safeHyperlinkTarget(target) === undefined) refused++;
  };
  const linksOf = (inline: Inline) => {
    for (let w: Wrapper | undefined = inline.wrap; w; w = w.parent) {
      if (w.kind === 'hyperlink') {
        // A link with neither a target nor a bookmark is no link: the layout shows it as text.
        if (w.target !== undefined || w.anchor !== undefined) refuse(w, w.anchor !== undefined ? (w.target === undefined ? '' : w.target) + '#' + w.anchor : w.target);
      } else if (w.kind === 'simpleField') {
        const field = hyperlinkFieldTarget(w.instruction);
        if (field) refuse(w, field.target);
      }
    }
    if (inline.type !== 'run') return;
    for (let i = 0; i < inline.content.length; i++) {
      const item = inline.content[i];
      if (item.type !== 'fieldChar' || item.charType !== 'begin' || !item.field) continue;
      const field = hyperlinkFieldTarget(item.field.instruction);
      if (field) refuse(item.field, field.target);
    }
  };
  const walk = (story: string, blocks: readonly Block[] | undefined, body: boolean) => {
    if (!Array.isArray(blocks)) return;
    if (budget < 0) {
      unread += blocks.length;
      return;
    }
    // Bookmarks between blocks start at the paragraph that comes next.
    let waiting: string[] = [];
    const stack: { blocks: readonly Block[]; at: number; path: number[] }[] = [{ blocks, at: 0, path: [] }];
    while (stack.length > 0) {
      const top = stack[stack.length - 1];
      if (top.at >= top.blocks.length) {
        stack.pop();
        continue;
      }
      if (--budget < 0) {
        // What is left of this story is not looked at (the block at `at` among it): counted, at the depth it was reached.
        for (let s = 0; s < stack.length; s++) unread += stack[s].blocks.length - stack[s].at;
        return;
      }
      const index = top.at++;
      const block = top.blocks[index];
      if (!block) continue;
      const path = top.path.concat([index]);
      if (block.type === 'mark') {
        if (body && block.mark && block.mark.type === 'bookmarkStart' && typeof block.mark.name === 'string' && waiting.length < 1000) waiting.push(block.mark.name);
        continue;
      }
      if (block.type === 'table') {
        if (stack.length >= MAX_WALK_DEPTH) {
          unread++;
          continue;
        }
        // The cells of the table, in reading order, each as a list of blocks: pushed last first.
        const rows = Array.isArray(block.rows) ? block.rows : [];
        for (let r = rows.length - 1; r >= 0; r--) {
          const cells = rows[r] && Array.isArray(rows[r].cells) ? rows[r].cells : [];
          for (let c = cells.length - 1; c >= 0; c--) if (cells[c] && Array.isArray(cells[c].blocks)) stack.push({ blocks: cells[c].blocks, at: 0, path: path.concat([r, c]) });
        }
        continue;
      }
      if (block.type !== 'paragraph') continue;
      const key = story + ':' + path.join(',');
      if (body) {
        for (let i = 0; i < waiting.length; i++) if (!bookmarks.has(waiting[i]) && bookmarks.size < MAX_BOOKMARKS) bookmarks.set(waiting[i], { key, spot: { offset: 0 } });
        waiting = [];
      }
      let offset = 0;
      const children = Array.isArray(block.children) ? block.children : [];
      for (let i = 0; i < children.length; i++) {
        const inline = children[i];
        if (!inline) continue;
        if (wantLinks) linksOf(inline);
        if (inline.type === 'bookmarkStart') {
          if (body && typeof inline.name === 'string' && !bookmarks.has(inline.name)) {
            if (bookmarks.size < MAX_BOOKMARKS) bookmarks.set(inline.name, { key, spot: { offset } });
          }
          continue;
        }
        if (inline.type === 'run' || inline.type === 'math') offset += paragraphText({ type: 'paragraph', props: block.props, children: [inline] } as Paragraph).length;
      }
      if (body && wantOutline) {
        const level = outlineLevelOf(model, block, levels);
        if (level !== undefined) {
          const title = outlineTitle(headingText(model, block));
          if (title === '') skipped++;
          else if (headings.length < MAX_OUTLINE_ENTRIES) headings.push({ key, heading: { title, level, spot: { offset: 0 } } });
          else dropped++;
        }
      }
    }
  };
  walk('body', model.body ? model.body.blocks : undefined, true);
  if (wantLinks) {
    const stories: [string, Map<string, { blocks: Block[] }> | undefined][] = [
      ['header', model.headers],
      ['footer', model.footers],
    ];
    for (let s = 0; s < stories.length; s++) {
      const map = stories[s][1];
      if (map) map.forEach((story, id) => walk(stories[s][0] + ':' + id, story.blocks, false));
    }
    const notes: [string, { items: Map<number, { blocks: Block[] }> } | undefined][] = [
      ['footnote', model.footnotes],
      ['endnote', model.endnotes],
    ];
    for (let s = 0; s < notes.length; s++) {
      const collection = notes[s][1];
      if (collection && collection.items) collection.items.forEach((note, id) => walk(notes[s][0] + ':' + id, note.blocks, false));
    }
  }
  return { bookmarks, headings, refused, dropped, skipped, unread };
}

/** The key of the paragraph a text item or a line comes from, in the body only. */
function bodyKey(source: { story: string; path: number[] } | undefined): string | undefined {
  if (!source || source.story !== 'body' || !Array.isArray(source.path)) return undefined;
  return 'body:' + source.path.join(',');
}

/**
 * The links of every page of a laid out document and the outline of its
 * headings, in points (see DocPdfLinks). Pages that are not laid out yet are
 * not looked at: call it with all of them.
 */
export function docPdfLinks(pages: readonly LayoutPage[], model: DocModel, options: DocPdfLinkOptions): DocPdfLinks {
  const scale = options.scale > 0 && isFinite(options.scale) ? options.scale : 0.75;
  const wantLinks = options.links !== false;
  const wantOutline = options.outline !== false;
  const out: DocPdfLinks = { pages: [], outline: [], refused: 0, unresolved: 0, dropped: 0, rotated: 0, outlineSkipped: 0, unread: 0 };
  for (let i = 0; i < pages.length; i++) out.pages.push([]);
  if (!wantLinks && !wantOutline) return out;
  const read = readModel(model, wantOutline, wantLinks);
  out.refused = read.refused;
  out.dropped = read.dropped;
  out.outlineSkipped = read.skipped;
  out.unread = read.unread;

  // The places to find on the pages: of the bookmarks that links go to, and of the headings.
  const wanted = new Map<string, Spot[]>();
  const want = (key: string, spot: Spot) => {
    const list = wanted.get(key);
    if (list) list.push(spot);
    else wanted.set(key, [spot]);
  };
  const targets = new Map<string, Spot | undefined>();
  const lower = new Map<string, string>();
  read.bookmarks.forEach((_entry, bookmark) => {
    const folded = bookmark.toLowerCase();
    if (!lower.has(folded)) lower.set(folded, bookmark);
  });
  /** The place of the bookmark a link goes to; Word finds bookmarks whatever the case of their letters. */
  const bookmarkSpot = (bookmark: string): Spot | undefined => {
    if (targets.has(bookmark)) return targets.get(bookmark);
    const entry = read.bookmarks.get(bookmark) || read.bookmarks.get(lower.get(bookmark.toLowerCase()) || '');
    const spot = entry ? entry.spot : undefined;
    if (entry && spot && !spot.asked) {
      spot.asked = true;
      want(entry.key, spot);
    }
    targets.set(bookmark, spot);
    return spot;
  };
  const metrics = (font: LayoutFont): { ascent: number; descent: number } => {
    if (options.metrics) {
      try {
        const got = options.metrics(font);
        if (got && got.ascent >= 0 && got.descent >= 0 && isFinite(got.ascent + got.descent)) return got;
      } catch {
        // The size of the font tells as much.
      }
    }
    const size = font && font.size > 0 && isFinite(font.size) ? font.size : 12;
    return { ascent: size * 0.9, descent: size * 0.25 };
  };

  // The links of each page, with the bookmarks they need.
  const internal: { link: PdfLink; spot: Spot }[] = [];
  if (wantLinks) {
    for (let p = 0; p < pages.length; p++) {
      const items = pages[p] && Array.isArray(pages[p].items) ? pages[p].items : [];
      const list = out.pages[p];
      let current: { target: string; link: PdfLink | undefined; rects: PdfLinkRect[]; last: TextItem; right: number } | undefined;
      /** The target of the turned stretch of text before this item, so that a stretch of pieces is counted once. */
      let turned: string | undefined;
      const close = () => {
        current = undefined;
      };
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (!item || item.type !== 'text') continue;
        if (item.link !== undefined && !item.deleted && item.rotation !== undefined) {
          // A link in text turned in a cell: its box would be in the wrong place, so it gets none - counted.
          if (turned !== item.link) out.rotated++;
          turned = item.link;
          close();
          continue;
        }
        turned = undefined;
        if (item.link === undefined || item.deleted || !(item.width > 0)) {
          if (item.link === undefined) close();
          continue;
        }
        const target = item.link;
        const box = metrics(item.font);
        const top = item.y - box.ascent;
        const height = box.ascent + box.descent;
        if (current && current.target === target) {
          const last = current.last;
          if (Math.abs(last.y - item.y) <= 0.5 && Math.abs(current.right - item.x) < 1) {
            // The next piece on the same line: one rectangle for both.
            const rect = current.rects[current.rects.length - 1];
            const bottom = Math.max(rect.y + rect.height, (top + height) * scale);
            rect.y = Math.min(rect.y, top * scale);
            rect.height = bottom - rect.y;
            rect.width = (item.x + item.width) * scale - rect.x;
            current.right = item.x + item.width;
            current.last = item;
            continue;
          }
          if (last.source && item.source && last.source.story === item.source.story && last.source.path.join(',') === item.source.path.join(',') && current.rects.length < 64) {
            // The same link on the next line of its paragraph.
            current.rects.push({ x: item.x * scale, y: top * scale, width: item.width * scale, height: height * scale });
            current.right = item.x + item.width;
            current.last = item;
            continue;
          }
        }
        // A new link.
        const rects: PdfLinkRect[] = [{ x: item.x * scale, y: top * scale, width: item.width * scale, height: height * scale }];
        current = { target, link: undefined, rects, last: item, right: item.x + item.width };
        if (list.length >= MAX_PAGE_LINKS) {
          out.dropped++;
          continue;
        }
        if (target.charAt(0) === '#') {
          const spot = bookmarkSpot(target.slice(1));
          if (!spot) {
            out.unresolved++;
            continue;
          }
          const link: PdfLink = { rects, dest: { page: -1 } };
          internal.push({ link, spot });
          current.link = link;
          list.push(link);
        } else {
          const link: PdfLink = { rects, uri: target };
          current.link = link;
          list.push(link);
        }
      }
    }
  }

  // The headings.
  if (wantOutline) for (let i = 0; i < read.headings.length; i++) want(read.headings[i].key, read.headings[i].heading.spot);

  // One pass over the pages: each place is the first line of its paragraph that holds its offset, or its last.
  if (wanted.size > 0) {
    for (let p = 0; p < pages.length; p++) {
      const items = pages[p] && Array.isArray(pages[p].items) ? pages[p].items : [];
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (!item || (item.type !== 'text' && item.type !== 'mark') || item.repeat) continue;
        if (item.type === 'text' && item.deleted) continue;
        const key = bodyKey(item.source);
        if (key === undefined) continue;
        const spots = wanted.get(key);
        if (!spots) continue;
        const source = item.source as { offset: number };
        let top: number;
        let end: number;
        if (item.type === 'text') {
          top = item.y - metrics(item.font).ascent;
          end = source.offset + item.text.length;
        } else {
          const mark = item as LineMark;
          top = mark.y;
          end = source.offset + mark.length + (mark.last ? 1 : 0);
        }
        for (let s = 0; s < spots.length; s++) {
          const spot = spots[s];
          if (spot.found) continue;
          // Taken for now when it does not hold the offset: a later line of the paragraph may.
          spot.page = p;
          spot.x = item.x;
          spot.y = top;
          if (source.offset <= spot.offset && spot.offset < end) spot.found = true;
        }
      }
    }
  }

  // The links into the document, where their bookmarks are.
  const nowhere = new Set<PdfLink>();
  for (let i = 0; i < internal.length; i++) {
    const { link, spot } = internal[i];
    if (spot.page === undefined) {
      // The bookmark is in the document but on no page (a paragraph that is hidden): the link is left out.
      nowhere.add(link);
      out.unresolved++;
      continue;
    }
    link.dest = { page: spot.page, x: (spot.x as number) * scale, y: (spot.y as number) * scale };
  }
  // Taken out of their pages in one pass.
  if (nowhere.size > 0) for (let p = 0; p < out.pages.length; p++) out.pages[p] = out.pages[p].filter((link) => !nowhere.has(link));

  for (let i = 0; i < read.headings.length; i++) {
    const heading = read.headings[i].heading;
    const spot = heading.spot;
    if (spot.page === undefined) {
      out.outlineSkipped++;
      continue;
    }
    out.outline.push({ title: heading.title, level: heading.level, dest: { page: spot.page, x: (spot.x as number) * scale, y: (spot.y as number) * scale } });
  }
  return out;
}
