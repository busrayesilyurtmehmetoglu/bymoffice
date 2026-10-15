// What Find and Replace may change, and how it changes a match.
//
// replaceText() puts the replacement in the place of the matched characters themselves: in the
// run (and the wrappers) of the first of them, so that the new text stays where the old one was
// - inside a bookmark or a comment range that begins or ends at the match, inside a content
// control whose whole text it was - and takes the formatting of the first matched character, as
// Word does. Where that cannot be done (changes are tracked, the replacement has a paragraph end,
// the match is not plain text) it is insertText().
//
// The rules of what cannot be replaced, shared by the search (which marks such matches) and the
// commands (which refuse them): content controls whose content is locked, alternate content
// around blocks or runs, and text boxes the writer could not write back (textBoxesWritable()).

import { createBodyContext } from '../../formats/docx/write-package';
import { textBoxesWritable, type BodyContext } from '../../formats/docx/write-body';
import type { DocModel, Paragraph, Wrapper } from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { buildEdit, type DocEdit, type DocEditOptions } from './draft';
import { cleanText, flatten, paragraphWith, rebuild, textPieces, withoutChange, type Piece } from './pieces';
import { asRange, orderRange, resolvePosition, storyOf, type DocPosition, type DocRange, type GraphicPlace, type ParagraphGraphic } from './position';
import { insertText } from './text';
import { trackingFor } from './track';

const localName = (name: string): string => name.slice(name.indexOf(':') + 1);

/** Whether a wrapper or one around it is a content control whose content may not be edited (<w:lock w:val="contentLocked"/> or "sdtContentLocked"). */
export function lockedControl(wrap: Wrapper | undefined): boolean {
  for (let w = wrap, depth = 0; w && depth < 256; w = w.parent, depth++) {
    if (w.kind !== 'sdt' || !w.pr) continue;
    for (const node of w.pr.children) {
      if (node.type !== 'element' || localName(node.name) !== 'lock') continue;
      let value: string | undefined;
      node.attrs.forEach((v, key) => {
        if (localName(key) === 'val') value = v;
      });
      if (value === 'contentLocked' || value === 'sdtContentLocked') return true;
    }
  }
  return false;
}

/** Whether a wrapper or one around it is alternate content: written as it was read while its branch is unchanged, so what is in it is not edited. */
export function inAlternate(wrap: Wrapper | undefined): boolean {
  for (let w = wrap, depth = 0; w && depth < 4096; w = w.parent, depth++) if (w.kind === 'alternate') return true;
  return false;
}

/** Whether what a wrapper holds may not be replaced: a locked content control or alternate content. */
export const fixedWrap = (wrap: Wrapper | undefined): boolean => wrap !== undefined && (lockedControl(wrap) || inAlternate(wrap));

/** The context the writer writes a story with (its part, its namespaces): what textBoxesWritable() is asked with. */
export function storyContext(doc: DocModel, story: string): BodyContext {
  const host = storyOf(doc, story).host;
  return createBodyContext(host.part, host.rootAttrs, host.relationships);
}

/**
 * Whether the text boxes of a graphic of a paragraph may not be replaced in: the run that holds
 * the graphic is in a locked content control or in alternate content, or the writer could not
 * write the change back in every branch (textBoxesWritable()).
 */
export function graphicFixed(ctx: BodyContext, graphic: ParagraphGraphic, place: GraphicPlace): boolean {
  if (fixedWrap(place.run.wrap)) return true;
  if (textLocked(graphic)) return true;
  return !textBoxesWritable(ctx, graphic, place.holder, place.nested);
}

const isTrue = (value: string | undefined): boolean => value === '1' || value === 't' || value === 'true' || value === 'on';

/**
 * Whether the shape says its text may not be edited: <a:spLocks noTextEdit="1"/> (DrawingML) or
 * <o:lock text="t"/> (VML). Only the shape itself is looked at, not the shapes in its text boxes.
 */
export function textLocked(graphic: ParagraphGraphic): boolean {
  const xml = graphic.xml as XmlElement | undefined;
  if (!xml) return false;
  let budget = 20000;
  const walk = (el: XmlElement, depth: number): boolean => {
    if (depth > 64 || --budget < 0) return false;
    const local = localName(el.name);
    if (local === 'txbxContent') return false;
    if (local === 'spLocks' && isTrue(attrOf(el, 'noTextEdit'))) return true;
    if (local === 'lock' && isTrue(attrOf(el, 'text'))) return true;
    for (const node of el.children) if (node.type === 'element' && walk(node, depth + 1)) return true;
    return false;
  };
  return walk(xml, 0);
}

function attrOf(el: XmlElement, local: string): string | undefined {
  let value: string | undefined;
  el.attrs.forEach((v, key) => {
    if (localName(key) === local) value = v;
  });
  return value;
}

/** Whether a paragraph as a whole may not be replaced in: it is in a locked content control or in alternate content around blocks. */
export const paragraphFixed = (paragraph: Paragraph): boolean => fixedWrap(paragraph.wrap);

const plainText = (piece: Piece): boolean => piece.item !== undefined && piece.item.type === 'text' && !piece.frozen && !piece.code && piece.lock === undefined;

const textPart = (piece: Piece, text: string): Piece => ({ ...piece, item: { type: 'text', text }, index: -1, len: text.length });

function inRevision(wrap: Wrapper | undefined): boolean {
  for (let w = wrap, depth = 0; w && depth < 4096; w = w.parent, depth++) if (w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo') return true;
  return false;
}

/** The innermost content control a wrapper is or is in. */
function controlOf(wrap: Wrapper | undefined): Wrapper | undefined {
  for (let w = wrap, depth = 0; w && depth < 4096; w = w.parent, depth++) if (w.kind === 'sdt') return w;
  return undefined;
}

function carries(piece: Piece, wrapper: Wrapper): boolean {
  for (let w = piece.wrap, depth = 0; w && depth < 4096; w = w.parent, depth++) if (w === wrapper) return true;
  return false;
}

/**
 * Replaces a range of one paragraph with text in the place of the replaced characters (see the
 * head of this file); insertText() where that cannot be done. One edit.
 */
export function replaceText(doc: DocModel, at: DocPosition | DocRange, text: string, options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  const clean = cleanText(text);
  const fallback = (): DocEdit => insertText(doc, range, text, undefined, options);
  if (clean.indexOf('\n') >= 0 || trackingFor(doc, options)) return fallback();
  const { start, end } = orderRange(range);
  if (start.story !== end.story || start.path.length !== end.path.length || start.path.some((value, i) => end.path[i] !== value) || start.offset >= end.offset) return fallback();
  const paragraph = resolvePosition(doc, start).paragraph;
  resolvePosition(doc, end);
  const next = replacedParagraph(paragraph, start.offset, end.offset, clean);
  if (!next) return fallback();
  const caret: DocPosition = { story: start.story, path: start.path.slice(), offset: start.offset + clean.length };
  return buildEdit(doc, (draft) => {
    draft.setParagraph(start.story, start.path, next);
    return { anchor: caret, focus: { ...caret, path: caret.path.slice() } };
  });
}

/**
 * A paragraph with the characters [start, end) of its text replaced by `clean` (text without a
 * paragraph end, cleaned by cleanText()) in the place of the replaced characters, as replaceText()
 * does; undefined where that cannot be done: the range holds something other than plain text of
 * runs, begins inside a tracked change, or is not in the text. Nothing is tracked.
 */
export function replacedParagraph(paragraph: Paragraph, start: number, end: number, clean: string): Paragraph | undefined {
  if (!(start >= 0 && start < end) || clean.indexOf('\n') >= 0) return undefined;
  const pieces = flatten(paragraph);
  const out: Piece[] = [];
  let first: Piece | undefined;
  let at0 = -1;
  let offset = 0;
  for (const piece of pieces) {
    const from = offset;
    const to = offset + piece.len;
    offset = to;
    if (piece.len === 0 || to <= start || from >= end) {
      out.push(piece);
      continue;
    }
    if (!plainText(piece)) return undefined;
    const own = (piece.item as { text: string }).text;
    const head = from < start ? own.slice(0, start - from) : '';
    const tail = to > end ? own.slice(end - from) : '';
    if (!first) {
      // Inside another author's tracked change the text would become part of it (rejecting that would take the new
      // text away): insertText() decides where it goes, as for typing.
      if (inRevision(piece.wrap)) return undefined;
      first = piece;
      if (head) out.push(textPart(piece, head));
      at0 = out.length;
      // As typed text: the formatting without the tracked change of formatting someone made to the old text.
      for (const made of textPieces(clean, withoutChange(piece.props), piece.wrap, piece.group)) out.push(made);
    }
    if (tail) out.push(textPart(piece, tail));
  }
  if (!first || offset < end) return undefined;
  // Text that was all a content control held, replaced by nothing: an empty text keeps the control (with its properties). Any
  // other wrapper (a link, a smart tag) that holds nothing any more goes, as it does when its text is deleted.
  const control = controlOf(first.wrap);
  if (control && !out.some((piece) => piece.item !== undefined && carries(piece, control))) out.splice(at0, 0, { ...textPart(first, ''), props: withoutChange(first.props), wrap: control });
  return paragraphWith(paragraph, rebuild(out));
}

/**
 * replacedParagraph() for many ranges of one paragraph at once, in one pass over its pieces, and
 * strictly, for text whose replacement is never tracked (a comment): besides what replacedParagraph()
 * refuses, a range is refused when any of its text is in a tracked change (insertion, deletion,
 * move) or a tracked deletion lies inside it - replacing it would take that change away without a
 * trace. `ranges`: [start, end) in the order of the text, none overlapping another. `done` says
 * which were replaced; `paragraph` is undefined when none was.
 */
export function replacedParagraphRanges(paragraph: Paragraph, ranges: readonly { start: number; end: number }[], clean: string): { paragraph?: Paragraph; done: boolean[] } {
  const done = ranges.map(() => false);
  if (clean.indexOf('\n') >= 0 || ranges.length === 0) return { done };
  const pieces = flatten(paragraph);
  const from: number[] = [];
  let total = 0;
  for (const piece of pieces) {
    from.push(total);
    total += piece.len;
  }
  // Which ranges can be replaced: each looks at the pieces it covers, and the ranges go on in the order of the text.
  let at = 0;
  let previous = -1;
  for (let i = 0; i < ranges.length; i++) {
    const { start, end } = ranges[i];
    if (!(start >= 0 && start < end && end <= total && start >= previous)) continue;
    previous = end;
    while (at < pieces.length && from[at] + pieces[at].len <= start) at++;
    let ok = false;
    for (let k = at; k < pieces.length && from[k] < end; k++) {
      const piece = pieces[k];
      if (piece.len === 0) {
        // Something without width inside the range: a deletion would be taken along.
        if (from[k] > start && inRevision(piece.wrap)) {
          ok = false;
          break;
        }
        continue;
      }
      if (!plainText(piece) || inRevision(piece.wrap)) {
        ok = false;
        break;
      }
      ok = true;
    }
    done[i] = ok;
  }
  if (!done.some(Boolean)) return { done };
  // The pass that replaces: each piece is cut by the ranges it holds a part of.
  const out: Piece[] = [];
  const firsts: { piece: Piece; at: number }[] = [];
  const used = ranges.filter((_range, i) => done[i]);
  let r = 0;
  for (let k = 0; k < pieces.length; k++) {
    const piece = pieces[k];
    const start = from[k];
    const end = start + piece.len;
    while (r < used.length && used[r].end <= start) r++;
    if (piece.len === 0 || r >= used.length || used[r].start >= end) {
      out.push(piece);
      continue;
    }
    const own = (piece.item as { text: string }).text;
    let cursor = start;
    for (let q = r; q < used.length && used[q].start < end; q++) {
      const range = used[q];
      if (range.start >= cursor && range.start >= start) {
        if (range.start > cursor) out.push(textPart(piece, own.slice(cursor - start, range.start - start)));
        firsts.push({ piece, at: out.length });
        for (const made of textPieces(clean, withoutChange(piece.props), piece.wrap, piece.group)) out.push(made);
      }
      cursor = Math.max(cursor, range.end);
    }
    if (cursor < end) out.push(textPart(piece, own.slice(cursor - start)));
  }
  // A content control whose whole text was replaced by nothing keeps an empty text (once), as in replacedParagraph().
  const carried = new Set<Wrapper>();
  for (const piece of out) {
    if (piece.item === undefined) continue;
    for (let w = piece.wrap, depth = 0; w && depth < 4096 && !carried.has(w); w = w.parent, depth++) carried.add(w);
  }
  const empties: { at: number; piece: Piece }[] = [];
  for (const first of firsts) {
    const control = controlOf(first.piece.wrap);
    if (!control || carried.has(control)) continue;
    carried.add(control);
    empties.push({ at: first.at, piece: { ...textPart(first.piece, ''), props: withoutChange(first.piece.props), wrap: control } });
  }
  for (let i = empties.length - 1; i >= 0; i--) out.splice(empties[i].at, 0, empties[i].piece);
  return { paragraph: paragraphWith(paragraph, rebuild(out)), done };
}
