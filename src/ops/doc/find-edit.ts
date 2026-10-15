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
  const pieces = flatten(paragraph);
  const out: Piece[] = [];
  let first: Piece | undefined;
  let at0 = -1;
  let offset = 0;
  for (const piece of pieces) {
    const from = offset;
    const to = offset + piece.len;
    offset = to;
    if (piece.len === 0 || to <= start.offset || from >= end.offset) {
      out.push(piece);
      continue;
    }
    if (!plainText(piece)) return fallback();
    const own = (piece.item as { text: string }).text;
    const head = from < start.offset ? own.slice(0, start.offset - from) : '';
    const tail = to > end.offset ? own.slice(end.offset - from) : '';
    if (!first) {
      // Inside another author's tracked change the text would become part of it (rejecting that would take the new
      // text away): insertText() decides where it goes, as for typing.
      if (inRevision(piece.wrap)) return fallback();
      first = piece;
      if (head) out.push(textPart(piece, head));
      at0 = out.length;
      // As typed text: the formatting without the tracked change of formatting someone made to the old text.
      for (const made of textPieces(clean, withoutChange(piece.props), piece.wrap, piece.group)) out.push(made);
    }
    if (tail) out.push(textPart(piece, tail));
  }
  if (!first) return fallback();
  // Text that was all a content control held, replaced by nothing: an empty text keeps the control (with its properties). Any
  // other wrapper (a link, a smart tag) that holds nothing any more goes, as it does when its text is deleted.
  const control = controlOf(first.wrap);
  if (control && !out.some((piece) => piece.item !== undefined && carries(piece, control))) out.splice(at0, 0, { ...textPart(first, ''), props: withoutChange(first.props), wrap: control });
  const next = paragraphWith(paragraph, rebuild(out));
  const caret: DocPosition = { story: start.story, path: start.path.slice(), offset: start.offset + clean.length };
  return buildEdit(doc, (draft) => {
    draft.setParagraph(start.story, start.path, next);
    return { anchor: caret, focus: { ...caret, path: caret.path.slice() } };
  });
}
