// The page of a document, section by section: the paper, its orientation and the margins of the
// sections a selection touches (or of every section), and section breaks. A section is the run of
// blocks of the body up to a paragraph whose properties end it (ParaProps.section), or up to the
// end of the body (Body.section). A change puts new properties in the place of the old ones: a new
// paragraph object for a section a paragraph ends, "setBodySection" for the last one. What else
// the section says (headers and footers, columns, page numbers, what the model keeps as XML) stays.
//
// Word records a change of section properties as a tracked change (w:sectPrChange). That is not
// made here: where changes are tracked these commands are refused (TRACK_UNSUPPORTED), so that
// nothing is changed without a trace.

import type { DocModel, Paragraph, SectionProps } from '../../model/doc/types';
import { buildEdit, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { caretIndex, flatten, paragraphWith, rebuild } from './pieces';
import { asRange, inTextBox, orderRange, resolvePosition, type DocPosition, type DocRange } from './position';
import { newMarkProps, withoutSection } from './text';

/** What a page of a section has when its properties do not say (Word's defaults: US Letter, one-inch margins), twips. */
export const DEFAULT_SECTION_PAGE = { width: 12240, height: 15840 };
export const DEFAULT_SECTION_MARGINS = { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720, gutter: 0 };
/** The sides of a page Word takes: from a tenth of an inch to 22 inches, twips. */
export const MIN_PAGE_TWIPS = 144;
export const MAX_PAGE_TWIPS = 31680;
/** What the margins have to leave of the page for the text, twips (a quarter of an inch). */
const MIN_TEXT_TWIPS = 360;

const mm = (value: number) => Math.round((value / 25.4) * 1440);

/** Papers by name, upright, twips. */
export const SECTION_PAPERS: { readonly [name: string]: { readonly width: number; readonly height: number } } = {
  letter: { width: 12240, height: 15840 },
  legal: { width: 12240, height: 20160 },
  executive: { width: 10440, height: 15120 },
  a3: { width: mm(297), height: mm(420) },
  a4: { width: mm(210), height: mm(297) },
  a5: { width: mm(148), height: mm(210) },
  b5: { width: mm(182), height: mm(257) },
};

export type SectionMarginName = 'top' | 'bottom' | 'left' | 'right' | 'header' | 'footer' | 'gutter';

/** The page of a section as it counts (stated or by default), twips. */
export interface SectionSetup {
  width: number;
  height: number;
  orientation: 'portrait' | 'landscape';
  /** The name of the paper (see SECTION_PAPERS) when the page is one of them, either way up. */
  paper?: string;
  /** Margins as lengths (a margin of the file that is fixed, written negative, is given as its length). */
  margins: { [name in SectionMarginName]: number };
}

/** A change of the page of sections: what is left out stays. */
export interface SectionPatch {
  /** A paper of SECTION_PAPERS, or a size in twips (upright or not: the orientation decides). */
  paper?: string | { width: number; height: number };
  orientation?: 'portrait' | 'landscape';
  /** Twips, 0..31680 each. */
  margins?: Partial<{ [name in SectionMarginName]: number }>;
}

/** Which sections a change is for: those a range touches (the section of a caret), or every section of the document. */
export type SectionScope = 'section' | 'document';

/** How the section after a break starts. */
export type SectionBreakKind = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';

const bad = (message: string): never => {
  throw new RangeError(message);
};

const isTwips = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && value >= min && value <= max && Math.floor(value) === value;

/** One section of the body: the block whose paragraph ends it, or -1 for the last one (Body.section). */
export interface SectionRef {
  /** Index of the paragraph of the body that ends the section; -1 for the section at the end of the body. */
  end: number;
  /** The first block of the body in the section. */
  start: number;
  props: SectionProps;
}

/** The sections of the body in order; there is always one, the last of which is the body's own. */
export function documentSections(doc: DocModel): SectionRef[] {
  const out: SectionRef[] = [];
  const blocks = doc.body.blocks;
  let start = 0;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type !== 'paragraph' || !block.props.section) continue;
    out.push({ end: i, start, props: block.props.section });
    start = i + 1;
  }
  out.push({ end: -1, start, props: doc.body.section || {} });
  return out;
}

/** Index (into documentSections) of the section a block of the body is in. */
function sectionOfBlock(sections: SectionRef[], block: number): number {
  for (let i = 0; i < sections.length - 1; i++) if (block <= sections[i].end) return i;
  return sections.length - 1;
}

/**
 * Indexes (into documentSections) of the sections a position or a range touches. A position
 * outside of the body (a header, a note) is in no section of its own: every section counts then.
 */
export function sectionsAt(doc: DocModel, at: DocPosition | DocRange): number[] {
  const sections = documentSections(doc);
  const range = orderRange(asRange(at));
  if (range.start.story !== 'body' || range.end.story !== 'body' || range.start.path.length === 0 || range.end.path.length === 0) return sections.map((_, i) => i);
  const first = sectionOfBlock(sections, range.start.path[0]);
  const last = sectionOfBlock(sections, range.end.path[0]);
  const out: number[] = [];
  for (let i = first; i <= last; i++) out.push(i);
  return out;
}

/** The page of a section as it counts. */
export function sectionSetup(props: SectionProps): SectionSetup {
  const page = props.page || {};
  const width = isTwips(page.width, 1, 10000000) ? page.width : DEFAULT_SECTION_PAGE.width;
  const height = isTwips(page.height, 1, 10000000) ? page.height : DEFAULT_SECTION_PAGE.height;
  const orientation = page.orientation === 'landscape' || (page.orientation === undefined && width > height) ? 'landscape' : 'portrait';
  const margins = {} as SectionSetup['margins'];
  const stated = (props.margins || {}) as { [name: string]: number | undefined };
  for (const name of Object.keys(DEFAULT_SECTION_MARGINS) as SectionMarginName[]) {
    const value = stated[name];
    margins[name] = typeof value === 'number' && value === value ? Math.abs(value) : DEFAULT_SECTION_MARGINS[name];
  }
  const setup: SectionSetup = { width, height, orientation, margins };
  const short = Math.min(width, height);
  const long = Math.max(width, height);
  for (const name of Object.keys(SECTION_PAPERS)) {
    const paper = SECTION_PAPERS[name];
    // A twip or two of rounding either way is the same paper.
    if (Math.abs(paper.width - short) <= 2 && Math.abs(paper.height - long) <= 2) {
      setup.paper = name;
      break;
    }
  }
  return setup;
}

/** The properties of a section with the patch applied; the very object when nothing changes. */
export function patchedSection(props: SectionProps, patch: SectionPatch): SectionProps {
  const now = sectionSetup(props);
  let width = now.width;
  let height = now.height;
  if (patch.paper !== undefined) {
    let size: { width: number; height: number } | undefined;
    if (typeof patch.paper === 'string') size = Object.prototype.hasOwnProperty.call(SECTION_PAPERS, patch.paper) ? SECTION_PAPERS[patch.paper] : undefined;
    else if (patch.paper && typeof patch.paper === 'object') size = patch.paper;
    if (!size || !isTwips(size.width, MIN_PAGE_TWIPS, MAX_PAGE_TWIPS) || !isTwips(size.height, MIN_PAGE_TWIPS, MAX_PAGE_TWIPS)) return bad('The paper is not one that is known');
    width = size.width;
    height = size.height;
  }
  let orientation = now.orientation;
  if (patch.orientation !== undefined) {
    if (patch.orientation !== 'portrait' && patch.orientation !== 'landscape') return bad('The orientation is portrait or landscape');
    orientation = patch.orientation;
  }
  // The long side goes across for landscape, down for portrait (Word turns the paper, it does not stretch it).
  if ((orientation === 'landscape') !== width > height && width !== height) {
    const swap = width;
    width = height;
    height = swap;
  }
  const margins = { ...now.margins };
  if (patch.margins !== undefined) {
    if (!patch.margins || typeof patch.margins !== 'object') return bad('The margins are an object');
    for (const name of Object.keys(DEFAULT_SECTION_MARGINS) as SectionMarginName[]) {
      const value = patch.margins[name];
      if (value === undefined) continue;
      if (!isTwips(value, 0, MAX_PAGE_TWIPS)) return bad('A margin is a whole number of twips from 0 to 31680');
      margins[name] = value;
    }
  }
  if (margins.left + margins.right + margins.gutter > width - MIN_TEXT_TWIPS || margins.top + margins.bottom > height - MIN_TEXT_TWIPS) return bad('The margins leave no room on the page');

  const page = props.page || {};
  const sizeChanged = width !== now.width || height !== now.height;
  const turned = orientation !== now.orientation;
  let out: SectionProps = props;
  if (sizeChanged || turned || (patch.paper !== undefined && (page.width === undefined || page.height === undefined))) {
    const next: NonNullable<SectionProps['page']> = { ...page, width, height };
    if (orientation === 'landscape') next.orientation = 'landscape';
    else if (next.orientation !== undefined) next.orientation = 'portrait';
    // UNSURE: the code of the printer's paper names the old paper; it goes with a paper of another size.
    if (sizeChanged && (Math.min(width, height) !== Math.min(now.width, now.height) || Math.max(width, height) !== Math.max(now.width, now.height))) delete next.code;
    out = { ...out, page: next };
  }
  if (patch.margins !== undefined) {
    const stated = props.margins || {};
    const next: NonNullable<SectionProps['margins']> = { ...stated };
    let changed = false;
    for (const name of Object.keys(DEFAULT_SECTION_MARGINS) as SectionMarginName[]) {
      if (patch.margins[name] === undefined || margins[name] === now.margins[name]) continue;
      // A margin the file fixed (negative: the text does not move down for a tall header) stays fixed.
      const fixed = (name === 'top' || name === 'bottom') && typeof stated[name] === 'number' && (stated[name] as number) < 0;
      next[name] = fixed && margins[name] > 0 ? -margins[name] : margins[name];
      changed = true;
    }
    if (changed) {
      // <w:pgMar> has to state all seven (CT_PageMar): what the section does not state is what counts for it now.
      for (const name of Object.keys(DEFAULT_SECTION_MARGINS) as SectionMarginName[]) if (typeof next[name] !== 'number') next[name] = now.margins[name];
      out = { ...out, margins: next };
    }
  }
  return out;
}

/** Puts other properties in the place of those of a section (see SectionRef). */
function putSection(doc: DocModel, draft: Draft, section: SectionRef, props: SectionProps): void {
  if (section.end < 0) {
    draft.apply({ type: 'setBodySection', section: props });
    return;
  }
  const paragraph = doc.body.blocks[section.end] as Paragraph;
  draft.setParagraph('body', [section.end], paragraphWith(paragraph, paragraph.children, { ...paragraph.props, section: props }));
}

/**
 * Changes the page of sections: of those `at` touches (a caret: its section), or with `scope`
 * "document" of every section. The selection stays. Refused where changes are tracked
 * (TRACK_UNSUPPORTED). No operation when nothing changes.
 */
export function setSectionSetup(doc: DocModel, at: DocPosition | DocRange, patch: SectionPatch, scope: SectionScope = 'section', options?: DocEditOptions): DocEdit {
  if (!patch || typeof patch !== 'object') bad('A change of the page is an object');
  if (scope !== 'section' && scope !== 'document') bad('A change of the page is for a section or for the document');
  const range = asRange(at);
  return buildEdit(
    doc,
    (draft) => {
      const sections = documentSections(doc);
      const chosen = scope === 'document' ? sections.map((_, i) => i) : sectionsAt(doc, range);
      for (const index of chosen) {
        const section = sections[index];
        const next = patchedSection(section.props, patch);
        if (next !== section.props) putSection(doc, draft, section, next);
      }
      return range;
    },
    options,
  );
}

/**
 * A section break at a position of the body (not in a table or a text box): the paragraph is cut
 * there, and the first half ends a new section with the properties of the section it was in (its
 * paper, margins, headers and footers). The section after the break starts as `kind` says. The
 * caret goes to the start of the second half. Refused where changes are tracked (TRACK_UNSUPPORTED).
 */
export function insertSectionBreak(doc: DocModel, position: DocPosition, kind: SectionBreakKind = 'nextPage', options?: DocEditOptions): DocEdit {
  if (kind !== 'nextPage' && kind !== 'continuous' && kind !== 'evenPage' && kind !== 'oddPage') bad('A section break starts the next section on a new page, an even or odd page, or continues');
  return buildEdit(
    doc,
    (draft) => {
      if (position.story !== 'body') bad('A section break can be put into the body only');
      if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
      if (position.path.length !== 1) bad('A section break cannot be put into a table cell');
      const at = resolvePosition(doc, position);
      const paragraph = at.paragraph;
      const blockIndex = position.path[0];
      const sections = documentSections(doc);
      const section = sections[sectionOfBlock(sections, blockIndex)];
      const pieces = flatten(paragraph);
      const index = caretIndex(pieces, position);
      // The new section is the one the text before the break was in, and starts the way that one did.
      const before: SectionProps = { ...section.props };
      delete before.change;
      delete before.attrs;
      // The section after the break starts as asked; "next page" is what a section without a type does.
      let after: SectionProps = section.props;
      const wanted = kind === 'nextPage' ? undefined : kind;
      if ((after.type || undefined) !== wanted && !(wanted === undefined && after.type === 'nextPage')) {
        after = { ...after };
        if (wanted === undefined) delete after.type;
        else after.type = wanted;
      }
      if (before.type === undefined) delete before.type;
      const firstProps = { ...newMarkProps(withoutSection(paragraph.props)), section: before };
      let secondProps = paragraph.props;
      if (section.end === blockIndex && after !== section.props) secondProps = { ...paragraph.props, section: after };
      draft.replace('body', [], blockIndex, 1, [paragraphWith(paragraph, rebuild(pieces.slice(0, index)), firstProps), paragraphWith(paragraph, rebuild(pieces.slice(index)), secondProps, false)]);
      // The section after the break ends elsewhere: its properties get the type there.
      if (section.end !== blockIndex && after !== section.props) {
        if (section.end < 0) draft.apply({ type: 'setBodySection', section: after });
        else {
          const end = section.end + 1;
          const last = doc.body.blocks[end] as Paragraph;
          draft.setParagraph('body', [end], paragraphWith(last, last.children, { ...last.props, section: after }));
        }
      }
      return { story: 'body', path: [blockIndex + 1], offset: 0 };
    },
    options,
  );
}
