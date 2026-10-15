// The panel that lists the comments of a document beside the pages: every
// thread in the order of the text it is about, with its author, its date, a
// short quote of that text, what it says and its answers. A thread that is
// picked in the panel is shown in the document; a thread the caret goes into
// is picked in the panel. The selected thread offers what the balloon offers:
// answer, edit, resolve or reopen, delete.
//
// The view does the work on the document; this module only builds elements
// and tells the view what the user asked for. Everything that comes from the
// file (names of authors, dates, the text of comments and of the document)
// reaches the DOM as text, never as markup. Nothing here touches the DOM at
// import time.
//
// Many comments: the list is built a batch at a time (more as it is scrolled
// or stepped through), and a refresh only builds anew the entries whose
// content changed.

import { formatValue } from '../format/number-format';
import type { FormatLocale } from '../format/locale';
import { paragraphText, type DocModel, type Paragraph } from '../model/doc/types';
import { resolvePosition, samePosition, type DocComment } from '../ops/doc/index';
import { fillMarked, type FoundMark } from './doc-comments';
import type { UiStrings } from './strings';

/** A comment with its answers, as the panel lists it. */
export interface CommentThread {
  root: DocComment;
  replies: DocComment[];
}

/** Threads the list is built with at first, and how many more each step builds. */
export const PANEL_BATCH = 100;
/** Most of the text of a comment an entry that is not selected shows, UTF-16 code units; the selected one shows all of it. */
export const PANEL_SHORT_TEXT = 400;
/** Most of the text a comment is about the panel quotes. */
export const QUOTE_LENGTH = 120;
/** Narrower than this the panel lies over the bottom of the pages instead of beside them. */
export const PANEL_NARROW = 600;
/** Width of the panel beside the pages, px. */
export const PANEL_WIDTH = 320;

/** Cosmetic rules of the panel; what its layout depends on is set inline. */
export const COMMENTS_PANEL_STYLE_TEXT =
  '.bym-doc-cpanel{background:#fff;color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
  '-webkit-user-select:text;user-select:text;border-left:1px solid #dadce0}' +
  '.bym-doc-cpanel[data-narrow="true"]{border-left:0;border-top:1px solid #dadce0;box-shadow:0 -4px 16px rgba(0,0,0,.18);border-radius:12px 12px 0 0}' +
  '.bym-doc-cpanel-head{border-bottom:1px solid #eceff1}' +
  '.bym-doc-cpanel-title{font-weight:600;font-size:14px}' +
  '.bym-doc-cpanel-count{color:#5f6368;font-size:12px}' +
  '.bym-doc-cpanel-filter{color:#3c4043;font-size:12px;cursor:pointer}' +
  '.bym-doc-cpanel-item{border:1px solid #e0e3e7;border-radius:8px;padding:8px;margin:0 0 8px;cursor:pointer;outline:none;background:#fff}' +
  '.bym-doc-cpanel-item[data-done="true"]{background:#f8f9fa;color:#5f6368}' +
  '.bym-doc-cpanel-item[aria-current="true"]{border-color:#1a7f45;box-shadow:0 0 0 1px #1a7f45;cursor:default}' +
  '.bym-doc-cpanel-item:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-doc-cpanel-item:focus:not(:focus-visible){outline:none}' +
  '.bym-doc-cpanel-author{font-weight:600}' +
  '.bym-doc-cpanel-date,.bym-doc-cpanel-replies{color:#5f6368;font-size:12px}' +
  '.bym-doc-cpanel-state{color:#0f5c31;font-size:12px;font-weight:600}' +
  '.bym-doc-cpanel-quote{color:#5f6368;font-style:italic;border-left:3px solid #dadce0;padding-left:6px;margin:4px 0;overflow-wrap:anywhere;word-break:break-word}' +
  '.bym-doc-cpanel-text{white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;margin:2px 0}' +
  '.bym-doc-cpanel-item:not([aria-current="true"]) .bym-doc-cpanel-text{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:3;overflow:hidden}' +
  '.bym-doc-cpanel-reply{margin:6px 0 0 12px;padding-top:6px;border-top:1px solid #eceff1}' +
  '.bym-doc-cpanel-empty{color:#5f6368;padding:16px 4px;text-align:center}';

/** A date of the file as the panel shows it: the short date and time of the locale for ISO text, anything else as it is (cut short). */
export function commentDateText(date: string | undefined, locale: FormatLocale): string {
  if (typeof date !== 'string' || date === '') return '';
  const match = /^(\d{4})-(\d\d)-(\d\d)(?:[T ](\d\d):(\d\d))?/.exec(date);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = match[4] !== undefined ? Number(match[4]) : undefined;
    const minute = match[5] !== undefined ? Number(match[5]) : undefined;
    const time = Date.UTC(year, month - 1, day);
    const check = new Date(time);
    // The date as the file writes it: its clock is the one of whoever wrote it, not turned to another zone.
    if (
      year >= 1900 &&
      check.getUTCFullYear() === year &&
      check.getUTCMonth() === month - 1 &&
      check.getUTCDate() === day &&
      (hour === undefined || (hour < 24 && (minute as number) < 60))
    ) {
      let days = (time - Date.UTC(1899, 11, 30)) / 86400000;
      // The number formats count a 29 February 1900 that never was (as Excel does): before March the serial is a day less.
      if (days < 61) days--;
      const serial = days + (hour === undefined ? 0 : (hour * 60 + (minute as number)) / 1440);
      const code = (locale.shortDate || 'm/d/yyyy') + (hour === undefined ? '' : ' ' + (locale.shortTime || 'h:mm'));
      return formatValue(serial, code, { locale }).text;
    }
  }
  return quoteText(date, 64);
}

/** Text on one line, at most `max` code units (an ellipsis where it was cut, never inside a pair of surrogates). */
export function quoteText(text: string, max = QUOTE_LENGTH): string {
  // Only what the length allows is looked at: a quote is never the whole of a long text. A cut there is a cut.
  let end = text.length;
  let cut = false;
  if (end > max * 4) {
    end = max * 4;
    if (highSurrogate(text.charCodeAt(end - 1))) end--;
    cut = true;
  }
  let line = text.slice(0, end).replace(/[\s\u0000-\u001f\u007f­￼]+/g, ' ').trim();
  if (!cut && line.length <= max) return line;
  if (line.length > max - 1) {
    let at = max - 1;
    if (highSurrogate(line.charCodeAt(at - 1))) at--;
    line = line.slice(0, at);
  }
  return line.replace(/\s+$/, '') + '…';
}

const highSurrogate = (code: number): boolean => code >= 0xd800 && code <= 0xdbff;

/**
 * A short quote of the text a comment is about: of its first paragraph and,
 * where it goes on, of its last ("…" between them when paragraphs are left
 * out). Never more than two paragraphs are read. "" for a comment about a
 * point, about nothing, or about a place the document does not have.
 */
export function commentQuote(doc: DocModel, comment: DocComment, textOf: (paragraph: Paragraph) => string = paragraphText): string {
  const start = comment.start;
  const end = comment.end;
  if (!start || !end || samePosition(start, end)) return '';
  try {
    const first = textOf(resolvePosition(doc, { story: start.story, path: start.path, offset: 0 }).paragraph);
    const samePath = start.path.length === end.path.length && start.path.every((step, i) => step === end.path[i]);
    if (start.story === end.story && samePath) return quoteText(first.slice(start.offset, end.offset));
    const head = first.slice(start.offset);
    if (head.length >= QUOTE_LENGTH) return quoteText(head);
    const last = textOf(resolvePosition(doc, { story: end.story, path: end.path, offset: 0 }).paragraph);
    const next =
      start.story === end.story &&
      start.path.length === end.path.length &&
      start.path.every((step, i) => (i < start.path.length - 1 ? step === end.path[i] : step + 1 === end.path[i]));
    return quoteText(head + (next ? ' ' : ' \u2026 ') + last.slice(0, end.offset));
  } catch {
    return '';
  }
}

/**
 * The text of paragraphs, each read once: what quotes are taken from while
 * the document stays as it is. Many comments about one long paragraph read it
 * once, not once each. A document that changed needs a new one.
 */
export function paragraphTexts(): (paragraph: Paragraph) => string {
  const read = new WeakMap<Paragraph, string>();
  return (paragraph) => {
    let text = read.get(paragraph);
    if (text === undefined) {
      text = paragraphText(paragraph);
      read.set(paragraph, text);
    }
    return text;
  };
}

/**
 * How much of the pages (`pages`: their box on the screen) a drawer whose top
 * edge is at `top` covers, in the px of the pages (`layoutHeight`: their
 * height as laid out, which a scaled page shows larger or smaller).
 */
export function coveredBelow(pages: { bottom: number; height: number }, top: number, layoutHeight: number): number {
  if (!(pages.height > 0) || !(layoutHeight > 0)) return 0;
  const hidden = pages.bottom - top;
  if (!(hidden > 0)) return 0;
  return Math.min(layoutHeight, hidden * (layoutHeight / pages.height));
}

/** The text of a comment an entry shows: all of it for the selected one, the start of it for the others. */
export function shownText(text: string, whole: boolean): string {
  if (whole || text.length <= PANEL_SHORT_TEXT) return text;
  let end = PANEL_SHORT_TEXT;
  const code = text.charCodeAt(end - 1);
  if (code >= 0xd800 && code <= 0xdbff) end--;
  return text.slice(0, end) + '…';
}

/** Where a key moves in a list of `count` entries from `index`; undefined for a key that does not move. */
export function stepIndex(index: number, count: number, key: string): number | undefined {
  if (count <= 0) return undefined;
  switch (key) {
    case 'ArrowDown':
      return Math.min(count - 1, index < 0 ? 0 : index + 1);
    case 'ArrowUp':
      return Math.max(0, index < 0 ? 0 : index - 1);
    case 'Home':
      return 0;
    case 'End':
      return count - 1;
    default:
      return undefined;
  }
}

/** The threads the panel lists: every one, or only the open ones. */
export function shownThreads(threads: CommentThread[], showResolved: boolean): CommentThread[] {
  return showResolved ? threads : threads.filter((thread) => !thread.root.done);
}

/** What the panel asks of the view. */
export interface CommentsPanelHost {
  /** The element the panel is put into: beside the pages, positioned. */
  readonly root: HTMLElement;
  strings(): UiStrings;
  /** The threads of the document in the order of their text; the ones nothing refers to at the end. */
  threads(): CommentThread[];
  /** Whether comments can be answered, edited, resolved and deleted now. */
  canChange(): boolean;
  /** A short quote of the text a comment is about; "" for none. */
  quote(comment: DocComment): string;
  /** A date of the file as it is shown. */
  date(date: string | undefined): string;
  /** The user picked a thread: the document shows its text. `toDocument`: the keys go back to the document. */
  reveal(id: number, toDocument: boolean): void;
  /** The thread that is selected changed (also to none): the pages mark it. */
  selectionChanged(id: number | undefined): void;
  /** An answer to the thread `id`, or the new text of the comment `id`. True when it was taken. */
  submit(mode: 'reply' | 'edit', id: number, text: string): boolean;
  resolve(id: number, done: boolean): void;
  remove(id: number): void;
  /** The panel was closed by the user; `refocus`: the focus was in it and has to go somewhere. */
  closed(refocus: boolean): void;
}

/** What is being written in the selected thread. */
interface Compose {
  mode: 'reply' | 'edit';
  thread: number;
  /** The comment that is edited; the first one of the thread for an answer. */
  comment: number;
  draft: string;
}

/** An entry of the list as it was built. */
interface Entry {
  element: HTMLDivElement;
  /** Everything the entry shows: the same text means the entry need not be built again. */
  key: string;
  /** Its buttons, by "action:comment". */
  buttons: Map<string, HTMLButtonElement>;
  field?: HTMLTextAreaElement;
}

function element<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text?: string): HTMLElementTagNameMap[K] {
  const out = document.createElement(tag);
  out.className = className;
  if (text !== undefined) out.textContent = text;
  return out;
}

const KEPT_EVENTS = ['mousedown', 'pointerdown', 'copy', 'cut', 'paste', 'dblclick', 'contextmenu'];

export class CommentsPanel {
  readonly element: HTMLElement;
  private readonly host: CommentsPanelHost;
  private readonly head: HTMLDivElement;
  private readonly title: HTMLDivElement;
  private readonly count: HTMLSpanElement;
  private readonly filter: HTMLInputElement;
  private readonly filterText: HTMLSpanElement;
  private readonly closeButton: HTMLButtonElement;
  private readonly scroll: HTMLDivElement;
  private readonly list: HTMLDivElement;
  private readonly more: HTMLButtonElement;
  private readonly empty: HTMLDivElement;
  private readonly entries = new Map<number, Entry>();
  /** The threads listed, as they were at the last refresh. */
  private visible: CommentThread[] = [];
  private all: CommentThread[] = [];
  /** How many of the visible threads have an entry. */
  private limit = PANEL_BATCH;
  private selected: number | undefined;
  private compose: Compose | undefined;
  private showResolved = true;
  private open = false;
  private narrow = false;
  private destroyed = false;
  /** The strings the entries were built with: entries are built anew when they change. */
  private builtWith: UiStrings | undefined;
  /** A match of Find that is marked in the selected thread, with the text of its comment as it was then: a text that changed is not marked. */
  private found: (FoundMark & { text: string; thread: number }) | undefined;

  constructor(host: CommentsPanelHost) {
    this.host = host;
    const panel = element('aside', 'bym-doc-cpanel');
    panel.setAttribute('role', 'complementary');
    panel.style.cssText = 'display:none;box-sizing:border-box;flex-direction:column;min-width:0;min-height:0;overflow:hidden';
    this.element = panel;

    const head = element('div', 'bym-doc-cpanel-head');
    head.style.cssText = 'display:flex;flex:none;align-items:center;flex-wrap:wrap;gap:4px 8px;box-sizing:border-box;padding:8px 8px 6px 12px';
    const title = element('div', 'bym-doc-cpanel-title');
    title.setAttribute('role', 'heading');
    title.setAttribute('aria-level', '2');
    title.style.cssText = 'flex:1 1 auto;min-width:0';
    const count = element('span', 'bym-doc-cpanel-count');
    count.style.cssText = 'display:block;font-weight:400';
    const label = element('label', 'bym-doc-cpanel-filter');
    label.style.cssText = 'display:inline-flex;align-items:center;gap:4px';
    const filter = element('input', 'bym-doc-cpanel-check');
    filter.type = 'checkbox';
    filter.checked = true;
    filter.style.cssText = 'margin:0';
    const filterText = element('span', '');
    label.appendChild(filter);
    label.appendChild(filterText);
    const close = element('button', 'bym-doc-comment-btn bym-doc-cpanel-close', '×');
    close.type = 'button';
    close.setAttribute('data-action', 'close');
    close.setAttribute('data-small', 'true');
    close.style.marginRight = '0';
    head.appendChild(title);
    head.appendChild(label);
    head.appendChild(close);

    const scroll = element('div', 'bym-doc-cpanel-scroll');
    scroll.style.cssText = 'flex:1 1 0;min-height:0;overflow:auto;overscroll-behavior:contain;-webkit-overflow-scrolling:touch;box-sizing:border-box;padding:8px';
    const list = element('div', 'bym-doc-cpanel-list');
    list.setAttribute('role', 'list');
    const more = element('button', 'bym-doc-comment-btn bym-doc-cpanel-more');
    more.type = 'button';
    more.setAttribute('data-action', 'more');
    more.style.cssText = 'display:none;width:100%';
    const empty = element('div', 'bym-doc-cpanel-empty');
    empty.style.display = 'none';
    scroll.appendChild(list);
    scroll.appendChild(more);
    scroll.appendChild(empty);
    panel.appendChild(head);
    panel.appendChild(scroll);

    this.head = head;
    this.title = title;
    this.count = count;
    this.filter = filter;
    this.filterText = filterText;
    this.closeButton = close;
    this.scroll = scroll;
    this.list = list;
    this.more = more;
    this.empty = empty;

    panel.addEventListener('click', this.onClick);
    panel.addEventListener('keydown', this.onKeyDown);
    panel.addEventListener('input', this.onInput);
    filter.addEventListener('change', this.onFilter);
    scroll.addEventListener('scroll', this.onScroll);
    // What happens in the panel is not the document's: the view listens on the element around both.
    for (const type of KEPT_EVENTS) panel.addEventListener(type, this.stop);
    host.root.appendChild(panel);
  }

  get isOpen(): boolean {
    return this.open;
  }

  /** The first comment of the thread that is selected; undefined when none is. */
  get selectedThread(): number | undefined {
    return this.selected;
  }

  /** Whether an answer or a new text is being written in it. */
  get composing(): boolean {
    return this.open && this.compose !== undefined;
  }

  get hasFocus(): boolean {
    const root = this.element.getRootNode ? (this.element.getRootNode() as Document | ShadowRoot) : document;
    return this.open && this.element.contains(root.activeElement);
  }

  /** How many threads have an entry now (tests, and what was measured). */
  get builtCount(): number {
    return this.entries.size;
  }

  /** Shows the panel. `focus`: the focus goes to the selected thread, else to the first one, else to the panel's first control. */
  show(focus = false): void {
    if (this.destroyed) return;
    this.open = true;
    this.element.style.display = 'flex';
    this.place();
    this.refresh();
    if (focus) this.focusList();
  }

  /** Takes the panel away. Nothing is said to the host: it asked for it. */
  hide(): void {
    if (!this.open) return;
    this.open = false;
    this.compose = undefined;
    this.found = undefined;
    this.element.style.display = 'none';
    // Nothing of a hidden panel is kept up to date: it is built anew when it is shown again.
    this.clearEntries();
  }

  /** Lies beside the pages, or (`narrow`) over the bottom of them. */
  setNarrow(narrow: boolean): void {
    if (this.narrow === narrow) return;
    this.narrow = narrow;
    this.place();
  }

  get isNarrow(): boolean {
    return this.narrow;
  }

  /**
   * Selects a thread (undefined: none), as the document asks when its text
   * or its mark is pressed or the caret goes into it. `scroll`: its entry is
   * brought into sight in the panel; `focus`: the focus goes to it.
   */
  select(id: number | undefined, options: { scroll?: boolean; focus?: boolean } = {}): void {
    if (this.destroyed) return;
    // A thread the document does not have is no thread to select.
    if (id !== undefined && this.open && !this.host.threads().some((thread) => thread.root.id === id)) id = undefined;
    if (id !== this.selected) {
      if (this.compose && this.compose.thread !== id) this.compose = undefined;
      if (this.found && this.found.thread !== id) this.found = undefined;
      this.selected = id;
      if (this.open) this.refresh();
      this.host.selectionChanged(id);
    }
    if (!this.open || id === undefined) return;
    if (this.indexOf(id) < 0) return;
    const entry = this.entries.get(id);
    if (!entry) return;
    if (options.scroll || options.focus) this.bringIntoSight(entry.element);
    if (options.focus) this.focusElement(entry.element);
  }

  /**
   * Marks a match of Find in the text of a comment of the selected thread (select() it first)
   * and brings it into sight in the list; a thread that the filter hides is shown with the
   * resolved ones. Undefined takes the mark away.
   */
  markText(mark: FoundMark | undefined): void {
    if (this.destroyed) return;
    let next: (FoundMark & { text: string; thread: number }) | undefined;
    if (mark && this.open && this.selected !== undefined) {
      for (const thread of this.host.threads()) {
        if (thread.root.id !== this.selected) continue;
        const comment = thread.root.id === mark.comment ? thread.root : thread.replies.filter((reply) => reply.id === mark.comment)[0];
        if (!comment) break;
        next = { comment: mark.comment, start: mark.start, end: mark.end, text: comment.text, thread: thread.root.id };
        if (thread.root.done && !this.showResolved) this.showResolved = true;
        break;
      }
    }
    const had = this.found;
    if (!next && !had) return;
    this.found = next;
    if (!this.open) return;
    this.refresh();
    const entry = next ? this.entries.get(next.thread) : undefined;
    const marked = entry ? (entry.element.querySelector('.bym-doc-find-mark') as HTMLElement | null) : null;
    if (marked) this.bringIntoSight(marked);
    else if (entry) this.bringIntoSight(entry.element);
  }

  /** Brings the list up to date with the threads of the document as they are now, keeping what is being written. */
  refresh(): void {
    if (!this.open || this.destroyed) return;
    this.all = this.host.threads();
    this.visible = shownThreads(this.all, this.showResolved);
    // A thread that is gone is no longer selected; one hidden by the filter stays selected (it is shown again with it).
    if (this.selected !== undefined && !this.all.some((thread) => thread.root.id === this.selected)) {
      this.selected = undefined;
      this.compose = undefined;
      this.host.selectionChanged(undefined);
    }
    if (this.compose && !this.stillComposable(this.compose)) this.compose = undefined;
    this.render();
  }

  /** The words of the panel changed: everything is built anew. */
  setStrings(): void {
    this.builtWith = undefined;
    this.refresh();
  }

  destroy(): void {
    if (this.destroyed) return;
    this.hide();
    this.destroyed = true;
    const panel = this.element;
    panel.removeEventListener('click', this.onClick);
    panel.removeEventListener('keydown', this.onKeyDown);
    panel.removeEventListener('input', this.onInput);
    this.filter.removeEventListener('change', this.onFilter);
    this.scroll.removeEventListener('scroll', this.onScroll);
    for (const type of KEPT_EVENTS) panel.removeEventListener(type, this.stop);
    if (panel.parentNode) panel.parentNode.removeChild(panel);
  }

  // ------------------------------------------------------------- building

  private place(): void {
    const style = this.element.style;
    if (this.narrow) {
      // A sheet over the bottom of the pages: they keep their width on a phone.
      style.position = 'absolute';
      style.left = '0';
      style.right = '0';
      style.bottom = '0';
      style.top = 'auto';
      style.width = 'auto';
      style.height = '60%';
      style.maxHeight = '480px';
      style.zIndex = '7';
      style.flex = 'none';
    } else {
      style.position = 'relative';
      style.left = '';
      style.right = '';
      style.bottom = '';
      style.top = '';
      style.width = `${PANEL_WIDTH}px`;
      style.maxWidth = '45%';
      style.height = 'auto';
      style.maxHeight = 'none';
      style.zIndex = '';
      style.flex = `0 0 ${PANEL_WIDTH}px`;
    }
    this.element.setAttribute('data-narrow', this.narrow ? 'true' : 'false');
  }

  private clearEntries(): void {
    this.entries.forEach((entry) => {
      if (entry.element.parentNode) entry.element.parentNode.removeChild(entry.element);
    });
    this.entries.clear();
    this.limit = PANEL_BATCH;
  }

  private indexOf(id: number): number {
    const visible = this.visible;
    for (let i = 0; i < visible.length; i++) if (visible[i].root.id === id) return i;
    return -1;
  }

  /** Whether what is being written still has its comment, and may still be written. */
  private stillComposable(compose: Compose): boolean {
    if (!this.host.canChange() || compose.thread !== this.selected) return false;
    for (const thread of this.all) {
      if (thread.root.id !== compose.thread) continue;
      if (compose.mode === 'reply') return thread.root.canChange;
      if (thread.root.id === compose.comment) return true;
      return thread.replies.some((reply) => reply.id === compose.comment);
    }
    return false;
  }

  private render(): void {
    const s = this.host.strings();
    // Whatever has the focus in an entry that is built anew gets it back in the new one.
    const restore = this.focusedPlace();
    if (s !== this.builtWith) {
      this.builtWith = s;
      this.element.setAttribute('aria-label', s.commentsPanel);
      this.list.setAttribute('aria-label', s.commentsPanel);
      this.title.textContent = s.commentsPanel;
      this.title.appendChild(this.count);
      this.filterText.textContent = s.commentsShowResolved;
      this.closeButton.setAttribute('aria-label', s.commentClose);
      this.closeButton.title = s.commentClose;
      // Every entry carries words.
      this.clearEntries();
    }
    const hidden = this.all.length - this.visible.length;
    this.count.textContent = s.commentsPanelCount(this.visible.length, hidden);
    if (this.filter.checked !== this.showResolved) this.filter.checked = this.showResolved;

    const can = this.host.canChange();
    const visible = this.visible;
    const count = Math.min(visible.length, this.limit);
    const keep = new Set<number>();
    // An entry is always built for the selected thread, wherever it is in the list.
    if (this.selected !== undefined && this.limit < visible.length) {
      const index = this.indexOf(this.selected);
      if (index >= this.limit) return this.growTo(index);
    }
    // The entries in the order of the threads: one that is in its place stays, the others are put there.
    let cursor: Node | null = this.list.firstChild;
    for (let i = 0; i < count; i++) {
      const thread = visible[i];
      const id = thread.root.id;
      keep.add(id);
      const quote = this.host.quote(thread.root);
      const key = this.keyOf(thread, can, quote);
      let entry = this.entries.get(id);
      if (!entry || entry.key !== key) {
        const made = this.build(thread, can, key, quote);
        if (entry && entry.element.parentNode) {
          if (cursor === entry.element) cursor = cursor.nextSibling;
          entry.element.parentNode.removeChild(entry.element);
        }
        entry = made;
        this.entries.set(id, entry);
      }
      if (cursor === entry.element) cursor = cursor.nextSibling;
      else this.list.insertBefore(entry.element, cursor);
    }
    // Entries of threads that are gone, hidden or beyond the ones built.
    this.entries.forEach((entry, id) => {
      if (keep.has(id)) return;
      if (entry.element.parentNode) entry.element.parentNode.removeChild(entry.element);
      this.entries.delete(id);
    });
    const left = visible.length - count;
    this.more.style.display = left > 0 ? 'block' : 'none';
    this.more.textContent = left > 0 ? s.commentsShowMore(Math.min(left, PANEL_BATCH)) : '';
    this.empty.style.display = visible.length === 0 ? 'block' : 'none';
    this.empty.textContent = visible.length === 0 ? s.commentsPanelEmpty(hidden) : '';
    this.keepTabStop();
    if (restore) this.restoreFocus(restore);
  }

  /** Builds entries up to the one at `index`, and a batch after it. */
  private growTo(index: number): void {
    this.limit = index + PANEL_BATCH;
    this.render();
  }

  /** Everything an entry shows, as one text. */
  private keyOf(thread: CommentThread, can: boolean, quote: string): string {
    const selected = thread.root.id === this.selected;
    const compose = selected && this.compose ? `${this.compose.mode}:${this.compose.comment}` : '';
    const found = selected && this.found ? `${this.found.comment}:${this.found.start}:${this.found.end}` : '';
    const parts: (string | number | boolean)[] = [selected, can, compose, quote, thread.root.start === undefined, found];
    const add = (comment: DocComment) => parts.push(comment.id, comment.author, comment.date || '', comment.done, comment.canChange, selected ? comment.text : shownText(comment.text, false));
    add(thread.root);
    if (selected) for (const reply of thread.replies) add(reply);
    else parts.push(thread.replies.length);
    return JSON.stringify(parts);
  }

  private button(entry: Entry, action: string, label: string, comment: number, small: boolean): HTMLButtonElement {
    const out = element('button', 'bym-doc-comment-btn', label);
    out.type = 'button';
    out.setAttribute('data-action', action);
    out.setAttribute('data-comment', String(comment));
    if (small) out.setAttribute('data-small', 'true');
    entry.buttons.set(`${action}:${comment}`, out);
    return out;
  }

  private meta(comment: DocComment, s: UiStrings, state: boolean): HTMLDivElement {
    const line = element('div', 'bym-doc-cpanel-meta');
    line.appendChild(element('span', 'bym-doc-cpanel-author', comment.author));
    const date = this.host.date(comment.date);
    if (date !== '') line.appendChild(element('span', 'bym-doc-cpanel-date', (comment.author !== '' ? ' · ' : '') + date));
    if (state && comment.done) line.appendChild(element('span', 'bym-doc-cpanel-state', ' · ' + s.commentResolved));
    return line;
  }

  private build(thread: CommentThread, can: boolean, key: string, quote: string): Entry {
    const s = this.host.strings();
    const root = thread.root;
    const selected = root.id === this.selected;
    const item = element('div', 'bym-doc-cpanel-item');
    const entry: Entry = { element: item, key, buttons: new Map() };
    item.setAttribute('role', 'listitem');
    item.setAttribute('data-thread', String(root.id));
    item.setAttribute('data-done', root.done ? 'true' : 'false');
    item.tabIndex = -1;
    if (selected) item.setAttribute('aria-current', 'true');
    item.setAttribute('aria-label', s.commentThreadLabel(root.author, this.host.date(root.date), quoteText(root.text, 200), thread.replies.length, root.done));

    item.appendChild(this.meta(root, s, true));
    if (root.start === undefined) item.appendChild(element('div', 'bym-doc-cpanel-quote', s.commentNotAnchored));
    else if (quote !== '') {
      const box = element('div', 'bym-doc-cpanel-quote', `“${quote}”`);
      box.title = `${s.commentAbout}: ${quote}`;
      item.appendChild(box);
    }
    item.appendChild(this.textOf(root, selected));
    const compose = selected ? this.compose : undefined;
    const writing = compose !== undefined;
    if (selected && can && !writing) {
      const actions = element('div', 'bym-doc-cpanel-actions');
      actions.appendChild(this.button(entry, 'edit', s.commentEdit, root.id, true));
      if (root.canChange) actions.appendChild(this.button(entry, 'delete', s.commentDelete, root.id, true));
      item.appendChild(actions);
    }
    if (compose && compose.mode === 'edit' && compose.comment === root.id) this.form(entry, compose, s);

    if (!selected) {
      if (thread.replies.length > 0) item.appendChild(element('div', 'bym-doc-cpanel-replies', s.commentReplies(thread.replies.length)));
      return entry;
    }
    for (const reply of thread.replies) {
      const box = element('div', 'bym-doc-cpanel-reply');
      box.setAttribute('data-comment', String(reply.id));
      box.appendChild(this.meta(reply, s, false));
      box.appendChild(this.textOf(reply, true));
      if (can && !writing) {
        const actions = element('div', 'bym-doc-cpanel-actions');
        actions.appendChild(this.button(entry, 'edit', s.commentEdit, reply.id, true));
        if (reply.canChange) actions.appendChild(this.button(entry, 'delete', s.commentDelete, reply.id, true));
        box.appendChild(actions);
      }
      item.appendChild(box);
      if (compose && compose.mode === 'edit' && compose.comment === reply.id) this.form(entry, compose, s);
    }
    if (compose && compose.mode === 'reply') this.form(entry, compose, s);
    if (can && !writing) {
      const foot = element('div', 'bym-doc-cpanel-foot');
      foot.style.marginTop = '4px';
      if (root.canChange) foot.appendChild(this.button(entry, 'reply', s.commentReply, root.id, false));
      foot.appendChild(this.button(entry, root.done ? 'reopen' : 'resolve', root.done ? s.commentReopen : s.commentResolve, root.id, false));
      item.appendChild(foot);
    }
    return entry;
  }

  /** The text of a comment as an entry shows it, with the match of Find in it marked. */
  private textOf(comment: DocComment, whole: boolean): HTMLDivElement {
    const out = element('div', 'bym-doc-cpanel-text');
    const found = this.found;
    const marked = whole && found && found.comment === comment.id && found.text === comment.text ? found : undefined;
    fillMarked(out, shownText(comment.text, whole), marked);
    return out;
  }

  private form(entry: Entry, compose: Compose, s: UiStrings): void {
    const form = element('div', 'bym-doc-cpanel-form');
    const field = element('textarea', 'bym-doc-comment-field');
    field.rows = 3;
    field.maxLength = 32767;
    field.setAttribute('aria-label', compose.mode === 'reply' ? s.commentReplyText : s.commentText);
    field.setAttribute('autocomplete', 'off');
    field.style.cssText = 'display:block;width:100%;min-height:64px;max-height:200px;margin:6px 0;padding:6px 8px;resize:vertical';
    field.value = compose.draft;
    entry.field = field;
    form.appendChild(field);
    form.appendChild(this.button(entry, 'save', s.commentSave, compose.comment, false));
    form.appendChild(this.button(entry, 'cancel', s.commentCancel, compose.comment, false));
    entry.element.appendChild(form);
  }

  /** Exactly one entry is a Tab stop: the selected one, else the first. */
  private keepTabStop(): void {
    let stop: HTMLElement | undefined;
    const selected = this.selected !== undefined ? this.entries.get(this.selected) : undefined;
    if (selected) stop = selected.element;
    else if (this.visible.length > 0) {
      const first = this.entries.get(this.visible[0].root.id);
      if (first) stop = first.element;
    }
    this.entries.forEach((entry) => {
      const tabIndex = entry.element === stop ? 0 : -1;
      if (entry.element.tabIndex !== tabIndex) entry.element.tabIndex = tabIndex;
    });
  }

  // --------------------------------------------------------------- focus

  /** Where the focus is in the list: the entry, one of its buttons or its field. */
  private focusedPlace(): { thread: number; what: string } | undefined {
    if (!this.hasFocus) return undefined;
    const root = this.element.getRootNode ? (this.element.getRootNode() as Document | ShadowRoot) : document;
    const active = root.activeElement;
    let found: { thread: number; what: string } | undefined;
    this.entries.forEach((entry, id) => {
      if (found || !active || !entry.element.contains(active)) return;
      if (active === entry.element) found = { thread: id, what: 'entry' };
      else if (active === entry.field) found = { thread: id, what: 'field' };
      else {
        let what = 'entry';
        entry.buttons.forEach((button, name) => {
          if (button === active) what = name;
        });
        found = { thread: id, what };
      }
    });
    return found;
  }

  private restoreFocus(place: { thread: number; what: string }): void {
    if (this.hasFocus) {
      // The element that had it is still there.
      return;
    }
    const entry = this.entries.get(place.thread);
    if (!entry) {
      this.focusList();
      return;
    }
    let target: HTMLElement | undefined;
    if (place.what === 'field') target = entry.field;
    else if (place.what !== 'entry') {
      // Resolve became Reopen: the button in the same place.
      const [action, comment] = place.what.split(':');
      target = entry.buttons.get(place.what) || (action === 'resolve' || action === 'reopen' ? entry.buttons.get(`${action === 'resolve' ? 'reopen' : 'resolve'}:${comment}`) : undefined);
    }
    this.focusElement(target || entry.field || entry.element);
  }

  /** The focus to the list: the selected entry, else the first, else the box of the filter. */
  private focusList(): void {
    const selected = this.selected !== undefined ? this.entries.get(this.selected) : undefined;
    const first = this.visible.length > 0 ? this.entries.get(this.visible[0].root.id) : undefined;
    const target = selected || first;
    this.focusElement(target ? target.element : this.filter);
  }

  private focusElement(target: HTMLElement): void {
    try {
      target.focus({ preventScroll: true });
    } catch {
      target.focus();
    }
    if (target !== this.filter) this.bringIntoSight(target);
  }

  /** Scrolls the list (and only the list) so that the element is seen. */
  private bringIntoSight(target: HTMLElement): void {
    const scroll = this.scroll;
    const view = scroll.clientHeight || 0;
    if (view <= 0 || typeof target.getBoundingClientRect !== 'function') return;
    // Where the element is on the screen against the list, in the px of the list (a page may scale the editor).
    const box = target.getBoundingClientRect();
    const frame = scroll.getBoundingClientRect();
    const scale = frame.height > 0 && scroll.offsetHeight > 0 ? scroll.offsetHeight / frame.height : 1;
    const top = (box.top - frame.top) * scale + scroll.scrollTop;
    const height = box.height * scale;
    if (top < scroll.scrollTop) scroll.scrollTop = Math.max(0, top - 8);
    else if (top + height > scroll.scrollTop + view) scroll.scrollTop = top + Math.min(height, view) - view + 8;
  }

  // ------------------------------------------------------------- commands

  private stop = (event: Event): void => {
    event.stopPropagation();
  };

  /** Closed from inside (its Close button, Escape, the shortcut): the user was in it, so the focus has to go somewhere. */
  private closeByUser(): void {
    const had = this.hasFocus;
    this.hide();
    this.host.closed(had);
  }

  private startCompose(mode: 'reply' | 'edit', comment: number): void {
    const thread = this.selected;
    if (thread === undefined || !this.host.canChange()) return;
    let text = '';
    if (mode === 'edit') {
      for (const entry of this.all) {
        if (entry.root.id !== thread) continue;
        if (entry.root.id === comment) text = entry.root.text;
        for (const reply of entry.replies) if (reply.id === comment) text = reply.text;
      }
    }
    this.compose = { mode, thread, comment, draft: text };
    this.render();
    const entry = this.entries.get(thread);
    if (entry && entry.field) this.focusElement(entry.field);
  }

  private cancelCompose(): void {
    const thread = this.compose ? this.compose.thread : undefined;
    this.compose = undefined;
    this.render();
    const entry = thread !== undefined ? this.entries.get(thread) : undefined;
    if (entry) this.focusElement(entry.element);
  }

  private save(): void {
    const compose = this.compose;
    if (!compose) return;
    // Written down before the host is asked: what it does refreshes the panel.
    const done = this.host.submit(compose.mode, compose.mode === 'reply' ? compose.thread : compose.comment, compose.draft);
    if (!done) return;
    if (this.compose === compose) this.compose = undefined;
    this.refresh();
    const entry = this.entries.get(compose.thread);
    if (entry) this.focusElement(entry.element);
    else this.focusList();
  }

  /** Picks a thread from the panel: it is selected, and the document shows its text. */
  private pick(id: number, toDocument: boolean): void {
    if (this.selected !== id) {
      this.compose = undefined;
      this.found = undefined;
      this.selected = id;
      this.render();
      this.host.selectionChanged(id);
    }
    this.host.reveal(id, toDocument);
  }

  private showMore(focusNew: boolean): void {
    const first = this.limit;
    if (first >= this.visible.length) return;
    this.limit += PANEL_BATCH;
    this.render();
    if (!focusNew) return;
    const entry = this.entries.get(this.visible[first].root.id);
    if (entry) this.focusElement(entry.element);
  }

  private entryOf(target: EventTarget | null): { id: number; entry: Entry } | undefined {
    let found: { id: number; entry: Entry } | undefined;
    const node = target as Node | null;
    if (!node) return undefined;
    this.entries.forEach((entry, id) => {
      if (!found && entry.element.contains(node)) found = { id, entry };
    });
    return found;
  }

  private onClick = (event: Event): void => {
    event.stopPropagation();
    if (this.destroyed) return;
    const target = event.target as Element | null;
    const button = target && typeof target.closest === 'function' ? (target.closest('button') as HTMLButtonElement | null) : null;
    if (button) {
      const action = button.getAttribute('data-action');
      const raw = button.getAttribute('data-comment');
      // The ids of a file may be negative (w:id is a signed number).
      const comment = raw !== null && /^-?\d+$/.test(raw) ? Number(raw) : undefined;
      switch (action) {
        case 'close':
          this.closeByUser();
          return;
        case 'more':
          this.showMore(true);
          return;
        case 'edit':
        case 'reply':
          if (comment !== undefined) this.startCompose(action, comment);
          return;
        case 'delete':
          if (comment !== undefined) {
            const thread = this.selected;
            this.host.remove(comment);
            // The answer went and the thread stays: the focus to its entry; the thread went: to the list.
            if (!this.hasFocus) {
              const entry = thread !== undefined ? this.entries.get(thread) : undefined;
              if (entry) this.focusElement(entry.element);
              else this.focusList();
            }
          }
          return;
        case 'resolve':
        case 'reopen':
          if (comment !== undefined) this.host.resolve(comment, action === 'resolve');
          return;
        case 'save':
          this.save();
          return;
        case 'cancel':
          this.cancelCompose();
          return;
      }
      return;
    }
    const hit = this.entryOf(target);
    // A press in the field or on the text of the selected thread (to select some of it) picks nothing anew.
    if (!hit || (hit.entry.field && hit.entry.field.contains(target as Node))) return;
    this.pick(hit.id, false);
    const entry = this.entries.get(hit.id);
    if (entry && !this.hasFocus) this.focusElement(entry.element);
  };

  private onInput = (event: Event): void => {
    const compose = this.compose;
    if (!compose) return;
    const entry = this.entries.get(compose.thread);
    if (entry && entry.field && event.target === entry.field) compose.draft = entry.field.value;
  };

  private onFilter = (): void => {
    this.showResolved = this.filter.checked;
    this.refresh();
  };

  private onScroll = (): void => {
    const scroll = this.scroll;
    // Near the end of what is built: the next batch.
    if (this.limit < this.visible.length && scroll.scrollTop + scroll.clientHeight >= scroll.scrollHeight - 200) this.showMore(false);
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    const command = event.ctrlKey || event.metaKey;
    const key = event.key;
    // Ctrl+P goes on: it prints the document, from wherever in the view it is pressed.
    if (!(command && !event.altKey && (key === 'p' || key === 'P'))) event.stopPropagation();
    if (this.destroyed || event.defaultPrevented) return;
    const target = event.target as Element | null;
    const entry = this.compose ? this.entries.get(this.compose.thread) : undefined;
    if (entry && entry.field && target === entry.field) {
      if (key === 'Escape') {
        event.preventDefault();
        this.cancelCompose();
      } else if (key === 'Enter' && command) {
        event.preventDefault();
        this.save();
      }
      return;
    }
    // The shortcut that opened the panel closes it.
    if (command && event.altKey && event.shiftKey && (event.code === 'KeyM' || key === 'm' || key === 'M')) {
      event.preventDefault();
      this.closeByUser();
      return;
    }
    if (key === 'Escape' && !command && !event.altKey) {
      event.preventDefault();
      this.closeByUser();
      return;
    }
    if (command || event.altKey) return;
    const hit = this.entryOf(target);
    // The keys of the list are those pressed on an entry itself: its buttons have their own.
    if (!hit || hit.entry.element !== target) return;
    if (key === 'Enter') {
      event.preventDefault();
      this.pick(hit.id, true);
      return;
    }
    if (key === ' ') {
      event.preventDefault();
      this.pick(hit.id, false);
      return;
    }
    const index = this.indexOf(hit.id);
    const next = stepIndex(index, this.visible.length, key);
    if (next === undefined) return;
    event.preventDefault();
    // pick() builds the entry of the thread it selects, wherever it is.
    const id = this.visible[next].root.id;
    this.pick(id, false);
    const moved = this.entries.get(id);
    if (moved) this.focusElement(moved.element);
  };
}
