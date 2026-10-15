// What the document view shows of comments beside the pages: the colours a
// comment is marked with, the mark in the margin, and the box ("balloon")
// that shows a comment with its answers and takes the text of a new one.
//
// The view does the work on the document; this module only builds elements
// and tells the view what the user asked for. Everything that comes from the
// file (names of authors, dates, the text of comments) reaches the DOM as
// text, never as markup. Nothing here touches the DOM at import time.

import type { DocComment } from '../ops/doc/index';
import type { UiStrings } from './strings';

/** The colours comments are marked with, one for each author (by the name), as red, green and blue. */
const COMMENT_COLORS: [number, number, number][] = [
  [251, 188, 4],
  [66, 133, 244],
  [234, 67, 53],
  [15, 157, 88],
  [171, 71, 188],
  [0, 172, 193],
  [255, 112, 67],
  [124, 179, 66],
];

/** The colour of an author: the same name has the same colour in every view. */
export function commentColor(author: string): [number, number, number] {
  let hash = 0;
  // A name is short; a hostile one is cut where the colour no longer depends on it.
  for (let i = 0; i < author.length && i < 256; i++) hash = (Math.imul(hash, 31) + author.charCodeAt(i)) | 0;
  return COMMENT_COLORS[(hash >>> 0) % COMMENT_COLORS.length];
}

/** How strongly the text of a comment is tinted: of an open one, of a resolved one, and what the one that is shown adds. */
export const COMMENT_ALPHA = { open: 0.3, resolved: 0.1, shown: 0.16 };

/** The fill the text of a comment is painted over with. */
export function commentFill(author: string, done: boolean, shown: boolean): string {
  const [r, g, b] = commentColor(author);
  return `rgba(${r}, ${g}, ${b}, ${(done ? COMMENT_ALPHA.resolved : COMMENT_ALPHA.open) + (shown ? COMMENT_ALPHA.shown : 0)})`;
}

/** A date of the file as it is shown: "2026-10-01 09:00" for ISO text, anything else as it is. */
export function shownDate(date: string | undefined): string {
  if (date === undefined) return '';
  return /^\d{4}-\d\d-\d\dT\d\d:\d\d/.test(date) ? `${date.slice(0, 10)} ${date.slice(11, 16)}` : date;
}

/** Cosmetic rules of the marks and the balloon; what their layout depends on is set inline. */
export const COMMENT_STYLE_TEXT =
  '.bym-doc-comment-mark{-webkit-appearance:none;appearance:none;background:#fff;border:2px solid #fbbc04;border-radius:50% 50% 50% 2px;color:#202124;' +
  'font:inherit;font-size:10px;line-height:1;cursor:pointer;-webkit-tap-highlight-color:transparent;touch-action:manipulation;opacity:.92}' +
  '.bym-doc-comment-mark[data-done="true"]{opacity:.5}' +
  '.bym-doc-comment-mark[aria-expanded="true"]{box-shadow:0 0 0 2px rgba(26,115,232,.45);opacity:1}' +
  '.bym-doc-comment-mark:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-doc-comment-mark:focus:not(:focus-visible){outline:none}' +
  '.bym-doc-comment-balloon{background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.2);color:#202124;' +
  'font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;-webkit-user-select:text;user-select:text}' +
  '.bym-doc-comment-item{padding:6px 0;border-top:1px solid #eceff1}' +
  '.bym-doc-comment-item:first-child{border-top:0;padding-top:0}' +
  '.bym-doc-comment-item[data-reply="true"]{margin-left:12px}' +
  '.bym-doc-comment-author{font-weight:600}' +
  '.bym-doc-comment-date,.bym-doc-comment-state{color:#5f6368;font-size:12px}' +
  '.bym-doc-comment-state{font-weight:600;color:#0f5c31}' +
  '.bym-doc-comment-body{white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;margin:2px 0}' +
  '.bym-doc-comment-btn{-webkit-appearance:none;appearance:none;box-sizing:border-box;min-height:28px;margin:2px 4px 2px 0;padding:0 10px;border:1px solid #dadce0;' +
  'border-radius:6px;background:#fff;color:#202124;font:inherit;line-height:1.2;cursor:pointer;touch-action:manipulation;-webkit-tap-highlight-color:transparent}' +
  '@media (pointer:coarse){.bym-doc-comment-btn{min-height:40px;padding:0 14px}}' +
  '@media (hover:hover){.bym-doc-comment-btn:hover{background:rgba(60,64,67,.09)}}' +
  '.bym-doc-comment-btn:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-doc-comment-btn:focus:not(:focus-visible){outline:none}' +
  '.bym-doc-comment-btn[data-action="save"]{background:#1a7f45;border-color:#1a7f45;color:#fff}' +
  '.bym-doc-comment-btn[data-small="true"]{min-height:24px;padding:0 6px;border-color:transparent;color:#5f6368;font-size:12px}' +
  '@media (pointer:coarse){.bym-doc-comment-btn[data-small="true"]{min-height:36px;padding:0 10px}}' +
  '.bym-doc-comment-field{box-sizing:border-box;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;font:inherit;outline:none}' +
  '.bym-doc-comment-field:focus{border-color:#1a7f45}' +
  // Text smaller than 16 px makes iOS zoom the page when a field takes the focus.
  '@media (pointer:coarse){.bym-doc-comment-field{font-size:16px}}';

/** What the balloon is for at the moment: showing a thread, or taking the text of a new comment, of an answer, or of a comment that is edited. */
export type BalloonMode = 'view' | 'new' | 'reply' | 'edit';

/** What the balloon asks of the view. */
export interface CommentBalloonHost {
  /** The element the balloon is put into: positioned, and what it must stay inside of. */
  readonly root: HTMLElement;
  strings(): UiStrings;
  /** Whether comments can be added and changed now. */
  canChange(): boolean;
  /** The text of the field was saved: a new comment, an answer to the thread, or the new text of the comment `id`. True when it was taken. */
  submit(mode: 'new' | 'reply' | 'edit', id: number | undefined, text: string): boolean;
  resolve(id: number, done: boolean): void;
  remove(id: number): void;
  /** The balloon was closed by the user; `refocus`: the focus was in it and has to go somewhere. */
  closed(refocus: boolean): void;
}

/** Where the balloon belongs, in px of the host's root: the right edge it lines up with, and the line of text it must not cover. */
export interface BalloonPlace {
  right: number;
  top: number;
  bottom: number;
}

/** Narrower than this, the balloon is a sheet along the bottom edge instead of a box beside the text. */
export const DOCKED_BELOW = 520;

function element<K extends keyof HTMLElementTagNameMap>(tag: K, className: string, text?: string): HTMLElementTagNameMap[K] {
  const out = document.createElement(tag);
  out.className = className;
  if (text !== undefined) out.textContent = text;
  return out;
}

export class CommentBalloon {
  readonly element: HTMLDivElement;
  private readonly host: CommentBalloonHost;
  private thread: DocComment[] = [];
  private mode: BalloonMode = 'view';
  /** The comment that is edited. */
  private editing: number | undefined;
  private field: HTMLTextAreaElement | undefined;
  private open = false;
  private place: BalloonPlace | undefined;
  private destroyed = false;

  constructor(host: CommentBalloonHost) {
    this.host = host;
    const box = element('div', 'bym-doc-comment-balloon');
    box.setAttribute('role', 'dialog');
    box.style.cssText = 'position:absolute;z-index:6;display:none;box-sizing:border-box;margin:0;padding:10px;overflow:auto;overscroll-behavior:contain;max-width:calc(100% - 16px)';
    this.element = box;
    box.addEventListener('click', this.onClick);
    box.addEventListener('keydown', this.onKeyDown);
    // What happens in the balloon is not the document's: the view listens on the element around both.
    for (const type of ['mousedown', 'copy', 'cut', 'paste', 'dblclick', 'contextmenu']) box.addEventListener(type, this.stop);
    host.root.appendChild(box);
  }

  get isOpen(): boolean {
    return this.open;
  }

  /** The first comment of the thread that is shown; undefined when a new comment is being written or nothing is shown. */
  get threadId(): number | undefined {
    return this.open && this.thread.length ? this.thread[0].id : undefined;
  }

  /** Whether text is being written in it: such a balloon does not go away by itself. */
  get composing(): boolean {
    return this.open && this.mode !== 'view';
  }

  get hasFocus(): boolean {
    const root = this.element.getRootNode() as Document | ShadowRoot;
    return this.open && this.element.contains(root.activeElement);
  }

  /** Whether the balloon lies along the bottom edge (a narrow view) instead of beside the text. */
  get docked(): boolean {
    return this.host.root.clientWidth < DOCKED_BELOW;
  }

  /**
   * Shows a thread (its first comment, then the answers), or with `mode`
   * "new" the field for a new comment. `focus`: the focus goes into the
   * balloon (to the field, else to its first button).
   */
  show(thread: DocComment[], place: BalloonPlace | undefined, options: { mode?: BalloonMode; editing?: number; focus?: boolean } = {}): void {
    if (this.destroyed) return;
    this.thread = thread;
    this.mode = options.mode || 'view';
    this.editing = this.mode === 'edit' ? options.editing : undefined;
    this.place = place;
    this.open = true;
    this.render();
    this.element.style.display = 'block';
    this.position();
    if (options.focus) this.focusInside();
  }

  /** Brings what is shown up to date with the thread as it is now, keeping what is being written. */
  refresh(thread: DocComment[]): void {
    if (!this.open || this.destroyed) return;
    const draft = this.field ? this.field.value : undefined;
    const focused = this.hasFocus;
    const root = this.element.getRootNode() as Document | ShadowRoot;
    const active = root.activeElement as HTMLElement | null;
    const action = active && this.element.contains(active) ? active.getAttribute('data-action') : null;
    this.thread = thread;
    this.render();
    if (draft !== undefined && this.field) this.field.value = draft;
    this.position();
    if (!focused) return;
    // The focus stays on the button that was pressed (Resolve became Reopen: the one in its place).
    const same = action === 'resolve' || action === 'reopen' ? this.element.querySelector('[data-action="resolve"],[data-action="reopen"]') : action ? this.element.querySelector(`[data-action="${action}"]`) : null;
    if (same instanceof HTMLElement) same.focus({ preventScroll: true });
    else this.focusInside();
  }

  moveTo(place: BalloonPlace | undefined): void {
    this.place = place;
    if (this.open) this.position();
  }

  /** Takes the balloon away. Nothing is said to the host: it asked for it. */
  hide(): void {
    if (!this.open) return;
    this.open = false;
    this.mode = 'view';
    this.field = undefined;
    this.thread = [];
    this.element.style.display = 'none';
    while (this.element.firstChild) this.element.removeChild(this.element.firstChild);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.hide();
    const box = this.element;
    box.removeEventListener('click', this.onClick);
    box.removeEventListener('keydown', this.onKeyDown);
    for (const type of ['mousedown', 'copy', 'cut', 'paste', 'dblclick', 'contextmenu']) box.removeEventListener(type, this.stop);
    if (box.parentNode) box.parentNode.removeChild(box);
  }

  private stop = (event: Event): void => {
    event.stopPropagation();
  };

  private focusInside(): void {
    const target = this.field || (this.element.querySelector('button') as HTMLElement | null);
    if (!target) return;
    try {
      target.focus({ preventScroll: true });
    } catch {
      target.focus();
    }
  }

  private button(action: string, label: string, id?: number, small = false): HTMLButtonElement {
    const out = element('button', 'bym-doc-comment-btn', label);
    out.type = 'button';
    out.setAttribute('data-action', action);
    if (id !== undefined) out.setAttribute('data-comment', String(id));
    if (small) out.setAttribute('data-small', 'true');
    return out;
  }

  private render(): void {
    const s = this.host.strings();
    const box = this.element;
    while (box.firstChild) box.removeChild(box.firstChild);
    this.field = undefined;
    const thread = this.thread;
    const can = this.host.canChange();
    const first = thread[0];
    box.setAttribute('aria-label', first ? s.commentBy(first.author) : s.newComment);
    box.setAttribute('data-mode', this.mode);
    if (first) box.setAttribute('data-thread', String(first.id));
    else box.removeAttribute('data-thread');

    const head = element('div', 'bym-doc-comment-headrow');
    head.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin:0 0 4px';
    const state = element('span', 'bym-doc-comment-state', first ? (first.done ? s.commentResolved : '') : s.newComment);
    head.appendChild(state);
    const close = this.button('close', '\u00d7', undefined, true);
    close.setAttribute('aria-label', s.commentClose);
    close.title = s.commentClose;
    close.style.marginRight = '0';
    head.appendChild(close);
    box.appendChild(head);

    if (thread.length) {
      const list = element('div', 'bym-doc-comment-list');
      for (let i = 0; i < thread.length; i++) {
        const comment = thread[i];
        const item = element('div', 'bym-doc-comment-item');
        item.setAttribute('data-comment', String(comment.id));
        if (i > 0) item.setAttribute('data-reply', 'true');
        const line = element('div', 'bym-doc-comment-head');
        line.appendChild(element('span', 'bym-doc-comment-author', comment.author));
        const date = shownDate(comment.date);
        if (date !== '') line.appendChild(element('span', 'bym-doc-comment-date', (comment.author !== '' ? ' \u00b7 ' : '') + date));
        item.appendChild(line);
        item.appendChild(element('div', 'bym-doc-comment-body', comment.text));
        if (can && this.mode === 'view') {
          const actions = element('div', 'bym-doc-comment-actions');
          actions.appendChild(this.button('edit', s.commentEdit, comment.id, true));
          if (comment.canChange) actions.appendChild(this.button('delete', s.commentDelete, comment.id, true));
          item.appendChild(actions);
        }
        list.appendChild(item);
      }
      box.appendChild(list);
    }

    if (this.mode !== 'view') {
      const form = element('div', 'bym-doc-comment-form');
      const field = element('textarea', 'bym-doc-comment-field');
      field.rows = 3;
      field.maxLength = 32767;
      field.setAttribute('aria-label', this.mode === 'reply' ? s.commentReplyText : s.commentText);
      field.setAttribute('autocomplete', 'off');
      field.style.cssText = 'display:block;width:100%;min-height:64px;max-height:200px;margin:6px 0;padding:6px 8px;resize:vertical';
      if (this.mode === 'edit') {
        const edited = thread.filter((comment) => comment.id === this.editing)[0];
        field.value = edited ? edited.text : '';
      }
      this.field = field;
      form.appendChild(field);
      form.appendChild(this.button('save', s.commentSave));
      form.appendChild(this.button('cancel', s.commentCancel));
      box.appendChild(form);
    } else if (can && first) {
      const foot = element('div', 'bym-doc-comment-foot');
      foot.style.marginTop = '4px';
      if (first.canChange) foot.appendChild(this.button('reply', s.commentReply));
      foot.appendChild(this.button(first.done ? 'reopen' : 'resolve', first.done ? s.commentReopen : s.commentResolve));
      box.appendChild(foot);
    }
  }

  /** Puts the balloon where it belongs: beside the line it is about, below it or above it, inside the root; along the bottom edge of a narrow root. */
  private position(): void {
    const root = this.host.root;
    const box = this.element;
    const style = box.style;
    const width = root.clientWidth;
    const height = root.clientHeight;
    if (this.docked || !this.place) {
      style.left = '8px';
      style.right = '8px';
      style.width = 'auto';
      style.top = 'auto';
      style.bottom = '8px';
      style.maxHeight = `${Math.max(120, Math.round(height * 0.6))}px`;
      if (!this.docked) {
        // A thread without a place on the pages that are shown: in the corner, out of the way of the text.
        style.left = 'auto';
        style.width = `${Math.min(320, Math.max(160, width - 16))}px`;
        style.right = '24px';
      }
      return;
    }
    const place = this.place;
    const boxWidth = Math.min(320, Math.max(160, width - 16));
    style.right = 'auto';
    style.bottom = 'auto';
    style.width = `${boxWidth}px`;
    style.maxHeight = `${Math.max(120, height - 16)}px`;
    const left = Math.min(Math.max(8, place.right - boxWidth), Math.max(8, width - boxWidth - 8));
    style.left = `${Math.round(left)}px`;
    const boxHeight = box.offsetHeight;
    let top = place.bottom + 6;
    // Not enough room below the line: above it; and inside the view in any case.
    if (top + boxHeight > height - 8 && place.top - 6 - boxHeight >= 8) top = place.top - 6 - boxHeight;
    top = Math.min(Math.max(8, top), Math.max(8, height - boxHeight - 8));
    style.top = `${Math.round(top)}px`;
  }

  private onClick = (event: Event): void => {
    event.stopPropagation();
    const target = event.target as Element | null;
    const button = target && typeof target.closest === 'function' ? (target.closest('button') as HTMLButtonElement | null) : null;
    if (!button || this.destroyed) return;
    const action = button.getAttribute('data-action');
    const raw = button.getAttribute('data-comment');
    // The ids of a file may be negative (w:id is a signed number).
    const id = raw !== null && /^-?\d+$/.test(raw) ? Number(raw) : undefined;
    const first = this.thread[0];
    switch (action) {
      case 'close':
        this.close();
        break;
      case 'edit':
        if (id === undefined) return;
        this.mode = 'edit';
        this.editing = id;
        this.render();
        this.position();
        this.focusInside();
        break;
      case 'reply':
        this.mode = 'reply';
        this.render();
        this.position();
        this.focusInside();
        break;
      case 'delete':
        if (id !== undefined) this.host.remove(id);
        // An answer went and the thread is still shown: the focus stays in the balloon (some browsers do not give
        // it to a button that is pressed with the mouse, and the button that was pressed is gone).
        if (this.open) this.focusInside();
        break;
      case 'resolve':
      case 'reopen':
        if (first) this.host.resolve(first.id, action === 'resolve');
        if (this.open) {
          // The button in the same place: Resolve became Reopen, or the other way round.
          const next = this.element.querySelector('[data-action="resolve"],[data-action="reopen"]');
          if (next instanceof HTMLElement) next.focus({ preventScroll: true });
        }
        break;
      case 'save':
        this.save();
        break;
      case 'cancel':
        this.cancel();
        break;
    }
  };

  private save(): void {
    const field = this.field;
    if (!field || this.mode === 'view') return;
    this.host.submit(this.mode, this.mode === 'edit' ? this.editing : this.thread.length ? this.thread[0].id : undefined, field.value);
  }

  /** Gives up what is being written: back to the thread, or away with a balloon that was only there for a new comment. */
  private cancel(): void {
    if (this.mode === 'new' || this.thread.length === 0) {
      this.close();
      return;
    }
    this.mode = 'view';
    this.editing = undefined;
    this.render();
    this.position();
    this.focusInside();
  }

  /** Closed from inside (its Close button, Cancel, Escape): the user was in it, so the focus has to go somewhere, also where a browser gave it to no button. */
  private close(): void {
    this.hide();
    this.host.closed(true);
  }

  private onKeyDown = (event: KeyboardEvent): void => {
    // The keys pressed here are the balloon's: the view would scroll with them, or take them as shortcuts.
    // Ctrl+P goes on: it prints the document, from wherever in the editor it is pressed.
    if (!((event.ctrlKey || event.metaKey) && !event.altKey && (event.key === 'p' || event.key === 'P'))) event.stopPropagation();
    if (this.destroyed || event.defaultPrevented) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      if (this.mode === 'view') this.close();
      else this.cancel();
    } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && this.mode !== 'view') {
      event.preventDefault();
      this.save();
    }
  };
}
