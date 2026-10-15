// The bar of Find and Replace, shared by the view of a workbook and the view of a
// document: a field for the text, buttons for the match before and after, a
// field for the replacement with Replace and Replace all, the options, and a line
// that says what was found. It floats at the top right of the editor and leaves
// the editor usable behind it; Escape closes it and hands the focus back.
//
// The bar only asks: the view it belongs to does the searching and says what
// came of it (setStatus). Every text reaches the DOM through textContent or the
// value of a field. Nothing here touches the DOM at import time.

import type { FindStrings } from './strings';

/** What the bar asks for. */
export interface FindBarQuery {
  text: string;
  matchCase: boolean;
  /** A whole word (document) or the entire cell (workbook). */
  whole: boolean;
  /** Document only: the body and its tables, not the headers, footers, notes and text boxes. */
  mainOnly: boolean;
  /** Workbook only. */
  lookIn: 'formulas' | 'values';
  within: 'sheet' | 'workbook';
}

export interface FindBarOptions {
  /** The element the bar is put into: the root element of the editor. */
  root: HTMLElement;
  /** The element the bar stands at the top of (the view in an editor with a toolbar); the root by default. */
  below?: HTMLElement;
  strings: FindStrings;
  kind: 'sheet' | 'doc';
  /** Open with the field of the replacement shown. */
  openReplace: boolean;
  /** The text to start with (what is selected), when there is one. */
  initial?: string;
  /** Whether replacing is possible now (the view can be edited). */
  canReplace(): boolean;
  find(query: FindBarQuery, backward: boolean): void;
  replace(query: FindBarQuery, replacement: string): void;
  replaceAll(query: FindBarQuery, replacement: string): void;
  /** The query changed: the view counts the matches anew, and says how many with setStatus(). */
  changed(query: FindBarQuery): void;
  /** The bar was closed by the user (Escape, the close button). */
  onClose(): void;
}

/** Events the editor around the bar must not take for its own. */
const CONTAINED = ['keypress', 'keyup', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];
const stop = (event: Event): void => event.stopPropagation();

let nextId = 1;

export class FindBar {
  readonly element: HTMLDivElement;
  private readonly options: FindBarOptions;
  private strings: FindStrings;
  private readonly findField: HTMLInputElement;
  private readonly replaceField: HTMLInputElement;
  private readonly replaceRow: HTMLDivElement;
  private readonly toggle: HTMLButtonElement;
  private readonly status: HTMLDivElement;
  private readonly matchCase: HTMLInputElement;
  private readonly whole: HTMLInputElement;
  private readonly mainOnly: HTMLInputElement | undefined;
  private readonly lookIn: HTMLSelectElement | undefined;
  private readonly within: HTMLSelectElement | undefined;
  private readonly labels: { node: HTMLElement; text: (s: FindStrings) => string; attr?: string }[] = [];
  private readonly replaceButtons: HTMLButtonElement[] = [];
  /** The position the root had, when the bar had to make it a positioned element. */
  private unplaced: string | undefined;
  private showReplace: boolean;
  private changeTimer: ReturnType<typeof setTimeout> | undefined;
  private closed = false;

  constructor(options: FindBarOptions) {
    this.options = options;
    this.strings = options.strings;
    this.showReplace = options.openReplace;
    const id = nextId++;
    const coarse = typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;
    const height = coarse ? 40 : 28;

    const bar = document.createElement('div');
    bar.className = 'bym-find';
    bar.setAttribute('role', 'dialog');
    bar.setAttribute('aria-modal', 'false');
    bar.style.cssText =
      'position:absolute;top:4px;right:8px;z-index:30;box-sizing:border-box;width:400px;max-width:calc(100% - 16px);display:flex;flex-direction:column;gap:6px;' +
      'padding:8px;background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.18);color:#202124;' +
      'font:13px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;text-align:left;direction:ltr';
    this.element = bar;
    this.label(bar, (s) => s.title, 'aria-label');

    const fieldCss = `-webkit-appearance:none;appearance:none;box-sizing:border-box;flex:1 1 auto;min-width:0;height:${height}px;margin:0;padding:0 8px;font:inherit;${coarse ? 'font-size:16px;' : ''}border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;outline-offset:-1px`;
    const buttonCss = `-webkit-appearance:none;appearance:none;box-sizing:border-box;flex:none;height:${height}px;min-width:${height}px;margin:0;padding:0 8px;font:inherit;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;cursor:pointer;white-space:nowrap`;
    const extras: HTMLElement[] = [];
    const row = (): HTMLDivElement => {
      const element = document.createElement('div');
      element.style.cssText = 'display:flex;gap:4px;align-items:center;flex-wrap:wrap;min-width:0';
      return element;
    };
    const field = (name: string, text: (s: FindStrings) => string): HTMLInputElement => {
      const input = document.createElement('input');
      input.type = 'text';
      input.className = `bym-find-${name}`;
      input.autocomplete = 'off';
      input.spellcheck = false;
      input.setAttribute('autocapitalize', 'off');
      input.maxLength = 1000;
      input.style.cssText = fieldCss;
      this.label(input, text, 'aria-label');
      this.label(input, text, 'placeholder');
      return input;
    };
    const button = (name: string, glyph: string | undefined, text: (s: FindStrings) => string): HTMLButtonElement => {
      const element = document.createElement('button');
      element.type = 'button';
      element.className = `bym-find-${name}`;
      element.setAttribute('data-bym-action', name);
      element.style.cssText = buttonCss;
      if (glyph !== undefined) {
        element.textContent = glyph;
        this.label(element, text, 'aria-label');
        this.label(element, text, 'title');
      } else this.label(element, text);
      return element;
    };
    const check = (name: string, text: (s: FindStrings) => string): HTMLInputElement => {
      const label = document.createElement('label');
      label.style.cssText = 'display:inline-flex;align-items:center;gap:4px;white-space:nowrap;cursor:pointer';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.className = `bym-find-${name}`;
      box.style.cssText = 'margin:0';
      const caption = document.createElement('span');
      this.label(caption, text);
      label.appendChild(box);
      label.appendChild(caption);
      extras.push(label);
      return box;
    };
    const choice = (name: string, caption: (s: FindStrings) => string, values: [string, (s: FindStrings) => string][]): HTMLSelectElement => {
      const label = document.createElement('label');
      label.style.cssText = 'display:inline-flex;align-items:center;gap:4px;white-space:nowrap';
      const span = document.createElement('span');
      this.label(span, caption);
      const select = document.createElement('select');
      select.className = `bym-find-${name}`;
      select.style.cssText = `box-sizing:border-box;height:${height}px;font:inherit;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;padding:0 4px`;
      for (const [value, text] of values) {
        const option = document.createElement('option');
        option.value = value;
        this.label(option, text);
        select.appendChild(option);
      }
      label.appendChild(span);
      label.appendChild(select);
      extras.push(label);
      return select;
    };

    // The text, the way through the matches, the replacement shown or not, and closing.
    const first = row();
    first.style.flexWrap = 'nowrap';
    this.findField = field('text', (s) => s.findLabel);
    this.findField.id = `bym-find-text-${id}`;
    if (options.initial) this.findField.value = options.initial.slice(0, 1000);
    first.appendChild(this.findField);
    const previous = button('previous', '↑', (s) => s.previous);
    const next = button('next', '↓', (s) => s.next);
    this.toggle = button('toggle', '⇄', (s) => s.toggleReplace);
    const close = button('close', '×', (s) => s.close);
    first.appendChild(previous);
    first.appendChild(next);
    first.appendChild(this.toggle);
    first.appendChild(close);
    bar.appendChild(first);

    // The replacement.
    const second = row();
    second.style.flexWrap = 'nowrap';
    this.replaceField = field('replacement', (s) => s.replaceLabel);
    second.appendChild(this.replaceField);
    const replaceOne = button('replace', undefined, (s) => s.replace);
    const replaceAll = button('replace-all', undefined, (s) => s.replaceAll);
    this.replaceButtons.push(replaceOne, replaceAll);
    second.appendChild(replaceOne);
    second.appendChild(replaceAll);
    this.replaceRow = second;
    bar.appendChild(second);

    // The options.
    this.matchCase = check('match-case', (s) => s.matchCase);
    this.whole = check('whole', (s) => (options.kind === 'sheet' ? s.entireCell : s.wholeWord));
    if (options.kind === 'doc') this.mainOnly = check('main-only', (s) => s.mainOnly);
    if (options.kind === 'sheet') {
      this.lookIn = choice('look-in', (s) => s.lookIn, [
        ['formulas', (s) => s.lookInFormulas],
        ['values', (s) => s.lookInValues],
      ]);
      this.within = choice('within', (s) => s.within, [
        ['sheet', (s) => s.withinSheet],
        ['workbook', (s) => s.withinWorkbook],
      ]);
    }
    const third = row();
    third.style.columnGap = '12px';
    for (const option of extras) third.appendChild(option);
    bar.appendChild(third);

    const status = document.createElement('div');
    status.className = 'bym-find-status';
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    status.style.cssText = 'min-height:1.3em;color:#5f6368;font-size:12px;overflow-wrap:anywhere';
    bar.appendChild(status);
    this.status = status;

    for (const type of CONTAINED) bar.addEventListener(type, stop);
    bar.addEventListener('keydown', this.onKey);
    bar.addEventListener('click', (event) => {
      const target = event.target as Element | null;
      const action = target && typeof target.closest === 'function' ? target.closest('[data-bym-action]') : null;
      if (!action) return;
      const name = action.getAttribute('data-bym-action');
      if (name === 'previous') this.go(true);
      else if (name === 'next') this.go(false);
      else if (name === 'toggle') this.setReplace(!this.showReplace, true);
      else if (name === 'close') this.options.onClose();
      else if (name === 'replace') this.replaceOnce();
      else if (name === 'replace-all') this.replaceEverywhere();
    });
    const watch = (): void => this.queryChanged();
    this.findField.addEventListener('input', watch);
    for (const control of [this.matchCase, this.whole, this.mainOnly, this.lookIn, this.within]) if (control) control.addEventListener('change', watch);

    const root = options.root;
    if (typeof window.getComputedStyle === 'function' && window.getComputedStyle(root).position === 'static') {
      this.unplaced = root.style.position;
      root.style.position = 'relative';
    }
    root.appendChild(bar);
    // At the top of the view, under the toolbar of an editor.
    const below = options.below;
    if (below && below !== root && typeof below.getBoundingClientRect === 'function') {
      const offset = below.getBoundingClientRect().top - root.getBoundingClientRect().top;
      if (offset > 0 && offset < root.clientHeight) bar.style.top = `${Math.round(offset + 4)}px`;
    }
    this.applyStrings();
    this.setReplace(this.showReplace, false);
    this.refresh();
  }

  /** The words of another language. */
  setStrings(strings: FindStrings): void {
    this.strings = strings;
    this.applyStrings();
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  get query(): FindBarQuery {
    return {
      text: this.findField.value,
      matchCase: this.matchCase.checked,
      whole: this.whole.checked,
      mainOnly: !!this.mainOnly && this.mainOnly.checked,
      lookIn: this.lookIn && this.lookIn.value === 'values' && !this.showReplace ? 'values' : 'formulas',
      within: this.within && this.within.value === 'workbook' ? 'workbook' : 'sheet',
    };
  }

  /** Puts the focus in a field: the text (selected), or the replacement. */
  focus(which: 'find' | 'replace' = 'find'): void {
    if (this.closed) return;
    const field = which === 'replace' && this.showReplace ? this.replaceField : this.findField;
    field.focus({ preventScroll: true });
    if (field === this.findField) field.select();
  }

  /** Takes a text to look for, as opened again with something selected. */
  setText(text: string): void {
    if (this.closed || text === '') return;
    if (this.findField.value !== text.slice(0, 1000)) {
      this.findField.value = text.slice(0, 1000);
      this.queryChanged();
    }
  }

  /** Whether the focus is in the bar. */
  get hasFocus(): boolean {
    const root = this.element.getRootNode ? (this.element.getRootNode() as Document | ShadowRoot) : document;
    return !this.closed && !!root.activeElement && this.element.contains(root.activeElement);
  }

  setReplace(on: boolean, focus = false): void {
    this.showReplace = on;
    this.replaceRow.style.display = on ? 'flex' : 'none';
    this.toggle.setAttribute('aria-expanded', on ? 'true' : 'false');
    this.toggle.setAttribute('aria-pressed', on ? 'true' : 'false');
    // Replacing reads formulas, as in Excel: what to look in is not a choice then.
    if (this.lookIn) this.lookIn.disabled = on;
    this.refresh();
    if (focus) this.focus(on ? 'replace' : 'find');
  }

  /** Says what came of a search or a replacement. */
  setStatus(text: string): void {
    if (this.closed) return;
    if (this.status.textContent !== text) this.status.textContent = text;
  }

  /** Brings the buttons in line with whether the view can be changed. */
  refresh(): void {
    if (this.closed) return;
    const can = this.options.canReplace();
    for (const button of this.replaceButtons) if (button.disabled === can) button.disabled = !can;
    this.replaceField.readOnly = !can;
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    if (this.changeTimer !== undefined) clearTimeout(this.changeTimer);
    const root = this.options.root;
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    if (this.unplaced !== undefined) root.style.position = this.unplaced;
  }

  // ---------------------------------------------------------------- inside

  private label(node: HTMLElement, text: (s: FindStrings) => string, attr?: string): void {
    this.labels.push({ node, text, attr });
  }

  private applyStrings(): void {
    for (const entry of this.labels) {
      const value = entry.text(this.strings);
      if (entry.attr) entry.node.setAttribute(entry.attr, value);
      else entry.node.textContent = value;
    }
  }

  private go(backward: boolean): void {
    if (this.findField.value === '') {
      this.findField.focus({ preventScroll: true });
      return;
    }
    this.options.find(this.query, backward);
  }

  private replaceOnce(): void {
    if (!this.options.canReplace() || this.findField.value === '') return;
    this.options.replace(this.query, this.replaceField.value);
  }

  private replaceEverywhere(): void {
    if (!this.options.canReplace() || this.findField.value === '') return;
    this.options.replaceAll(this.query, this.replaceField.value);
  }

  private queryChanged(): void {
    if (this.changeTimer !== undefined) clearTimeout(this.changeTimer);
    // Counted once the typing pauses.
    this.changeTimer = setTimeout(() => {
      this.changeTimer = undefined;
      if (!this.closed) this.options.changed(this.query);
    }, 200);
  }

  private onKey = (event: KeyboardEvent): void => {
    event.stopPropagation();
    if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
    const command = event.ctrlKey || event.metaKey;
    const key = event.key;
    if (key === 'Escape') {
      event.preventDefault();
      this.options.onClose();
      return;
    }
    if (command && !event.altKey && (key === 'f' || key === 'F')) {
      event.preventDefault();
      this.focus('find');
      return;
    }
    if (command && !event.altKey && (key === 'h' || key === 'H')) {
      event.preventDefault();
      this.setReplace(true, true);
      return;
    }
    if (key !== 'Enter' || command || event.altKey) return;
    const target = event.target as Element | null;
    if (target === this.findField) {
      event.preventDefault();
      this.go(event.shiftKey);
    } else if (target === this.replaceField) {
      event.preventDefault();
      this.replaceOnce();
    }
  };
}
