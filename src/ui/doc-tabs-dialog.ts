// The dialog of tab stops of a document (Paragraph -> Tabs): the stops of the paragraph at the
// caret in a list, a stop set with its position, its alignment and its leader, stops cleared
// one by one or all at once, and the distance between the automatic stops of the document.
// Nothing is changed until "OK": then what was done in the dialog is handed over as one change
// of tab stops (only what differs from what the paragraph had) and the default distance.
//
// It covers the editor inside its container, as the question of a conflict does; the editor
// behind it takes no input while it is open. Escape cancels; Enter in the position field sets
// the stop. Every text reaches the DOM through textContent. Nothing here touches the DOM at
// import time.

import type { TabStop } from '../model/doc/types';
import type { TabStopsPatch } from '../ops/doc/format';

export interface TabStopsStrings {
  title: string;
  /** The label of the list of stops, and what it says when there are none. */
  stops: string;
  none: string;
  position: string;
  alignment: string;
  leader: string;
  aligns: { left: string; center: string; right: string; decimal: string; bar: string };
  leaders: { none: string; dot: string; hyphen: string; underscore: string; heavy: string; middleDot: string };
  set: string;
  /** The button of a stop of the list: `stop` names it. */
  clear(stop: string): string;
  clearAll: string;
  defaultStops: string;
  ok: string;
  cancel: string;
  /** "cm" or "in": the unit positions are given in. */
  unit: 'cm' | 'in';
  /** The decimal separator of numbers typed and shown. */
  decimal: string;
  badPosition: string;
  badDefault: string;
  /** Under the field of the default distance where it cannot be changed: the document is edited together. */
  defaultShared: string;
}

export interface TabStopsDialogOptions {
  /** The element the dialog covers and is put into: the root element of the editor. */
  root: HTMLElement;
  strings: TabStopsStrings;
  /** The stops the paragraph has now (twips). */
  stops: TabStop[];
  /** The distance between the automatic stops of the document now (twips). */
  defaultTab: number;
  /** The distance cannot be changed (the document is edited together, where it does not travel): its field is shown, not taken. */
  defaultFixed?: boolean;
  /**
   * Called on "OK" with the change of the stops (undefined: none) and the distance between the
   * automatic stops (undefined: unchanged). Returns false to keep the dialog open.
   */
  apply(patch: TabStopsPatch | undefined, defaultTab: number | undefined): boolean;
  onClose(): void;
}

/** Twips in a unit. */
const TWIPS: { [unit in TabStopsStrings['unit']]: number } = { cm: 1440 / 2.54, in: 1440 };
/** How far a stop can be, twips (see MAX_TAB_POS of the commands). */
const MAX_POS = 31680;
const MAX_STOPS = 64;
const ALIGNS = ['left', 'center', 'right', 'decimal', 'bar'] as const;
const LEADERS = ['none', 'dot', 'hyphen', 'underscore'] as const;

const CONTAINED = ['keypress', 'keyup', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];
const stop = (event: Event): void => event.stopPropagation();

let nextId = 1;

/** A number as typed: digits with one decimal separator ("," or "." as the language has it, the other one too). */
export function parseLength(text: string, decimal: string): number | undefined {
  const trimmed = text.trim().replace(/\s+/g, '');
  if (!/^-?\d{1,6}([.,]\d{0,4})?$/.test(trimmed)) return undefined;
  // Either separator is taken, whichever the language writes.
  void decimal;
  const value = Number(trimmed.replace(',', '.'));
  return Number.isFinite(value) ? value : undefined;
}

/** A position in twips as the dialog shows it: in the unit, with two decimals at most. */
export function formatLength(twips: number, unit: TabStopsStrings['unit'], decimal: string): string {
  const value = Math.round((twips / TWIPS[unit]) * 100) / 100;
  return String(value).replace('.', decimal);
}

/** A position typed in the unit, in twips; undefined when it is no position a stop can have. */
export function twipsOf(text: string, unit: TabStopsStrings['unit'], decimal: string, min = -MAX_POS): number | undefined {
  const value = parseLength(text, decimal);
  if (value === undefined) return undefined;
  const twips = Math.round(value * TWIPS[unit]);
  return twips >= min && twips <= MAX_POS ? twips : undefined;
}

/**
 * The default distance the dialog hands over: undefined when its field shows what it showed when
 * the dialog opened (the shown value is rounded: read back it would be another distance) or the
 * distance it has; null when what was typed is no distance.
 */
export function defaultTabChange(shown: string, typed: string, current: number, unit: TabStopsStrings['unit'], decimal: string): number | undefined | null {
  if (typed.trim() === shown.trim()) return undefined;
  const twips = twipsOf(typed, unit, decimal, 1);
  if (twips === undefined) return null;
  return twips === current ? undefined : twips;
}

/**
 * What the dialog hands over: the stops that were set or changed, the positions whose stops
 * went, and whether everything was cleared first. Undefined when the stops are what they were.
 */
export function tabStopsChange(before: TabStop[], after: TabStop[], clearedAll: boolean): TabStopsPatch | undefined {
  const was = new Map<number, TabStop>();
  for (const tab of before) was.set(tab.pos, tab);
  const now = new Map<number, TabStop>();
  for (const tab of after) now.set(tab.pos, tab);
  const set: TabStop[] = [];
  for (const tab of after) {
    const old = was.get(tab.pos);
    if (clearedAll || !old || old.align !== tab.align || (old.leader || 'none') !== (tab.leader || 'none')) set.push(tab.leader && tab.leader !== 'none' ? { pos: tab.pos, align: tab.align, leader: tab.leader } : { pos: tab.pos, align: tab.align });
  }
  const clear: number[] = [];
  if (!clearedAll) for (const pos of was.keys()) if (!now.has(pos)) clear.push(pos);
  if (!clearedAll && set.length === 0 && clear.length === 0) return undefined;
  const patch: TabStopsPatch = {};
  if (clearedAll) patch.clearAll = true;
  if (set.length > 0) patch.set = set;
  if (clear.length > 0) patch.clear = clear;
  return patch;
}

/** The alignment a dialog shows for a stop of the file ("start" and "num" are left ones, "end" a right one). */
const shownAlign = (align: string): (typeof ALIGNS)[number] => (align === 'end' ? 'right' : align === 'start' || align === 'num' ? 'left' : (ALIGNS as readonly string[]).indexOf(align) >= 0 ? (align as (typeof ALIGNS)[number]) : 'left');

export class TabStopsDialog {
  private readonly options: TabStopsDialogOptions;
  private readonly backdrop: HTMLDivElement;
  private readonly list: HTMLUListElement;
  private readonly position: HTMLInputElement;
  private readonly align: HTMLSelectElement;
  private readonly leader: HTMLSelectElement;
  private readonly defaultField: HTMLInputElement;
  private readonly message: HTMLDivElement;
  private readonly silenced: Element[] = [];
  /** The position the root had, when the dialog had to make it a positioned element. */
  private unplaced: string | undefined;
  private readonly before: TabStop[];
  private stops: TabStop[];
  private clearedAll = false;
  /** What the field of the default distance showed when the dialog opened. */
  private readonly defaultShown: string;
  private closed = false;

  constructor(options: TabStopsDialogOptions) {
    this.options = options;
    const s = options.strings;
    this.before = options.stops.map((tab) => ({ ...tab }));
    this.stops = this.before.map((tab) => ({ ...tab }));
    const id = nextId++;
    const coarse = typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;

    const backdrop = document.createElement('div');
    backdrop.className = 'bym-tabs-backdrop';
    backdrop.style.cssText =
      'position:absolute;left:0;top:0;right:0;bottom:0;z-index:40;display:flex;align-items:center;justify-content:center;box-sizing:border-box;' +
      'padding:8px;overflow:hidden;background:rgba(32,33,36,.32);color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
      'letter-spacing:normal;text-transform:none;text-align:left;direction:ltr';
    const panel = document.createElement('div');
    panel.className = 'bym-tabs-dialog';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'true');
    panel.style.cssText =
      'display:flex;flex-direction:column;box-sizing:border-box;width:360px;max-width:100%;max-height:100%;min-width:0;overflow:auto;' +
      'background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.28);padding:12px 16px';
    const title = document.createElement('div');
    title.id = `bym-tabs-title-${id}`;
    title.style.cssText = 'font-size:15px;font-weight:600;margin:0 0 8px';
    title.textContent = s.title;
    panel.setAttribute('aria-labelledby', title.id);
    panel.appendChild(title);

    const controlCss = `box-sizing:border-box;height:${coarse ? 40 : 30}px;font:inherit;border:1px solid #dadce0;border-radius:4px;padding:0 6px;background:#fff;color:#202124;min-width:0`;
    const buttonCss = `-webkit-appearance:none;appearance:none;box-sizing:border-box;height:${coarse ? 40 : 30}px;padding:0 12px;font:inherit;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;cursor:pointer;white-space:nowrap`;
    const labelled = (text: string, control: HTMLElement): HTMLLabelElement => {
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;flex-direction:column;gap:2px;flex:1 1 0;min-width:0';
      const caption = document.createElement('span');
      caption.textContent = text;
      label.appendChild(caption);
      label.appendChild(control);
      return label;
    };
    const select = (values: readonly string[], names: { [key: string]: string }, className: string): HTMLSelectElement => {
      const element = document.createElement('select');
      element.className = className;
      element.style.cssText = controlCss;
      for (const value of values) {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = names[value];
        element.appendChild(option);
      }
      return element;
    };
    const button = (text: string, name: string): HTMLButtonElement => {
      const element = document.createElement('button');
      element.type = 'button';
      element.className = `bym-tabs-${name}`;
      element.setAttribute('data-bym-action', name);
      element.style.cssText = buttonCss;
      element.textContent = text;
      return element;
    };

    // The stops, each with its button that clears it.
    const listLabel = document.createElement('div');
    listLabel.id = `bym-tabs-list-${id}`;
    listLabel.textContent = s.stops;
    panel.appendChild(listLabel);
    const list = document.createElement('ul');
    list.className = 'bym-tabs-list';
    list.setAttribute('aria-labelledby', listLabel.id);
    list.style.cssText = 'list-style:none;margin:4px 0 8px;padding:4px;border:1px solid #dadce0;border-radius:4px;max-height:140px;overflow:auto;min-height:24px';
    panel.appendChild(list);
    this.list = list;

    // A stop to set.
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap';
    const position = document.createElement('input');
    position.type = 'text';
    position.className = 'bym-tabs-position';
    position.setAttribute('inputmode', 'decimal');
    position.autocomplete = 'off';
    position.spellcheck = false;
    position.style.cssText = controlCss;
    this.position = position;
    row.appendChild(labelled(`${s.position} (${s.unit})`, position));
    this.align = select(ALIGNS, s.aligns, 'bym-tabs-align');
    row.appendChild(labelled(s.alignment, this.align));
    this.leader = select(LEADERS, s.leaders, 'bym-tabs-leader');
    row.appendChild(labelled(s.leader, this.leader));
    panel.appendChild(row);
    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex;gap:8px;margin:8px 0;flex-wrap:wrap';
    const setButton = button(s.set, 'set');
    const clearAll = button(s.clearAll, 'clear-all');
    actions.appendChild(setButton);
    actions.appendChild(clearAll);
    panel.appendChild(actions);

    const defaultField = document.createElement('input');
    defaultField.type = 'text';
    defaultField.className = 'bym-tabs-default';
    defaultField.setAttribute('inputmode', 'decimal');
    defaultField.autocomplete = 'off';
    defaultField.spellcheck = false;
    defaultField.style.cssText = controlCss;
    defaultField.value = formatLength(options.defaultTab, s.unit, s.decimal);
    this.defaultShown = defaultField.value;
    this.defaultField = defaultField;
    panel.appendChild(labelled(`${s.defaultStops} (${s.unit})`, defaultField));
    if (options.defaultFixed) {
      // Said beside the field, so that the stops of the dialog are not lost with a change that would be refused.
      defaultField.disabled = true;
      const why = document.createElement('div');
      why.className = 'bym-tabs-fixed';
      why.id = `bym-tabs-fixed-${id}`;
      why.style.cssText = 'color:#5f6368;font-size:12px;margin:2px 0 0';
      why.textContent = s.defaultShared;
      defaultField.setAttribute('aria-describedby', why.id);
      panel.appendChild(why);
    }

    const message = document.createElement('div');
    message.className = 'bym-tabs-message';
    message.setAttribute('role', 'alert');
    message.style.cssText = 'min-height:18px;margin:4px 0;color:#b3261e;font-size:12px';
    this.message = message;
    panel.appendChild(message);

    const footer = document.createElement('div');
    footer.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';
    const ok = button(s.ok, 'ok');
    ok.style.background = '#1a73e8';
    ok.style.borderColor = '#1a73e8';
    ok.style.color = '#fff';
    const cancel = button(s.cancel, 'cancel');
    footer.appendChild(cancel);
    footer.appendChild(ok);
    panel.appendChild(footer);

    backdrop.appendChild(panel);
    this.backdrop = backdrop;

    setButton.addEventListener('click', () => this.setStop());
    clearAll.addEventListener('click', () => {
      this.stops = [];
      this.clearedAll = true;
      this.render();
    });
    ok.addEventListener('click', () => this.confirm());
    cancel.addEventListener('click', () => this.close());
    position.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.setStop();
      }
    });
    defaultField.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.confirm();
      }
    });
    backdrop.addEventListener('keydown', this.onKeyDown);
    backdrop.addEventListener('mousedown', this.onMouseDown);
    for (const name of CONTAINED) backdrop.addEventListener(name, stop);

    // The editor behind the dialog takes no input while it is open.
    const root = options.root;
    for (let i = 0; i < root.children.length; i++) {
      const child = root.children[i];
      if (child.hasAttribute('inert')) continue;
      child.setAttribute('inert', '');
      this.silenced.push(child);
    }
    if (typeof window.getComputedStyle === 'function' && window.getComputedStyle(root).position === 'static') {
      this.unplaced = root.style.position;
      root.style.position = 'relative';
    }
    root.appendChild(backdrop);
    this.render();
    position.focus({ preventScroll: true });
  }

  /**
   * Says what is wrong and gives the field it is about the focus: a press on a button does not focus it in every
   * browser (Safari), and a focus left outside the dialog would leave Escape to nobody.
   */
  private refuse(message: string, field: HTMLInputElement): void {
    this.message.textContent = message;
    field.focus({ preventScroll: true });
    if (typeof field.select === 'function') field.select();
  }

  /**
   * A press on anything but a field keeps the focus where it is: Safari gives a pressed button no focus and takes it
   * from the field that had it, which would leave the dialog without the keys (Escape) it answers to.
   */
  private onMouseDown = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    const tag = target && target.tagName ? target.tagName.toLowerCase() : '';
    if (tag === 'input' || tag === 'select' || tag === 'option' || tag === 'label') return;
    const active = (this.backdrop.getRootNode() as Document | ShadowRoot).activeElement;
    if (active && this.backdrop.contains(active)) event.preventDefault();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    event.stopPropagation();
    if (event.key === 'Escape' || event.key === 'Esc') {
      event.preventDefault();
      this.close();
    }
  };

  /** Shows the stops of the list as they are now. */
  private render(): void {
    const s = this.options.strings;
    while (this.list.firstChild) this.list.removeChild(this.list.firstChild);
    if (this.stops.length === 0) {
      const empty = document.createElement('li');
      empty.style.cssText = 'color:#5f6368;padding:2px 4px';
      empty.textContent = s.none;
      this.list.appendChild(empty);
      return;
    }
    for (const tab of this.stops) {
      const item = document.createElement('li');
      item.className = 'bym-tabs-stop';
      item.style.cssText = 'display:flex;justify-content:space-between;align-items:center;gap:8px;padding:2px 4px';
      const text = document.createElement('span');
      const align = s.aligns[shownAlign(tab.align)];
      const leader = tab.leader && tab.leader !== 'none' && (s.leaders as { [key: string]: string })[tab.leader] ? ` · ${(s.leaders as { [key: string]: string })[tab.leader]}` : '';
      const name = `${formatLength(tab.pos, s.unit, s.decimal)} ${s.unit} · ${align}${leader}`;
      text.textContent = name;
      item.appendChild(text);
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'bym-tabs-clear';
      remove.setAttribute('data-bym-action', 'clear');
      remove.setAttribute('aria-label', s.clear(name));
      remove.title = s.clear(name);
      remove.textContent = '×';
      remove.style.cssText = 'border:none;background:none;font:inherit;font-size:16px;cursor:pointer;color:#5f6368;padding:0 6px';
      remove.addEventListener('click', () => {
        this.stops = this.stops.filter((other) => other !== tab);
        this.render();
      });
      item.appendChild(remove);
      this.list.appendChild(item);
    }
  }

  /** Puts the stop of the fields into the list, in the place of one at that position. */
  private setStop(): void {
    const s = this.options.strings;
    const pos = twipsOf(this.position.value, s.unit, s.decimal);
    if (pos === undefined) {
      this.refuse(s.badPosition, this.position);
      return;
    }
    const others = this.stops.filter((tab) => tab.pos !== pos);
    if (others.length >= MAX_STOPS) {
      this.refuse(s.badPosition, this.position);
      return;
    }
    this.message.textContent = '';
    const tab: TabStop = { pos, align: this.align.value };
    if (this.leader.value !== 'none') tab.leader = this.leader.value;
    this.stops = others.concat([tab]).sort((a, b) => a.pos - b.pos);
    this.position.value = '';
    this.render();
    this.position.focus({ preventScroll: true });
  }

  private confirm(): void {
    const s = this.options.strings;
    // A position typed and not set yet is set with "OK", as Word does.
    if (this.position.value.trim() !== '') {
      this.setStop();
      if (this.message.textContent) return;
    }
    // The distance is only changed when its field was: what it shows is rounded, and read back it would be another one.
    const defaultTab = this.options.defaultFixed ? undefined : defaultTabChange(this.defaultShown, this.defaultField.value, this.options.defaultTab, s.unit, s.decimal);
    if (defaultTab === null) {
      this.refuse(s.badDefault, this.defaultField);
      return;
    }
    const patch = tabStopsChange(this.before, this.stops, this.clearedAll);
    if ((patch || defaultTab !== undefined) && !this.options.apply(patch, defaultTab)) return;
    this.close();
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.backdrop.removeEventListener('keydown', this.onKeyDown);
    this.backdrop.removeEventListener('mousedown', this.onMouseDown);
    for (const name of CONTAINED) this.backdrop.removeEventListener(name, stop);
    for (const element of this.silenced) element.removeAttribute('inert');
    this.silenced.length = 0;
    if (this.backdrop.parentNode) this.backdrop.parentNode.removeChild(this.backdrop);
    if (this.unplaced !== undefined) this.options.root.style.position = this.unplaced;
    this.options.onClose();
  }
}
