// What drops down from a cell of the grid: the menu of a filter button (sort,
// search, a list of values with check marks) and the list of a data
// validation. Both are placed inside the view, next to the cell they belong
// to, and are built for a mouse, a finger and the keyboard alike.
//
// The lists are virtual: only the rows in sight are elements, so ten thousand
// values cost what twenty do. Text that comes from a file only ever reaches
// the DOM through textContent.

export interface PanelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Space kept between a panel and the edges of its parent, and between it and its cell. */
const GAP = 4;
const PANEL_WIDTH = 264;
/** Rows of a list shown at once at most; fewer when the view is small. */
const MAX_VISIBLE_ROWS = 9;

let nextId = 0;

/** Whether the main pointer is a finger: rows and buttons are made larger then. */
function coarsePointer(): boolean {
  try {
    return typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;
  } catch (error) {
    return false;
  }
}

/**
 * Calls `onOutside` for every press outside `element`. Presses are listened for on the
 * document and on every shadow root between it and the element: a listener on the
 * document alone only ever sees the shadow host as the target. Returns what stops it.
 */
function watchOutsidePresses(element: HTMLElement, onPress: () => void, onOutside: () => void): () => void {
  // Pointer events where there are any; the mouse events a browser makes up after a touch would count twice.
  const types = 'PointerEvent' in window ? ['pointerdown'] : ['mousedown', 'touchstart'];
  const watched: { root: Node; listener: (event: Event) => void }[] = [];
  let root: Node = element.getRootNode();
  let inner: Element | undefined;
  // Ends at the document; the bound only guards against a tree that is not one.
  for (let depth = 0; depth < 64; depth++) {
    const through = inner;
    const listener = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      onPress();
      // A press on the host that leads towards the element is judged by the listener inside it.
      if (through ? target !== through : !element.contains(target)) onOutside();
    };
    for (let i = 0; i < types.length; i++) root.addEventListener(types[i], listener, true);
    watched.push({ root, listener });
    const host = (root as ShadowRoot).host;
    if (root.nodeType !== 11 || !host) break;
    inner = host;
    root = host.getRootNode();
  }
  return () => {
    for (let i = 0; i < watched.length; i++) for (let k = 0; k < types.length; k++) watched[i].root.removeEventListener(types[k], watched[i].listener, true);
  };
}

/**
 * Puts a panel under its cell, or above it when there is more room there, and never
 * outside its parent: a panel larger than the room it gets is cut to it (its list
 * takes what is left and scrolls). Sizes are those of the parent's own layout, so a
 * CSS transform on an ancestor changes nothing.
 */
function placePanel(element: HTMLElement, parent: HTMLElement, anchor: PanelRect, width: number): void {
  const style = element.style;
  const parentWidth = parent.clientWidth;
  const parentHeight = parent.clientHeight;
  const panelWidth = Math.max(Math.min(width, parentWidth - 2 * GAP), 0);
  style.width = `${panelWidth}px`;
  const below = parentHeight - (anchor.y + anchor.height) - 2 * GAP;
  const above = anchor.y - 2 * GAP;
  const wanted = element.offsetHeight;
  const up = wanted > below && above > below;
  style.maxHeight = `${Math.max(Math.floor(up ? above : below), 0)}px`;
  const height = element.offsetHeight;
  const left = Math.min(Math.max(anchor.x, GAP), Math.max(parentWidth - panelWidth - GAP, GAP));
  const top = up ? Math.max(anchor.y - GAP - height, GAP) : anchor.y + anchor.height + GAP;
  style.left = `${Math.round(left)}px`;
  style.top = `${Math.round(top)}px`;
}

interface VirtualListOptions {
  label: string;
  rowHeight: number;
  /** Several rows can be chosen (check marks) or one. */
  multiple: boolean;
  /** Fills the element of a row; called again whenever the row comes into sight or changes. */
  render(index: number, row: HTMLDivElement): void;
  /** A row was pressed, or Space was pressed on it. */
  onActivate(index: number): void;
  /** Enter was pressed in the list. */
  onEnter(index: number): void;
}

/** A list of any length of which only the rows in sight exist as elements. */
class VirtualList {
  readonly element: HTMLDivElement;
  private readonly sizer: HTMLDivElement;
  private readonly options: VirtualListOptions;
  private readonly id: string;
  private count = 0;
  private active = 0;
  private drawn = new Map<number, HTMLDivElement>();

  constructor(options: VirtualListOptions) {
    this.options = options;
    this.id = `bym-list-${++nextId}`;
    const element = document.createElement('div');
    element.className = 'bym-list';
    element.setAttribute('role', 'listbox');
    element.setAttribute('aria-label', options.label);
    if (options.multiple) element.setAttribute('aria-multiselectable', 'true');
    element.tabIndex = 0;
    element.style.cssText = 'position:relative;box-sizing:border-box;overflow-x:hidden;overflow-y:auto;flex:1 1 auto;overscroll-behavior:contain';
    const sizer = document.createElement('div');
    sizer.style.cssText = 'position:relative;width:100%;height:0';
    element.appendChild(sizer);
    this.element = element;
    this.sizer = sizer;
    element.addEventListener('scroll', this.onScroll, { passive: true });
    element.addEventListener('click', this.onClick);
    element.addEventListener('keydown', this.onKeyDown);
  }

  get activeIndex(): number {
    return this.active;
  }

  /** Sets the number of rows and draws the ones in sight again. */
  setCount(count: number, active = 0): void {
    this.count = Math.max(count, 0);
    this.active = Math.min(Math.max(active, 0), Math.max(this.count - 1, 0));
    this.sizer.style.height = `${this.count * this.options.rowHeight}px`;
    // No taller than its rows, and no taller than a screenful of them; in a small view at least two rows stay.
    this.element.style.maxHeight = `${Math.min(this.count, MAX_VISIBLE_ROWS) * this.options.rowHeight + 2}px`;
    this.element.style.minHeight = `${Math.min(this.count, 2) * this.options.rowHeight}px`;
    this.refresh();
  }

  /** Draws the rows in sight again, after their content or the scroll position changed. */
  refresh(): void {
    const rowHeight = this.options.rowHeight;
    const element = this.element;
    // Before the list has a size (not in the document yet) a screenful is drawn.
    const height = element.clientHeight > 0 ? element.clientHeight : MAX_VISIBLE_ROWS * rowHeight;
    const first = Math.max(Math.floor(element.scrollTop / rowHeight) - 1, 0);
    const last = Math.min(Math.ceil((element.scrollTop + height) / rowHeight) + 1, this.count) - 1;
    for (const [index, row] of this.drawn) {
      if (index < first || index > last) {
        this.sizer.removeChild(row);
        this.drawn.delete(index);
      }
    }
    for (let index = first; index <= last; index++) {
      let row = this.drawn.get(index);
      if (!row) {
        row = document.createElement('div');
        row.className = 'bym-list-row';
        row.setAttribute('role', 'option');
        row.id = `${this.id}-${index}`;
        row.setAttribute('data-index', String(index));
        row.style.cssText = `position:absolute;left:0;right:0;top:${index * rowHeight}px;height:${rowHeight}px;box-sizing:border-box`;
        this.sizer.appendChild(row);
        this.drawn.set(index, row);
      }
      this.options.render(index, row);
      if (index === this.active) row.setAttribute('data-active', '');
      else row.removeAttribute('data-active');
    }
    if (this.count > 0) element.setAttribute('aria-activedescendant', `${this.id}-${this.active}`);
    else element.removeAttribute('aria-activedescendant');
  }

  focus(): void {
    this.element.focus({ preventScroll: true });
  }

  destroy(): void {
    this.element.removeEventListener('scroll', this.onScroll);
    this.element.removeEventListener('click', this.onClick);
    this.element.removeEventListener('keydown', this.onKeyDown);
  }

  private setActive(index: number): void {
    if (this.count === 0) return;
    this.active = Math.min(Math.max(index, 0), this.count - 1);
    const rowHeight = this.options.rowHeight;
    const element = this.element;
    const top = this.active * rowHeight;
    if (top < element.scrollTop) element.scrollTop = top;
    else if (top + rowHeight > element.scrollTop + element.clientHeight) element.scrollTop = top + rowHeight - element.clientHeight;
    this.refresh();
  }

  private onScroll = (): void => {
    this.refresh();
  };

  private onClick = (event: Event): void => {
    let node = event.target as HTMLElement | null;
    while (node && node !== this.element && !node.hasAttribute('data-index')) node = node.parentElement;
    if (!node || node === this.element) return;
    const index = Number(node.getAttribute('data-index'));
    if (!(index >= 0 && index < this.count)) return;
    this.active = index;
    this.options.onActivate(index);
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    const page = Math.max(Math.floor(this.element.clientHeight / this.options.rowHeight) - 1, 1);
    switch (event.key) {
      case 'ArrowDown':
        this.setActive(this.active + 1);
        break;
      case 'ArrowUp':
        this.setActive(this.active - 1);
        break;
      case 'PageDown':
        this.setActive(this.active + page);
        break;
      case 'PageUp':
        this.setActive(this.active - page);
        break;
      case 'Home':
        this.setActive(0);
        break;
      case 'End':
        this.setActive(this.count - 1);
        break;
      case ' ':
      case 'Spacebar':
        if (this.count > 0) this.options.onActivate(this.active);
        break;
      case 'Enter':
        this.options.onEnter(this.active);
        break;
      default:
        return;
    }
    event.preventDefault();
    event.stopPropagation();
  };
}

/** What every panel does: it lives in the view, closes on Escape and on a press outside it, and keeps Tab inside. */
abstract class Panel {
  readonly element: HTMLDivElement;
  private stopWatching: (() => void) | undefined;
  private closed = false;
  /**
   * False while the touch that opened the panel may still be on the screen. Lifting that finger
   * makes the browser send a click at the same spot, which may now be on the panel; it belongs
   * to the opening gesture and must not press anything.
   */
  protected armed: boolean;

  protected constructor(
    className: string,
    label: string,
    openedByTouch: boolean,
    private readonly closing: () => void,
  ) {
    this.armed = !openedByTouch;
    const element = document.createElement('div');
    element.className = `bym-panel ${className}`;
    element.setAttribute('role', 'dialog');
    element.setAttribute('aria-label', label);
    element.style.cssText = 'position:absolute;z-index:4;left:0;top:0;display:flex;flex-direction:column;box-sizing:border-box;overflow-x:hidden';
    this.element = element;
    // Seen before the controls see it: a key ends the opening gesture as a new press does.
    element.addEventListener('keydown', this.onArm, true);
    element.addEventListener('keydown', this.onPanelKeyDown);
    // Stops the browser's own menu from opening over this one.
    element.addEventListener('contextmenu', this.onContextMenu);
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  /** Whether the keyboard focus is inside the panel. */
  get hasFocus(): boolean {
    const root = this.element.getRootNode() as Document | ShadowRoot;
    return this.element.contains(root.activeElement);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    if (this.stopWatching) this.stopWatching();
    this.element.removeEventListener('keydown', this.onArm, true);
    this.element.removeEventListener('keydown', this.onPanelKeyDown);
    this.element.removeEventListener('contextmenu', this.onContextMenu);
    this.dispose();
    // The owner moves the focus while the panel is still in the document: an element that is
    // removed with the focus on it leaves the page without one.
    this.closing();
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
  }

  protected attach(parent: HTMLElement, anchor: PanelRect, width: number): void {
    parent.appendChild(this.element);
    placePanel(this.element, parent, anchor, width);
    this.stopWatching = watchOutsidePresses(
      this.element,
      () => {
        this.armed = true;
      },
      () => this.close(),
    );
  }

  protected abstract dispose(): void;

  private onContextMenu = (event: Event): void => {
    event.preventDefault();
  };

  private onArm = (): void => {
    this.armed = true;
  };

  private onPanelKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      event.preventDefault();
      // The host page must not also act on it (close a dialog of its own, say).
      event.stopPropagation();
      this.close();
      return;
    }
    if (event.key !== 'Tab') return;
    // Tab stays inside: from the last control it goes to the first and back.
    const stops = Array.prototype.slice.call(this.element.querySelectorAll('button:not(:disabled), input, [tabindex="0"]')) as HTMLElement[];
    if (stops.length === 0) return;
    const root = this.element.getRootNode() as Document | ShadowRoot;
    const at = stops.indexOf(root.activeElement as HTMLElement);
    const next = event.shiftKey ? (at <= 0 ? stops.length - 1 : at - 1) : at < 0 || at === stops.length - 1 ? 0 : at + 1;
    event.preventDefault();
    event.stopPropagation();
    stops[next].focus({ preventScroll: true });
  };
}

export interface FilterMenuStrings {
  sortAscending: string;
  sortDescending: string;
  clearFilter: string;
  searchValues: string;
  selectAll: string;
  blanks: string;
  apply: string;
  cancel: string;
  noMatches: string;
  valueList: string;
}

export interface FilterMenuOptions {
  /** Positioned element the menu is placed in; it is kept inside it. */
  parent: HTMLElement;
  /** The header cell the menu belongs to, in the parent's own pixels. */
  anchor: PanelRect;
  label: string;
  strings: FilterMenuStrings;
  /** The values of the column, as its cells show them, with whether they are shown now. */
  values: { text: string; checked: boolean }[];
  /** The column has empty cells, and whether they are shown now. */
  blanks: boolean;
  blanksChecked: boolean;
  /** Shown under the list when it does not hold every value of the column. */
  note?: string;
  /** The column has a criterion that can be removed. */
  canClear: boolean;
  openedByTouch?: boolean;
  onSort(descending: boolean): void;
  /** The values to show and whether empty cells are shown; undefined removes the criterion. */
  onApply(choice: { values: string[]; blank: boolean } | undefined): void;
  /** Runs once, whichever way the menu closes, before the action that closed it. */
  onClose(): void;
}

// Rows of the value list that are not values.
const ROW_ALL = -1;
const ROW_BLANKS = -2;

function fold(text: string): string {
  return text.replace(/[İIı]/g, 'i').toLowerCase();
}

function button(className: string, label: string): HTMLButtonElement {
  const element = document.createElement('button');
  element.type = 'button';
  element.className = className;
  element.textContent = label;
  return element;
}

/** The menu of a filter button: sort, clear, and the values of the column to choose from. */
export class FilterMenu extends Panel {
  private readonly options: FilterMenuOptions;
  private readonly list: VirtualList;
  private readonly search: HTMLInputElement;
  private readonly applyButton: HTMLButtonElement;
  private readonly empty: HTMLDivElement;
  private readonly checked: boolean[];
  private blanksChecked: boolean;
  /** Folded texts of the values, for the search. */
  private readonly folded: string[];
  /** What the list shows: indexes of values, ROW_ALL first and ROW_BLANKS last. */
  private rows: number[] = [];

  constructor(options: FilterMenuOptions) {
    super('bym-filter', options.label, options.openedByTouch === true, options.onClose);
    this.options = options;
    const strings = options.strings;
    const coarse = coarsePointer();
    this.checked = options.values.map((value) => value.checked);
    this.blanksChecked = options.blanksChecked;
    this.folded = options.values.map((value) => fold(value.text));

    const element = this.element;
    const ascending = button('bym-panel-item', strings.sortAscending);
    ascending.setAttribute('data-action', 'sort-asc');
    const descending = button('bym-panel-item', strings.sortDescending);
    descending.setAttribute('data-action', 'sort-desc');
    const clear = button('bym-panel-item', strings.clearFilter);
    clear.setAttribute('data-action', 'clear');
    clear.disabled = !options.canClear;
    const separator = document.createElement('div');
    separator.className = 'bym-panel-separator';

    const search = document.createElement('input');
    search.type = 'text';
    search.className = 'bym-panel-search';
    search.setAttribute('aria-label', strings.searchValues);
    search.placeholder = strings.searchValues;
    search.setAttribute('autocomplete', 'off');
    search.setAttribute('autocorrect', 'off');
    search.setAttribute('autocapitalize', 'off');
    search.spellcheck = false;
    this.search = search;

    this.list = new VirtualList({
      label: strings.valueList,
      rowHeight: coarse ? 40 : 28,
      multiple: true,
      render: (index, row) => this.renderRow(index, row),
      onActivate: (index) => this.toggle(index),
      onEnter: () => this.apply(),
    });

    const empty = document.createElement('div');
    empty.className = 'bym-panel-note bym-panel-empty';
    empty.setAttribute('role', 'status');
    empty.textContent = strings.noMatches;
    empty.style.display = 'none';
    this.empty = empty;

    const actions = document.createElement('div');
    actions.className = 'bym-panel-actions';
    const applyButton = button('bym-panel-button bym-panel-primary', strings.apply);
    applyButton.setAttribute('data-action', 'apply');
    const cancel = button('bym-panel-button', strings.cancel);
    cancel.setAttribute('data-action', 'cancel');
    actions.appendChild(applyButton);
    actions.appendChild(cancel);
    this.applyButton = applyButton;

    element.appendChild(ascending);
    element.appendChild(descending);
    element.appendChild(separator);
    element.appendChild(clear);
    element.appendChild(search);
    element.appendChild(this.list.element);
    element.appendChild(empty);
    if (options.note) {
      const note = document.createElement('div');
      note.className = 'bym-panel-note';
      note.setAttribute('role', 'note');
      note.textContent = options.note;
      element.appendChild(note);
    }
    element.appendChild(actions);

    element.addEventListener('click', this.onClick);
    search.addEventListener('input', this.onSearch);
    search.addEventListener('keydown', this.onSearchKey);

    this.filterRows();
    this.attach(options.parent, options.anchor, PANEL_WIDTH);
    // Now that the list has its size, the rows in sight are the right ones.
    this.list.refresh();
    this.list.focus();
  }

  protected dispose(): void {
    this.element.removeEventListener('click', this.onClick);
    this.search.removeEventListener('input', this.onSearch);
    this.search.removeEventListener('keydown', this.onSearchKey);
    this.list.destroy();
  }

  /** Brings the rows of the list in line with the search text. */
  private filterRows(): void {
    const query = fold(this.search.value.trim());
    const rows: number[] = [];
    for (let i = 0; i < this.folded.length; i++) if (query === '' || this.folded[i].indexOf(query) >= 0) rows.push(i);
    if (this.options.blanks && query === '') rows.push(ROW_BLANKS);
    if (rows.length > 0) rows.unshift(ROW_ALL);
    this.rows = rows;
    this.empty.style.display = rows.length === 0 ? 'block' : 'none';
    this.list.element.style.display = rows.length === 0 ? 'none' : 'block';
    this.list.setCount(rows.length);
    this.updateApply();
  }

  private isChecked(row: number): boolean {
    return row === ROW_BLANKS ? this.blanksChecked : this.checked[row];
  }

  /** Of the rows in the list (other than "select all"): 1 all checked, 0 none, -1 some. */
  private allState(): number {
    let on = 0;
    let off = 0;
    for (let i = 1; i < this.rows.length; i++) {
      if (this.isChecked(this.rows[i])) on++;
      else off++;
      if (on > 0 && off > 0) return -1;
    }
    return off === 0 && on > 0 ? 1 : 0;
  }

  private renderRow(index: number, element: HTMLDivElement): void {
    const row = this.rows[index];
    const strings = this.options.strings;
    const state = row === ROW_ALL ? this.allState() : this.isChecked(row) ? 1 : 0;
    const text = row === ROW_ALL ? strings.selectAll : row === ROW_BLANKS ? strings.blanks : this.options.values[row].text;
    let box = element.firstChild as HTMLSpanElement | null;
    let label = element.lastChild as HTMLSpanElement | null;
    if (!box || !label || box === label) {
      box = document.createElement('span');
      box.className = 'bym-check';
      box.setAttribute('aria-hidden', 'true');
      label = document.createElement('span');
      label.className = 'bym-list-text';
      element.appendChild(box);
      element.appendChild(label);
    }
    box.setAttribute('data-state', state === 1 ? 'on' : state === 0 ? 'off' : 'mixed');
    if (label.textContent !== text) label.textContent = text;
    element.setAttribute('aria-selected', state === 1 ? 'true' : 'false');
    element.setAttribute('data-kind', row === ROW_ALL ? 'all' : row === ROW_BLANKS ? 'blanks' : 'value');
  }

  private toggle(index: number): void {
    const row = this.rows[index];
    if (row === undefined || !this.armed) return;
    if (row === ROW_ALL) {
      // "Select all" speaks for the rows in the list: with a search, for what was found.
      const on = this.allState() !== 1;
      for (let i = 1; i < this.rows.length; i++) {
        if (this.rows[i] === ROW_BLANKS) this.blanksChecked = on;
        else this.checked[this.rows[i]] = on;
      }
    } else if (row === ROW_BLANKS) this.blanksChecked = !this.blanksChecked;
    else this.checked[row] = !this.checked[row];
    this.list.refresh();
    this.updateApply();
  }

  /** What Apply would set: undefined for "show everything", null when nothing is chosen. */
  private choice(): { values: string[]; blank: boolean } | undefined | null {
    const searching = this.search.value.trim() !== '';
    const values: string[] = [];
    let all = true;
    if (searching) {
      // With a search, the checked rows that were found are what is shown, as in Excel.
      for (let i = 1; i < this.rows.length; i++) if (this.rows[i] >= 0 && this.checked[this.rows[i]]) values.push(this.options.values[this.rows[i]].text);
      return values.length > 0 ? { values, blank: false } : null;
    }
    for (let i = 0; i < this.checked.length; i++) {
      if (this.checked[i]) values.push(this.options.values[i].text);
      else all = false;
    }
    const blank = this.options.blanks && this.blanksChecked;
    if (this.options.blanks && !blank) all = false;
    if (all) return undefined;
    return values.length > 0 || blank ? { values, blank } : null;
  }

  private updateApply(): void {
    this.applyButton.disabled = this.choice() === null;
  }

  private apply(): void {
    const choice = this.choice();
    if (choice === null) return;
    this.close();
    this.options.onApply(choice);
  }

  private onSearch = (): void => {
    this.filterRows();
  };

  private onSearchKey = (event: KeyboardEvent): void => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      event.stopPropagation();
      this.apply();
    } else if (event.key === 'ArrowDown' && this.rows.length > 0) {
      event.preventDefault();
      this.list.focus();
    }
  };

  private onClick = (event: Event): void => {
    if (!this.armed) return;
    const target = event.target as Element | null;
    const pressed = target && typeof target.closest === 'function' ? (target.closest('button') as HTMLButtonElement | null) : null;
    if (!pressed || pressed.disabled) return;
    switch (pressed.getAttribute('data-action')) {
      case 'sort-asc':
      case 'sort-desc': {
        const descending = pressed.getAttribute('data-action') === 'sort-desc';
        this.close();
        this.options.onSort(descending);
        break;
      }
      case 'clear':
        this.close();
        this.options.onApply(undefined);
        break;
      case 'apply':
        this.apply();
        break;
      case 'cancel':
        this.close();
        break;
      default:
    }
  };
}

export interface ListMenuOptions {
  parent: HTMLElement;
  /** The cell the list belongs to, in the parent's own pixels. */
  anchor: PanelRect;
  label: string;
  items: string[];
  /** Index of the item the cell holds now, if it holds one of them. */
  selected?: number;
  note?: string;
  openedByTouch?: boolean;
  onChoose(index: number): void;
  onClose(): void;
}

/** The list of a data validation: one of its values is chosen for the cell. */
export class ListMenu extends Panel {
  private readonly options: ListMenuOptions;
  private readonly list: VirtualList;

  constructor(options: ListMenuOptions) {
    super('bym-dropdown', options.label, options.openedByTouch === true, options.onClose);
    this.options = options;
    const selected = options.selected !== undefined && options.selected >= 0 ? options.selected : -1;
    this.list = new VirtualList({
      label: options.label,
      rowHeight: coarsePointer() ? 40 : 28,
      multiple: false,
      render: (index, row) => {
        const text = options.items[index];
        if (row.textContent !== text) row.textContent = text;
        row.setAttribute('aria-selected', index === selected ? 'true' : 'false');
      },
      onActivate: (index) => this.choose(index),
      onEnter: (index) => this.choose(index),
    });
    this.element.appendChild(this.list.element);
    if (options.note) {
      const note = document.createElement('div');
      note.className = 'bym-panel-note';
      note.setAttribute('role', 'note');
      note.textContent = options.note;
      this.element.appendChild(note);
    }
    this.list.setCount(options.items.length, Math.max(selected, 0));
    // As wide as its cell, and wide enough to read.
    this.attach(options.parent, options.anchor, Math.max(options.anchor.width, 140));
    this.list.refresh();
    this.list.focus();
  }

  protected dispose(): void {
    this.list.destroy();
  }

  private choose(index: number): void {
    if (!this.armed || !(index >= 0 && index < this.options.items.length)) return;
    this.close();
    this.options.onChoose(index);
  }
}
