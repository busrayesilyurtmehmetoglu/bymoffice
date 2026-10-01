// What drops down from a toolbar button: a list of commands (borders, rows
// and columns) or a colour palette. It closes on Escape, on Tab, on a press
// outside it and when a command is chosen. Labels only ever reach the DOM
// through textContent.

export interface PopoverItem {
  label: string;
  /** Shown before the label of a list item. */
  icon?: SVGElement;
  /** A CSS colour: the item is a palette cell that shows it, named by its label. */
  swatch?: string;
  /** Written to the cell as data-color, for the page and its tests to find the cell by. */
  value?: string;
  /** The item is what the selection has now. */
  selected?: boolean;
  action(): void;
}

export type PopoverCloseReason = 'select' | 'escape' | 'tab' | 'outside' | 'owner';

export interface PopoverOptions {
  /** Positioned element the popover is placed in; it is kept inside its width. */
  parent: HTMLElement;
  /** The button it belongs to; the popover opens under it. */
  anchor: HTMLElement;
  label: string;
  /** List items, one per row. With `swatches` they come first, above the palette. */
  items: PopoverItem[];
  /** Palette cells, laid out in rows of `columns`. */
  swatches?: PopoverItem[];
  columns?: number;
  /** Runs once, whichever way the popover closes, before the chosen item's action. */
  onClose(reason: PopoverCloseReason): void;
}

/** Space kept between the popover and the anchor, and the edges of the parent. */
const GAP = 4;

export class Popover {
  readonly element: HTMLDivElement;
  readonly anchor: HTMLElement;
  private readonly actions: (() => void)[] = [];
  private readonly buttons: HTMLButtonElement[] = [];
  /** Buttons before the palette, and cells per palette row (0: a plain list). */
  private readonly listCount: number;
  private readonly columns: number;
  private readonly onClose: (reason: PopoverCloseReason) => void;
  private closed = false;
  /**
   * Where presses are listened for: the document, and every shadow root between
   * it and the popover. A listener on the document alone only ever sees the
   * shadow host as the target, so it cannot tell a press on an item from one
   * beside the popover.
   */
  private readonly watched: { root: Node; listener: (event: Event) => void }[] = [];

  constructor(options: PopoverOptions) {
    this.anchor = options.anchor;
    this.onClose = options.onClose;
    const swatches = options.swatches || [];
    this.listCount = options.items.length;
    this.columns = swatches.length > 0 ? Math.max(1, Math.floor(options.columns || 1)) : 0;

    const element = document.createElement('div');
    element.className = 'bym-popover';
    element.setAttribute('role', swatches.length > 0 ? 'dialog' : 'menu');
    element.setAttribute('aria-label', options.label);
    element.style.cssText = 'position:absolute;z-index:10;left:0;top:0;box-sizing:border-box;overflow:auto';

    for (let i = 0; i < options.items.length; i++) {
      const item = options.items[i];
      const button = this.button('bym-pop-item', item);
      if (swatches.length === 0) button.setAttribute('role', 'menuitem');
      if (item.icon) button.appendChild(item.icon);
      const text = document.createElement('span');
      text.textContent = item.label;
      button.appendChild(text);
      element.appendChild(button);
    }
    if (swatches.length > 0) {
      const grid = document.createElement('div');
      grid.className = 'bym-pop-grid';
      grid.style.gridTemplateColumns = `repeat(${this.columns}, auto)`;
      for (let i = 0; i < swatches.length; i++) {
        const item = swatches[i];
        const button = this.button('bym-pop-swatch', item);
        button.setAttribute('aria-label', item.label);
        button.title = item.label;
        if (item.value !== undefined) button.setAttribute('data-color', item.value);
        button.style.backgroundColor = item.swatch || 'transparent';
        grid.appendChild(button);
      }
      element.appendChild(grid);
    }

    this.element = element;
    options.parent.appendChild(element);
    this.place(options.parent);

    element.addEventListener('click', this.onClick);
    element.addEventListener('keydown', this.onKeyDown);
    this.watchPresses();

    // The keys work from inside: the item the selection has now takes the focus, else the first.
    let first = 0;
    for (let i = 0; i < this.buttons.length; i++) {
      if (this.buttons[i].getAttribute('aria-pressed') === 'true') {
        first = i;
        break;
      }
    }
    if (this.buttons[first]) this.buttons[first].focus({ preventScroll: true });
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  /** Whether the keyboard focus is on one of the items. */
  get hasFocus(): boolean {
    const root = this.element.getRootNode() as Document | ShadowRoot;
    return this.element.contains(root.activeElement);
  }

  close(reason: PopoverCloseReason = 'owner'): void {
    if (this.closed) return;
    this.closed = true;
    this.element.removeEventListener('click', this.onClick);
    this.element.removeEventListener('keydown', this.onKeyDown);
    for (let i = 0; i < this.watched.length; i++) {
      const { root, listener } = this.watched[i];
      root.removeEventListener('pointerdown', listener, true);
      root.removeEventListener('mousedown', listener, true);
      root.removeEventListener('touchstart', listener, true);
    }
    this.watched.length = 0;
    // The owner moves the focus while the items are still in the document: an element that is
    // removed with the focus on it leaves the page without one.
    this.onClose(reason);
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
  }

  private button(className: string, item: PopoverItem): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = className;
    // Reached with the arrow keys, not with Tab.
    button.tabIndex = -1;
    if (item.selected !== undefined) button.setAttribute('aria-pressed', item.selected ? 'true' : 'false');
    this.buttons.push(button);
    this.actions.push(item.action);
    return button;
  }

  /**
   * Under the anchor, or above it when there is more room there, and moved
   * left as far as needed: the popover stays inside the parent, and scrolls
   * within itself when the parent is smaller than it is.
   */
  private place(parent: HTMLElement): void {
    const style = this.element.style;
    const box = parent.getBoundingClientRect();
    const anchor = this.anchor.getBoundingClientRect();
    // Rectangles are in screen pixels, positions in the parent's own: they differ under a CSS transform.
    const scaleX = parent.offsetWidth > 0 && box.width > 0 ? box.width / parent.offsetWidth : 1;
    const scaleY = parent.offsetHeight > 0 && box.height > 0 ? box.height / parent.offsetHeight : 1;
    const anchorLeft = (anchor.left - box.left) / scaleX;
    const anchorTop = (anchor.top - box.top) / scaleY;
    const anchorBottom = (anchor.bottom - box.top) / scaleY;
    const parentWidth = parent.clientWidth;
    const parentHeight = parent.clientHeight;

    style.maxWidth = `${Math.max(parentWidth - 2 * GAP, 0)}px`;
    const below = parentHeight - anchorBottom - 2 * GAP;
    const above = anchorTop - 2 * GAP;
    const wanted = this.element.offsetHeight;
    const up = wanted > below && above > below;
    style.maxHeight = `${Math.max(Math.floor(up ? above : below), 0)}px`;
    const width = this.element.offsetWidth;
    const height = this.element.offsetHeight;
    const left = Math.min(Math.max(anchorLeft, GAP), Math.max(parentWidth - width - GAP, GAP));
    const top = up ? anchorTop - GAP - height : anchorBottom + GAP;
    style.left = `${Math.round(left)}px`;
    style.top = `${Math.round(top)}px`;
  }

  private watchPresses(): void {
    // Pointer events where there are any; the mouse events a browser makes up after a touch would count twice.
    const types = 'PointerEvent' in window ? ['pointerdown'] : ['mousedown', 'touchstart'];
    let root: Node = this.element.getRootNode();
    let inner: Element | undefined;
    // Ends at the document; the bound only guards against a tree that is not one.
    for (let depth = 0; depth < 64; depth++) {
      const through = inner;
      const listener = (event: Event) => this.onPress(event, through);
      for (let i = 0; i < types.length; i++) root.addEventListener(types[i], listener, true);
      this.watched.push({ root, listener });
      const host = (root as ShadowRoot).host;
      if (root.nodeType !== 11 || !host) break;
      inner = host;
      root = host.getRootNode();
    }
  }

  /**
   * A press seen from one of the watched roots. `through` is the shadow host
   * that leads towards the popover from there: a press on it is judged by the
   * listener inside that host, which sees the real target.
   */
  private onPress(event: Event, through: Element | undefined): void {
    const target = event.target as Node | null;
    if (!target) return;
    if (through) {
      if (target !== through) this.close('outside');
      return;
    }
    if (this.element.contains(target)) return;
    // A press on the anchor is the owner's to answer: it closes the popover it opened.
    if (this.anchor.contains(target)) return;
    this.close('outside');
  }

  private onClick = (event: Event): void => {
    const target = event.target as Node;
    for (let i = 0; i < this.buttons.length; i++) {
      if (this.buttons[i].contains(target)) {
        this.close('select');
        this.actions[i]();
        return;
      }
    }
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    const buttons = this.buttons;
    const count = buttons.length;
    const root = this.element.getRootNode() as Document | ShadowRoot;
    const index = buttons.indexOf(root.activeElement as HTMLButtonElement);
    const list = this.listCount;
    const columns = this.columns;
    let next = -1;
    switch (event.key) {
      case 'ArrowDown':
        if (index < 0) next = 0;
        else if (columns === 0) next = (index + 1) % count;
        else if (index < list) next = index + 1;
        else next = index + columns < count ? index + columns : index;
        break;
      case 'ArrowUp':
        if (index < 0) next = count - 1;
        else if (columns === 0) next = (index + count - 1) % count;
        else if (index < list) next = Math.max(index - 1, 0);
        else next = index - columns >= list ? index - columns : list > 0 ? list - 1 : index;
        break;
      case 'ArrowRight':
        if (columns === 0) return;
        next = Math.min(Math.max(index, 0) + 1, count - 1);
        break;
      case 'ArrowLeft':
        if (columns === 0) return;
        next = Math.max(index - 1, 0);
        break;
      case 'Home':
        next = 0;
        break;
      case 'End':
        next = count - 1;
        break;
      case 'Escape':
        this.close('escape');
        break;
      case 'Tab':
        this.close('tab');
        break;
      default:
        // Enter and Space press the focused button without help.
        return;
    }
    event.preventDefault();
    event.stopPropagation();
    // Scrolls a popover that is cut short by a small editor to the item.
    if (next >= 0 && next < count) buttons[next].focus();
  };
}
