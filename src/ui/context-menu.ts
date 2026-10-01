// A small pop-up menu: a list of buttons placed at a point of its parent.
// It closes on Escape, on Tab, on a press outside it and when asked to; the
// owner decides what else closes it (scrolling, say). Labels only ever reach
// the DOM through textContent.

export interface MenuItem {
  /** Text of the item; unused for a separator. */
  label: string;
  action?(): void;
  /** A line between groups of items instead of an item. */
  separator?: boolean;
  /** Shown but not available: it cannot be pressed and the arrow keys pass over it. */
  disabled?: boolean;
}

export class ContextMenu {
  readonly element: HTMLDivElement;
  private readonly items: MenuItem[];
  /** The element of each item, in the order of `items`. */
  private readonly nodes: HTMLElement[] = [];
  private readonly onClose: () => void;
  /** The document or shadow root the menu lives in: where presses and the focus are seen as they are. */
  private readonly rootNode: Document | ShadowRoot;
  private closed = false;
  /**
   * False while the touch that opened the menu is still on the screen. Lifting
   * that finger makes the browser send mouse events and a click at the same
   * spot, which is now inside or beside the menu; they belong to the opening
   * gesture and must neither press an item nor count as a press outside.
   */
  private armed: boolean;

  /**
   * Shows the menu inside `parent` (which must be positioned) with its corner
   * at (x, y), moved as far as needed to stay inside the parent. The first
   * available item takes the focus. `onClose` runs once, whichever way the
   * menu closes.
   */
  constructor(parent: HTMLElement, items: MenuItem[], x: number, y: number, onClose: () => void, openedByTouch = false) {
    this.onClose = onClose;
    this.items = items;
    this.armed = !openedByTouch;
    const element = document.createElement('div');
    element.className = 'bym-menu';
    element.setAttribute('role', 'menu');
    // Never taller than its parent: a menu that does not fit scrolls.
    element.style.cssText =
      'position:absolute;z-index:3;display:flex;flex-direction:column;left:0;top:0;box-sizing:border-box;overflow-x:hidden;overflow-y:auto;' +
      `max-height:${Math.max(parent.clientHeight - 8, 48)}px`;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      let node: HTMLElement;
      if (item.separator) {
        node = document.createElement('div');
        node.className = 'bym-menu-separator';
        node.setAttribute('role', 'separator');
      } else {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'bym-menu-item';
        button.setAttribute('role', 'menuitem');
        // Reached with the arrow keys, not with Tab.
        button.tabIndex = -1;
        button.textContent = item.label;
        if (item.disabled) {
          button.disabled = true;
          button.setAttribute('aria-disabled', 'true');
        }
        node = button;
      }
      // Items keep their height in a menu that scrolls.
      node.style.flex = '0 0 auto';
      this.nodes.push(node);
      element.appendChild(node);
    }
    this.element = element;
    parent.appendChild(element);
    const left = Math.max(Math.min(x, parent.clientWidth - element.offsetWidth), 0);
    const top = Math.max(Math.min(y, parent.clientHeight - element.offsetHeight), 0);
    element.style.left = `${left}px`;
    element.style.top = `${top}px`;

    element.addEventListener('click', this.onClick);
    element.addEventListener('keydown', this.onKeyDown);
    // Stops the browser's own menu from opening over this one.
    element.addEventListener('contextmenu', this.onContextMenu);
    // Pointer events where there are any. The mouse events a browser makes up after a touch would
    // otherwise count a second time, and a handler that cancels pointerdown cancels them anyway.
    //
    // Inside a shadow root the document only sees the host as the target of a press. The root sees the real
    // target, so it decides for presses inside it, and the document only for presses elsewhere on the page.
    this.rootNode = element.getRootNode() as Document | ShadowRoot;
    const types = 'PointerEvent' in window ? ['pointerdown'] : ['mousedown', 'touchstart'];
    for (let i = 0; i < types.length; i++) {
      this.rootNode.addEventListener(types[i], this.onPress, true);
      if (this.rootNode !== document) document.addEventListener(types[i], this.onPagePress, true);
    }
    this.focusItem(this.step(-1, 1));
  }

  /** Focuses an item and scrolls the menu, and nothing else, to show it. */
  private focusItem(index: number): void {
    const node = this.nodes[index];
    if (!node) return;
    node.focus({ preventScroll: true });
    const menu = this.element;
    const top = node.offsetTop;
    const bottom = top + node.offsetHeight;
    if (top < menu.scrollTop) menu.scrollTop = top;
    else if (bottom > menu.scrollTop + menu.clientHeight) menu.scrollTop = bottom - menu.clientHeight;
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.element.removeEventListener('click', this.onClick);
    this.element.removeEventListener('keydown', this.onKeyDown);
    this.element.removeEventListener('contextmenu', this.onContextMenu);
    const types = ['pointerdown', 'mousedown', 'touchstart'];
    for (let i = 0; i < types.length; i++) {
      this.rootNode.removeEventListener(types[i], this.onPress, true);
      document.removeEventListener(types[i], this.onPagePress, true);
    }
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    this.onClose();
  }

  private usable(index: number): boolean {
    const item = this.items[index];
    return item !== undefined && !item.separator && !item.disabled;
  }

  /** The next available item from `index` in a direction, going round; -1 when there is none. */
  private step(index: number, direction: number): number {
    const count = this.items.length;
    for (let n = 1; n <= count; n++) {
      const next = (((index + direction * n) % count) + count) % count;
      if (this.usable(next)) return next;
    }
    return -1;
  }

  private onClick = (event: Event): void => {
    if (!this.armed) return;
    for (let i = 0; i < this.nodes.length; i++) {
      if (this.nodes[i].contains(event.target as Node)) {
        if (!this.usable(i)) return;
        const action = this.items[i].action;
        this.close();
        if (action) action();
        return;
      }
    }
  };

  private onContextMenu = (event: Event): void => {
    event.preventDefault();
  };

  /** A press somewhere in the menu's own document or shadow root. */
  private onPress = (event: Event): void => {
    // Any new press ends the opening gesture.
    this.armed = true;
    if (!this.element.contains(event.target as Node)) this.close();
  };

  /** A press as the page sees it, when the menu is in a shadow root: only presses outside that root count here. */
  private onPagePress = (event: Event): void => {
    const host = (this.rootNode as ShadowRoot).host;
    const target = event.target as Node | null;
    if (host && target && (target === host || host.contains(target))) return;
    this.armed = true;
    this.close();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    let index = -1;
    const focused = this.rootNode.activeElement;
    for (let i = 0; i < this.nodes.length; i++) if (this.nodes[i] === focused) index = i;
    let next = -1;
    switch (event.key) {
      case 'ArrowDown':
        next = this.step(index, 1);
        break;
      case 'ArrowUp':
        next = this.step(index < 0 ? 0 : index, -1);
        break;
      case 'Home':
        next = this.step(-1, 1);
        break;
      case 'End':
        next = this.step(0, -1);
        break;
      case 'Escape':
      case 'Tab':
        this.armed = true;
        this.close();
        break;
      default:
        // Enter and Space press the focused button without help.
        this.armed = true;
        return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (next >= 0) this.focusItem(next);
  };
}
