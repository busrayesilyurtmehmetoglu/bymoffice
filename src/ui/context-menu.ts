// A small pop-up menu: a list of buttons placed at a point of its parent.
// It closes on Escape, on Tab, on a press outside it and when asked to; the
// owner decides what else closes it (scrolling, say). Labels only ever reach
// the DOM through textContent.

export interface MenuItem {
  label: string;
  action(): void;
}

export class ContextMenu {
  readonly element: HTMLDivElement;
  private readonly items: MenuItem[];
  private readonly onClose: () => void;
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
   * item takes the focus. `onClose` runs once, whichever way the menu closes.
   */
  constructor(parent: HTMLElement, items: MenuItem[], x: number, y: number, onClose: () => void, openedByTouch = false) {
    this.onClose = onClose;
    this.items = items;
    this.armed = !openedByTouch;
    const element = document.createElement('div');
    element.className = 'bym-menu';
    element.setAttribute('role', 'menu');
    element.style.cssText = 'position:absolute;z-index:3;display:flex;flex-direction:column;left:0;top:0';
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'bym-menu-item';
      button.setAttribute('role', 'menuitem');
      // Reached with the arrow keys, not with Tab.
      button.tabIndex = -1;
      button.textContent = item.label;
      element.appendChild(button);
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
    if ('PointerEvent' in window) document.addEventListener('pointerdown', this.onOutside, true);
    else {
      document.addEventListener('mousedown', this.onOutside, true);
      document.addEventListener('touchstart', this.onOutside, true);
    }
    const first = element.firstChild as HTMLElement | null;
    if (first) first.focus({ preventScroll: true });
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
    document.removeEventListener('pointerdown', this.onOutside, true);
    document.removeEventListener('mousedown', this.onOutside, true);
    document.removeEventListener('touchstart', this.onOutside, true);
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    this.onClose();
  }

  private onClick = (event: Event): void => {
    if (!this.armed) return;
    const buttons = this.element.children;
    for (let i = 0; i < buttons.length; i++) {
      if (buttons[i].contains(event.target as Node)) {
        this.close();
        this.items[i].action();
        return;
      }
    }
  };

  private onContextMenu = (event: Event): void => {
    event.preventDefault();
  };

  private onOutside = (event: Event): void => {
    // Any new press ends the opening gesture.
    this.armed = true;
    if (!this.element.contains(event.target as Node)) this.close();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    const items = this.element.children;
    let index = -1;
    for (let i = 0; i < items.length; i++) if (items[i] === document.activeElement) index = i;
    let next = -1;
    switch (event.key) {
      case 'ArrowDown':
        next = (index + 1) % items.length;
        break;
      case 'ArrowUp':
        next = (index + items.length - 1) % items.length;
        break;
      case 'Home':
        next = 0;
        break;
      case 'End':
        next = items.length - 1;
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
    if (next >= 0 && items[next]) (items[next] as HTMLElement).focus({ preventScroll: true });
  };
}
