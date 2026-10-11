// The focus a press asks for, in a frame that is not the active one.
//
// Firefox does not move the focus to an element of a frame while a press on
// that frame is being handled and the frame is not the active one yet: focus()
// on the hidden input does nothing, the frame becomes active with the release
// and its body has the focus. What the user types next goes nowhere. (Other
// engines move it at once.) The focus is asked for again when the press ends.
//
// Only a press is followed up, and only its own release: a focus that was asked
// for in any other way and did not take (a view that is hidden, a call of the
// application) is not taken later, at a release that has nothing to do with it.

/** The press a focus was asked for in: what tells its release from another one. */
export interface FocusPress {
  /** The pointer of a pointer event; undefined for a mouse or a touch event. */
  pointerId: number | undefined;
  /** A finger, not a mouse or a pen. */
  touch: boolean;
}

/** What a mousedown, pointerdown or touchstart (or the pointerup of a tap) says of its press. */
export function pressOf(event: Event): FocusPress {
  const pointer = event as PointerEvent;
  const touch = (typeof event.type === 'string' && event.type.indexOf('touch') === 0) || pointer.pointerType === 'touch';
  return { pointerId: typeof pointer.pointerId === 'number' ? pointer.pointerId : undefined, touch };
}

/** Whether `event` (pointerup, mouseup, touchend) ends the press. */
function releases(press: FocusPress, event: Event): boolean {
  const pointer = event as PointerEvent;
  if (event.type === 'pointerup') {
    if (press.pointerId !== undefined) return pointer.pointerId === press.pointerId;
    // A press that was seen as a mouse event: the mouse is the one pointer of its kind.
    return !press.touch && (pointer.pointerType === undefined || pointer.pointerType === 'mouse');
  }
  if (event.type === 'mouseup') return !press.touch;
  return press.touch;
}

/** How deep shadow roots are followed up to the document; deeper than that the focus is left alone. */
const MAX_ROOTS = 32;

/**
 * Where the focus is, seen from `input`: on it, nowhere (the body, the root element, a host of
 * the shadow roots the input is in, or nothing), or on something else, which keeps it: an element,
 * or another document when this one is a frame that does not have the focus. Inside a
 * shadow root `activeElement` is null for a focus outside that root, so the roots around it are asked.
 */
function focusOf(input: HTMLElement): 'input' | 'nowhere' | 'elsewhere' {
  const doc = input.ownerDocument;
  let root = input.getRootNode() as Document | ShadowRoot;
  let active: Element | null | undefined = root.activeElement;
  if (active === input) return 'input';
  // The document is a frame that is not the active one: the focus is in the page around it (a field that page
  // focused while the press was held), and to that page this document's body looks like nothing has it.
  if (typeof doc.hasFocus === 'function' && !doc.hasFocus()) return 'elsewhere';
  let host: Element | undefined;
  for (let depth = 0; depth < MAX_ROOTS; depth++) {
    const outer = (root as ShadowRoot).host as Element | undefined;
    if (root === doc || !outer) return !active || active === doc.body || active === doc.documentElement || active === host ? 'nowhere' : 'elsewhere';
    // Something in this tree has it: a menu of the view, a field of the application.
    if (active && active !== host) return 'elsewhere';
    host = outer;
    root = outer.getRootNode() as Document | ShadowRoot;
    active = root.activeElement;
  }
  return 'elsewhere';
}

const RELEASES = ['pointerup', 'mouseup', 'touchend'];
/** What ends a press without a release, or makes its release no longer the user's wish for the focus. */
const CANCELS = ['pointercancel', 'touchcancel', 'dragend', 'visibilitychange'];

/** The one follow-up a view can have waiting for the end of a press. */
export class FocusRetry {
  private readonly input: HTMLElement;
  private readonly alive: () => boolean;
  /** Takes the listeners of the follow-up that waits away. */
  private off: (() => void) | undefined;

  constructor(input: HTMLElement, alive: () => boolean) {
    this.input = input;
    this.alive = alive;
  }

  /** Whether a follow-up waits for the end of a press. */
  get waiting(): boolean {
    return this.off !== undefined;
  }

  /**
   * After `input.focus()` was called while `press` is handled: when the input did not get the
   * focus, `focus` is called once more when that press is released, if the focus is still
   * nowhere then and the view is alive. A follow-up that waited before is given up.
   */
  after(press: FocusPress, focus: () => void): void {
    this.cancel();
    const input = this.input;
    const doc = input.ownerDocument;
    if (!doc || typeof doc.addEventListener !== 'function') return;
    if (focusOf(input) === 'input') return;
    const types = RELEASES.concat(CANCELS);
    const off = (): void => {
      for (let i = 0; i < types.length; i++) doc.removeEventListener(types[i], on, true);
      if (this.off === off) this.off = undefined;
    };
    const on = (event: Event): void => {
      if (RELEASES.indexOf(event.type) < 0) return off();
      // Another finger, another pointer: the press goes on.
      if (!releases(press, event)) return;
      off();
      // Something else took the focus meanwhile (a menu, a field of the application): it keeps it.
      if (this.alive() && focusOf(input) === 'nowhere') focus();
    };
    for (let i = 0; i < types.length; i++) doc.addEventListener(types[i], on, true);
    this.off = off;
  }

  /** Gives up the follow-up that waits, if any: the view is destroyed, or another press began. */
  cancel(): void {
    if (this.off) this.off();
  }
}
