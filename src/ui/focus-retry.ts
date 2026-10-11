// The focus a press asks for, in a frame that is not the active one.
//
// Firefox does not move the focus to an element of a frame while a press on
// that frame is being handled and the frame is not the active one yet: focus()
// on the hidden input does nothing, the frame becomes active with the release
// and its body has the focus. What the user types next goes nowhere. (Other
// engines move it at once.) The focus is asked for again when the press ends.

/**
 * After `input.focus()` was called: when the input did not get the focus,
 * `focus` is called once more at the end of the press, unless the focus has
 * gone to something else by then or `alive` says the view is gone.
 */
export function focusOnceActive(input: HTMLElement, focus: () => void, alive: () => boolean): void {
  const doc = input.ownerDocument;
  if (!doc || typeof doc.addEventListener !== 'function') return;
  const active = (): Element | null => (input.getRootNode() as Document | ShadowRoot).activeElement;
  if (active() === input) return;
  let done = false;
  const retry = (): void => {
    if (done) return;
    done = true;
    doc.removeEventListener('pointerup', retry, true);
    doc.removeEventListener('mouseup', retry, true);
    doc.removeEventListener('touchend', retry, true);
    if (!alive()) return;
    const now = active();
    // Something else took the focus meanwhile (a menu, a field of the application): it keeps it.
    if (now && now !== input && now !== doc.body && now !== doc.documentElement) return;
    if (now !== input) focus();
  };
  doc.addEventListener('pointerup', retry, true);
  doc.addEventListener('mouseup', retry, true);
  doc.addEventListener('touchend', retry, true);
}
