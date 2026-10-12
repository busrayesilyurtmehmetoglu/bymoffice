// When a drag is over. A view that follows a pointer hears of its end from the window,
// and a pointer does not always end there: the capture can be lost, the window can lose
// the focus with the button still down, the page can be hidden. Whoever misses those
// keeps "dragging" for ever, and everything that waits for the drag to end waits with it.

type Listener = (event: Event) => void;

interface Listens {
  addEventListener(type: string, listener: Listener): void;
  removeEventListener(type: string, listener: Listener): void;
}

export interface DragEndTargets {
  window: Listens;
  document: Listens & { visibilityState?: string };
  /** The element the drag began on: where a lost capture is told. */
  element: Listens;
  /** Whether the pointer is followed with pointer events; mouse events otherwise. */
  pointerEvents: boolean;
}

/**
 * Calls `end` when the drag that has just begun is over, however it ends: the pointer is let
 * go or cancelled, its capture is lost, the window loses the focus, the page is hidden. `end`
 * may be called more than once. The result takes the listeners away again; safe to call twice.
 */
export function watchDragEnd(targets: DragEndTargets, end: () => void): () => void {
  const { window: win, document: doc, element } = targets;
  const onEnd: Listener = () => end();
  const onVisibility: Listener = () => {
    if (doc.visibilityState === 'hidden') end();
  };
  const up = targets.pointerEvents ? 'pointerup' : 'mouseup';
  win.addEventListener(up, onEnd);
  if (targets.pointerEvents) {
    win.addEventListener('pointercancel', onEnd);
    element.addEventListener('lostpointercapture', onEnd);
  }
  win.addEventListener('blur', onEnd);
  doc.addEventListener('visibilitychange', onVisibility);
  let watching = true;
  return () => {
    if (!watching) return;
    watching = false;
    win.removeEventListener(up, onEnd);
    if (targets.pointerEvents) {
      win.removeEventListener('pointercancel', onEnd);
      element.removeEventListener('lostpointercapture', onEnd);
    }
    win.removeEventListener('blur', onEnd);
    doc.removeEventListener('visibilitychange', onVisibility);
  };
}
