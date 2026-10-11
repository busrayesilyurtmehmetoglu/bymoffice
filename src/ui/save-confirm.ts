// The question an editor asks when a save ran into a conflict: the file that
// is saved to was changed since it was opened, and the user decides whether
// their own file goes over it. It covers the editor inside its container, and
// is the same for a workbook and a document, with or without their toolbars.
//
// An alert dialog with two answers. The focus goes to "cancel" and comes back
// to where it was; Escape cancels. The focus is not trapped: Tab leaves the
// question for the page around the editor. The editor behind the question
// takes no input while it is open. Every text reaches the DOM through
// textContent. Nothing here touches the DOM at import time.
//
// The question comes when a server answers, which is not when the user
// expects it: they may be in the middle of moving through cells with the
// arrow keys, or of a second press on "save". Nothing they were doing for the
// editor may end up as "overwrite". So only Tab moves between the two answers
// (the arrow keys do nothing here), "overwrite" answers nothing for a moment
// after the question came, and a key that is held down answers nothing at all.

export interface OverwritePromptStrings {
  saveConflictTitle: string;
  saveConflictQuestion: string;
  saveOverwrite: string;
  saveCancel: string;
}

export interface OverwritePromptOptions {
  /**
   * The element the question covers and is put into: the root element of the editor (of a view
   * without its shell, the root of the view). Everything else in it takes no input meanwhile.
   * Where it is not positioned, it is for as long as the question is open.
   */
  root: HTMLElement;
  strings: OverwritePromptStrings;
  /** What the application said of the conflict: shown under the question, as text. May be empty. */
  detail?: string;
  /** Called with the answer, once: true to overwrite, false to leave the file unsaved. */
  answer(overwrite: boolean): void;
  /** Gives the focus to the view, for when the element that had it before the question is gone. */
  focusView?(): void;
  /** Calls back after `ms` and returns what calls that off. Default: the window's timer. */
  timer?(callback: () => void, ms: number): () => void;
}

/**
 * For how long after the question came "overwrite" answers nothing: a press that was on its way
 * to the editor (the second of a double press on "save", Enter after an arrow key) is not an answer.
 */
export const OVERWRITE_ARM_MS = 400;

/** How far up the root of an editor is looked for. */
const MAX_DEPTH = 64;

/**
 * The root element of the editor a view belongs to: the element that holds the view's own root
 * (`element`) and was put into the container, with the shell around the view when there is one.
 * A view that names no element inside the container leaves the container itself.
 */
export function editorRootIn(container: HTMLElement, element: HTMLElement | null | undefined): HTMLElement {
  let root: HTMLElement | null = element || null;
  for (let i = 0; root && root.parentElement && root.parentElement !== container && i < MAX_DEPTH; i++) root = root.parentElement;
  return root && root.parentElement === container ? root : container;
}

/** How much of what an application said is shown. */
const MAX_DETAIL_CHARS = 200;

const stop = (event: Event): void => event.stopPropagation();

/** Events of the question that the editor around it must not take for its own. */
const CONTAINED = ['keypress', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];

/** Keys that move through an editor, and would scroll the page: in the question they do nothing. */
const NAVIGATION = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Left', 'Right', 'Up', 'Down', 'Home', 'End', 'PageUp', 'PageDown'];

let nextId = 1;

const defaultTimer = (callback: () => void, ms: number): (() => void) => {
  const handle = setTimeout(callback, ms);
  return () => clearTimeout(handle);
};

/** A key whose press on a button is a press of the button. */
const pressesButton = (key: string): boolean => key === 'Enter' || key === ' ' || key === 'Spacebar';

/**
 * Asks whether the file is to be saved over what the server holds. Returns what takes the question
 * away again (safe to call twice); the caller calls it when the question is answered as well.
 */
export function askOverwrite(options: OverwritePromptOptions): () => void {
  const parent = options.root;
  const strings = options.strings;
  const coarse = typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;
  const root = parent.getRootNode() as Document | ShadowRoot;
  /** Where the focus was: it goes back there. */
  const before = root.activeElement as HTMLElement | null;

  const backdrop = document.createElement('div');
  backdrop.className = 'bym-confirm-backdrop';
  // Everything is set inline: the question looks the same in both editors, and where a style element is not allowed.
  backdrop.style.cssText =
    'position:absolute;left:0;top:0;right:0;bottom:0;z-index:40;display:flex;align-items:center;justify-content:center;box-sizing:border-box;' +
    'padding:8px;overflow:hidden;background:rgba(32,33,36,.32);color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
    'letter-spacing:normal;text-transform:none;text-align:left;direction:ltr;-webkit-user-select:text;user-select:text';
  const panel = document.createElement('div');
  panel.className = 'bym-confirm';
  panel.setAttribute('role', 'alertdialog');
  panel.setAttribute('data-bym-confirm', 'save-conflict');
  panel.style.cssText =
    'display:flex;flex-direction:column;box-sizing:border-box;width:360px;max-width:100%;max-height:100%;min-width:0;min-height:0;overflow:hidden;' +
    'background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.28)';
  const id = nextId++;
  const title = document.createElement('div');
  title.className = 'bym-confirm-title';
  title.id = `bym-confirm-title-${id}`;
  title.style.cssText = 'flex:none;margin:0;padding:12px 16px 6px;font-size:15px;font-weight:600';
  title.textContent = strings.saveConflictTitle;
  const body = document.createElement('div');
  body.className = 'bym-confirm-body';
  body.id = `bym-confirm-text-${id}`;
  // A container that is lower than the question scrolls its text; the answers stay in view.
  body.style.cssText = 'flex:1 1 auto;min-height:0;margin:0;padding:0 16px 4px;overflow-x:hidden;overflow-y:auto;overflow-wrap:break-word;overscroll-behavior:contain';
  const text = document.createElement('div');
  text.className = 'bym-confirm-text';
  text.textContent = strings.saveConflictQuestion;
  body.appendChild(text);
  // One line of plain text, never markup: it comes from the application.
  let detail = typeof options.detail === 'string' ? options.detail.replace(/[\u0000-\u001F\u007F]+/g, ' ').trim() : '';
  if (detail.length > MAX_DETAIL_CHARS) detail = `${detail.slice(0, MAX_DETAIL_CHARS)}…`;
  if (detail !== '') {
    const said = document.createElement('div');
    said.className = 'bym-confirm-detail';
    said.style.cssText = 'margin:6px 0 0;color:#5f6368;font-size:12px';
    said.textContent = detail;
    body.appendChild(said);
  }
  const actions = document.createElement('div');
  actions.className = 'bym-confirm-actions';
  actions.style.cssText = 'display:flex;flex:none;flex-wrap:wrap;justify-content:flex-end;align-items:center;margin:0;padding:4px 12px 12px';
  panel.setAttribute('aria-labelledby', title.id);
  panel.setAttribute('aria-describedby', body.id);

  let closed = false;
  /** "Overwrite" answers: not for a moment after the question came (see OVERWRITE_ARM_MS). */
  let armed = false;
  /** A key that presses a button is held down and repeats: what it leads to is no answer. */
  let held = false;
  const reply = (overwrite: boolean): void => {
    if (closed) return;
    try {
      options.answer(overwrite);
    } finally {
      // Whoever asked takes the question away with the answer; when it does not (or fails), the question still goes.
      close();
    }
  };
  const button = (label: string, name: 'overwrite' | 'cancel', overwrite: boolean): HTMLButtonElement => {
    const element = document.createElement('button');
    element.type = 'button';
    element.className = `bym-confirm-button bym-confirm-${name}`;
    element.setAttribute('data-bym-action', name);
    // Every property a rule of the host page for bare buttons could reach is stated.
    element.style.cssText =
      '-webkit-appearance:none;appearance:none;display:block;flex:none;box-sizing:border-box;float:none;width:auto;' +
      `height:${coarse ? 40 : 32}px;min-width:72px;min-height:0;max-width:100%;max-height:none;margin:8px 0 0 8px;padding:0 14px;` +
      'border:1px solid #dadce0;border-radius:6px;font:inherit;line-height:1.2;letter-spacing:normal;text-transform:none;text-indent:0;text-shadow:none;' +
      'text-decoration:none;opacity:1;white-space:nowrap;cursor:pointer;touch-action:manipulation;' +
      (overwrite ? 'background:#b3261e;border-color:#b3261e;color:#fff' : 'background:#fff;color:#202124');
    element.textContent = label;
    element.addEventListener('click', (event) => {
      // Too early to be an answer to the question, or the end of a key that was held down (a click
      // that comes from a key says `detail` 0). "Cancel" changes nothing and always answers.
      if (overwrite && (!armed || (held && event.detail === 0))) return;
      reply(overwrite);
    });
    actions.appendChild(element);
    return element;
  };
  // In reading order the answer that changes nothing comes last, and has the focus.
  const overwriteButton = button(strings.saveOverwrite, 'overwrite', true);
  const cancelButton = button(strings.saveCancel, 'cancel', false);
  // Not `disabled`: such a button cannot be reached with Tab, and one that turns enabled is not announced.
  // It is there and can be focused; it is said and shown to be unavailable until it answers.
  overwriteButton.setAttribute('aria-disabled', 'true');
  overwriteButton.style.opacity = '.55';
  overwriteButton.style.cursor = 'default';

  panel.appendChild(title);
  panel.appendChild(body);
  panel.appendChild(actions);
  backdrop.appendChild(panel);

  const onKeyDown = (event: KeyboardEvent): void => {
    // The editor behind the question hears of no key that is pressed in it.
    event.stopPropagation();
    // A key that is held down was pressed for something else, or by accident: it answers nothing,
    // moves nothing, and what its release would press is not pressed either.
    if (event.repeat) {
      if (pressesButton(event.key)) held = true;
      event.preventDefault();
      return;
    }
    held = false;
    if (event.key === 'Escape' || event.key === 'Esc') {
      event.preventDefault();
      reply(false);
      return;
    }
    // The keys that move through cells and pages do nothing here: they neither move the focus between
    // the two answers (Tab does) nor scroll the page behind the question.
    if (NAVIGATION.indexOf(event.key) >= 0) event.preventDefault();
  };
  const onKeyUp = (event: KeyboardEvent): void => {
    event.stopPropagation();
    // Space presses a button when it is let go.
    if (held && pressesButton(event.key)) event.preventDefault();
  };
  // No press takes the focus out of the question: Escape still cancels after it. A press on an answer
  // gives that answer the focus in every browser (Safari gives a button that is pressed none, and takes
  // it from the one that had it); a press beside the answers leaves it where it is.
  const onMouseDown = (event: MouseEvent): void => {
    const target = event.target as Node | null;
    event.preventDefault();
    const pressed = target && overwriteButton.contains(target) ? overwriteButton : target && cancelButton.contains(target) ? cancelButton : undefined;
    if (pressed) return pressed.focus({ preventScroll: true });
    const active = (backdrop.getRootNode() as Document | ShadowRoot).activeElement;
    if (!active || !backdrop.contains(active)) cancelButton.focus({ preventScroll: true });
  };
  backdrop.addEventListener('keydown', onKeyDown);
  backdrop.addEventListener('keyup', onKeyUp);
  backdrop.addEventListener('mousedown', onMouseDown);
  for (let i = 0; i < CONTAINED.length; i++) backdrop.addEventListener(CONTAINED[i], stop);

  // The editor behind the question takes no input and no focus while it is open. Tab still leaves the
  // question, for the page around the editor: the focus is not trapped.
  const silenced: Element[] = [];
  const inertKnown = typeof HTMLElement !== 'undefined' && 'inert' in HTMLElement.prototype;
  const children = parent.children;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child.hasAttribute('inert')) continue;
    child.setAttribute('inert', '');
    silenced.push(child);
  }
  // A browser that does not know `inert`: a focus that goes into the editor comes back to the question.
  const onFocusIn = (event: Event): void => {
    const target = event.target as Node | null;
    if (closed || !target || backdrop.contains(target)) return;
    cancelButton.focus({ preventScroll: true });
  };
  if (!inertKnown) parent.addEventListener('focusin', onFocusIn);

  // The question is placed by the edges of the root: a root that is not positioned (a rule of the page
  // can undo what the editor set) is, for as long as the question is open.
  let unplaced: { value: string; priority: string } | undefined;
  if (typeof window !== 'undefined' && typeof window.getComputedStyle === 'function' && window.getComputedStyle(parent).position === 'static') {
    unplaced = { value: parent.style.getPropertyValue('position'), priority: parent.style.getPropertyPriority('position') };
    parent.style.setProperty('position', 'relative', 'important');
  }

  parent.appendChild(backdrop);
  cancelButton.focus({ preventScroll: true });
  let disarm: (() => void) | undefined = (options.timer || defaultTimer)(() => {
    disarm = undefined;
    if (closed) return;
    armed = true;
    overwriteButton.removeAttribute('aria-disabled');
    overwriteButton.style.opacity = '1';
    overwriteButton.style.cursor = 'pointer';
  }, OVERWRITE_ARM_MS);

  function close(): void {
    if (closed) return;
    closed = true;
    armed = false;
    if (disarm) disarm();
    disarm = undefined;
    backdrop.removeEventListener('keydown', onKeyDown);
    backdrop.removeEventListener('keyup', onKeyUp);
    backdrop.removeEventListener('mousedown', onMouseDown);
    for (let i = 0; i < CONTAINED.length; i++) backdrop.removeEventListener(CONTAINED[i], stop);
    if (!inertKnown) parent.removeEventListener('focusin', onFocusIn);
    for (let i = 0; i < silenced.length; i++) silenced[i].removeAttribute('inert');
    silenced.length = 0;
    // The focus goes back while the question is still in the document: an element that is removed with
    // the focus on it leaves the page without one. A focus the user moved elsewhere meanwhile stays there.
    const active = (backdrop.getRootNode() as Document | ShadowRoot).activeElement as Node | null;
    const body = backdrop.ownerDocument ? backdrop.ownerDocument.body : null;
    if (!active || active === body || backdrop.contains(active)) {
      try {
        if (before && before !== body && before.isConnected && typeof before.focus === 'function') before.focus({ preventScroll: true });
        else if (options.focusView) options.focusView();
      } catch {
        // Nothing to give the focus to: it stays where the browser puts it.
      }
    }
    if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    if (unplaced) {
      if (unplaced.value === '') parent.style.removeProperty('position');
      else parent.style.setProperty('position', unplaced.value, unplaced.priority);
    }
  }
  return close;
}
