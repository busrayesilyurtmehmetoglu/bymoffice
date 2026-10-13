// The question mount() asks for the password of a password-protected file,
// inside the container the file will be shown in, when the application gave
// no password and no `onPassword` of its own.
//
// A dialog with one field (type "password", autocomplete "new-password", no
// name, no form around it: never filled in by the browser, never
// spell-checked) and two answers. Enter in the field answers, Escape
// gives up. The field is emptied the moment its content is taken, and again
// when the dialog goes: the password is not left in the DOM. While the
// password is being checked the field takes nothing; a wrong password brings
// the question back with a line that says so. Every text reaches the DOM
// through textContent. Nothing here touches the DOM at import time.

export interface PasswordPromptStrings {
  passwordTitle: string;
  passwordQuestion: string;
  passwordLabel: string;
  passwordWrong(attempts: number): string;
  passwordChecking: string;
  passwordOpen: string;
  passwordCancel: string;
}

export interface PasswordPromptOptions {
  /** The element the question is put into and covers. Where it is not positioned, it is for as long as the question is there. */
  container: HTMLElement;
  strings: PasswordPromptStrings;
  /** Called when the user gives up, also while a password is being checked (when no question is open to answer). */
  onCancel?(): void;
}

/** What a question about a password is told. */
export interface PasswordRequest {
  /** How many passwords were tried and did not open the file. */
  attempts: number;
  /** Whether the last password that was tried was wrong (false the first time, unless a password of the options was). */
  wrong: boolean;
}

/** Events of the question that the page around it must not take for its own. */
const CONTAINED = ['keypress', 'keyup', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];

const stop = (event: Event): void => event.stopPropagation();

let nextId = 1;

export class PasswordPrompt {
  private readonly container: HTMLElement;
  private readonly backdrop: HTMLElement;
  private readonly field: HTMLInputElement;
  private readonly message: HTMLElement;
  private readonly openButton: HTMLButtonElement;
  private readonly cancelButton: HTMLButtonElement;
  private readonly strings: PasswordPromptStrings;
  private pending: ((answer: string | null) => void) | undefined;
  private closed = false;
  private shown = false;
  private unplaced: { value: string; priority: string } | undefined;
  private readonly before: HTMLElement | null;
  private readonly onCancel: (() => void) | undefined;

  constructor(options: PasswordPromptOptions) {
    const strings = (this.strings = options.strings);
    const container = (this.container = options.container);
    this.onCancel = options.onCancel;
    const root = container.getRootNode ? (container.getRootNode() as Document | ShadowRoot) : document;
    this.before = (root.activeElement as HTMLElement | null) || null;
    const coarse = typeof window !== 'undefined' && typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;
    const id = nextId++;

    const backdrop = (this.backdrop = document.createElement('div'));
    backdrop.className = 'bym-password-backdrop';
    // Everything is set inline: the question looks the same wherever it is, and where a style element is not allowed.
    backdrop.style.cssText =
      'position:absolute;left:0;top:0;right:0;bottom:0;z-index:40;display:flex;align-items:center;justify-content:center;box-sizing:border-box;' +
      'padding:8px;overflow:hidden;background:rgba(32,33,36,.08);color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
      'letter-spacing:normal;text-transform:none;text-align:left;direction:ltr';
    const panel = document.createElement('div');
    panel.className = 'bym-password';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'true');
    panel.setAttribute('data-bym-dialog', 'password');
    panel.style.cssText =
      'display:flex;flex-direction:column;box-sizing:border-box;width:340px;max-width:100%;max-height:100%;min-width:0;min-height:0;overflow:auto;' +
      'background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.28)';
    const title = document.createElement('div');
    title.className = 'bym-password-title';
    title.id = `bym-password-title-${id}`;
    title.style.cssText = 'flex:none;margin:0;padding:12px 16px 6px;font-size:15px;font-weight:600';
    title.textContent = strings.passwordTitle;
    const text = document.createElement('div');
    text.className = 'bym-password-text';
    text.id = `bym-password-text-${id}`;
    text.style.cssText = 'flex:none;margin:0;padding:0 16px 8px;overflow-wrap:break-word';
    text.textContent = strings.passwordQuestion;
    panel.setAttribute('aria-labelledby', title.id);
    panel.setAttribute('aria-describedby', text.id);

    const label = document.createElement('label');
    label.className = 'bym-password-label';
    label.style.cssText = 'display:block;flex:none;margin:0;padding:0 16px 4px;font-size:12px;color:#5f6368';
    label.textContent = strings.passwordLabel;
    const field = (this.field = document.createElement('input'));
    field.type = 'password';
    field.className = 'bym-password-field';
    field.id = `bym-password-field-${id}`;
    label.setAttribute('for', field.id);
    // Not remembered, not offered, not corrected: the password is the file's, not the page's. "new-password" is what
    // keeps a browser from filling in the password it has stored for the site ("off" is not honoured for passwords).
    // The field has no name and is in no form, so that nothing offers to save what is typed either.
    field.setAttribute('autocomplete', 'new-password');
    field.setAttribute('autocapitalize', 'off');
    field.setAttribute('autocorrect', 'off');
    field.setAttribute('spellcheck', 'false');
    field.setAttribute('data-bym-field', 'password');
    field.style.cssText =
      `display:block;flex:none;box-sizing:border-box;width:auto;height:${coarse ? 40 : 32}px;margin:0 16px;padding:0 8px;border:1px solid #80868b;border-radius:6px;` +
      'background:#fff;color:#202124;font:inherit;font-size:16px;letter-spacing:normal;outline-offset:1px';
    const message = (this.message = document.createElement('div'));
    message.className = 'bym-password-message';
    message.id = `bym-password-message-${id}`;
    // Read out when it changes: "wrong password", "checking".
    message.setAttribute('role', 'alert');
    message.setAttribute('aria-live', 'assertive');
    message.style.cssText = 'flex:none;min-height:18px;margin:0;padding:4px 16px 0;color:#b3261e;font-size:12px;overflow-wrap:break-word';
    field.setAttribute('aria-describedby', message.id);

    const actions = document.createElement('div');
    actions.className = 'bym-password-actions';
    actions.style.cssText = 'display:flex;flex:none;flex-wrap:wrap;justify-content:flex-end;align-items:center;margin:0;padding:4px 12px 12px';
    const button = (labelText: string, name: 'open' | 'cancel', primary: boolean): HTMLButtonElement => {
      const element = document.createElement('button');
      element.type = 'button';
      element.className = `bym-password-button bym-password-${name}`;
      element.setAttribute('data-bym-action', name);
      element.style.cssText =
        '-webkit-appearance:none;appearance:none;display:block;flex:none;box-sizing:border-box;float:none;width:auto;' +
        `height:${coarse ? 40 : 32}px;min-width:72px;min-height:0;max-width:100%;max-height:none;margin:8px 0 0 8px;padding:0 14px;` +
        'border:1px solid #dadce0;border-radius:6px;font:inherit;line-height:1.2;letter-spacing:normal;text-transform:none;text-indent:0;text-shadow:none;' +
        'text-decoration:none;opacity:1;white-space:nowrap;cursor:pointer;touch-action:manipulation;' +
        (primary ? 'background:#1a73e8;border-color:#1a73e8;color:#fff' : 'background:#fff;color:#202124');
      element.textContent = labelText;
      actions.appendChild(element);
      return element;
    };
    this.cancelButton = button(strings.passwordCancel, 'cancel', false);
    this.openButton = button(strings.passwordOpen, 'open', true);

    panel.appendChild(title);
    panel.appendChild(text);
    panel.appendChild(label);
    panel.appendChild(field);
    panel.appendChild(message);
    panel.appendChild(actions);
    backdrop.appendChild(panel);

    backdrop.addEventListener('keydown', this.onKeyDown);
    this.openButton.addEventListener('click', this.submit);
    this.cancelButton.addEventListener('click', this.cancel);
    for (let i = 0; i < CONTAINED.length; i++) backdrop.addEventListener(CONTAINED[i], stop);
  }

  /** Whether the question is in the page. */
  get open(): boolean {
    return this.shown && !this.closed;
  }

  /**
   * Shows the question, or shows it again after a password that was wrong.
   * Resolves to what was typed (which may be nothing: an empty password is a
   * password), or to null when the user gave up. The field is empty by the
   * time the promise resolves.
   */
  ask(request: PasswordRequest): Promise<string | null> {
    if (this.closed || this.cancelled) return Promise.resolve(null);
    // A question that is still open is given up: it has only one answer.
    if (this.pending) this.settle(null);
    this.show();
    const field = this.field;
    field.value = '';
    field.disabled = false;
    this.openButton.disabled = false;
    this.message.textContent = request.wrong ? this.strings.passwordWrong(request.attempts) : '';
    if (request.wrong) field.setAttribute('aria-invalid', 'true');
    else field.removeAttribute('aria-invalid');
    try {
      field.focus({ preventScroll: true });
    } catch {
      // A field that cannot take the focus is still there to be clicked.
    }
    return new Promise((resolve) => {
      this.pending = resolve;
    });
  }

  /** Takes the question away; a question that is open is answered with null. Safe to call twice. */
  close(): void {
    if (this.closed) return;
    this.settle(null);
    this.closed = true;
    // Whatever was typed and not sent goes with the dialog.
    this.field.value = '';
    const backdrop = this.backdrop;
    backdrop.removeEventListener('keydown', this.onKeyDown);
    this.openButton.removeEventListener('click', this.submit);
    this.cancelButton.removeEventListener('click', this.cancel);
    for (let i = 0; i < CONTAINED.length; i++) backdrop.removeEventListener(CONTAINED[i], stop);
    if (!this.shown) return;
    const root = backdrop.getRootNode ? (backdrop.getRootNode() as Document | ShadowRoot) : document;
    const active = root.activeElement as Node | null;
    const body = backdrop.ownerDocument ? backdrop.ownerDocument.body : null;
    if (!active || active === body || backdrop.contains(active)) {
      try {
        const before = this.before;
        if (before && before !== body && before.isConnected && typeof before.focus === 'function') before.focus({ preventScroll: true });
      } catch {
        // Nothing to give the focus to: it stays where the browser puts it.
      }
    }
    if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    const unplaced = this.unplaced;
    if (unplaced) {
      if (unplaced.value === '') this.container.style.removeProperty('position');
      else this.container.style.setProperty('position', unplaced.value, unplaced.priority);
    }
  }

  private show(): void {
    if (this.shown) return;
    this.shown = true;
    const container = this.container;
    if (typeof window !== 'undefined' && typeof window.getComputedStyle === 'function' && window.getComputedStyle(container).position === 'static') {
      this.unplaced = { value: container.style.getPropertyValue('position'), priority: container.style.getPropertyPriority('position') };
      container.style.setProperty('position', 'relative', 'important');
    }
    container.appendChild(this.backdrop);
  }

  private settle(answer: string | null): void {
    const pending = this.pending;
    this.pending = undefined;
    if (pending) pending(answer);
  }

  private readonly submit = (): void => {
    if (!this.pending) return;
    const field = this.field;
    const value = field.value;
    // Taken out of the DOM before anything is done with it.
    field.value = '';
    field.disabled = true;
    this.openButton.disabled = true;
    field.removeAttribute('aria-invalid');
    this.message.textContent = this.strings.passwordChecking;
    this.settle(value);
  };

  private readonly cancel = (): void => {
    this.field.value = '';
    // Also while a password is being checked: the caller hears of it when it asks next, or through close().
    this.cancelled = true;
    this.settle(null);
    if (this.onCancel) this.onCancel();
  };

  /** Set when the user gave up: no question is asked after that. */
  cancelled = false;

  private readonly onKeyDown = (event: KeyboardEvent): void => {
    // The page around the question hears of no key that is pressed in it.
    event.stopPropagation();
    if (event.key === 'Escape' || event.key === 'Esc') {
      event.preventDefault();
      this.cancel();
    } else if (event.key === 'Enter' && event.target === this.field && !event.isComposing) {
      event.preventDefault();
      this.submit();
    }
  };
}
