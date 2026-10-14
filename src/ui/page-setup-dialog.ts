// The dialog of the page setup (File -> Page Setup) of both editors: a form of fields in groups
// (lists, lengths in the unit of the language, numbers, check boxes, references) and a row of
// actions that act at once (a page break, a section break). Nothing is changed until "OK": then
// the fields whose value differs from what they showed when the dialog opened are handed over,
// as they read; a field the user did not touch is never handed over, so that a value the dialog
// shows rounded is not written back as another one. What the values mean is the business of the
// view that opened the dialog (see sheet-view.ts, doc-view.ts), which may refuse them with a
// message: the dialog then stays open and says it.
//
// It covers the editor inside its container, as the dialog of tab stops does; the editor behind
// it takes no input while it is open. Escape cancels; Enter in a text field is "OK". Every text
// reaches the DOM through textContent. Nothing here touches the DOM at import time.

/** One field of the dialog. `value` is what it shows when the dialog opens: text, or a check box's state. */
export interface PageDialogField {
  key: string;
  label: string;
  kind: 'select' | 'text' | 'check';
  value: string | boolean;
  /** The entries of a list. */
  options?: { value: string; label: string }[];
  /** For a text field: what the keyboard of a phone offers. */
  inputMode?: 'decimal' | 'numeric' | 'text';
  /** The field is shown and cannot be changed. */
  disabled?: boolean;
}

export interface PageDialogGroup {
  title: string;
  fields: PageDialogField[];
}

/** A button that acts at once and closes the dialog; what it does is undone like any edit. */
export interface PageDialogAction {
  key: string;
  label: string;
  run(): string | true;
}

export interface PageSetupDialogStrings {
  ok: string;
  cancel: string;
  /** The label of the row of actions. */
  actions: string;
}

export interface PageSetupDialogOptions {
  /** The element the dialog covers and is put into: the root element of the editor. */
  root: HTMLElement;
  title: string;
  strings: PageSetupDialogStrings;
  groups: PageDialogGroup[];
  actions?: PageDialogAction[];
  /** Explains what the fields that are disabled cannot do (a document edited together, say). */
  note?: string;
  /**
   * Called on "OK" with the fields whose value differs from what they showed (an empty object
   * when none does: the call is made anyway, and changes nothing). Returns true to close the
   * dialog, or a message that says what is wrong, which keeps it open; `field` names the field
   * the message is about, which gets the focus.
   */
  apply(changed: { [key: string]: string | boolean }): true | { message: string; field?: string };
  /** A field changed (a list, a check box): the dialog may show or hide others (see setHidden). */
  onInput?(key: string, dialog: PageSetupDialog): void;
  onClose(): void;
}

const CONTAINED = ['keypress', 'keyup', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];
const stop = (event: Event): void => event.stopPropagation();

let nextId = 1;

/** The fields whose value differs from the one they opened with; text is compared without the spaces at its ends. */
export function changedFields(groups: readonly PageDialogGroup[], values: { [key: string]: string | boolean }): { [key: string]: string | boolean } {
  const out: { [key: string]: string | boolean } = {};
  for (const group of groups) {
    for (const field of group.fields) {
      if (field.disabled || !Object.prototype.hasOwnProperty.call(values, field.key)) continue;
      const now = values[field.key];
      const was = field.value;
      if (typeof was === 'boolean' ? now !== was : String(now).trim() !== String(was).trim()) out[field.key] = typeof now === 'string' ? now.trim() : now;
    }
  }
  return out;
}

export class PageSetupDialog {
  private readonly options: PageSetupDialogOptions;
  private readonly backdrop: HTMLDivElement;
  private readonly message: HTMLDivElement;
  private readonly controls = new Map<string, HTMLInputElement | HTMLSelectElement>();
  private readonly rows = new Map<string, HTMLElement>();
  private readonly silenced: Element[] = [];
  private unplaced: string | undefined;
  private closed = false;

  constructor(options: PageSetupDialogOptions) {
    this.options = options;
    const id = nextId++;
    const coarse = typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;

    const backdrop = document.createElement('div');
    backdrop.className = 'bym-page-setup-backdrop';
    backdrop.style.cssText =
      'position:absolute;left:0;top:0;right:0;bottom:0;z-index:40;display:flex;align-items:center;justify-content:center;box-sizing:border-box;' +
      'padding:8px;overflow:hidden;background:rgba(32,33,36,.32);color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;' +
      'letter-spacing:normal;text-transform:none;text-align:left;direction:ltr';
    const panel = document.createElement('div');
    panel.className = 'bym-page-setup-dialog';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'true');
    panel.style.cssText =
      'display:flex;flex-direction:column;box-sizing:border-box;width:440px;max-width:100%;max-height:100%;min-width:0;overflow:auto;' +
      'background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.28);padding:12px 16px';
    const title = document.createElement('div');
    title.id = `bym-page-setup-title-${id}`;
    title.style.cssText = 'font-size:15px;font-weight:600;margin:0 0 8px';
    title.textContent = options.title;
    panel.setAttribute('aria-labelledby', title.id);
    panel.appendChild(title);

    const controlCss = `box-sizing:border-box;height:${coarse ? 40 : 30}px;font:inherit;border:1px solid #dadce0;border-radius:4px;padding:0 6px;background:#fff;color:#202124;min-width:0;width:100%`;
    const buttonCss = `-webkit-appearance:none;appearance:none;box-sizing:border-box;min-height:${coarse ? 40 : 30}px;padding:0 12px;font:inherit;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;cursor:pointer`;
    const button = (text: string, name: string): HTMLButtonElement => {
      const element = document.createElement('button');
      element.type = 'button';
      element.className = `bym-page-setup-${name}`;
      element.setAttribute('data-bym-action', name);
      element.style.cssText = buttonCss;
      element.textContent = text;
      return element;
    };

    let first: HTMLElement | undefined;
    for (const group of options.groups) {
      const fieldset = document.createElement('fieldset');
      fieldset.style.cssText = 'border:1px solid #dadce0;border-radius:6px;margin:0 0 8px;padding:6px 10px 8px;min-width:0';
      const legend = document.createElement('legend');
      legend.style.cssText = 'padding:0 4px;font-weight:600';
      legend.textContent = group.title;
      fieldset.appendChild(legend);
      const grid = document.createElement('div');
      grid.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px 10px';
      for (const field of group.fields) {
        const label = document.createElement('label');
        label.setAttribute('data-bym-field', field.key);
        let control: HTMLInputElement | HTMLSelectElement;
        if (field.kind === 'check') {
          label.style.cssText = 'display:flex;align-items:center;gap:6px;flex:1 1 180px;min-width:0';
          const input = document.createElement('input');
          input.type = 'checkbox';
          input.checked = field.value === true;
          input.className = `bym-page-setup-field bym-page-setup-${field.key}`;
          label.appendChild(input);
          const caption = document.createElement('span');
          caption.textContent = field.label;
          label.appendChild(caption);
          control = input;
        } else {
          label.style.cssText = 'display:flex;flex-direction:column;gap:2px;flex:1 1 120px;min-width:0';
          const caption = document.createElement('span');
          caption.textContent = field.label;
          label.appendChild(caption);
          if (field.kind === 'select') {
            const select = document.createElement('select');
            for (const entry of field.options || []) {
              const option = document.createElement('option');
              option.value = entry.value;
              option.textContent = entry.label;
              select.appendChild(option);
            }
            select.value = String(field.value);
            control = select;
          } else {
            const input = document.createElement('input');
            input.type = 'text';
            input.value = String(field.value);
            input.autocomplete = 'off';
            input.spellcheck = false;
            if (field.inputMode) input.setAttribute('inputmode', field.inputMode);
            input.addEventListener('keydown', (event) => {
              if (event.key === 'Enter') {
                event.preventDefault();
                this.confirm();
              }
            });
            control = input;
          }
          control.className = `bym-page-setup-field bym-page-setup-${field.key}`;
          control.style.cssText = controlCss;
          label.appendChild(control);
        }
        control.setAttribute('data-bym-key', field.key);
        if (field.disabled) control.disabled = true;
        control.addEventListener('change', () => {
          if (this.options.onInput) this.options.onInput(field.key, this);
        });
        if (!first && !field.disabled) first = control;
        this.controls.set(field.key, control);
        this.rows.set(field.key, label);
        grid.appendChild(label);
      }
      fieldset.appendChild(grid);
      panel.appendChild(fieldset);
    }

    if (options.actions && options.actions.length > 0) {
      const row = document.createElement('div');
      row.className = 'bym-page-setup-actions';
      row.setAttribute('role', 'group');
      row.setAttribute('aria-label', options.strings.actions);
      row.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;margin:0 0 8px;align-items:center';
      const caption = document.createElement('span');
      caption.style.cssText = 'flex:1 0 100%;font-weight:600';
      caption.textContent = options.strings.actions;
      row.appendChild(caption);
      for (const action of options.actions) {
        const element = button(action.label, `action-${action.key}`);
        element.addEventListener('click', () => {
          const result = action.run();
          if (result === true) this.close();
          else this.message.textContent = result;
        });
        row.appendChild(element);
      }
      panel.appendChild(row);
    }

    if (options.note) {
      const note = document.createElement('div');
      note.className = 'bym-page-setup-note';
      note.style.cssText = 'color:#5f6368;font-size:12px;margin:0 0 4px';
      note.textContent = options.note;
      panel.appendChild(note);
    }

    const message = document.createElement('div');
    message.className = 'bym-page-setup-message';
    message.setAttribute('role', 'alert');
    message.style.cssText = 'min-height:18px;margin:4px 0;color:#b3261e;font-size:12px';
    this.message = message;
    panel.appendChild(message);

    const footer = document.createElement('div');
    footer.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';
    const ok = button(options.strings.ok, 'ok');
    ok.style.background = '#1a73e8';
    ok.style.borderColor = '#1a73e8';
    ok.style.color = '#fff';
    const cancel = button(options.strings.cancel, 'cancel');
    footer.appendChild(cancel);
    footer.appendChild(ok);
    panel.appendChild(footer);
    backdrop.appendChild(panel);
    this.backdrop = backdrop;

    ok.addEventListener('click', () => this.confirm());
    cancel.addEventListener('click', () => this.close());
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
    if (this.options.onInput) for (const key of this.controls.keys()) this.options.onInput(key, this);
    (first || ok).focus({ preventScroll: true });
  }

  /** What a field holds now: its text, the value of its list, or whether its box is checked. */
  value(key: string): string | boolean | undefined {
    const control = this.controls.get(key);
    if (!control) return undefined;
    return control instanceof HTMLInputElement && control.type === 'checkbox' ? control.checked : control.value;
  }

  /** Shows or hides a field (one that is hidden is still handed over when it was changed before). */
  setHidden(key: string, hidden: boolean): void {
    const row = this.rows.get(key);
    if (row) row.style.display = hidden ? 'none' : 'flex';
  }

  get isOpen(): boolean {
    return !this.closed;
  }

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

  private confirm(): void {
    const values: { [key: string]: string | boolean } = {};
    for (const key of this.controls.keys()) values[key] = this.value(key) as string | boolean;
    const result = this.options.apply(changedFields(this.options.groups, values));
    if (result === true) {
      this.close();
      return;
    }
    this.message.textContent = result.message;
    const control = result.field ? this.controls.get(result.field) : undefined;
    if (control) {
      control.focus({ preventScroll: true });
      if (control instanceof HTMLInputElement && control.type === 'text') control.select();
    }
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

/** Twips in a unit of the dialogs. */
const TWIPS_PER: { [unit in 'cm' | 'in']: number } = { cm: 1440 / 2.54, in: 1440 };

/** A length in twips as a dialog shows it: in the unit, with two decimals at most, the decimal separator of the language. */
export function formatTwips(twips: number, unit: 'cm' | 'in', decimal: string): string {
  const value = Math.round((twips / TWIPS_PER[unit]) * 100) / 100;
  return String(value).replace('.', decimal);
}

/** A length typed in a unit, in inches; undefined for what is no length of 0 or more. */
export function parseInches(text: string, unit: 'cm' | 'in'): number | undefined {
  const trimmed = text.trim().replace(/\s+/g, '');
  if (!/^\d{1,6}([.,]\d{0,4})?$/.test(trimmed)) return undefined;
  const value = Number(trimmed.replace(',', '.'));
  return Number.isFinite(value) ? (unit === 'cm' ? value / 2.54 : value) : undefined;
}

/** A length typed in a unit, in twips (rounded to a whole twip); undefined for what is no length of 0 or more. */
export function parseTwips(text: string, unit: 'cm' | 'in'): number | undefined {
  const inches = parseInches(text, unit);
  return inches === undefined ? undefined : Math.round(inches * TWIPS_PER.in);
}

/** A whole number typed (digits only); undefined otherwise. */
export function parseWhole(text: string): number | undefined {
  const trimmed = text.trim();
  return /^\d{1,6}$/.test(trimmed) ? Number(trimmed) : undefined;
}
