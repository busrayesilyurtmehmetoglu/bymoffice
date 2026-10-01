// The formula bar: the name box (address of the selection, and a way to go
// to one) and a field with the content of the active cell, which edits the
// cell together with the editor inside the grid. The field is one line high
// but holds every line of the cell: a text with line breaks keeps them.

import { cellName, parseRange, rangeName } from '../model/sheet/types';
import type { EditorStrings } from './editor-strings';
import type { ShellHost } from './toolbar';

/** Longest address the name box takes ("XFD1048576:XFD1048576" is 21 characters). */
const MAX_NAME_LENGTH = 32;

export class FormulaBar {
  readonly element: HTMLDivElement;
  private readonly host: ShellHost;
  private readonly nameBox: HTMLInputElement;
  private readonly content: HTMLTextAreaElement;
  /** Address of the selection, which the name box goes back to when what was typed is not one. */
  private name = '';
  private destroyed = false;

  constructor(host: ShellHost, strings: EditorStrings) {
    this.host = host;
    const element = document.createElement('div');
    element.className = 'bym-formula';
    element.setAttribute('role', 'group');

    const nameBox = document.createElement('input');
    nameBox.type = 'text';
    nameBox.className = 'bym-namebox';
    nameBox.maxLength = MAX_NAME_LENGTH;
    plainInput(nameBox);

    const fx = document.createElement('span');
    fx.className = 'bym-fx';
    fx.setAttribute('aria-hidden', 'true');
    fx.textContent = 'fx';

    // A textarea, not an input: an input drops the line breaks of what it is given.
    const content = document.createElement('textarea');
    content.className = 'bym-formula-input';
    content.rows = 1;
    content.setAttribute('wrap', 'off');
    plainInput(content);

    element.appendChild(nameBox);
    element.appendChild(fx);
    element.appendChild(content);
    this.element = element;
    this.nameBox = nameBox;
    this.content = content;
    this.setStrings(strings);

    nameBox.addEventListener('keydown', this.onNameKeyDown);
    nameBox.addEventListener('focus', this.onNameFocus);
    nameBox.addEventListener('blur', this.onNameBlur);
    content.addEventListener('input', this.onContentInput);
    content.addEventListener('keydown', this.onContentKeyDown);
  }

  setStrings(strings: EditorStrings): void {
    this.element.setAttribute('aria-label', strings.formulaBar);
    this.nameBox.setAttribute('aria-label', strings.nameBox);
    this.nameBox.title = strings.nameBox;
    this.content.setAttribute('aria-label', strings.cellContent);
  }

  /** Shows the selection's address and the active cell's content. */
  update(single: boolean): void {
    if (this.destroyed) return;
    const view = this.host.view;
    const selection = view.selection;
    this.name = single ? cellName(selection.active.row, selection.active.col) : rangeName(selection.range);
    if (!this.focused(this.nameBox)) this.nameBox.value = this.name;
    this.content.readOnly = view.readOnly;
    // While the field is being typed in, the text in it is the newest there is.
    if (!(view.isEditing && this.focused(this.content))) this.setContent(view.editText);
  }

  /** The grid's editor opened, changed or closed. */
  editingChanged(state: { editing: boolean; text: string }): void {
    if (this.destroyed) return;
    // Typing here comes back as an echo; writing it into the field again would move the caret.
    if (state.editing && this.focused(this.content)) return;
    this.setContent(state.editing ? state.text : this.host.view.editText);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.nameBox.removeEventListener('keydown', this.onNameKeyDown);
    this.nameBox.removeEventListener('focus', this.onNameFocus);
    this.nameBox.removeEventListener('blur', this.onNameBlur);
    this.content.removeEventListener('input', this.onContentInput);
    this.content.removeEventListener('keydown', this.onContentKeyDown);
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
  }

  private focused(input: HTMLInputElement | HTMLTextAreaElement): boolean {
    const root = input.getRootNode() as Document | ShadowRoot;
    return root.activeElement === input;
  }

  private setContent(text: string): void {
    // As the field will hold it: a text field turns every kind of line break into "\n".
    const value = String(text === undefined || text === null ? '' : text).replace(/\r\n?/g, '\n');
    if (this.content.value !== value) {
      this.content.value = value;
      this.content.scrollTop = 0;
    }
  }

  private onNameFocus = (): void => {
    this.nameBox.select();
  };

  private onNameBlur = (): void => {
    this.nameBox.value = this.name;
  };

  private onNameKeyDown = (event: KeyboardEvent): void => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      const range = parseRange(this.nameBox.value.trim());
      if (!range) {
        // Not an address: the old one comes back, selected, ready to be typed over.
        this.nameBox.value = this.name;
        this.nameBox.select();
        return;
      }
      // The blur that follows puts the new address in the box.
      this.name = rangeName(range);
      this.host.run(() => this.host.view.select(range));
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.nameBox.value = this.name;
      this.host.view.focus();
    }
  };

  private onContentInput = (): void => {
    const view = this.host.view;
    if (this.destroyed || view.readOnly) return;
    view.setEditText(this.content.value);
  };

  private onContentKeyDown = (event: KeyboardEvent): void => {
    // Enter that confirms an input method's candidate is not Enter for the cell.
    if (event.isComposing || event.keyCode === 229) return;
    const view = this.host.view;
    // Tab only belongs to the cell while one is edited; otherwise it moves through the page as usual.
    if (event.key === 'Enter' && event.altKey) {
      // A line break in the cell, as in the grid's own editor.
      event.preventDefault();
      if (view.readOnly) return;
      const field = this.content;
      field.setRangeText('\n', field.selectionStart, field.selectionEnd, 'end');
      view.setEditText(field.value);
      // The field shows one line at a time, each as high as the field: the caret's line comes into view.
      field.scrollTop = field.value.slice(0, field.selectionStart).split('\n').length * field.clientHeight - field.clientHeight;
    } else if (event.key === 'Enter' || (event.key === 'Tab' && view.isEditing)) {
      event.preventDefault();
      // A cell that cannot be written keeps its editor, and the reason is shown.
      this.host.run(() => {
        view.commitEditing();
      });
    } else if (event.key === 'Escape') {
      event.preventDefault();
      view.cancelEditing();
      this.setContent(view.editText);
      view.focus();
    }
  };
}

function plainInput(input: HTMLInputElement | HTMLTextAreaElement): void {
  input.setAttribute('autocomplete', 'off');
  input.setAttribute('autocorrect', 'off');
  input.setAttribute('autocapitalize', 'off');
  input.spellcheck = false;
}
