// The toolbar of the document editor: undo and redo, the paragraph style, the
// font and its size, character and paragraph formatting, lists, tables, a
// picture, a link, a page break, the tracked changes (tracking on and off,
// the change before and after, accept, reject, and a menu for all of them) and
// the comments (a new one, the one before and after), and Find and Replace. It has the looks and the keys of the
// spreadsheet editor's toolbar: its buttons are a single Tab stop with the
// arrow keys moving between them, each list is a Tab stop of its own, and a
// command ends with the focus back in the document.
//
// Names of styles and fonts come from the document: they only ever reach the
// DOM as text.

import type { RangeFormat } from '../ops/doc/index';
import { withShortcut, type EditorStrings } from './editor-strings';
import type { DocumentView } from './doc-view';
import { ICON_BAR_CLASS, createDocIcon, createIcon, type DocIconName, type IconName } from './icons';
import { Popover, type PopoverCloseReason, type PopoverItem } from './popover';
import type { SaveCommand } from './save-command';
import { onFontsChanged } from '../render/fonts';
import { fontChoices } from './font-list';

/** What the toolbar needs from the editor that holds it. */
export interface DocShellHost {
  readonly view: DocumentView;
  /** The editor's root element: positioned, and the parent of everything that drops down. */
  readonly root: HTMLElement;
  /** Runs a command of the view; afterwards the focus goes back to the document (unless it is to be kept) and the toolbar is brought up to date. */
  run(action: () => void, keepFocus?: boolean): void;
  /** With it the toolbar starts with a Save button. */
  readonly save?: SaveCommand;
}

type ButtonKey =
  | 'save'
  | 'undo'
  | 'redo'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'textColor'
  | 'highlight'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'alignJustify'
  | 'bulletList'
  | 'numberedList'
  | 'decreaseIndent'
  | 'increaseIndent'
  | 'tabStops'
  | 'insertTable'
  | 'rowsColumns'
  | 'insertPicture'
  | 'insertLink'
  | 'pageBreak'
  | 'pageSetup'
  | 'trackChanges'
  | 'previousChange'
  | 'nextChange'
  | 'acceptChange'
  | 'rejectChange'
  | 'changesMenu'
  | 'newComment'
  | 'previousComment'
  | 'nextComment'
  | 'find';

type CommentKey = 'newComment' | 'previousComment' | 'nextComment';

/** The icons of the comment buttons, drawn like the editor's other icons (24 x 24, the colour of the text): a speech balloon with a plus, an arrow up, an arrow down. */
const COMMENT_ICONS: { [key in CommentKey]: string[] } = {
  newComment: ['M5 5h14v11h-7l-4 3.5V16H5z', 'M12 8v5 M9.5 10.5h5'],
  previousComment: ['M5 5h14v11h-7l-4 3.5V16H5z', 'M12 13.5v-6 M9.5 10l2.5-2.5 2.5 2.5'],
  nextComment: ['M5 5h14v11h-7l-4 3.5V16H5z', 'M12 7.5v6 M9.5 11l2.5 2.5 2.5-2.5'],
};

const SVG_NS = 'http://www.w3.org/2000/svg';

function commentIcon(key: CommentKey): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
  const attrs: { [name: string]: string } = {
    viewBox: '0 0 24 24',
    width: '20',
    height: '20',
    fill: 'none',
    stroke: 'currentColor',
    'stroke-width': '1.75',
    'stroke-linecap': 'round',
    'stroke-linejoin': 'round',
    'aria-hidden': 'true',
    focusable: 'false',
  };
  for (const name of Object.keys(attrs)) svg.setAttribute(name, attrs[name]);
  svg.setAttribute('class', 'bym-icon');
  for (const d of COMMENT_ICONS[key]) {
    const path = document.createElementNS(SVG_NS, 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  return svg;
}

const SHARED_ICONS: { [key: string]: IconName } = {
  save: 'save',
  undo: 'undo',
  redo: 'redo',
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strikethrough: 'strikethrough',
  textColor: 'textColor',
  alignLeft: 'alignLeft',
  alignCenter: 'alignCenter',
  alignRight: 'alignRight',
  rowsColumns: 'rowsColumns',
  pageSetup: 'pageSetup',
  find: 'find',
};

/** Font sizes offered, in points. */
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const MAX_LIST_ENTRIES = 120;

const PALETTE_COLUMNS = 10;
const PALETTE = (
  '000000 434343 666666 999999 B7B7B7 CCCCCC D9D9D9 EFEFEF F3F3F3 FFFFFF ' +
  '980000 FF0000 FF9900 FFFF00 00FF00 00FFFF 4A86E8 0000FF 9900FF FF00FF ' +
  'E6B8AF F4CCCC FCE5CD FFF2CC D9EAD3 D0E0E3 C9DAF8 CFE2F3 D9D2E9 EAD1DC ' +
  'DD7E6B EA9999 F9CB9C FFE599 B6D7A8 A2C4C9 A4C2F4 9FC5E8 B4A7D6 D5A6BD ' +
  'A61C00 CC0000 E69138 F1C232 6AA84F 45818E 3C78D8 3D85C6 674EA7 A64D79 ' +
  '5B0F00 660000 783F04 7F6000 274E13 0C343D 1C4587 073763 20124D 4C1130'
).split(' ');

/** The highlight colours of Word, by the names the file format has for them. */
const HIGHLIGHTS: [name: string, color: string][] = [
  ['yellow', '#FFFF00'],
  ['green', '#00FF00'],
  ['cyan', '#00FFFF'],
  ['magenta', '#FF00FF'],
  ['blue', '#0000FF'],
  ['red', '#FF0000'],
  ['darkBlue', '#000080'],
  ['darkCyan', '#008080'],
  ['darkGreen', '#008000'],
  ['darkMagenta', '#800080'],
  ['darkRed', '#800000'],
  ['darkYellow', '#808000'],
  ['darkGray', '#808080'],
  ['lightGray', '#C0C0C0'],
  ['black', '#000000'],
];

const TABLE_GRID_ROWS = 6;
const TABLE_GRID_COLUMNS = 8;
/** Largest picture read from a file the user picks, bytes. */
const MAX_PICTURE_FILE = 20000000;

const apple = (): boolean => typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');

function option(value: string, label: string): HTMLOptionElement {
  const element = document.createElement('option');
  element.value = value;
  element.textContent = label;
  return element;
}

type Control = HTMLButtonElement;

export class DocToolbar {
  readonly element: HTMLDivElement;
  private readonly host: DocShellHost;
  private strings: EditorStrings;
  /** Every button in visual order, for the arrow keys. */
  private readonly controls: Control[] = [];
  private readonly buttons: { [key in ButtonKey]?: HTMLButtonElement } = {};
  private readonly actions = new Map<HTMLButtonElement, () => void>();
  /** Buttons whose press opens something: their action runs as it is, not as a command of the view. */
  private readonly openers = new Set<HTMLButtonElement>();
  /** Buttons whose command puts the focus where it belongs itself. */
  private readonly keepsFocus = new Set<HTMLButtonElement>();
  private readonly styleSelect: HTMLSelectElement;
  private readonly fontSelect: HTMLSelectElement;
  private readonly unwatchFonts: () => void;
  private readonly sizeSelect: HTMLSelectElement;
  private readonly fileInput: HTMLInputElement;
  private readonly bars: { textColor?: SVGElement; highlight?: SVGElement } = {};
  private styleKey = '';
  private fontKey = '';
  private sizeKey = '';
  private popover: Popover | undefined;
  private linkField: HTMLDivElement | undefined;
  private linkWatch: ((event: Event) => void) | undefined;
  private openedAt = 0;
  private stop: Control | undefined;
  private listByKey = false;
  private destroyed = false;

  constructor(host: DocShellHost, strings: EditorStrings) {
    this.host = host;
    this.strings = strings;
    const element = document.createElement('div');
    element.className = 'bym-toolbar bym-doc-toolbar';
    element.setAttribute('role', 'toolbar');
    // Some browsers make an element that scrolls a Tab stop of its own; the stop is one of the controls.
    element.tabIndex = -1;
    this.element = element;

    const view = () => host.view;
    const save = host.save;
    if (save) {
      this.addButton('save', () => save.run()).setAttribute('aria-busy', 'false');
      this.addSeparator();
    }
    this.addButton('undo', () => view().undo());
    this.addButton('redo', () => view().redo());
    this.addSeparator();
    this.styleSelect = this.addSelect('bym-tb-style');
    this.fontSelect = this.addSelect('bym-tb-font');
    this.sizeSelect = this.addSelect('bym-tb-size');
    this.addSeparator();
    this.addButton('bold', () => view().toggleFormat('bold'), true);
    this.addButton('italic', () => view().toggleFormat('italic'), true);
    this.addButton('underline', () => view().toggleFormat('underline'), true);
    this.addButton('strikethrough', () => view().toggleFormat('strike'), true);
    this.addSeparator();
    this.addMenuButton('textColor', 'dialog', () => this.openColors());
    this.addMenuButton('highlight', 'dialog', () => this.openHighlights());
    this.addSeparator();
    this.addButton('alignLeft', () => view().formatParagraph({ align: 'left' }), true);
    this.addButton('alignCenter', () => view().formatParagraph({ align: 'center' }), true);
    this.addButton('alignRight', () => view().formatParagraph({ align: 'right' }), true);
    this.addButton('alignJustify', () => view().formatParagraph({ align: 'both' }), true);
    this.addSeparator();
    this.addButton('bulletList', () => view().toggleList('bullet'), true);
    this.addButton('numberedList', () => view().toggleList('number'), true);
    this.addButton('decreaseIndent', () => view().indent(-1));
    this.addButton('increaseIndent', () => view().indent(1));
    // A dialog of the view: it takes the focus, and hands it back when it closes.
    this.addMenuButton('tabStops', 'dialog', () => view().openTabStopsDialog());
    this.addSeparator();
    this.addMenuButton('insertTable', 'dialog', () => this.openTableGrid());
    this.addMenuButton('rowsColumns', 'menu', () => this.openRowsColumns());
    this.addSeparator();
    this.addMenuButton('insertPicture', 'dialog', () => this.fileInput.click());
    this.addMenuButton('insertLink', 'dialog', () => this.openLink());
    this.addButton('pageBreak', () => view().insertPageBreak());
    // The dialog of the page setup (and of section breaks) takes the focus, and hands it back when it closes.
    this.addMenuButton('pageSetup', 'dialog', () => view().openPageSetup());
    this.addSeparator();
    // The comments, before the tracked changes as in Word. The field of a new one takes the focus: the command
    // does not hand it back to the document.
    this.addButton('newComment', () => view().startComment());
    this.keepsFocus.add(this.buttons.newComment as HTMLButtonElement);
    this.addButton('previousComment', () => view().goToComment('previous'));
    this.addButton('nextComment', () => view().goToComment('next'));
    this.addSeparator();
    // The tracked changes: Word's order. Accept and reject go on to the next change, as they do there.
    this.addButton('trackChanges', () => view().setTrackChanges(!view().trackingChanges), true);
    this.addButton('previousChange', () => view().goToRevision('previous'));
    this.addButton('nextChange', () => view().goToRevision('next'));
    this.addButton('acceptChange', () => view().acceptRevision(undefined, true));
    this.addButton('rejectChange', () => view().rejectRevision(undefined, true));
    this.addMenuButton('changesMenu', 'menu', () => this.openChangesMenu());
    this.addSeparator();
    // The bar of Find and Replace takes the focus itself.
    const find = this.addButton('find', () => view().openFind(false));
    find.setAttribute('aria-haspopup', 'dialog');
    this.openers.add(find);

    // The picture button only opens the browser's own file dialog: it has nothing that stays open.
    const pictureButton = this.buttons.insertPicture as HTMLButtonElement;
    pictureButton.removeAttribute('aria-expanded');

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/png,image/jpeg,image/gif,image/bmp';
    fileInput.className = 'bym-tb-file';
    fileInput.tabIndex = -1;
    fileInput.setAttribute('aria-hidden', 'true');
    fileInput.style.cssText = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;border:0;overflow:hidden;clip:rect(0 0 0 0);opacity:0';
    element.appendChild(fileInput);
    this.fileInput = fileInput;

    this.bars.textColor = this.barOf('textColor');
    this.bars.highlight = this.barOf('highlight');

    element.addEventListener('click', this.onClick);
    element.addEventListener('mousedown', this.onMouseDown);
    element.addEventListener('keydown', this.onKeyDown);
    element.addEventListener('focusin', this.onFocusIn);
    element.addEventListener('scroll', this.onScroll, { passive: true });
    this.styleSelect.addEventListener('change', this.onStyleChange);
    this.fontSelect.addEventListener('change', this.onFontChange);
    this.sizeSelect.addEventListener('change', this.onSizeChange);
    fileInput.addEventListener('change', this.onFileChange);
    // The list of fonts follows the fonts that are registered.
    this.unwatchFonts = onFontsChanged(() => this.update());

    this.applyStrings();
  }

  setStrings(strings: EditorStrings): void {
    this.strings = strings;
    this.closePopover();
    this.styleKey = '';
    this.applyStrings();
    this.update();
  }

  /** Brings every control in line with the selection, the history and the read-only state. */
  update(): void {
    if (this.destroyed) return;
    const view = this.host.view;
    const editable = !view.readOnly;
    const format = editable ? view.selectionFormat : undefined;
    const mixed = format ? format.mixed : [];

    this.setDisabled('undo', !view.canUndo);
    this.setDisabled('redo', !view.canRedo);
    const save = this.host.save;
    const saveButton = this.buttons.save;
    if (save && saveButton) {
      // Dimmed while there is nothing to save, busy while a save is under way (a press then saves once more
      // when something was changed meanwhile). The whole toolbar is gone from a document that is read-only.
      this.setDisabled('save', !(editable || view.canComment) || !(save.dirty || save.saving));
      const busy = save.saving ? 'true' : 'false';
      if (saveButton.getAttribute('aria-busy') !== busy) saveButton.setAttribute('aria-busy', busy);
    }
    const others: ButtonKey[] = [
      'bold', 'italic', 'underline', 'strikethrough', 'textColor', 'highlight', 'alignLeft', 'alignCenter', 'alignRight', 'alignJustify',
      'bulletList', 'numberedList', 'decreaseIndent', 'increaseIndent', 'tabStops', 'insertTable', 'insertPicture', 'insertLink', 'pageBreak', 'pageSetup',
    ];
    // Where changes are tracked all of these are offered too: each makes its change as a tracked one, and the few
    // it cannot make so (a column through merged cells, a link over a picture) are refused when asked for, and said.
    for (let i = 0; i < others.length; i++) this.setDisabled(others[i], !editable);
    // In a text box only its text changes: no picture, link or page break goes into one.
    if (editable && view.inTextBox) for (const key of ['insertPicture', 'insertLink', 'pageBreak'] as ButtonKey[]) this.setDisabled(key, true);
    // Rows and columns are those of the table the caret is in.
    this.setDisabled('rowsColumns', !editable || !view.inTable);
    this.styleSelect.disabled = !editable;
    this.fontSelect.disabled = !editable;
    this.sizeSelect.disabled = !editable;
    if (!editable) this.closePopover();
    this.updateChanges(editable);
    // The comments: also where the protection of the document allows comments and nothing else.
    const commented = view.commentCount > 0;
    this.setDisabled('newComment', !view.canComment);
    this.setDisabled('previousComment', !commented);
    this.setDisabled('nextComment', !commented);

    const pressed = (key: ButtonKey, name: keyof RangeFormat, value: unknown) => this.setPressed(key, mixed.indexOf(name) >= 0 ? 'mixed' : format !== undefined && format[name] === value ? 'true' : 'false');
    pressed('bold', 'bold', true);
    pressed('italic', 'italic', true);
    pressed('underline', 'underline', true);
    pressed('strikethrough', 'strike', true);
    pressed('alignLeft', 'align', 'left');
    pressed('alignCenter', 'align', 'center');
    pressed('alignRight', 'align', 'right');
    pressed('alignJustify', 'align', 'both');
    pressed('bulletList', 'list', 'bullet');
    pressed('numberedList', 'list', 'number');

    this.showStyle(format && mixed.indexOf('paraStyle') < 0 ? format.paraStyle : undefined);
    this.showFont(format && mixed.indexOf('font') < 0 ? format.font : undefined);
    this.showSize(format && mixed.indexOf('size') < 0 && typeof format.size === 'number' ? format.size / 2 : undefined);
    const color = format && typeof format.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(format.color) ? `#${format.color.toUpperCase()}` : '#000000';
    this.paintBar(this.bars.textColor, color);
    let highlight: string | undefined;
    if (format && format.highlight !== undefined) for (const entry of HIGHLIGHTS) if (entry[0] === format.highlight) highlight = entry[1];
    this.paintBar(this.bars.highlight, highlight);
    this.keepTabStop();
  }

  /** The buttons of the tracked changes: what the document holds, what the selection touches, what its protection allows. */
  private updateChanges(editable: boolean): void {
    const view = this.host.view;
    const s = this.strings;
    const count = editable ? view.revisions().length : 0;
    const selected = editable && count > 0 ? view.selectedRevisions() : [];
    // An enforced protection allows no change to be accepted or rejected (Word does not either).
    const locked = view.protection !== undefined;
    const enforced = view.trackingEnforced;
    this.setDisabled('trackChanges', !editable || enforced);
    this.setPressed('trackChanges', view.trackingChanges ? 'true' : 'false');
    this.setTitle('trackChanges', enforced ? `${s.trackChanges} (${s.trackingEnforced})` : s.trackChanges);
    this.setDisabled('previousChange', count === 0);
    this.setDisabled('nextChange', count === 0);
    this.setDisabled('acceptChange', locked || !selected.some((revision) => revision.canAccept));
    this.setDisabled('rejectChange', locked || !selected.some((revision) => revision.canReject));
    this.setDisabled('changesMenu', locked || count === 0);
    // What the buttons would work on: the name of an author is set as text.
    const about = locked ? ` (${s.changesProtected})` : selected.length === 1 ? `: ${view.describeRevision(selected[0])}` : '';
    this.setTitle('acceptChange', s.acceptChange + about);
    this.setTitle('rejectChange', s.rejectChange + about);
    if (this.popover && this.popover.anchor === this.buttons.changesMenu && (locked || count === 0)) this.closePopover();
  }

  private setTitle(key: ButtonKey, title: string): void {
    const button = this.buttons[key];
    if (button && button.title !== title) button.title = title;
  }

  closePopover(): void {
    if (this.popover) this.popover.close('owner');
    this.closeLink(false);
  }

  /** Opens the field for the address of a link; also what Ctrl+K does. */
  openLink(): void {
    const anchor = this.buttons.insertLink;
    if (this.destroyed || !anchor || anchor.disabled) return;
    this.closePopover();
    const s = this.strings;
    const field = document.createElement('div');
    field.className = 'bym-popover bym-doc-linkfield';
    field.setAttribute('role', 'dialog');
    field.setAttribute('aria-label', s.insertLink);
    field.style.cssText = 'position:absolute;z-index:10;left:0;top:0;display:flex;align-items:center;box-sizing:border-box;max-width:calc(100% - 8px)';
    const input = document.createElement('input');
    input.type = 'url';
    input.className = 'bym-doc-linkinput';
    input.setAttribute('aria-label', s.linkAddress);
    input.placeholder = 'https://';
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.spellcheck = false;
    const apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'bym-pop-item bym-doc-linkapply';
    apply.textContent = s.applyLink;
    apply.style.width = 'auto';
    field.appendChild(input);
    field.appendChild(apply);
    const submit = () => {
      const target = input.value;
      this.closeLink(false);
      // The view refuses what is not an address a link may have, and says so.
      this.host.run(() => this.host.view.insertLink(target));
    };
    apply.addEventListener('click', submit);
    field.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        event.stopPropagation();
        submit();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        this.closeLink(true);
      } else if (event.key !== 'Tab') event.stopPropagation();
    });
    const parent = this.host.root;
    parent.appendChild(field);
    // Under the button, inside the editor.
    const box = parent.getBoundingClientRect();
    const rect = anchor.getBoundingClientRect();
    const scaleX = parent.offsetWidth > 0 && box.width > 0 ? box.width / parent.offsetWidth : 1;
    const scaleY = parent.offsetHeight > 0 && box.height > 0 ? box.height / parent.offsetHeight : 1;
    const left = Math.min(Math.max((rect.left - box.left) / scaleX, 4), Math.max(parent.clientWidth - field.offsetWidth - 4, 4));
    field.style.left = `${Math.round(left)}px`;
    field.style.top = `${Math.round((rect.bottom - box.top) / scaleY + 4)}px`;
    anchor.setAttribute('aria-expanded', 'true');
    this.linkField = field;
    this.openedAt = this.element.scrollLeft;
    const watch = (event: Event) => {
      const path = typeof event.composedPath === 'function' ? event.composedPath() : [event.target as EventTarget];
      if (path.indexOf(field) >= 0 || path.indexOf(anchor) >= 0) return;
      this.closeLink(false);
    };
    this.linkWatch = watch;
    document.addEventListener('PointerEvent' in window ? 'pointerdown' : 'mousedown', watch, true);
    input.focus({ preventScroll: true });
  }

  private closeLink(toAnchor: boolean): void {
    const field = this.linkField;
    if (!field) return;
    this.linkField = undefined;
    if (this.linkWatch) {
      document.removeEventListener('pointerdown', this.linkWatch, true);
      document.removeEventListener('mousedown', this.linkWatch, true);
      this.linkWatch = undefined;
    }
    const anchor = this.buttons.insertLink;
    if (anchor) anchor.setAttribute('aria-expanded', 'false');
    const root = field.getRootNode() as Document | ShadowRoot;
    const hadFocus = field.contains(root.activeElement);
    // The focus must not stay on an element that is about to go.
    if (hadFocus && !this.destroyed) {
      if (toAnchor && anchor) anchor.focus({ preventScroll: true });
      else this.host.view.focus();
    }
    if (field.parentNode) field.parentNode.removeChild(field);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.closePopover();
    this.destroyed = true;
    this.unwatchFonts();
    const element = this.element;
    element.removeEventListener('click', this.onClick);
    element.removeEventListener('mousedown', this.onMouseDown);
    element.removeEventListener('keydown', this.onKeyDown);
    element.removeEventListener('focusin', this.onFocusIn);
    element.removeEventListener('scroll', this.onScroll);
    this.styleSelect.removeEventListener('change', this.onStyleChange);
    this.fontSelect.removeEventListener('change', this.onFontChange);
    this.sizeSelect.removeEventListener('change', this.onSizeChange);
    this.fileInput.removeEventListener('change', this.onFileChange);
    if (element.parentNode) element.parentNode.removeChild(element);
  }

  // ------------------------------------------------------------- building

  private addButton(key: ButtonKey, action: () => void, toggles = false): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'bym-tb-btn';
    button.setAttribute('data-command', key);
    button.tabIndex = -1;
    if (toggles) button.setAttribute('aria-pressed', 'false');
    button.appendChild(SHARED_ICONS[key] ? createIcon(SHARED_ICONS[key]) : key in COMMENT_ICONS ? commentIcon(key as CommentKey) : createDocIcon(key as DocIconName));
    this.buttons[key] = button;
    this.actions.set(button, action);
    this.controls.push(button);
    this.element.appendChild(button);
    return button;
  }

  private addMenuButton(key: ButtonKey, popup: 'menu' | 'dialog', open: () => void): void {
    const button = this.addButton(key, open);
    button.setAttribute('aria-haspopup', popup);
    button.setAttribute('aria-expanded', 'false');
    this.openers.add(button);
  }

  private addSelect(className: string): HTMLSelectElement {
    const select = document.createElement('select');
    select.className = `bym-tb-select ${className}`;
    this.element.appendChild(select);
    return select;
  }

  private addSeparator(): void {
    const separator = document.createElement('span');
    separator.className = 'bym-tb-sep';
    separator.setAttribute('role', 'separator');
    separator.setAttribute('aria-orientation', 'vertical');
    this.element.appendChild(separator);
  }

  private barOf(key: ButtonKey): SVGElement | undefined {
    const button = this.buttons[key];
    const bar = button ? button.getElementsByClassName(ICON_BAR_CLASS)[0] : undefined;
    return bar as SVGElement | undefined;
  }

  private applyStrings(): void {
    const s = this.strings;
    this.element.setAttribute('aria-label', s.toolbar);
    const labels: { [key in ButtonKey]: string } = {
      save: withShortcut(s.save, 'S'),
      undo: withShortcut(s.undo, 'Z'),
      redo: withShortcut(s.redo, 'Y'),
      bold: withShortcut(s.bold, 'B'),
      italic: withShortcut(s.italic, 'I'),
      underline: withShortcut(s.underline, 'U'),
      strikethrough: s.strikethrough,
      textColor: s.textColor,
      highlight: s.highlightColor,
      alignLeft: withShortcut(s.alignLeft, 'L'),
      alignCenter: withShortcut(s.alignCenter, 'E'),
      alignRight: withShortcut(s.alignRight, 'R'),
      alignJustify: withShortcut(s.alignJustify, 'J'),
      bulletList: s.bulletList,
      numberedList: s.numberedList,
      decreaseIndent: s.decreaseIndent,
      increaseIndent: s.increaseIndent,
      tabStops: s.tabStops,
      insertTable: s.insertTable,
      rowsColumns: s.tableRowsColumns,
      insertPicture: s.insertPicture,
      insertLink: withShortcut(s.insertLink, 'K'),
      pageBreak: s.pageBreak,
      pageSetup: s.pageSetup,
      trackChanges: s.trackChanges,
      previousChange: s.previousChange,
      nextChange: s.nextChange,
      acceptChange: s.acceptChange,
      rejectChange: s.rejectChange,
      changesMenu: s.allChanges,
      newComment: `${s.newComment} (${apple() ? '\u2325\u2318M' : 'Ctrl+Alt+M'})`,
      previousComment: s.previousComment,
      nextComment: s.nextComment,
      find: withShortcut(s.findReplace, 'F'),
    };
    for (const key of Object.keys(labels) as ButtonKey[]) {
      const button = this.buttons[key];
      if (!button) continue;
      button.setAttribute('aria-label', labels[key]);
      button.title = labels[key];
    }
    const name = (select: HTMLSelectElement, label: string) => {
      select.setAttribute('aria-label', label);
      select.title = label;
    };
    name(this.styleSelect, s.paragraphStyle);
    name(this.fontSelect, s.fontName);
    name(this.sizeSelect, s.fontSize);
  }

  // ---------------------------------------------------------------- state

  private setDisabled(key: ButtonKey, disabled: boolean): void {
    const button = this.buttons[key];
    if (button && button.disabled !== disabled) button.disabled = disabled;
  }

  private setPressed(key: ButtonKey, pressed: 'true' | 'false' | 'mixed'): void {
    const button = this.buttons[key];
    if (button && button.getAttribute('aria-pressed') !== pressed) button.setAttribute('aria-pressed', pressed);
  }

  /** Fills a list when what it holds has changed, and selects a value; a value the list does not have is added, an unknown one shows nothing. */
  private fill(select: HTMLSelectElement, entries: [value: string, label: string][], key: string, old: string, value: string | undefined): string {
    if (key !== old) {
      while (select.firstChild) select.removeChild(select.firstChild);
      // An entry without text stands for "several values" and for "none".
      const blank = option('\u0000', '');
      blank.hidden = true;
      select.appendChild(blank);
      for (let i = 0; i < entries.length; i++) select.appendChild(option(entries[i][0], entries[i][1]));
    }
    const wanted = value === undefined ? '\u0000' : value;
    if (select.value !== wanted) select.value = wanted;
    return key;
  }

  private showStyle(style: string | undefined): void {
    const styles = this.host.view.document.styles;
    const fallback = styles.defaultIds.paragraph;
    const entries: [string, string][] = [];
    let hasDefault = false;
    styles.styles.forEach((entry, id) => {
      if (entry.type !== 'paragraph' || entries.length >= MAX_LIST_ENTRIES) return;
      // Styles Word keeps out of its own list stay out of this one, unless the selection has one of them.
      if ((entry.hidden || entry.semiHidden) && id !== style) return;
      if (id === fallback) hasDefault = true;
      entries.push([id, entry.name !== undefined && entry.name !== '' ? entry.name : id]);
    });
    // A document without a default paragraph style still has "Normal": no style at all.
    if (!hasDefault) entries.unshift(['', this.strings.normalStyle]);
    let shown = style;
    if (shown === '') shown = hasDefault ? fallback : '';
    if (shown !== undefined && shown !== '' && !styles.styles.has(shown)) entries.push([shown, shown]);
    this.styleKey = this.fill(this.styleSelect, entries, JSON.stringify(entries), this.styleKey, shown);
  }

  private showFont(font: string | undefined): void {
    const { names, shown } = fontChoices(this.host.view.document.fonts.keys(), font);
    const entries = names.map((name): [string, string] => [name, name]);
    this.fontKey = this.fill(this.fontSelect, entries, names.join('\u0001'), this.fontKey, shown);
  }

  private showSize(size: number | undefined): void {
    const usable = size !== undefined && size > 0 && size <= 1638;
    const sizes = !usable || FONT_SIZES.indexOf(size as number) >= 0 ? FONT_SIZES : FONT_SIZES.concat([size as number]).sort((a, b) => a - b);
    const entries = sizes.map((value): [string, string] => [String(value), String(value)]);
    this.sizeKey = this.fill(this.sizeSelect, entries, sizes.join(' '), this.sizeKey, usable ? String(size) : undefined);
  }

  /** The strip under a colour icon: the colour, or an empty outline for none. */
  private paintBar(bar: SVGElement | undefined, color: string | undefined): void {
    if (!bar) return;
    bar.style.fill = color || 'none';
    bar.setAttribute('data-color', color || 'none');
  }

  /** Exactly one enabled control is a Tab stop: the one last focused while it stays enabled. */
  private keepTabStop(): void {
    let stop = this.stop && !this.stop.disabled ? this.stop : undefined;
    if (!stop) {
      for (let i = 0; i < this.controls.length; i++) {
        if (!this.controls[i].disabled) {
          stop = this.controls[i];
          break;
        }
      }
    }
    this.stop = stop;
    for (let i = 0; i < this.controls.length; i++) {
      const tabIndex = this.controls[i] === stop ? 0 : -1;
      if (this.controls[i].tabIndex !== tabIndex) this.controls[i].tabIndex = tabIndex;
    }
  }

  // ------------------------------------------------------------- commands

  private openColors(): void {
    const view = this.host.view;
    const format = view.selectionFormat;
    const current = format && typeof format.color === 'string' ? format.color.toUpperCase() : '';
    const swatches: PopoverItem[] = [];
    for (let i = 0; i < PALETTE.length; i++) {
      const hex = `#${PALETTE[i]}`;
      const name = this.strings.colorName(Math.floor(i / PALETTE_COLUMNS), i % PALETTE_COLUMNS);
      swatches.push({ label: `${name} (${hex})`, swatch: hex, value: hex, selected: current === PALETTE[i], action: () => this.host.run(() => view.format({ color: PALETTE[i] })) });
    }
    const reset: PopoverItem = { label: this.strings.automatic, action: () => this.host.run(() => view.format({ color: 'auto' })) };
    this.open('textColor', this.strings.textColor, [reset], swatches, PALETTE_COLUMNS);
  }

  private openHighlights(): void {
    const view = this.host.view;
    const format = view.selectionFormat;
    const current = format ? format.highlight : undefined;
    const swatches = HIGHLIGHTS.map(
      (entry): PopoverItem => ({ label: entry[1], swatch: entry[1], value: entry[1], selected: current === entry[0], action: () => this.host.run(() => view.format({ highlight: entry[0] })) }),
    );
    const reset: PopoverItem = { label: this.strings.noHighlight, action: () => this.host.run(() => view.format({ highlight: 'none' })) };
    this.open('highlight', this.strings.highlightColor, [reset], swatches, 5);
  }

  /** A grid of cells: the one chosen says how many rows and columns the new table has. */
  private openTableGrid(): void {
    const view = this.host.view;
    const cells: PopoverItem[] = [];
    for (let row = 1; row <= TABLE_GRID_ROWS; row++) {
      for (let column = 1; column <= TABLE_GRID_COLUMNS; column++) {
        cells.push({ label: this.strings.tableSize(row, column), swatch: '#FFFFFF', value: `${row}x${column}`, action: () => this.host.run(() => view.insertTable(row, column)) });
      }
    }
    this.open('insertTable', this.strings.insertTable, [], cells, TABLE_GRID_COLUMNS);
  }

  private openRowsColumns(): void {
    const s = this.strings;
    const view = this.host.view;
    const items: PopoverItem[] = [];
    const add = (label: string, action: () => void) => items.push({ label, action: () => this.host.run(action) });
    add(s.insertRowAbove, () => view.insertRow('above'));
    add(s.insertRowBelow, () => view.insertRow('below'));
    add(s.insertColumnLeft, () => view.insertColumn('left'));
    add(s.insertColumnRight, () => view.insertColumn('right'));
    add(s.deleteRow, () => view.deleteRow());
    add(s.deleteColumn, () => view.deleteColumn());
    this.open('rowsColumns', s.tableRowsColumns, items);
  }

  private openChangesMenu(): void {
    const s = this.strings;
    const view = this.host.view;
    const items: PopoverItem[] = [
      { label: s.acceptAllChanges, action: () => this.host.run(() => view.acceptAllRevisions()) },
      { label: s.rejectAllChanges, action: () => this.host.run(() => view.rejectAllRevisions()) },
    ];
    this.open('changesMenu', s.allChanges, items);
  }

  private open(key: ButtonKey, label: string, items: PopoverItem[], swatches?: PopoverItem[], columns?: number): void {
    const anchor = this.buttons[key];
    if (!anchor) return;
    this.closePopover();
    anchor.setAttribute('aria-expanded', 'true');
    const onClose = (reason: PopoverCloseReason) => {
      const popover = this.popover;
      this.popover = undefined;
      anchor.setAttribute('aria-expanded', 'false');
      if (this.destroyed || !popover || !popover.hasFocus) return;
      // Escape and Tab hand the focus back to the button; a choice gets it to the document through the command.
      if (reason === 'escape' || reason === 'tab') anchor.focus({ preventScroll: true });
      else this.host.view.focus();
    };
    this.openedAt = this.element.scrollLeft;
    this.popover = new Popover({ parent: this.host.root, anchor, label, items, swatches, columns, onClose });
  }

  // --------------------------------------------------------------- events

  private onClick = (event: Event): void => {
    const target = event.target as Element | null;
    const button = target && typeof target.closest === 'function' ? (target.closest('button') as HTMLButtonElement | null) : null;
    if (!button || button.disabled) return;
    const action = this.actions.get(button);
    if (!action) return;
    if (this.openers.has(button)) {
      // A second press on the button closes what the first one opened.
      if ((this.popover && this.popover.anchor === button) || (this.linkField && button === this.buttons.insertLink)) this.closePopover();
      else action();
      return;
    }
    this.closePopover();
    this.host.run(action, this.keepsFocus.has(button));
  };

  /** A press on the toolbar leaves the focus where it is (in the document); only the lists need it to open. */
  private onMouseDown = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (target && typeof target.closest === 'function' && target.closest('select')) {
      this.listByKey = false;
      return;
    }
    event.preventDefault();
  };

  private onFocusIn = (event: FocusEvent): void => {
    const index = this.controls.indexOf(event.target as Control);
    if (index < 0) return;
    this.stop = this.controls[index];
    this.keepTabStop();
  };

  private onScroll = (): void => {
    if (this.element.scrollLeft !== this.openedAt) this.closePopover();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    if (event.defaultPrevented) return;
    const target = event.target as Element | null;
    if (target === this.styleSelect || target === this.fontSelect || target === this.sizeSelect) {
      // Every key is the list's own; Enter and Escape are how the keyboard says it is done there.
      if (event.key === 'Enter' || event.key === 'Escape') {
        event.preventDefault();
        this.listByKey = false;
        this.host.view.focus();
      } else if (event.key !== 'Tab') this.listByKey = true;
      return;
    }
    if (event.altKey || event.ctrlKey || event.metaKey) return;
    const index = this.controls.indexOf(event.target as Control);
    if (index < 0) return;
    let next: Control | undefined;
    switch (event.key) {
      case 'ArrowRight':
        next = this.enabledFrom(index + 1, 1);
        break;
      case 'ArrowLeft':
        next = this.enabledFrom(index - 1, -1);
        break;
      case 'Home':
        next = this.enabledFrom(0, 1);
        break;
      case 'End':
        next = this.enabledFrom(this.controls.length - 1, -1);
        break;
      case 'Escape':
        this.host.view.focus();
        break;
      default:
        return;
    }
    event.preventDefault();
    if (next) next.focus({ preventScroll: false });
  };

  private enabledFrom(start: number, step: number): Control | undefined {
    for (let i = start; i >= 0 && i < this.controls.length; i += step) if (!this.controls[i].disabled) return this.controls[i];
    return undefined;
  }

  private onStyleChange = (): void => {
    const value = this.styleSelect.value;
    if (value === '\u0000') return;
    // No style at all is the default one.
    this.host.run(() => this.host.view.formatParagraph({ style: value === '' ? null : value }), this.listByKey);
  };

  private onFontChange = (): void => {
    const value = this.fontSelect.value;
    if (value === '\u0000' || value === '') return;
    this.host.run(() => this.host.view.format({ font: value }), this.listByKey);
  };

  private onSizeChange = (): void => {
    const size = Number(this.sizeSelect.value);
    if (!(size > 0)) return;
    // Sizes are kept in half points.
    this.host.run(() => this.host.view.format({ size: Math.round(size * 2) }), this.listByKey);
  };

  private onFileChange = (): void => {
    const input = this.fileInput;
    const file = input.files && input.files[0];
    // The same file can be chosen again next time.
    input.value = '';
    if (!file || this.destroyed) return;
    if (file.size > MAX_PICTURE_FILE) {
      this.host.run(() => this.host.view.insertPicture(new Uint8Array(0)));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (this.destroyed || !(reader.result instanceof ArrayBuffer)) return;
      const bytes = new Uint8Array(reader.result);
      // What kind of picture it is, the view reads from the bytes. The name of the file stays out of the document.
      this.host.run(() => this.host.view.insertPicture(bytes));
    };
    reader.readAsArrayBuffer(file);
  };
}
