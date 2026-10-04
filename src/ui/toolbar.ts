// The editor's toolbar: one row of buttons and two lists that format the
// selection, and show what the active cell has. Its buttons are a single Tab
// stop with the arrow keys moving between them; each list is a Tab stop of
// its own, because a list needs the arrow keys for itself. A command ends
// with the focus back in the grid, so work with the keyboard is not
// interrupted; a list that is being walked with the keyboard keeps the focus.

import { resolveColor, type ResolvedCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, type Alignment } from '../model/sheet/types';
import type { BorderPreset } from '../ops/format';
import { EDITOR_STRINGS_EN, EDITOR_STRINGS_TR, withShortcut, type EditorStrings } from './editor-strings';
import { ICON_BAR_CLASS, createBorderIcon, createIcon, type IconName } from './icons';
import { Popover, type PopoverCloseReason, type PopoverItem } from './popover';
import type { WorkbookView } from './sheet-view';

/** What the parts of the shell need from the editor that holds them. */
export interface ShellHost {
  readonly view: WorkbookView;
  /** The editor's root element: positioned, and the parent of everything that drops down. */
  readonly root: HTMLElement;
  /**
   * Runs a command of the view. A failure is shown in the status bar instead
   * of thrown; either way the focus goes back to the grid (unless it is to be
   * kept where it is) and the shell is brought up to date.
   */
  run(action: () => void, keepFocus?: boolean): void;
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
const DEFAULT_FONT_SIZE = 11;

/** Ten columns: greys, the plain colours, then two lighter and two darker shades of each. */
const PALETTE_COLUMNS = 10;
const PALETTE = (
  '000000 434343 666666 999999 B7B7B7 CCCCCC D9D9D9 EFEFEF F3F3F3 FFFFFF ' +
  '980000 FF0000 FF9900 FFFF00 00FF00 00FFFF 4A86E8 0000FF 9900FF FF00FF ' +
  'E6B8AF F4CCCC FCE5CD FFF2CC D9EAD3 D0E0E3 C9DAF8 CFE2F3 D9D2E9 EAD1DC ' +
  'DD7E6B EA9999 F9CB9C FFE599 B6D7A8 A2C4C9 A4C2F4 9FC5E8 B4A7D6 D5A6BD ' +
  'A61C00 CC0000 E69138 F1C232 6AA84F 45818E 3C78D8 3D85C6 674EA7 A64D79 ' +
  '5B0F00 660000 783F04 7F6000 274E13 0C343D 1C4587 073763 20124D 4C1130'
).split(' ');

const BORDER_PRESETS: BorderPreset[] = ['all', 'outline', 'bottom', 'top', 'left', 'right', 'none'];

type FormatKind = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'time' | 'text';
const FORMAT_KINDS: FormatKind[] = ['general', 'number', 'currency', 'percent', 'date', 'time', 'text'];
const CUSTOM = 'custom';

function formatCode(kind: FormatKind, strings: EditorStrings): string {
  switch (kind) {
    case 'general':
      return 'General';
    case 'number':
      return '#,##0.00';
    case 'currency':
      return strings.currencyCode;
    case 'percent':
      return '0%';
    case 'date':
      return strings.dateCode;
    case 'time':
      return 'hh:mm';
    default:
      return '@';
  }
}

/** The list entry a number format belongs to; a code applied in the other language still counts. */
function formatKindOf(code: string, strings: EditorStrings): FormatKind | typeof CUSTOM {
  for (let i = 0; i < FORMAT_KINDS.length; i++) if (formatCode(FORMAT_KINDS[i], strings) === code) return FORMAT_KINDS[i];
  if (code === EDITOR_STRINGS_EN.currencyCode || code === EDITOR_STRINGS_TR.currencyCode) return 'currency';
  if (code === EDITOR_STRINGS_EN.dateCode || code === EDITOR_STRINGS_TR.dateCode) return 'date';
  return CUSTOM;
}

function fontSizeOf(format: ResolvedCellStyle): number {
  const size = format.font.size;
  // From the file: anything that is not a usable size reads as the default.
  return typeof size === 'number' && isFinite(size) && size > 0 && size <= 409 ? Math.round(size * 100) / 100 : DEFAULT_FONT_SIZE;
}

/** Turns a font property of the selection on, or off when the active cell has it. */
function toggleFont(view: WorkbookView, what: 'bold' | 'italic' | 'underline' | 'strike'): void {
  const font = view.selectionFormat.font;
  if (what === 'underline') view.format({ underline: font.underline !== undefined && font.underline !== 'none' ? 'none' : 'single' });
  else if (what === 'bold') view.format({ bold: !font.bold });
  else if (what === 'italic') view.format({ italic: !font.italic });
  else view.format({ strike: !font.strike });
}

type Control = HTMLButtonElement;
type ButtonKey = IconName;

export class Toolbar {
  readonly element: HTMLDivElement;
  private readonly host: ShellHost;
  private strings: EditorStrings;
  /** Every button in visual order, for the arrow keys. */
  private readonly controls: Control[] = [];
  private readonly buttons: { [key in ButtonKey]?: HTMLButtonElement } = {};
  private readonly sizeSelect: HTMLSelectElement;
  private readonly formatSelect: HTMLSelectElement;
  private readonly bars: { textColor?: SVGElement; fillColor?: SVGElement } = {};
  /** Font sizes the size list holds now; it is rebuilt only when this changes. */
  private sizeKey = '';
  private customShown = false;
  private popover: Popover | undefined;
  /** Where the toolbar was scrolled to when the popover opened; it closes when the toolbar moves on. */
  private openedAt = 0;
  /** The one control that is a Tab stop. */
  private stop: Control | undefined;
  /** The list in use is being worked with the keyboard, not with a pointer. */
  private listByKey = false;
  private destroyed = false;

  constructor(host: ShellHost, strings: EditorStrings) {
    this.host = host;
    this.strings = strings;
    const element = document.createElement('div');
    element.className = 'bym-toolbar';
    element.setAttribute('role', 'toolbar');
    // Some browsers make an element that scrolls a Tab stop of its own; the stop is one of the controls.
    element.tabIndex = -1;
    this.element = element;

    this.addButton('undo', () => host.view.undo());
    this.addButton('redo', () => host.view.redo());
    // Works in a read-only view as well; what goes wrong is said by the view.
    this.addButton('print', () => void host.view.print().catch(() => undefined)).setAttribute('aria-busy', 'false');
    this.addSeparator();
    this.sizeSelect = this.addSelect('bym-tb-size');
    this.addSeparator();
    this.addButton('bold', () => toggleFont(host.view, 'bold'), true);
    this.addButton('italic', () => toggleFont(host.view, 'italic'), true);
    this.addButton('underline', () => toggleFont(host.view, 'underline'), true);
    this.addButton('strikethrough', () => toggleFont(host.view, 'strike'), true);
    this.addSeparator();
    this.addMenuButton('textColor', 'dialog', () => this.openPalette('textColor'));
    this.addMenuButton('fillColor', 'dialog', () => this.openPalette('fillColor'));
    this.addSeparator();
    this.addMenuButton('borders', 'menu', () => this.openBorders());
    this.addSeparator();
    this.addButton('alignLeft', () => this.align('left'), true);
    this.addButton('alignCenter', () => this.align('center'), true);
    this.addButton('alignRight', () => this.align('right'), true);
    this.addSeparator();
    this.addButton('wrapText', () => host.view.format({ wrapText: host.view.selectionFormat.alignment.wrapText !== true }), true);
    this.addButton('merge', () => (host.view.selectionMerged ? host.view.unmerge() : host.view.merge()), true);
    this.addSeparator();
    this.formatSelect = this.addSelect('bym-tb-format');
    this.addSeparator();
    this.addButton('sortAscending', () => host.view.sort(false));
    this.addButton('sortDescending', () => host.view.sort(true));
    this.addButton('filter', () => host.view.toggleFilter(), true);
    this.addSeparator();
    this.addMenuButton('rowsColumns', 'menu', () => this.openRowsColumns());

    this.bars.textColor = this.barOf('textColor');
    this.bars.fillColor = this.barOf('fillColor');

    element.addEventListener('click', this.onClick);
    element.addEventListener('mousedown', this.onMouseDown);
    element.addEventListener('keydown', this.onKeyDown);
    element.addEventListener('focusin', this.onFocusIn);
    element.addEventListener('scroll', this.onScroll, { passive: true });
    this.sizeSelect.addEventListener('change', this.onSizeChange);
    this.formatSelect.addEventListener('change', this.onFormatChange);

    this.applyStrings();
  }

  setStrings(strings: EditorStrings): void {
    this.strings = strings;
    this.closePopover();
    this.applyStrings();
    this.update();
  }

  /** Brings every control in line with the active cell, the history and the read-only state. */
  update(): void {
    if (this.destroyed) return;
    const view = this.host.view;
    const sheet = view.workbook.sheets[view.activeSheet];
    const editable = !view.readOnly && !!sheet && sheet.kind === 'worksheet';
    const format = view.selectionFormat;
    const font = format.font;
    const horizontal = format.alignment.horizontal;

    // Not during an edit either: the view would answer by dropping what was typed.
    this.setDisabled('undo', view.readOnly || view.isEditing || !view.history.canUndo);
    this.setDisabled('redo', view.readOnly || view.isEditing || !view.history.canRedo);
    const others: ButtonKey[] = ['bold', 'italic', 'underline', 'strikethrough', 'textColor', 'fillColor', 'borders', 'alignLeft', 'alignCenter', 'alignRight', 'wrapText', 'merge', 'sortAscending', 'sortDescending', 'filter', 'rowsColumns'];
    for (let i = 0; i < others.length; i++) this.setDisabled(others[i], !editable);
    this.sizeSelect.disabled = !editable;
    this.formatSelect.disabled = !editable;
    if (!editable) this.closePopover();

    this.setPressed('bold', font.bold === true);
    this.setPressed('italic', font.italic === true);
    this.setPressed('underline', font.underline !== undefined && font.underline !== 'none');
    this.setPressed('strikethrough', font.strike === true);
    this.setPressed('alignLeft', horizontal === 'left');
    this.setPressed('alignCenter', horizontal === 'center');
    this.setPressed('alignRight', horizontal === 'right');
    this.setPressed('wrapText', format.alignment.wrapText === true);
    this.setPressed('merge', view.selectionMerged);
    this.setPressed('filter', view.hasFilter);

    this.showFontSize(fontSizeOf(format));
    this.showFormat(formatKindOf(format.numFmt, this.strings));
    this.paintBar(this.bars.textColor, this.textColorOf(format) || '#000000');
    this.paintBar(this.bars.fillColor, this.fillColorOf(format));
    this.keepTabStop();
  }

  /** Marks the Print button as busy while the pages of a print job are prepared. */
  setPrinting(busy: boolean): void {
    const button = this.buttons.print;
    const value = busy ? 'true' : 'false';
    if (button && button.getAttribute('aria-busy') !== value) button.setAttribute('aria-busy', value);
  }

  closePopover(): void {
    if (this.popover) this.popover.close('owner');
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.closePopover();
    const element = this.element;
    element.removeEventListener('click', this.onClick);
    element.removeEventListener('mousedown', this.onMouseDown);
    element.removeEventListener('keydown', this.onKeyDown);
    element.removeEventListener('focusin', this.onFocusIn);
    element.removeEventListener('scroll', this.onScroll);
    this.sizeSelect.removeEventListener('change', this.onSizeChange);
    this.formatSelect.removeEventListener('change', this.onFormatChange);
    if (element.parentNode) element.parentNode.removeChild(element);
  }

  // ------------------------------------------------------------- building

  private readonly actions = new Map<HTMLButtonElement, () => void>();
  /** Buttons whose click opens something: their action runs as it is, not as a command of the view. */
  private readonly openers = new Set<HTMLButtonElement>();

  private addButton(key: ButtonKey, action: () => void, toggles = false): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'bym-tb-btn';
    button.setAttribute('data-command', key);
    button.tabIndex = -1;
    if (toggles) button.setAttribute('aria-pressed', 'false');
    button.appendChild(createIcon(key));
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
      undo: withShortcut(s.undo, 'Z'),
      redo: withShortcut(s.redo, 'Y'),
      // The same label as the button of the document bar.
      print: s.print,
      bold: withShortcut(s.bold, 'B'),
      italic: withShortcut(s.italic, 'I'),
      underline: withShortcut(s.underline, 'U'),
      strikethrough: s.strikethrough,
      textColor: s.textColor,
      fillColor: s.fillColor,
      borders: s.borders,
      alignLeft: s.alignLeft,
      alignCenter: s.alignCenter,
      alignRight: s.alignRight,
      wrapText: s.wrapText,
      merge: s.mergeCells,
      sortAscending: s.sortAscending,
      sortDescending: s.sortDescending,
      filter: s.filter,
      rowsColumns: s.rowsAndColumns,
    };
    for (const key of Object.keys(labels) as ButtonKey[]) {
      const button = this.buttons[key];
      if (!button) continue;
      button.setAttribute('aria-label', labels[key]);
      button.title = labels[key];
    }
    this.sizeSelect.setAttribute('aria-label', s.fontSize);
    this.sizeSelect.title = s.fontSize;
    this.formatSelect.setAttribute('aria-label', s.numberFormat);
    this.formatSelect.title = s.numberFormat;

    const names: { [kind in FormatKind]: string } = {
      general: s.formatGeneral,
      number: s.formatNumber,
      currency: s.formatCurrency,
      percent: s.formatPercent,
      date: s.formatDate,
      time: s.formatTime,
      text: s.formatText,
    };
    const select = this.formatSelect;
    const value = select.value;
    while (select.firstChild) select.removeChild(select.firstChild);
    for (let i = 0; i < FORMAT_KINDS.length; i++) select.appendChild(option(FORMAT_KINDS[i], names[FORMAT_KINDS[i]]));
    this.customShown = false;
    if (value) this.showFormat(value as FormatKind | typeof CUSTOM);
  }

  // ---------------------------------------------------------------- state

  private setDisabled(key: ButtonKey, disabled: boolean): void {
    const button = this.buttons[key];
    if (button && button.disabled !== disabled) button.disabled = disabled;
  }

  private setPressed(key: ButtonKey, pressed: boolean): void {
    const button = this.buttons[key];
    if (button) button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
  }

  /** Selects a size in the list; a size the list does not have is added in its place. */
  private showFontSize(size: number): void {
    const sizes = FONT_SIZES.indexOf(size) >= 0 ? FONT_SIZES : FONT_SIZES.concat([size]).sort((a, b) => a - b);
    const key = sizes.join(' ');
    const select = this.sizeSelect;
    if (key !== this.sizeKey) {
      this.sizeKey = key;
      while (select.firstChild) select.removeChild(select.firstChild);
      for (let i = 0; i < sizes.length; i++) select.appendChild(option(String(sizes[i]), String(sizes[i])));
    }
    if (select.value !== String(size)) select.value = String(size);
  }

  /** Selects a number format in the list; "Custom" is only an entry while the cell has such a format. */
  private showFormat(kind: FormatKind | typeof CUSTOM): void {
    const select = this.formatSelect;
    const custom = kind === CUSTOM;
    if (custom !== this.customShown) {
      this.customShown = custom;
      if (custom) select.appendChild(option(CUSTOM, this.strings.formatCustom));
      else if (select.lastChild) select.removeChild(select.lastChild);
    }
    if (select.value !== kind) select.value = kind;
  }

  private textColorOf(format: ResolvedCellStyle): string | undefined {
    return resolveColor(format.font.color, this.host.view.workbook.styles);
  }

  private fillColorOf(format: ResolvedCellStyle): string | undefined {
    const fill = format.fill;
    if (fill.type !== 'pattern' || fill.pattern !== 'solid') return undefined;
    return resolveColor(fill.fg, this.host.view.workbook.styles);
  }

  /** The strip under a colour icon: the colour, or an empty outline for "no fill". */
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

  /** The three alignments work like radio buttons; pressing the one that is on goes back to "general". */
  private align(horizontal: NonNullable<Alignment['horizontal']>): void {
    const view = this.host.view;
    view.format({ horizontal: view.selectionFormat.alignment.horizontal === horizontal ? null : horizontal });
  }

  private openPalette(key: 'textColor' | 'fillColor'): void {
    const view = this.host.view;
    const text = key === 'textColor';
    const format = view.selectionFormat;
    const current = text ? this.textColorOf(format) : this.fillColorOf(format);
    const apply = (rgb: string | null) => () =>
      this.host.run(() => {
        const color = rgb === null ? null : { rgb: `FF${rgb}` };
        view.format(text ? { color } : { fill: color });
      });
    const swatches: PopoverItem[] = [];
    for (let i = 0; i < PALETTE.length; i++) {
      const hex = `#${PALETTE[i]}`;
      const name = this.strings.colorName(Math.floor(i / PALETTE_COLUMNS), i % PALETTE_COLUMNS);
      swatches.push({ label: `${name} (${hex})`, swatch: hex, value: hex, selected: current === hex, action: apply(PALETTE[i]) });
    }
    const reset: PopoverItem = { label: text ? this.strings.automatic : this.strings.noFill, action: apply(null) };
    this.open(key, text ? this.strings.textColor : this.strings.fillColor, [reset], swatches);
  }

  private openBorders(): void {
    const s = this.strings;
    const labels: { [preset in BorderPreset]: string } = {
      all: s.borderAll,
      outline: s.borderOutline,
      bottom: s.borderBottom,
      top: s.borderTop,
      left: s.borderLeft,
      right: s.borderRight,
      none: s.borderNone,
    };
    const items = BORDER_PRESETS.map(
      (preset): PopoverItem => ({
        label: labels[preset],
        icon: createBorderIcon(preset),
        action: () => this.host.run(() => this.host.view.format({ border: { preset } })),
      }),
    );
    this.open('borders', s.borders, items);
  }

  private openRowsColumns(): void {
    const s = this.strings;
    const view = this.host.view;
    const range = view.selection.range;
    const rows = range.bottom - range.top + 1;
    const columns = range.right - range.left + 1;
    const items: PopoverItem[] = [];
    const add = (label: string, action: () => void) => items.push({ label, action: () => this.host.run(action) });
    // Whole columns are selected: there are no rows to speak of, and the other way round.
    const withRows = rows < MAX_ROWS;
    const withColumns = columns < MAX_COLS;
    if (withRows) {
      add(s.insertRowsAbove(rows), () => view.insertRows(range.top, rows));
      add(s.insertRowsBelow(rows), () => view.insertRows(range.bottom + 1, rows));
    }
    if (withColumns) {
      add(s.insertColumnsLeft(columns), () => view.insertColumns(range.left, columns));
      add(s.insertColumnsRight(columns), () => view.insertColumns(range.right + 1, columns));
    }
    if (withRows) add(s.deleteRows(rows), () => view.deleteRows(range.top, rows));
    if (withColumns) add(s.deleteColumns(columns), () => view.deleteColumns(range.left, columns));
    if (items.length > 0) this.open('rowsColumns', s.rowsAndColumns, items);
  }

  private open(key: ButtonKey, label: string, items: PopoverItem[], swatches?: PopoverItem[]): void {
    const anchor = this.buttons[key];
    if (!anchor) return;
    this.closePopover();
    anchor.setAttribute('aria-expanded', 'true');
    const onClose = (reason: PopoverCloseReason) => {
      const popover = this.popover;
      this.popover = undefined;
      anchor.setAttribute('aria-expanded', 'false');
      if (this.destroyed || !popover || !popover.hasFocus) return;
      // The focus must not stay on an element that is about to go. Escape and Tab hand it back to
      // the button; a choice gets it to the grid through the command, and so does everything else.
      if (reason === 'escape' || reason === 'tab') anchor.focus({ preventScroll: true });
      else this.host.view.focus();
    };
    this.openedAt = this.element.scrollLeft;
    this.popover = new Popover({ parent: this.host.root, anchor, label, items, swatches, columns: PALETTE_COLUMNS, onClose });
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
      if (this.popover && this.popover.anchor === button) this.closePopover();
      else action();
      return;
    }
    this.closePopover();
    this.host.run(action);
  };

  /** A press on the toolbar leaves the focus where it is (in the grid); only the lists need it to open. */
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
    // The event comes a frame late: a scroll that brought the button into view before the press that
    // opened the popover is not a reason to close it.
    if (this.element.scrollLeft !== this.openedAt) this.closePopover();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    if (event.defaultPrevented) return;
    const target = event.target as Element | null;
    if (target === this.sizeSelect || target === this.formatSelect) {
      // Every key is the list's own (arrows, Home, End and letters choose in it); Enter and Escape
      // are how the keyboard says it is done there.
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

  private onSizeChange = (): void => {
    const size = Number(this.sizeSelect.value);
    if (!(size > 0)) return;
    // A value chosen with the keyboard may be one step on the way to another: the focus stays in the list.
    this.host.run(() => this.host.view.format({ fontSize: size }), this.listByKey);
  };

  private onFormatChange = (): void => {
    const kind = this.formatSelect.value;
    if (FORMAT_KINDS.indexOf(kind as FormatKind) < 0) return;
    this.host.run(() => this.host.view.format({ numFmt: formatCode(kind as FormatKind, this.strings) }), this.listByKey);
  };
}

function option(value: string, label: string): HTMLOptionElement {
  const element = document.createElement('option');
  element.value = value;
  element.textContent = label;
  return element;
}
