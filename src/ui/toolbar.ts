// The editor's toolbar: one row of buttons and three lists that format the
// selection, and show what the active cell has. Its buttons are a single Tab
// stop with the arrow keys moving between them; each list is a Tab stop of
// its own, because a list needs the arrow keys for itself. A command ends
// with the focus back in the grid, so work with the keyboard is not
// interrupted; a list that is being walked with the keyboard keeps the focus.

import { resolveColor, type ResolvedCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, type Alignment } from '../model/sheet/types';
import type { BorderPreset } from '../ops/format';
import { EDITOR_STRINGS_EN, EDITOR_STRINGS_TR, withCtrlKey, withShortcut, type EditorStrings } from './editor-strings';
import { ICON_BAR_CLASS, createBorderIcon, createIcon, type IconName } from './icons';
import { Popover, type PopoverCloseReason, type PopoverItem } from './popover';
import { fontChoices } from './font-list';
import type { SaveCommand } from './save-command';
import { onFontsChanged } from '../render/fonts';
import type { WorkbookView } from './sheet-view';
import { MAX_PICTURE_BYTES } from '../ops/pictures';
import type { ShapeInfo, ShapeKind } from '../ops/shapes';

/** The weights a line of a shape is given from the toolbar, in points. */
const SHAPE_LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

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
  /** With it the toolbar starts with a Save button, which is not there while the view is read-only. */
  readonly save?: SaveCommand;
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
const DEFAULT_FONT_SIZE = 11;

/** Ten columns: greys, the plain colours, then two lighter and two darker shades of each. */
export const PALETTE_COLUMNS = 10;
export const PALETTE = (
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

/** How a cell is aligned vertically as far as the buttons go: what Excel draws as the top, the middle, the bottom. */
function verticalOf(vertical: Alignment['vertical']): 'top' | 'center' | 'bottom' {
  return vertical === 'top' || vertical === 'justify' ? 'top' : vertical === 'center' || vertical === 'distributed' ? 'center' : 'bottom';
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
  private readonly fontSelect: HTMLSelectElement;
  private readonly sizeSelect: HTMLSelectElement;
  private readonly formatSelect: HTMLSelectElement;
  private readonly unwatchFonts: () => void;
  /** Font names the font list holds now; it is rebuilt only when this changes. */
  private fontKey = '';
  private readonly bars: { textColor?: SVGElement; fillColor?: SVGElement } = {};
  /** The line after the Save button: it goes when the button does. */
  private saveSeparator: HTMLElement | undefined;
  /** Font sizes the size list holds now; it is rebuilt only when this changes. */
  private sizeKey = '';
  private customShown = false;
  private popover: Popover | undefined;
  /** The label of each button, which is its title too unless the button is off for a reason it then shows. */
  private labels: { [key in ButtonKey]?: string } = {};
  /** Where the toolbar was scrolled to when the popover opened; it closes when the toolbar moves on. */
  private openedAt = 0;
  /** The one control that is a Tab stop. */
  private stop: Control | undefined;
  /** The list in use is being worked with the keyboard, not with a pointer. */
  private listByKey = false;
  private destroyed = false;
  /** The browser's own file dialog the picture button opens; never shown itself. */
  private readonly fileInput: HTMLInputElement;

  constructor(host: ShellHost, strings: EditorStrings) {
    this.host = host;
    this.strings = strings;
    const element = document.createElement('div');
    element.className = 'bym-toolbar';
    element.setAttribute('role', 'toolbar');
    // Some browsers make an element that scrolls a Tab stop of its own; the stop is one of the controls.
    element.tabIndex = -1;
    this.element = element;

    const save = host.save;
    if (save) {
      this.addButton('save', () => save.run()).setAttribute('aria-busy', 'false');
      this.saveSeparator = this.addSeparator();
    }
    this.addButton('undo', () => host.view.undo());
    this.addButton('redo', () => host.view.redo());
    // Works in a read-only view as well; what goes wrong is said by the view.
    this.addButton('print', () => void host.view.print().catch(() => undefined)).setAttribute('aria-busy', 'false');
    // The print preview covers the editor and takes the focus; it gives it back when it closes.
    this.addDialogButton('printPreview', () => void host.view.printPreview());
    // One file at a time: a press while one is being made does not ask for a second.
    this.addButton('pdf', () => {
      if (!host.view.exportingPdf) void host.view.downloadPdf().catch(() => undefined);
    }).setAttribute('aria-busy', 'false');
    // The dialog of the page setup of the shown sheet takes the focus, and hands it back when it closes.
    this.addDialogButton('pageSetup', () => void host.view.openPageSetup());
    this.addSeparator();
    this.fontSelect = this.addSelect('bym-tb-font');
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
    this.addButton('alignTop', () => this.alignVertically('top'), true);
    this.addButton('alignMiddle', () => this.alignVertically('center'), true);
    this.addButton('alignBottom', () => this.alignVertically('bottom'), true);
    this.addSeparator();
    this.addButton('wrapText', () => host.view.format({ wrapText: host.view.selectionFormat.alignment.wrapText !== true }), true);
    // Values the merge would discard are asked about first.
    this.addButton('merge', () => (host.view.selectionMerged ? host.view.unmerge() : void host.view.requestMerge()), true);
    this.addMenuButton('mergeOptions', 'menu', () => this.openMerge());
    this.addSeparator();
    this.formatSelect = this.addSelect('bym-tb-format');
    this.addSeparator();
    this.addButton('sortAscending', () => host.view.sort(false));
    this.addButton('sortDescending', () => host.view.sort(true));
    this.addButton('filter', () => host.view.toggleFilter(), true);
    this.addMenuButton('conditionalFormat', 'menu', () => this.openConditional());
    this.addDialogButton('validation', () => void host.view.openValidationDialog());
    this.addMenuButton('chart', 'menu', () => this.openCharts());
    // The picture button only opens the browser's own file dialog: it has nothing that stays open.
    this.addDialogButton('insertPicture', () => this.fileInput.click());
    // The kinds of shape that are added, and what changes the shape that is selected.
    this.addMenuButton('insertShape', 'menu', () => this.openShapes());
    this.addSeparator();
    this.addMenuButton('rowsColumns', 'menu', () => this.openRowsColumns());
    this.addSeparator();
    // The bar of Find and Replace takes the focus itself; finding works in a read-only view too.
    this.addDialogButton('find', () => void host.view.openFind(false));

    this.bars.textColor = this.barOf('textColor');
    this.bars.fillColor = this.barOf('fillColor');

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = 'image/png,image/jpeg,image/gif,image/bmp';
    fileInput.className = 'bym-tb-file';
    fileInput.tabIndex = -1;
    fileInput.setAttribute('aria-hidden', 'true');
    fileInput.style.display = 'none';
    fileInput.addEventListener('change', this.onFileChange);
    element.appendChild(fileInput);
    this.fileInput = fileInput;

    element.addEventListener('click', this.onClick);
    element.addEventListener('mousedown', this.onMouseDown);
    element.addEventListener('keydown', this.onKeyDown);
    element.addEventListener('focusin', this.onFocusIn);
    element.addEventListener('scroll', this.onScroll, { passive: true });
    this.fontSelect.addEventListener('change', this.onFontChange);
    this.sizeSelect.addEventListener('change', this.onSizeChange);
    this.formatSelect.addEventListener('change', this.onFormatChange);
    // The list of fonts follows the fonts that are registered.
    this.unwatchFonts = onFontsChanged(() => this.update());

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

    // Not during an edit either: the view would answer by dropping what was typed. Nor while a note
    // or the name of a sheet is typed in its field, whose own Ctrl+Z undoes within it.
    const typing = view.isEditing || view.isEditingComment || view.isEditingShapeText || view.isRenaming;
    // The button of charts is named in the words of the view, which may have changed its language.
    const chart = this.buttons.chart;
    const insert = view.chartStrings.insert;
    if (chart && chart.getAttribute('aria-label') !== insert) {
      chart.setAttribute('aria-label', insert);
      chart.title = insert;
    }
    this.setDisabled('undo', view.readOnly || typing || !view.canUndo);
    this.setDisabled('redo', view.readOnly || typing || !view.canRedo);
    this.updateSave(typing);
    const others: ButtonKey[] = ['pageSetup', 'bold', 'italic', 'underline', 'strikethrough', 'textColor', 'fillColor', 'borders', 'alignLeft', 'alignCenter', 'alignRight', 'alignTop', 'alignMiddle', 'alignBottom', 'wrapText', 'merge', 'mergeOptions', 'sortAscending', 'sortDescending', 'filter', 'conditionalFormat', 'validation', 'chart', 'rowsColumns'];
    for (let i = 0; i < others.length; i++) this.setDisabled(others[i], !editable);
    // What the protection of the sheet forbids is off as well, and its title says why.
    const guarded: [ButtonKey[], 'formatCells' | 'merge' | 'sort' | 'filterOnOff' | 'rules' | 'objects'][] = [
      [['bold', 'italic', 'underline', 'strikethrough', 'textColor', 'fillColor', 'borders', 'alignLeft', 'alignCenter', 'alignRight', 'alignTop', 'alignMiddle', 'alignBottom', 'wrapText'], 'formatCells'],
      [['merge', 'mergeOptions'], 'merge'],
      [['sortAscending', 'sortDescending'], 'sort'],
      [['filter'], 'filterOnOff'],
      [['conditionalFormat', 'validation'], 'rules'],
      [['chart'], 'objects'],
    ];
    for (const [keys, action] of guarded) {
      const why = editable ? view.protectionReason(action) : undefined;
      for (const key of keys) {
        const button = this.buttons[key];
        if (!button) continue;
        if (why !== undefined) this.setDisabled(key, true);
        const title = why !== undefined ? why : key === 'chart' ? insert : this.labels[key];
        if (title !== undefined && title !== '' && button.title !== title) button.title = title;
      }
    }
    const formatWhy = editable ? view.protectionReason('formatCells') : undefined;
    // The picture button says why it is off where pictures cannot be changed (edited together, a protected sheet).
    const picture = this.buttons.insertPicture;
    if (picture) {
      const why = editable ? view.pictureProblem : undefined;
      this.setDisabled('insertPicture', !editable || why !== undefined);
      const title = why !== undefined ? why : this.strings.insertPicture;
      if (picture.title !== title) picture.title = title;
    }
    // So does the button of shapes; it is named in the words of the view.
    const shapeButton = this.buttons.insertShape;
    if (shapeButton) {
      const why = editable ? view.shapeProblem : undefined;
      this.setDisabled('insertShape', !editable || why !== undefined);
      const label = view.shapeStrings.insert;
      if (shapeButton.getAttribute('aria-label') !== label) shapeButton.setAttribute('aria-label', label);
      const title = why !== undefined ? why : label;
      if (shapeButton.title !== title) shapeButton.title = title;
    }
    this.fontSelect.disabled = !editable || formatWhy !== undefined;
    this.sizeSelect.disabled = !editable || formatWhy !== undefined;
    this.formatSelect.disabled = !editable || formatWhy !== undefined;
    const selects: [HTMLSelectElement, string][] = [
      [this.fontSelect, this.strings.fontName],
      [this.sizeSelect, this.strings.fontSize],
      [this.formatSelect, this.strings.numberFormat],
    ];
    for (const [select, label] of selects) {
      const title = formatWhy !== undefined ? formatWhy : label;
      if (select.title !== title) select.title = title;
    }
    if (!editable) this.closePopover();

    this.setPressed('bold', font.bold === true);
    this.setPressed('italic', font.italic === true);
    this.setPressed('underline', font.underline !== undefined && font.underline !== 'none');
    this.setPressed('strikethrough', font.strike === true);
    this.setPressed('alignLeft', horizontal === 'left');
    this.setPressed('alignCenter', horizontal === 'center');
    this.setPressed('alignRight', horizontal === 'right');
    // A cell without a vertical alignment is aligned at the bottom.
    const vertical = verticalOf(format.alignment.vertical);
    this.setPressed('alignTop', vertical === 'top');
    this.setPressed('alignMiddle', vertical === 'center');
    this.setPressed('alignBottom', vertical === 'bottom');
    this.setPressed('wrapText', format.alignment.wrapText === true);
    this.setPressed('merge', view.selectionMerged);
    this.setPressed('filter', view.hasFilter);

    this.showFont(typeof font.name === 'string' ? font.name : undefined);
    // A shape whose look can be changed is selected: the colours and the size of its text are shown, and are what the controls change.
    const shape = this.formattableShape();
    this.showFontSize(shape && shape.textSize !== undefined ? shape.textSize : fontSizeOf(format));
    this.showFormat(formatKindOf(format.numFmt, this.strings));
    this.paintBar(this.bars.textColor, shape ? shape.textColor || '#000000' : this.textColorOf(format) || '#000000');
    this.paintBar(this.bars.fillColor, shape ? shape.fill : this.fillColorOf(format));
    this.keepTabStop();
  }

  /**
   * The Save button: gone from a view that is read-only, dimmed while there is nothing to save, busy
   * while a save is under way. What is being typed counts as something to save: saving applies it first.
   */
  private updateSave(typing: boolean): void {
    const save = this.host.save;
    const button = this.buttons.save;
    if (!save || !button) return;
    const hidden = this.host.view.readOnly;
    if (button.hidden !== hidden) {
      button.hidden = hidden;
      if (this.saveSeparator) this.saveSeparator.hidden = hidden;
    }
    // A press while a save is under way is taken: it saves once more when something was changed meanwhile.
    this.setDisabled('save', hidden || !(save.dirty || save.saving || typing));
    const busy = save.saving ? 'true' : 'false';
    if (button.getAttribute('aria-busy') !== busy) button.setAttribute('aria-busy', busy);
  }

  /** Marks the Print button as busy while the pages of a print job are prepared. */
  setPrinting(busy: boolean): void {
    const button = this.buttons.print;
    const value = busy ? 'true' : 'false';
    if (button && button.getAttribute('aria-busy') !== value) button.setAttribute('aria-busy', value);
  }

  /** Marks the PDF button as busy while the pages of a PDF are painted. */
  setExporting(busy: boolean): void {
    const button = this.buttons.pdf;
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
    this.unwatchFonts();
    const element = this.element;
    element.removeEventListener('click', this.onClick);
    element.removeEventListener('mousedown', this.onMouseDown);
    element.removeEventListener('keydown', this.onKeyDown);
    element.removeEventListener('focusin', this.onFocusIn);
    element.removeEventListener('scroll', this.onScroll);
    this.fontSelect.removeEventListener('change', this.onFontChange);
    this.sizeSelect.removeEventListener('change', this.onSizeChange);
    this.formatSelect.removeEventListener('change', this.onFormatChange);
    this.fileInput.removeEventListener('change', this.onFileChange);
    if (element.parentNode) element.parentNode.removeChild(element);
  }

  /** A file chosen in the dialog of the picture button: read, and inserted at the active cell. Its name stays out of the workbook. */
  private onFileChange = (): void => {
    const input = this.fileInput;
    const file = input.files && input.files[0];
    // The same file can be chosen again next time.
    input.value = '';
    if (!file || this.destroyed) return;
    const view = this.host.view;
    if (file.size > MAX_PICTURE_BYTES) {
      // Refused before anything of it is read: the view says why in its own words.
      this.host.run(() => {
        throw new RangeError(`A picture cannot be larger than ${MAX_PICTURE_BYTES / 1024 / 1024} MB`);
      });
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (this.destroyed || !(reader.result instanceof ArrayBuffer)) return;
      const bytes = new Uint8Array(reader.result);
      // What kind of picture it is, the view reads from the bytes.
      this.host.run(() => void view.insertPicture(bytes));
    };
    // A file that cannot be read (gone, no permission) is said in the words of the view.
    reader.onerror = () => {
      if (!this.destroyed) {
        this.host.run(() => {
          throw new RangeError('The picture cannot be read');
        });
      }
    };
    reader.readAsArrayBuffer(file);
  };

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

  /** A button that opens a dialog of the view: the dialog takes the focus, so the click is not followed by a command's way back to the grid. */
  private addDialogButton(key: ButtonKey, open: () => void): void {
    const button = this.addButton(key, open);
    button.setAttribute('aria-haspopup', 'dialog');
    this.openers.add(button);
  }

  private addSelect(className: string): HTMLSelectElement {
    const select = document.createElement('select');
    select.className = `bym-tb-select ${className}`;
    this.element.appendChild(select);
    return select;
  }

  private addSeparator(): HTMLElement {
    const separator = document.createElement('span');
    separator.className = 'bym-tb-sep';
    separator.setAttribute('role', 'separator');
    separator.setAttribute('aria-orientation', 'vertical');
    this.element.appendChild(separator);
    return separator;
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
      // The same label as the button of the document bar.
      print: s.print,
      printPreview: withCtrlKey(s.printPreview, 'F2'),
      pdf: s.pdfDownload,
      pageSetup: s.pageSetup,
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
      alignTop: s.alignTop,
      alignMiddle: s.alignMiddle,
      alignBottom: s.alignBottom,
      wrapText: s.wrapText,
      merge: s.mergeCells,
      mergeOptions: s.mergeOptions,
      sortAscending: s.sortAscending,
      sortDescending: s.sortDescending,
      filter: s.filter,
      validation: s.dataValidation,
      conditionalFormat: s.conditionalFormat,
      // The words of charts are the view's (its menu and its dialog use the same): set once there is a view, see update().
      chart: '',
      insertPicture: s.insertPicture,
      // The words of shapes are the view's: set once there is a view, see update().
      insertShape: '',
      rowsColumns: s.rowsAndColumns,
      find: withShortcut(s.findReplace, 'F'),
    };
    this.labels = labels;
    for (const key of Object.keys(labels) as ButtonKey[]) {
      const button = this.buttons[key];
      if (!button || labels[key] === '') continue;
      button.setAttribute('aria-label', labels[key]);
      button.title = labels[key];
    }
    this.fontSelect.setAttribute('aria-label', s.fontName);
    this.fontSelect.title = s.fontName;
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

  /**
   * Selects a font in the list: the common fonts, the registered ones, those the workbook names
   * and the font of the active cell. A cell without a font name shows an entry without text.
   */
  private showFont(name: string | undefined): void {
    const fromFile: string[] = [];
    const fonts = this.host.view.workbook.styles.fonts;
    for (let i = 0; i < fonts.length && i < 4096; i++) if (typeof fonts[i].name === 'string') fromFile.push(fonts[i].name as string);
    const { names, shown } = fontChoices(fromFile, name);
    const key = names.join('\u0001');
    const select = this.fontSelect;
    if (key !== this.fontKey) {
      this.fontKey = key;
      while (select.firstChild) select.removeChild(select.firstChild);
      // An entry without text stands for "no name".
      const blank = option('\u0000', '');
      blank.hidden = true;
      select.appendChild(blank);
      for (let i = 0; i < names.length; i++) select.appendChild(option(names[i], names[i]));
    }
    const wanted = shown === undefined ? '\u0000' : shown;
    if (select.value !== wanted) select.value = wanted;
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

  /**
   * The three vertical alignments work like radio buttons. A cell without one is aligned at the bottom: "bottom"
   * is that, and pressing the one that is on goes back to it, as in Excel.
   */
  private alignVertically(vertical: 'top' | 'center' | 'bottom'): void {
    const view = this.host.view;
    const now = verticalOf(view.selectionFormat.alignment.vertical);
    view.format({ vertical: vertical === 'bottom' || now === vertical ? null : vertical });
  }

  /** The shape that is selected, while its look can be changed: the colours and the size of text of the toolbar are its own then. */
  private formattableShape(): ShapeInfo | undefined {
    const info = this.host.view.selectedShapeInfo;
    return info && info.formattable ? info : undefined;
  }

  private openPalette(key: 'textColor' | 'fillColor'): void {
    const view = this.host.view;
    const text = key === 'textColor';
    const format = view.selectionFormat;
    const shape = this.formattableShape();
    const current = shape ? (text ? shape.textColor : shape.fill) : text ? this.textColorOf(format) : this.fillColorOf(format);
    const apply = (rgb: string | null) => () =>
      this.host.run(() => {
        // A shape that is selected gets the colour; a line has neither a fill nor text of its own.
        if (shape) {
          if (!shape.line) view.formatShape(shape.id, text ? { textColor: rgb === null ? '#000000' : `#${rgb}` } : { fill: rgb === null ? null : `#${rgb}` });
          return;
        }
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

  /** The menu of conditional formats: the quick rules (each opens the dialog of a rule), clearing, and the list of rules. */
  private openConditional(): void {
    const view = this.host.view;
    const rules = view.ruleStrings;
    // A dialog takes the focus itself: its item is not followed by a command's way back to the grid.
    const dialog = (label: string, what: string): PopoverItem => ({ label, action: () => void view.openConditionalDialog(what) });
    const quick = rules.quickRules;
    const items: PopoverItem[] = [
      dialog(quick.greaterThan, 'greaterThan'),
      dialog(quick.lessThan, 'lessThan'),
      dialog(quick.between, 'between'),
      dialog(quick.equal, 'equal'),
      dialog(quick.containsText, 'containsText'),
      dialog(quick.duplicateValues, 'duplicateValues'),
      dialog(quick.top10, 'top10'),
      dialog(quick.colorScale, 'colorScale'),
      dialog(quick.dataBar, 'dataBar'),
      { label: rules.clearFromSelection, action: () => this.host.run(() => view.clearConditionalRules(view.selection.range)) },
      { label: rules.clearFromSheet, action: () => this.host.run(() => view.clearConditionalRules()) },
      dialog(rules.manageRules, 'manage'),
    ];
    this.open('conditionalFormat', this.strings.conditionalFormat, items);
  }

  /** The types of chart that can be added for the selection, and the dialog of the chart that is selected. */
  private openCharts(): void {
    const view = this.host.view;
    const charts = view.chartStrings;
    const add = (label: string, type: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter', stacked = false): PopoverItem => ({
      label,
      action: () => this.host.run(() => void view.addChart(stacked ? { type, grouping: 'stacked' } : { type })),
    });
    const items: PopoverItem[] = [
      add(charts.types.column, 'column'),
      add(charts.types.stackedColumn, 'column', true),
      add(charts.types.bar, 'bar'),
      add(charts.types.line, 'line'),
      add(charts.types.area, 'area'),
      add(charts.types.pie, 'pie'),
      add(charts.types.doughnut, 'doughnut'),
      add(charts.types.scatter, 'scatter'),
    ];
    // A dialog takes the focus itself: its item is not followed by a command's way back to the grid.
    if (view.selectedChart !== undefined) items.push({ label: charts.edit, action: () => void view.openChartDialog() });
    this.open('chart', charts.insert, items);
  }

  /**
   * The kinds of shape that can be inserted at the active cell, and, for a shape that is selected whose look can be
   * changed, its line: a colour from the palette (or none) and a weight from a list. Its fill and the colour and size
   * of its text are those of the fill and text colour buttons and the list of sizes.
   */
  private openShapes(): void {
    const view = this.host.view;
    const strings = view.shapeStrings;
    const add = (kind: ShapeKind): PopoverItem => ({ label: strings.kinds[kind], action: () => this.host.run(() => void view.insertShape(kind)) });
    const items: PopoverItem[] = (['rect', 'roundRect', 'ellipse', 'line', 'arrow', 'doubleArrow', 'textBox'] as ShapeKind[]).map(add);
    const shape = this.formattableShape();
    if (shape) {
      // The popover of the line takes the place of this one, under the same button.
      items.push({ label: strings.lineColor, action: () => this.openShapeLine(shape, 'color') });
      items.push({ label: strings.lineWidth, action: () => this.openShapeLine(shape, 'width') });
    }
    this.open('insertShape', strings.insert, items);
  }

  /** The colours (a palette, and none) or the weights (a list) of the line of the selected shape. */
  private openShapeLine(shape: ShapeInfo, what: 'color' | 'width'): void {
    const view = this.host.view;
    const strings = view.shapeStrings;
    // Opened from an item of the menu, which is closing: the new popover comes once it has.
    setTimeout(() => {
      if (this.destroyed) return;
      if (what === 'width') {
        const items = SHAPE_LINE_WEIGHTS.map(
          (points): PopoverItem => ({ label: strings.points(points), action: () => this.host.run(() => view.formatShape(shape.id, shape.outline ? { lineWidth: points } : { lineWidth: points, lineColor: '#000000' })) }),
        );
        this.open('insertShape', strings.lineWidth, items);
        return;
      }
      const current = shape.outline ? shape.outline.color : undefined;
      const swatches: PopoverItem[] = [];
      for (let i = 0; i < PALETTE.length; i++) {
        const hex = `#${PALETTE[i]}`;
        const name = this.strings.colorName(Math.floor(i / PALETTE_COLUMNS), i % PALETTE_COLUMNS);
        swatches.push({ label: `${name} (${hex})`, swatch: hex, value: hex, selected: current === hex, action: () => this.host.run(() => view.formatShape(shape.id, { lineColor: hex })) });
      }
      // A line keeps a line: it has nothing else to show.
      const reset: PopoverItem[] = shape.line ? [] : [{ label: strings.noLine, action: () => this.host.run(() => view.formatShape(shape.id, { lineColor: null })) }];
      this.open('insertShape', strings.lineColor, reset, swatches);
    }, 0);
  }

  /** The ways of merging the selection, and of splitting it again. */
  private openMerge(): void {
    const s = this.strings;
    const view = this.host.view;
    const items: PopoverItem[] = [
      { label: s.mergeAndCenter, action: () => this.host.run(() => void view.requestMerge({ center: true })) },
      { label: s.mergeCells, action: () => this.host.run(() => void view.requestMerge()) },
      { label: s.unmergeCells, action: () => this.host.run(() => view.unmerge()) },
    ];
    this.open('mergeOptions', s.mergeOptions, items);
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
    // Frozen panes: at the active cell (when it is not the first one), the top row, the first column, and none.
    const active = view.selection.active;
    const frozen = view.frozenPanes;
    if (active.row > 0 || active.col > 0) add(s.freezePanes, () => view.freezePanes());
    add(s.freezeTopRow, () => view.freezePanes(1, 0));
    add(s.freezeFirstColumn, () => view.freezePanes(0, 1));
    if (frozen.rows > 0 || frozen.cols > 0) add(s.unfreezePanes, () => view.unfreezePanes());
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
    if (target === this.fontSelect || target === this.sizeSelect || target === this.formatSelect) {
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

  private onFontChange = (): void => {
    const name = this.fontSelect.value;
    if (name === '\u0000' || name === '') return;
    // A value chosen with the keyboard may be one step on the way to another: the focus stays in the list.
    this.host.run(() => this.host.view.format({ fontName: name }), this.listByKey);
  };

  private onSizeChange = (): void => {
    const size = Number(this.sizeSelect.value);
    if (!(size > 0)) return;
    const shape = this.formattableShape();
    if (shape) {
      if (!shape.line) this.host.run(() => this.host.view.formatShape(shape.id, { textSize: size }), this.listByKey);
      return;
    }
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
