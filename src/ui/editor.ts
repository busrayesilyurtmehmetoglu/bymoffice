// The spreadsheet editor: the grid with a toolbar and a formula bar above it
// and a status bar below. The grid (WorkbookView) does all the work on the
// workbook; the shell only calls its public methods and shows its state.
//
// Nothing in this module touches the DOM at import time, so it can be
// imported where there is no document.

import type { People } from '../collab/people';
import { getLocale, normalizeLocale, type FormatLocale } from '../format/locale';
import { PeopleButton } from './collab-people';
import type { CellRange, Sheet, Workbook } from '../model/sheet/types';
import type { GridSelection } from '../render/grid';
import { getEditorStrings, type EditorStrings } from './editor-strings';
import { FormulaBar } from './formula-bar';
import type { SaveCommand } from './save-command';
import type { DocumentPrintResult } from './doc-print';
import type { SheetPrintOptions } from './sheet-print';
import type { PdfDownload } from '../pdf/export';
import type { SheetPdfExportOptions, SheetPreviewOptions } from './sheet-view';
import { WorkbookView, type WorkbookViewOptions } from './sheet-view';
import { StatusBar } from './status-bar';
import type { CollabPeer, CollabState } from '../collab/session';
import { Toolbar, type ShellHost } from './toolbar';

export interface SpreadsheetEditorOptions extends WorkbookViewOptions {
  /** Show the toolbar. Default true. */
  toolbar?: boolean;
  /** Show the formula bar. Default true. */
  formulaBar?: boolean;
  /** Show the status bar. Default true. */
  statusBar?: boolean;
  /**
   * What saves the workbook: with it the toolbar has a Save button (not while
   * the view is read-only). mount() hands it in for its `saveTo` option.
   */
  saveCommand?: SaveCommand;
}

// Colours are custom properties of the editor's root, so that a host page (or
// a theme, later) can change them in one place. The palette follows the grid:
// the accent is the colour of its selection border.
const STYLE_TEXT = [
  '.bym-editor{--bym-bg:#fff;--bym-surface:#f8f9fa;--bym-border:#dadce0;--bym-text:#202124;--bym-muted:#5f6368;',
  '--bym-hover:rgba(60,64,67,.09);--bym-active:rgba(60,64,67,.17);--bym-accent:#1a7f45;--bym-accent-soft:#d2e3d8;--bym-accent-text:#0f5c31;',
  '--bym-control:32px;--bym-bar:28px;--bym-swatch:22px;--bym-radius:6px;',
  'font:13px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:var(--bym-text);background:var(--bym-bg);',
  'letter-spacing:normal;text-transform:none;text-align:left;-webkit-text-size-adjust:100%;text-size-adjust:100%}',
  // Nothing a finger has to hit is smaller than 40 px (a colour is 36 px, with no gap to the next).
  '@media (pointer:coarse){.bym-editor{--bym-control:40px;--bym-bar:40px;--bym-swatch:36px}}',

  // A host page often styles bare elements ("button{min-width:120px}", "svg{width:100%}"). Every
  // property such a rule could reach is stated here for the shell's controls; a class beats an
  // element selector, so no rule needs to be marked important.
  '.bym-tb-btn,.bym-tb-select,.bym-namebox,.bym-formula-input,.bym-pop-item,.bym-pop-swatch{-webkit-appearance:none;appearance:none;',
  'box-sizing:border-box;float:none;margin:0;min-width:0;min-height:0;max-width:none;max-height:none;font:inherit;line-height:1.2;',
  'letter-spacing:normal;text-transform:none;text-indent:0;text-shadow:none;text-decoration:none;vertical-align:baseline;opacity:1}',
  '.bym-icon{display:block;flex:none;box-sizing:content-box;float:none;width:20px;height:20px;min-width:0;min-height:0;max-width:none;max-height:none;',
  'margin:0;padding:0;border:0;pointer-events:none}',
  '.bym-icon-bar{stroke:rgba(0,0,0,.28);stroke-width:.75}',

  '.bym-toolbar{display:flex;flex:none;align-items:center;box-sizing:border-box;padding:4px 8px;background:var(--bym-surface);',
  'border-bottom:1px solid var(--bym-border);overflow-x:auto;overflow-y:hidden;scrollbar-width:none;-webkit-overflow-scrolling:touch;outline:none}',
  '.bym-toolbar::-webkit-scrollbar{display:none}',
  '.bym-toolbar>*{flex:none;margin:0 1px}',
  '.bym-tb-btn{display:inline-flex;align-items:center;justify-content:center;width:var(--bym-control);height:var(--bym-control);padding:0;border:0;',
  'border-radius:var(--bym-radius);background:transparent;color:var(--bym-text);cursor:pointer;-webkit-tap-highlight-color:transparent;',
  'touch-action:manipulation}',
  '@media (hover:hover){.bym-tb-btn:hover,.bym-tb-select:hover{background-color:var(--bym-hover)}}',
  '.bym-tb-btn:active{background:var(--bym-active)}',
  '.bym-tb-btn[aria-pressed="true"],.bym-tb-btn[aria-expanded="true"]{background:var(--bym-accent-soft);color:var(--bym-accent-text)}',
  '.bym-tb-btn:disabled,.bym-tb-select:disabled{opacity:.38;cursor:default;background-color:transparent}',
  '.bym-tb-btn[aria-busy="true"]{cursor:progress}',
  '.bym-tb-btn:focus,.bym-tb-select:focus{outline:2px solid var(--bym-accent);outline-offset:-2px}',
  '.bym-tb-btn:focus:not(:focus-visible){outline:none}',
  '.bym-tb-sep{display:block;width:1px;height:20px;margin:0 5px;background:var(--bym-border)}',
  // The Save button of a view that is read-only: the rules above would show what is only marked as hidden.
  '.bym-tb-btn[hidden],.bym-tb-sep[hidden]{display:none}',
  // The arrow of a list is drawn with two gradients: no image, and it takes the text colour.
  '.bym-tb-select{display:block;height:var(--bym-control);padding:0 22px 0 8px;border:0;border-radius:var(--bym-radius);color:var(--bym-text);',
  'cursor:pointer;background-color:transparent;',
  'background-image:linear-gradient(45deg,transparent 50%,currentColor 50%),linear-gradient(135deg,currentColor 50%,transparent 50%);',
  'background-position:right 11px top 50%,right 6px top 50%;background-size:5px 5px,5px 5px;background-repeat:no-repeat}',
  '.bym-tb-size{width:58px}',
  '.bym-tb-format{width:112px;text-overflow:ellipsis}',

  '.bym-formula{display:flex;flex:none;align-items:stretch;box-sizing:content-box;height:var(--bym-bar);background:var(--bym-bg);',
  'border-bottom:1px solid var(--bym-border)}',
  '.bym-namebox,.bym-formula-input{display:block;height:var(--bym-bar);border:0;border-radius:0;padding:0 8px;background:transparent;',
  'color:var(--bym-text);outline:none}',
  '.bym-namebox{flex:none;width:88px;border-right:1px solid var(--bym-border)}',
  '.bym-fx{flex:none;display:flex;align-items:center;justify-content:center;width:34px;color:var(--bym-muted);',
  'font:italic 600 14px Georgia,"Times New Roman",serif;border-right:1px solid var(--bym-border);-webkit-user-select:none;user-select:none}',
  // One line of the cell at a time: a line is as high as the field, so the arrow keys bring one line after the other into view.
  '.bym-formula-input{flex:1 1 0;width:auto;line-height:var(--bym-bar);resize:none;overflow:hidden;white-space:pre;scrollbar-width:none}',
  '.bym-namebox:focus,.bym-formula-input:focus{box-shadow:inset 0 -2px 0 var(--bym-accent)}',
  // Text smaller than 16 px makes iOS zoom the page when a field takes the focus.
  '@media (pointer:coarse){.bym-namebox,.bym-formula-input,.bym-tb-select{font-size:16px}',
  '.bym-tb-size{width:68px}.bym-tb-format{width:128px}}',

  '.bym-status{display:flex;flex:none;align-items:center;box-sizing:border-box;height:24px;padding:0 10px;background:var(--bym-surface);',
  'border-top:1px solid var(--bym-border);color:var(--bym-muted);font-size:12px;white-space:nowrap;overflow:hidden}',
  '.bym-status-badge{flex:none;margin-right:10px;padding:1px 8px;border-radius:9px;background:var(--bym-accent-soft);color:var(--bym-accent-text);font-weight:600}',
  '.bym-status-badge[hidden]{display:none}',
  '.bym-status-message{flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis}',
  '.bym-status-summary{flex:0 1 auto;min-width:0;margin-left:auto;padding-left:12px;overflow:hidden;text-overflow:ellipsis}',

  '.bym-popover{padding:6px;background:var(--bym-bg);border:1px solid var(--bym-border);border-radius:8px;',
  'box-shadow:0 4px 16px rgba(0,0,0,.18);color:var(--bym-text);overscroll-behavior:contain}',
  '.bym-pop-item{display:flex;align-items:center;width:100%;height:auto;min-height:var(--bym-control);',
  'padding:0 14px 0 8px;border:0;border-radius:var(--bym-radius);background:transparent;color:inherit;text-align:left;',
  'white-space:nowrap;cursor:pointer}',
  '.bym-pop-item .bym-icon{margin-right:10px}',
  '@media (hover:hover){.bym-pop-item:hover{background:var(--bym-hover)}}',
  '.bym-pop-item:focus{background:var(--bym-hover);outline:none}',
  '.bym-pop-grid{display:grid;grid-gap:3px;gap:3px;margin-top:6px;padding:2px}',
  '.bym-pop-swatch{display:block;width:var(--bym-swatch);height:var(--bym-swatch);padding:0;border:1px solid rgba(0,0,0,.16);border-radius:50%;cursor:pointer}',
  // The whole 36 px square takes the finger; the colour is the circle inside a clear border.
  '@media (pointer:coarse){.bym-pop-grid{grid-gap:0;gap:0}',
  '.bym-pop-swatch{border:3px solid transparent;background-clip:padding-box;box-shadow:inset 0 0 0 1px rgba(0,0,0,.16)}}',
  '@media (hover:hover){.bym-pop-swatch:hover{box-shadow:0 0 0 2px var(--bym-active)}}',
  '.bym-pop-swatch:focus{outline:2px solid var(--bym-accent);outline-offset:1px}',
  '.bym-pop-swatch[aria-pressed="true"]{box-shadow:0 0 0 2px var(--bym-bg),0 0 0 4px var(--bym-accent)}',
].join('');

/** One style element per document or shadow root, shared by the editors in it. */
const styleUsers = new WeakMap<Node, { element: HTMLStyleElement; count: number }>();

function acquireStyle(container: HTMLElement): Node {
  const root = container.getRootNode();
  const host: Node = root.nodeType === 9 ? (root as Document).head || (root as Document).documentElement : root;
  const entry = styleUsers.get(host);
  if (entry) {
    entry.count++;
    return host;
  }
  const element = document.createElement('style');
  element.setAttribute('data-bymoffice-editor', '');
  element.textContent = STYLE_TEXT;
  host.appendChild(element);
  styleUsers.set(host, { element, count: 1 });
  return host;
}

function releaseStyle(host: Node): void {
  const entry = styleUsers.get(host);
  if (!entry || --entry.count > 0) return;
  if (entry.element.parentNode) entry.element.parentNode.removeChild(entry.element);
  styleUsers.delete(host);
}

const browserLanguage = () => (typeof navigator !== 'undefined' ? navigator.language : undefined);

function resolveLocale(locale: string | FormatLocale | undefined): FormatLocale {
  if (typeof locale === 'object' && locale !== null) return normalizeLocale(locale);
  return getLocale(locale !== undefined ? locale : browserLanguage());
}

const sameRange = (a: CellRange, b: CellRange) => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;

/** Whether a selection is one cell: a single cell, or exactly one merged cell. */
function isSingleCell(sheet: Sheet | undefined, selection: GridSelection): boolean {
  const range = selection.range;
  if (range.top === range.bottom && range.left === range.right) return true;
  if (!sheet || range.top !== selection.active.row || range.left !== selection.active.col) return false;
  const merges = sheet.merges;
  for (let i = 0; i < merges.length; i++) if (sameRange(merges[i], range)) return true;
  return false;
}

export class SpreadsheetEditor {
  readonly view: WorkbookView;
  private readonly root: HTMLDivElement;
  private readonly gridHost: HTMLDivElement;
  private readonly toolbar: Toolbar | undefined;
  private readonly formulaBar: FormulaBar | undefined;
  private readonly statusBar: StatusBar | undefined;
  private readonly styleHost: Node;
  private strings: EditorStrings;
  private locale: FormatLocale;
  /** What a command or the view last had to say; it stays until the selection moves on. */
  private notice = '';
  /** What the status bar shows while the pages of a print job are prepared. */
  private printProgress = '';
  /** The same while the pages of a PDF are painted. */
  private pdfProgress = '';
  private frame = 0;
  private destroyed = false;
  /** Ends the watch on the save command. */
  private unwatchSave: (() => void) | undefined;

  constructor(container: HTMLElement, workbook: Workbook, options: SpreadsheetEditorOptions = {}) {
    this.strings = getEditorStrings(typeof options.locale === 'string' ? options.locale : browserLanguage());
    this.locale = resolveLocale(options.locale);

    const root = document.createElement('div');
    root.className = 'bym-editor';
    // What the layout depends on is set inline, as in the view: it holds where a style element is not allowed.
    root.style.cssText = 'position:relative;display:flex;flex-direction:column;width:100%;height:100%;box-sizing:border-box;direction:ltr';
    const gridHost = document.createElement('div');
    gridHost.className = 'bym-editor-grid';
    gridHost.style.cssText = 'position:relative;flex:1 1 0;min-height:0;min-width:0;overflow:hidden';
    this.root = root;
    this.gridHost = gridHost;

    const self = this;
    const host: ShellHost = {
      get view() {
        return self.view;
      },
      root,
      run: (action, keepFocus) => this.run(action, keepFocus === true),
      save: options.saveCommand,
    };
    if (options.toolbar !== false) {
      this.toolbar = new Toolbar(host, this.strings);
      root.appendChild(this.toolbar.element);
    }
    if (options.formulaBar !== false) {
      this.formulaBar = new FormulaBar(host, this.strings);
      root.appendChild(this.formulaBar.element);
    }
    root.appendChild(gridHost);
    if (options.statusBar !== false) {
      this.statusBar = new StatusBar(this.strings);
      root.appendChild(this.statusBar.element);
    }

    this.styleHost = acquireStyle(container);
    container.appendChild(root);

    // The host page's callbacks still run; the shell listens first.
    const viewOptions: WorkbookViewOptions = {
      ...options,
      focusScope: root,
      onSelectionChange: (selection, sheetIndex) => {
        this.notice = '';
        if (this.statusBar) this.statusBar.showMessage('');
        this.refresh();
        if (options.onSelectionChange) options.onSelectionChange(selection, sheetIndex);
      },
      onSheetChange: (sheetIndex) => {
        this.refresh();
        if (options.onSheetChange) options.onSheetChange(sheetIndex);
      },
      onNotice: (notice) => {
        // What the view tells assistive technology (a command from a key or its menu failed, say) is shown as well.
        // What the view also has as its standing status (how many rows a filter shows) is shown from
        // there, and goes when the status does; anything else stays until the selection moves on.
        this.notice = this.view && notice.message === this.view.statusText ? '' : notice.message;
        if (this.statusBar) this.statusBar.showMessage(notice.message);
        if (options.onNotice) options.onNotice(notice);
      },
      onChange: (change) => {
        this.refresh();
        if (options.onChange) options.onChange(change);
      },
      onEditingChange: (state) => {
        if (this.formulaBar && this.view) this.formulaBar.editingChanged(state);
        // Undo and redo are off while a cell is edited.
        this.refresh();
        if (options.onEditingChange) options.onEditingChange(state);
      },
      onPrintProgress: (progress) => {
        if (this.toolbar) this.toolbar.setPrinting(progress !== undefined);
        this.printProgress = progress && progress.total > 0 ? this.strings.printProgress(progress.done, progress.total) : '';
        // In the status bar only: assistive technology heard that printing is being prepared, and hears when it is over.
        if (this.statusBar && this.view) this.statusBar.showMessage(this.printProgress || this.pdfProgress || this.notice || this.view.statusText);
        if (options.onPrintProgress) options.onPrintProgress(progress);
      },
      onPdfProgress: (progress) => {
        if (this.toolbar) this.toolbar.setExporting(progress !== undefined);
        this.pdfProgress = progress && progress.total > 0 ? this.strings.pdfProgress(progress.done, progress.total) : '';
        // In the status bar only, as for printing.
        if (this.statusBar && this.view) this.statusBar.showMessage(this.printProgress || this.pdfProgress || this.notice || this.view.statusText);
        if (options.onPdfProgress) options.onPdfProgress(progress);
      },
    };
    try {
      this.view = new WorkbookView(gridHost, workbook, viewOptions);
    } catch (error) {
      // Nothing of a failed editor stays behind.
      this.destroyed = true;
      if (this.frame) cancelAnimationFrame(this.frame);
      if (this.toolbar) this.toolbar.destroy();
      if (this.formulaBar) this.formulaBar.destroy();
      if (root.parentNode) root.parentNode.removeChild(root);
      releaseStyle(this.styleHost);
      throw error;
    }

    window.addEventListener('resize', this.onResize);
    root.addEventListener('keydown', this.onKeyDown);
    // The Save button follows what there is to save.
    if (options.saveCommand && this.toolbar) this.unwatchSave = options.saveCommand.watch(() => this.refresh());
    this.sync();
  }

  /** The row that says whether the workbook is edited together and who else is there; made when it is first needed. */
  private collabBar: HTMLDivElement | undefined;
  private peopleButton: PeopleButton | undefined;

  /**
   * Shows the state of editing together and the other participants above the grid; undefined
   * takes the row away. Names are shown as text.
   */
  setCollab(info: { state: CollabState; peers: readonly CollabPeer[]; pending: number; autoWaits?: 'room' | 'offline'; people?: People } | undefined): void {
    if (this.destroyed) return;
    // Buttons that a room turns off (pictures) follow.
    this.refresh();
    if (!info) {
      if (this.peopleButton) this.peopleButton.destroy();
      this.peopleButton = undefined;
      if (this.collabBar && this.collabBar.parentNode) this.collabBar.parentNode.removeChild(this.collabBar);
      this.collabBar = undefined;
      return;
    }
    const words = this.view.collabStrings;
    let bar = this.collabBar;
    if (!bar) {
      bar = document.createElement('div');
      bar.className = 'bym-collab';
      bar.setAttribute('role', 'status');
      bar.style.cssText = 'display:flex;align-items:center;gap:6px;flex:0 0 auto;padding:3px 8px;font:12px system-ui,sans-serif;color:#3c4043;background:#f8f9fa;border-bottom:1px solid #dadce0;overflow:hidden;white-space:nowrap';
      this.root.insertBefore(bar, this.gridHost);
      this.collabBar = bar;
    }
    let button = this.peopleButton;
    if (!button) {
      button = new PeopleButton(this.root, words);
      this.peopleButton = button;
      bar.appendChild(button.element);
    }
    // Everything but the button is made anew (the button stays where it is: it may have the focus, and its list may be open).
    for (const child of Array.from(bar.childNodes)) if (child !== button.element) bar.removeChild(child);
    bar.setAttribute('data-state', info.state);
    bar.setAttribute('data-peers', String(info.peers.length));
    const dot = document.createElement('span');
    dot.setAttribute('aria-hidden', 'true');
    dot.style.cssText = `display:inline-block;flex:0 0 auto;width:8px;height:8px;border-radius:50%;background:${info.state === 'online' ? '#188038' : info.state === 'connecting' ? '#E37400' : '#D93025'}`;
    const label = document.createElement('span');
    label.className = 'bym-collab-state';
    label.style.cssText = 'flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis';
    label.textContent = words.state[info.state];
    bar.insertBefore(dot, button.element);
    bar.insertBefore(label, button.element);
    // Auto-save that holds still on purpose says so, and why.
    if (info.autoWaits) {
      bar.setAttribute('data-auto-waits', info.autoWaits);
      const waits = document.createElement('span');
      waits.className = 'bym-collab-auto';
      waits.style.cssText = 'flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis';
      waits.textContent = words.autoWaits[info.autoWaits];
      bar.insertBefore(waits, button.element);
    } else bar.removeAttribute('data-auto-waits');
    // Who is in the file: how many (this participant too), always; the first of the others by their letters, where there is room.
    const people: People = info.people || { list: [], more: 0, count: 1 + info.peers.length };
    const view = this.view;
    button.update(people, words, (index) => {
      const sheet = view.workbook.sheets[index];
      return sheet ? sheet.name : undefined;
    });
    bar.setAttribute('data-people', String(people.count));
    const narrow = this.root.clientWidth > 0 && this.root.clientWidth < 480;
    if (narrow) bar.setAttribute('data-narrow', 'true');
    else bar.removeAttribute('data-narrow');
    const others = people.list.filter((who) => !who.self);
    const shown = narrow ? [] : others.slice(0, 3);
    for (const who of shown) bar.appendChild(PeopleButton.avatar(who.name, who.color, who.id));
    const rest = people.count - 1 - shown.length;
    if (!narrow && rest > 0) {
      const more = document.createElement('span');
      more.className = 'bym-collab-more';
      more.setAttribute('aria-hidden', 'true');
      more.textContent = `+${rest}`;
      bar.appendChild(more);
    }
    bar.setAttribute('aria-label', `${words.state[info.state]}. ${words.peopleCount(people.count)}`);
  }

  /** Opens the print preview: see WorkbookView.printPreview(). */
  printPreview(options?: SheetPreviewOptions): boolean {
    return this.view.printPreview(options);
  }

  /** Closes the print preview: see WorkbookView.closePrintPreview(). */
  closePrintPreview(): boolean {
    return this.view.closePrintPreview();
  }

  /** Prints the workbook: see WorkbookView.print(). */
  print(options?: SheetPrintOptions): Promise<DocumentPrintResult> {
    return this.view.print(options);
  }

  /** The workbook as a PDF file: see WorkbookView.exportPdf(). */
  exportPdf(options?: SheetPdfExportOptions): Promise<Uint8Array> {
    return this.view.exportPdf(options);
  }

  /** Makes the PDF and hands it to the browser as a download: see WorkbookView.downloadPdf(). */
  downloadPdf(options?: SheetPdfExportOptions): Promise<PdfDownload> {
    return this.view.downloadPdf(options);
  }

  /** Changes the language of number and date formats and, for a language tag, of every label. */
  setLocale(locale: string | FormatLocale): void {
    if (this.destroyed) return;
    this.view.setLocale(locale);
    this.locale = resolveLocale(locale);
    // As in the view: a locale object only changes formats.
    if (typeof locale === 'string') {
      this.strings = getEditorStrings(locale);
      if (this.toolbar) this.toolbar.setStrings(this.strings);
      if (this.formulaBar) this.formulaBar.setStrings(this.strings);
      if (this.statusBar) this.statusBar.setStrings(this.strings);
    }
    this.sync();
  }

  /** Turns editing off or on in the grid and in the shell. */
  setReadOnly(readOnly: boolean): void {
    if (this.destroyed) return;
    this.view.setReadOnly(readOnly);
    this.sync();
  }

  /** Removes the editor from its container with every listener and pending frame. Safe to call twice. */
  /** The root element of the editor: the shell with the view in it, as it was put into the container. */
  get element(): HTMLElement {
    return this.root;
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.frame) cancelAnimationFrame(this.frame);
    this.frame = 0;
    window.removeEventListener('resize', this.onResize);
    if (this.peopleButton) this.peopleButton.destroy();
    this.peopleButton = undefined;
    this.root.removeEventListener('keydown', this.onKeyDown);
    if (this.unwatchSave) this.unwatchSave();
    this.unwatchSave = undefined;
    if (this.toolbar) this.toolbar.destroy();
    if (this.formulaBar) this.formulaBar.destroy();
    if (this.statusBar) this.statusBar.destroy();
    this.view.destroy();
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
  }

  private run(action: () => void, keepFocus: boolean): void {
    if (this.destroyed) return;
    this.notice = '';
    if (this.statusBar) this.statusBar.showMessage('');
    try {
      action();
    } catch (error) {
      // In the language of the interface; the operations themselves speak English.
      this.notice = this.view.localizeError(error);
      if (this.statusBar) this.statusBar.showMessage(this.notice);
    }
    if (!keepFocus) this.view.focus();
    this.refresh();
  }

  /** Brings the shell up to date in the next frame; any number of changes before it cost one update. */
  private refresh(): void {
    if (this.destroyed || this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.sync();
    });
  }

  private sync(): void {
    const view = this.view;
    // A callback of the view can come before the constructor has the view.
    if (this.destroyed || !view) return;
    const selection = view.selection;
    const sheet: Sheet | undefined = view.workbook.sheets[view.activeSheet];
    const single = isSingleCell(sheet, selection);
    if (this.toolbar) this.toolbar.update();
    if (this.formulaBar) this.formulaBar.update(single);
    if (this.statusBar) {
      // With nothing else to say: what the view knows about the active cell and the filter of the sheet.
      this.statusBar.showMessage(this.printProgress || this.pdfProgress || this.notice || view.statusText);
      this.statusBar.setReadOnly(view.readOnly);
      this.statusBar.showSummary(sheet, selection.range, single, this.locale);
    }
  }

  /** Ctrl+P in the toolbar and the formula bar prints the workbook as it does in the grid (which handles it itself); Ctrl+F and Ctrl+H find. */
  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented || !(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey) return;
    const key = event.key;
    // Ctrl+F and Ctrl+H open the bar of Find and Replace there too.
    if ((key === 'f' || key === 'F' || key === 'h' || key === 'H') && !this.view.isEditing) {
      event.preventDefault();
      this.view.openFind(key === 'h' || key === 'H');
      return;
    }
    // Ctrl+F2: the print preview, as from the grid.
    if (key === 'F2' && event.ctrlKey && !event.metaKey && !this.view.isEditing) {
      event.preventDefault();
      this.view.printPreview();
      return;
    }
    if (key !== 'p' && key !== 'P') return;
    event.preventDefault();
    this.view.print().catch(() => undefined);
  };

  private onResize = (): void => {
    if (this.toolbar) this.toolbar.closePopover();
  };
}
