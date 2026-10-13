// The document editor: the pages (DocumentView) with a formatting toolbar
// above them. The view does all the work on the document; the shell only
// calls its public methods and shows its state. A document that is shown
// read-only has no toolbar: it looks as the view alone does (one whose
// protection allows comments has the toolbar, with the comments alone).
//
// Nothing in this module touches the DOM at import time.

import type { DocModel } from '../model/doc/types';
import type { DocRange } from '../ops/doc/index';
import { DocToolbar, type DocShellHost } from './doc-toolbar';
import type { DocumentPrintOptions, DocumentPrintResult } from './doc-print';
import type { PdfDownload, PdfExportOptions } from '../pdf/export';
import { DocumentView, type DocumentViewOptions } from './doc-view';
import { getEditorStrings, type EditorStrings } from './editor-strings';
import type { SaveCommand } from './save-command';

export interface DocumentEditorOptions extends DocumentViewOptions {
  /** Show the formatting toolbar (while the document can be edited) and the bar with the zoom and the page number. Default true. */
  toolbar?: boolean;
  /**
   * What saves the document: with it the toolbar has a Save button. mount()
   * hands it in for its `saveTo` option.
   */
  saveCommand?: SaveCommand;
}

// The same custom properties and classes as the spreadsheet editor has, so that the two look alike and a
// host page themes both in one place. The rules are stated here again: either editor may be the only one on a page.
const STYLE_TEXT = [
  '.bym-editor{--bym-bg:#fff;--bym-surface:#f8f9fa;--bym-border:#dadce0;--bym-text:#202124;--bym-muted:#5f6368;',
  '--bym-hover:rgba(60,64,67,.09);--bym-active:rgba(60,64,67,.17);--bym-accent:#1a7f45;--bym-accent-soft:#d2e3d8;--bym-accent-text:#0f5c31;',
  '--bym-control:32px;--bym-bar:28px;--bym-swatch:22px;--bym-radius:6px;',
  'font:13px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:var(--bym-text);background:var(--bym-bg);',
  'letter-spacing:normal;text-transform:none;text-align:left;-webkit-text-size-adjust:100%;text-size-adjust:100%}',
  '@media (pointer:coarse){.bym-editor{--bym-control:40px;--bym-bar:40px;--bym-swatch:36px}}',
  // Every property a rule of the host page for bare elements could reach is stated for the shell's controls.
  '.bym-tb-btn,.bym-tb-select,.bym-pop-item,.bym-pop-swatch,.bym-doc-linkinput{-webkit-appearance:none;appearance:none;',
  'box-sizing:border-box;float:none;margin:0;min-width:0;min-height:0;max-width:none;max-height:none;font:inherit;line-height:1.2;',
  'letter-spacing:normal;text-transform:none;text-indent:0;text-shadow:none;text-decoration:none;vertical-align:baseline;opacity:1}',
  '.bym-icon{display:block;flex:none;box-sizing:content-box;float:none;width:20px;height:20px;min-width:0;min-height:0;max-width:none;max-height:none;',
  'margin:0;padding:0;border:0;pointer-events:none}',
  '.bym-icon-bar{stroke:rgba(0,0,0,.28);stroke-width:.75}',
  // One row that scrolls sideways when the editor is narrower than its buttons.
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
  // Part of the selection has it, part has not.
  '.bym-tb-btn[aria-pressed="mixed"]{background:var(--bym-hover)}',
  '.bym-tb-btn:disabled,.bym-tb-select:disabled{opacity:.38;cursor:default;background-color:transparent}',
  '.bym-tb-btn[aria-busy="true"]{cursor:progress}',
  '.bym-tb-btn:focus,.bym-tb-select:focus{outline:2px solid var(--bym-accent);outline-offset:-2px}',
  '.bym-tb-btn:focus:not(:focus-visible){outline:none}',
  '.bym-tb-sep{display:block;width:1px;height:20px;margin:0 5px;background:var(--bym-border)}',
  '.bym-tb-select{display:block;height:var(--bym-control);padding:0 22px 0 8px;border:0;border-radius:var(--bym-radius);color:var(--bym-text);',
  'cursor:pointer;background-color:transparent;text-overflow:ellipsis;',
  'background-image:linear-gradient(45deg,transparent 50%,currentColor 50%),linear-gradient(135deg,currentColor 50%,transparent 50%);',
  'background-position:right 11px top 50%,right 6px top 50%;background-size:5px 5px,5px 5px;background-repeat:no-repeat}',
  '.bym-tb-size{width:58px}.bym-tb-style{width:118px}.bym-tb-font{width:128px}',
  // Text smaller than 16 px makes iOS zoom the page when a field takes the focus.
  '@media (pointer:coarse){.bym-tb-select,.bym-doc-linkinput{font-size:16px}.bym-tb-size{width:68px}.bym-tb-style{width:132px}.bym-tb-font{width:140px}}',
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
  '@media (pointer:coarse){.bym-pop-grid{grid-gap:0;gap:0}',
  '.bym-pop-swatch{border:3px solid transparent;background-clip:padding-box;box-shadow:inset 0 0 0 1px rgba(0,0,0,.16)}}',
  '@media (hover:hover){.bym-pop-swatch:hover{box-shadow:0 0 0 2px var(--bym-active)}}',
  '.bym-pop-swatch:focus{outline:2px solid var(--bym-accent);outline-offset:1px}',
  '.bym-pop-swatch[aria-pressed="true"]{box-shadow:0 0 0 2px var(--bym-bg),0 0 0 4px var(--bym-accent)}',
  // The cells of the grid that inserts a table are squares.
  '.bym-popover[aria-label] .bym-pop-swatch[data-color*="x"]{border-radius:2px}',
  '.bym-doc-linkinput{display:block;width:220px;max-width:60vw;height:var(--bym-control);margin:0 6px 0 0;padding:0 8px;border:1px solid var(--bym-border);',
  'border-radius:var(--bym-radius);background:var(--bym-bg);color:var(--bym-text);outline:none}',
  '.bym-doc-linkinput:focus{border-color:var(--bym-accent)}',
].join('');

/** One style element per document or shadow root, shared by the editors in it. */
const styleUsers = new WeakMap<Node, { element: HTMLStyleElement; count: number }>();

function acquireStyle(container: HTMLElement): Node {
  const root: Node = typeof container.getRootNode === 'function' ? container.getRootNode() : container.ownerDocument || document;
  const host: Node = root.nodeType === 9 ? (root as Document).head || (root as Document).documentElement : root;
  const entry = styleUsers.get(host);
  if (entry) {
    entry.count++;
    return host;
  }
  const element = document.createElement('style');
  element.setAttribute('data-bymoffice-doc-editor', '');
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

export class DocumentEditor {
  readonly view: DocumentView;
  private readonly root: HTMLDivElement;
  private readonly toolbar: DocToolbar | undefined;
  private readonly styleHost: Node;
  private strings: EditorStrings;
  private frame = 0;
  private destroyed = false;
  /** Ends the watch on the save command. */
  private unwatchSave: (() => void) | undefined;

  constructor(container: HTMLElement, doc: DocModel, options: DocumentEditorOptions = {}) {
    this.strings = getEditorStrings(typeof options.locale === 'string' ? options.locale : browserLanguage());

    const root = document.createElement('div');
    root.className = 'bym-editor bym-doc-editor';
    // What the layout depends on is set inline, as in the view: it holds where a style element is not allowed.
    root.style.cssText = 'position:relative;display:flex;flex-direction:column;width:100%;height:100%;box-sizing:border-box;direction:ltr';
    const pages = document.createElement('div');
    pages.className = 'bym-doc-editor-pages';
    pages.style.cssText = 'position:relative;flex:1 1 0;min-height:0;min-width:0;overflow:hidden';
    this.root = root;

    const self = this;
    const host: DocShellHost = {
      get view() {
        return self.view;
      },
      root,
      run: (action, keepFocus) => this.run(action, keepFocus === true),
      save: options.saveCommand,
    };
    if (options.toolbar !== false) {
      this.toolbar = new DocToolbar(host, this.strings);
      // Shown once the view says the document can be edited.
      this.toolbar.element.style.display = 'none';
      root.appendChild(this.toolbar.element);
    }
    root.appendChild(pages);

    this.styleHost = acquireStyle(container);
    container.appendChild(root);

    // The host page's callbacks still run; the shell listens first.
    const viewOptions: DocumentViewOptions = {
      ...options,
      onSelectionChange: (selection: DocRange | undefined) => {
        this.refresh();
        if (options.onSelectionChange) options.onSelectionChange(selection);
      },
      onChange: () => {
        this.refresh();
        if (options.onChange) options.onChange();
      },
      onLinkShortcut: () => {
        if (this.toolbar && !this.view.readOnly) this.toolbar.openLink();
        if (options.onLinkShortcut) options.onLinkShortcut();
      },
    };
    try {
      this.view = new DocumentView(pages, doc, viewOptions);
    } catch (error) {
      // Nothing of a failed editor stays behind.
      this.destroyed = true;
      if (this.frame) cancelAnimationFrame(this.frame);
      if (this.toolbar) this.toolbar.destroy();
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

  /** Prints the document: see DocumentView.print(). */
  print(options?: DocumentPrintOptions): Promise<DocumentPrintResult> {
    return this.view.print(options);
  }

  /** The document as a PDF file: see DocumentView.exportPdf(). */
  exportPdf(options?: PdfExportOptions): Promise<Uint8Array> {
    return this.view.exportPdf(options);
  }

  /** Makes the PDF and hands it to the browser as a download: see DocumentView.downloadPdf(). */
  downloadPdf(options?: PdfExportOptions): Promise<PdfDownload> {
    return this.view.downloadPdf(options);
  }

  /** Changes the language of every label. */
  setLocale(locale: string | object): void {
    if (this.destroyed) return;
    this.view.setLocale(locale);
    if (typeof locale === 'string') {
      this.strings = getEditorStrings(locale);
      if (this.toolbar) this.toolbar.setStrings(this.strings);
    }
    this.sync();
  }

  /** Turns editing off or on, in the pages and in the shell. */
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
    this.root.removeEventListener('keydown', this.onKeyDown);
    if (this.unwatchSave) this.unwatchSave();
    this.unwatchSave = undefined;
    if (this.toolbar) this.toolbar.destroy();
    this.view.destroy();
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
  }

  private run(action: () => void, keepFocus: boolean): void {
    if (this.destroyed) return;
    // The view says by itself what it could not do (a part that cannot be edited, an address that is no link).
    action();
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
    if (this.destroyed || !view || !this.toolbar) return;
    // Shown while the document can be edited, or at least commented (a protection that allows comments only).
    const display = view.readOnly && !view.canComment ? 'none' : '';
    if (this.toolbar.element.style.display !== display) this.toolbar.element.style.display = display;
    this.toolbar.update();
  }

  /** Ctrl+P on the toolbar prints the document as it does in the pages (which handle it themselves). */
  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented || !(event.ctrlKey || event.metaKey) || event.altKey || event.shiftKey) return;
    if (event.key !== 'p' && event.key !== 'P') return;
    event.preventDefault();
    this.view.print().catch(() => undefined);
  };

  private onResize = (): void => {
    if (this.toolbar) this.toolbar.closePopover();
  };
}
