// Text input for the document editor.
//
// The pages are canvases, so what is typed arrives in a hidden textarea that
// holds the focus. Two ways lead from there to the document:
//
// 1. Plain keys. `beforeinput` says what is about to happen (a character, a
//    deletion); it is cancelled and done to the document instead. The textarea
//    never changes.
//
// 2. Everything a browser will not let be cancelled or does in steps of its
//    own: input methods (compositions), the word an Android keyboard is still
//    "composing", the replacement iOS makes when it corrects a word. For
//    these the textarea is a MIRROR of the text around the caret: it holds
//    some of the paragraph before the caret and some after it, with its own
//    caret between the two. The browser changes the mirror as it likes; after
//    every `input` event the new content is compared with the old, and the
//    difference (so many characters gone before the caret, so many after it,
//    this text in their place) is done to the document. Then, unless a
//    composition is under way, the mirror is filled again from the document.
//    During a composition it is left alone, because writing to the field
//    would end the composition; the edits of one composition are one undo
//    step, and a composition that is cancelled takes its text back through
//    the same comparison.
//
// The mirror also gives the keyboard of a phone what it needs for its
// suggestions (the words around the caret), and its place on screen (the
// editor keeps the field at the caret) tells an input method where to open.
//
// Nothing in this module touches the DOM at import time.

/** The text around the caret (or the selection) as the mirror shows it. */
export interface InputContext {
  /** Text of the paragraph before the caret; with `atStart`, a line end comes before it that stands for the paragraph boundary. */
  before: string;
  after: string;
  /** Stands for the selected text when there is a selection: typing over it or deleting it replaces the selection. */
  selected: string;
  /** The context reaches the start (end) of the paragraph. */
  atStart: boolean;
  atEnd: boolean;
}

/** One change of the mirror, in terms of the document. */
export interface InputChange {
  /** Characters removed before and after the caret; one more than the paragraph has there joins it with its neighbour. */
  removeBefore: number;
  removeAfter: number;
  /** The selection was replaced (or removed). */
  replaceSelection: boolean;
  text: string;
  /** Where the caret of the field is after the change, counted back from the end of `text`: 0 when it is right after it. */
  caretBack: number;
  /** The text that is being composed, when a composition is under way: shown underlined. */
  composing: number;
}

export interface InputHost {
  /** Undefined when nothing can be typed (no caret, read-only). */
  context(): InputContext | undefined;
  /** A change the browser made in the field. `composition` names the composition it is part of. */
  change(change: InputChange, composition: string | undefined): void;
  /** Cancelled keys: what they would have done. */
  type(text: string): void;
  remove(direction: -1 | 1, unit: 'char' | 'word' | 'line'): void;
  paragraph(): void;
  undo(): void;
  redo(): void;
  compositionStart(): void;
  compositionEnd(): void;
}

const LINE_END = '\n';

export class DocInput {
  private value = ' ';
  private start = 0;
  private end = 1;
  private selected = false;
  private composition: string | undefined;
  private compositions = 0;
  private composed = '';
  private destroyed = false;

  constructor(
    readonly element: HTMLTextAreaElement,
    private readonly host: InputHost,
  ) {
    element.addEventListener('beforeinput', this.onBeforeInput as EventListener);
    element.addEventListener('input', this.onInput);
    element.addEventListener('compositionstart', this.onCompositionStart);
    element.addEventListener('compositionupdate', this.onCompositionUpdate as EventListener);
    element.addEventListener('compositionend', this.onCompositionEnd);
  }

  /** A composition is under way: the field must be left alone. */
  get composing(): boolean {
    return this.composition !== undefined;
  }

  destroy(): void {
    this.destroyed = true;
    const element = this.element;
    element.removeEventListener('beforeinput', this.onBeforeInput as EventListener);
    element.removeEventListener('input', this.onInput);
    element.removeEventListener('compositionstart', this.onCompositionStart);
    element.removeEventListener('compositionupdate', this.onCompositionUpdate as EventListener);
    element.removeEventListener('compositionend', this.onCompositionEnd);
  }

  /** Fills the mirror from the document. Does nothing during a composition. */
  sync(): void {
    if (this.destroyed || this.composition !== undefined) return;
    const context = this.host.context();
    let value = ' ';
    let start = 0;
    let end = 1;
    this.selected = true;
    if (context) {
      const before = (context.atStart ? LINE_END : '') + context.before;
      // Something is always selected when there is a selection: some browsers send no copy event otherwise.
      const middle = context.selected;
      value = before + middle + context.after + (context.atEnd ? LINE_END : '');
      start = before.length;
      end = start + middle.length;
      this.selected = middle.length > 0;
    }
    this.value = value;
    this.start = start;
    this.end = end;
    const element = this.element;
    if (element.value !== value) element.value = value;
    this.select();
  }

  /** Puts the selection of the field where the mirror has it; only when the field has the focus (elsewhere some browsers would move the focus). */
  select(): void {
    const element = this.element;
    const root = typeof element.getRootNode === 'function' ? (element.getRootNode() as Document | ShadowRoot) : element.ownerDocument;
    if (!root || root.activeElement !== element) return;
    try {
      if (element.selectionStart !== this.start || element.selectionEnd !== this.end) element.setSelectionRange(this.start, this.end);
    } catch {
      // A field that is not shown has no selection to set.
    }
  }

  /** Ends a composition that is under way: what it has typed so far stays. */
  commit(): void {
    if (this.composition === undefined) return;
    this.read();
    this.finish();
    // Writing to the field ends the composition in the browser as well.
    this.element.value = '';
    this.sync();
  }

  private finish(): void {
    if (this.composition === undefined) return;
    this.composition = undefined;
    this.composed = '';
    this.host.compositionEnd();
  }

  private onCompositionStart = (): void => {
    if (this.destroyed || this.composition !== undefined) return;
    if (!this.host.context()) return;
    this.composition = 'ime:' + ++this.compositions;
    this.composed = '';
    this.host.compositionStart();
  };

  private onCompositionUpdate = (event: CompositionEvent): void => {
    this.composed = typeof event.data === 'string' ? event.data : '';
  };

  private onCompositionEnd = (): void => {
    if (this.destroyed) return;
    // Some browsers change the field before this event, some after it; what is there now is read either way.
    this.composed = '';
    this.read();
    this.finish();
    this.sync();
  };

  private onBeforeInput = (event: InputEvent): void => {
    if (this.destroyed) return;
    const host = this.host;
    if (!host.context()) {
      // Nothing can be typed: the field stays as it is.
      if (event.cancelable) event.preventDefault();
      return;
    }
    if (this.composition !== undefined || event.isComposing || !event.cancelable) return;
    switch (event.inputType) {
      case 'insertText':
        if (typeof event.data !== 'string') return;
        event.preventDefault();
        if (event.data !== '') host.type(event.data);
        return;
      case 'insertLineBreak':
      case 'insertParagraph':
        event.preventDefault();
        host.paragraph();
        return;
      case 'deleteContentBackward':
        event.preventDefault();
        host.remove(-1, 'char');
        return;
      case 'deleteContentForward':
        event.preventDefault();
        host.remove(1, 'char');
        return;
      case 'deleteWordBackward':
        event.preventDefault();
        host.remove(-1, 'word');
        return;
      case 'deleteWordForward':
        event.preventDefault();
        host.remove(1, 'word');
        return;
      case 'deleteSoftLineBackward':
      case 'deleteHardLineBackward':
        event.preventDefault();
        host.remove(-1, 'line');
        return;
      case 'deleteSoftLineForward':
      case 'deleteHardLineForward':
        event.preventDefault();
        host.remove(1, 'line');
        return;
      case 'historyUndo':
        event.preventDefault();
        host.undo();
        return;
      case 'historyRedo':
        event.preventDefault();
        host.redo();
        return;
      case 'insertFromPaste':
      case 'insertFromDrop':
      case 'insertFromYank':
      case 'deleteByCut':
      case 'deleteByDrag':
        // The clipboard events do these; nothing is dropped into the field.
        event.preventDefault();
        return;
      default:
      // Replacements and whatever else the browser does on its own: seen in the field afterwards.
    }
  };

  private onInput = (): void => {
    if (this.destroyed) return;
    this.read();
    this.sync();
  };

  /** Compares the field with the mirror and hands the difference to the host. */
  private read(): void {
    const element = this.element;
    const now = element.value;
    const old = this.value;
    if (now === old) return;
    if (!this.host.context()) {
      this.value = now;
      return;
    }
    let caret = now.length;
    try {
      if (typeof element.selectionStart === 'number') caret = Math.max(0, Math.min(now.length, element.selectionEnd as number));
    } catch {
      // The end of the field then.
    }
    // What is the same at both ends did not change. The carets say where a change was when the text alone does not
    // ("aa" with one "a" removed): the part before both carets is compared first.
    const limit = Math.min(this.start, caret, old.length, now.length);
    let prefix = 0;
    while (prefix < limit && old.charCodeAt(prefix) === now.charCodeAt(prefix)) prefix++;
    const most = Math.min(old.length - Math.max(prefix, this.selected ? this.end : this.start), now.length - prefix);
    let suffix = 0;
    while (suffix < most && old.charCodeAt(old.length - 1 - suffix) === now.charCodeAt(now.length - 1 - suffix)) suffix++;
    const text = now.slice(prefix, now.length - suffix);
    const removedEnd = old.length - suffix;
    const change: InputChange = {
      removeBefore: Math.max(0, this.start - prefix),
      removeAfter: Math.max(0, removedEnd - this.end),
      replaceSelection: this.selected && removedEnd > this.start && prefix < this.end,
      text,
      caretBack: Math.max(0, Math.min(text.length, prefix + text.length - caret)),
      composing: 0,
    };
    if (this.composition !== undefined && this.composed !== '') {
      // The composed text ends at the caret as a rule; it is underlined as far as the field agrees.
      const length = this.composed.length;
      if (caret >= length && now.slice(caret - length, caret) === this.composed) change.composing = length;
    }
    this.value = now;
    this.start = this.end = caret;
    this.selected = false;
    this.host.change(change, this.composition);
  }
}
