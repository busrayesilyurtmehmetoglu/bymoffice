// One flow of text while it is read: the body, a header, a note, the text of
// a text box. It gathers characters into runs, runs into paragraphs and,
// with the table assembler, paragraphs into blocks.

import type { Block, Inline, ParaProps, Paragraph, Run, RunItem, RunProps, Wrapper } from '../../model/doc/types';
import { RowState, TableAssembler, type TableBudget } from './tables';
import type { FormatContext } from './state';

export type StoryKind = 'body' | 'header' | 'footnote' | 'comment' | 'textbox';

/** What the text being added is: shown text, the code of a field, text of a tracked deletion, or the code of a field in one. */
export const TEXT_SHOWN = 0;
export const TEXT_CODE = 1;
export const TEXT_DELETED = 2;
export const TEXT_CODE_DELETED = 3;

const TEXT_TYPES = ['text', 'instrText', 'deletedText', 'instrText'] as const;

/** Wrappers around one inline that are followed. */
const MAX_CHAIN = 64;

/**
 * Gives every stretch of inlines that follow each other in a tracked change a wrapper object of its own, as the
 * elements of a docx file are: the reader hands out one object for a kind, an author and a date, wherever in the
 * file they come, and two changes that stand apart must not be one object (resolving one would resolve the
 * other). A deletion inside an insertion keeps the insertion of the inlines around it.
 */
export function separateRevisions(children: Inline[]): void {
  /** For the wrappers around the inline before: the object that stands for each in this stretch. */
  let open = new Map<Wrapper, Wrapper>();
  for (const inline of children) {
    if (!inline.wrap) {
      if (open.size) open = new Map();
      continue;
    }
    const chain: Wrapper[] = [];
    for (let w: Wrapper | undefined = inline.wrap; w && chain.length < MAX_CHAIN; w = w.parent) chain.push(w);
    const next = new Map<Wrapper, Wrapper>();
    let parent: Wrapper | undefined;
    // Outermost first: a wrapper goes on only inside the wrapper it was in.
    for (let k = chain.length - 1; k >= 0; k--) {
      const source = chain[k];
      const known = open.get(source);
      let own: Wrapper;
      if (known && known.parent === parent) own = known;
      else {
        own = { ...source };
        if (parent) own.parent = parent;
        else delete own.parent;
      }
      next.set(source, own);
      parent = own;
    }
    inline.wrap = parent;
    open = next;
  }
}

export class Story {
  readonly assembler: TableAssembler;
  readonly row: RowState;
  /** The children of the paragraph being read. */
  children: Inline[] = [];
  /** For each run of the paragraph being read, the run properties in full it was made from. */
  fulls: RunProps[] = [];
  runs: Run[] = [];
  /** Id of the paragraph style the runs of the paragraph were made for; undefined before the first run. */
  styleUsed: string | undefined;
  hasRuns = false;
  /** The text a writer puts in front of a numbered paragraph for readers that know no lists. */
  listText: Inline[] | undefined;
  /** A note mark ("\chftn") that waits for its note: where it is. */
  pendingNote: { run: Run; index: number } | undefined;
  pendingComment: { run: Run; index: number } | undefined;
  /** "\ftnalt": the note being read is an endnote. */
  endnote = false;
  /** The automatic marks inside the note being read. */
  noteRefs: { run: Run; index: number }[] = [];
  /** "\atnref": what ties the comment being read to the marks of the text it is about. */
  commentRef: string | undefined;
  paragraphs = 0;

  private run: Run | undefined;
  private text = '';
  private textKind = TEXT_SHOWN;
  private saved: { children: Inline[]; fulls: RunProps[]; runs: Run[]; styleUsed: string | undefined; hasRuns: boolean } | undefined;

  constructor(
    readonly kind: StoryKind,
    ctx: FormatContext,
    budget?: TableBudget,
  ) {
    this.row = new RowState(ctx);
    this.assembler = new TableAssembler(budget);
  }

  private flushText(): void {
    if (!this.text) return;
    const run = this.run!;
    if (this.textKind === TEXT_CODE_DELETED) run.content.push({ type: 'instrText', text: this.text, deleted: true });
    else run.content.push({ type: TEXT_TYPES[this.textKind], text: this.text } as RunItem);
    this.text = '';
  }

  /** Ends the run being read: what follows starts a new one. */
  endRun(): void {
    const run = this.run;
    if (!run) return;
    this.flushText();
    // An array that grew by one item at a time holds room for many more; the model keeps one of the right size.
    run.content = run.content.slice();
    this.run = undefined;
  }

  private open(props: RunProps, full: RunProps, wrap: Wrapper | undefined): Run {
    let run = this.run;
    if (run && run.props === props && run.wrap === wrap) return run;
    this.endRun();
    run = { type: 'run', props, content: [] };
    if (wrap) run.wrap = wrap;
    this.run = run;
    this.children.push(run);
    this.runs.push(run);
    this.fulls.push(full);
    this.hasRuns = true;
    return run;
  }

  addText(text: string, kind: number, props: RunProps, full: RunProps, wrap: Wrapper | undefined): void {
    this.open(props, full, wrap);
    if (kind !== this.textKind) {
      this.flushText();
      this.textKind = kind;
    }
    this.text += text;
  }

  /** Adds an item to the run being read; returns where it is, for an item that is changed later. */
  addItem(item: RunItem, props: RunProps, full: RunProps, wrap: Wrapper | undefined): { run: Run; index: number } {
    const run = this.open(props, full, wrap);
    this.flushText();
    run.content.push(item);
    return { run, index: run.content.length - 1 };
  }

  /** Adds an item in a run of its own, as a field mark has it in a docx file. */
  addAlone(item: RunItem, props: RunProps, full: RunProps, wrap?: Wrapper): Run {
    this.endRun();
    const run = this.open(props, full, wrap);
    run.content.push(item);
    this.endRun();
    return run;
  }

  addInline(inline: Inline): void {
    this.endRun();
    this.children.push(inline);
  }

  /** What follows is the text in front of a numbered paragraph, kept apart until the paragraph says whether it is numbered. */
  beginCapture(): void {
    this.endRun();
    if (this.saved) return;
    this.saved = { children: this.children, fulls: this.fulls, runs: this.runs, styleUsed: this.styleUsed, hasRuns: this.hasRuns };
    this.children = [];
    this.fulls = [];
    this.runs = [];
    this.styleUsed = undefined;
    this.hasRuns = false;
  }

  endCapture(): void {
    this.endRun();
    const saved = this.saved;
    if (!saved) return;
    this.listText = this.children;
    this.children = saved.children;
    this.fulls = saved.fulls;
    this.runs = saved.runs;
    this.styleUsed = saved.styleUsed;
    this.hasRuns = saved.hasRuns;
    this.saved = undefined;
  }

  /** Ends the paragraph being read and hands it out; the next one starts empty. */
  takeParagraph(props: ParaProps, keepListText: boolean): Paragraph {
    this.endRun();
    if (this.saved) this.endCapture();
    let children = this.children;
    if (this.listText) {
      if (keepListText && this.listText.length) children = this.listText.concat(children);
      this.listText = undefined;
    }
    this.children = [];
    this.fulls = [];
    this.runs = [];
    this.styleUsed = undefined;
    this.hasRuns = false;
    this.paragraphs++;
    separateRevisions(children);
    return { type: 'paragraph', props, children: children.length ? children.slice() : [] };
  }

  finish(): Block[] {
    return this.assembler.finish();
  }
}
