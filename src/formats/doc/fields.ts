// Fields of a Word binary file -> the complex fields of the document model.
//
// In the text a field is three marks around its code and its result:
// 0x13 code 0x14 result 0x15 (the separator and the result may be missing).
// The story reader turns the marks into fieldChar items and leaves the code
// as text; here the code becomes instrText items, as in a docx file, and
// then the fields are assembled by the same function the docx reader uses.
//
// Marks that do not pair up (a begin without an end, a stray separator or
// end) are taken out, so that a damaged file cannot hide the text after them.

import type { Block, FieldCharItem, Run, RunItem } from '../../model/doc/types';
import { linkFields } from '../docx/body';

function eachRun(blocks: Block[], fn: (run: Run) => void): void {
  for (const block of blocks) {
    if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) eachRun(cell.blocks, fn);
    } else if (block.type === 'paragraph') {
      for (const inline of block.children) if (inline.type === 'run') fn(inline);
    }
  }
}

/** Turns the field marks and code text of a story into the fields of the model. */
export function linkDocFields(blocks: Block[]): void {
  // First pass: which marks pair up.
  const open: { begin: FieldCharItem; separator?: FieldCharItem }[] = [];
  const stray = new Set<RunItem>();
  let any = false;
  eachRun(blocks, (run) => {
    for (const item of run.content) {
      if (item.type !== 'fieldChar') continue;
      any = true;
      if (item.charType === 'begin') open.push({ begin: item });
      else if (item.charType === 'separate') {
        const top = open[open.length - 1];
        if (!top || top.separator) stray.add(item);
        else top.separator = item;
      } else if (!open.pop()) stray.add(item);
    }
  });
  if (!any) return;
  for (const field of open) {
    stray.add(field.begin);
    if (field.separator) stray.add(field.separator);
  }

  // Second pass: text in a field code is the instruction of the field.
  const stack: boolean[] = [];
  eachRun(blocks, (run) => {
    const content = run.content;
    let kept = 0;
    for (let i = 0; i < content.length; i++) {
      let item = content[i];
      if (item.type === 'fieldChar') {
        if (stray.has(item)) continue;
        if (item.charType === 'begin') stack.push(false);
        else if (item.charType === 'separate') stack[stack.length - 1] = true;
        else stack.pop();
      } else if (item.type === 'text' && stack.length && !stack[stack.length - 1]) item = { type: 'instrText', text: item.text };
      content[kept++] = item;
    }
    content.length = kept;
  });
  linkFields(blocks);
}
