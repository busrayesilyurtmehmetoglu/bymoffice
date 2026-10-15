// The panel of the fields of a pivot table: the fields of its source, and four boxes they stand in
// (filters, columns, rows, values). A field is put in and moved by dragging it into a box, taken
// out by unticking it or by the button of its entry, and every move can be made with the buttons
// and lists each entry has (for the keyboard and for fingers); a value is summarized
// by a function and shown with a number format; a field of the rows or columns has its subtotals
// and its order; the items of a field can be hidden. Every change is applied at once, as one undo
// step, and the panel shows the table as it is then. It stands beside the cells (at the right, or
// along the bottom on a narrow view) and leaves them usable.
//
// The part that works out a change of the table from what is done in the panel has no DOM in it
// and is tested on its own; the class below it builds the panel. What the file names (fields,
// items) is shown as text only.

import { PIVOT_AGGREGATES, type PivotAggregate } from '../model/sheet/pivot';
import type { PivotFieldsChange, PivotFieldsInfo, PivotSort, PivotValueChange } from '../ops/pivot-fields';
import type { PivotStrings } from './strings';

/** The boxes fields stand in. */
export type PivotBox = 'filters' | 'cols' | 'rows' | 'values';
export const PIVOT_BOXES: PivotBox[] = ['filters', 'cols', 'rows', 'values'];

/** The functions the panel offers a value to be summarized by (a value of the file may have another: it is kept). */
export const PANEL_AGGREGATES: PivotAggregate[] = ['sum', 'count', 'average', 'max', 'min'];

/** The number formats the panel offers; '' is that of the cells. */
export const PANEL_FORMATS = ['', '0', '#,##0', '0.00', '#,##0.00', '0%', '0.00%'];

/** An entry of a box: a field, at a place. */
export interface PivotSpot {
  box: PivotBox;
  at: number;
}

/** The fields of each box, from what the table has. */
export function boxesOf(info: PivotFieldsInfo): Record<PivotBox, number[]> {
  return { filters: info.filters.map((one) => one.field), cols: info.cols.slice(), rows: info.rows.slice(), values: info.values.map((one) => one.field) };
}

/** The values of a table as a change states them again: what is left out keeps what the value has (its name follows its function while it is Excel's). */
function valuesAsChange(info: PivotFieldsInfo): PivotValueChange[] {
  return info.values.map((one) => ({ field: one.field, subtotal: one.subtotal }));
}

/** The whole layout of a table as a change, with `boxes` in the place of its fields. */
function layoutChange(info: PivotFieldsInfo, boxes: Record<PivotBox, number[]>, values: PivotValueChange[]): PivotFieldsChange {
  // A filter keeps the item it shows.
  const filters = boxes.filters.map((field) => {
    const had = info.filters.find((one) => one.field === field);
    return had ? had : { field };
  });
  return { rows: boxes.rows, cols: boxes.cols, filters, values };
}

/**
 * The change that moves the entry at `from` (or puts the field `field` in, when `from` is
 * undefined) to the box `to` at the place `at` (the end by default); `to` null takes it out. A
 * field stands in one of the filters, the columns and the rows: put in another of them, it leaves
 * the one it was in. A field may stand in the values more than once.
 */
export function moveChange(info: PivotFieldsInfo, field: number, from: PivotSpot | undefined, to: PivotBox | null, at?: number): PivotFieldsChange {
  const boxes = boxesOf(info);
  let values = valuesAsChange(info);
  let moved: PivotValueChange | undefined;
  if (from) {
    if (from.box === 'values') {
      moved = values[from.at];
      values = values.filter((_, i) => i !== from.at);
      boxes.values = boxes.values.filter((_, i) => i !== from.at);
    } else boxes[from.box] = boxes[from.box].filter((_, i) => i !== from.at);
  }
  if (to !== null) {
    if (to !== 'values') {
      // Out of the other places of the axes.
      for (const box of ['filters', 'cols', 'rows'] as PivotBox[]) boxes[box] = boxes[box].filter((one) => one !== field);
    }
    const list = boxes[to];
    const place = at === undefined || at < 0 || at > list.length ? list.length : at;
    list.splice(place, 0, field);
    if (to === 'values') values.splice(place, 0, moved && from && from.box === 'values' ? moved : { field, subtotal: 'sum' });
  }
  return layoutChange(info, boxes, values);
}

/** The change that puts a field in (a field of numbers into the values, any other into the rows, as Excel does) or takes it out of every box. */
export function toggleChange(info: PivotFieldsInfo, field: number, on: boolean, numeric: boolean): PivotFieldsChange {
  if (on) return moveChange(info, field, undefined, numeric ? 'values' : 'rows');
  const boxes = boxesOf(info);
  for (const box of ['filters', 'cols', 'rows'] as PivotBox[]) boxes[box] = boxes[box].filter((one) => one !== field);
  const values = valuesAsChange(info).filter((one) => one.field !== field);
  return layoutChange(info, boxes, values);
}

/** The change of the function or the number format of the value at a place. */
export function valueChange(info: PivotFieldsInfo, at: number, subtotal: PivotAggregate | undefined, numFmt: string | undefined): PivotFieldsChange {
  const values = valuesAsChange(info).map((one, i) => {
    if (i !== at) return one;
    const next: PivotValueChange = { ...one };
    if (subtotal !== undefined) next.subtotal = subtotal;
    if (numFmt !== undefined) next.numFmt = numFmt;
    return next;
  });
  return { values };
}

/** Items a list of the panel shows at most. */
export const PANEL_ITEMS = 500;

/** What the list of a filter offers: all items, then the items; a long list is cut, but the item that is shown is always in it. */
export function filterChoices(items: { text: string }[], shown: number | undefined, all: string): [string, string][] {
  const out: [string, string][] = [['', all]];
  items.forEach((one, i) => {
    if (i < PANEL_ITEMS || i === shown) out.push([String(i), one.text]);
  });
  return out;
}

/** Whether a field is in any box. */
export function isUsed(info: PivotFieldsInfo, field: number): boolean {
  return info.rows.indexOf(field) >= 0 || info.cols.indexOf(field) >= 0 || info.filters.some((one) => one.field === field) || info.values.some((one) => one.field === field);
}

export interface PivotPanelOptions {
  parent: HTMLElement;
  strings: PivotStrings;
  /** Opened by a finger: controls are larger. */
  coarse?: boolean;
  /** The fields of the table as they are now; undefined when there is no such table any more (the panel closes). */
  info(): PivotFieldsInfo | undefined;
  /** Whether each field of the source holds numbers only (where it goes when it is ticked). */
  numeric(field: number): boolean;
  /** Applies a change; what it throws is shown in the panel. */
  onChange(change: PivotFieldsChange): void;
  /** Puts an error into the words of the view. */
  explain(error: unknown): string;
  onClose(): void;
}

let nextId = 0;

/** The panel of the fields of a pivot table. */
export class PivotPanel {
  readonly element: HTMLDivElement;
  private readonly options: PivotPanelOptions;
  private readonly content: HTMLDivElement;
  private readonly alert: HTMLDivElement;
  private info: PivotFieldsInfo | undefined;
  /** The field whose items are listed (to hide some); undefined: none. */
  private itemsOf: number | undefined;
  /** The entry that is dragged. */
  private dragged: { field: number; from?: PivotSpot } | undefined;
  private closed = false;

  constructor(options: PivotPanelOptions) {
    this.options = options;
    const strings = options.strings;
    const element = document.createElement('div');
    element.className = 'bym-pivot-panel';
    element.setAttribute('role', 'complementary');
    element.setAttribute('data-bym-panel', 'pivot');
    // The panel itself takes the focus when the control that had it is gone.
    element.tabIndex = -1;
    const title = document.createElement('div');
    title.className = 'bym-pivot-title';
    title.id = `bym-pivot-title-${nextId++}`;
    title.textContent = strings.panelTitle;
    element.setAttribute('aria-labelledby', title.id);
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'bym-pivot-close';
    close.textContent = '×';
    close.setAttribute('aria-label', strings.close);
    close.title = strings.close;
    close.setAttribute('data-bym-action', 'close');
    close.addEventListener('click', () => this.close());
    title.appendChild(close);
    const content = document.createElement('div');
    content.className = 'bym-pivot-content';
    const alert = document.createElement('div');
    alert.className = 'bym-dialog-alert';
    alert.setAttribute('role', 'alert');
    element.appendChild(title);
    element.appendChild(content);
    element.appendChild(alert);
    if (options.coarse) element.classList.add('bym-pivot-coarse');
    this.element = element;
    this.content = content;
    this.alert = alert;
    element.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.close();
      }
      // What is typed here is not the view's.
      event.stopPropagation();
    });
    for (const type of ['keyup', 'keypress', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input']) {
      element.addEventListener(type, (event) => event.stopPropagation());
    }
    options.parent.appendChild(element);
    this.refresh();
    const first = (content.querySelector('input:not(:disabled),button:not(:disabled),select:not(:disabled)') || close) as HTMLElement;
    if (first) first.focus({ preventScroll: true });
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  /** Shows what went wrong; '' clears it. */
  showError(message: string): void {
    this.alert.textContent = message;
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    this.options.onClose();
  }

  /** Shows the table as it is now; closes the panel when there is no table any more. The control that had the focus keeps it where it can. */
  refresh(): void {
    if (this.closed) return;
    const info = this.options.info();
    if (!info) {
      this.close();
      return;
    }
    this.info = info;
    const active = this.element.ownerDocument ? (this.element.ownerDocument.activeElement as HTMLElement | null) : null;
    const key = active && this.element.contains(active) ? active.getAttribute('data-bym-key') : null;
    while (this.content.firstChild) this.content.removeChild(this.content.firstChild);
    this.build(info);
    if (key) {
      const again = this.content.querySelector(`[data-bym-key="${key.replace(/["\\]/g, '')}"]`) as HTMLElement | null;
      if (again && !(again as HTMLButtonElement).disabled) again.focus({ preventScroll: true });
      else this.element.focus();
    }
  }

  /** Applies a change and shows the table as it comes out (or what went wrong). */
  private change(change: PivotFieldsChange): void {
    this.showError('');
    try {
      this.options.onChange(change);
    } catch (error) {
      this.showError(this.options.explain(error));
    }
    this.refresh();
  }

  private button(text: string, label: string, key: string, run: () => void): HTMLButtonElement {
    const el = document.createElement('button');
    el.type = 'button';
    el.className = 'bym-pivot-button';
    el.textContent = text;
    el.setAttribute('aria-label', label);
    el.title = label;
    el.setAttribute('data-bym-key', key);
    el.addEventListener('click', run);
    return el;
  }

  private select(items: [string, string][], value: string, label: string, key: string, run: (value: string) => void): HTMLSelectElement {
    const el = document.createElement('select');
    el.className = 'bym-dialog-input bym-dialog-select bym-pivot-select';
    for (const [option, text] of items) {
      const item = document.createElement('option');
      item.value = option;
      item.textContent = text;
      el.appendChild(item);
    }
    el.value = value;
    el.setAttribute('aria-label', label);
    el.title = label;
    el.setAttribute('data-bym-key', key);
    el.addEventListener('change', () => run(el.value));
    return el;
  }

  private build(info: PivotFieldsInfo): void {
    const strings = this.options.strings;
    if (info.refusal !== undefined) {
      const note = document.createElement('div');
      note.className = 'bym-dialog-note';
      note.textContent = this.options.explain(new RangeError(info.refusal));
      this.content.appendChild(note);
    }
    const locked = info.refusal !== undefined;
    // The fields of the source, each with a box to tick.
    const list = document.createElement('div');
    list.className = 'bym-pivot-fields';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', strings.fieldList);
    for (const field of info.fields) {
      const row = document.createElement('div');
      row.className = 'bym-pivot-field';
      row.setAttribute('data-bym-field', String(field.index));
      const label = document.createElement('label');
      label.className = 'bym-dialog-check bym-pivot-check';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.className = 'bym-dialog-box';
      box.checked = isUsed(info, field.index);
      box.disabled = locked;
      box.setAttribute('data-bym-key', `use:${field.index}`);
      box.addEventListener('change', () => this.change(toggleChange(info, field.index, box.checked, this.options.numeric(field.index))));
      const name = document.createElement('span');
      name.className = 'bym-pivot-name';
      name.textContent = field.name;
      label.appendChild(box);
      label.appendChild(name);
      row.appendChild(label);
      this.draggable(row, field.index, undefined, locked);
      if (field.items && !locked) {
        const open = this.itemsOf === field.index;
        const toggle = this.button('▾', strings.filterItems(field.name), `items:${field.index}`, () => {
          this.itemsOf = open ? undefined : field.index;
          this.refresh();
        });
        toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
        row.appendChild(toggle);
      }
      list.appendChild(row);
      if (this.itemsOf === field.index && field.items && !locked) list.appendChild(this.itemList(info, field.index));
    }
    this.content.appendChild(list);

    // The four boxes.
    const grid = document.createElement('div');
    grid.className = 'bym-pivot-boxes';
    for (const box of PIVOT_BOXES) grid.appendChild(this.boxElement(info, box, locked));
    this.content.appendChild(grid);

    // Grand totals.
    const totals = document.createElement('div');
    totals.className = 'bym-pivot-totals';
    const check = (text: string, checked: boolean, key: string, run: (on: boolean) => void) => {
      const label = document.createElement('label');
      label.className = 'bym-dialog-check';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.className = 'bym-dialog-box';
      input.checked = checked;
      input.disabled = locked;
      input.setAttribute('data-bym-key', key);
      input.addEventListener('change', () => run(input.checked));
      const span = document.createElement('span');
      span.textContent = text;
      label.appendChild(input);
      label.appendChild(span);
      totals.appendChild(label);
    };
    check(strings.rowGrandTotals, info.rowGrandTotals, 'rowGrandTotals', (on) => this.change({ rowGrandTotals: on }));
    check(strings.colGrandTotals, info.colGrandTotals, 'colGrandTotals', (on) => this.change({ colGrandTotals: on }));
    this.content.appendChild(totals);
  }

  /** The items of a field, each with a box: unticked items are hidden. */
  private itemList(info: PivotFieldsInfo, field: number): HTMLElement {
    const strings = this.options.strings;
    const items = info.fields[field].items || [];
    const list = document.createElement('div');
    list.className = 'bym-pivot-items';
    list.setAttribute('role', 'group');
    list.setAttribute('aria-label', strings.filterItems(info.fields[field].name));
    // A long list is shown in part: hiding item by item is not what a list of thousands is for.
    const shown = Math.min(items.length, PANEL_ITEMS);
    for (let i = 0; i < shown; i++) {
      const label = document.createElement('label');
      label.className = 'bym-dialog-check';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.className = 'bym-dialog-box';
      box.checked = !items[i].hidden;
      box.setAttribute('data-bym-key', `item:${field}:${i}`);
      box.addEventListener('change', () => {
        const hidden: number[] = [];
        items.forEach((item, k) => {
          if (k === i ? !box.checked : item.hidden) hidden.push(k);
        });
        this.change({ hidden: [{ field, items: hidden }] });
      });
      const text = document.createElement('span');
      text.textContent = items[i].text;
      label.appendChild(box);
      label.appendChild(text);
      list.appendChild(label);
    }
    if (items.length > shown) {
      const more = document.createElement('div');
      more.className = 'bym-dialog-hint';
      more.textContent = strings.moreItems(items.length - shown);
      list.appendChild(more);
    }
    return list;
  }

  private boxElement(info: PivotFieldsInfo, box: PivotBox, locked: boolean): HTMLElement {
    const strings = this.options.strings;
    const fields = boxesOf(info)[box];
    const element = document.createElement('div');
    element.className = 'bym-pivot-box';
    element.setAttribute('data-bym-box', box);
    const heading = document.createElement('div');
    heading.className = 'bym-pivot-box-title';
    heading.id = `bym-pivot-box-${nextId++}`;
    heading.textContent = strings.boxes[box];
    element.appendChild(heading);
    const list = document.createElement('ul');
    list.className = 'bym-pivot-list';
    list.setAttribute('aria-labelledby', heading.id);
    element.appendChild(list);
    if (fields.length === 0) {
      const empty = document.createElement('li');
      empty.className = 'bym-dialog-hint bym-pivot-empty';
      empty.textContent = strings.dropHere;
      list.appendChild(empty);
    }
    fields.forEach((field, at) => {
      const name = info.fields[field] ? info.fields[field].name : '';
      const shown = box === 'values' ? info.values[at].name : name;
      const item = document.createElement('li');
      item.className = 'bym-pivot-entry';
      item.setAttribute('data-bym-entry', `${box}:${at}`);
      const text = document.createElement('span');
      text.className = 'bym-pivot-name';
      text.textContent = shown;
      item.appendChild(text);
      const tools = document.createElement('span');
      tools.className = 'bym-pivot-tools';
      const from: PivotSpot = { box, at };
      const up = this.button('↑', strings.moveUp(shown), `up:${box}:${at}`, () => this.change(moveChange(info, field, from, box, at - 1)));
      up.disabled = locked || at === 0;
      const down = this.button('↓', strings.moveDown(shown), `down:${box}:${at}`, () => this.change(moveChange(info, field, from, box, at + 1)));
      down.disabled = locked || at === fields.length - 1;
      const to = this.select(
        PIVOT_BOXES.map((one) => [one, strings.boxes[one]] as [string, string]),
        box,
        strings.moveTo(shown),
        `to:${box}:${at}`,
        (value) => this.change(moveChange(info, field, from, value as PivotBox)),
      );
      to.disabled = locked;
      const remove = this.button('×', strings.remove(shown), `remove:${box}:${at}`, () => this.change(moveChange(info, field, from, null)));
      remove.disabled = locked;
      tools.appendChild(up);
      tools.appendChild(down);
      tools.appendChild(to);
      tools.appendChild(remove);
      item.appendChild(tools);
      const more = document.createElement('div');
      more.className = 'bym-pivot-more';
      if (box === 'values') {
        const value = info.values[at];
        const functions: PivotAggregate[] = PANEL_AGGREGATES.indexOf(value.subtotal) >= 0 ? PANEL_AGGREGATES : PANEL_AGGREGATES.concat([value.subtotal]);
        const fn = this.select(
          functions.map((one) => [one, strings.aggregates[one] || one] as [string, string]),
          value.subtotal,
          strings.summarize(shown),
          `fn:${at}`,
          (picked) => {
            if ((PIVOT_AGGREGATES as readonly string[]).indexOf(picked) >= 0) this.change(valueChange(info, at, picked as PivotAggregate, undefined));
          },
        );
        fn.disabled = locked;
        const own = value.numFmt === undefined ? '' : value.numFmt;
        const formats = PANEL_FORMATS.indexOf(own) >= 0 ? PANEL_FORMATS : PANEL_FORMATS.concat([own]);
        const format = this.select(
          formats.map((one) => [one, one === '' ? strings.generalFormat : one] as [string, string]),
          own,
          strings.numberFormat(shown),
          `fmt:${at}`,
          (picked) => this.change(valueChange(info, at, undefined, picked)),
        );
        format.disabled = locked;
        more.appendChild(fn);
        more.appendChild(format);
      } else if (box === 'rows' || box === 'cols') {
        const field2 = info.fields[field];
        const label = document.createElement('label');
        label.className = 'bym-dialog-check';
        const sub = document.createElement('input');
        sub.type = 'checkbox';
        sub.className = 'bym-dialog-box';
        sub.checked = field2.subtotal;
        sub.disabled = locked;
        sub.setAttribute('data-bym-key', `sub:${field}`);
        sub.setAttribute('aria-label', strings.subtotals(name));
        sub.addEventListener('change', () => this.change({ subtotals: [{ field, on: sub.checked }] }));
        const span = document.createElement('span');
        span.textContent = strings.subtotal;
        label.appendChild(sub);
        label.appendChild(span);
        more.appendChild(label);
        const sort = this.select(
          (['manual', 'ascending', 'descending'] as PivotSort[]).map((one) => [one, strings.sorts[one]] as [string, string]),
          field2.sort,
          strings.sort(name),
          `sort:${field}`,
          (picked) => this.change({ sort: [{ field, order: picked as PivotSort }] }),
        );
        sort.disabled = locked;
        more.appendChild(sort);
      } else if (box === 'filters') {
        const field2 = info.fields[field];
        const items = field2.items || [];
        const filter = info.filters[at];
        const choices = filterChoices(items, filter.item, strings.allItems);
        const show = this.select(choices, filter.item === undefined ? '' : String(filter.item), strings.shownItem(name), `show:${field}`, (picked) => {
          const filters = info.filters.map((one) => (one.field === field ? (picked === '' ? { field, item: undefined } : { field, item: Number(picked) }) : one));
          this.change({ filters });
        });
        show.disabled = locked;
        more.appendChild(show);
      }
      if (more.firstChild) item.appendChild(more);
      this.draggable(item, field, from, locked);
      list.appendChild(item);
    });
    // Dropped on a box: there, at the end.
    if (!locked) {
      element.addEventListener('dragover', (event) => {
        if (!this.dragged) return;
        event.preventDefault();
        if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
        element.classList.add('bym-pivot-over');
      });
      element.addEventListener('dragleave', () => element.classList.remove('bym-pivot-over'));
      element.addEventListener('drop', (event) => {
        element.classList.remove('bym-pivot-over');
        const dragged = this.dragged;
        this.dragged = undefined;
        if (!dragged || !this.info) return;
        event.preventDefault();
        // Dropped on an entry: before it.
        const target = (event.target as HTMLElement).closest ? ((event.target as HTMLElement).closest('[data-bym-entry]') as HTMLElement | null) : null;
        const spot = target ? target.getAttribute('data-bym-entry') : null;
        let at = spot && spot.split(':')[0] === box ? Number(spot.split(':')[1]) : undefined;
        // Before an entry further down its own box: counted without the entry that moves.
        if (at !== undefined && dragged.from && dragged.from.box === box && dragged.from.at < at) at--;
        this.change(moveChange(this.info, dragged.field, dragged.from, box, at));
      });
    }
    return element;
  }

  /** Makes an element one that is dragged to a box: an entry of the list of fields, or of a box (it moves to the box it is dropped on). */
  private draggable(element: HTMLElement, field: number, from: PivotSpot | undefined, locked: boolean): void {
    if (locked) return;
    element.draggable = true;
    element.addEventListener('dragstart', (event) => {
      this.dragged = { field, from };
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = 'move';
        // Some browsers start no drag without data.
        event.dataTransfer.setData('text/plain', '');
      }
    });
    element.addEventListener('dragend', () => {
      this.dragged = undefined;
    });
  }
}

/** The cosmetic rules of the panel; it uses those of the dialogs (DIALOG_STYLE) as well. */
export const PIVOT_PANEL_STYLE =
  '.bym-pivot-panel{position:absolute;top:0;right:0;bottom:0;z-index:25;display:flex;flex-direction:column;box-sizing:border-box;width:300px;max-width:100%;' +
  'background:#fff;color:#202124;border-left:1px solid #dadce0;box-shadow:-4px 0 12px rgba(0,0,0,.12);font:13px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;text-align:left;direction:ltr}' +
  // A narrow view has the panel along its bottom, as a drawer.
  '@media (max-width:600px){.bym-pivot-panel{top:auto;left:0;width:auto;height:55%;border-left:0;border-top:1px solid #dadce0;box-shadow:0 -4px 12px rgba(0,0,0,.12)}}' +
  '.bym-pivot-title{display:flex;align-items:center;justify-content:space-between;flex:none;padding:10px 12px;font-size:15px;font-weight:600;border-bottom:1px solid #dadce0}' +
  '.bym-pivot-content{flex:1 1 auto;min-height:0;overflow:auto;overscroll-behavior:contain;padding:6px 12px}' +
  '.bym-pivot-close,.bym-pivot-button{-webkit-appearance:none;appearance:none;flex:none;box-sizing:border-box;min-width:24px;height:24px;margin:0 2px;padding:0 4px;border:1px solid #dadce0;border-radius:4px;background:#fff;color:#202124;font:inherit;line-height:1;cursor:pointer}' +
  '.bym-pivot-button:disabled{opacity:.38;cursor:default}' +
  '.bym-pivot-close:focus,.bym-pivot-button:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-pivot-field{display:flex;align-items:center;justify-content:space-between}' +
  '.bym-pivot-check{flex:1 1 auto;min-width:0;margin:3px 0}' +
  '.bym-pivot-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}' +
  '.bym-pivot-items{margin:0 0 6px 24px;max-height:200px;overflow:auto}' +
  '.bym-pivot-boxes{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:10px 0}' +
  '.bym-pivot-box{min-width:0;min-height:60px;padding:4px 6px;border:1px solid #dadce0;border-radius:6px}' +
  '.bym-pivot-over{border-color:#1a7f45;background:#e6f4ea}' +
  '.bym-pivot-box-title{font-size:12px;font-weight:600;color:#5f6368;margin:0 0 4px}' +
  '.bym-pivot-list{list-style:none;margin:0;padding:0}' +
  '.bym-pivot-entry{display:flex;flex-wrap:wrap;align-items:center;margin:0 0 4px;padding:3px 4px;border:1px solid #e0e0e0;border-radius:4px;background:#f8f9fa;cursor:grab}' +
  '.bym-pivot-entry .bym-pivot-name{flex:1 1 100%}' +
  '.bym-pivot-tools{display:flex;align-items:center;flex-wrap:wrap}' +
  '.bym-pivot-select{flex:none;height:24px;max-width:100%;font-size:12px;padding:0 20px 0 4px}' +
  '.bym-pivot-more{display:flex;flex-wrap:wrap;align-items:center;flex:1 1 100%;gap:4px;margin-top:2px}' +
  '.bym-pivot-more .bym-dialog-check{margin:0}' +
  '.bym-pivot-totals{margin:4px 0 8px}' +
  // A finger needs more to hit than a mouse pointer.
  '.bym-pivot-coarse .bym-pivot-button,.bym-pivot-coarse .bym-pivot-close,.bym-pivot-coarse .bym-pivot-select{height:40px;min-width:40px}' +
  '@media (pointer:coarse){.bym-pivot-button,.bym-pivot-close,.bym-pivot-select{height:40px;min-width:40px}}';
