// The dialog of a chart: its type, title, legend, axes, the range of cells it shows and its
// series one by one. Like the dialogs of rules it produces a patch: what was only looked at
// is not changed, so a dialog that is confirmed as it was opened makes no step, and a chart
// of the file keeps everything the dialog does not show.
//
// The part that works out the fields from a chart and the patch from the fields has no DOM
// in it and is tested on its own; the class below it builds the form.

import type { FormatLocale } from '../format/locale';
import { quoteSheetName } from '../calc/refs';
import { CHART_TYPES, type ChartInfo, type ChartPatch, type ChartSeriesEdit, type ChartType } from '../ops/charts';
import { Dialog, FieldError, parseRangeList, rangeListText, type DialogPointer } from './rules-dialog';
import type { ChartStrings } from './strings';

/** One series as the dialog shows it. */
export interface ChartSeriesFields {
  /** The place the series had in the chart when the dialog was opened; undefined for one added in the dialog. */
  from?: number;
  /** The name as a text, or "=" and the cell it is read from. */
  name: string;
  values: string;
  /** "#rrggbb", or '' for the colour the chart gives by itself. */
  color: string;
}

/** What the fields of the dialog hold. */
export interface ChartFields {
  /** One of CHART_TYPES; '' for a chart of a type the dialog does not offer. */
  type: string;
  grouping: string;
  title: string;
  /** The block of cells the series are made of ("A1:C5"); '' when they are not one block. */
  range: string;
  /** "auto", "columns" or "rows". */
  seriesIn: string;
  legend: string;
  dataLabels: boolean;
  categoryTitle: string;
  valueTitle: string;
  minimum: string;
  maximum: string;
  /** The categories every series shares; '' when there are none or they differ. */
  categories: string;
  series: ChartSeriesFields[];
}

/** The fields a chart is shown with. `sheetName` is the sheet the chart is on: references to its cells are shown without it. */
export function chartFields(info: ChartInfo, locale: FormatLocale, sheetName: string): ChartFields {
  const own = `${quoteSheetName(sheetName)}!`;
  const shown = (ref: string | undefined): string => (ref === undefined ? '' : ref.slice(0, own.length) === own && ref.indexOf(',') < 0 ? ref.slice(own.length) : ref);
  const number = (value: number | undefined): string => (value === undefined ? '' : String(value).replace('.', locale.decimal));
  const first = info.series[0];
  const shared = first && info.series.every((series) => series.categories === first.categories) ? shown(first.categories) : '';
  return {
    type: info.type || '',
    grouping: info.grouping,
    title: info.title || '',
    range: info.range ? rangeListText([info.range], locale) : '',
    seriesIn: info.range && info.seriesIn ? info.seriesIn : 'auto',
    legend: info.legend,
    dataLabels: info.dataLabels,
    categoryTitle: info.axisTitles.category || '',
    valueTitle: info.axisTitles.value || '',
    minimum: number(info.axisBounds.min),
    maximum: number(info.axisBounds.max),
    categories: shared,
    series: info.series.map((series, i) => ({
      from: i,
      name: series.nameRef !== undefined ? `=${shown(series.nameRef)}` : series.name || '',
      values: shown(series.values),
      color: series.color ? series.color.toLowerCase() : '',
    })),
  };
}

const PLAIN_NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/** What the dialog needs words for when it says what is wrong. */
export type ChartFieldStrings = Pick<ChartStrings, 'badRange' | 'badNumber' | 'needValues' | 'needRange'>;

function numberOf(text: string, locale: FormatLocale, strings: ChartFieldStrings, field: string): number | null {
  const typed = text.trim();
  if (typed === '') return null;
  const plain = locale.decimal === '.' ? typed : typed.split(locale.group || '\u0000').join('').replace(locale.decimal, '.');
  if (!PLAIN_NUMBER.test(plain) || !Number.isFinite(Number(plain))) throw new FieldError(strings.badNumber, field);
  return Number(plain);
}

/** A reference as a field has it, as the commands take it: without a leading "=". */
const refOf = (text: string): string => text.trim().replace(/^=\s*/, '');

/**
 * What the fields ask to change in the chart they were opened with (`initial`): nothing for
 * fields that were left as they were. Throws FieldError for what cannot be read (a range that
 * is none, a number that is none, a series without values), naming the field.
 */
export function chartPatch(initial: ChartFields, fields: ChartFields, locale: FormatLocale, strings: ChartFieldStrings): ChartPatch {
  const patch: ChartPatch = {};
  if (fields.type !== initial.type && CHART_TYPES.indexOf(fields.type as ChartType) >= 0) patch.type = fields.type as ChartType;
  if (fields.grouping !== initial.grouping && (fields.grouping === 'standard' || fields.grouping === 'stacked' || fields.grouping === 'percentStacked')) patch.grouping = fields.grouping;
  if (fields.title !== initial.title) patch.title = fields.title.trim() === '' ? null : fields.title;
  if (fields.legend !== initial.legend) patch.legend = fields.legend as ChartPatch['legend'];
  if (fields.dataLabels !== initial.dataLabels) patch.dataLabels = fields.dataLabels;
  const titles: NonNullable<ChartPatch['axisTitles']> = {};
  if (fields.categoryTitle !== initial.categoryTitle) titles.category = fields.categoryTitle.trim() === '' ? null : fields.categoryTitle;
  if (fields.valueTitle !== initial.valueTitle) titles.value = fields.valueTitle.trim() === '' ? null : fields.valueTitle;
  if (Object.keys(titles).length > 0) patch.axisTitles = titles;
  const bounds: NonNullable<ChartPatch['axisBounds']> = {};
  if (fields.minimum.trim() !== initial.minimum.trim()) bounds.min = numberOf(fields.minimum, locale, strings, 'minimum');
  if (fields.maximum.trim() !== initial.maximum.trim()) bounds.max = numberOf(fields.maximum, locale, strings, 'maximum');
  if (Object.keys(bounds).length > 0) patch.axisBounds = bounds;

  // The series anew from a block of cells, when the block or its direction is another one.
  const ranged = fields.range.trim() !== initial.range.trim() || fields.seriesIn !== initial.seriesIn;
  if (ranged && fields.range.trim() !== '') {
    const ranges = parseRangeList(fields.range, locale);
    if (!ranges || ranges.length !== 1) throw new FieldError(strings.badRange, 'range');
    patch.range = ranges[0];
    if (fields.seriesIn === 'columns' || fields.seriesIn === 'rows') patch.seriesIn = fields.seriesIn;
    return patch;
  }
  // Otherwise the series one by one, as far as the list or the shared categories changed.
  const categories = fields.categories.trim() !== initial.categories.trim();
  const sameList = fields.series.length === initial.series.length && fields.series.every((series, i) => series.from === i && series.name === initial.series[i].name && series.values === initial.series[i].values && series.color === initial.series[i].color);
  if (sameList && !categories) return patch;
  if (fields.series.length === 0) throw new FieldError(strings.needValues, 'series');
  patch.series = fields.series.map((series, i) => {
    const was = series.from !== undefined ? initial.series[series.from] : undefined;
    const edit: ChartSeriesEdit = {};
    if (series.from !== undefined && was) edit.from = series.from;
    if (!was || series.name !== was.name) {
      const name = series.name.trim();
      if (name.charAt(0) === '=') edit.nameRef = refOf(name);
      else if (name === '') {
        edit.name = null;
        edit.nameRef = null;
      } else edit.name = series.name;
    }
    if (!was || series.values.trim() !== was.values.trim()) {
      if (refOf(series.values) === '') throw new FieldError(strings.needValues, `values${i}`);
      edit.values = refOf(series.values);
    }
    if (!was || series.color !== was.color) edit.color = series.color === '' ? null : series.color;
    if (categories) edit.categories = refOf(fields.categories) === '' ? null : refOf(fields.categories);
    return edit;
  });
  return patch;
}

export interface ChartDialogOptions {
  parent: HTMLElement;
  strings: ChartStrings;
  locale: FormatLocale;
  fields: ChartFields;
  /** "all": every field; "parts": every field but the type. */
  editable: 'all' | 'parts';
  /** Whether a series can be added, and whether series can be taken away or put in another order (ChartInfo); both by default. */
  canAddSeries?: boolean;
  canReorderSeries?: boolean;
  coarse?: boolean;
  /** Lets cells be pointed at for the range and for the series. */
  pointer?: DialogPointer;
  /** Applies what the dialog asks for; what it throws is shown in the dialog, which then stays open. */
  onApply(patch: ChartPatch): void;
  /** Puts an error into the words of the view. */
  explain(error: unknown): string;
  onClose(): void;
}

/** The dialog that changes a chart. */
export class ChartDialog {
  readonly dialog: Dialog;
  private readonly options: ChartDialogOptions;
  private readonly initial: ChartFields;
  private readonly controls: { [name: string]: HTMLInputElement | HTMLSelectElement } = {};
  private readonly list: HTMLDivElement;
  private readonly axes: HTMLFieldSetElement;
  private rows: { from?: number; name: HTMLInputElement; values: HTMLInputElement; own: HTMLInputElement; color: HTMLInputElement }[] = [];

  constructor(options: ChartDialogOptions) {
    this.options = options;
    this.initial = options.fields;
    const strings = options.strings;
    const fields = options.fields;
    const dialog = new Dialog({ parent: options.parent, title: strings.dialogTitle, name: 'chart', onClose: options.onClose, coarse: options.coarse, pointer: options.pointer });
    this.dialog = dialog;
    const body = dialog.body;

    const types: [string, string][] = CHART_TYPES.map((type) => [type, strings.types[type] || type]);
    if (fields.type === '') types.unshift(['', strings.otherType]);
    const type = dialog.select(types);
    type.value = fields.type;
    if (options.editable !== 'all' || fields.type === '') {
      // The type of a chart of the file stays: said where the list is.
      type.disabled = true;
      type.title = strings.typeFixed;
    }
    dialog.addField(body, strings.type, type, 'type');
    const grouping = dialog.select((['standard', 'stacked', 'percentStacked'] as const).map((key) => [key, strings.groupings[key] || key]));
    grouping.value = fields.grouping;
    dialog.addField(body, strings.grouping, grouping, 'grouping');
    const title = dialog.input(255);
    title.value = fields.title;
    dialog.addField(body, strings.title, title, 'title');
    const range = dialog.input(255);
    range.value = fields.range;
    dialog.pointable(range, 'ranges');
    dialog.addField(body, strings.range, range, 'range');
    if (options.canAddSeries === false || options.canReorderSeries === false) {
      // Another range may ask for more series, or fewer: not offered where the chart cannot take that.
      range.readOnly = true;
      range.title = strings.seriesFixed;
      dialog.pointable(range, undefined);
    }
    const seriesIn = dialog.select((['auto', 'columns', 'rows'] as const).map((key) => [key, strings.directions[key] || key]));
    seriesIn.value = fields.seriesIn;
    dialog.addField(body, strings.seriesIn, seriesIn, 'seriesIn');
    if (options.canAddSeries === false || options.canReorderSeries === false) seriesIn.disabled = true;
    const legend = dialog.select((['r', 'l', 't', 'b', 'tr', 'none'] as const).map((key) => [key, strings.legends[key] || key]));
    legend.value = fields.legend;
    dialog.addField(body, strings.legend, legend, 'legend');
    const labels = dialog.addCheck(body, strings.dataLabels, 'dataLabels');
    labels.checked = fields.dataLabels;

    const axes = dialog.section(strings.axes);
    this.axes = axes;
    const text = (name: string, label: string, value: string): HTMLInputElement => {
      const input = dialog.input(255);
      input.value = value;
      dialog.addField(axes, label, input, name);
      return input;
    };
    this.controls.categoryTitle = text('categoryTitle', strings.categoryTitle, fields.categoryTitle);
    this.controls.valueTitle = text('valueTitle', strings.valueTitle, fields.valueTitle);
    this.controls.minimum = text('minimum', strings.minimum, fields.minimum);
    this.controls.maximum = text('maximum', strings.maximum, fields.maximum);

    const series = dialog.section(strings.series);
    const categories = dialog.input(8192);
    categories.value = fields.categories;
    dialog.pointable(categories, 'source');
    dialog.addField(series, strings.categories, categories, 'categories');
    const list = document.createElement('div');
    list.className = 'bym-chart-series';
    series.appendChild(list);
    this.list = list;
    const add = document.createElement('button');
    add.type = 'button';
    add.className = 'bym-dialog-button';
    add.setAttribute('data-bym-field', 'addSeries');
    add.textContent = strings.addSeries;
    add.addEventListener('click', () => {
      this.addRow({ name: '', values: '', color: '' });
      this.renumber();
      const last = this.rows[this.rows.length - 1];
      if (last) last.values.focus();
    });
    if (options.canAddSeries === false) {
      // What the chart cannot take is not offered, and says why.
      add.disabled = true;
      add.title = strings.seriesFixed;
    }
    series.appendChild(add);

    Object.assign(this.controls, { type, grouping, title, range, seriesIn, legend, dataLabels: labels, categories });
    for (const item of fields.series) this.addRow(item);
    this.renumber();
    type.addEventListener('change', () => this.sync());
    this.sync();

    dialog.addAction(strings.cancel, () => dialog.close()).setAttribute('data-bym-action', 'cancel');
    dialog.addAction(strings.ok, () => this.apply(), 'primary').setAttribute('data-bym-action', 'ok');
    dialog.markPristine();
    dialog.focusFirst();
    // The type cannot be chosen: the first field that can be is where the dialog starts.
    if (type.disabled) title.focus({ preventScroll: true });
  }

  /** What the fields hold now. */
  fields(): ChartFields {
    const c = this.controls;
    return {
      type: c.type.value,
      grouping: c.grouping.value,
      title: c.title.value,
      range: c.range.value,
      seriesIn: c.seriesIn.value,
      legend: c.legend.value,
      dataLabels: (c.dataLabels as HTMLInputElement).checked,
      categoryTitle: c.categoryTitle.value,
      valueTitle: c.valueTitle.value,
      minimum: c.minimum.value,
      maximum: c.maximum.value,
      categories: c.categories.value,
      series: this.rows.map((row) => {
        const out: ChartSeriesFields = { name: row.name.value, values: row.values.value, color: row.own.checked ? row.color.value.toLowerCase() : '' };
        if (row.from !== undefined) out.from = row.from;
        return out;
      }),
    };
  }

  /** What a type has no use for is not offered: the axes of a pie, the grouping of a scatter chart. */
  private sync(): void {
    const type = this.controls.type.value;
    const round = type === 'pie' || type === 'doughnut';
    // Set inline: the style of a section of a dialog would show it all the same.
    this.axes.style.display = round ? 'none' : '';
    (this.controls.grouping as HTMLSelectElement).disabled = round || type === 'scatter' || type === '';
  }

  private addRow(item: ChartSeriesFields): void {
    const dialog = this.dialog;
    const row = document.createElement('div');
    row.className = 'bym-chart-series-row';
    row.setAttribute('role', 'group');
    row.style.cssText = 'display:flex;align-items:center;flex-wrap:wrap;margin:0 0 6px';
    const name = dialog.input(255);
    name.value = item.name;
    name.style.flex = '1 1 90px';
    name.style.minWidth = '0';
    const values = dialog.input(8192);
    values.value = item.values;
    values.style.flex = '1 1 110px';
    values.style.minWidth = '0';
    dialog.pointable(values, 'source');
    const own = document.createElement('input');
    own.type = 'checkbox';
    own.className = 'bym-dialog-box';
    own.checked = item.color !== '';
    const color = dialog.color();
    color.value = item.color !== '' ? item.color : '#4472c4';
    color.disabled = !own.checked;
    own.addEventListener('change', () => {
      color.disabled = !own.checked;
    });
    const button = (text: string, act: () => void): HTMLButtonElement => {
      const el = document.createElement('button');
      el.type = 'button';
      el.className = 'bym-dialog-button';
      el.textContent = text;
      el.style.minWidth = '0';
      el.style.padding = '2px 8px';
      el.addEventListener('click', act);
      return el;
    };
    const entry = { from: item.from, name, values, own, color };
    const place = () => this.rows.indexOf(entry);
    const move = (by: number) => {
      const at = place();
      const to = at + by;
      if (at < 0 || to < 0 || to >= this.rows.length) return;
      this.rows.splice(at, 1);
      this.rows.splice(to, 0, entry);
      if (by < 0) this.list.insertBefore(row, this.list.children[to]);
      else this.list.insertBefore(row, this.list.children[to].nextSibling);
      this.renumber();
    };
    const up = button('↑', () => move(-1));
    const down = button('↓', () => move(1));
    const remove = button('×', () => {
      const at = place();
      if (at < 0) return;
      this.rows.splice(at, 1);
      this.list.removeChild(row);
      this.renumber();
    });
    for (const [el, field] of [
      [name, 'seriesName'],
      [values, 'seriesValues'],
      [own, 'seriesOwnColor'],
      [color, 'seriesColor'],
      [up, 'seriesUp'],
      [down, 'seriesDown'],
      [remove, 'seriesRemove'],
    ] as [HTMLElement, string][]) {
      el.setAttribute('data-bym-field', field);
      row.appendChild(el);
    }
    this.rows.push(entry);
    this.list.appendChild(row);
  }

  /** Names every control of the list by the place of its series, for assistive technology. */
  private renumber(): void {
    const strings = this.options.strings;
    this.rows.forEach((row, i) => {
      const n = i + 1;
      const buttons = (row.name.parentNode as HTMLElement).querySelectorAll('button');
      row.name.setAttribute('aria-label', strings.seriesName(n));
      row.name.placeholder = strings.seriesName(n);
      row.values.setAttribute('aria-label', strings.seriesValues(n));
      row.values.placeholder = strings.seriesValues(n);
      row.own.setAttribute('aria-label', strings.ownColor(n));
      row.color.setAttribute('aria-label', strings.seriesColor(n));
      if (buttons[0]) buttons[0].setAttribute('aria-label', strings.moveUp(n));
      if (buttons[1]) buttons[1].setAttribute('aria-label', strings.moveDown(n));
      if (buttons[2]) buttons[2].setAttribute('aria-label', strings.removeSeries(n));
      const fixed = this.options.canReorderSeries === false;
      if (buttons[0]) (buttons[0] as HTMLButtonElement).disabled = fixed || i === 0;
      if (buttons[1]) (buttons[1] as HTMLButtonElement).disabled = fixed || i === this.rows.length - 1;
      if (buttons[2]) (buttons[2] as HTMLButtonElement).disabled = fixed || this.rows.length === 1;
      for (let k = 0; k < 3 && fixed; k++) if (buttons[k]) (buttons[k] as HTMLButtonElement).title = strings.seriesFixed;
    });
  }

  private fieldOf(name: string): HTMLElement | undefined {
    const match = /^values(\d+)$/.exec(name);
    if (match) return this.rows[Number(match[1])] ? this.rows[Number(match[1])].values : undefined;
    if (name === 'series') return this.rows[0] ? this.rows[0].values : undefined;
    return this.controls[name];
  }

  private apply(): void {
    const dialog = this.dialog;
    dialog.showError('');
    let patch: ChartPatch;
    try {
      patch = chartPatch(this.initial, this.fields(), this.options.locale, this.options.strings);
    } catch (error) {
      dialog.showError(this.options.explain(error), error instanceof FieldError ? this.fieldOf(error.field) : undefined);
      return;
    }
    try {
      // What was only looked at is not changed: no step is made.
      if (Object.keys(patch).length > 0) this.options.onApply(patch);
    } catch (error) {
      dialog.showError(this.options.explain(error));
      return;
    }
    dialog.close();
  }
}
