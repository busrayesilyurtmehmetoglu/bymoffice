// The words of the editor shell: toolbar, formula bar and status bar. The
// grid's own words are in strings.ts. Number format codes that differ by
// language (currency, date) are here too, next to the labels they belong to.

export interface EditorStrings {
  toolbar: string;
  undo: string;
  redo: string;
  fontSize: string;
  bold: string;
  italic: string;
  underline: string;
  strikethrough: string;
  textColor: string;
  fillColor: string;
  automatic: string;
  noFill: string;
  /** Name of the palette colour in a row (0-5) and column (0-9) of the palette. */
  colorName(row: number, column: number): string;
  borders: string;
  borderAll: string;
  borderOutline: string;
  borderBottom: string;
  borderTop: string;
  borderLeft: string;
  borderRight: string;
  borderNone: string;
  alignLeft: string;
  alignCenter: string;
  alignRight: string;
  wrapText: string;
  mergeCells: string;
  numberFormat: string;
  formatGeneral: string;
  formatNumber: string;
  formatCurrency: string;
  formatPercent: string;
  formatDate: string;
  formatTime: string;
  formatText: string;
  formatCustom: string;
  /** Format codes the "Currency" and "Date" entries apply. */
  currencyCode: string;
  dateCode: string;
  rowsAndColumns: string;
  insertRowsAbove(count: number): string;
  insertRowsBelow(count: number): string;
  insertColumnsLeft(count: number): string;
  insertColumnsRight(count: number): string;
  deleteRows(count: number): string;
  deleteColumns(count: number): string;
  formulaBar: string;
  nameBox: string;
  cellContent: string;
  statusBar: string;
  readOnly: string;
  /** The line about the selected numbers; every argument is already formatted. */
  summary(sum: string, average: string, count: string): string;
  actionFailed: string;
}

/**
 * The palette has a row of greys, a row of plain colours and four rows with
 * two lighter and two darker shades of each plain colour.
 */
function paletteName(row: number, column: number, greys: string[], hues: string[], shade: (name: string, light: boolean, step: number) => string): string {
  if (row <= 0) return greys[column] || '';
  const hue = hues[column] || '';
  if (row === 1) return hue;
  return shade(hue, row < 4, row === 2 || row === 5 ? 2 : 1);
}

const GREYS_EN = ['Black', 'Dark grey 4', 'Dark grey 3', 'Dark grey 2', 'Dark grey 1', 'Grey', 'Light grey 1', 'Light grey 2', 'Light grey 3', 'White'];
const HUES_EN = ['Red berry', 'Red', 'Orange', 'Yellow', 'Green', 'Cyan', 'Cornflower blue', 'Blue', 'Purple', 'Magenta'];
const GREYS_TR = ['Siyah', 'Koyu gri 4', 'Koyu gri 3', 'Koyu gri 2', 'Koyu gri 1', 'Gri', 'Açık gri 1', 'Açık gri 2', 'Açık gri 3', 'Beyaz'];
const HUES_TR = ['Vişne', 'Kırmızı', 'Turuncu', 'Sarı', 'Yeşil', 'Camgöbeği', 'Peygamber çiçeği mavisi', 'Mavi', 'Mor', 'Macenta'];
const lower = (text: string) => text.charAt(0).toLowerCase() + text.slice(1);

const rows = (count: number) => (count === 1 ? 'row' : `${count} rows`);
const columns = (count: number) => (count === 1 ? 'column' : `${count} columns`);

export const EDITOR_STRINGS_EN: EditorStrings = {
  toolbar: 'Formatting',
  undo: 'Undo',
  redo: 'Redo',
  fontSize: 'Font size',
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  strikethrough: 'Strikethrough',
  textColor: 'Text colour',
  fillColor: 'Fill colour',
  automatic: 'Automatic',
  noFill: 'No fill',
  colorName: (row, column) => paletteName(row, column, GREYS_EN, HUES_EN, (name, light, step) => `${light ? 'Light' : 'Dark'} ${lower(name)} ${step}`),
  borders: 'Borders',
  borderAll: 'All borders',
  borderOutline: 'Outside borders',
  borderBottom: 'Bottom border',
  borderTop: 'Top border',
  borderLeft: 'Left border',
  borderRight: 'Right border',
  borderNone: 'No borders',
  alignLeft: 'Align left',
  alignCenter: 'Align centre',
  alignRight: 'Align right',
  wrapText: 'Wrap text',
  mergeCells: 'Merge cells',
  numberFormat: 'Number format',
  formatGeneral: 'General',
  formatNumber: 'Number',
  formatCurrency: 'Currency',
  formatPercent: 'Percent',
  formatDate: 'Date',
  formatTime: 'Time',
  formatText: 'Text',
  formatCustom: 'Custom',
  currencyCode: '"$"#,##0.00',
  dateCode: 'm/d/yyyy',
  rowsAndColumns: 'Rows and columns',
  insertRowsAbove: (count) => `Insert ${rows(count)} above`,
  insertRowsBelow: (count) => `Insert ${rows(count)} below`,
  insertColumnsLeft: (count) => `Insert ${columns(count)} left`,
  insertColumnsRight: (count) => `Insert ${columns(count)} right`,
  deleteRows: (count) => `Delete ${rows(count)}`,
  deleteColumns: (count) => `Delete ${columns(count)}`,
  formulaBar: 'Formula bar',
  nameBox: 'Name box',
  cellContent: 'Cell content',
  statusBar: 'Status bar',
  readOnly: 'Read-only',
  summary: (sum, average, count) => `Sum: ${sum} · Average: ${average} · Count: ${count}`,
  actionFailed: 'The action could not be completed',
};

const satir = (count: number) => (count === 1 ? 'satır' : `${count} satır`);
const sutun = (count: number) => (count === 1 ? 'sütun' : `${count} sütun`);

export const EDITOR_STRINGS_TR: EditorStrings = {
  toolbar: 'Biçimlendirme',
  undo: 'Geri al',
  redo: 'Yinele',
  fontSize: 'Yazı tipi boyutu',
  bold: 'Kalın',
  italic: 'İtalik',
  underline: 'Altı çizili',
  strikethrough: 'Üstü çizili',
  textColor: 'Yazı rengi',
  fillColor: 'Dolgu rengi',
  automatic: 'Otomatik',
  noFill: 'Dolgu yok',
  // Lower case by hand: none of the names starts with the letter that Turkish lower-cases differently.
  colorName: (row, column) => paletteName(row, column, GREYS_TR, HUES_TR, (name, light, step) => `${light ? 'Açık' : 'Koyu'} ${lower(name)} ${step}`),
  borders: 'Kenarlıklar',
  borderAll: 'Tüm kenarlıklar',
  borderOutline: 'Dış kenarlıklar',
  borderBottom: 'Alt kenarlık',
  borderTop: 'Üst kenarlık',
  borderLeft: 'Sol kenarlık',
  borderRight: 'Sağ kenarlık',
  borderNone: 'Kenarlık yok',
  alignLeft: 'Sola hizala',
  alignCenter: 'Ortaya hizala',
  alignRight: 'Sağa hizala',
  wrapText: 'Metni kaydır',
  mergeCells: 'Hücreleri birleştir',
  numberFormat: 'Sayı biçimi',
  formatGeneral: 'Genel',
  formatNumber: 'Sayı',
  formatCurrency: 'Para',
  formatPercent: 'Yüzde',
  formatDate: 'Tarih',
  formatTime: 'Saat',
  formatText: 'Metin',
  formatCustom: 'Özel',
  currencyCode: '#,##0.00 "₺"',
  dateCode: 'dd.mm.yyyy',
  rowsAndColumns: 'Satırlar ve sütunlar',
  insertRowsAbove: (count) => `Üste ${satir(count)} ekle`,
  insertRowsBelow: (count) => `Alta ${satir(count)} ekle`,
  insertColumnsLeft: (count) => `Sola ${sutun(count)} ekle`,
  insertColumnsRight: (count) => `Sağa ${sutun(count)} ekle`,
  deleteRows: (count) => (count === 1 ? 'Satırı sil' : `${count} satırı sil`),
  deleteColumns: (count) => (count === 1 ? 'Sütunu sil' : `${count} sütunu sil`),
  formulaBar: 'Formül çubuğu',
  nameBox: 'Ad kutusu',
  cellContent: 'Hücre içeriği',
  statusBar: 'Durum çubuğu',
  readOnly: 'Salt okunur',
  summary: (sum, average, count) => `Toplam: ${sum} · Ortalama: ${average} · Sayı: ${count}`,
  actionFailed: 'İşlem tamamlanamadı',
};

/** Strings for a language tag such as "tr" or "en-US"; English for anything unknown. */
export function getEditorStrings(tag: string | undefined): EditorStrings {
  return typeof tag === 'string' && /^tr(?:[-_]|$)/i.test(tag) ? EDITOR_STRINGS_TR : EDITOR_STRINGS_EN;
}

/** "Bold (Ctrl+B)": a label with its keyboard shortcut, named the way the platform names the key. */
export function withShortcut(label: string, key: string, shift = false): string {
  const apple = typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');
  return apple ? `${label} (${shift ? '⇧' : ''}⌘${key})` : `${label} (Ctrl+${shift ? 'Shift+' : ''}${key})`;
}
