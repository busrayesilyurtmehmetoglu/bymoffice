// The words of the editor shell: toolbar, formula bar and status bar. The
// grid's own words are in strings.ts. Number format codes that differ by
// language (currency, date) are here too, next to the labels they belong to.

export interface EditorStrings {
  toolbar: string;
  /** The Save button, shown when the editor was given somewhere to save to. */
  save: string;
  undo: string;
  redo: string;
  print: string;
  /** Shown in the status bar while the pages are prepared: "Preparing to print… 3 / 12". */
  printProgress(done: number, total: number): string;
  /** The button of the print preview. */
  printPreview: string;
  /** The button that downloads the workbook as a PDF file. */
  pdfDownload: string;
  /** Shown in the status bar while the pages of the PDF are painted: "Making the PDF… 3 / 12". */
  pdfProgress(done: number, total: number): string;
  /** The button that opens the dialog of the page setup. */
  pageSetup: string;
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
  /** Vertical alignment of the cells: top, middle, bottom. */
  alignTop: string;
  alignMiddle: string;
  alignBottom: string;
  wrapText: string;
  mergeCells: string;
  /** The menu beside the merge button, and its entries. */
  mergeOptions: string;
  mergeAndCenter: string;
  unmergeCells: string;
  /** Entries of the menu of rows and columns: frozen panes. */
  freezePanes: string;
  freezeTopRow: string;
  freezeFirstColumn: string;
  unfreezePanes: string;
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
  sortAscending: string;
  sortDescending: string;
  filter: string;
  /** The button that opens the data validation dialog. */
  dataValidation: string;
  /** The button that opens the menu of conditional formats. */
  conditionalFormat: string;
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
  // The document editor.
  paragraphStyle: string;
  /** Name of the default paragraph style when the document gives it none. */
  normalStyle: string;
  fontName: string;
  /** The button that opens the bar of Find and Replace. */
  findReplace: string;
  highlightColor: string;
  noHighlight: string;
  alignJustify: string;
  bulletList: string;
  numberedList: string;
  decreaseIndent: string;
  increaseIndent: string;
  /** The button of the dialog of tab stops. */
  tabStops: string;
  insertTable: string;
  /** A table size in the grid that inserts one: "3 x 2 table". */
  tableSize(rows: number, columns: number): string;
  tableRowsColumns: string;
  insertRowAbove: string;
  insertRowBelow: string;
  insertColumnLeft: string;
  insertColumnRight: string;
  deleteRow: string;
  deleteColumn: string;
  insertPicture: string;
  insertLink: string;
  linkAddress: string;
  applyLink: string;
  pageBreak: string;
  // Tracked changes of a document.
  trackChanges: string;
  /** Added to the label of the tracking button where tracking cannot be turned off. */
  trackingEnforced: string;
  previousChange: string;
  nextChange: string;
  acceptChange: string;
  rejectChange: string;
  /** Added to the labels of accept and reject where the protection of the document allows neither. */
  changesProtected: string;
  /** The button that opens the menu with "accept all" and "reject all". */
  allChanges: string;
  /** The comments of a document: a new one for the selection, the one before and the one after. */
  newComment: string;
  previousComment: string;
  nextComment: string;
  acceptAllChanges: string;
  rejectAllChanges: string;
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
  save: 'Save',
  undo: 'Undo',
  redo: 'Redo',
  print: 'Print',
  printProgress: (done, total) => `Preparing to print\u2026 ${done} / ${total}`,
  printPreview: 'Print preview',
  pdfDownload: 'Download as PDF',
  pdfProgress: (done, total) => `Making the PDF\u2026 ${done} / ${total}`,
  pageSetup: 'Page setup…',
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
  alignTop: 'Align top',
  alignMiddle: 'Align middle',
  alignBottom: 'Align bottom',
  wrapText: 'Wrap text',
  mergeCells: 'Merge cells',
  mergeOptions: 'Merge options',
  mergeAndCenter: 'Merge and centre',
  unmergeCells: 'Unmerge cells',
  freezePanes: 'Freeze panes',
  freezeTopRow: 'Freeze top row',
  freezeFirstColumn: 'Freeze first column',
  unfreezePanes: 'Unfreeze panes',
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
  sortAscending: 'Sort A to Z',
  sortDescending: 'Sort Z to A',
  filter: 'Filter',
  dataValidation: 'Data validation',
  conditionalFormat: 'Conditional formatting',
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
  paragraphStyle: 'Paragraph style',
  normalStyle: 'Normal',
  fontName: 'Font',
  findReplace: 'Find and replace',
  highlightColor: 'Highlight colour',
  noHighlight: 'No highlight',
  alignJustify: 'Justify',
  bulletList: 'Bulleted list',
  numberedList: 'Numbered list',
  decreaseIndent: 'Decrease indent',
  increaseIndent: 'Increase indent',
  tabStops: 'Tab stops…',
  insertTable: 'Insert table',
  tableSize: (rows, columns) => `${rows} x ${columns} table`,
  tableRowsColumns: 'Table rows and columns',
  insertRowAbove: 'Insert row above',
  insertRowBelow: 'Insert row below',
  insertColumnLeft: 'Insert column left',
  insertColumnRight: 'Insert column right',
  deleteRow: 'Delete row',
  deleteColumn: 'Delete column',
  insertPicture: 'Insert picture',
  insertLink: 'Insert link',
  linkAddress: 'Address',
  applyLink: 'Apply',
  pageBreak: 'Page break',
  trackChanges: 'Track changes',
  trackingEnforced: 'cannot be turned off in this document',
  previousChange: 'Previous change',
  nextChange: 'Next change',
  acceptChange: 'Accept change',
  rejectChange: 'Reject change',
  changesProtected: 'the protection of this document does not allow it',
  allChanges: 'All changes',
  newComment: 'Comment',
  previousComment: 'Previous comment',
  nextComment: 'Next comment',
  acceptAllChanges: 'Accept all changes',
  rejectAllChanges: 'Reject all changes',
};

const satir = (count: number) => (count === 1 ? 'satır' : `${count} satır`);
const sutun = (count: number) => (count === 1 ? 'sütun' : `${count} sütun`);

export const EDITOR_STRINGS_TR: EditorStrings = {
  toolbar: 'Biçimlendirme',
  save: 'Kaydet',
  undo: 'Geri al',
  redo: 'Yinele',
  print: 'Yazdır',
  printProgress: (done, total) => `Yazdırmaya hazırlanıyor\u2026 ${done} / ${total}`,
  printPreview: 'Baskı önizleme',
  pdfDownload: 'PDF olarak indir',
  pdfProgress: (done, total) => `PDF hazırlanıyor\u2026 ${done} / ${total}`,
  pageSetup: 'Sayfa yapısı…',
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
  alignTop: 'Üste hizala',
  alignMiddle: 'Ortaya hizala',
  alignBottom: 'Alta hizala',
  wrapText: 'Metni kaydır',
  mergeCells: 'Hücreleri birleştir',
  mergeOptions: 'Birleştirme seçenekleri',
  mergeAndCenter: 'Birleştir ve ortala',
  unmergeCells: 'Hücreleri ayır',
  freezePanes: 'Bölmeleri dondur',
  freezeTopRow: 'Üst satırı dondur',
  freezeFirstColumn: 'İlk sütunu dondur',
  unfreezePanes: 'Bölmeleri çöz',
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
  sortAscending: "A'dan Z'ye sırala",
  sortDescending: "Z'den A'ya sırala",
  filter: 'Süzgeç',
  dataValidation: 'Veri doğrulama',
  conditionalFormat: 'Koşullu biçimlendirme',
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
  paragraphStyle: 'Paragraf stili',
  normalStyle: 'Normal',
  fontName: 'Yazı tipi',
  findReplace: 'Bul ve değiştir',
  highlightColor: 'Vurgu rengi',
  noHighlight: 'Vurgu yok',
  alignJustify: 'İki yana yasla',
  bulletList: 'Madde imli liste',
  numberedList: 'Numaralı liste',
  decreaseIndent: 'Girintiyi azalt',
  increaseIndent: 'Girintiyi artır',
  tabStops: 'Sekme durakları…',
  insertTable: 'Tablo ekle',
  tableSize: (rows, columns) => `${rows} x ${columns} tablo`,
  tableRowsColumns: 'Tablo satır ve sütunları',
  insertRowAbove: 'Üste satır ekle',
  insertRowBelow: 'Alta satır ekle',
  insertColumnLeft: 'Sola sütun ekle',
  insertColumnRight: 'Sağa sütun ekle',
  deleteRow: 'Satırı sil',
  deleteColumn: 'Sütunu sil',
  insertPicture: 'Görsel ekle',
  insertLink: 'Bağlantı ekle',
  linkAddress: 'Adres',
  applyLink: 'Uygula',
  pageBreak: 'Sayfa sonu',
  trackChanges: 'Değişiklikleri izle',
  trackingEnforced: 'bu belgede kapatılamaz',
  previousChange: 'Önceki değişiklik',
  nextChange: 'Sonraki değişiklik',
  acceptChange: 'Değişikliği kabul et',
  rejectChange: 'Değişikliği reddet',
  changesProtected: 'bu belgenin koruması izin vermiyor',
  allChanges: 'Tüm değişiklikler',
  newComment: 'Yorum',
  previousComment: 'Önceki yorum',
  nextComment: 'Sonraki yorum',
  acceptAllChanges: 'Tüm değişiklikleri kabul et',
  rejectAllChanges: 'Tüm değişiklikleri reddet',
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

/** A label with a shortcut of the Control key, which is Control on Apple's machines as well: "Print preview (Ctrl+F2)", "(⌃F2)". */
export function withCtrlKey(label: string, key: string): string {
  const apple = typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');
  return apple ? `${label} (⌃${key})` : `${label} (Ctrl+${key})`;
}
