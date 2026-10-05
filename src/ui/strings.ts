// The few words the views themselves say: labels for assistive technology, the
// header menu, notes, the document view's bar. Content of files is never translated.

export interface UiStrings {
  spreadsheet: string;
  cellEditor: string;
  actionFailed: string;
  /** Label of the grid: "Sheet <name>". */
  sheet(name: string): string;
  chartSheetNote: string;
  otherSheetNote: string;
  /** Said after the active cell when more than one cell is selected. */
  selected(range: string): string;
  insertRowsAbove(count: number): string;
  insertRowsBelow(count: number): string;
  deleteRows(count: number): string;
  insertColumnsLeft(count: number): string;
  insertColumnsRight(count: number): string;
  deleteColumns(count: number): string;
  cut: string;
  copy: string;
  paste: string;
  clearContents: string;
  /** Said when the menu cannot read the clipboard and the keyboard has to be used. */
  pasteWithKeyboard: string;
  // What the operations refuse to do, in the user's language. `axis` is 'row' or 'col'.
  tooManyCells: string;
  tooManyRows(action: 'format' | 'resize' | 'change'): string;
  cannotMergeWhole: string;
  insertPushesOff(count: number, axis: string): string;
  arrayFormulaBlocks(kind: 'insert' | 'delete', axis: string, ref: string): string;
  /** `what` names the content in the way: comments, drawings, tables. */
  notSupportedWith(kind: 'insert' | 'delete', axis: string, what: string): string;
  pasteTooLarge(max: string): string;
  tooManyFormats: string;
  badFormula: string;
  copyTooLarge: string;
  // Sorting and filtering.
  /** Label of the menu of a filter button: "Sort and filter column B". */
  filterMenu(column: string): string;
  sortAscending: string;
  sortDescending: string;
  clearFilter: string;
  searchValues: string;
  selectAll: string;
  blanks: string;
  apply: string;
  cancel: string;
  noMatches: string;
  valueList: string;
  /** Under a value list that does not hold every value of its column. */
  tooManyValues(max: number): string;
  /** What a filter with criteria shows: "12 of 340 records shown". */
  recordsShown(shown: number, total: number): string;
  /** Said when choosing values took the place of a criterion that cannot be shown as values (a colour, an icon). */
  filterReplaced: string;
  nothingToFilter: string;
  cannotSortMerged: string;
  cannotSortArray(ref: string): string;
  /** `what` names the content in the way: comments, tables, pivot tables. */
  sortNotSupportedWith(action: 'sort' | 'filter', what: string): string;
  // Data validation.
  /** Label of the list a cell offers its values in. */
  chooseFromList: string;
  /** What a rule says when it has no message of its own. */
  validationFailed: string;
  /** Said when a value breaks a rule that only warns: the value is kept. */
  validationWarning(message: string): string;
  // Formulas this engine has something to say about (the cell has a mark in its corner).
  formulaNotCalculated: string;
  formulaCircular: string;
  formulaStale: string;
  /** Said before the comment of a cell, and the label of the box that shows it. */
  comment: string;
  /** Said once when a sheet is shown of which some drawings or comments could not be read. */
  drawingsNotShown: string;
  // The document view.
  document: string;
  documentTools: string;
  /** Label of one page: "Page 3". */
  page(number: number): string;
  /** The page indicator: "Page 3 / 12". */
  pageOf(page: number, count: number): string;
  zoomIn: string;
  zoomOut: string;
  fitWidth: string;
  /** Label of the button that shows the zoom and resets it to 100%. */
  zoomLevel(percent: number): string;
  /** Shown while the pages are still being laid out. */
  preparingPages: string;
  documentTooLong: string;
  documentLayoutFailed: string;
  /** What stands in a placeholder: `kind` is "picture", "chart", "diagram", "shape", "vml", "object", ... */
  placeholder(kind: string): string;
  // Editing a document.
  /** Label of the field that takes what is typed. */
  documentEditor: string;
  cannotEditHere: string;
  /** Said when a change is refused because the document tracks its changes and this kind of change cannot be tracked yet. */
  cannotTrackChange: string;
  /** Said once for a document whose protection keeps it read-only; `kind` is "readOnly", "comments" or "forms". */
  protectedDocument(kind: string): string;
  /** Said after a copy was pasted as text alone because changes are tracked. */
  pastedAsText: string;
  cannotCutObjects: string;
  cannotWriteClipboard: string;
  /** A copy made in an editor of the page does not fit where it is pasted. */
  cannotPasteHere: string;
  /** A paste waits for the pictures of the copy to be read from the document they are in. */
  pasteWaiting: string;
  /** The caret moved or the document changed while a paste was waiting: it was not made. */
  pasteInterrupted: string;
  /** Said after Escape: the next Tab moves the focus out of the editor instead of typing a tab. */
  tabLeaves: string;
  documentPasteTooLarge: string;
  /** A file was converted to be opened for editing; the formats as they are written ("doc", "docx"). */
  openedConverted(from: string, to: string): string;
  /** A file could not be converted to be edited and is shown as it is, read-only. */
  notConverted(from: string): string;
  pictureNotSupported: string;
  linkNotAllowed: string;
  /** Said when the caret enters another part of the document: `kind` is "body", "header", "footer", "footnote" or "endnote". */
  editingStory(kind: string): string;
  /** Said for an empty line and for a space the caret passes. */
  blankLine: string;
  space: string;
  endOfDocument: string;
  /** The bubble over a selection made by touch, and its button that selects the whole text. */
  selectionTools: string;
  selectWholeText: string;
  // Printing a document.
  print: string;
  /** Said when the pages are being prepared for the print dialog. */
  printPreparing: string;
  /** Said when the print dialog was closed; replaces `printPreparing`. */
  printClosed: string;
  /** Shown in the bar while they are: "Preparing to print… 3 / 12". */
  printProgress(done: number, total: number): string;
  printNotAvailable: string;
  printTooLong: string;
  printFailed: string;
  // Printing a workbook.
  printSheetTooLong: string;
  printSheetFailed: string;
  /** Said when the sheets that were to be printed hold nothing. */
  printNothing: string;
  // The sheet tabs.
  /** Label of the button beside the tabs that adds a sheet after the last one. */
  addSheet: string;
  /** Items of the menu of a tab. */
  insertSheet: string;
  renameSheet: string;
  deleteSheet: string;
  /** Label of the field a sheet is renamed in. */
  sheetName: string;
  /** Said after a sheet was deleted. */
  sheetDeleted(name: string): string;
  /** Said when a reference was put into the formula being typed by pointing at cells: "Reference B5:D9". */
  referenceAdded(reference: string): string;
  /** Said when what was typed could not be written because the sheet of the cell is no longer in the workbook. */
  editSheetGone: string;
  // What the operations on sheets refuse to do.
  badSheetName: string;
  sheetNameTaken(name: string): string;
  cannotDeleteLastSheet: string;
  /** A chart sheet or a sheet of another kind that is only kept. */
  sheetKindFixed(action: 'rename' | 'delete'): string;
  /** `what` names the content in the way: chart, pivotTable, table, form controls. */
  sheetBlockedBy(action: 'rename' | 'delete', what: string): string;
  /** `where` is the cell or the name whose formula cannot be read. */
  sheetFormulaUnread(action: 'rename' | 'delete', where: string): string;
  /** Two sheets whose names differ only by the case of their letters: the sheets of such a workbook are not changed. */
  sheetNamesClash(first: string, second: string): string;
}

const rows = (count: number) => (count === 1 ? 'row' : `${count} rows`);
const columns = (count: number) => (count === 1 ? 'column' : `${count} columns`);

export const STRINGS_EN: UiStrings = {
  spreadsheet: 'Spreadsheet',
  cellEditor: 'Cell editor',
  actionFailed: 'The action could not be completed',
  sheet: (name) => `Sheet ${name}`,
  chartSheetNote: 'This sheet is a chart sheet. Charts are not shown yet.',
  otherSheetNote: 'This kind of sheet is not shown yet.',
  selected: (range) => `${range} selected`,
  insertRowsAbove: (count) => `Insert ${rows(count)} above`,
  insertRowsBelow: (count) => `Insert ${rows(count)} below`,
  deleteRows: (count) => `Delete ${rows(count)}`,
  insertColumnsLeft: (count) => `Insert ${columns(count)} left`,
  insertColumnsRight: (count) => `Insert ${columns(count)} right`,
  deleteColumns: (count) => `Delete ${columns(count)}`,
  cut: 'Cut',
  copy: 'Copy',
  paste: 'Paste',
  clearContents: 'Clear contents',
  pasteWithKeyboard: 'The clipboard cannot be read here. Press Ctrl+V to paste.',
  tooManyCells: 'Too many cells to format at once',
  tooManyRows: (action) => `Too many rows to ${action} at once`,
  cannotMergeWhole: 'Whole rows or columns cannot be merged',
  insertPushesOff: (count, axis) => `Cannot insert ${count} ${axis === 'row' ? 'row' : 'column'}(s): cells would be pushed off the sheet`,
  arrayFormulaBlocks: (kind, axis, ref) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} through the array formula in ${ref}`,
  notSupportedWith: (kind, axis, what) => `Cannot ${kind} ${axis === 'row' ? 'rows' : 'columns'} in a sheet with ${what}: not supported yet`,
  pasteTooLarge: (max) => `Cannot paste more than ${max} cells at once`,
  tooManyFormats: 'The workbook has too many cell formats',
  badFormula: 'This formula cannot be read. Use English function names and commas between arguments.',
  copyTooLarge: 'The range is too large to copy',
  filterMenu: (column) => `Sort and filter column ${column}`,
  sortAscending: 'Sort A to Z',
  sortDescending: 'Sort Z to A',
  clearFilter: 'Clear filter',
  searchValues: 'Search',
  selectAll: '(Select all)',
  blanks: '(Blanks)',
  apply: 'Apply',
  cancel: 'Cancel',
  noMatches: 'Nothing matches',
  valueList: 'Values',
  tooManyValues: (max) => `Only the first ${max} values are listed`,
  recordsShown: (shown, total) => `${shown} of ${total} records shown`,
  filterReplaced: 'The colour or icon filter of this column was replaced by the chosen values.',
  nothingToFilter: 'There is no data to filter here',
  cannotSortMerged: 'A range with merged cells cannot be sorted',
  cannotSortArray: (ref) => `The array formula in ${ref} cannot be sorted through`,
  sortNotSupportedWith: (action, what) => `${action === 'sort' ? 'Sorting' : 'Filtering'} is not supported yet in a sheet with ${what}`,
  chooseFromList: 'Choose from the list',
  validationFailed: 'This value does not match the data validation rule of the cell.',
  validationWarning: (message) => `${message} (the value was entered anyway)`,
  formulaNotCalculated: 'This formula cannot be calculated here; the last value in the file is shown',
  formulaCircular: 'Circular reference',
  formulaStale: 'The calculation is not finished; the value may be out of date',
  comment: 'Comment',
  drawingsNotShown: 'Some drawings or comments of this sheet could not be read and are not shown',
  document: 'Document',
  documentTools: 'Document tools',
  page: (number) => `Page ${number}`,
  pageOf: (page, count) => `Page ${page} / ${count}`,
  zoomIn: 'Zoom in',
  zoomOut: 'Zoom out',
  fitWidth: 'Fit to width',
  zoomLevel: (percent) => `Zoom ${percent}%, reset to 100%`,
  preparingPages: 'Preparing pages\u2026',
  documentTooLong: 'The document is too long: only its first pages are shown',
  documentLayoutFailed: 'The document could not be laid out completely',
  placeholder: (kind) => (kind === 'chart' ? 'Chart' : kind === 'diagram' ? 'Diagram' : kind === 'object' ? 'Embedded object' : kind === 'picture' ? 'Picture' : 'Drawing'),
  documentEditor: 'Document editor',
  cannotEditHere: 'This part of the document cannot be edited',
  cannotTrackChange: 'Changes are tracked in this document, and this kind of change cannot be tracked yet; nothing was changed',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'This document is protected: only comments are allowed. It is shown read-only'
      : kind === 'forms'
        ? 'This document is protected: only its form fields may be filled in. It is shown read-only'
        : 'This document is protected against changes. It is shown read-only',
  pastedAsText: 'Changes are tracked in this document: the copy was pasted as text, without its formatting and pictures',
  cannotCutObjects: 'The selection holds something that cannot be cut (a note, an equation, a text box, a content control, a tracked change); it was left in place',
  cannotWriteClipboard: 'The clipboard cannot be written here; nothing was cut',
  cannotPasteHere: 'What was copied cannot be pasted here',
  pasteWaiting: 'Reading the pictures of the copy; it is pasted when they are there',
  pasteInterrupted: 'The selection or the document changed while the pictures of the copy were read; nothing was pasted',
  tabLeaves: 'Press Tab to leave the editor',
  documentPasteTooLarge: 'The text on the clipboard is too long to paste',
  openedConverted: (from, to) => `This ${from} file was converted to ${to} to be edited; it is saved as ${to}`,
  notConverted: (from) => `This ${from} file could not be converted to be edited; it is shown read-only`,
  pictureNotSupported: 'This picture cannot be inserted',
  linkNotAllowed: 'This address cannot be used as a link',
  editingStory: (kind) => (kind === 'header' ? 'Editing the header' : kind === 'footer' ? 'Editing the footer' : kind === 'footnote' ? 'Editing a footnote' : kind === 'endnote' ? 'Editing an endnote' : 'Editing the text'),
  blankLine: 'blank',
  space: 'space',
  endOfDocument: 'End of the document',
  selectionTools: 'Selection',
  selectWholeText: 'Select all',
  print: 'Print',
  printPreparing: 'Preparing to print\u2026',
  printClosed: 'The print dialog was closed',
  printProgress: (done, total) => `Preparing to print\u2026 ${done} / ${total}`,
  printNotAvailable: 'Printing is not available here',
  printTooLong: 'The document is too long to print at once',
  printFailed: 'The document could not be printed',
  printSheetTooLong: 'The sheet is too long to print at once',
  printSheetFailed: 'The sheet could not be printed',
  printNothing: 'There is nothing to print',
  addSheet: 'Add sheet',
  insertSheet: 'Insert sheet',
  renameSheet: 'Rename',
  deleteSheet: 'Delete',
  sheetName: 'Sheet name',
  sheetDeleted: (name) => `Sheet ${name} was deleted. Undo brings it back.`,
  referenceAdded: (reference) => `Reference ${reference}`,
  editSheetGone: 'The sheet of the cell is gone; what was typed was not written.',
  badSheetName: 'A sheet name has 1 to 31 characters, none of [ ] : * ? / \\, no apostrophe at its start or end, and is not "History"',
  sheetNameTaken: (name) => `Another sheet is already named "${name}"`,
  cannotDeleteLastSheet: 'The last visible sheet cannot be deleted',
  sheetKindFixed: (action) => `This kind of sheet cannot be ${action === 'rename' ? 'renamed' : 'deleted'} yet`,
  sheetBlockedBy: (action, what) =>
    what === 'table'
      ? `A sheet with a table cannot be ${action === 'rename' ? 'renamed' : 'deleted'} yet`
      : `Sheets cannot be ${action === 'rename' ? 'renamed' : 'deleted'} yet in a workbook with ${what === 'pivotTable' ? 'a pivot table' : what === 'chart' ? 'a chart' : what}`,
  sheetFormulaUnread: (action, where) => `The sheet cannot be ${action === 'rename' ? 'renamed' : 'deleted'}: the formula of ${where} cannot be read`,
  sheetNamesClash: (first, second) => `The sheets of this workbook cannot be changed: "${first}" and "${second}" differ only by case`,
};

const satir = (count: number) => (count === 1 ? 'satır' : `${count} satır`);
const sutun = (count: number) => (count === 1 ? 'sütun' : `${count} sütun`);

export const STRINGS_TR: UiStrings = {
  spreadsheet: 'Hesap tablosu',
  cellEditor: 'Hücre düzenleyici',
  actionFailed: 'İşlem tamamlanamadı',
  sheet: (name) => `Sayfa ${name}`,
  chartSheetNote: 'Bu sayfa bir grafik sayfası. Grafikler henüz gösterilmiyor.',
  otherSheetNote: 'Bu tür sayfalar henüz gösterilmiyor.',
  selected: (range) => `${range} seçili`,
  insertRowsAbove: (count) => `Üste ${satir(count)} ekle`,
  insertRowsBelow: (count) => `Alta ${satir(count)} ekle`,
  deleteRows: (count) => (count === 1 ? 'Satırı sil' : `${count} satırı sil`),
  insertColumnsLeft: (count) => `Sola ${sutun(count)} ekle`,
  insertColumnsRight: (count) => `Sağa ${sutun(count)} ekle`,
  deleteColumns: (count) => (count === 1 ? 'Sütunu sil' : `${count} sütunu sil`),
  cut: 'Kes',
  copy: 'Kopyala',
  paste: 'Yapıştır',
  clearContents: 'İçeriği temizle',
  pasteWithKeyboard: 'Pano buradan okunamıyor. Yapıştırmak için Ctrl+V kullanın.',
  tooManyCells: 'Bir kerede biçimlendirilemeyecek kadar çok hücre seçili',
  tooManyRows: (action) =>
    action === 'format'
      ? 'Bir kerede biçimlendirilemeyecek kadar çok satır seçili'
      : action === 'resize'
        ? 'Bir kerede boyutlandırılamayacak kadar çok satır seçili'
        : 'Bir kerede değiştirilemeyecek kadar çok satır seçili',
  cannotMergeWhole: 'Tam satırlar ya da sütunlar birleştirilemez',
  insertPushesOff: (count, axis) => `${count} ${axis === 'row' ? 'satır' : 'sütun'} eklenemiyor: hücreler sayfanın dışına itilirdi`,
  arrayFormulaBlocks: (kind, axis, ref) =>
    `${ref} aralığındaki dizi formülünün içinden ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'eklenemez' : 'silinemez'}`,
  notSupportedWith: (kind, axis, what) =>
    `Bu sayfada ${axis === 'row' ? 'satır' : 'sütun'} ${kind === 'insert' ? 'ekleme' : 'silme'} henüz desteklenmiyor (sayfada bulunan: ${what})`,
  pasteTooLarge: (max) => `Bir kerede ${max} hücreden fazlası yapıştırılamaz`,
  tooManyFormats: 'Çalışma kitabında çok fazla hücre biçimi var',
  badFormula: 'Bu formül okunamadı. Argümanları noktalı virgülle (;) ayırın, ondalıkları virgülle yazın; tanınmayan fonksiyonlarda İngilizce adı kullanın.',
  copyTooLarge: 'Aralık kopyalanamayacak kadar büyük',
  filterMenu: (column) => `${column} sütununu sırala ve süz`,
  sortAscending: "A'dan Z'ye sırala",
  sortDescending: "Z'den A'ya sırala",
  clearFilter: 'Süzgeci temizle',
  searchValues: 'Ara',
  selectAll: '(Tümünü seç)',
  blanks: '(Boşlar)',
  apply: 'Uygula',
  cancel: 'İptal',
  noMatches: 'Eşleşen değer yok',
  valueList: 'Değerler',
  tooManyValues: (max) => `Yalnız ilk ${max} değer listeleniyor`,
  recordsShown: (shown, total) => `${total} kayıttan ${shown} tanesi gösteriliyor`,
  filterReplaced: 'Bu sütunun renk ya da simge süzgecinin yerini seçilen değerler aldı.',
  nothingToFilter: 'Burada süzülecek veri yok',
  cannotSortMerged: 'Birleştirilmiş hücre içeren aralık sıralanamaz',
  cannotSortArray: (ref) => `${ref} aralığındaki dizi formülü bölünerek sıralanamaz`,
  sortNotSupportedWith: (action, what) => `Bu sayfada ${action === 'sort' ? 'sıralama' : 'süzme'} henüz desteklenmiyor (sayfada bulunan: ${what})`,
  chooseFromList: 'Listeden seçin',
  validationFailed: 'Bu değer hücrenin veri doğrulama kuralına uymuyor.',
  validationWarning: (message) => `${message} (değer yine de yazıldı)`,
  formulaNotCalculated: 'Bu formül burada hesaplanamıyor; dosyadaki son değer gösteriliyor',
  formulaCircular: 'Döngüsel başvuru',
  formulaStale: 'Hesap bitmedi; değer eski olabilir',
  comment: 'Yorum',
  drawingsNotShown: 'Bu sayfadaki bazı çizimler ya da yorumlar okunamadı; gösterilmiyor',
  document: 'Belge',
  documentTools: 'Belge araçları',
  page: (number) => `Sayfa ${number}`,
  pageOf: (page, count) => `Sayfa ${page} / ${count}`,
  zoomIn: 'Yakınlaştır',
  zoomOut: 'Uzaklaştır',
  fitWidth: 'Genişliğe sığdır',
  zoomLevel: (percent) => `Yakınlaştırma %${percent}, %100'e döndür`,
  preparingPages: 'Sayfalar hazırlanıyor\u2026',
  documentTooLong: 'Belge çok uzun: yalnız ilk sayfaları gösteriliyor',
  documentLayoutFailed: 'Belgenin tamamı sayfalara yerleştirilemedi',
  placeholder: (kind) => (kind === 'chart' ? 'Grafik' : kind === 'diagram' ? 'Diyagram' : kind === 'object' ? 'Gömülü nesne' : kind === 'picture' ? 'Görsel' : 'Çizim'),
  documentEditor: 'Belge düzenleyici',
  cannotEditHere: 'Belgenin bu bölümü düzenlenemez',
  cannotTrackChange: 'Bu belgede değişiklikler izleniyor ve bu tür değişiklik henüz izlenemiyor; hiçbir şey değiştirilmedi',
  protectedDocument: (kind) =>
    kind === 'comments'
      ? 'Bu belge korumalı: yalnız yorum eklenebilir. Salt okunur gösteriliyor'
      : kind === 'forms'
        ? 'Bu belge korumalı: yalnız form alanları doldurulabilir. Salt okunur gösteriliyor'
        : 'Bu belge değişikliğe karşı korumalı. Salt okunur gösteriliyor',
  pastedAsText: 'Bu belgede değişiklikler izleniyor: kopya biçimi ve görselleri olmadan, metin olarak yapıştırıldı',
  cannotCutObjects: 'Seçimde kesilemeyen bir öğe var (dipnot, denklem, metin kutusu, içerik denetimi, izlenen değişiklik); yerinde bırakıldı',
  cannotWriteClipboard: 'Pano buradan yazılamıyor; hiçbir şey kesilmedi',
  cannotPasteHere: 'Kopyalanan içerik buraya yapıştırılamıyor',
  pasteWaiting: 'Kopyanın görselleri okunuyor; okununca yapıştırılacak',
  pasteInterrupted: 'Kopyanın görselleri okunurken seçim ya da belge değişti; hiçbir şey yapıştırılmadı',
  tabLeaves: 'Düzenleyiciden çıkmak için Sekme tuşuna basın',
  documentPasteTooLarge: 'Panodaki metin yapıştırılamayacak kadar uzun',
  openedConverted: (from, to) => `Bu ${from} dosyası düzenlenebilmesi için ${to} biçimine çevrildi; ${to} olarak kaydedilir`,
  notConverted: (from) => `Bu ${from} dosyası düzenlenebilecek biçime çevrilemedi; salt okunur gösteriliyor`,
  pictureNotSupported: 'Bu görsel eklenemiyor',
  linkNotAllowed: 'Bu adres bağlantı olarak kullanılamaz',
  editingStory: (kind) => (kind === 'header' ? 'Üstbilgi düzenleniyor' : kind === 'footer' ? 'Altbilgi düzenleniyor' : kind === 'footnote' ? 'Dipnot düzenleniyor' : kind === 'endnote' ? 'Sonnot düzenleniyor' : 'Metin düzenleniyor'),
  blankLine: 'boş',
  space: 'boşluk',
  endOfDocument: 'Belgenin sonu',
  selectionTools: 'Seçim',
  selectWholeText: 'Tümünü seç',
  print: 'Yazdır',
  printPreparing: 'Yazdırmaya hazırlanıyor\u2026',
  printClosed: 'Yazdırma penceresi kapandı',
  printProgress: (done, total) => `Yazdırmaya hazırlanıyor\u2026 ${done} / ${total}`,
  printNotAvailable: 'Burada yazdırma kullanılamıyor',
  printTooLong: 'Belge bir kerede yazdırılamayacak kadar uzun',
  printFailed: 'Belge yazdırılamadı',
  printSheetTooLong: 'Sayfa bir kerede yazdırılamayacak kadar uzun',
  printSheetFailed: 'Sayfa yazdırılamadı',
  printNothing: 'Yazdırılacak bir şey yok',
  addSheet: 'Sayfa ekle',
  insertSheet: 'Sayfa ekle',
  renameSheet: 'Yeniden adlandır',
  deleteSheet: 'Sil',
  sheetName: 'Sayfa adı',
  sheetDeleted: (name) => `${name} sayfası silindi. Geri al ile geri gelir.`,
  referenceAdded: (reference) => `Başvuru ${reference}`,
  editSheetGone: 'Hücrenin sayfası artık yok; yazılan kaydedilmedi.',
  badSheetName: 'Sayfa adı 1-31 karakter olmalı; [ ] : * ? / \\ içeremez, kesme işaretiyle başlayıp bitemez ve "History" olamaz',
  sheetNameTaken: (name) => `"${name}" adında başka bir sayfa var`,
  cannotDeleteLastSheet: 'Görünen son sayfa silinemez',
  sheetKindFixed: (action) => `Bu tür sayfalar henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : 'silinemiyor'}`,
  sheetBlockedBy: (action, what) =>
    what === 'table'
      ? `Tablo içeren sayfa henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : 'silinemiyor'}`
      : `${what === 'pivotTable' ? 'Pivot tablo' : what === 'chart' ? 'Grafik' : what === 'form controls' ? 'Form denetimi' : what} içeren çalışma kitabında sayfalar henüz ${action === 'rename' ? 'yeniden adlandırılamıyor' : 'silinemiyor'}`,
  sheetFormulaUnread: (action, where) => `Sayfa ${action === 'rename' ? 'yeniden adlandırılamadı' : 'silinemedi'}: ${where} formülü okunamıyor`,
  sheetNamesClash: (first, second) => `Bu çalışma kitabının sayfaları değiştirilemiyor: "${first}" ve "${second}" adları yalnız harf büyüklüğüyle ayrılıyor`,
};

/** Strings for a language tag such as "tr" or "en-US"; English for anything unknown. */
export function getUiStrings(tag: string | undefined): UiStrings {
  return typeof tag === 'string' && /^tr(?:[-_]|$)/i.test(tag) ? STRINGS_TR : STRINGS_EN;
}

/**
 * The strings a view speaks for the `locale` it was given: those of a language
 * tag, and for anything else (no locale, a locale object, which only changes
 * formats) those of the browser's language.
 */
export function resolveUiStrings(locale: string | object | undefined): UiStrings {
  return getUiStrings(typeof locale === 'string' ? locale : typeof navigator !== 'undefined' ? navigator.language : undefined);
}
