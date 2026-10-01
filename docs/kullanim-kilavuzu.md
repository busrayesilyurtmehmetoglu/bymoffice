# BYM Office — geliştirici kullanım kılavuzu

Bu kılavuz, BYM Office'i kendi uygulamasına ekleyecek geliştiriciler içindir. Her kod değişikliğinde güncellenir;
burada yazan her şey depodaki sürümde çalışır durumdadır, henüz yapılmamış olanlar "Henüz olmayanlar" başlığında açıkça listelenir.

**Kılavuzun karşılık geldiği sürüm:** 0.0.1 (2026-10-01) — kapsayıcı katmanı, xlsx okuma modeli, tablo görünümü, sayı biçimleri ve ekranda hücre düzenleme.

## 1. Bu sürüm ne yapar

- docx, xlsx, doc ve xls dosyalarını **içeriğinden** tanır (dosya adına bakmaz).
- docx/xlsx içindeki her parçayı (XML, görsel, bilinmeyen ikili parça) okur.
- Bir parçayı değiştirip dosyayı yeniden kaydeder; dokunulmayan parçalar özgün sıkıştırılmış baytlarıyla, **bayt bayt aynı** kalır.
- xlsx dosyasını **tablo modeline** okur: sayfalar, hücre değerleri ve türleri, formül metinleri, stiller, birleştirilmiş hücreler, sütun genişlikleri, satır yükseklikleri, dondurulmuş bölmeler, gizli satır/sütun/sayfa, köprüler, adlandırılmış aralıklar.
- xlsx dosyasını sayfada **gösterir** (`BymOffice.mount()`): Canvas üzerinde yalnız görünen alan çizilir; kaydırma, hücre/aralık seçimi, klavye ile gezinme, dokunmatik kullanım, dondurulmuş bölmeler, sayfa sekmeleri.
- Hücreleri **Excel sayı biçimleriyle** gösterir: tarih, saat, para, yüzde, bilimsel, kesir ve özel biçim kodları; Türkçe ve İngilizce yerel ayar (`1.234,50 ₺` / `15.03.2024` ↔ `1,234.50` / `3/15/2024`).
- Hücreleri **düzenler** (şimdilik yalnız ekranda): yazarak giriş, geri al/yinele, kopyala/kes/yapıştır, satır ve sütun ekle-sil. Formül metinleri, birleştirilmiş hücreler ve adlandırılmış aralıklar ekleme-silmeyle birlikte kayar.
- doc/xls içindeki akışları (stream) okur.
- Tarayıcıda ve Node'da aynı kodla, çalışma zamanı bağımlılığı olmadan çalışır.

### Henüz olmayanlar

| Özellik | Durum |
|---|---|
| **Düzenlemeleri dosyaya yazma** | Yok (görev 1.5). `save()` hâlâ özgün dosyayı verir; ekranda yaptığınız değişiklikler kaydedilmez |
| Formül hesaplama | Yok (Aşama 2). Yazılan ya da yapıştırılan formül hücrede `=metin` olarak görünür; var olan formüllerin sonuçları, baktıkları hücreler değişse de eski değerinde kalır |
| Biçimlendirme (kalın, renk, kenarlık), sütun genişliği/satır yüksekliği değiştirme, birleştirme, sayfa ekleme-silme | Yok (Aşama 2) |
| Belge (docx/doc) ve xls görüntüleme | Yok; `mount()` `UnsupportedFileError` ile reddeder |
| Görünümde: döndürülmüş metin, sığdırmak için küçültme, iki yana yaslama, sağdan sola sayfa, yakınlaştırma, çapraz kenarlık, desenli/gradyan dolgu, başlıkları gizleme, görsel ve grafikler | Yok; içerik düz çizilir ya da atlanır |
| `onSave` seçeneği; arayüz metinlerinin Türkçesi | Yok (görev 1.6); menü ve ekran okuyucu metinleri şimdilik İngilizce |
| Sayı biçimlerinde: Türkçe ve İngilizce dışındaki diller, `*` ile doldurma, `_` boşluğunun gerçek genişliği, Hicri/Japon takvimleri, `[DBNum]` | Yok; `*` hiçbir şey yazmaz, `_x` tek boşluk olur |
| xls ve docx/doc için model | Yok; `workbook()` `UnsupportedFileError` ile reddeder |
| Koşullu biçim, veri doğrulama, grafik, görsel | Modelde yok; dosyada korunur |
| doc/xls yazma | Yok; `save()` özgün baytları döndürür |
| RTF, parola korumalı dosyalar | Yok; `UnsupportedFileError` fırlatılır |
| npm ve CDN'den kurulum | Paket henüz yayımlanmadı; depodan derlenir |
| 4 GB ve üzeri arşiv yazma | Yok |

## 2. Kurulum

Paket henüz npm'de yayımlanmadığı için depodan derlenir:

```sh
pnpm install
pnpm build
```

Çıktılar:

| Dosya | Kullanım |
|---|---|
| `dist/bymoffice.min.js` | `<script>` ile; global `BymOffice` nesnesini tanımlar |
| `dist/bymoffice.esm.js` | ES modülü (`import`) |
| `dist/types/index.d.ts` | TypeScript tipleri |

Yayımlandığında kurulum şöyle olacak (şimdilik çalışmaz):

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js"></script>
```

```sh
npm i bymoffice
```

## 3. Hızlı başlangıç

### Düz HTML (`<script>`)

```html
<input type="file" id="file" />
<script src="dist/bymoffice.min.js"></script>
<script>
  document.getElementById('file').addEventListener('change', async (event) => {
    try {
      const file = await BymOffice.open(event.target.files[0]);
      console.log(file.format, file.kind); // 'xlsx' 'spreadsheet'
      const bytes = await file.save();     // Uint8Array
    } catch (error) {
      console.error('Açılamadı:', error.message);
    }
  });
</script>
```

### ES modülü

```js
import { open, UnsupportedFileError } from './dist/bymoffice.esm.js';

const response = await fetch('/files/report.xlsx');
const file = await open(await response.arrayBuffer());
```

### Node

```js
import { readFile, writeFile } from 'node:fs/promises';
import { open } from './dist/bymoffice.esm.js';

const file = await open(new Uint8Array(await readFile('report.xlsx')));
await writeFile('copy.xlsx', await file.save());
```

## 4. API

### `mount(container, options): Promise<MountedView>`

Bir xlsx dosyasını açar ve `container` içinde gösterir. Kapsayıcının bir boyutu olmalıdır (ör. `height: 600px`); görünüm onu doldurur ve boyutu değişince kendini uyarlar.

```html
<div id="editor" style="height: 600px"></div>
<script src="dist/bymoffice.min.js"></script>
<script>
  (async () => {
    const mounted = await BymOffice.mount(document.getElementById('editor'), {
      file: '/files/report.xlsx',           // URL, Blob/File, ArrayBuffer ya da Uint8Array
      onSelectionChange: (selection, sheetIndex) => {
        console.log(BymOffice.rangeName(selection.range), selection.active);
      },
      onSheetChange: (sheetIndex) => console.log('sayfa', sheetIndex),
    });
    // mounted.file, mounted.workbook, mounted.view
    // mounted.destroy();  // görünümü kaldırır
  })();
</script>
```

| Seçenek | İş |
|---|---|
| `file` | Dosyanın kendisi ya da getirileceği adres. Adres verilirse `fetch` ile alınır; sunucu hata dönerse `Could not load the file: HTTP 404` gibi bir hata fırlatılır |
| `locale` | Sayı ve tarih biçimlerinin dili: `'tr'`, `'en'` gibi bir etiket ya da tam bir `FormatLocale` nesnesi. Verilmezse tarayıcının dili (`navigator.language`); Türkçe dışındaki her dil İngilizce (en-US) kurallarıyla gösterilir |
| `sheet` | İlk gösterilecek sayfanın dizini. Verilmezse dosyanın etkin sayfası; o gizliyse ilk görünür sayfa |
| `readOnly` | `true` ise düzenleme kapalıdır (editör, yapıştırma, kesme, silme, geri al, satır/sütun komutları çalışmaz); kopyalama çalışır. Varsayılan `false` |
| `onChange({ sheets })` | Çalışma kitabı her değiştiğinde çağrılır (geri al ve yinele dahil); `sheets` etkilenmiş olabilecek sayfa dizinleridir |
| `onSelectionChange(selection, sheetIndex)` | Seçim değişince çağrılır (açılışta ve sayfa değişiminde çağrılmaz) |
| `onSheetChange(sheetIndex)` | Gösterilen sayfa değişince çağrılır |

xlsx dışındaki dosyalarda `UnsupportedFileError` ile reddeder ve kapsayıcıya bir şey eklemez.

`mounted.view` bir `WorkbookView`'dir (modeli zaten elinizde tutuyorsanız `new BymOffice.WorkbookView(container, workbook, options)` ile doğrudan da kurabilirsiniz):

| Üye | İş |
|---|---|
| `activeSheet`, `setActiveSheet(index)` | Gösterilen sayfa; gizli sayfa da programla gösterilebilir |
| `selection` | `{ active: { row, col }, range: { top, left, bottom, right } }` (0 tabanlı; kopya döner) |
| `select(range, active?)` | Seçimi değiştirir ve görünür alana getirir; sınır dışı değerler sayfa sınırına çekilir |
| `scrollTo(row, col)` | Hücreyi görünür alana getirir |
| `setLocale(locale)` | Sayı ve tarih biçimlerinin dilini değiştirir ve yeniden çizer |
| `setCellInput(row, col, text)` | Hücreye, kullanıcı yazmış gibi metin girer (tek geri alma adımı) |
| `startEditing(text?)`, `commitEditing()`, `cancelEditing()`, `isEditing` | Hücre editörünü programla açar, yazar, kapatır |
| `clearSelection()` | Seçili aralığın içeriğini siler, biçimi korur |
| `undo()`, `redo()`, `history.canUndo`, `history.canRedo` | Geri al ve yinele; seçim ve sayfa da geri döner |
| `copy()`, `cut()`, `paste(text, html?)` | Pano işlemlerinin çekirdeği; `copy()` `{ text, html, … }` döndürür. `paste`'e yalnız metin verilirse dışarıdan gelmiş sayılır; kendi kopyanızı biçim ve formülleriyle yapıştırmak için `copy()`'nin `html`'ini de verin |
| `readOnly`, `setReadOnly(readOnly)` | Salt okunur kipini görünümü yeniden kurmadan değiştirir; geri alma geçmişi korunur |
| `insertRows(at, count)`, `deleteRows(at, count)`, `insertColumns(at, count)`, `deleteColumns(at, count)` | Satır ve sütun ekler ya da siler (0 tabanlı). Yapılamazsa `RangeError` fırlatır |
| `getViewState()` | `{ scrollX, scrollY, width, height, devicePixelRatio }` |
| `render()` | Hemen yeniden çizer (normalde gerekmez) |
| `destroy()` | Tüm DOM'u ve dinleyicileri kaldırır |

Kullanıcı etkileşimi:

- **Fare:** tıkla (seç), sürükle (aralık; kenarda kendiliğinden kayar), Shift+tık, satır/sütun başlığı (tüm satır/sütun), köşe (tümü). Birleştirilmiş hücre bütün olarak seçilir.
- **Dokunma:** dokun (seç), sürükle (kaydır; atalet tarayıcınındır).
- **Düzenleme:** hücreye yazmaya başlayın (içerik değişir), F2'ye basın ya da çift tıklayın/çift dokunun (mevcut içerik düzenlenir). Enter yazar ve aşağı iner, Shift+Enter yukarı çıkar, Tab sağa, Shift+Tab sola gider, Esc vazgeçer, Alt+Enter satır sonu ekler. Delete ya da Backspace seçili aralığı temizler. Ctrl/Cmd+Z geri alır, Ctrl+Y ya da Ctrl/Cmd+Shift+Z yineler. Ctrl/Cmd+C, X, V kopyalar, keser, yapıştırır.
- **Satır ve sütun:** satır ya da sütun başlığına sağ tıklayın (dokunmatikte uzun basın); açılan menüden ekleyin ya da silin. Tam satır ya da sütun seçiliyken Ctrl/Cmd+`+` ekler, Ctrl/Cmd+`-` siler.
- **Klavye** (önce görünüme tıklayın ya da Tab ile odaklanın): oklar, Shift+oklar, Ctrl/Cmd+oklar (veri bloğunun kenarı), Home, Ctrl+Home, Ctrl+End, PageUp/PageDown, Enter/Shift+Enter, Ctrl+A; sayfa değiştirmek için Alt+PageDown/PageUp (Ctrl+PageDown/PageUp da çalışır, ama bazı tarayıcılar onu kendine ayırır). Tab ve Shift+Tab hücre değiştirmez: odağı ızgaradan sayfadaki sonraki/önceki öğeye taşır, böylece klavye kullanıcıları görünümde hapsolmaz. Sayfa sekmeleri Tab ile odaklanıp Enter ya da Boşluk ile seçilir.

Bilinmesi gerekenler:

- Hücreler dosyadaki sayı biçimiyle gösterilir. `[Red]` gibi renk bölümleri yazı rengini değiştirir; biçimiyle gösterilemeyen değer (ör. negatif tarih) ve sütuna sığmayan sayı `####` olarak çizilir.
- Çizimde hücrenin kendi fontu kullanılır; cihazda yoksa `Carlito`, `Calibri`, `Arial` sırasıyla yedeğe düşülür. Font farkı metin genişliklerini biraz değiştirebilir.
- Sütun genişlikleri ve satır yükseklikleri Excel'in piksel yuvarlamasıyla hesaplanır (varsayılan sütun 64 px, satır 20 px).
- Sayfa her zaman 1.048.576 satır × 16.384 sütundur; kaydırma alanı kullanılan alanla başlar ve sona yaklaşıldıkça büyür.
- Erişilebilirlik: klavye odağı görünümdeki gizli bir metin alanındadır (`aria-label="Cell editor"`); kaydırılan bölüm `role="grid"` taşır, sekmeler `role="tab"` düğmeleridir; etkin hücrenin adresi ve metni ekran okuyucular için bir canlı bölgede duyurulur. Ekran okuyucuyla henüz denenmedi.
- CSS ile ölçeklenmiş (`transform: scale(...)`) bir kapsayıcıda da tıklamalar doğru hücreye gider.
- Görünüm kendi stillerini sayfaya tek bir `<style data-bymoffice>` öğesiyle ekler; sınıf adları `bym-` ile başlar. Dosyadan gelen hiçbir metin HTML olarak basılmaz.

### Düzenleme: yazılan metin nasıl yorumlanır

Yazılan metin Excel'deki gibi yorumlanır; kurallar seçili dile (`locale`) göre değişir:

| Yazılan (TR) | Sonuç |
|---|---|
| `12,5` · `1.234,56` · `-5` · `(5)` | Sayı |
| `15%` · `%15` | Yüzde (0,15); hücre "Genel" biçimdeyse yüzde biçimi alır |
| `15.03.2024` · `15/3/24` · `2024-03-15` · `15 Mart 2024` | Tarih (seri numarası); hücre "Genel" biçimdeyse tarih biçimi alır |
| `13:45` · `13:45:30` · `1:30 ÖS` | Saat |
| `12,50 ₺` · `$1,234.50` | Para; sayı ve para biçimi |
| `DOĞRU` · `YANLIŞ` · `TRUE` · `FALSE` | Mantıksal değer |
| `=A1*2` | Formül (henüz hesaplanmaz; `=A1*2` olarak görünür) |
| `'0123` | Baştaki `'` ile her şey metin olur: `0123` |
| diğer her şey | Metin |

Bilinmesi gerekenler:

- **Metin biçimli hücreler** (`@`): yazılan her şey olduğu gibi metin saklanır; `0123` sıfırını korur, `=A1` formül olmaz.
- **Uzun numaralar korunur:** sayıya tam sığmayan rakam dizileri (ör. 20 haneli bir numara, IBAN) Excel'in aksine yuvarlanmaz, metin olarak saklanır. Baştaki sıfırlar ise Excel'deki gibi düşer (`0123` → 123); korumak için hücreyi Metin biçimli yapın ya da başına `'` koyun.
- Hücrenin zaten bir sayı biçimi varsa yazılan tarih ya da yüzde o biçimi değiştirmez.
- İçeriği değiştirmeden editörden çıkmak (F2, sonra Enter) hücreyi değiştirmez ve geri alma adımı üretmez.
- **Satır/sütun ekle-sil:** formüllerdeki başvurular (diğer sayfalardakiler dahil), birleştirilmiş hücreler, köprüler, sütun genişlikleri ve adlandırılmış aralıklar birlikte kayar; silinen hücreye bakan başvuru `#REF!` olur. Bir dizi formülünü ortasından kesen ya da dolu hücreleri sayfanın dışına iten işlem Excel'deki gibi reddedilir.
- **Geri alma** en çok 200 adım tutar; büyük sayfalarda satır/sütun ekleme gibi ağır adımlar belleği korumak için daha erken düşebilir.
- **Düzenlerken geri al** yalnız açık düzenlemeyi iptal eder; önceki adımı geri almak için önce editörden çıkın.
- Bir milyon hücreden büyük yapıştırma, yarısı yazılıp kalmasın diye tümüyle reddedilir. Birleştirilmiş bir alanın altında kalan hücrelere yapıştırma yapılmaz.
- **Kesme** kopyalayıp temizler; Excel'deki gibi "taşıma" değildir (kesilen hücrelere bakan formüller yeni yeri izlemez).
- Kendi kopyaladığınız aralığı yapıştırdığınızda biçimler ve formüller birlikte gelir, formül başvuruları yeni yere göre kayar. Başka bir programdan gelen metin hücre hücre yazılmış gibi yorumlanır.

Aynı işlemler arayüz olmadan da kullanılabilir (ör. Node'da): `setCellInput(workbook, sheetIndex, row, col, text, { locale })`, `clearRange`, `copyRange`, `pasteAt` işlem listesi (`Op[]`) üretir; `new History(workbook).apply(ops, before, after)` uygular, `undo()`/`redo()` geri alır. Satır eklemek için doğrudan `{ type: 'insert', sheet, axis: 'row', at, count }` işlemi verilir. Yardımcılar: `parseCellInput(text, options)`, `cellEditText(cell, numFmt, options)`, `shiftFormula(text, dRow, dCol)`, `adjustFormula(text, sheetName, edit)`.

### `open(input): Promise<OfficeFile>`

`input`: `Uint8Array`, `ArrayBuffer` ya da `Blob`/`File`.

```ts
interface OfficeFile {
  format: 'docx' | 'xlsx' | 'doc' | 'xls';
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb';
  package?: OpcPackage; // docx/xlsx
  cfb?: CfbFile;        // doc/xls
  workbook(): Promise<Workbook>; // tablo modeli; şimdilik yalnız xlsx
  save(): Promise<Uint8Array>;
}
```

Hatalar:

- `UnsupportedFileError`: girdi Word/Excel dosyası değil ya da henüz desteklenmeyen bir tür (RTF, parola korumalı).
- `Error`: dosya bozuk (iletisi `Corrupt ZIP archive: …`, `Corrupt ZIP entry …` vb. ile başlar).

### Tablo modeli: `file.workbook()`

İlk çağrıda dosyayı okur, sonraki çağrılarda aynı nesneyi döndürür. Modeli okumak dosyayı değiştirmez; `save()` yine özgün baytları verir.
Satır ve sütun numaraları modelde **0 tabanlıdır** (`A1` → satır 0, sütun 0).

```js
const file = await BymOffice.open(bytes);
const workbook = await file.workbook();

for (const sheet of workbook.sheets) {
  console.log(sheet.name, sheet.state, sheet.rowCount, sheet.colCount);
}

const sheet = workbook.sheets[workbook.activeSheet];
const cell = BymOffice.getCell(sheet, 0, 0);     // A1; hücre yoksa undefined
cell.kind;      // 'blank' | 'number' | 'string' | 'boolean' | 'error' | 'date'
cell.value;     // number | string | boolean | null
cell.formula;   // { text: 'SUM(B2:B3)', type: 'normal' } ya da undefined
cell.runs;      // zengin metin parçaları (varsa)

const style = BymOffice.resolveCellStyle(workbook.styles, cell.style);
style.font.bold;                                          // boolean
style.numFmt;                                             // ör. '#,##0.00'
BymOffice.resolveColor(style.font.color, workbook.styles); // '#RRGGBB' ya da undefined
```

`Workbook`:

| Alan | İçerik |
|---|---|
| `sheets: Sheet[]` | Sayfalar, dosyadaki sırayla |
| `styles: StyleSheet` | Font, dolgu, kenarlık, hücre biçimi tabloları ve tema renkleri |
| `sharedStrings` | Paylaşılan metin tablosu |
| `definedNames` | Adlandırılmış aralıklar (`name`, `value`, `localSheet`, `hidden`) |
| `date1904` | Tarihler 1904 sistemine göre mi |
| `activeSheet` | Açılışta gösterilen sayfanın dizini |

`Sheet`:

| Alan | İçerik |
|---|---|
| `name`, `sheetId`, `state`, `kind`, `part` | Ad, kimlik, `'visible' \| 'hidden' \| 'veryHidden'`, `'worksheet' \| 'chartsheet' \| …`, kaynak parça |
| `rows: Map<number, Row>` | Seyrek satırlar; `Row`: `cells` (sütuna göre seyrek dizi), `height` (punto), `hidden`, `style`, `outlineLevel` |
| `columns: ColumnRange[]` | Sütun aralıkları: `min`, `max`, `width` (karakter), `hidden`, `style` |
| `merges: CellRange[]` | Birleştirilmiş alanlar (`top`, `left`, `bottom`, `right`; uçlar dahil) |
| `hyperlinks` | `range`, `target` (dış adres) ya da `location` (kitap içi), `tooltip` |
| `rowCount`, `colCount` | Hücre içeren son satır/sütun + 1 |
| `dimension` | Dosyanın bildirdiği kullanılan alan (eksik ya da yanlış olabilir) |
| `defaultRowHeight`, `defaultColWidth` | Varsayılan ölçüler |
| `view` | `frozenRows`, `frozenCols`, `topLeftCell` (kaydırılan alanın ilk hücresi), `firstVisibleCell` (görünümün sol üst hücresi; dondurulan alan buradan başlar), `showGridLines`, `zoomScale`, `rightToLeft`, `activeCell`, … |
| `tabColor` | Sekme rengi |

Yardımcılar: `getCell(sheet, row, col)`, `columnName(col)`, `cellName(row, col)`, `parseCellRef('B7')`, `parseRange('A1:C3')`, `rangeName(range)`, `resolveCellStyle(styles, index)`, `resolveColor(color, styles)`, `formatCodeOf(styles, numFmtId)`.

Bilinmesi gerekenler:

- Tarihler Excel'de olduğu gibi **sayı** olarak gelir; tarih olduğunu `style.numFmt` biçim kodu belirtir (`isDateFormat(code)`). Ekranda görünen metin için `cellDisplayText(workbook, cell, locale)`, takvim parçaları için `serialToDate(serial, workbook.date1904)` kullanın.
- Formüller hesaplanmaz; `cell.value` dosyaya en son kaydedilen sonuçtur.
- Grafik sayfaları (`kind: 'chartsheet'`) listede görünür ama hücre içermez.
- `activeSheet` gizli bir sayfayı gösterebilir (dosya öyle kaydedilmişse); kullanıcıya göstermeden önce `state` alanına bakın.
- Sayılar XML'deki ondalık yazımla okunur; `0x1F`, `NaN`, `INF` gibi değerler içeren dosya bozuk sayılır.
- Bozuk sayfa verisi `Error` ile reddedilir (iletisi `Invalid XML …` ya da `Invalid xlsx …`/`Invalid worksheet …` ile başlar).
- Başarım: 100.000 satır × 5 sütunluk sayfa güncel tarayıcılarda bir saniyenin altında okunur (her değişiklikte Playwright ile ölçülür).
- Görünüm için ek yardımcılar da dışa açıktır: `cellText(cell)` (ekranda görünen metin), `createSheetMetrics(sheet)`, `createLayout(...)`, `cellRect(...)`, `hitTest(...)`, `renderGrid(ctx, scene)`. Bunlar kendi tuvalinize çizmek ya da test yazmak içindir; arayüzleri ilk kararlı sürüme dek değişebilir.

### Sayı biçimleri

```js
const tr = BymOffice.getLocale('tr');

BymOffice.formatValue(1234.5, '#,##0.00 "₺"', { locale: tr }).text;   // '1.234,50 ₺'
BymOffice.formatValue(45366, 'd mmmm yyyy dddd', { locale: tr }).text; // '15 Mart 2024 Cuma'
BymOffice.formatValue(-5, '0.00;[Red](0.00)');                         // { text: '(5.00)', color: '#FF0000' }
BymOffice.formatValue(-1, 'yyyy-mm-dd');                               // { text: '####', overflow: true }

BymOffice.isDateFormat('dd/mm/yyyy');        // true
BymOffice.serialToDate(45366);               // { year: 2024, month: 3, day: 15, hour: 0, ..., weekday: 5 }
BymOffice.cellDisplayText(workbook, cell, tr); // hücrenin ekranda görünen metni
```

| İşlev | İş |
|---|---|
| `formatValue(value, code, { locale?, date1904?, builtinId? })` | Değeri biçim koduyla metne çevirir: `{ text, color?, overflow? }`. `builtinId` 14 ya da 22 ise kod yerine yerel ayarın kısa tarihi kullanılır |
| `formatCell(cell, code, options)` | Aynısı, model hücresi için. Hata değerleri olduğu gibi kalır; ISO tarih hücreleri (`kind: 'date'`) tarih biçimiyle gösterilir |
| `isoToSerial(text, date1904?)` | ISO 8601 metni → seri numarası (tarih değilse `undefined`) |
| `cellDisplayText(workbook, cell, locale?)` | Hücrenin stilindeki biçim koduyla ve çalışma kitabının tarih sistemiyle görünen metin |
| `isDateFormat(code)` | Kod tarih/saat biçimi mi |
| `serialToDate(serial, date1904?)` | Seri numarası → yıl, ay, gün, saat, dakika, saniye, milisaniye, haftanın günü (0 = Pazar) |
| `getLocale(tag)`, `LOCALE_TR`, `LOCALE_EN`, `normalizeLocale(partial)` | Yerel ayarlar; kendi `FormatLocale` nesnenizi de verebilirsiniz, eksik bıraktığınız alanlar en-US'ten tamamlanır |

Desteklenenler: dört bölüm (`pozitif;negatif;sıfır;metin`), koşullar (`[>100]`), renkler (`[Red]`, `[Color 3]`), `0 # ?` yer tutucuları, binlik ayırıcı ve ölçekleme virgülleri, yüzde, bilimsel gösterim, kesirler, sabit metin ve kaçışlar, para ve yerel ayar kodları (`[$₺-41F]`, `[$-409]`), tarih ve saat belirteçleri, geçen süre (`[h]:mm:ss`), AM/PM, 1900 ve 1904 tarih sistemleri (Excel'in 29.02.1900 hatasıyla birlikte).

Bilinmesi gerekenler:

- Çıktılar bağımsız bir uygulamayla (SheetJS SSF) binlerce kod–değer çiftinde karşılaştırılarak sınanır; gerçek Excel'le doğrudan karşılaştırma yapılmadı. Nadir kodlarda (ör. `00.0E+00`, `0` yer tutuculu kesirler, günün son yarım saniyesi) Excel'den küçük sapmalar olabilir.
- Yuvarlama Excel gibi ondalık gösterim üzerinden yarım yukarıdır: `1.005` → `1.01`. Sıfıra yuvarlanan negatif sayı Excel'deki gibi `-0.00` görünür.
- Yerleşik tarih biçimleri 14 ve 22 Excel'deki gibi yerel ayarın kısa tarihine çevrilir: TR `15.03.2024` ve `15.03.2024 06:00`, EN `3/15/2024` ve `3/15/2024 6:00`. Dosya aynı kodu (`mm-dd-yy`) kendi özel biçimi olarak yazmışsa kod olduğu gibi uygulanır (`03-15-24`).
- `[$-tr-TR]`, `[$-41F]`, `[$-409]` gibi yerel ayar kodları ay ve gün adlarının dilini seçer; ayırıcılar her zaman seçtiğiniz yerel ayardan gelir.
- 1900 tarih sisteminde negatif tarih `####` olur; 1904 sisteminde Excel'deki gibi eksi işaretiyle gösterilir.
- `NaN` ve sonsuz değerler `#NUM!` olarak gösterilir.
- Bozuk ya da tanınmayan bir biçim kodu hata fırlatmaz; değer "Genel" biçimiyle gösterilir.

### `OpcPackage` (docx/xlsx)

Parçalar mutlak adlarla anılır: `/xl/workbook.xml`, `/word/document.xml`.

| Üye | İş |
|---|---|
| `partNames(): string[]` | Tüm parça adları, dosyadaki sırayla |
| `hasPart(part): boolean` | Parça var mı (büyük/küçük harf duyarsız) |
| `contentTypeOf(part): string \| undefined` | İçerik türü |
| `mainPart(): Promise<string \| undefined>` | Ana parça (`/xl/workbook.xml`, `/word/document.xml`) |
| `relationships(part = '/'): Promise<Relationship[]>` | Bir parçanın ilişkileri |
| `readPart(part): Promise<Uint8Array>` | Parçanın açılmış baytları |
| `readXml(part): Promise<XmlDocument>` | Parçayı XML ağacı olarak okur |
| `writePart(part, data, contentType?)` | Parça ekler ya da değiştirir (`Uint8Array` ya da `XmlDocument`) |
| `save(): Promise<Uint8Array>` | Paketi yazar; dokunulmayan parçalar aynen kalır |

Örnek — bir parçayı değiştirip kaydetme:

```js
const file = await BymOffice.open(bytes);
const pkg = file.package;

const sheet = await pkg.readXml('/xl/worksheets/sheet1.xml');
sheet.root.attrs.set('edited', 'yes');
pkg.writePart('/xl/worksheets/sheet1.xml', sheet);

const saved = await file.save();
```

### `CfbFile` (doc/xls)

| Üye | İş |
|---|---|
| `entries: CfbEntry[]` | Akış ve klasörler (`path`, `type`, `size`) |
| `has(path): boolean` | Akış var mı |
| `stream(path): Uint8Array` | Akışın baytları (ör. `'Workbook'`, `'WordDocument'`) |

### XML

`parseXml(input)` ve `serializeXml(doc)`; ad önekleri, öznitelik sırası, yorumlar, CDATA ve boşluklar korunur.
DTD içeren belgeler güvenlik nedeniyle reddedilir. Yardımcılar: `childElements(element, localName?)`, `localName(name)`.

### `ZipArchive`

Düşük düzey ZIP erişimi: `ZipArchive.read(bytes)`, `names()`, `has(name)`, `entry(name)`, `read(name)`, `set(name, data)`, `delete(name)`, `write()`.

- ZIP64 arşivleri okunur; 65.534'ten fazla girdi ZIP64 kaydıyla yazılır.
- Girdi yorumları, ek alanlar, öznitelikler, arşiv yorumu ve UTF-8 olmayan adlar aynen geri yazılır.
- Aynı adı taşıyan iki girdi içeren arşiv, biri kaybolmasın diye reddedilir.
- Bildirdiği boyuttan büyük açılan ("zip bombası") girdi reddedilir.

## 5. Tarayıcı ve platform desteği

- Güncel Chrome, Firefox, Safari; iOS ve Android tarayıcıları ve WebView'leri. Her değişiklikte Playwright ile Chromium, Firefox, WebKit ve iki mobil profilde (iPhone, Pixel) sınanır.
- Tarayıcının yerleşik sıkıştırması (`DecompressionStream`) yoksa ya da `deflate-raw` biçimini desteklemiyorsa motor kendi inflate/deflate kodunu kullanır (iOS 16.4 öncesi, eski Android WebView). Ek ayar gerekmez.
- Paket ES2019 sözdizimine indirilmiş olarak derlenir.
- Sınır: gerçek eski cihazlarda deneme yapılmadı; eski ortam, testlerde yerleşik API'ler kaldırılarak taklit edilir.

## 6. Test arayüzü

Elle deneme için bir sayfa vardır:

```sh
pnpm demo
```

Ardından tarayıcıda `http://127.0.0.1:4318/demo/index.html` adresini açın (sunucuyu durdurmak için Ctrl+C). `pnpm build` sonrasında `demo/index.html` dosyasını çift tıklayarak da açabilirsiniz.

Sayfada yapabilecekleriniz:

- Bir ya da birden çok dosyayı sürükleyip bırakmak ya da seçmek.
- Her dosya için biçimi, boyutu, açılış süresini ve "kaydet → yeniden aç" sonucunu görmek (yeşil: tüm parçalar aynı; sarı: desteklenmiyor ya da ikili biçim; kırmızı: hata ya da içerik değişti).
- Parçaya ya da akışa tıklayıp içeriğini görmek (XML metin olarak, diğerleri onaltılık döküm).
- xlsx dosyalarında **tablo görünümü**: gerçek ızgara (kaydırma, seçim, klavye, sayfa sekmeleri, dondurulmuş bölmeler) ve **hücre düzenleme** (yazma, geri al/yinele, kopyala/yapıştır, satır/sütun ekle-sil; üstteki düğmeler ve "Salt okunur" seçimi); altında seçili aralık, etkin hücrenin görünen değeri, ham değeri ve formülü yazar. Düzenlemeler henüz dosyaya yazılmaz: "Yeniden kaydet ve indir" özgün dosyayı verir. "Sayı ve tarih biçimi dili" seçimiyle Türkçe ve İngilizce gösterim arasında geçilir.
- "Yeniden kaydet ve indir" ile kaydedilen dosyayı indirip Word/Excel'de açmak.
- "Örnek xlsx üret ve aç" ile elinizde dosya olmadan denemek.
- "Eski cihaz kipi" ile tarayıcının yerleşik sıkıştırmasını kapatıp motorun kendi kodunu sınamak.

Dosyalar yalnız tarayıcıda işlenir; hiçbir yere gönderilmez.

## 7. Geliştirme komutları

| Komut | İş |
|---|---|
| `pnpm verify` | **Her şeyi sıfırdan sınar:** tip denetimi, birim testleri, derleme, tüm tarayıcı testleri, gerçek dosya havuzu (iki sıkıştırma kipinde) |
| `pnpm typecheck` | Tip denetimi |
| `pnpm test` | Birim testleri |
| `pnpm build` | `dist/` çıktıları |
| `pnpm e2e` | Derler ve Playwright testlerini 5 tarayıcı profilinde çalıştırır |
| `pnpm corpus <klasör> …` | Gerçek dosyalarda aç → kaydet → yeniden aç denetimi (`BYM_NO_NATIVE=1` ile motorun kendi sıkıştırma koduyla) |
| `pnpm demo` | Derler ve test arayüzünü sunar |

Tarayıcıların ilk kurulumu: `PLAYWRIGHT_SKIP_BROWSER_GC=1 pnpm exec playwright install chromium firefox webkit`.

Gerçek dosya havuzu isteğe bağlıdır. Kendi belgelerinizle sınamak için depo köküne, depoya girmeyen bir `.corpus.local` dosyası koyun; her satırı bir klasör ya da seçenektir:

```
# gerçek Office dosyalarının bulunduğu klasörler (yalnız okunur)
/yol/belgeler
--min-passed=100
--allow-failures=0
```

Dosya yoksa `pnpm verify` havuz adımını atlar ve bunu çıktıda belirtir.
