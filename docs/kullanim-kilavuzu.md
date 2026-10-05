# BYM Office — geliştirici kullanım kılavuzu

Bu kılavuz, BYM Office'i kendi uygulamasına ekleyecek geliştiriciler içindir. Her kod değişikliğinde güncellenir;
burada yazan her şey depodaki sürümde çalışır durumdadır, henüz yapılmamış olanlar "Henüz olmayanlar" başlığında açıkça listelenir.

**Kılavuzun karşılık geldiği sürüm:** 0.0.1 (2026-10-01) — xlsx görüntüleme, düzenleme, biçimlendirme, formül hesaplama ve kayıpsız kaydetme; editör arayüzü; WebView köprüsü.

## 1. Bu sürüm ne yapar

- docx, xlsx, doc ve xls dosyalarını **içeriğinden** tanır (dosya adına bakmaz).
- docx/xlsx içindeki her parçayı (XML, görsel, bilinmeyen ikili parça) okur.
- Bir parçayı değiştirip dosyayı yeniden kaydeder; dokunulmayan parçalar özgün sıkıştırılmış baytlarıyla, **bayt bayt aynı** kalır.
- xlsx dosyasını **tablo modeline** okur: sayfalar, hücre değerleri ve türleri, formül metinleri, stiller, birleştirilmiş hücreler, sütun genişlikleri, satır yükseklikleri, dondurulmuş bölmeler, gizli satır/sütun/sayfa, köprüler, adlandırılmış aralıklar.
- xlsx dosyasını sayfada **gösterir** (`BymOffice.mount()`): Canvas üzerinde yalnız görünen alan çizilir; kaydırma, hücre/aralık seçimi, klavye ile gezinme, dokunmatik kullanım, dondurulmuş bölmeler, sayfa sekmeleri.
- Hücreleri **Excel sayı biçimleriyle** gösterir: tarih, saat, para, yüzde, bilimsel, kesir ve özel biçim kodları; Türkçe ve İngilizce yerel ayar (`1.234,50 ₺` / `15.03.2024` ↔ `1,234.50` / `3/15/2024`).
- Hücreleri **düzenler ve dosyaya kaydeder**: yazarak giriş, geri al/yinele, kopyala/kes/yapıştır, satır ve sütun ekle-sil. Formül metinleri, birleştirilmiş hücreler ve adlandırılmış aralıklar ekleme-silmeyle birlikte kayar.
- **Editör arayüzü** sunar: araç çubuğu (geri al/yinele, yazı boyutu, kalın/italik/altı çizili/üstü çizili, yazı ve dolgu rengi, kenarlık, hizalama, metni kaydır, birleştir, sayı biçimi, satır/sütun), formül çubuğu ve ad kutusu, durum çubuğu (toplam/ortalama/sayı), sağ tık menüsü, sürükleyerek sütun/satır boyutlandırma. Türkçe ve İngilizce.
- **Word belgelerini (docx) sayfa sayfa gösterir ve düzenler**: imleç ve seçim (fare, klavye, dokunmatik tutamaçlar), yazma ve silme, giriş yöntemleri (IME), geri al/yinele, kopyala/kes/yapıştır, biçimlendirme araç çubuğu (stil, font, boyut, kalın/italik/altı çizili, renk, hizalama, listeler, tablo, bağlantı, görsel); `save()` yalnız değişen blokları yeniden yazar.
- **Eski Word dosyalarını (doc, Word 97–2003) salt okunur açar ve gösterir**: metin, karakter ve paragraf biçimi, stiller, listeler, tablolar (birleşik ve iç içe hücreler), bölümler, üstbilgi/altbilgi, dipnotlar, alanlar, yer imleri, satır içi ve kayan resimler, metin kutuları, şekiller (dikdörtgen, yuvarlatılmış dikdörtgen, elips, düz çizgi; düz renk dolgu, çizgi rengi, kalınlığı ve kesikliği, çevirme ve döndürme — modelde `drawing.shape`). `file.document()` docx ile aynı `DocModel`'i verir; resimler pakette değil `model.media` içindedir (`BymOffice.documentPictures(model, file.package)` ikisini tek arayüzle sunar).
- **RTF dosyalarını salt okunur açar ve gösterir** (`.doc` uzantısıyla kaydedilmiş RTF dahil; biçim içerikten tanınır): metin ve kod sayfaları (Türkçe 1254 dahil), karakter/paragraf biçimi, stiller, listeler, tablolar (birleşik ve iç içe), bölümler, üstbilgi/altbilgi, dipnotlar, alanlar, resimler, metin kutuları, şekiller (doc ile aynı kapsam: dikdörtgen, yuvarlatılmış dikdörtgen, elips, düz çizgi; dolgu ve çizgi). Yarıda kesilmiş RTF hata vermez; okunabilen kısmı gösterilir.
- **Eski Excel dosyalarını (xls, Excel 97–2003) salt okunur açar**: hücreler, stiller, birleştirmeler, sütun/satır ölçüleri, dondurulmuş bölmeler, köprüler, tanımlı adlar, formül metinleri ve dosyadaki sonuçları, sayfa yapısı (kenar boşlukları, kâğıt, yön, ölçek/sığdırma, üstbilgi/altbilgi, sayfa sonları; `sheet.pageSetup`). `mount()` xls'i her zaman salt okunur kurar. **`file.saveAsXlsx()` xls'i xlsx olarak yazar** (aşağıda "xls dosyasını xlsx olarak kaydetme"). Ayrıntı için `BymOffice.readXlsWithReport(file.cfb)` okunmayan kayıtları ve çevrilemeyen formülleri sayar. Excel 5/95 dosyaları da okunur ama gerçek dosyayla sınanmamıştır.
- **Excel sayfalarında görselleri, basit şekilleri/metin kutularını ve hücre yorumlarını gösterir** (xlsx ve xls): görseller (png, jpeg, gif, bmp, webp) çapalandıkları hücrelerde, satır/sütun eklenip silinince yazıcının dosyaya yazdığıyla aynı yere kayar; yorumlu hücrenin köşesinde kırmızı üçgen vardır ve üzerine gelince, hücre seçilince ya da dokununca yorum balonu açılır. Dış bağlantılı görseller yüklenmez; SVG/EMF/WMF yerinde yer tutucu görünür. Model: `sheet.drawings`, `sheet.comments`, `sheet.warnings`.
- **Sayfa yapısını (yazdırma ayarlarını) modele okur** (xlsx ve xls): `sheet.pageSetup` — kenar boşlukları, kâğıt boyutu, yön, ölçek ve sayfaya sığdırma, üstbilgi/altbilgi metinleri, yatay/dikey ortalama, kılavuz çizgisi ve başlık yazdırma, sayfa sonları. xls → xlsx dönüştürmede yeni dosyaya yazılır. Yazdırmada uygulanır ("Tabloyu yazdırma"); bu ayarları değiştiren bir arayüz yoktur.
- **Word belgelerini yazdırır** (docx, doc, RTF; salt okunur görünümde de): belge çubuğundaki Yazdır düğmesi, Ctrl/Cmd+P ya da `print()` tarayıcının yazdırma penceresini belgenin kâğıt boyutunda açar; "PDF olarak kaydet" de oradan yapılır. Sayfalar görüntü olarak basılır (ayrıntı ve sınırlar: "Belgeyi yazdırma").
- **Tabloları yazdırır** (xlsx, xls; salt okunur görünümde de): araç çubuğundaki Yazdır düğmesi, Ctrl/Cmd+P ya da `print()` etkin sayfayı (istenirse görünür bütün sayfaları) sayfa yapısına göre kâğıda dizer: kâğıt ve yön, kenar boşlukları, ölçek ya da sayfaya sığdırma, yazdırma alanı, yinelenen başlıklar, elle sayfa sonları, üstbilgi/altbilgi, kılavuz çizgisi ve satır/sütun başlığı seçenekleri (ayrıntı ve sınırlar: "Tabloyu yazdırma").
- **Koşullu biçimleri gösterir**: hücre kuralları, formüllü kurallar, renk ölçekleri, veri çubukları, ikon kümeleri, ilk/son N, ortalama üstü/altı, yinelenen/benzersiz, metin ve tarih kuralları. Hücre düzenlenince ve geri alınınca güncellenir; kaydederken kurallar dosyada aynen korunur.
- **Formülleri hesaplar**: 225 Excel fonksiyonu; bir hücre değişince yalnız ona bağlı formüller yeniden hesaplanır.
- **WebView köprüsü** ile mobil ve masaüstü uygulamalara gömülür (ayrıntı: [webview-koprusu.md](webview-koprusu.md)).
- doc/xls içindeki akışları (stream) okur.
- Tarayıcıda ve Node'da aynı kodla, çalışma zamanı bağımlılığı olmadan çalışır.

### Henüz olmayanlar

| Özellik | Durum |
|---|---|
| Formüllerde: dinamik dizi taşması, `LET`/`LAMBDA`, tablo başvuruları (`Tablo[Sütun]`), başka dosyaya ve 3B (`Sayfa1:Sayfa3!A1`) başvurular, `;` ile argüman ayırma | Yok. Böyle bir formül içeren hücre dosyadaki son değerini korur ya da `#REF!`/`#NAME?` verir; formül metni kaydederken bozulmaz |
| Tablo ya da form denetimi içeren sayfalarda, ve içinde grafik ya da pivot tablo bulunan çalışma kitaplarında satır/sütun ekleme-silme | Reddedilir (bu içeriklerin başvuruları henüz kaydırılmıyor; sessizce bozulmasınlar diye); hücre düzenleme ve biçimlendirme serbesttir. Resim ve yorum içeren sayfalarda serbesttir |
| Formül yazarken hücre göstermede: F4 ile mutlak başvuru (`$A$1`) ve başvuru türleri arasında dolaşma; formüldeki bütün başvuruların renklendirilmesi (yalnız son eklenen çerçevelenir); çerçeveyi sürükleyerek başvuruyu taşıma/boyutlandırma; dokunmatikte aralık gösterme (tek dokunuş hücre, satır ya da sütun ekler; sürükleme kaydırmadır); köşe düğmesiyle tüm sayfa; Ctrl+tık ile birden çok alan; adlandırılmış aralık ya da tablo adı önerme; formül çubuğundan klavyeyle hücre gösterme ve sayfa değiştirme; başka sayfa gösterilirken hücre içi düzenleyicinin görünmesi | Yok. Başvuru her zaman görelidir; `$` elle yazılır |
| Yazı tipi adı seçimi; dikey hizalama düğmeleri; koşullu biçim ve veri doğrulama KURALI ekleme/değiştirme | Yok |
| Sayfalarda: taşıma/sıra değiştirme, gizleme/gösterme, kopyalama, sekme rengi değiştirme | Yok (ekleme, silme ve yeniden adlandırma var; bkz. "Sayfa ekleme, silme, yeniden adlandırma") |
| İçinde grafik, pivot tablo ya da form denetimi bulunan çalışma kitabında sayfa silme ve yeniden adlandırma; tablo içeren sayfayı silme; grafik sayfasını silme ve yeniden adlandırma; adları yalnız harf büyüklüğüyle ayrılan iki sayfası (`IL` / `İL`) olan kitapta her sayfa işlemi (ekleme dahil) | Reddedilir (bu içerikler sayfaları adıyla anar ve henüz yeniden yazılmıyor); menüde öğe devre dışı görünür. Sayfa **ekleme** her zaman serbesttir |
| WebView köprüsünde sayfa listesini bildiren ileti ve sayfa ekle/sil/yeniden adlandır komutları | Yok. Sekme şeridi köprüde de çalışır, ama `loaded` iletisindeki sayfa listesi sonradan güncellenmez ve yeniden adlandırma bildirilmez; `changed.sheets` ve `sheet.index` işlemden sonraki dizinlerdir |
| Yeniden adlandırmada ve silmede çizimlerin (şekil, görsel) iç bağlantılarındaki sayfa adlarının güncellenmesi | Yok; bu bağlantılar eski adı taşımaya devam eder. Formüller, tanımlı adlar, hücre bağlantıları, koşullu biçim ve doğrulama formülleri, tablo sütun formülleri güncellenir |
| Renge ve ikona göre süzme, tablo içi süzgeç, yorumlu ya da tablolu sayfada sıralama, uyarı türü doğrulamada onay penceresi, Türkçe hata adları (`#YOK`) | Yok. Uygulanamayan süzgeç ölçütleri dosyada korunur |
| Koşullu biçimde: kural ekleme/değiştirme arayüzü, Excel'in özgün ikon görselleri (basit vektör ikonlar çizilir), kuraldaki yazı tipi adı/boyutu ve desenli/gradyan dolgu | Yok. `calculate: false` iken formüllü kurallar uygulanmaz |
| Word 95 ve öncesi, parola korumalı doc | Yok; `UnsupportedFileError` |
| RTF'te matematik, çerçeve yerleşimi, Word 95 çizim nesnelerinin (`\do`) çizgi ve dolguları (metin kutusu olanların yalnız metni okunur, diğerleri atlanır), eski `\pn` listeleri, tablo stilleri, Uzak Doğu kod sayfaları; RTF'e yazma | Yok. RTF salt okunur açılır, `save()` özgün baytları verir |
| doc düzenleme ve kaydetme; doc'ta EMF/WMF görseller, resim madde imleri, eski otomatik numaralar, yorum aralıkları | Yok. doc salt okunur açılır, `save()` özgün baytları verir |
| Sayfa yapısını (kenar boşluğu, kâğıt, ölçek, üstbilgi/altbilgi, sayfa sonu, yazdırma alanı) düzenleme; sayfa sonu önizlemesi; yazdırma önizlemesi | Yok. Sayfa yapısı dosyadan okunur (`sheet.pageSetup`), yazdırmada uygulanır, dosyada korunur ve xls → xlsx dönüştürmede taşınır |
| Tablo yazdırmada: bir sayfaya sığmayan satır/sütunun ortadan bölünmesi (kendi sayfasına konur, taşan kısmı kenar boşluğunda kesilir); yorumların basılması (`cellComments`: sayfa sonunda ya da göründüğü gibi); hata hücrelerinin boş/tire/`#N/A` basılması (`errors`); siyah-beyaz ve taslak kipleri; üstbilgi/altbilgi görselleri (`&G`), tema renkleri (`&K04+000`), ana hat ve gölge (`&O`, `&H`), dosya yolu (`&Z` boş kalır); grafik sayfaları; kopya sayısı ve yazıcı çözünürlüğü ayarları; yalnız seçili aralığı yazdırma; sağdan sola sayfalar; metni seçilebilir PDF | Yok. Bu ayarlar dosyada korunur, çıktıyı etkilemez |
| Belge yazdırmada: metni seçilebilir/aranabilir PDF (sayfalar görüntü olarak basılır), doğrudan PDF dosyası üretme (yalnız tarayıcının yazdırma penceresi üzerinden), sayfa aralığı/kopya sayısı seçenekleri (tarayıcının penceresinden seçilir), WebView'lerde kendiliğinden yazdırma (`window.print()` olmayan ya da hiçbir şey yapmayan WebView'de uygulama kendi yazdırmasını bağlar; bkz. [webview-koprusu.md](webview-koprusu.md)) | Yok |
| xls düzenleme ve xls olarak kaydetme; xls'te koşullu biçim, veri doğrulama, otomatik süzgeç ölçütleri, yazıcı sürücüsü verisi (PLS), özel görünümlerin (custom view) sayfa yapısı, üstbilgi/altbilgi görselleri, özet (outline) ayarları, sayfa koruması, grafikler, form denetimleri, makrolar, dış çalışma kitabı yolları; parola korumalı xls; Excel 4 ve öncesi | Yok. xls salt okunur açılır, `save()` özgün baytları verir. Bunlar modele okunmadığı için `saveAsXlsx()` ile üretilen xlsx'te de bulunmaz |
| Belge düzenlemede: dipnot/sonnot başvurusu, yorum, denklem, içerik denetimi, izlenen değişiklik, metin kutusu ve metinli şekil, grafik, diyagram, gömülü nesne (OLE), VML resim, `mc:AlternateContent`, birden çok paragrafa yayılan alan (içindekiler tablosu), akıllı etiket ve özel XML'in panoyla taşınması (kopyada bunların yalnız metni gider; bunları içeren seçim **kesilemez**, bildirimle yerinde bırakılır), başka sekmedeki/sayfadaki bir düzenleyiciye ya da başka uygulamaya zengin içerik taşıma (pano dışarıya yalnız metin ve basit HTML verir), dışarıdan gelen HTML'in biçimiyle yapıştırılması (yalnız metni alınır), tablo hücrelerinin üzerine hücre hücre yapıştırma (tablo, iç içe tablo olarak eklenir), değişiklik izleme ve yorum düzenleme, metin kutularının içini düzenleme | Yok. Görsel, bağlantı, alan, sayfa/sütun sonu, tablo, paragraf biçimi, stil ve liste aynı sayfadaki düzenleyiciler arasında taşınır (bkz. Pano) |
| Belgede: belge içi bağlantılar, arama, sağdan sola yazı, heceleme, grafik ve vektör görsel (EMF/WMF/SVG) çizimi, belgeye gömülü fontlar | Yok. Çizilemeyen nesnelerin yerinde yer tutucu görünür |
| Belgede şekiller (docx, doc, RTF): dikdörtgen, yuvarlatılmış dikdörtgen, elips ve düz çizgi dışındaki geometriler (oklar, serbest biçimler, akış şeması şekilleri…); ok başları; çift/üçlü (bileşik) çizgi; özel kesik çizgi; gradyan, desen ve resim dolgu; saydamlık; gölge ve diğer efektler; şekil gruplarının ve çizim tuvalinin parçaları; metin kutusunun metniyle birlikte döndürülmesi | Yok. **Tek şekilli çizimde** (grup değil) çizilenler: hazır geometri `rect` (kesik/tek köşesi yuvarlak dikdörtgen türleri ve `flowChartProcess` de düz dikdörtgen olarak), `roundRect`, `ellipse`, `line`, `straightConnector1`; düz renk dolgu; çizginin rengi, kalınlığı (0 = en ince çizgi) ve hazır kesiklik türü; metinsiz şekilde döndürme, çizgide çevirme. **Yaklaşık çizilenler:** ok başlı çizgi başsız, bileşik çizgi tek, özel kesiklik düz çizilir; gölge çizilmez; kısmi saydam renk opak boyanır (tam saydam boyanmaz); gradyan/desen olan *çizgi* hiç çizilmez; metinli şekil döndürülmez. **Yer tutucu olanlar:** metinsiz şekil başka bir geometrideyse (ya da docx'te özel geometri taşıyor ya da geometri belirtmiyorsa) ya da dolgusu gradyan/desen/resimse; metinsiz gruplar, tuvaller, grafikler, diyagramlar. Bu durumlardaki **metin kutusunda** metin gösterilir: tanınmayan geometride docx'te çerçevesiz, doc/RTF'te dikdörtgen çerçeveyle; gösterilemeyen dolguda dolgusuz. **Gruplarda** yalnız parçaların metinleri (alt alta) ve resimleri gösterilir, parçaların geometrisi çizilmez. doc/RTF'te ayrıca: palet/şema/sistem dizinli renkler çözülmez (dolgu boyanmaz, çizgi siyah olur), çizim grubunun varsayılan özellik tablosu okunmaz (şekil kendi belirtmediği özellikte belirtimin varsayılanını alır), ok başı, bileşik çizgi ve gölge okunmaz; resmi bulunamayan resim şekli yer tutucudur. docx'te tüm bunlar dosyada korunur |
| Görünümde: döndürülmüş metin, sığdırmak için küçültme, iki yana yaslama, sağdan sola sayfa, yakınlaştırma, çapraz kenarlık, desenli/gradyan dolgu, başlıkları gizleme, görsel ve grafikler | Yok; içerik düz çizilir ya da atlanır |
| Sayı biçimlerinde: Türkçe ve İngilizce dışındaki diller, `*` ile doldurma, `_` boşluğunun gerçek genişliği, Hicri/Japon takvimleri, `[DBNum]` | Yok; `*` hiçbir şey yazmaz, `_x` tek boşluk olur |
| Model | doc ve RTF `document()` ile docx'le aynı belge modeline, xls `workbook()` ile aynı tablo modeline okunur. docx modeli `BymOffice.docOps` işlemleriyle düzenlenir ve `file.save()` ile dosyaya yazılır; doc ve RTF modeli yalnız `saveAsDocx()` ile docx olarak, xls modeli yalnız `saveAsXlsx()` ile xlsx olarak yazılır |
| doc/xls biçimine yazma | Yok; `save()` özgün baytları döndürür (`saveAsDocx()`/`saveAsXlsx()` yeni biçime çevirir) |
| Parola korumalı dosyalar | Yok; `UnsupportedFileError` fırlatılır |
| Excel'de grafik ve pivot tablo gösterimi (yerinde yer tutucu görünür); görsel/şekil ekleme, taşıma, boyutlandırma; dikdörtgen, elips ve çizgi dışındaki şekil geometrileri; gradyan/desen dolgu, gölge; yorum ekleme/düzenleme | Yok; dosyada korunur |
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
| `dist/types/index.d.ts` | TypeScript tipleri (`moduleResolution` `bundler`, `node16` ve `nodenext` ile çözülür) |
| `dist/bymoffice.webview.html` | WebView'ler için tek dosyalık köprü sayfası (pakette `bymoffice/webview.html`) |

Paket yalnız ES modülü ve `<script>` olarak dağıtılır; ayrı bir CommonJS çıktısı yoktur (`require()` yalnız ES modülünü yükleyebilen Node sürümlerinde çalışır). Çalışma zamanı bağımlılığı yoktur.

`pnpm pack:check` (derlemeden sonra) paketi yayımlanacağı haliyle sınar: arşivi üretir, boş bir projeye kurar, oradan ES modülü olarak kullanır, tipleri üç çözümleme kipiyle derler ve `<script>` dosyasının global `BymOffice`'i tanımladığına bakar. `pnpm verify` bunu da çalıştırır.

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

Bir dosyayı açar ve `container` içinde gösterir: xlsx ve xls tablo düzenleyicide, docx, doc ve RTF belge düzenleyicide (aşağıda "Word belgesini gösterme ve düzenleme"). Kapsayıcının bir boyutu olmalıdır (ör. `height: 600px`); görünüm onu doldurur ve boyutu değişince kendini uyarlar.

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
| `toolbar`, `formulaBar`, `statusBar` | `false` verilirse ilgili bölüm gösterilmez (varsayılan hepsi `true`). Üçü de kapalıyken yalnız ızgara ve sayfa sekmeleri kalır |
| `calculate` | `false` verilirse formüller yeniden hesaplanmaz (varsayılan `true`) |
| `onNotice({ message, kind })` | Kullanıcıya gösterilen hata ve bilgi iletileri (arayüz dilinde) |
| `sheet` | İlk gösterilecek sayfanın dizini. Verilmezse dosyanın etkin sayfası; o gizliyse ilk görünür sayfa |
| `readOnly` | `true` ise düzenleme kapalıdır (editör, yapıştırma, kesme, silme, geri al, satır/sütun komutları çalışmaz); kopyalama çalışır. Varsayılan `false` |
| `onChange({ sheets })` | Çalışma kitabı her değiştiğinde çağrılır (geri al ve yinele dahil); `sheets` etkilenmiş olabilecek sayfa dizinleridir |
| `convert` | `true` ise yazılamayan biçimler (doc, RTF, xls) **dönüştürülerek düzenlenebilir açılır**: doc ve RTF docx'e, xls xlsx'e çevrilir; gösterilen, düzenlenen ve `save()` ile alınan dosya yeni biçimdedir. Dönüş değerinde `conversion: { from, to, warnings }` bulunur ve `onNotice` ile bildirilir (bildirimin dili görünümlerle aynı kuralla seçilir: `locale` bir dil etiketiyse o, verilmemişse ya da nesneyse tarayıcının dili). `readOnly: true` ile birlikte verilirse dönüştürme yapılmaz. Varsayılan `false`: bu dosyalar olduğu gibi, salt okunur gösterilir |
| `onSave(bytes)` | Kullanıcı görünümde Ctrl+S / Cmd+S'e basınca dosyanın baytlarıyla çağrılır; verilmezse kısayol tarayıcıya bırakılır. Hata olursa `onError(error)` çağrılır |
| `onSelectionChange(selection, sheetIndex)` | Seçim değişince çağrılır (açılışta ve sayfa değişiminde çağrılmaz) |
| `onSheetChange(sheetIndex)` | Gösterilen sayfa değişince çağrılır |

Tanınmayan ya da desteklenmeyen dosyalarda `UnsupportedFileError` ile reddeder ve kapsayıcıya bir şey eklemez.

`mounted.editor` kabuğu (`SpreadsheetEditor`), `mounted.view` içindeki ızgarayı (`WorkbookView`) verir. Modeli zaten elinizde tutuyorsanız `new BymOffice.SpreadsheetEditor(container, workbook, options)` ile kabuklu, `new BymOffice.WorkbookView(container, workbook, options)` ile yalnız ızgarayı doğrudan kurabilirsiniz. `editor.setLocale(...)` ve `editor.setReadOnly(...)` kabuğu ve ızgarayı birlikte değiştirir; `editor.destroy()` hepsini kaldırır.

Izgaranın (`mounted.view`) üyeleri:

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
| `insertSheet(at?, name?)`, `renameSheet(index, name)`, `deleteSheet(index)`, `startRename(index)`, `isRenaming` | Sayfa ekler, yeniden adlandırır, siler; sekmede ad alanını açar. Aşağıda "Sayfa ekleme, silme, yeniden adlandırma" |
| `copy()`, `cut()`, `paste(text, html?)` | Pano işlemlerinin çekirdeği; `copy()` `{ text, html, … }` döndürür. `paste`'e yalnız metin verilirse dışarıdan gelmiş sayılır; kendi kopyanızı biçim ve formülleriyle yapıştırmak için `copy()`'nin `html`'ini de verin |
| `readOnly`, `setReadOnly(readOnly)` | Salt okunur kipini görünümü yeniden kurmadan değiştirir; geri alma geçmişi korunur |
| `insertRows(at, count)`, `deleteRows(at, count)`, `insertColumns(at, count)`, `deleteColumns(at, count)` | Satır ve sütun ekler ya da siler (0 tabanlı). Yapılamazsa `RangeError` fırlatır |
| `format(patch)` | Seçime biçim uygular: `{ bold, italic, underline, strike, fontName, fontSize, color, fill, horizontal, vertical, wrapText, numFmt, border: { preset } }`; `null` ilgili özelliği kaldırır |
| `merge()`, `unmerge()`, `selectionMerged` | Seçimi tek hücrede birleştirir (yalnız sol üst hücrenin içeriği kalır) ya da ayırır |
| `selectionFormat` | Etkin hücrenin biçimi (araç çubuğu durumu için) |
| `setColumnWidth(first, last, width)`, `setRowHeight(first, last, height)` | Genişlik (karakter) ve yükseklik (punto); `undefined` varsayılana döndürür, `0` gizler |
| `editText`, `setEditText(text)`, `focus()` | Formül çubuğu gibi dış düzenleyiciler için |
| `getViewState()` | `{ scrollX, scrollY, width, height, devicePixelRatio }` |
| `render()` | Hemen yeniden çizer (normalde gerekmez) |
| `destroy()` | Tüm DOM'u ve dinleyicileri kaldırır |

Kullanıcı etkileşimi:

- **Fare:** tıkla (seç), sürükle (aralık; kenarda kendiliğinden kayar), Shift+tık, satır/sütun başlığı (tüm satır/sütun), köşe (tümü). Birleştirilmiş hücre bütün olarak seçilir.
- **Dokunma:** dokun (seç), sürükle (kaydır; atalet tarayıcınındır).
- **Düzenleme:** hücreye yazmaya başlayın (içerik değişir), F2'ye basın ya da çift tıklayın/çift dokunun (mevcut içerik düzenlenir). Enter yazar ve aşağı iner, Shift+Enter yukarı çıkar, Tab sağa, Shift+Tab sola gider, Esc vazgeçer, Alt+Enter satır sonu ekler. Delete ya da Backspace seçili aralığı temizler. Ctrl/Cmd+Z geri alır, Ctrl+Y ya da Ctrl/Cmd+Shift+Z yineler. Ctrl/Cmd+C, X, V kopyalar, keser, yapıştırır.
- **Formül yazarken hücre gösterme:** `=` ile başlayan bir giriş sürerken, imleç formülün başvuru beklediği bir yerdeyse (`=`'in, bir işlecin `+ - * / ^ & < > :`, bir ayırıcının `,` `;` ya da `(`'in hemen ardı; araya boşluk girebilir) bir hücreye tıklamak/dokunmak düzenlemeyi bitirmez, başvurusunu imlecin olduğu yere yazar (`B5`); basılı sürüklemek aralık (`B5:D9`, sürükledikçe güncellenir), sütun/satır başlığına tıklamak tüm sütun/satır (`C:C`, `3:3`; başlıklar üzerinde sürüklemek `B:D`) yazar; Shift+tık son eklenen başvuruyu aralığa uzatır; birleşik hücre sol üst hücresiyle yazılır. Hemen ardından yapılan yeni tıklama son eklenen başvurunun **yerine** geçer; bir şey yazdıktan sonra (ör. `+`, `;`) yapılan tıklama yenisini **ekler**. Eklenen başvurunun hücreleri ızgarada mavi kesikli çerçeveyle gösterilir (yalnız son eklenen; yazmaya devam edince ya da giriş bitince kalkar) ve canlı bölgeden duyurulur ("Başvuru B5:D9"). Başvuru metni dilden bağımsızdır ve her zaman görelidir (`$` yok); ayırıcıyı (`;` ya da `,`) siz yazarsınız. İmleç bir işlenenin ardındayken (`=A1`), tırnak içindeyken ya da giriş `=` ile başlamıyorken tıklama eskisi gibi girişi yazar ve o hücreyi seçer. Formül girişi sürerken (metin `=` ile başlıyorken) başlık kenarlığı sütun/satır **boyutlandırmaz**: fareyle ya da parmakla kenarlığa basış ve çift tıklama yok sayılır, giriş açık kalır, hiçbir şey yazılmaz ve bildirilmez (`=` ile başlamayan girişte eskisi gibi giriş yazılır ve boyutlandırma başlar). Tekerlek ve kaydırma çubuğu başvuru eklemez ve girişi bitirmez; hücredeki süzgeç düğmesi bu sırada açılmaz. Başvuru yeri sayılmayan durumlar: bilimsel gösterimin işareti (`=1E+`), hata sabitinin içi (`=#N/`, `=#DIV/`), tablo başvurusunun köşeli ayraçlarının içi (`=T[[A]:`), `{}` içi, `!` ardı ve ondalık ayırıcısı virgül olan dilde bir rakamın hemen ardındaki virgül (`=1,`; `;` her zaman ayırıcıdır). Salt okunur görünümde ve giriş yöntemi (IME) bileşimi sürerken devreye girmez.
  - **Klavyeyle:** giriş *yazarak* başlatıldıysa ok tuşları, başvuru beklenen yerde, düzenlenen hücreden başlayarak bir hücre gösterir ve başvurusunu yazar/değiştirir; Shift+ok aralığa uzatır, Ctrl/Cmd+ok veri bloğunun kenarına atlar, PageUp/PageDown bir ekran gider. Sayfanın kenarında hareket edemeyen tuş, henüz bir hücre gösterilmemişse hiçbir şey yazmaz (hücre kendine başvurmaz). Giriş bitince (Enter, Esc) düzenlenen hücre, gösterme sırasında kaydırıldıysa yeniden görünüme gelir. Başvuru beklenmeyen yerde oklar eskisi gibi metin imlecini gezdirir (Excel'in aksine girişi bitirmez). Giriş *F2 ya da çift tıklamayla* başlatıldıysa oklar her zaman metin imlecini gezdirir. Düzenleme sırasında **F2** bu iki kip arasında geçirir; düzenleyicinin içine tıklamak imleç kipine alır. Formül çubuğunda oklar her zaman çubuğun metninde gezer.
  - **Başka sayfa:** formül yazarken bir sekmeye tıklamak (ya da hücre düzenleyicisinde Alt+PageDown/PageUp) girişi bitirmez; o sayfa gösterilir ve tıklanan hücre sayfa adıyla yazılır (`Sayfa2!B5`, gerekirse `'Satış 2026'!B5`). Enter/Tab ile onaylayınca ya da Esc ile vazgeçince girişin başladığı sayfaya ve hücreye dönülür; değer o hücreye yazılır, geri alma da oraya döner. Başka sayfa gösterilirken hücre düzenleyicisi görünmez; metin formül çubuğunda izlenir (`formulaBar: false` ile ya da çıplak `WorkbookView`'de o sırada metin ekranda görünmez). Formül girişi sürerken sayfa listesi arayüzden değiştirilemez: **+** düğmesi devre dışıdır (`aria-disabled`), sekme menüsü açılmaz, çift tıklama ve F2 ad alanını açmaz; başka sayfada hücre menüsü de açılmaz. Başka sayfadayken başvuru beklenmeyen bir yerde hücreye tıklamak girişi kendi hücresine yazar ve tıklanan sayfada o hücreyi seçer. `view.setActiveSheet()`, `select()`, `insertSheet()`, `setReadOnly(true)` gibi yöntemler ve odağın görünümden çıkması eskisi gibi önce girişi yazar: değer girişin başladığı hücreye yazılır; yazılamıyorsa (okunamayan formül, veri doğrulamanın durdurduğu değer) o sayfaya dönülür ve düzenleyici metniyle açık kalır. Girişin sayfası bu sırada kitaptan dışarıdan çıkarılmışsa giriş iptal edilir ve `onNotice` ile bildirilir ("Hücrenin sayfası artık yok; yazılan kaydedilmedi.").
  - **Formül çubuğundan:** formül çubukta yazılırken ızgaraya ya da sekmeye tıklamak odağı çubuktan almaz; başvuru çubuktaki imlecin yerine eklenir.
  - **Düzenleyicide Ctrl/Cmd+Z** metin alanının kendi geri almasıdır: Chromium ve Firefox'ta eklenen başvuruyu tek adımda geri alır (sürüklemede aralık her değiştiğinde bir adım oluşur; aynı hücrenin içindeki hareket adım üretmez ve `onEditingChange` her yazımda bir kez çağrılır); WebKit/Safari bir seferde yazılan her şeyi tek adım saydığından orada başvuruyla birlikte ondan önce yazılanlar da gider (giriş açık kalır). Sürükleme sırasında bir şey yazılırsa sürükleme biter.
- **Biçimlendirme:** araç çubuğu ya da Ctrl/Cmd+B, I, U.
- **Sıralama ve süzgeç:** araç çubuğundaki A→Z, Z→A ve Süzgeç düğmeleri. Tek hücre seçiliyken çevresindeki dolu bölge alınır, ilk satır başlık sayılır. Süzgeç açıkken başlık hücrelerindeki düğme (ya da Alt+Aşağı ok) menüyü açar: sırala, ara, değerleri seç, süzgeci temizle. Durum çubuğu "N kayıttan M tanesi" gösterir. Sıralama satırları formülleri ve biçimleriyle taşır (satır yükseklikleri taşınmaz); birleşik hücre, çok hücreli dizi formülü içeren aralık ve yorumlu/tablolu sayfa sıralanmaz. Kaydederken değişmeyen süzgeç bayt bayt korunur; değişen süzgeç Excel'in beklediği biçimde yazılır. API: `view.sort(descending)`, `view.toggleFilter()`, `view.setColumnFilter(col, choice)`. Süzgeçli listede temizle, kopyala/kes, biçimlendir ve satır sil yalnız GÖRÜNEN satırlara uygulanır (Excel gibi); yapıştırma gizli satırlara da yazar; tam sütun biçimlendirmesi gizli satırları da etkiler. Ölçüt sütunu silinince o ölçütün gizlediği satırlar geri gelir; başlık satırı silinince süzgeç kalkar. Sıralamada aksanlı harfler taban harfiyle birlikte sıralanır, başka sayfaya giden başvurular kaymaz; tek hücreyle sıralamada başlık satırı türe ya da kalınlığa göre sezilir. Süzgeçte ve listelerde ı/i ayrımı yerele göredir.
- **Veri doğrulama:** dosyadaki kurallar uygulanır. "Dur" türündeki kurala uymayan değer yazılamaz ve kuralın iletisi gösterilir; "uyarı" ve "bilgi" türünde değer yazılır ve bildirim verilir. Liste kurallı hücrede açılır ok (Alt+Aşağı ok) değerleri gösterir; hücre seçilince giriş iletisi balonda görünür. Yapıştırma doğrulanmaz (Excel de doğrulamaz). Kural ekleme/değiştirme arayüzü yoktur. Formüllü kurallar, yazılan değere bağlı formüller hesaplanmış haliyle değerlendirilir; reddedilen giriş hiçbir iz bırakmaz (dosya bayt bayt aynı kalır). Motorun hesaplayamadığı kural (desteklenmeyen fonksiyon) uygulanmaz.
- **Sağ tık menüsü** (dokunmatikte uzun basma): kes, kopyala, yapıştır, içeriği temizle, satır/sütun ekle-sil. Menüden "Yapıştır" tarayıcı izni ister ve her zaman düz metin yapıştırır; biçim ve formülleriyle yapıştırmak için Ctrl/Cmd+V kullanın.
- **Boyutlandırma:** sütun ya da satır başlığının kenarını sürükleyin (Esc vazgeçer); kenara çift tıklamak sütunu içeriğine sığdırır. Sıfıra sürüklemek gizler.
- **Satır ve sütun:** satır ya da sütun başlığına sağ tıklayın (dokunmatikte uzun basın); açılan menüden ekleyin ya da silin. Tam satır ya da sütun seçiliyken Ctrl/Cmd+`+` ekler, Ctrl/Cmd+`-` siler.
- **Sayfalar:** sekmelerin solundaki **+** düğmesi sona yeni sayfa ekler ve onu gösterir. Sekmeye sağ tıklamak (dokunmatikte uzun basmak; klavyede odaktaki sekmede Shift+F10 ya da Menü tuşu) menüyü açar: "Sayfa ekle" (o sekmenin sağına), "Yeniden adlandır", "Sil". Sekmeye çift tıklamak/çift dokunmak ya da odaktaki sekmede F2 adı yerinde düzenletir: Enter ya da alandan çıkmak uygular, Esc vazgeçer. Silmede onay sorulmaz; Ctrl/Cmd+Z geri getirir. Salt okunur görünümde düğme, menü ve ad alanı yoktur. Ayrıntı aşağıda.
- **Klavye** (önce görünüme tıklayın ya da Tab ile odaklanın): oklar, Shift+oklar, Ctrl/Cmd+oklar (veri bloğunun kenarı), Home, Ctrl+Home, Ctrl+End, PageUp/PageDown, Enter/Shift+Enter, Ctrl+A; sayfa değiştirmek için Alt+PageDown/PageUp (Ctrl+PageDown/PageUp da çalışır, ama bazı tarayıcılar onu kendine ayırır). Tab ve Shift+Tab hücre değiştirmez: odağı ızgaradan sayfadaki sonraki/önceki öğeye taşır, böylece klavye kullanıcıları görünümde hapsolmaz. Sayfa sekmeleri Tab ile odaklanıp Enter ya da Boşluk ile seçilir; Tab sırası: ızgara → **+** düğmesi → sekmeler.

Bilinmesi gerekenler:

- Hücreler dosyadaki sayı biçimiyle gösterilir. `[Red]` gibi renk bölümleri yazı rengini değiştirir; biçimiyle gösterilemeyen değer (ör. negatif tarih) ve sütuna sığmayan sayı `####` olarak çizilir.
- Çizimde hücrenin kendi fontu kullanılır; cihazda yoksa `Carlito`, `Calibri`, `Arial` sırasıyla yedeğe düşülür. Font farkı metin genişliklerini biraz değiştirebilir.
- Sütun genişlikleri ve satır yükseklikleri Excel'in piksel yuvarlamasıyla hesaplanır (varsayılan sütun 64 px, satır 20 px).
- Sayfa her zaman 1.048.576 satır × 16.384 sütundur; kaydırma alanı kullanılan alanla başlar ve sona yaklaşıldıkça büyür.
- Erişilebilirlik: klavye odağı görünümdeki gizli bir metin alanındadır (`aria-label="Cell editor"`); kaydırılan bölüm `role="grid"` taşır, sekmeler `role="tab"` düğmeleridir; etkin hücrenin adresi ve metni ekran okuyucular için bir canlı bölgede duyurulur. Ekran okuyucuyla henüz denenmedi.
- CSS ile ölçeklenmiş (`transform: scale(...)`) bir kapsayıcıda da tıklamalar doğru hücreye gider.
- Görünüm kendi stillerini sayfaya tek bir `<style data-bymoffice>` öğesiyle ekler; sınıf adları `bym-` ile başlar. Dosyadan gelen hiçbir metin HTML olarak basılmaz.

### Kaydetme

```js
const mounted = await BymOffice.mount(el, { file, onSave: (bytes) => upload(bytes) }); // Ctrl+S / Cmd+S
const bytes = await mounted.save();   // istediğiniz an; Uint8Array
```

`save()` (hem `mounted.save()` hem `file.save()`) düzenlemeleri dosyaya yazar ve baytlarını döndürür. Nasıl çalıştığı:

- **Yalnız değişen yeniden yazılır.** Değiştirmediğiniz sayfalar, görseller, tema, makrolar ve motorun tanımadığı her parça özgün baytlarıyla kalır. Değişen bir sayfada bile dokunmadığınız satırlar özgün XML'iyle, bayt bayt aynen kopyalanır.
- **Hiçbir şey değişmediyse** kaydedilen dosya açılanla bayt bayt aynıdır.
- Sayfanın modelden yazılmayan bölümleri (koşullu biçimler, veri doğrulama, sayfa yapısı/yazdırma ayarları ve yazıcı ayarı parçası, uzantılar) özgün baytlarıyla aynen korunur; satır/sütun eklenip silindiyse içlerindeki hücre başvuruları, resimlerin konumları ve yorumlar birlikte kayar.
- Yeni biçimler stil tablosuna yalnız **eklenir**; var olan stil kayıtları değişmez.
- Hesaplanan formül sonuçları dosyaya yazılır. Yine de dosyaya "açılışta yeniden hesapla" işareti konur ve eski hesap zinciri (`calcChain`) kaldırılır; Excel dosyayı açınca tüm sonuçları kendi motoruyla günceller.
- Düzenlediğiniz metin hücreleri satır içi metin (`inlineStr`) olarak yazılır; dokunulmayan hücreler paylaşılan metin tablosundaki yerlerini korur.
- Aynı çalışma kitabı art arda kaydedilebilir; kaydetmeden sonra geri al/yinele çalışmayı sürdürür.

Doğrulama: her değişiklikte, gerçek dosyalardan oluşan bir havuzda "düzenle → kaydet → yeniden aç → modelle karşılaştır" ve "dokunulmayan parçalar bayt bayt aynı" denetimi çalışır; kaydedilen dosyalar ayrıca bağımsız bir okuyucuyla (openpyxl) karşılaştırıldı. Excel'in kendisiyle otomatik deneme yapılmadı: kaydettiğiniz dosyayı Excel'de açıp gözle denetlemeniz önerilir.

Değişen satırlar yeniden üretilirken şunlar o satırlarda korunmaz (dokunulmayan satırlarda korunur): `spans` ipucu, hücrelerin fonetik (furigana) metni, zengin metinde `outline`/`shadow` gibi nadir font ayrıntıları.

### Eski biçimleri düzenleme: `convert`

```js
const mounted = await BymOffice.mount(container, { file: docBytes, convert: true, locale: 'tr' });
if (mounted.conversion) {
  // { from: 'doc', to: 'docx', warnings: [...] }: dosya artık docx; sunucuda adını ve türünü buna göre saklayın
  console.log(mounted.file.format);         // 'docx'
}
const bytes = await mounted.save();          // docx baytları
```

- doc ve RTF docx'e, xls xlsx'e çevrilir (`saveAsDocx()` / `saveAsXlsx()` ile aynı yazıcılar); özgün dosya değişmez, `save()` yeni biçimi verir. Uygulamanız dosya adını ve türünü `mounted.conversion.to`'ya göre güncellemelidir.
- **Modelde bulunup** yeni dosyaya taşınamayanlar `mounted.conversion.warnings` içindedir (İngilizce metinler; her biri kaç kez olduğunu söyler). Hiçbir şey kaybolmadıysa dizi boştur. Belgede (doc, RTF → docx) uyarıya düşenler:
  - şekiller: dikdörtgen, yuvarlatılmış dikdörtgen, elips ve düz çizgi geometrisi, düz renk dolgusu, çizgisi (renk, kalınlık, kesiklik), çevirme ve döndürmesiyle yazılır ve uyarı çıkmaz. Uyarıya düşenler (her biri sayısıyla):
    - okuyucunun geometrisini tanımadığı şekil (ok, serbest biçim…) dolgusu ve çizgisiyle dikdörtgen olarak yazılır (`N shape(s) were written as rectangles: their geometry was not carried over`);
    - görünümü modelde bulunmayan şekil (RTF'teki Word 95 çizim nesnesi, çizim verisinde karşılığı olmayan doc şekli, resmi bulunamayan resimli metin kutusu) çizgisiz ve dolgusuz dikdörtgen olur (`N shape(s) were written as plain rectangles…`);
    - gradyan, desen, doku ya da resim dolgulu şekil dolgusuz yazılır (`N shape(s) with a fill that is not a plain colour… were written without a fill`);
    - kısmi saydam dolgu opak yazılır (`N shape(s) with a partly transparent fill were written with an opaque one`);
    - palet, şema ya da sistem dizinli renk çözülmez: dolgu yazılmaz, çizgi siyah yazılır (`N shape(s) with a colour that is an index into a palette or a system colour…`);
    - gruplarda parçalar grubun bütün alanına, çizgisiz ve dolgusuz yerleştirilir (`N group(s) of shapes…`);
  - baytları modelde bulunmayan görseller ve resmi dosyada bulunamayan resim şekilleri boş çerçeve olarak yazılır (`N picture(s) whose data is not available…`);
  - XML'in taşıyamadığı karakterler (denetim karakterleri, eşi olmayan vekil kodlar) atılır (`N character(s) that a docx file cannot hold were left out`);
  - izlenmeyen adresli bağlantılar düz metne çevrilir ve tema başvuruları düz değerlere çevrilir. Bu iki uyarı doc ve RTF'te bugün çıkmaz (bu okuyucular bağlantıyı alan olarak, temayı boş kurar); bir docx modeli `buildDocx` ile yazılırken görülür.

  Tabloda (xls → xlsx) `saveAsXlsx()` raporundaki uyarıların aynısı gelir (aşağıda "xls dosyasını xlsx olarak kaydetme"): içeriği modelde olmayan grafik ve makro sayfaları, baytları okunamayan görseller (metafile), başka kitaba başvuran formüller vb.
- Okuyucunun modele **hiç almadığı** içerik ("Henüz olmayanlar" tablosu: ör. şekillerde ok başları ve gölge, doc'ta EMF/WMF görseller, xls'te grafikler ve yazıcı sürücüsü verisi) yeni dosyada yoktur ve **uyarılmaz**; bu yüzden **arşivlik bir dosyanın özgün halini saklamaya devam edin**.
- Yeni dosya yazılamaz ya da geri okunamazsa `mount()` reddetmez: dosya dönüştürülmeden, olduğu gibi **salt okunur** gösterilir, `conversion` bulunmaz ve `onNotice` `kind: 'error'` ile "çevrilemedi; salt okunur gösteriliyor" bildirir. (Böyle bir dosya bilinmiyor. Geri düşüş, ortak dönüştürme işlevinde ve WebView köprüsünde yapay bir hata üretilerek sınandı; `mount()` üzerinden bu yol sınanmadı.)
- `convert` verilmezse ya da `readOnly: true` ise dosya olduğu gibi, salt okunur gösterilir ve `save()` özgün baytları verir. Bu durumda düzenleme **açılamaz**: doc, RTF ve xls görünümlerinde `view.setReadOnly(false)` ve `editor.setReadOnly(false)` etkisizdir (düzenlemeler kaydedilemeyeceği için).
- Örnek sayfada `convert=1` parametresiyle kullanılır; kaydetme isteği biçimi `X-Bym-Format` başlığında bildirir, üst sayfaya giden iletilerde `format` ve `convertedFrom` bulunur. Adın sonundaki `.doc`, `.rtf` ya da `.xls` (büyük/küçük harf fark etmez) yeni uzantıyla değiştirilir; ad böyle bitmiyorsa (`Rapor 2024.10`, `teklif.v2`) yeni uzantı sona eklenir. WebView köprü sayfasında aynı işi `load.convert` yapar ([webview-koprusu.md](webview-koprusu.md)).
- Dönüştürmeyi görünüm kurmadan yapmak için: `BymOffice.convertForEditing(file)` → `{ from, to, bytes, warnings }` (dönüştürme gerekmeyen dosyada `undefined`); `BymOffice.openConverted(file, BymOffice.open)` dönüştürüp yeni dosyayı açar ve `{ file, conversion?, failed? }` verir.

### Yeni (boş) dosya: `createBlank(kind, options?)`

```js
const docx = await BymOffice.createBlank('docx', { locale: 'tr' });  // Uint8Array
const xlsx = await BymOffice.createBlank('xlsx', { locale: 'tr', sheetName: 'Bütçe' });
const mounted = await BymOffice.mount(container, { file: docx, locale: 'tr' });
```

- `'docx'`: tek boş paragraflı belge; Normal, Başlık 1–3, Konu Başlığı (Title) ve Köprü stilleri; Calibri 11. `locale` bir dil etiketidir (`tr`, `tr-TR`, `en-GB`; büyük/küçük harf ve `_` fark etmez): Türkçe A4 ve Türkçe ayırıcılar, `en`/`en-US`/`en-CA` Letter, diğer bütün diller A4 verir; etiket değilse ya da verilmezse `en-GB` sayılır.
- `'xlsx'`: tek boş sayfalı kitap; sayfa adı Türkçede `Sayfa1`, diğerlerinde `Sheet1`, ya da `sheetName` (1–31 karakter; `[ ] : * ? / \` ve denetim karakteri içeremez, yalnız boşluk olamaz, kesme işaretiyle başlayıp bitemez, `History` olamaz; aksi halde `RangeError`).
- Aynı seçenekler her zaman aynı baytları verir. Başka bir tür istenirse `RangeError`.
- Dosyalar yeniden okunarak ve bağımsız bir okuyucuyla (xlsx) açılarak sınandı; Microsoft Word/Excel'de açılarak sınanmadı.

### Bir uygulamaya gömme: `examples/html/index.html`

Sunucu tarafı ne olursa olsun (PHP, .NET, Node…) çalışan, çatı gerektirmeyen tek sayfa: dosyayı sitenin kendi adresinden alır, editörde gösterir, kaydedilince baytları `PUT` ile geri yollar. Sekmede ya da `<iframe>` içinde açılır:

```
examples/html/index.html?file=/files/42/download&save=/files/42&name=rapor.xlsx&lang=tr
examples/html/index.html?file=/files/42/download&mode=view        salt okunur
examples/html/index.html?new=docx&save=/files&name=yazi.docx      yeni, boş belge
```

| Parametre | İş |
|---|---|
| `file` | Dosyanın alınacağı adres (sayfayla aynı kökten olmalı) |
| `save` | Kaydedilen baytların `PUT` ile yollanacağı adres (aynı kök); verilmezse "Kaydet" dosyayı indirir |
| `mode` | `edit` (varsayılan) ya da `view` |
| `lang` | `tr` ya da `en`: editörün ve yeni dosyanın dili |
| `new` | `docx` ya da `xlsx`: dosya almak yerine boş dosyayla başlar |
| `convert` | `1`: doc, RTF ve xls dosyalarını docx/xlsx'e çevirerek düzenlenebilir açar; ad yeni uzantıyı alır, `PUT` isteği biçimi `X-Bym-Format` başlığında bildirir |
| `name` | Gösterilen ve indirmede kullanılan dosya adı |

- Kaydet düğmesi ve Ctrl/Cmd+S aynı işi yapar. Dosya yoldayken yazılan şey "kaydedildi" sayılmaz: durum "Kaydedilmedi" kalır ve `saved` iletisi gitmez. Kaydedilmemiş değişiklik varken sayfadan çıkılırsa tarayıcı sorar. doc, xls ve RTF kendi biçimlerine yazılamadığı için salt okunur gösterilir ve Kaydet gizlenir.
- `<iframe>` içindeyken sayfa üst sayfaya (yalnız aynı kökte) `{ source: 'bymoffice', type: 'ready' | 'changed' | 'saved' | 'error', message?, format?, convertedFrom? }` iletileri yollar.
- Başka kökten adresler reddedilir: sayfa yalnız kendi sitesinden dosya alır ve yalnız oraya yollar. Sunucunuz `PUT` isteğinde oturumu ve yetkiyi kendi denetlemelidir; CSRF belirteci gerekiyorsa `fetch` çağrısına başlık olarak ekleyin.
- Sayfayı kendi projenize kopyalayıp `<script src>` adresini kendi sunucunuza ya da CDN'e çevirin. React ve Vue için ayrı örnek yoktur; `mount()` çağrısı bir `ref`'e bağlanmış kapsayıcıyla aynı şekilde kullanılır ve bileşen kaldırılırken `mounted.destroy()` çağrılır (bu kullanım sınanmadı).

### xls dosyasını xlsx olarak kaydetme

```js
const file = await BymOffice.open(xlsBytes);      // format: 'xls'
const xlsxBytes = await file.saveAsXlsx();        // yeni bir xlsx; file.save() özgün xls baytlarını vermeye devam eder
```

`saveAsXlsx()` xlsx'te `save()` ile aynıdır; xls'te tablo modelini **sıfırdan yeni bir pakete** yazar (`BymOffice.buildXlsx(workbook, { report })`). Aynı işlev kodla kurulmuş bir `Workbook` için de kullanılabilir. `mount()` davranışı değişmez: xls salt okunur gösterilir.

Yazılanlar: sayfa adları, sırası ve durumu (gizli, çok gizli), etkin sayfa, 1904 tarih sistemi; hücre değerleri ve türleri; formül metinleri ile dosyadaki sonuçları (paylaşılan ve dizi formülleri, veri tabloları öyle kalır); stiller (yazı tipi, dolgu, kenarlık, hizalama, sayı biçimi, koruma, hücre stilleri, renk paleti, tema renkleri); zengin metin; birleştirmeler; sütun genişlikleri, satır yükseklikleri, gizli/özetlenmiş satır ve sütunlar; görünüm (dondurulmuş bölmeler, etkin hücre, yakınlaştırma, kılavuz çizgileri, sağdan sola, sekme rengi); köprüler; tanımlı adlar; yorumlar (yazar, zengin metin, kutunun yeri, hep görünür olması); görseller, şekiller ve metin kutuları, gruplar; sayfa yapısı (`sheet.pageSetup`: kenar boşlukları, kâğıt, yön, ölçek/sığdırma, ilk sayfa numarası, sayfa sırası, üstbilgi/altbilgi, ortalama, kılavuz çizgisi/başlık yazdırma, sayfa sonları; yazdırma alanı ve yinelenen başlıklar tanımlı ad olarak). Model bir xlsx'ten okunmuşsa koşullu biçimler, veri doğrulamaları ve otomatik süzgeç de yazılır.

**Modelde bulunup** yeni dosyaya yazılamayan ya da değiştirilerek yazılan içerik `report.warnings` içinde türüne göre, kaç kez olduğuyla birlikte listelenir (aşağıdaki maddeler). Modelde **hiç bulunmayan** şeyler için uyarı çıkmaz: xls okuyucusunun okumadığı içerik ("Henüz olmayanlar" tablosu) ve bir xlsx'in modele alınmayan parçaları (yazıcı ayarı parçası, sayfa koruması, belge özellikleri) sessizce yoktur; tek istisna, modelin varlığını bildiği tablo/pivot/form denetimi gibi sayfaya bağlı parçalardır, onlar uyarılır. Kodla kurulmuş geçersiz bir modelde biçimsel düzeltmeler (yinelenen sayfa kimliği, sınır dışı yakınlaştırma ya da özet düzeyi) uyarısız yapılır.

```js
const report = { parts: [], warnings: [] };
const bytes = await BymOffice.buildXlsx(await file.workbook(), { report });
// report.warnings: ['Drawing: a chart is not in the model and is left out', ...]
```

- **Grafik ve makro sayfaları** modelde içeriksizdir; adları ve sıraları korunsun diye **boş çalışma sayfası** olarak yazılır. Sayfadaki grafik, diyagram ve gömülü nesnelerin yer tutucuları yazılmaz.
- **Baytları okunamayan görseller** (xls içindeki EMF/WMF) resimsiz çerçeve olarak yazılır.
- **Hedefi olmayan köprüler** (okuyucunun tanımadığı türde bağlantı), değeri olmayan ya da olmayan bir sayfaya bağlı adlar, sayfa sınırı dışındaki hücre/birleştirme/yorumlar yazılmaz.
- **Başka dosyaya başvuran formüller** metinleriyle (`[1]Sayfa1!A1`) ve son değerleriyle yazılır; xls okuyucusu dış dosyanın yolunu modele almadığı için bağlantının kendisi (`externalLink`) yazılmaz; böyle her formül ve tanımlı ad sayılarak uyarılır.
- **Satır/sütun eklenip silinmiş model** (xlsx'ten okunup `applyOp` ile düzenlenmiş): yorumlar, koşullu biçimler, veri doğrulamaları ve çizim çapaları `save()`'in uyguladığı kurallarla kaydırılarak yazılır.
- **Dosyanın ifade edemeyeceği değerler düzeltilir ve uyarılır:** tabloda olmayan biçim dizini (varsayılan biçim), ilk iki dolgu (`none`, `gray125` başa eklenir, dizinler kayar; uyarısız, çünkü biçimler aynı kalır), sonlu olmayan sayı (`#NUM!`), formül/hata/ad içindeki denetim karakterleri (U+FFFD), çakışan sütun aralıkları, geçersiz dizi formülü aralığı, metinsiz formül, aynı kapsamda ikinci ad, aynı hücrede ikinci köprü, gizli etkin sayfa (ilk görünür sayfa), satırın yazıcıya ait ya da başka ad alanından öznitelikleri. Köprü hedefindeki denetim karakterleri yüzde kodlamasıyla yazılır.
- **Sayfa yapısı:** `sheet.pageSetup` şema sırasıyla yazılır (`sheetPr/pageSetUpPr`, `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`, `rowBreaks`, `colBreaks`). Dosyanın ifade edemeyeceği değerler yazılmaz ve sayılarak uyarılır: aralık dışı sayı (ör. 10–400 dışı ölçek, negatif ilk sayfa numarası), tanınmayan ad, sayfa dışındaki ya da aynı yerdeki ikinci sayfa sonu; uzunluk olmayan kenar boşluğu varsayılanla (0,7 / 0,75 / 0,3 inç) yazılır. Üstbilgi/altbilgi metni olduğu gibi yazılır (255 karakterden uzun olsa da; XML'in taşıyamadığı karakterler `_xHHHH_` kaçışıyla). Bir yönde 1026'dan çok sayfa sonu yazılır ama uyarılır. **Taşınmayanlar (uyarısız, çünkü modelde yok):** yazıcı sürücüsüne özgü ikili veri (xls'te PLS kaydı, xlsx'te `pageSetup r:id` ile bağlı yazıcı ayarı parçası), özel görünümlerin kendi sayfa yapısı, üstbilgi/altbilgi görselleri (`&G`).
- xlsx'ten okunmuş modellerde: 2010 uzantısına özgü koşullu biçim kuralları, renge/ikona göre süzgeç ölçütleri, konuşma biçimindeki yorumların yanıtları, hücre meta verisi (dinamik dizi) ve modelde hiç bulunmayan parçalar (tablolar, pivotlar, yazıcı ayarı parçası) yazılmaz. Bir xlsx'i kaydetmek için `save()` kullanın; o hiçbir şeyi kaybetmez.
- xls okuyucusunun modele almadığı her şey ("Henüz olmayanlar" tablosu) doğal olarak yeni dosyada da yoktur.

Aynı model her zaman aynı baytları verir. Üretilen dosyalar yeniden okunarak (içerik karşılaştırması), şema sırası denetlenerek ve gerçek xls dosyalarıyla sınandı; **Microsoft Excel'de açılarak sınanmadı**.

### Word belgesi modeli

```js
const file = await BymOffice.open(bytes);          // file.format === 'docx'
const doc = await file.document();                 // DocModel
for (const block of doc.body.blocks) {
  if (block.type === 'paragraph') console.log(BymOffice.paragraphText(block));
}
```

- Model: gövde blokları (paragraf, tablo), çalıştırmalar, alanlar, köprüler, yer imleri, çizimler ve görseller, stiller, numaralandırma, tema, bölümler, üstbilgi/altbilgi, dipnotlar, yorumlar. Birimler dosyadaki gibidir (twips, yarım punto, EMU).
- Modelde hiçbir şey miras alınmaz; etkin biçimi `BymOffice.resolveParaProps(doc, props)` ve `BymOffice.resolveRunProps(doc, paraProps, runProps)` hesaplar. Liste sayaçları için `new BymOffice.NumberingCounters(...)`.
- Köprü hedefini bağlamadan önce `BymOffice.safeHyperlinkTarget(target)` kullanın: yalnız `http(s)`, `mailto`, `tel` ve belge içi `#yerimi` geçer.
- Anlaşılmayan öğeler modelde saklanır; her üst düzey bloğun özgün XML'i `BymOffice.rawOf(story, block)` ile alınır.
- **Kaydetme:** docx'te `file.save()` modelde düzenlenen blokları dosyaya yazar; düzenlenmeyen bloklar ve parçalar bayt bayt korunur (`BymOffice.writeDocx(pkg, doc)`). doc ve RTF kendi biçimlerine yazılamaz; `file.saveAsDocx()` modeli yeni bir docx olarak verir (`BymOffice.buildDocx(doc)`). `buildDocx(doc, { report })` taşınamayanları `report.warnings` içinde sayılarıyla bildirir: şekiller geometri, dolgu ve çizgileriyle DrawingML olarak yazılır; geometrisi tanınmayan şekil dikdörtgen, görünümü modelde olmayan şekil çizgisiz/dolgusuz dikdörtgen olur, düz renk olmayan dolgu yazılmaz, kısmi saydamlık ve dizinli renkler taşınmaz (hepsi sayısıyla uyarılır), gruplarda parçalar grubun bütün alanına yazılır, baytı bulunmayan görsel boş çerçeve olur, `javascript:` gibi izlenmeyen adresli bağlantılar düz metne çevrilir, tema başvuruları (yalnız docx'ten okunmuş bir modelde bulunur) düz renk ve yazı tipi olarak yazılır; XML'in taşıyamadığı, atılan karakterlerin sayısı `report.droppedCharacters` içindedir. Üretilen dosyalar yeniden okunarak ve şema sırası denetlenerek sınandı; Microsoft Word'de açılarak sınanmadı.
- **Düzenleme işlemleri** (`BymOffice.docOps`): `insertText`, `deleteRange`, `splitParagraph`, `setRunFormat`, `setParaFormat`, `toggleList`, `insertTable`, `insertImage`, `insertHyperlink`… her biri `DocEdit` döndürür; `docOps.applyDocOps(doc, edit.ops)` uygular, `docOps.DocHistory` geri al/yinele sağlar.

### Word belgesini gösterme ve düzenleme

```js
const mounted = await BymOffice.mount(container, { file: docxBytes, locale: 'tr', zoom: 'fit-width' });
if (mounted.kind === 'document') {
  mounted.view.goToPage(3);          // DocumentView
  console.log(mounted.view.pageCount);
  mounted.view.insertText('Merhaba'); // imlecin olduğu yere
  const bytes = await mounted.save(); // değişikliklerle birlikte docx
}
```

- **docx varsayılan olarak düzenlenebilir açılır** (`readOnly: true` ile salt okunur); doc ve RTF, `convert: true` verilmedikçe her zaman salt okunurdur, çünkü kendi biçimlerine yazılamazlar; bu görünümlerde `setReadOnly(false)` düzenlemeyi açmaz. `mounted.save()` docx'te değişiklikleri içeren baytları verir ve belgeyi "kaydedildi" sayar (`view.isModified` → `false`); yazma sürerken yapılan değişiklik kaydedilmiş sayılmaz.
- `mount()` dönüşünde `kind` alanı vardır: `'spreadsheet'` (`workbook`, `view`, `editor`) ya da `'document'` (`document`, `view`, `editor`: `DocumentEditor`, araç çubuğuyla kabuk). TypeScript'te `workbook`/`editor` için `kind` ile daraltın (0.x'te tip düzeyinde kırıcı değişiklik).
- Seçenekler: `zoom` (sayı, `'fit-width'`, `'fit-page'`), `toolbar` (sayfa göstergesi ve yakınlaştırma çubuğu), `readOnly`, `onDocumentChange()` (belge her değiştiğinde, geri al ve yinele dahil), `onPageChange(page, count)`, `onLinkClick(target)` (`false` dönerse bağlantı açılmaz; yalnız http, https, mailto, tel hedefleri bağlantı olur), `locale`, `onNotice`, `onSave`.
- `DocumentView`: `pageCount`, `currentPage`, `layoutDone`, `zoom`, `setZoom()`, `zoomIn()`, `zoomOut()`, `goToPage(n)`, `selectAll()`, `getSelectedText()`, `setLocale()`, `focus()`, `print(options?)`, `printing`, `destroy()`.
- `DocumentView` düzenleme üyeleri: `readOnly`, `setReadOnly()`, `isModified`, `markSaved()`, `changeCount`, `canUndo`, `canRedo`, `undo()`, `redo()`, `getSelection()`, `setSelection(konum ya da aralık)`, `insertText()`, `format({ bold, italic, font, size, color… })`, `toggleFormat('bold')`, `formatParagraph({ align, style… })`, `toggleList('bullet' | 'number')`, `indent(±1)`, `insertTable(satır, sütun)`, `insertRow()`, `deleteRow()`, `insertColumn()`, `deleteColumn()`, `insertPageBreak()`, `insertLink(adres)`, `insertPicture(bytes)`, `selectionFormat`, `inTable`, `commitEditing()`. Konum `{ story: 'body', path: [blok dizini…], offset }` biçimindedir.
- **Pano:** kopyalanan seçim sistem panosuna düz metin ve (kaçışlanmış metinden üretilen) basit HTML olarak konur; dosyadan gelen hiçbir şey panoya işaretleme olarak yazılmaz, panodan gelen HTML hiç ayrıştırılmaz. Kopyanın kendisi (**zengin kopya**) sayfanın belleğinde tutulur ve HTML'deki bir işaretle eşlenir: yapıştırılan HTML bu işareti taşıyor **ve** düz metin o kopyanın metniyle aynıysa bellekteki kopya yapıştırılır, değilse panodaki düz metin. Böylece kopyala/kes → yapıştır **aynı düzenleyicide ve aynı sayfadaki başka bir BYM belge düzenleyicisinde** şunları taşır:
  - **Metin ve biçimi**, sekme, satır sonu, **sayfa sonu, sütun sonu**, yumuşak tire, bölünmez tire, simge (`w:sym`).
  - **Görseller** (satır içi ve kayan) ve metinsiz çizimler: çizimin saklanan XML'i yalnız görsellere ve dış adreslere başvuruyorsa. Aynı belgede var olan medya parçası ve ilişki yeniden kullanılır; başka belgede parça `DocModel.media`'ya boş bir adla (`/word/media/imageN.ext`) eklenir, öykünün parçasına ilişki eklenir, aynı görselin ikinci yapıştırması aynı parçayı kullanır. Her yapıştırılan çizim belgede benzersiz bir `wp:docPr id` alır.
  - **Bağlantılar** metinleriyle: dış adres ilişkisi hedef parçada aranır, yoksa eklenir; belge içi hedef (`w:anchor`) olduğu gibi kalır (hedef belgede o yer imi olmayabilir). Bağlantının yalnız bir bölümü kopyalanırsa o bölüm de bağlantıdır.
  - **Alanlar:** basit alan (`w:fldSimple`) komutuyla; karmaşık alan, başından sonuna seçimdeyse (ya da sonucunun tamamı seçimdeyse) kodu ve sonucu ile. Sonucunun yalnız bir bölümü seçilen alanın o metni düz metin olarak gider.
  - **Tablolar:** seçim tabloyu bütünüyle kapsıyorsa tablo; tablonun içinde hücreler seçiliyse o hücrelerin kapladığı sütunlardan oluşan tablo (tam satırlar seçiliyse satırlar olduğu gibi); seçim tabloya girip çıkıyorsa silmenin boşaltacağı hücreleri taşıyan satırlar (o satırların öbür hücreleri boş). Izgara, hücre/satır/tablo özellikleri, birleşmeler ve hücrelerin içindeki her şey (iç içe tablolar dahil) taşınır. **Yapıştırma kuralı:** tablo, imlecin olduğu yere blok olarak girer; imleç bir hücredeyse **o hücrenin içine iç içe tablo** olarak eklenir (hücrelerin üzerine yazılmaz). Yapıştırma sonucu 16 düzeyden derin iç içe tablo oluşacaksa reddedilir. İki tablo yan yana gelecekse ya da tablo listenin sonunda kalacaksa araya boş paragraf konur. **Tablo sınırını aşan kesme bir taşıma değildir:** silme, tablonun satırlarını kaldırmaz, yalnız hücrelerini boşaltır; bu yüzden böyle bir seçimi kesip aynı yere yapıştırınca özgün tablo boşalmış hücreleriyle yerinde kalır ve kesilen hücreler **yeni bir tablo** olarak gelir (içerik yitmez, tablo yapısı yinelenir). Bütünüyle seçilen tablo ise kesilince kalkar ve yapıştırınca tek tablo olarak geri gelir. Tabloların, satırların, hücrelerin ve bütün olarak taşınan paragrafların öğe öznitelikleri (`w:rsidR` gibi) korunur; bir öğeyi adlandıran ve belgede benzersiz olması gereken `w14:paraId` ve `w14:textId` yapıştırmada **her zaman düşürülür** (aynı kopya iki kez yapıştırılabildiği için; Word kimliği olmayan öğeye yenisini verir).
  - **Paragraf özellikleri, stil ve liste:** sonu (paragraf işareti) kopyalanan her paragraf kendi özellikleriyle gelir; imlecin olduğu paragrafın imleçten önceki bölümü ilk yapıştırılan paragrafın biçimini alır, imleçten sonraki bölümü kendi biçiminde kalır (Word'deki gibi). **Tek ayrık durum, kesilenin geri konmasıdır:** bir paragraftan öbürüne uzanan seçim silinince iki paragraf birleşir ve yalnız birinin biçimi kalır (ilki boş değilse ilkinin). Kopya, seçimin bittiği paragrafın özelliklerini de taşır; **kesilerek** alınmış bir kopya, biçimi kesmenin bıraktığı biçimle aynı olan bir paragrafa (öncelikle kesildiği yere) yapıştırılınca imleçten sonraki bölüm o özellikleri geri alır. Böylece kes → aynı yere yapıştır, her paragrafı stili, listesi, hizalaması ve saklanan özellikleriyle eski haline getirir. Kopyalanmış (kesilmemiş) içerikte ve başka biçimdeki bir paragrafta kural değişmez: imleçten sonrası kendi biçiminde kalır. **Aynı belgede** stil ve liste kimlikleri korunur. **Başka belgede:** aynı kimlikli stil varsa hedefin stili kullanılır ("hedef stilleri kullan"); yoksa tanımı, dayandığı stillerle (`basedOn`, `link`, `next`) birlikte hedefe kopyalanır (varsayılan stil işareti olmadan); liste her zaman yeni bir liste olarak (yeni `numId`/`abstractNumId`) eklenir. Belge varsayılanları, tema ve font tablosu taşınmaz: tema rengi/fontu hedefin temasına göre çözülür. Bölüm sonu taşıyan paragraf, gövdeye (tablo dışına) yapıştırılınca bölüm sonunu korur — başka belgede üstbilgi/altbilgi başvuruları olmadan; hücreye ya da üstbilgiye yapıştırılınca bölüm sonu düşer.
  - İmleç bir bağlantının ya da içerik denetiminin içindeyse yapıştırılan metin onun içinde kalır (denetim bölünmez); kopyadaki bağlantı, içine yapıştırıldığı bağlantıya katılır.
  - Kopyadaki saklanan XML (tanınmayan biçim öğeleri, çizim) hedef parçanın kökünde aynı biçimde bildirilmemiş bir ad alanı öneki kullanıyorsa bildirimi öğenin kendisine eklenir. Başka belgeye kopyalanan **stil ve liste tanımları** için ölçü, gittikleri parçanın (`styles.xml`, `numbering.xml`) köküdür, gövdeninki değil: o kökte aynı bildirim yoksa (ya da parça henüz yoksa) öneki kullanan öğe, özniteliklerdeki önekler dahil, bildirimi kendi taşır. Strict ↔ Transitional belgeler arasında ve WordprocessingML'i başka bir önekle yazan belgeye zengin yapıştırma **reddedilir** (bildirim: "What was copied cannot be pasted here"); kopya bellekte kalır, kaynak belgeye geri yapıştırılabilir.
  - **Taşınmayanlar** (kopyada yalnız metinleri gider): dipnot/sonnot başvurusu, yorum başvurusu, denklem, içerik denetimi, izlenen ekleme, metin kutusu/metinli şekil, grafik, diyagram, OLE nesnesi, VML resim, `mc:AlternateContent`, ilişkisi bulunamayan görsel, birden çok paragrafa yayılan alan, akıllı etiket/özel XML, bir kopyadan fazlası olan sarmalayıcılar. **Kesme kuralı:** kesmeden önce silme denenir, belgeden çıkacak her şey türüne göre sayılır ve kopyada aranır; kopyada olmayan bir şey gidecekse ya da panoya yazılamıyorsa **kesme yapılmaz**, `onNotice` ile bildirilir (Delete/Backspace ile silinebilir). Paragraf biçimi de sayılır: birleşen ya da kalkan paragrafların özellikleri kopyada olmalıdır; zengin kopya bunu her zaman taşır, yalnız düz metin taşıyan bir kopya (çok büyük seçim) taşımaz, bu yüzden biçimli bir paragrafın sonunu aşan öyle bir kesme reddedilir. Tarayıcı panoya HTML yazdırmazsa (kopyanın işareti panoya ulaşmazsa) kesme düz metin kesmesi olarak değerlendirilir: metinden fazlası gidecekse yapılmaz. Silmenin yerinde bıraktıkları (yer imi, yorum işareti, seçimin yalnız ucuna değen alan) kesmeye engel değildir. Blok düzeyinde içerik denetiminin içine giren bir seçim de kesilemez.
  - **Tek geri alma adımı:** yapıştırma, eklediği medya parçası, ilişki, stil ve liste tanımlarıyla birlikte tek adımdır; geri alınınca hepsi modelden çıkar ve kaydedilen pakette yetim parça kalmaz.
  - **Bellek:** sayfa son 8 kopyayı, toplam 64 MB'a kadar (görsel baytları dahil) tutar; eskiler atılır. Paketten okunan görsel baytları kopyadan hemen sonra eşzamansız okunur; baytlar gelmeden başka belgeye yapıştırılırsa bu bildirilir ve yapıştırma baytlar gelince, **istendiği yere** yapılır: bu arada imleç taşınmış, seçim ya da belge değişmişse yapıştırma yapılmaz ve bildirilir (yeniden yapıştırmak yeter). Sınırı aşıp görsel baytları atılan kopyaya sonradan okunan baytlar da eklenmez. Sınıra sığmayan kopya görsel baytlarını yitirir (yalnız görselin zaten bulunduğu belgeye yapıştırılabilir). 5.000 paragraftan ya da 20.000 öğe/hücreden büyük seçim zengin kopyalanmaz (düz metin olarak gider) ve metin dışında bir şey içeriyorsa kesilemez. Uzun yapıştırma sınırı (5 milyon karakter) zengin kopya için de geçerlidir.
  - Yalnız görsel seçiliyken kopyalanan panoda düz metin boştur; dokunmatik balonundaki Yapıştır düğmesi panonun HTML'ini okuyabildiği tarayıcılarda (`navigator.clipboard.read`) zengin kopyayı tanır, okuyamadığında düz metni yapıştırır ya da klavyeyle yapıştırmayı ister.
  - Salt okunur görünümden kopyalanan seçim de zengin kopyadır. Paket olmayan bir dosyadan (RTF) alınan kopyanın bir docx'e görseli, alanı ve stiliyle yapıştırılıp kaydedilmesi birim testiyle sınanmıştır; ikili doc için ayrı bir test yoktur.
  - Programdan: `BymOffice.docOps.fragmentOf(doc, range)` kopyayı (`DocFragment`), `docOps.insertFragment(doc, konum | aralık, fragment, { taken? })` yapıştırma düzenlemesini (`DocEdit`) verir; `docOps.fragmentNeeds(doc, fragment)` hedef belgenin baytını beklediği görselleri sayar. `insertFragment`'e `moved: true` vermek, kopyanın kesilerek alındığını söyler (yukarıdaki "kesilenin geri konması" kuralı). `sliceOf()` dönüşündeki `DocSlice` artık `fragment` alanını da taşır.
- Çok sayıda bloğu değiştiren işlemde (ör. bütün belgeyi kalın yapmak) sayfalar açılıştaki gibi dilimler halinde yeniden dizilir; arayüz donmaz.
- Tuşlar: Ctrl/Cmd ile `+`, `-`, `0` ve tekerlek yakınlaştırır (dokunmatikte iki parmak); PageUp/PageDown/Boşluk/Home/End/oklar gezdirir; fareyle sürükleyerek ya da Ctrl/Cmd+A ile seçilen metin Ctrl/Cmd+C ile kopyalanır.
- Sayfalar aşamalı dizilir: ilk sayfalar hemen görünür, kalanı arka planda; binlerce sayfalık belgede yalnız ekrandaki sayfaların tuvali tutulur.
- **Fontlar ve sayfalama:** metin cihazdaki fontlarla ölçülür. Belgenin fontu cihazda yoksa ölçüleri uyumlu açık fontlar denenir (Calibri→Carlito, Cambria→Caladea, Arial→Liberation Sans, Times New Roman→Liberation Serif, Courier New→Liberation Mono), sonra genel aile. Bu fontları sayfanıza `@font-face` ile eklerseniz kendiliğinden kullanılır. Font farklıysa satır ve sayfa kırılmaları Word'dekinden farklı olabilir; Word'le yan yana karşılaştırma yapılmamıştır.
- Alt düzey: `BymOffice.createCanvasMeasurer({ fonts: doc.fonts })` → `new BymOffice.DocLayout(doc, measurer)` → `layout.layout({ pages: 5 })` → `BymOffice.renderPage(ctx, layout.pages[0], { scale })`. Sayfanın öğeleri (`page.items`) `text`, `rect`, `line`, `image`, `shape` (şekil: `geometry`, `fill`, `stroke`, `rotation`) ve `mark` türündedir. Model düzenlendikten sonra `layout.edited({ milliseconds })` yalnız değişen sayfadan itibaren yeniden dizer; `undefined` dönerse ya da sonraki `layout()` çağrılarından sonra `layout.broken` `true` olursa yeni bir `DocLayout` kurulur.

### Belgeyi yazdırma

```js
const mounted = await BymOffice.mount(container, { file: docxBytes });
if (mounted.kind === 'document') {
  const { pages, dpi } = await mounted.print();   // ya da mounted.view.print({ dpi: 300 })
}
```

- **Nasıl başlatılır:** belge çubuğundaki (yakınlaştırma düğmelerinin yanındaki) **Yazdır** düğmesi, görünüm ya da araç çubuğu odaktayken **Ctrl/Cmd+P** (tarayıcının sayfayı olduğu gibi yazdırması engellenir; yoksa kaydırılan alanın yalnız görünen kısmı basılırdı) ya da `mounted.print(options?)` / `view.print(options?)` / `editor.print(options?)`. Salt okunur görünümde ve doc/RTF dosyalarında da çalışır; `toolbar: false` ile çubuk gizliyse API yeter. Tablolar için bir sonraki başlığa bakın.
- **Ne yapar:** önce henüz dizilmemiş sayfaların dizilmesini bekler, sonra her sayfayı yazdırma çözünürlüğünde bir tuvale çizip PNG'ye çevirir (sayfa başına bir dilim; arayüz donmaz, belge o sırada okunmaya devam edilebilir) ve bunları belge düzeyinde (`<body>`'nin doğrudan çocuğu, sınıfı `bym-print`), ekranda görünmeyen bir kaba `<img>` olarak koyar. Yazdırma ortamı kuralları (`@media print`) barındıran sayfanın geri kalanını gizler ve `@page { size: …; margin: 0 }` ile kâğıdı belgenin sayfa boyutuna ayarlar; sonra görünümün içinde bulunduğu pencerenin `print()`'i çağrılır. Pencere kapanınca kap, stil ve görüntüler silinir. Hazırlık sırasında belge düzenlenirse sayfalar baştan hazırlanır (en çok 3 kez, sonra `failed`).
- **Dönüş:** yazdırma penceresi kapanınca `{ pages, dpi }` ile çözülen `Promise`. Kullanıcının yazdırdığı mı vazgeçtiği mi **ayırt edilemez** (tarayıcılar bunu bildirmez). Hata durumunda `DocumentPrintError` ile reddedilir; `code`: `'unavailable'` (pencerede `print()` yok), `'too-long'` (sınırların ötesinde belge), `'empty'` (sayfa yok), `'cancelled'` (hazırlık sürerken görünüm `destroy()` edildi), `'failed'` (diğer; düzeni sonuna kadar yapılamayan belge de böyle reddedilir, yarım belge basılmaz). Geçersiz `dpi` her durumda `RangeError` verir. Kullanıcıya gösterilecek ileti ayrıca `onNotice` ile bildirilir (başlangıçta "Yazdırmaya hazırlanıyor…", pencere kapanınca "Yazdırma penceresi kapandı", hata olursa nedeni); hazırlık sürerken çubuktaki sayfa göstergesi ilerlemeyi yazar ve düğme `aria-busy="true"` taşır.
- **Aynı anda ikinci çağrı:** sayfalar hazırlanırken yapılan çağrı ikinci bir pencere açmaz, **ilk çağrının `Promise`'ini** alır (kendi seçenekleri dikkate alınmaz). `view.printing` bir iş sürerken `true`'dur.
- **`destroy()`:** hazırlık sürerken çağrılırsa iş `cancelled` ile reddedilir; pencere açıkken çağrılırsa iş pencere kapanmış gibi biter. İki durumda da kap, stil ve görüntü adresleri hemen temizlenir.
- **Çözünürlük:** `dpi` seçeneği, varsayılan **200**, 72–600 arasına çekilir. A4 sayfa 200 dpi'de 1654 × 2339 pikseldir.
- **Bellek sınırları** (dokunmatik cihazlarda ikinci değer): bir sayfa en çok 16 / 8 milyon piksel (daha büyüğü yalnız o sayfa için düşük çözünürlükte çizilir); bütün sayfalar toplam 600 / 200 milyon piksel — aşılırsa çözünürlük bütün belge için düşürülür (masaüstünde 155 A4 sayfaya kadar 200 dpi; 500 sayfa ≈ 111 dpi), 72 dpi'nin altına inmesi gerekiyorsa (masaüstünde ≈ 1.200, dokunmatikte ≈ 399 A4 sayfa) `too-long` ile reddedilir; en çok 2.000 sayfa; PNG'lerin toplamı en çok 768 / 192 MB. Sayfalar tek bir tuvalde sırayla çizilir; aynı anda yalnız bir sayfa piksel olarak bellekte durur. Görseller ekrandakinden ayrı, aynı bütçeyle çalışan bir depoda yazdırma boyutunda çözülür; 30 saniyede çözülemeyen ya da bütçeye sığmayan görselin yerinde yer tutucu basılır.
- **Sınırlar:** sayfalar görüntü olarak basıldığı için **PDF'te metin seçilemez ve aranamaz**, dosya boyutu metin PDF'inden büyüktür ve bağlantılar tıklanamaz. Seçim vurgusu ve imleç basılmaz. Çıktı, ekrandaki sayfalarla aynı düzen motorundan gelir (aynı font ve sayfalama sınırlamaları geçerlidir).
- **Farklı boyutta/yönde bölümler:** her farklı kâğıt boyutu için adlandırılmış bir `@page` kuralı yazılır ve sayfa kendi kâğıdına basılır (dikey ve yatay bölümler aynı çıktıda). Bu, CSS `page` özelliğini bilen tarayıcılarda çalışır (Chromium tabanlılar ve güncel Firefox; Chromium'da üretilen PDF'in sayfa boyutlarıyla sınanır). `page` özelliğini bilmeyen tarayıcıda bütün yapraklar **ilk sayfanın boyutundadır**; farklı boyuttaki sayfa o yaprağa bütün olarak, ortalanıp gerekirse küçültülerek yerleştirilir (döndürülmez).
- **Tarayıcı farkları:** kâğıt boyutunu `@page size` belirler; Chromium tabanlı tarayıcılar buna uyar (yazdırma penceresinde ve "PDF olarak kaydet"te). **Safari ve Firefox `@page size`'ı her durumda uygulamayabilir**: o zaman kâğıt, yazdırma penceresinde seçilendir; sayfa görüntüleri gerçek ölçüsünde (inç olarak) yerleştirildiğinden aynı boyutta kâğıt seçildiğinde çıktı birebirdir, daha küçük kâğıtta kenarlar kesilir. Pencerede "kenar boşluğu: yok / varsayılan" ve "ölçek: %100" seçili olmalıdır. Gerçek yazıcıda, Safari'de ve iOS/Android'de elle deneme **yapılmadı**; otomatik testler yazdırma penceresini açmaz (`window.print` yerine sahte işlev, yazdırma ortamı öykünmesi ve Chromium'un PDF çıktısı kullanılır).
- **Pencerenin kapandığı nasıl anlaşılır:** `afterprint` olayıyla. `print()` dönmüş ama olay gelmemişse (ör. `window.print`'i uygulamanın değiştirdiği bir WebView) sayfadaki ilk tuş ya da işaretçi basışı (pencere açıkken sayfaya ulaşamaz), yeni bir `print()` çağrısı ya da 5 dakikalık üst sınır işi bitirir.
- **iframe ve Shadow DOM:** görünüm bir `<iframe>` içindeyse o çerçevenin penceresi yazdırılır (çevresindeki sayfa basılmaz; `sandbox` kullanılıyorsa `allow-modals` gerekir). Shadow DOM içinde kurulu görünümde de çalışır: yazdırma kabı gölge kökün dışına, belgenin `<body>`'sine eklenir. Yazdırma sürerken barındıran sayfanın `<html>` ve `<body>` öğelerine yalnız yazdırma ortamında geçerli kurallar uygulanır (yükseklik, taşma, kenar boşluğu, arka plan sıfırlanır); ekrandaki görünüm değişmez.
- **Güvenlik:** dosyadan gelen hiçbir metin yazdırma kabına girmez; kapta yalnız sayfa görüntüleri ve görünümün kendi "Sayfa N" etiketleri (`alt`) vardır. Görüntüler `blob:` adreslidir (`toBlob` olmayan eski tarayıcıda `data:`); sıkı içerik güvenlik ilkesinde `img-src blob: data:` ve `style-src 'unsafe-inline'` yeterlidir (köprü sayfasının ilkesi böyledir).
- Alt düzey: `BymOffice.planPrint(sayfaBoyutları, { dpi })` (çözünürlük ve piksel boyutları), `BymOffice.printCss(plan, adlandırılmışSayfa)` (stil), `BymOffice.DocumentPrintError`.

### Tabloyu yazdırma

```js
const mounted = await BymOffice.mount(container, { file: xlsxBytes });
if (mounted.kind === 'spreadsheet') {
  await mounted.print();                                        // etkin sayfa
  await mounted.print({ sheets: 'all', dpi: 300, fileName: 'butce.xlsx' });
}
```

Belge yazdırmayla **aynı altyapıdır**: sayfalar görüntü olarak hazırlanır, yazdırma ortamı kurallarıyla sayfanın geri kalanı gizlenir, pencerenin `print()`'i çağrılır. Dönüş değeri (`{ pages, dpi }`), hata sınıfı ve kodları (`DocumentPrintError`: `unavailable`, `too-long`, `empty`, `cancelled`, `failed`; geçersiz seçenekte `RangeError`), çözünürlük ve bellek sınırları, aynı anda ikinci çağrı, `destroy()`, pencerenin kapandığının anlaşılması, iframe/Shadow DOM, tarayıcı farkları ve güvenlik, "Belgeyi yazdırma" başlığında anlatılanla aynıdır. Buradakiler tabloya özgü olanlardır.

- **Nasıl başlatılır:** editör araç çubuğundaki **Yazdır** düğmesi (Geri al/Yinele'nin yanında; salt okunur kipte etkin kalan tek düğmedir; sayfalar hazırlanırken `aria-busy="true"` taşır ve durum çubuğu "Yazdırmaya hazırlanıyor… 3 / 12" yazar), ızgara, formül çubuğu ya da araç çubuğu odaktayken **Ctrl/Cmd+P**, ya da `mounted.print(options?)` / `view.print(options?)` / `editor.print(options?)`. xls dosyalarında ve salt okunur görünümde de çalışır. Düzenlenmekte olan hücre önce yazılır.
- **Seçenekler:** `dpi` (belgedeki gibi), `sheets`: `'active'` (varsayılan; gösterilen sayfa) ya da `'all'` (görünür bütün çalışma sayfaları, sekme sırasıyla; gizli sayfalar, grafik sayfaları ve boş sayfalar atlanır), `fileName` (üstbilgi/altbilgideki `&F` için; verilmezse boş). `WorkbookViewOptions.onPrintProgress({ done, total } | undefined)` hazırlık ilerlemesini bildirir (hazırlık bitince `undefined`). `view.printing` iş sürerken `true`'dur.
- **Ne yazdırılır:** sayfanın **yazdırma alanı** (`_xlnm.Print_Area` tanımlı adı; birden çok alan ayrı sayfalara, sırayla); yoksa **kullanılan alan**: A1'den, değeri ya da formülü olan, dolgusu ya da kenarlığı olan son hücreye, son birleşik hücreye ve bir görselin/şeklin altındaki son hücreye kadar. Yalnız satır ya da sütun düzeyinde verilmiş biçim ve koşullu biçim dolgusu kullanılan alanı büyütmez. Hiçbir şey yoksa `empty` ("Yazdırılacak bir şey yok"). **Gizli satır ve sütunlar** yer kaplamaz (süzgecin gizledikleri de).
- **Kâğıt:** `pageSetup.page.paperSize` kodu (1 Letter, 5 Legal, 8 A3, 9 A4, 11 A5 ve yaygın diğerleri: 3, 4, 6, 7, 12–20, 27, 28, 34, 66, 70) ya da `paperWidth`/`paperHeight` (mm, cm, in, pt); `orientation: 'landscape'` kâğıdı çevirir. **Dosya kâğıt belirtmiyorsa ya da kod tanınmıyorsa:** görünümün dili ABD/Kanada ise (`en`, `en-US`, `fr-CA`…) **Letter**, diğer bütün dillerde **A4** (Excel varsayılan yazıcının kâğıdını kullanır; tarayıcı bunu bildirmez). **Kenar boşlukları** dosyadan; yoksa Excel'in "Normal" değerleri (sol/sağ 0,7 inç, üst/alt 0,75 inç, üstbilgi/altbilgi 0,3 inç). Hücrelere yer bırakmayan kenar boşlukları varsayılanla değiştirilir.
- **Sayfalara bölme:** sayfa, bütün satır ve sütunlarla doldurulur; **elle sayfa sonları** (`rowBreaks`, `colBreaks`; sonradan eklenen/silinen satır-sütunlarla birlikte kayar) yeni sayfa başlatır. **Kâğıttan geniş bir sütun ya da yüksek bir satır ortadan bölünmez** (Excel böler): kendi sayfasına konur ve taşan kısmı kenar boşluğunda kesilir. Sayfa sırası `pageOrder` (`downThenOver` varsayılan, `overThenDown`).
- **Ölçek:** `scale` (%10–400; aralık dışıysa %100) ya da `fitToPage` ile sığdırma: `fitToWidth` × `fitToHeight` sayfaya (0 = sınırsız, belirtilmemişse 1) sığan **en büyük tam yüzde** bulunur; hiçbir zaman büyütülmez (%100 üst sınır), %10'un altına inilmez (o durumda %10'da gereken kadar sayfa basılır). Sığdırma açıkken elle sayfa sonları dikkate alınmaz.
- **Yinelenen başlıklar** (`_xlnm.Print_Titles`): başlık satırları, onlardan sonra başlayan her sayfanın üstünde; başlık sütunları solunda yinelenir ve sayfadan yer alır. Hücrelere yer bırakmayacak kadar büyük başlıklar yinelenmez.
- **Yazdırma seçenekleri:** `gridLines` (`gridLinesSet` kapalı değilse) kılavuz çizgilerini ince siyah çizgiyle basar — ekrandaki kılavuz çizgisi ayarından bağımsızdır, varsayılan basılmaz; `headings` satır numaralarını ve sütun harflerini basar; `horizontalCentered` / `verticalCentered` hücre bloğunu kenar boşlukları arasında ortalar.
- **Üstbilgi/altbilgi:** `&L`, `&C`, `&R` bölümleri; `&P` sayfa no (`&P+n`, `&P-n`), `&N` sayfa sayısı, `&D` tarih ve `&T` saat (görünümün yerel ayarıyla), `&A` sayfa adı, `&F` dosya adı, `&&`; biçim kodları `&"Font,Stil"`, `&12` (boyut), `&B`, `&I`, `&U`, `&E` (çift altı çizili), `&S` (üstü çizili), `&X`/`&Y` (üst/alt simge), `&KRRGGBB` (renk). Tanınmayan ya da uygulanmayan kodlar (`&G`, `&O`, `&H`, tema rengi, `&Z`) **metne sızmaz**, atlanır. Satır sonları korunur. `differentFirst` (sayfanın ilk yaprağı) ve `differentOddEven` (çift numaralı yapraklar) uygulanır; `scaleWithDoc` (varsayılan açık) yazıları sayfanın ölçeğiyle küçültür; `alignWithMargins` kapalıysa bölümler kenar boşluğu yerine kâğıt kenarından 0,7 inç içeride başlar. Üstbilgi kâğıdın üstünden `margins.header`, altbilgi altından `margins.footer` uzaklıktadır; yazı uzunsa hücrelerin üstüne taşabilir (Excel'de de öyle). **Sayfa numaraları** bütün iş boyunca sürer (`sheets: 'all'`'da sayfadan sayfaya devam eder); `useFirstPageNumber` + `firstPageNumber` veren sayfa oradan başlar; `&N` işteki toplam yaprak sayısıdır.
- **Çizilenler:** hücreler ekrandaki çiziciyle çizilir (değerler ve sayı biçimleri, yazı tipleri, dolgular, kenarlıklar, hizalama, metin kaydırma, **birleşik hücreler**, **koşullu biçimler**), üstlerinde **görseller ve şekiller**. **Basılmayanlar** (ekrana aittir): seçim ve etkin hücre çerçevesi, süzgeç ve liste düğmeleri, yorum işaretleri ve yorum kutuları (Excel'in varsayılanı da yorumları basmamaktır), formül uyarı işaretleri, dondurulmuş bölme çizgileri.
- **Hazırlık sırasında değişiklik:** çalışma kitabı değişirse (düzenleme, geri al/yinele, yeniden hesap) sayfalar baştan dizilip hazırlanır (en çok 3 kez, sonra `failed`).
- **Sınırlar:** belgedekiyle aynı (en çok 2.000 sayfa; toplam piksel ve bayt sınırları). Sınırın ötesindeki sayfa — ör. kullanılan alanı 1.048.576. satıra uzanan bir sayfa — daha dizilirken, hiçbir sayfa çizilmeden `too-long` ile reddedilir ("Sayfa bir kerede yazdırılamayacak kadar uzun").
- **Excel'le doğrulanamayanlar** (elde Excel yok; kodda `UNSURE` ile işaretli): sığdırma açıkken elle sayfa sonlarının yok sayılması; silinen satırların içinde kalan sayfa sonunun silme yerine taşınması; `sheets: 'all'`'da sayfa numaralarının sayfalar arasında sürmesi; `alignWithMargins` kapalıyken 0,7 inçlik iç boşluk.
- Alt düzey (DOM gerektirmez, Node'da da çalışır): `BymOffice.paginateSheet(sheet, { areas, titleRows, titleCols, styles, drawings, defaultPaper, maxPages })` → `{ paper, margins, scale, headings, gridLines, pages: [{ rows, cols, titleRows?, titleCols?, x, y, width, height }] }` (px, inç başına 96; boş sayfada ya da sınır aşımında `SheetPagesError`), `BymOffice.printRanges(workbook, sayfaDizini)` (yazdırma alanı ve başlıklar), `BymOffice.usedRange(sheet, …)`, `BymOffice.paperSize(page, varsayılan)`, `BymOffice.planSheetPages(workbook, sayfaDizinleri, …)` (bir işin bütün yaprakları, numaralarıyla), `BymOffice.parseHeaderFooter(kod)` + `BymOffice.headerFooterLines(parçalar, değerler)`, `BymOffice.paintSheetPage(ctx, …)` (bir yaprağı tuvale çizer).

### Formüller

```js
const engine = new BymOffice.CalcEngine(workbook);
engine.recalculate();                                  // tüm formüller
engine.recalculate([{ sheet: 0, row: 1, col: 1 }]);    // yalnız B2'ye bağlı olanlar
engine.evaluateText('SUM(A1:A3)*2', { sheet: 0, row: 0, col: 0 });   // saklamadan hesapla
```

Editör bunu kendisi yapar: her düzenlemeden, geri al ve yineleden sonra değişen hücrelere bağlı formüller yeniden hesaplanır. Kendi kodunuzda modeli doğrudan değiştirirseniz `recalculate` çağırmak size kalır; satır/sütun eklenip silindiyse önce `engine.invalidate()`.

- **Fonksiyonlar** (`BymOffice.FUNCTIONS` kayıt haritası): matematik (SUM, SUMIF(S), SUMPRODUCT, ROUND ailesi, MOD, …), istatistik (AVERAGE, COUNT ailesi, MIN/MAX, MEDIAN, STDEV, PERCENTILE, RANK, …), mantıksal (IF, IFS, IFERROR, AND, OR, SWITCH, CHOOSE), metin (CONCAT, TEXTJOIN, LEFT/MID/RIGHT, FIND/SEARCH, SUBSTITUTE, TEXT, VALUE, …), tarih (DATE, TODAY, EDATE, EOMONTH, DATEDIF, NETWORKDAYS, WORKDAY, WEEKNUM, …), arama (VLOOKUP, HLOOKUP, XLOOKUP, MATCH, INDEX, OFFSET, INDIRECT, …), bilgi (ISNUMBER, ISERROR, …), finans (PMT, FV, PV, RATE, NPV, IRR, …).
- Formüller dosyada ve modelde Excel'in İngilizce fonksiyon adları, `,` argüman ayırıcısı ve `.` ondalık ayırıcısıyla saklanır. **Türkçe yerel ayarda** (ve ondalık ayırıcısı virgül olan diğer yerellerde) kullanıcı Türkçe Excel'deki gibi yazar: `=EĞER(A1>5;1,5;0)`; argümanlar `;` ile, ondalıklar `,` ile. Yaklaşık 85 Türkçe fonksiyon adı (TOPLA, EĞER, DÜŞEYARA, ORTALAMA…) girişte tanınır; formül çubuğu adı İngilizce, ayırıcıları yerel gösterir (`=IF(A1>5;1,5;0)`). Türkçe yerelde virgül yalnız ondalık ayırıcıdır: ondalık olmayan virgül içeren formül (`=IF(A1>5,1,0)`, `=ROUND(A1,2)`) Türkçe Excel'deki gibi REDDEDİLİR ve `;` kullanılması istenir. Türkçe hata adları (`#YOK`, `#DEĞER!`…) girişte tanınır. İngilizce yerelde davranış değişmez. Çeviri: `BymOffice.formulaFromLocal(text, locale)` / `formulaToLocal`.
- Okunamayan formül (`=1+`) yazılamaz: editör açık kalır ve nedenini söyler (`onNotice`); `setCellInput` `RangeError` fırlatır. Dışarıdan yapıştırılan okunamayan formül metin olarak kalır.
- `SUBTOTAL` ve `AGGREGATE` gizli satırları Excel gibi ayırır: süzgecin gizlediği satırlar her zaman, elle gizlenenler 101–111 kodlarıyla atlanır; süzgeç değişince sonuç güncellenir. Kendi kodunuzda satır gizliliğini değiştirirseniz `engine.rowsChanged(sheet)` ardından `engine.recalculate([])` çağırın; tanımlı adları değiştirirseniz `engine.namesChanged()`. `engine.evaluateRange(text, at)` aralık döndüren bir formülün tüm değerlerini tek çağrıda verir.
- **Formül işaretleri:** hücrenin sol üst köşesindeki üçgen, formülün burada hesaplanamadığını söyler: turuncu = dosyadaki son değer gösteriliyor (desteklenmeyen fonksiyon vb.), kırmızı = döngüsel başvuru, gri = hesap sürüyor ya da değer eski. Hücre seçilince durum çubuğu nedenini yazar.
- Dosyadaki değerlere güvenilir: açılışta tüm kitap yeniden hesaplanmaz, yalnız değişikliklerden etkilenen formüller hesaplanır. Sonucu hiç olmayan formüller (hesaplama yapmayan kitaplıkların yazdığı dosyalar) açılışta bir kez hesaplanır (`engine.uncalculated()`).
- Girdisi değişmediği halde bizim sonucumuz dosyadakinden farklı çıkan formülde dosyadaki değer kalır (`engine.mismatched`): Excel'in bildiği bir şeyi bilmiyoruz demektir (ör. Türkçe biçim kodlu `TEXT`). Satır/sütun eklenip silindikten sonra ise tüm sonuçlar yeniden hesaplanıp yazılır (`recalculate(undefined, { overwrite: true })`).
- Süre bütçesini tek başına aşan formül sert sınıra (`budget.hardMilliseconds`) kadar denenir; onu da aşarsa kendisi ve ona bağlı olanlar `engine.stale` listesine girer. Koşullu biçim gibi çok sayıda küçük hesap için `engine.batch(() => …)` ve `engine.evaluateShifted(text, anchor, at)` vardır.
- Döngüsel başvurudaki hücreler 0 olur (`engine.circular` bunları listeler).
- **Hesaplayamadığımız formül dosyadaki değerini korur.** Ayrıştırılamayan formül (`engine.unparsed`; ör. tablo başvurusu) ve desteklenmeyen bir şey kullanan formül (`engine.unsupported`; kayıtlı olmayan fonksiyon, `LET`/`LAMBDA`, 3B ya da başka dosyaya başvuru, tanımsız ad) hesaplanmaz; formül metni olduğu gibi kaydedilir. Yeni yazılan böyle bir formül ise `#NAME?`/`#REF!` gösterir.
- **Örtük kesişim:** normal bir formülde çok hücreli başvuru tek değer beklenen yerde formülün satırındaki (ya da sütunundaki) hücreyi verir (`=A1:A3*2`, 2. satırda A2*2); dizi formüllerinde ve `SUMPRODUCT` gibi dizi bekleyen fonksiyonlarda dizi olarak hesaplanır. Dinamik dizi taşması (spill) yoktur.
- **Süre bütçesi:** `new CalcEngine(workbook, { budget: { milliseconds } })` (varsayılan 5000). Süre dolunca kalan formüller eski değerinde kalır ve `engine.incomplete` doğru olur; `engine.recalculate([])` kaldığı yerden sürdürür. Editör 250 ms'lik dilimlerle hesaplar ve kalanını kendisi sürdürür; sayfa donmaz. Tek başına bütçeyi aşan formüller `engine.overBudget`'ta listelenir ve değerini korur.
- **Doğrulama sınırı:** sonuçlar Excel belgelerindeki örneklerle ve bağımsız hesaplarla sınandı; gerçek Excel ile doğrudan karşılaştırma yapılmadı. Nadir kenar durumlarında Excel'den sapma olabilir. Kaydedilen dosyayı Excel açınca kendi sonuçlarını yazar.
- Başarım: 100.000 formüllü bir kitapta tek hücre değişikliği birkaç milisaniyede yeniden hesaplanır; 100.000 satırlık tabloya bakan binlerce tam eşleşmeli `VLOOKUP` onlarca milisaniyede.

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
| `=A1*2` | Formül; hemen hesaplanır ve sonucu görünür |
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
- **Biçimlendirme** tam satır ya da sütun seçiliyken satırın/sütunun kendi biçimini değiştirir (hücre hücre değil); tek seferde en çok 200.000 hücre biçimlendirilebilir.
- **Kesme** kopyalayıp temizler; Excel'deki gibi "taşıma" değildir (kesilen hücrelere bakan formüller yeni yeri izlemez).
- Kendi kopyaladığınız aralığı yapıştırdığınızda biçimler ve formüller birlikte gelir, formül başvuruları yeni yere göre kayar. Başka bir programdan gelen metin hücre hücre yazılmış gibi yorumlanır.

Aynı işlemler arayüz olmadan da kullanılabilir (ör. Node'da): `setCellInput(workbook, sheetIndex, row, col, text, { locale })`, `clearRange`, `copyRange`, `pasteAt` işlem listesi (`Op[]`) üretir; `new History(workbook).apply(ops, before, after)` uygular, `undo()`/`redo()` geri alır. Satır eklemek için doğrudan `{ type: 'insert', sheet, axis: 'row', at, count }` işlemi verilir. Yardımcılar: `parseCellInput(text, options)`, `cellEditText(cell, numFmt, options)`, `shiftFormula(text, dRow, dCol)`, `adjustFormula(text, sheetName, edit)`.

### `open(input): Promise<OfficeFile>`

`input`: `Uint8Array`, `ArrayBuffer` ya da `Blob`/`File`.

```ts
interface OfficeFile {
  format: 'docx' | 'xlsx' | 'doc' | 'xls' | 'rtf';
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb';
  package?: OpcPackage; // docx/xlsx
  cfb?: CfbFile;        // doc/xls
  workbook(): Promise<Workbook>; // tablo modeli: xlsx ve xls
  saveAsXlsx?(): Promise<Uint8Array>; // xlsx'te save() ile aynı; xls'te modelden yeni bir xlsx
  save(): Promise<Uint8Array>;
}
```

Hatalar:

- `UnsupportedFileError`: girdi Word/Excel dosyası değil ya da henüz desteklenmeyen bir tür (parola korumalı, Word 95 ve öncesi).
- `Error`: dosya bozuk (iletisi `Corrupt ZIP archive: …`, `Corrupt ZIP entry …` vb. ile başlar).

### Sayfa ekleme, silme, yeniden adlandırma

Tablo görünümünün sekme şeridinden ya da programla yapılır; her biri **tek geri alma adımıdır** ve `save()` ile dosyaya yazılır (ek bir çağrı gerekmez).

```js
const view = mounted.view;                 // WorkbookView
const index = view.insertSheet();          // sona ekler, yeni sayfayı gösterir; dizinini döndürür ("Sayfa1" / "Sheet1")
view.insertSheet(0, 'Özet');               // en başa, adıyla
view.renameSheet(index, 'Satışlar 2026');  // true; ad zaten buysa false
view.deleteSheet(0);                       // true; onay sormaz
view.undo();                               // silinen sayfa seçimi ve kaydırma konumuyla geri gelir
view.startRename(view.activeSheet);        // sekmede ad alanını açar (kullanıcı yazar)
```

| Üye | İş |
|---|---|
| `insertSheet(at?, name?)` | `at` dizinine (verilmezse sona) boş bir çalışma sayfası ekler ve **onu gösterir**. Ad verilmezse görünümün dilindeki ilk boş varsayılan ad kullanılır (`tr` → `Sayfa1`, diğerleri `Sheet1`). Yeni dizini döndürür; salt okunur görünümde `-1` |
| `renameSheet(index, name)` | Sayfanın adını değiştirir. Değişiklik yapıldıysa `true`; ad aynıysa ya da görünüm salt okunursa `false` |
| `deleteSheet(index)` | Sayfayı siler. Gösterilen sayfa silindiyse sağındaki ilk görünür sayfa (sondaysa solundaki) gösterilir. Salt okunurda `false` |
| `startRename(index)`, `isRenaming` | Sekmenin yerinde ad alanını açar (sayfa önce gösterilir); alan açık mı. Gizli sayfada (sekmesi yoktur; gösterilen sayfa değişmez) ve salt okunurda `false` döner |

Davranış:

- **Arayüz:** `+` düğmesi (`.bym-tab-add`), sekme menüsü (`.bym-menu`) ve ad alanı (`.bym-tab-input`, en çok 31 karakter, mevcut ad seçili gelir). Ad alanında Enter ya da odağın alandan çıkması uygular, Esc vazgeçer; giriş yöntemi (IME) bileşimi sürerken Enter adı uygulamaz; alandaki tuşlar ızgaranın kısayollarına ulaşmaz (Ctrl/Cmd+Z alanın kendi geri almasıdır). Geçersiz ya da başka sayfada kullanılan bir ad Enter ile uygulanmak istenirse alan açık kalır, `aria-invalid` işaretlenir ve neden `onNotice` (`kind: 'error'`) ile ve canlı bölgede bildirilir. Odak alandan böyle bir adla çıkarsa sayfa eski adını korur ve neden yine bildirilir. Ad yazılırken başka bir sekmeye basmak adı uygular ve o sekmeyi gösterir; Tab ile çıkınca odak sıradaki sekmeye geçer. Alan, sekme listesinin (`role="tablist"`) içinde `role="none"` taşıyan bir sarmalayıcıda (`.bym-tab-edit`) durur. Menüde yapılamayan öğe `aria-disabled="true"` taşır ama ok tuşlarıyla ulaşılabilir; Enter ya da dokunma işlemi yapmaz, nedenini bildirir (menü açık kalır).
- **Geçerli ad:** 1–31 karakter; `[ ] : * ? / \` ve denetim karakteri içermez, kesme işaretiyle (`'`) başlamaz/bitmez, `History` değildir ve (büyük/küçük harf ayrımı olmadan) başka bir sayfanın adı değildir. Baştaki/sondaki boşluklar atılmaz.
- **Formüller:** yeniden adlandırmada sayfayı anan formüller, tanımlı adlar ve kitap içi hücre bağlantıları **yeni adla yazılır** (gerekirse tırnaklanarak: `'Satışlar 2026'!A3`); hiçbir değer değişmez ve hiçbir formül yeniden hesaplanmaz (`RAND()` gibi her hesapta değişenler de). **Ekleme** de hiçbir şeyi hesaplamaz: boş bir sayfa hiçbir değeri değiştirmez, dosyadaki sonuçlar olduğu gibi kalır; "ekle → geri al → kaydet" özgün dosyayı bayt bayt verir. Sayfa eklenmeden önce yazılmış `=Yeni!A1` gibi bir formül, sayfa eklendiği anda değil, okuduğu bir hücre ilk değiştiğinde yeniden hesaplanır. **Silmede** o sayfayı okuyan başvurular `#REF!` olur (Excel `#REF!A1` yazar; burada başvurunun tamamı `#REF!` olur), sayfaya yerel tanımlı adlar sayfayla birlikte gider; bir 3B başvurunun (`Sayfa1:Sayfa3!A1`) ucu silinirse uç aralığın içindeki komşu sayfaya çekilir, tek sayfa kalırsa düz başvuru olur. Koşullu biçim ve doğrulama formülleri ile tablo sütun formülleri de aynı kurallarla izler; yalnız formülü böyle değişen hücreler ve onları okuyan formüller yeniden hesaplanır, kitabın geri kalanı dosyadaki değerini korur. Şu durumlarda silme (ve geri alınması) bütün formülleri yeniden hesaplar ve dosyadaki sonuçların yerini hesaplananlar alır: silme tanımlı bir adı değiştirdiyse, o sırada süren bir hesap bitirilemediyse, 100.000'den çok hücre değiştiyse ya da bağımlılık grafiği kurulamadıysa. Geri alma hepsini eski haline getirir. Sayfa işleminden sonra formül işaretleri (hücre köşesindeki "burada hesaplanamıyor" işareti) bir sonraki hesaba dek görünmez.
- **Geri al/yinele:** silme geri alınınca aynı sayfa, seçimi ve kaydırma konumuyla geri gelir ve gösterilir (gizli bir sayfa geri gelince gösterilen sayfa değişmez); ekleme geri alınınca eklemeden önce gösterilen sayfaya dönülür. Diğer sayfaların seçimi ve kaydırması sayfa dizinleri kaysa da kendi sayfasında kalır.
- **Pano:** sayfa listesi her değiştiğinde (işlem, geri al, yinele) görünümün kendi kopyası unutulur; ardından yapılan yapıştırma, eski sayfa adını taşıyan formülleri değil panodaki düz metni (görünen değerleri) yazar.
- **Olaylar:** her işlemden (ve geri al/yinele'den) sonra `onChange({ sheets })` kitabın o andaki **bütün** sayfa dizinleriyle çağrılır; gösterilen sayfa ya da onun dizini değiştiyse `onSheetChange(index)` yeni dizinle çağrılır (yalnız ad değiştiyse çağrılmaz). Önceden elinizde tuttuğunuz dizinler ekleme/silmeden sonra kaymış olabilir; sayfa listesini `view.workbook.sheets`'ten yeniden okuyun.
- **Retler:** yapılamayan işlem programda `RangeError` fırlatır; arayüzde aynı neden görünümün dilinde bildirilir (`view.localizeError(error)`), menüde ilgili öğe devre dışıdır ve nedeni `title` olarak taşır. Reddedilenler: geçersiz ya da kullanılan ad; görünür son sayfayı silmek; içinde grafik, pivot tablo ya da form denetimi bulunan kitapta silme/yeniden adlandırma; adları yalnız harf büyüklüğüyle ayrılan iki sayfası olan kitapta her sayfa işlemi; tablo içeren sayfayı silmek; grafik sayfasını silmek/yeniden adlandırmak; sayfa adı içerebilecek bir formül sonuna dek okunamıyorsa (kapanmamış tırnak) silme/yeniden adlandırma. Reddedilen işlem modeli değiştirmez.
- **Kaydetme:** yeni sayfa pakete kendi parçasıyla (`/xl/worksheets/sheetN.xml`, ilk boş numara) eklenir; silinen sayfanın parçası ve yalnız onun kullandığı parçalar (ilişkileri, çizimi, yorumları, yazıcı ayarı, başka hiçbir yerde kullanılmayan görselleri) paketten çıkar; yeniden adlandırma `workbook.xml`'e ve değişen formüllerin sayfalarına yazılır. Dokunulmayan sayfaların parçaları bayt bayt aynı kalır. Her şey geri alınıp kaydedilirse dosya geldiği sayfalarla yazılır.
- **Salt okunur:** `readOnly: true` görünümde ve dönüştürülmemiş xls'te `+` düğmesi gizlidir, sekme menüsü açılmaz (tarayıcının kendi menüsü engellenmez), çift tıklama ve F2 ad alanını açmaz, yöntemler hiçbir şey yapmaz.
- **WebView köprüsü:** köprüde sekme şeridi aynı biçimde çalışır; sayfa işlemleri için ayrı bir köprü komutu yoktur. `changed` iletisinin `sheets` alanı işlemden sonraki dizinleri taşır; `loaded` iletisindeki sayfa listesi sonradan güncellenmez (yeniden adlandırma ayrıca bildirilmez).

Görünüm olmadan, model üzerinde (Node'da da) aynı işlemler `Op` üreten işlevlerle yapılır; işlemler bir `History` üzerinden uygulanır:

```js
const history = new BymOffice.History(workbook);
history.apply(BymOffice.insertSheet(workbook, workbook.sheets.length, { locale: 'tr' }), undefined, undefined);
history.apply(BymOffice.renameSheet(workbook, 0, 'Veri'), undefined, undefined);
if (BymOffice.canDeleteSheet(workbook, 1) === undefined) history.apply(BymOffice.deleteSheet(workbook, 1), undefined, undefined);
history.undo();
```

| İşlev | İş |
|---|---|
| `insertSheet(workbook, at, { name?, locale? })`, `renameSheet(workbook, index, name)`, `deleteSheet(workbook, index)` | İşlemi (`Op[]`) üretir, modeli değiştirmez; yapılamayacaksa `RangeError` fırlatır. `renameSheet` ad aynıysa boş dizi döndürür. Ekleme `workbook.activeSheet`'in gösterdiği sayfayı değiştirmez (dizin kayar); silme onu geçerli, görünür bir sayfaya ayarlar |
| `canInsertSheet(workbook, at, name?)`, `canRenameSheet(workbook, index, name?)`, `canDeleteSheet(workbook, index)` | Yapılamıyorsa İngilizce neden, yapılabiliyorsa `undefined`. Hücrelere bakmaz (ucuzdur); okunamayan formül reddi yalnız işlem uygulanırken çıkar |
| `defaultSheetName(workbook, locale?)` | İlk boş varsayılan ad (`Sayfa1` / `Sheet1`) |
| `checkSheetName(name)` | Adı döndürür; sayfa adı olamayacaksa `RangeError` fırlatır (çakışmaya bakmaz) |
| `quoteSheetName(name)` | Adı formülde yazıldığı gibi verir (gerekirse tırnaklı) |
| `renameSheetsInFormula(text, map)`, `formulaNamesSheetsUnread(text)` | Formül metnindeki sayfa adlarını `map(ad)` ile değiştirir (`string` yeni ad, `null` silindi → `#REF!`, `undefined` dokunma); formül sonuna dek okunamıyor ve okunmayan kısmı sayfa adı içerebiliyor mu |
| `fileSheetName(sheet)`, `currentSheetName(workbook, nameInFile)`, `sheetNameMap(workbook)` | Sayfanın dosyadaki adı (dosyada olmayan yeni sayfada `undefined`); dosyadaki bir adın bugünkü karşılığı (`null`: silindi); hepsi için eşleme işlevi (hiçbir ad değişmediyse `undefined`). Dosyadan olduğu gibi tutulan içerik (koşullu biçim, veri doğrulama kuralları) sayfaları dosyadaki adlarıyla anar |
| `Op` türleri `insertSheet`, `deleteSheet`, `renameSheet`, `restoreSheets`; `SheetsSnapshot` | İlk üçünün tersi `restoreSheets`'tir (kendi tersidir) ve sayfa listesinin önceki halini `SheetsSnapshot` olarak taşır. Hepsinde `sheet` etkilenen dizindir |

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
| `pageSetup` | Sayfa yapısı (yazdırma ayarları); dosya hiçbir şey belirtmiyorsa `undefined`. Aşağıda |

`sheet.pageSetup` (`SheetPageSetup`) SpreadsheetML öğeleriyle bire bir eşleşir; her üye dosyada ilgili öğe (ya da xls kaydı) varsa bulunur, üyelerin alanları da yalnız dosya belirtiyorsa bulunur (olmayan alan, öğenin varsayılanıdır):

| Üye | Öğe | İçerik |
|---|---|---|
| `margins` | `pageMargins` | `left`, `right`, `top`, `bottom`, `header`, `footer` (inç; altısı da hep vardır) |
| `page` | `pageSetup` | `paperSize` (kâğıt kodu: 1 Letter, 9 A4, …), `paperWidth`/`paperHeight`, `orientation` (`'default' \| 'portrait' \| 'landscape'`), `scale` (10–400), `fitToWidth`, `fitToHeight` (0: gerektiği kadar), `firstPageNumber`, `useFirstPageNumber`, `pageOrder` (`'downThenOver' \| 'overThenDown'`), `blackAndWhite`, `draft`, `cellComments` (`'none' \| 'asDisplayed' \| 'atEnd'`), `errors` (`'displayed' \| 'blank' \| 'dash' \| 'NA'`), `horizontalDpi`, `verticalDpi`, `copies`, `usePrinterDefaults` |
| `print` | `printOptions` | `horizontalCentered`, `verticalCentered`, `headings` (satır/sütun başlıkları), `gridLines`, `gridLinesSet` |
| `headerFooter` | `headerFooter` | `oddHeader`, `oddFooter`, `evenHeader`, `evenFooter`, `firstHeader`, `firstFooter` (Excel'in kodlu metni: `&L` `&C` `&R` bölümler, `&P` sayfa, `&N` sayfa sayısı, `&D` tarih, `&"font,stil"`, `&&` = `&`); `differentOddEven`, `differentFirst`, `scaleWithDoc`, `alignWithMargins` |
| `fitToPage`, `autoPageBreaks` | `sheetPr/pageSetUpPr` | Sayfaya sığdırma etkin mi (`page.fitToWidth` × `page.fitToHeight`); otomatik sayfa sonları gösterilsin mi |
| `rowBreaks`, `colBreaks` | `rowBreaks`, `colBreaks` | Sayfa sonları, artan sırada: `at` (yeni sayfanın ilk satırı/sütunu, 0 tabanlı; sonun üstünde/solunda kaldığı satır/sütun), `min`, `max` (kapsadığı sütun/satır aralığı), `manual`, `pivot` |

- Bu alan **okuma görünümüdür**: xlsx `save()` bu öğeleri modelden değil özgün baytlardan yazar (tanınmayan öznitelikler ve yazıcı ayarı bağlantısı dahil aynen kalır); `pageSetup`'ı kodla değiştirmek yalnız `buildXlsx()` çıktısını etkiler. Satır/sütun eklenip silinince modeldeki sayfa sonları değişmez; kaydırma dosya yazılırken uygulanır.
- xlsx'te değer taşımayan öznitelik (aralık dışı ölçek, tanınmayan ad) alınmaz; öğe iki kez yazılmışsa ilki geçerlidir; özel görünümlerin (`customSheetViews`) sayfa yapısı okunmaz.
- xls'te okunan kayıtlar: `HEADER`, `FOOTER`, `HEADERFOOTER` (çift/ilk sayfa), `LEFTMARGIN`…`BOTTOMMARGIN`, `SETUP`, `HCENTER`, `VCENTER`, `PRINTHEADERS`, `PRINTGRIDLINES`, `GRIDSET`, `WSBOOL` (sığdırma ve otomatik sayfa sonu bitleri), `HORIZONTALPAGEBREAKS`, `VERTICALPAGEBREAKS`; Excel 5/95'te dizgiler kod sayfasından çözülür, sayfa sonlarının aralığı yoktur (tüm sayfa). `SETUP` "yazıcı ayarları hiç belirlenmedi" diyorsa (`fNoPls`) kâğıt, ölçek, yön, çözünürlük ve kopya sayısı alınmaz; "yön belirlenmedi" diyorsa (`fNoOrient`) yön alınmaz. Kaydı olmayan kenar boşluğu Excel 97–2003 varsayılanıdır (0,75 / 1 / 0,5 inç). Varsayılanla aynı olan Boole değerleri ve 1×1 sığdırma modele yazılmaz. Eski biçimin tüm genişliğini/yüksekliğini kapsayan sayfa sonu modelde tüm sayfayı kapsar. Kısa ya da anlamsız kayıt sayfayı reddettirmez; o alan atlanır ve kayıt `readXlsWithReport().report.skippedRecords` içinde sayılır. Aynı kayıt iki kez varsa ilki geçerlidir. `GRIDSET` modele alınmaz: kılavuz çizgilerinin basılıp basılmadığını yalnız `PRINTGRIDLINES` belirler (xlsx'te çizgiler iki öznitelik de doğruyken basıldığı için `GRIDSET`'i taşımak çizgileri kapatabilirdi). `SETUP`'taki "hücre hataları nasıl yazdırılır" alanında yalnız "tire" ve "#N/A" alınır; "boş" değeri alınmaz, çünkü eldeki gerçek dosyaların tamamı bu değeri taşıyor ve bunun gerçek bir kullanıcı ayarı olduğu Excel'de doğrulanamadı (yanlışsa her dönüştürülen dosyada hata hücreleri boş basılırdı). Özel görünüm bloğundaki (`USERSVIEWBEGIN`…`USERSVIEWEND`) kayıtlar ve yazıcı sürücüsü verisi (`PLS`) okunmaz, sayılır.
- Yazdırma alanı ve yinelenen başlık satırları tanımlı addır (`_xlnm.Print_Area`, `_xlnm.Print_Titles`; `workbook.definedNames`).

Yardımcılar: `getCell(sheet, row, col)`, `columnName(col)`, `cellName(row, col)`, `parseCellRef('B7')`, `parseRange('A1:C3')`, `rangeName(range)`, `resolveCellStyle(styles, index)`, `resolveColor(color, styles)`, `formatCodeOf(styles, numFmtId)`.

Bilinmesi gerekenler:

- Tarihler Excel'de olduğu gibi **sayı** olarak gelir; tarih olduğunu `style.numFmt` biçim kodu belirtir (`isDateFormat(code)`). Ekranda görünen metin için `cellDisplayText(workbook, cell, locale)`, takvim parçaları için `serialToDate(serial, workbook.date1904)` kullanın.
- Dosya açıldığında formüller yeniden hesaplanmaz; `cell.value` dosyaya en son kaydedilen sonuçtur. Bir hücre düzenlenince ona bağlı formüller hesaplanır ve `cell.value` güncellenir.
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
- xlsx dosyalarında **tablo görünümü**: gerçek ızgara (kaydırma, seçim, klavye, sayfa sekmeleri, dondurulmuş bölmeler) ve **hücre düzenleme** (yazma, geri al/yinele, kopyala/yapıştır, satır/sütun ekle-sil; üstteki düğmeler ve "Salt okunur" seçimi); araç çubuğundaki **Yazdır** düğmesi (ya da Ctrl/Cmd+P) etkin sayfayı sayfa yapısına göre yazdırır; altında seçili aralık, etkin hücrenin görünen değeri, ham değeri ve formülü yazar. "Yeniden kaydet ve indir" düzenlemelerinizi dosyaya yazar; indirilen dosyayı Excel'de açıp denetleyebilirsiniz. "Sayı ve tarih biçimi dili" seçimiyle Türkçe ve İngilizce gösterim arasında geçilir.
- docx dosyalarında **belge görünümü**: sayfalar ve belge düzenleyici (yazma, seçim, biçimlendirme araç çubuğu, geri al/yinele, pano); doc ve RTF salt okunur gösterilir. Belge çubuğundaki **Yazdır** düğmesi (ya da Ctrl/Cmd+P) tarayıcının yazdırma penceresini açar.
- "Yeniden kaydet ve indir" ile kaydedilen dosyayı indirip Word/Excel'de açmak.
- Bir xls dosyası seçiliyken "xlsx olarak indir" ile `saveAsXlsx()` çıktısını indirmek.
- "Örnek xlsx üret ve aç" ile elinizde dosya olmadan denemek.
- "Boş docx oluştur ve aç" / "Boş xlsx oluştur ve aç" ile `createBlank()` çıktısını açmak.
- "Eski cihaz kipi" ile tarayıcının yerleşik sıkıştırmasını kapatıp motorun kendi kodunu sınamak.

Dosyalar yalnız tarayıcıda işlenir; hiçbir yere gönderilmez.

## 7. Geliştirme komutları

| Komut | İş |
|---|---|
| `pnpm verify` | **Her şeyi sıfırdan sınar:** tip denetimi, birim testleri, derleme, tüm tarayıcı testleri, gerçek dosya havuzu (iki sıkıştırma kipinde) |
| `pnpm typecheck` | Tip denetimi |
| `pnpm test` | Birim testleri |
| `pnpm build` | `dist/` çıktıları |
| `pnpm pack:check` | Paketi arşivleyip boş bir projeye kurar ve oradan kullanır (ES modülü, tipler, `<script>`); önce `pnpm build` gerekir |
| `pnpm e2e` | Derler ve Playwright testlerini 5 tarayıcı profilinde çalıştırır |
| `pnpm corpus <klasör> …` | Gerçek dosyalarda aç → kaydet → yeniden aç denetimi (`BYM_NO_NATIVE=1` ile motorun kendi sıkıştırma koduyla); her xls ayrıca xlsx'e çevrilip yeniden okunur ve iki model karşılaştırılır ("converted to xlsx and read back") |
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
