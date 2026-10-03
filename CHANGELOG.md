# Sürüm notları

Bu dosya [Keep a Changelog](https://keepachangelog.com/tr/1.1.0/) düzenini izler; sürümler [anlamsal sürümleme](https://semver.org/lang/tr/) ile numaralanır. 1.0'dan önce (0.x) küçük sürümler kırıcı değişiklik içerebilir; bunlar "Kırıcı" başlığı altında yazılır.

## Yayımlanmadı

İlk sürüm için hazırlanan içerik. Paket henüz npm'de yayımlanmadı.

### Eklendi

- **Dosya biçimleri:** docx, xlsx, doc (Word 97–2003), xls (Excel 97–2003) ve RTF içerikten tanınır ve açılır (`open()`).
- **Kayıpsız kaydetme:** docx ve xlsx'te yalnız değişen parça yeniden üretilir; dokunulmayan parçalar bayt bayt aynı kalır.
- **Tablo düzenleyici** (`mount()` ile xlsx): Canvas ızgara, hücre düzenleme, geri al/yinele, pano, satır/sütun ekle-sil, biçimlendirme araç çubuğu, formül çubuğu, durum çubuğu, süzgeç ve sıralama, veri doğrulama, koşullu biçim gösterimi, görsel/şekil/yorum gösterimi.
- **Formüller:** ayrıştırıcı, bağımlılık grafiği, artımlı yeniden hesaplama, 177 fonksiyon; Türkçe yerel ayarda `;` ayırıcı ve Türkçe fonksiyon adlarıyla giriş.
- **Sayı biçimleri:** Excel biçim kodları, Türkçe ve İngilizce yerel ayar.
- **Belge görünümü ve düzenleyici** (`mount()` ile docx): sayfa düzeni, imleç ve seçim, yazma, giriş yöntemleri (IME), geri al/yinele, pano, biçimlendirme araç çubuğu, tablolar, bağlantılar, görseller; doc ve RTF salt okunur gösterilir.
- **Yeni dosya:** `createBlank('docx' | 'xlsx')` boş belge ya da tablo üretir.
- **Gömme örneği:** `examples/html/index.html` — dosyayı sitenin adresinden alıp düzenleten ve `PUT` ile geri yollayan, sekmede ya da `<iframe>` içinde açılan sayfa.
- **Dönüştürme:** doc ve RTF → docx (`saveAsDocx()`), xls → xlsx (`saveAsXlsx()`). `mount()`'ta `convert: true` bu dosyaları dönüştürerek düzenlenebilir açar.
- **WebView köprüsü:** docx köprüde düzenlenir (`changed`, `undo`, `redo`, `setReadOnly` belgede de çalışır; belgenin `changed` iletisinde `sheets` boş dizidir); `load.convert` doc/RTF/xls'i docx/xlsx'e çevirerek düzenlenebilir açar (`loaded`/`saved` iletilerinde `format` ve `convertedFrom`, `loaded.warnings`); `load.blank` boş docx/xlsx ile başlar; `load.toolbar` belgeyi biçimlendirme araç çubuğuyla gösterir. Dönüştürülmemiş doc ve RTF `loaded.readOnly: true` bildirir ve `setReadOnly(false)` ile açılamaz.
- **Sayfa yapısı (yazdırma ayarları):** `sheet.pageSetup` — kenar boşlukları, kâğıt, yön, ölçek/sayfaya sığdırma, ilk sayfa numarası, sayfa sırası, üstbilgi/altbilgi (tek/çift/ilk sayfa), ortalama, kılavuz çizgisi ve başlık yazdırma, sayfa sonları. xlsx ve xls'ten (Excel 97–2003 ve 5/95) okunur; xls → xlsx dönüştürmede (`saveAsXlsx()`, `buildXlsx()`, `convert: true`) yeni dosyaya yazılır. xlsx `save()` bu öğeleri eskisi gibi özgün baytlarıyla korur. Yazıcı sürücüsü verisi taşınmaz; yazdırma/PDF çıktısı yoktur.
- **Dönüştürme raporu:** `buildDocx(…, { report })` ve `conversion.warnings` taşınamayanları sayılarıyla bildirir (dikdörtgene çevrilen şekiller, gruplar, baytı olmayan görseller, düz metne çevrilen bağlantılar, tema, atılan karakterler). `convertForEditing()` ve `openConverted()` dışa açıldı.
- **doc ve RTF'te şekiller:** dikdörtgen, yuvarlatılmış dikdörtgen, elips ve düz çizgi; düz renk dolgu, çizgi rengi, kalınlığı ve kesikliği, çevirme ve döndürme okunur (`drawing.shape`), sayfada çizilir ve docx'e dönüştürmede DrawingML olarak (`a:prstGeom`, dolgu, `a:ln`) yazılır. Metin kutuları dolgu ve çerçeveleriyle görünür. docx'te de tek şekilli çizimlerde (grup değil) metinsiz dikdörtgen, elips, çizgi ve yuvarlatılmış dikdörtgen artık yer tutucu yerine kendi biçimiyle çizilir.
- **Gömme:** tek `<script>` (global `BymOffice`), ES modülü ve TypeScript tipleri; WebView'ler için tek dosyalık köprü sayfası (`bymoffice/webview.html`) ve mesaj sözleşmesi.
- **Eski cihazlar:** yerleşik sıkıştırma akışı olmayan ortamlar için kendi inflate/deflate kodu; derleme hedefi ES2019.

### Değişti

- doc/RTF → docx dönüştürmede `N shape(s) were written as plain rectangles…` uyarısı artık yalnız görünümü modelde bulunmayan şekiller için çıkar; geometrisi tanınmayan şekiller için yeni uyarı: `N shape(s) were written as rectangles: their geometry was not carried over`.
- doc/RTF → docx dönüştürmede yeni uyarılar: düz renk olmayan dolgu (`…with a fill that is not a plain colour…`), kısmi saydam dolgu (`…with a partly transparent fill…`), palet/şema/sistem dizinli renk (`…with a colour that is an index into a palette or a system colour…`); resmi dosyada bulunamayan resim şekli artık `N picture(s) whose data is not available…` ile sayılır.
- `PageItem` birleşimine `ShapeItem` (`type: 'shape'`) eklendi; `page.items` üzerinde tür türü dolaşan kod bu türü de görür (kendi çizicisini yazanlar için tip düzeyinde kırıcı).
- Belge görünümünde kalınlığı 0 olan şekil çizgisi (`<a:ln w="0">`) artık en ince çizgi olarak çizilir (önceden çizilmiyordu).
- `mount()`: dönüştürülmemiş doc ve RTF görünümlerinde `setReadOnly(false)` artık düzenlemeyi açmaz (düzenlemeler kaydedilemiyordu; xls'teki kuralın aynısı).
- `mount()`: dönüştürme bildirimi (`onNotice`), `locale` verilmediğinde ya da nesne olduğunda görünümler gibi tarayıcının dilinde verilir (önceden hep İngilizceydi). Dönüştürme başarısız olursa dosya reddedilmez, salt okunur gösterilir ve bildirilir.
- Köprü: `load`'da `data`, `url`, `blank`'ten tam olarak biri gerekir; hata iletisinin metni buna göre değişti.
- Örnek sayfa (`examples/html`): dönüştürmede addan yalnız `.doc`/`.rtf`/`.xls` uzantısı atılır.

### Bilinen sınırlar

Güncel liste [kullanım kılavuzundaki](docs/kullanim-kilavuzu.md) "Henüz olmayanlar" tablosundadır.
