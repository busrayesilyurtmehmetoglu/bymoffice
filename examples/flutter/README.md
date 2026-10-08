# BYM Office — Flutter örneği

> **Uyarı: bu örnek hiç derlenmedi ve hiç çalıştırılmadı.** Yazıldığı ortamda Flutter/Dart kurulu değildi. Kod `webview_flutter` 4.x arayüzüne (`WebViewController`, `addJavaScriptChannel`, `loadFlutterAsset`, `runJavaScript`) göre yazıldı; ilk derlemede küçük düzeltmeler gerekebilir. Emin olunmayan satırlar kodda `UNVERIFIED` diye işaretlidir. Köprü sayfasının kendisi tarayıcılarda (Chromium, Firefox, WebKit, mobil görünüm) otomatik testlerle sınandı; gerçek bir cihazdaki WebView'de sınanmadı.

Örnek, BYM Office editörünü bir WebView içinde gösterir: bir Excel ya da Word dosyası (xlsx, docx; dönüştürülerek xls, doc, RTF) seçilir, düzenlenir, kaydedilir. Editör uygulamanın içine gömülüdür; internet gerekmez.

## Dosyalar

| Dosya | İş |
|---|---|
| `lib/bym_office_view.dart` | Yeniden kullanılabilir parça: `BymOfficeController` (yükle, kaydet, geri al, akışlar) ve `BymOfficeView` (widget). Kendi uygulamanıza bu dosyayı kopyalayın |
| `lib/main.dart` | Dosya seçip açan ve kaydeden küçük uygulama |
| `pubspec.yaml` | Bağımlılıklar (`webview_flutter`, örnek için `file_picker`) ve varlık bildirimi |

## Kurulum

1. Platform klasörlerini üretin (bu depoda yalnız Dart kodu vardır):

   ```sh
   cd examples/flutter
   flutter create --platforms=android,ios,macos --org com.example --project-name bym_office_example .
   ```

2. Köprü sayfasını varlık olarak kopyalayın. Sayfa motoru içinde taşıyan tek bir HTML dosyasıdır ve motorun derlemesiyle üretilir:

   ```sh
   # depo kökünde
   pnpm install && pnpm build
   mkdir -p examples/flutter/assets
   cp dist/bymoffice.webview.html examples/flutter/assets/
   ```

   Motor her güncellendiğinde bu kopyalama yinelenir. Dosya `pubspec.yaml` içinde `assets/bymoffice.webview.html` olarak bildirilmiştir.

3. Çalıştırın:

   ```sh
   flutter pub get
   flutter run
   ```

## Kullanım

```dart
final editor = BymOfficeController();

// Widget ağacında, sınırlı boyutlu bir yerde:
BymOfficeView(controller: editor);

// Dosya açma: sayfa hazır olana dek kendisi bekler.
final document = await editor.load(bytes, name: 'rapor.xlsx', locale: 'tr');
print(document.sheets.map((s) => s.name));

// Değişiklikler ve hatalar:
editor.onChanged.listen((change) => print('değişti, geri alınabilir: ${change.canUndo}'));
editor.onError.listen((error) => print('${error.code}: ${error.message}'));

// Kaydetme: düzenlenmiş dosyanın baytları.
final Uint8List saved = await editor.save();

await editor.undo();
await editor.setReadOnly(true);

// Word belgesi: docx düzenlenir; `toolbar: true` biçimlendirme araç çubuğunu gösterir.
final belge = await editor.load(docxBytes, name: 'yazı.docx', toolbar: true);
print(belge.isDocument);   // true; belge.sheets boştur

// Eski biçimler (doc, RTF, xls): `convert: true` ile docx/xlsx'e çevrilip düzenlenir.
final eski = await editor.load(docBytes, name: 'eski.doc', convert: true);
print('${eski.convertedFrom} -> ${eski.format}');   // doc -> docx
print(eski.warnings);                                // taşınamayanlar (İngilizce), boş olabilir
final kayit = await editor.saveFile();               // kayit.bytes bir docx'tir; kayit.format == 'docx'

// `convert` verilmezse doc/RTF/xls salt okunurdur: eski.readOnly == true,
// ve `await editor.setReadOnly(false)` true (hâlâ salt okunur) döner.

// Boş dosyayla başlama:
await editor.loadBlank('docx', locale: 'tr');

// Belgedeki bağlantılar: sayfa açmaz, uygulamaya bildirir.
editor.onLink.listen((url) => print('bağlantı: $url'));

editor.dispose();
```

`onChanged` belge için de gelir; belgede `change.sheets` hep boştur. Dönüştürülen dosyada uygulama dosyanın adını ve türünü `format`'a göre güncellemelidir (`main.dart` adın sonundaki `.doc`/`.rtf`/`.xls` uzantısını değiştirir). Bu eklemeler de derlenmedi; `warnings` listesinin çözülmesi kodda `UNVERIFIED` işaretlidir.

Sayfa komutları: `insertSheet`, `renameSheet`, `deleteSheet` ve (2026-10-08'de eklenen, **derlenmemiş ve çalıştırılmamış**) `moveSheet(from, to)`, `setSheetState(index, state)`, `copySheet(index, {at, name})`, `setTabColor(index, color)`. `sheets` iletisindeki `tabColor` alanı sarmalayıcıda okunmuyor.

İzlenen değişiklikler (yalnız belgede; 2026-10-08'de eklendi, **derlenmemiş ve çalıştırılmamış**): `acceptAllRevisions()`, `rejectAllRevisions()`, `setTrackChanges(on)` kalan değişiklik sayısını ve izlemenin açık olup olmadığını (`BymOfficeRevisions`) döndürür; sayı ya da izleme durumu değişince `onRevisionsChanged` akışı haber verir. `loaded` iletisindeki `revisions` alanı sarmalayıcıda okunmuyor.

Her çağrı (`load`, `save`, `undo`, `redo`, `setReadOnly`, `setLocale`, `close`) iletiye bir `id` ekler ve aynı `id`'yi taşıyan yanıtı (`loaded`, `saved`, `done` ya da `error`) bekler. Yanıt 60 saniyede gelmezse (sayfanın hazır olmasını bekleme süresi dahil) `BymOfficeException('timeout', …)` fırlatılır. `dispose()` sayfaya `close` gönderir, JavaScript kanalını kaldırır ve bekleyen çağrıları `closed` koduyla sonlandırır; sonrasında yapılan çağrılar hemen hata verir. Sayfa yüklenince denetleyici bir `hello` gönderir; `ready` kanal kurulmadan önce gönderilip kaybolduysa sayfa onu yineler. Açılamayan bir dosya, o sırada açık olan dosyayı kapatmaz. Kullanıcı sekme değiştirince `onSheetChanged` akışı haber verir. Sayfa listesi değişince (sayfa eklendi, silindi, adı değişti; komutla, kullanıcı eliyle ya da geri almayla) `onSheetsChanged` akışı güncel listeyi ve gösterilen sayfanın dizinini verir; `insertSheet({at, name})` (yeni sayfanın dizinini döndürür), `renameSheet(index, name)` ve `deleteSheet(index)` sayfa işlemlerini yapar (bunlar da **derlenmedi**). Hata kodları: `unsupported`, `corrupt`, `load-failed`, `save-failed`, `bad-message`, `not-loaded`, `print-unavailable`, `print-failed`, `read-only`, `no-sheet`, `sheet-refused` (sonuncusunda `message` editörün dilinde, kullanıcıya gösterilebilir nedendir). Sözleşmenin tamamı `docs/webview-koprusu.md` dosyasındadır.

## Platform notları

- **Android:** `loadFlutterAsset` sayfayı `file:///android_asset/…` adresinden yükler; ek izin gerekmez. Uygulama internete çıkmıyorsa `INTERNET` izni de gerekmez. Eski cihazlarda sistem WebView'inin sürümü belirleyicidir; motor ES2019'a derlenir.
- **iOS:** `WKWebView` kullanılır. Sayfa güvenli alanı (çentik, ana ekran çizgisi) kendisi boş bırakır; `Scaffold` gövdesini ayrıca `SafeArea` ile sarmayın, yoksa boşluk iki kez uygulanır. Sayfanın `viewport` etiketi `maximum-scale=1` içerir; hücre editörü odaklandığında iOS'un sayfayı kendiliğinden yakınlaştırmasını önler.
- **macOS:** `webview_flutter` macOS'u `WKWebView` ile destekler (yeterince yeni bir `webview_flutter_wkwebview` sürümü gerekir). Uygulama korumalı alanda (sandbox) çalışıyorsa dosya seçmek için `com.apple.security.files.user-selected.read-write` yetkisi gerekir.
- **Windows:** `webview_flutter` Windows'u desteklemez. Windows'ta WebView2 tabanlı bir paket (ör. `webview_windows` ya da `flutter_inappwebview`) gerekir. Köprü sayfası bu kanalları da tanır (`window.chrome.webview.postMessage`, `window.flutter_inappwebview.callHandler('bymoffice', …)`), ama `bym_office_view.dart` yalnız `webview_flutter` için yazılmıştır; Windows için denetleyicinin WebView'e dokunan üç satırı (kanal kurma, sayfa yükleme, `runJavaScript`) o paketin karşılıklarıyla değiştirilmelidir. Bu yapılmadı ve denenmedi.

## Büyük dosyalarda base64 maliyeti

Dosya, uygulama ile sayfa arasında base64 metni olarak taşınır, çünkü her platformun ortak desteklediği tek taşıma biçimi dizgidir. Sonuçları:

- Base64 metni dosyadan üçte bir büyüktür; JavaScript dizgileri karakter başına 2 bayt tutabildiği için 20 MB'lık bir dosya, yükleme sırasında kısa süreliğine 100 MB'ın üzerinde bellek kullanabilir (Dart tarafındaki baytlar + base64 metni + `runJavaScript`'e verilen betik metni + sayfadaki dizgi + çözülmüş baytlar).
- `load` iletisi `runJavaScript` ile tek bir betik olarak gönderilir. Onlarca MB'lık betiklerin WebView'e geçirilmesi yavaştır ve platforma göre sınıra takılabilir; gerçek cihazda ölçülmedi.
- Sayfanın kendi kodlama/çözme süresi küçüktür (masaüstü tarayıcıda 5 MB için yükleme ≈ 0,3 sn, kaydetme < 0,1 sn ölçüldü).
- Denetleyici, Dart tarafındaki base64 dönüşümünü `compute` ile ayrı bir yalıtımda (isolate) yapar; arayüz takılmaz.

Çok büyük dosyalar düzenli olarak açılacaksa dosyayı yerel bir adresten sunup `load` iletisinde `url` vermek daha uygundur; bunun için sayfanın içerik güvenliği ilkesinin (`connect-src`) genişletilmesi gerekir, ayrıntısı `docs/webview-koprusu.md` içindedir.
