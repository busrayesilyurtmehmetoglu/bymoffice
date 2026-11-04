# BYM Office — Flutter örneği

> **Sınanma durumu (2026-11-04, Flutter 3.41, `webview_flutter` 4.x):** `flutter analyze` temiz; uygulama Android (APK), iOS ve macOS için derlendi. `integration_test/bridge_test.dart` denetleyiciyi gerçek bir WebView'de köprü sayfasına karşı çalıştırır (boş tablo → sayfa ekle → kaydet → yeniden aç; boş belge; Office dosyası olmayan dosyanın reddi; RTF'in docx'e dönüştürülmesi) ve macOS'ta (WKWebView) 4/4 geçti. Bu test Android ve iOS cihazda/öykünücüde koşturulmadı; dosya seçme (`file_picker`) elle denenmedi. Kodda `UNVERIFIED` işaretli satırlar testin ulaşmadığı yollardır.

Örnek, BYM Office editörünü bir WebView içinde gösterir: bir Excel ya da Word dosyası (xlsx, docx; dönüştürülerek xls, doc, RTF) seçilir, düzenlenir, kaydedilir. Editör uygulamanın içine gömülüdür; internet gerekmez.

## Dosyalar

| Dosya | İş |
|---|---|
| `lib/bym_office_view.dart` | Yeniden kullanılabilir parça: `BymOfficeController` (yükle, kaydet, geri al, akışlar) ve `BymOfficeView` (widget). Kendi uygulamanıza bu dosyayı kopyalayın |
| `lib/main.dart` | Dosya seçip açan ve kaydeden küçük uygulama |
| `pubspec.yaml` | Bağımlılıklar (`webview_flutter`, örnek için `file_picker`) ve varlık bildirimi |

## Kurulum

1. Platform klasörlerini üretin (bu depoda yalnız Dart kodu vardır; üretilenler `.gitignore`'dadır):

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

3. Platform ayarları (Xcode 27 ile görüldü):

   - **macOS:** korumalı alandaki (sandbox) uygulamada WKWebView, `file://` adresindeki köprü sayfasını bile **`com.apple.security.network.client` yetkisi olmadan açmaz**: sayfa yüklenmez, `ready` hiç gelmez, her istek `timeout` ile düşer. `macos/Runner/DebugProfile.entitlements` ve `Release.entitlements` dosyalarına ekleyin:

     ```xml
     <key>com.apple.security.network.client</key>
     <true/>
     ```

   - **Dağıtım hedefi:** `flutter create`'in yazdığı eski hedefler (macOS 10.15, iOS 9–13) güncel Xcode'da derlenmez. `macos/Podfile`'da `platform :osx, '12.0'`, `ios/Podfile`'da `platform :ios, '15.0'` yazın; iki Podfile'ın `post_install` döngüsüne de hedefi ekleyin, ör. macOS için:

     ```ruby
     target.build_configurations.each { |c| c.build_settings['MACOSX_DEPLOYMENT_TARGET'] = '12.0' }
     ```

     (iOS'ta `IPHONEOS_DEPLOYMENT_TARGET` = `15.0`.) `Runner.xcodeproj` içindeki aynı ayarı da Xcode'dan ya da metin düzenleyiciyle yükseltin.

4. Çalıştırın ve sınayın:

   ```sh
   flutter pub get
   flutter run
   flutter test integration_test -d macos     # ya da bağlı bir Android/iOS cihazın kimliği
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

// Kullanıcının kaydetmesi: `userSaves: true` ile sayfanın araç çubuğunda Kaydet düğmesi çıkar ve
// Ctrl/Cmd+S çalışır; dosya `onUserSave` ile gelir, uygulama saklayıp sonucu bildirir (derlendi; testte çalıştırılmadı).
await editor.load(bytes, name: 'rapor.xlsx', toolbar: true, userSaves: true);
editor.onUserSave.listen((file) async {
  try {
    await upload(file.bytes, file.format);          // file.format: 'docx' | 'xlsx'; file.convertedFrom dönüştürüldüyse
    await editor.saveResult(file, true);            // aynı kaydın kimliğiyle (file.saveId); sayfa "Kaydedildi" der
  } catch (_) {
    await editor.saveResult(file, false, message: 'Sunucuya ulaşılamadı');   // sayfa "Dosya kaydedilemedi (…)" der
  }
});

// Kendiliğinden kaydetme: `autoSave: true` (ya da `{'delayMs': 5000, 'maxWaitMs': 60000}`) ile sayfa, son
// değişiklikten 2 sn sonra dosyayı kendisi verir; aynı `onUserSave` akışına `file.byAuto == true` ile gelir ve
// aynı biçimde `saveResult` ile yanıtlanır. Yazılmakta olan hücreyi bölmez; `conflict: true` yanıtında onay
// açılmaz, kendiliğinden kayıt durur ve onay kullanıcının Kaydet'ine kalır (derlendi; testte çalıştırılmadı).
// Çakışma denetimi uygulamadadır: `byAuto` dosyayı saklamadan önce elinizdeki sürümle karşılaştırın; değiştiyse
// yazmadan `saveResult(file, false, conflict: true)` deyin. Sayfa köprüde sürüm (ETag) bilgisi tutmaz.
await editor.load(bytes, name: 'rapor.xlsx', toolbar: true, autoSave: true);

// Çakışma: dosya, uygulamanın elinde (sunucuda) o arada değiştiyse üzerine yazmayın; sayfaya söyleyin.
// Sayfa kullanıcıya sorar ("Üzerine yaz" / "Vazgeç"); "Üzerine yaz" denirse dosya `onUserSave` ile
// yeniden gelir ve bu kez `file.overwrite == true`'dur: koşulsuz yazın. "Vazgeç"te hiçbir şey gelmez.
//   if (!file.overwrite && await changedOnServer()) {
//     await editor.saveResult(file, false, conflict: true, message: 'Başka biri az önce kaydetti');
//   }

await editor.undo();
await editor.setReadOnly(true);

// Word belgesi: docx düzenlenir; `toolbar: true` biçimlendirme araç çubuğunu gösterir.
final belge = await editor.load(docxBytes, name: 'yazı.docx', toolbar: true);
print(belge.isDocument);   // true; belge.sheets boştur

// Parola korumalı dosya: parola biliniyorsa `password` ile; bilinmiyorsa sayfa kullanıcıya kendi kutusunda sorar
// ve `onPasswordRequired` de ateşlenir (kendi kutunuzu gösterecekseniz `passwordPrompt: false` verin ve
// `sendPassword(parola)` ile yanıtlayın; `sendPassword(null)` vazgeçer → BymOfficeException('password-cancelled')).
// Dosya aynı parolayla şifreli kaydedilir (`sifreli.encrypted == true`). Yüklemenin 60 sn'lik zaman aşımı
// parola beklenirken de işler: `BymOfficeController(requestTimeout: …)` ile uzatın. (Derlendi; testte çalıştırılmadı.)
final sifreli = await editor.load(bytes, name: 'gizli.xlsx', password: 'Parola1234_');

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

`onChanged` belge için de gelir; belgede `change.sheets` hep boştur. Dönüştürülen dosyada uygulama dosyanın adını ve türünü `format`'a göre güncellemelidir (`main.dart` adın sonundaki `.doc`/`.rtf`/`.xls` uzantısını değiştirir). `convert` ile açma ve kaydetme testte RTF dosyasıyla çalıştırıldı; `warnings` listesinin çözülmesi kodda `UNVERIFIED` işaretlidir (testteki dosyada liste boştu).

Sayfa komutları: `insertSheet`, `renameSheet`, `deleteSheet` ve (2026-10-08'de eklenen; derlendi, testte çalıştırılmadı) `moveSheet(from, to)`, `setSheetState(index, state)`, `copySheet(index, {at, name})`, `setTabColor(index, color)`. `sheets` iletisindeki `tabColor` alanı sarmalayıcıda okunmuyor.

İzlenen değişiklikler (yalnız belgede; 2026-10-08'de eklendi; derlendi, testte çalıştırılmadı): `acceptAllRevisions()`, `rejectAllRevisions()`, `setTrackChanges(on)` kalan değişiklik sayısını ve izlemenin açık olup olmadığını (`BymOfficeRevisions`) döndürür; sayı ya da izleme durumu değişince `onRevisionsChanged` akışı haber verir. `loaded` iletisindeki `revisions` alanı sarmalayıcıda okunmuyor.

Yorumlar (yalnız belgede; 2026-10-12'de eklendi; derlendi, testte çalıştırılmadı): `addComment(text, {anchor, focus})` ve `replyComment(commentId, text)` yeni yorumun kimliğini döndürür; `editComment(commentId, text)`, `resolveComment(commentId, {done})`, `removeComment(commentId)`. Liste her değiştiğinde `onCommentsChanged` akışı tamamını (`BymOfficeComment`: `id`, `author`, `text`, `done`, `date`, `parent`) verir. `loaded` iletisindeki `comments` ve `commentsOnly` alanları sarmalayıcıda okunmuyor; `comments` korumalı belgede `loaded.readOnly` gelmez (metin sayfada kilitlidir, yorum komutları çalışır).

Her çağrı (`load`, `save`, `undo`, `redo`, `setReadOnly`, `setLocale`, `close`) iletiye bir `id` ekler ve aynı `id`'yi taşıyan yanıtı (`loaded`, `saved`, `done` ya da `error`) bekler. Yanıt 60 saniyede gelmezse (sayfanın hazır olmasını bekleme süresi dahil) `BymOfficeException('timeout', …)` fırlatılır. `dispose()` sayfaya `close` gönderir, JavaScript kanalını kaldırır ve bekleyen çağrıları `closed` koduyla sonlandırır; sonrasında yapılan çağrılar hemen hata verir. Sayfa yüklenince denetleyici bir `hello` gönderir; `ready` kanal kurulmadan önce gönderilip kaybolduysa sayfa onu yineler. Açılamayan bir dosya, o sırada açık olan dosyayı kapatmaz. Kullanıcı sekme değiştirince `onSheetChanged` akışı haber verir. Sayfa listesi değişince (sayfa eklendi, silindi, adı değişti; komutla, kullanıcı eliyle ya da geri almayla) `onSheetsChanged` akışı güncel listeyi ve gösterilen sayfanın dizinini verir; `insertSheet({at, name})` (yeni sayfanın dizinini döndürür), `renameSheet(index, name)` ve `deleteSheet(index)` sayfa işlemlerini yapar (testte yalnız `insertSheet` çalıştırıldı). `exportPdf({dpi, image, quality, sheets, title, author})` açık dosyanın PDF'ini (`BymOfficePdf`: `bytes`, `name`, `pages`) döndürür; uzun belgede 60 saniyelik yanıt süresi yetmeyebilir (`requestTimeout`'u artırın) — derlendi, testte çalıştırılmadı. Hata kodları: `unsupported`, `corrupt`, `load-failed`, `save-failed`, `bad-message`, `not-loaded`, `print-unavailable`, `print-failed`, `pdf-unavailable`, `pdf-failed`, `pdf-too-large` (dosya 64 MB'ı aşıyor: daha düşük `dpi` isteyin), `read-only`, `no-sheet`, `sheet-refused` (sonuncusunda `message` editörün dilinde, kullanıcıya gösterilebilir nedendir). Sözleşmenin tamamı `docs/webview-koprusu.md` dosyasındadır.

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
