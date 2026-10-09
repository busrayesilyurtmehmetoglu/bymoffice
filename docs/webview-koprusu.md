# BYM Office — WebView köprüsü

Mobil ve masaüstü uygulamalar editörü bir WebView içinde gösterir. Bunun için tek bir dosya yeter: `dist/bymoffice.webview.html`. Dosya motoru içinde taşır, başka hiçbir dosyaya ve internete ihtiyaç duymaz; uygulamaya varlık (asset) olarak konur. Uygulama ile sayfa, aşağıdaki mesaj sözleşmesiyle konuşur. Aynı sayfa bir web uygulamasında `<iframe>` içinde de kullanılabilir.

> **Durum:** Sayfa ve sözleşme Chromium, Firefox, WebKit ve mobil görünüm profillerinde otomatik testlerle sınandı. **Gerçek bir cihazdaki WebView'de (Android, iOS, Windows, macOS) sınanmadı.** Aşağıdaki platform örnekleri derlenmedi; ilgili platformların belgelenmiş arayüzlerine göre yazılmış en küçük iskeletlerdir.

## 1. Sayfayı elde etme

```sh
pnpm install && pnpm build   # dist/bymoffice.webview.html
```

Sayfa `src/embed/host.html` şablonundan üretilir; derleme motoru (`bymoffice.min.js`) ve paketle gelen fontu içine gömer. Motor güncellenince dosya yeniden kopyalanır. Sözleşmenin sürümü paket sürümüdür ve `ready` iletisinde bildirilir.

**Paketle gelen font sayfanın içindedir.** Bookman Old Style için kullanılan TeX Gyre Bonum'un dört dosyası sayfaya base64 olarak gömülüdür (`<script type="text/plain" data-bym-font="…">` veri blokları; tarayıcı bunları çalıştırmaz) ve yanlarında lisans metni yorum olarak durur. Bir dosya Bookman Old Style ile yazılmış metin çizdiğinde, cihazda o font yoksa ve `load.fonts` ile o ad için font verilmemişse, baytlar çözülüp sayfaya eklenir: **ağ isteği yapılmaz**, uygulamanın bir şey yapması gerekmez. Bu aileyi kullanmayan dosyada bloklara dokunulmaz. Kurallar ve sınırlar kullanım kılavuzunun "Fontlar" bölümündedir (4. madde).

Fontlar sayfaya 770.873 bayt ekler (fontsuz 1.510.719 bayt, fontlu 2.281.592 bayt). Küçük dosya isteyen fontsuz sayfa derleyebilir:

```sh
BYM_NO_FONTS=1 pnpm build    # dist/bymoffice.webview.html fontsuz; dist/fonts/ yine üretilir
```

Fontsuz sayfada paketle gelen font kapalıdır (istenecek bir adres de yoktur): Bookman Old Style cihazın yedek fontuyla gösterilir, bildirim yapılmaz. Kendi fontunuzu `load.fonts` ile yine verebilirsiniz.

## 2. Taşıma

Her iki yönde de iletiler **JSON metni (dizgi)** olarak taşınır; böylece her platform aynı kodu kullanır.

**Uygulama → sayfa**

- `window.BymOfficeBridge.receive(jsonMetni)` — uygulama bunu WebView'in "JavaScript çalıştır" çağrısıyla (`evaluateJavascript`, `runJavaScript`, `evaluateJavaScript`, `ExecuteScriptAsync`, `injectJavaScript`) çağırır. İşlev bir şey döndürmez.
- `window` üzerindeki `message` olayı — yalnız sayfa bir `<iframe>` içindeyken, yalnız kapsayan sayfadan gelen ve yalnız o sayfanın kökeni sayfanın izin listesindeyse (Bölüm 5). Varsayılan olarak **kapalıdır**.
- `window.chrome.webview` üzerindeki `message` olayı — Windows WebView2 (`PostWebMessageAsString` / `PostWebMessageAsJson`).

`receive` ve olaylar dizginin yanında, ayrıştırılmış nesneyi de kabul eder.

**Sayfa → uygulama**

Sayfa, aşağıdaki kanallardan **var olan ilkine** gönderir (her gönderimde yeniden bakılır):

| Sıra | Kanal | Platform |
|---|---|---|
| 1 | `window.BymOfficeHost.postMessage(json)` | Flutter `webview_flutter` (JavaScriptChannel), Android `addJavascriptInterface` |
| 2 | `window.webkit.messageHandlers.bymoffice.postMessage(json)` | iOS / macOS `WKWebView` |
| 3 | `window.flutter_inappwebview.callHandler('bymoffice', json)` | `flutter_inappwebview` (Windows'ta WebView2 üstünde çalıştığı için WebView2'nin kendi kanalından önce gelir) |
| 4 | `window.chrome.webview.postMessage(json)` | Windows WebView2 |
| 5 | `window.ReactNativeWebView.postMessage(json)` | React Native |
| 6 | `window.parent.postMessage(json, köken)` | `<iframe>`; yalnız izin listesindeki kökene, liste boşsa hiç |

Kanal, sayfa yüklendikten sonra da eklenebilir. Kanal yokken gönderilemeyen iletiler saklanır ve kanal 50 ms aralıklarla, **son başarılı gönderimden sonra** en çok 10 saniye aranır. Süre dolduktan sonra da uygulamanın her `receive` çağrısında saklanan iletiler (önce `ready`) yeniden denenir. Saklananlar için kurallar:

- `ready`, `loaded`, `saved`, `error`, `done` asla düşürülmez ve sıraları korunur.
- Durum bildiren iletilerden (`changed`, `selection`, `sheet`, `sheets`) her türden yalnız sonuncusu tutulur; birleştirilen `changed` iletisinin `sheets` alanı, yerine geçtiklerinin sayfalarını da içerir.
- Saklanan iletilerin toplamı 64 milyon karakteri aşarsa en eski `saved` iletileri, aynı `id` ile `error` / `save-failed` iletisine çevrilir (dosya baytları bırakılır, isteğin yanıtı kaybolmaz).

Uygulama, sayfanın gönderdiği iletiye **aynı çağrının içinden** `receive` ile yanıt verebilir (ör. `ready` gelir gelmez `load`); köprünün gönderim kuyruğu buna dayanıklıdır.

`ready`'yi kaçırdığını düşünen uygulama (ör. dinleyicisini sayfa yüklendikten sonra kuran bir üst sayfa) `{ "type": "hello" }` gönderir; sayfa `ready`'yi yeniden gönderir.

## 3. Sözleşme

`id` isteğe bağlı bir dizgidir; verilirse o iletinin yanıtında (`loaded`, `saved`, `done`, `error`) aynen geri döner. İstek/yanıt eşleştirmesi bununla yapılır. `id` verilen her ileti tam olarak bir yanıt alır.

İletiler geliş sırasıyla, birer birer işlenir: `load`'un hemen ardından gönderilen `save`, yükleme bitince çalışır. İki istisna vardır ve ikisi de sıra beklemez, böylece asılı kalmış bir yükleme köprüyü kilitleyemez:

- `close` ve `hello` geldiği anda işlenir.
- `close` ve yeni bir `load`, sürmekte olan (ya da sırada bekleyen) yüklemeyi **iptal eder**: getirme isteği kesilir, o yükleme `loaded` göndermez; `id`'si varsa `error` / `load-failed` ("The load was cancelled") alır.

### Uygulama → sayfa

| İleti | Alanlar | İş |
|---|---|---|
| `load` | `data` (base64), `url` **ya da** `blank` (`"docx"` / `"xlsx"`); isteğe bağlı `id`, `name`, `locale`, `readOnly`, `sheet`, `convert`, `toolbar`, `save`, `author`, `trackChanges`, `fonts` | Dosyayı açar (ya da boş bir dosyayla başlar) ve gösterir. Yanıt: `loaded` ya da `error` |
| `save` | isteğe bağlı `id` | Düzenlenmiş dosyanın baytlarını ister. Tabloda önce yazılmakta olan uygulanır: açık hücre düzenleyicisi, **not alanı** ve **sekme adı alanı** (`commitPending`). Uygulanamayan bir not ya da ad varsa alan açık kalır ve hiçbir şey kaydedilmez. **Açık kural kutusu** (veri doğrulama, koşullu biçim) kaydetmeyle uygulanmaz: içinde değişiklik yoksa kapatılır ve kaydedilir; değişiklik varsa kutu açık kalır ve hiçbir şey kaydedilmez (`save-failed`; kullanıcıya "Önce açık kural kutusunu onaylayın ya da kapatın" denir). Yanıt: `saved` ya da `error` (`save-failed`) |
| `setReadOnly` | `readOnly` (boolean); isteğe bağlı `id` | Salt okunur kipini açar/kapatır (tablo ve belge). Yazılamayan biçimlerde (dönüştürülmemiş doc, RTF, xls) düzenleme açılamaz: yanıt `readOnly: true` taşır |
| `setLocale` | `locale` (ör. `"tr"`, `"en"`); isteğe bağlı `id` | Sayı/tarih biçimlerinin, görünüm metinlerinin ve sayfanın kendi yazılarının dilini değiştirir |
| `undo`, `redo` | isteğe bağlı `id` | Geri al, yinele (tablo ve belge). Bir şey değişirse `changed` gelir. Tabloda not ya da sekme adı yazılırken gelen komut yazılanı **atmaz**: önce uygular ve o komutta başka bir şey yapmaz (sonraki `undo` onu geri alır); hücre düzenlenirken eskisi gibi düzenlemeyi bırakır. **Kural kutusu açıkken** komut hiçbir şey yapmaz (kutu kiplidir; yanıt yine `done`'dır, `changed` gelmez) |
| `print` | isteğe bağlı `id`, `dpi` (sayı, varsayılan 200), `sheets` (`"active"` varsayılan / `"all"`; yalnız tabloda) | Açık belgeyi ya da tabloyu yazdırır: sayfaları hazırlar ve sayfanın `window.print()`'ini çağırır. Yanıt: yazdırma penceresi kapanınca `done`, yapılamazsa `error` (aşağıda "Yazdırma") |
| `insertSheet` | isteğe bağlı `id`, `at` (0 tabanlı yer; verilmezse sona), `name` (verilmezse görünümün dilinde sıradaki varsayılan ad: `Sayfa2`, `Sheet2`) | Boş bir çalışma sayfası ekler ve gösterir (tek geri alma adımı). Yanıt: `done` (`index` ve `name` ile) ya da `error` |
| `renameSheet` | `index`, `name`; isteğe bağlı `id` | Sayfayı yeniden adlandırır; formüller yeni adı izler. Ad aynıysa hiçbir şey değişmez, yine `done` gelir |
| `deleteSheet` | `index`; isteğe bağlı `id` | Sayfayı siler (onay sorulmaz; `undo` geri getirir) |
| `moveSheet` | `from`, `to` (sayfanın taşındıktan sonraki 0 tabanlı dizini); isteğe bağlı `id` | Sayfayı taşır; formüller değişmez |
| `setSheetState` | `index`, `state` (`"visible"`, `"hidden"`, `"veryHidden"`); isteğe bağlı `id` | Sayfayı gizler ya da gösterir (görünür son sayfa gizlenemez) |
| `copySheet` | `index`; isteğe bağlı `id`, `at` (kopyanın yeri; verilmezse hemen arkası), `name` (verilmezse `Ad (2)`) | Sayfanın kopyasını ekler ve gösterir |
| `setTabColor` | `index`; isteğe bağlı `id`, `color` (`"#RRGGBB"`; verilmezse ya da `null` ise renk kaldırılır) | Sekme rengini değiştirir |
| `acceptAllRevisions`, `rejectAllRevisions` | isteğe bağlı `id` | Belgedeki izlenen değişikliklerin tümünü kabul ya da reddeder (tek geri alma adımı). Çözülemeyen türler ve korumalı belge değişmez. Yanıt: `done` (`revisions`: kalan değişiklik sayısı, `trackChanges`: izleme açık mı) ya da `error` |
| `setTrackChanges` | `on` (boolean); isteğe bağlı `id` | Belgenin "değişiklikleri izle" ayarını açar ya da kapatır (geri alınabilir; `save` ile dosyaya yazılır). Kapatılamayan belgede (uygulanan `trackedChanges` koruması, ya da `load.trackChanges: true`) hiçbir şey değişmez; `load.trackChanges: true` ile izlenen belgede `on: true` da dosyanın ayarına dokunmaz (izleme zaten açıktır). Yanıt: `done` (`revisions`, `trackChanges`: **şimdi geçerli olan** durum) ya da `error` |
| `close` | isteğe bağlı `id` | Dosyayı kapatır, süren yüklemeyi iptal eder, boş duruma döner. Dosya yokken de hata vermez |
| `saveResult` | `saveId` (o `saved` iletisinin `saveId`'si), `ok` (boolean); isteğe bağlı `message` (kısa metin), `id` | **Kullanıcının** başlattığı kaydın (`by: "user"` taşıyan `saved` iletisi; aşağıda "Kullanıcının kaydetmesi") sonucunu bildirir: `ok: true` ise sayfa "Kaydedildi" der, `false` ise "Dosya kaydedilemedi". Yalnız **o anda beklenen** kaydın `saveId`'siyle geldiğinde uygulanır; değilse `bad-message`. Sıra beklemez. `id` verilmişse `done` ile onaylanır |
| `hello` | — | Sayfa `ready`'yi yeniden gönderir |

`setReadOnly`, `setLocale`, `undo`, `redo`, `print`, `insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `saveResult` ve `close` başarılı olduğunda, **`id` verilmişse** `{ "type": "done", "id": …, "command": "<type>" }` ile onaylanır; `id` verilmemişse yanıt gelmez. `undo`/`redo` için `done`, komutun işlendiğini bildirir; geri alınacak bir şey olup olmadığını `changed` iletisi gösterir.

```json
{ "type": "load", "id": "1", "data": "UEsDBBQ…", "name": "rapor.xlsx", "locale": "tr", "readOnly": false, "sheet": 0 }
{ "type": "save", "id": "2" }
{ "type": "setReadOnly", "readOnly": true }
{ "type": "load", "id": "3", "data": "0M8R4KGx…", "name": "eski.doc", "convert": true, "toolbar": true }
{ "type": "load", "id": "4", "blank": "docx", "locale": "tr" }
```

Sayfa işlemleri (`insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`):

- Yalnız tabloda (xlsx; `convert` ile açılmış xls) çalışır ve görünümün kendi sekme şeridindeki işlemlerle aynı koddur: her biri tek geri alma adımıdır, kurallar ve sınırlar aynıdır ([kullanım kılavuzu](kullanim-kilavuzu.md), "Sayfa ekleme, silme, yeniden adlandırma"). O sırada düzenlenen hücre, not ve sekme adı önce yazılır (`commitPending`); içinde değişiklik olmayan açık kural kutusu kapatılır.
- Güncel liste **her zaman** `sheets` iletisiyle gelir (aşağıda), `id` verilmemiş olsa da; `done` komutun **yapıldığını** (ya da yapılacak bir şey olmadığını: sayfa zaten o adı, yeri, durumu ya da rengi taşıyor) bildirir; görünümün yapmadığı komuta `done` denmez. `insertSheet`'in `done` yanıtı yeni sayfanın `index` ve `name` alanlarını da taşır: `{ "type": "done", "id": …, "command": "insertSheet", "index": 1, "name": "Sayfa2" }`; `copySheet`'inki de kopyanın `index` ve `name` alanlarını taşır. Kopyaya alınamayan içerik (koşullu biçim, doğrulama, görsel…) köprüde ayrıca bildirilmez; kurallar kullanım kılavuzunda, "Sayfa taşıma, gizleme, kopyalama, sekme rengi" başlığındadır.
- Alanlar denetlenir: `index`, `at`, `from` ve `to` 0 ya da daha büyük tam sayı, `name` en çok 255 karakterlik dizge, `state` üç değerden biri, `color` `#RRGGBB` biçiminde olmalıdır; değilse `bad-message` (hiçbir şey denenmez). Adın sayfa adı kurallarına uyması (en çok 31 karakter, `\ / ? * [ ] :` yok, başka sayfanın adı değil…) görünümün işidir ve aşağıdaki `sheet-refused` ile reddedilir.
- Hatalar: açık dosya yoksa `not-loaded`; açık dosya bir belgeyse ya da `index` hiçbir sayfanın dizini değilse (`at` için: sayfa sayısından büyükse) `no-sheet`; görünüm salt okunursa (`load.readOnly`, `setReadOnly(true)` ya da yazılamayan biçim) `read-only`; görünüm işlemi reddederse ya da **yapamazsa** `sheet-refused`: uygulanamayan bir not ya da sekme adı alanı açıksa, içinde değişiklik yapılmış bir kural kutusu açıksa ("… a dialog of rules is open with changes: apply or close it first"; kutu açık kalır, sayfalara dokunulmaz), ya da salt okunur olmayan görünüm sayfa ekleyemedi/kopyalayamadı (önceden bu durumlarda `done` ya da yanlış olarak `read-only` dönüyordu). `sheet-refused` iletisinin `message` alanı nedeni **görünümün dilinde** verir ve kullanıcıya gösterilebilir (ör. "Başka bir sayfa bu adı taşıyor", son görünür sayfa silinemez, grafik/pivot içeren kitapta silme yok, düzenlenen hücredeki değer veri doğrulamaya takıldı); kod hepsinde aynıdır.

```json
{ "type": "insertSheet", "id": "7", "name": "Özet", "at": 0 }
{ "type": "renameSheet", "id": "8", "index": 1, "name": "Veri 2026" }
{ "type": "deleteSheet", "id": "9", "index": 2 }
{ "type": "moveSheet", "id": "10", "from": 0, "to": 2 }
{ "type": "setSheetState", "id": "11", "index": 1, "state": "hidden" }
{ "type": "copySheet", "id": "12", "index": 0 }
{ "type": "setTabColor", "id": "13", "index": 0, "color": "#FF8800" }
```

Yazdırma (`print`):

- Belgede (docx, doc, RTF) ve tabloda (xlsx, xls) çalışır, salt okunur dosyada da; `DocumentView.print()` / `WorkbookView.print()` ile aynı koddur (ayrıntı ve sınırlar: [kullanım kılavuzu](kullanim-kilavuzu.md), "Belgeyi yazdırma" ve "Tabloyu yazdırma").
- **Tabloda** `sheets: "active"` (varsayılan) gösterilen sayfayı, `"all"` görünür bütün çalışma sayfalarını yazdırır; başka bir değer, belge açıkken de, `bad-message` döner (belgede geçerli değerler yok sayılır). `load` iletisindeki `name`, üstbilgi/altbilgideki `&F` (dosya adı) olarak kullanılır. Yazdırılacak hiçbir şey yoksa (boş sayfa) ya da sayfa sınırların ötesinde uzunsa `error` / `print-failed` döner.
- `print` **sırayı tutmaz**: sayfalar hazırlanırken ve pencere açıkken sonraki iletiler işlenir. `id` verilmişse yanıt, yazdırma penceresi kapandığında gelir: `{ "type": "done", "id": …, "command": "print" }`. Kullanıcının yazdırdığı mı vazgeçtiği mi ayırt edilemez. Sayfalar hazırlanırken gelen ikinci `print` aynı işe katılır (ikinci pencere açılmaz; iki `id` de yanıtlanır).
- Sayfalar hazırlanırken dosya kapatılır ya da değiştirilirse (`close`, yeni `load`) iş `error` / `print-failed` ile biter.
- **WebView'lerde `window.print()`:** pencerede `print` işlevi hiç yoksa `error` / `print-unavailable` döner. Android `WebView` ve `WKWebView` ise çoğunlukla işlevi tanımlar ama **hiçbir şey yapmaz**; sayfa bunu ayırt edemez. Bu platformlarda uygulama, `print` göndermeden önce sayfaya kendi `window.print` işlevini yerleştirmelidir: işlev yerel yazdırmayı başlatır (Android: `webView.createPrintDocumentAdapter()` + `PrintManager`; iOS: `UIPrintInteractionController` ile `webView.viewPrintFormatter()`), yerel yazdırma bittiğinde sayfada `window.dispatchEvent(new Event('afterprint'))` çalıştırır. Yerel yazdırma WebView'in içeriğini yazdırma ortamı kurallarıyla basar; o sırada sayfada yalnız hazırlanmış belge sayfaları görünür. `afterprint` gelmezse iş, kullanıcının sayfaya ilk dokunuşunda ya da en geç 5 dakikada biter ve `done` o zaman gönderilir. Bu akış gerçek cihazda **denenmedi**.
- Sayfanın içerik güvenlik ilkesi yazdırma için gevşetilmez: sayfa görüntüleri `blob:` adresli `<img>` öğeleridir, kurallar satır içi bir `<style>` öğesidir; ikisine de ilke zaten izin verir.

Kullanıcının kaydetmesi (`load.save: true`, `saved` + `by: "user"`, `saveResult`):

- Kullanıcı Kaydet düğmesine ya da Ctrl/Cmd+S'e basınca sayfa, uygulama `save` göndermiş gibi dosyayı yazar ve **`saved`** iletisini gönderir: `{ "type": "saved", "by": "user", "saveId": "u1", "data": "<base64>", "format": "docx" | "xlsx", "name"?, "convertedFrom"? }`. Bu iletide **`id` yoktur** (hiçbir isteğin yanıtı değildir), `format` her zaman vardır ve **`saveId`** o kaydın kimliğidir (sayfa açık kaldıkça hiç yinelenmeyen bir dizge; biçimine güvenmeyin, aynen geri gönderin); `by: "user"` onu uygulamanın kendi `save` isteğinin yanıtından ayırır (o yanıt `by` taşımaz ve eskisi gibidir).
- Uygulama dosyayı saklar (diske yazar, sunucuya yollar…) ve sonucu **aynı `saveId` ile** bildirir: `{ "type": "saveResult", "saveId": "u1", "ok": true }` ya da `{ "type": "saveResult", "saveId": "u1", "ok": false, "message": "Sunucuya ulaşılamadı" }`. Sonuç gelene dek düğme "meşgul" görünür (`aria-busy`). `ok: true` → "Kaydedildi", düğme solar. `ok: false` → "Dosya kaydedilemedi (…)"; `message`'ın ilk 200 karakteri tek satır olarak, **metin olarak** gösterilir (HTML olarak basılmaz) ve değişiklikler kaydedilmemiş sayılmaya devam eder.
- **60 saniye** içinde `saveResult` gelmezse kayıt başarısız sayılır ("Dosya kaydedilemedi: zamanında yanıt gelmedi"); sonradan gelen sonuç hiçbir şeyi değiştirmez. **`saveResult` yalnız kendi kaydına uygulanır:** `saveId`'si o anda beklenen kaydınkiyle aynı değilse (zaman aşımına uğramış kaydın geç yanıtı, iki kez gönderilen yanıt, kapatılmış dosyanın kaydı, bekleyen kayıt yokken gelen yanıt) hiçbir şey değişmez ve uygulamaya `error` / `bad-message` ("No save with this "saveId" is waiting for a result") döner — böylece geç gelen bir "tamam", bir sonraki kaydı (ya da başka bir dosyanın kaydını) yanlışlıkla "kaydedildi" yapamaz. `saveId` yoksa ya da metin değilse, `ok` boolean değilse ya da `message` metin değilse de `bad-message` döner.
- Aynı anda tek kullanıcı kaydı bekler: sonuç beklenirken yeniden basılırsa ve arada bir şey değişmediyse ikinci ileti gönderilmez; değiştiyse sonuç geldikten sonra bir kez daha `saved` gönderilir. Sonuç beklenirken yazılan şey "kaydedildi" sayılmaz.
- Kaydetmeden önce yazılmakta olan uygulanır (`save` ile aynı kural); uygulanamıyorsa hiçbir şey gönderilmez ve neden görünümde söylenir. Kaydedilecek değişiklik yoksa düğme soluktur ve tuş hiçbir şey göndermez.
- Uygulamanın kendi `save` isteği her an çalışır ve kullanıcının "kaydedilmemiş değişiklik" durumunu **değiştirmez** (sayfa, baytların saklandığını yalnız `saveResult` ile öğrenir). `close` ya da yeni bir `load`, bekleyen kullanıcı kaydını sessizce bırakır.
- Sayfa, kaydedilmemiş değişiklik varken kapatılmaya karşı kendisi uyarmaz; bunu uygulama `changed` ve `saved` iletilerinden izler.

```json
{ "type": "load", "id": "1", "data": "UEsDBBQ…", "name": "rapor.xlsx", "toolbar": true, "save": true }
{ "type": "saved", "by": "user", "saveId": "u1", "data": "UEsDBBQ…", "format": "xlsx", "name": "rapor.xlsx" }
{ "type": "saveResult", "saveId": "u1", "ok": true }
```

`load` ayrıntıları:

- `data`, `url` ve `blank`'ten **tam olarak biri** verilir; hiçbiri ya da birden fazlası `bad-message` döner. `data` standart ya da URL-güvenli base64 olabilir; satır sonları ve boşluklar yok sayılır, dolgu (`=`) eksik olabilir. `data:` öneki konmaz.
- `url` sayfanın kendi adresine göre çözülür ve çözülen adres denetlenir. Kabul edilenler: `blob:` ve `data:` adresleri ile **sayfanın kendi kökenindeki** `http(s):` adresleri (göreli adresler dahil). Reddedilenler (`bad-message`): başka köken (`https://başka.example/…`, `//başka.example/…`), kullanıcı adı/parola içeren adres, `javascript:`, `file:` ve diğer şemalar, boşluk, denetim karakteri ya da ters eğik çizgi içeren adresler. Sayfa `http(s)` dışından yüklendiyse (uygulama varlığı: `file:`, özel şema) göreli ve `file:` adresler de reddedilir; yalnız `blob:` ve `data:` çalışır.
- Sayfa bir `<iframe>` içindeyse `url` **hiç kabul edilmez** (`bad-message`); dosya yalnız `data` ile verilir. Gerekçe: getirme, sayfanın kendi kökeninin çerezleriyle yapılır; `url` kabul edilseydi sayfayı gömen taraf o kökendeki dosyaları `load` + `save` ile okuyabilirdi.
- Getirme 30 saniyede yanıt vermezse ya da dosya 256 MB'tan büyükse (`Content-Length` ve okunan bayt sayısı) `load-failed` döner.
- `name` yalnız saklanır ve `saved` iletisinde geri verilir; biçim her zaman içerikten algılanır.
- `sheet` ilk gösterilecek sayfanın 0 tabanlı dizinidir; dosyada yoksa `bad-message` döner. Gizli (`hidden`, `veryHidden`) bir sayfanın dizini de verilebilir: sayfa gösterilir, ama sekmesi görünmez; kullanıcı başka sekmeye geçince ona sekmelerden dönemez.
- `blank`: `"docx"` ya da `"xlsx"`; dosya yerine yeni, boş bir belge ya da tabloyla başlar (`BymOffice.createBlank` ile aynı dosya; `locale` verilmişse onunla: Türkçede sayfa adı `Sayfa1`, belge A4). Başka bir değer `bad-message` döner. Boş dosya normal bir docx/xlsx'tir: düzenlenir, `save` ile alınır.
- `convert` (boolean, varsayılan `false`): `true` ise ve `readOnly` `true` değilse, kendi biçimine yazılamayan dosyalar **dönüştürülerek düzenlenebilir açılır**: doc ve RTF docx'e, xls xlsx'e çevrilir (`mount()`'un `convert` seçeneğiyle aynı kod). Gösterilen, düzenlenen ve `save` ile alınan dosya yeni biçimdedir; `loaded` ve o dosyanın her `saved` iletisi `format` (yeni biçim) ve `convertedFrom` (özgün biçim) taşır, `loaded.warnings` dönüştürmede taşınamayanları listeler (boş dizi olabilir). Uygulama dosyanın adını ve türünü buna göre güncellemelidir; `name` olduğu gibi geri verilir, sayfa onu değiştirmez. docx ve xlsx'te `convert` hiçbir şey değiştirmez. Dönüştürme başarısız olursa (yeni dosya yazılamadı ya da geri okunamadı) dosya reddedilmez: olduğu gibi, salt okunur gösterilir; `loaded` özgün `format`'ı, `readOnly: true`'yu ve nedeni anlatan tek bir `warnings` girdisini taşır, `convertedFrom` bulunmaz.
- `author` (string, en çok 255 karakter): belgede izlenen değişikliklerde yazar adı olarak yazılır (`w:author`); tabloda kullanıcının eklediği **hücre notlarının** yazarı olur (sağ tık → Not ekle ya da Shift+F2; dosyada `<author>`). Verilmezse `"Author"`. Metin olmayan ya da çok uzun değer `bad-message` hatasıdır. Not eklemek, düzenlemek ve silmek için köprüde ayrı bir komut yoktur; kullanıcı görünümden yapar, her değişiklik `changed` iletisiyle bildirilir ve `undo`/`redo`/`save` ile birlikte çalışır.
- `trackChanges` (boolean, varsayılan `false`; yalnız belgede): `true` ise belge izleme istemese de düzenlemeler izlenir. Belge izleme istiyorsa (ayarı açık ya da "yalnız izlenen değişiklik" korumalı) bu alan ne olursa olsun izlenir; `false` izlemeyi kapatmaz.
- **Korumalı belge:** `w:documentProtection` uygulanıyorsa ve türü `readOnly`, `comments` ya da `forms` ise belge `readOnly` ne olursa olsun **salt okunur** gösterilir; `loaded` iletisi `readOnly: true` ve `protection` taşır, `setReadOnly` ile açılamaz (yanıt `done` + `readOnly: true`), `save` `read-only` hatası verir. Ayrıntı: kullanım kılavuzunda "Belge koruması ve değişiklik izleme".
- `toolbar` (boolean, varsayılan `false`): `true` ise belge, sayfaların üstünde **biçimlendirme araç çubuğuyla** gösterilir (`DocumentEditor`; belge salt okunurken araç çubuğu gizlenir). `false` iken belge yalnız sayfalar olarak gösterilir (altta sayfa numarası ve yakınlaştırma çubuğu vardır); klavyeyle yazma ve kısayollar iki kipte de çalışır. **Tablo, köprü sayfasında her iki kipte de editör kabuğuyla (araç çubuğu, formül çubuğu, durum çubuğu) gösterilir**; `toolbar` tabloda bir şey değiştirmez. Komutlar ve iletiler iki kipte de aynıdır.
- `save` (boolean, varsayılan `false`): `true` ise **kullanıcı sayfanın içinden kaydedebilir**: araç çubuğunun başında Kaydet düğmesi görünür (belgede `toolbar: true` de gerekir; tablo hep kabukla gösterilir) ve Ctrl/Cmd+S aynı işi yapar (tarayıcının "sayfayı kaydet" penceresi açılmaz). Ayrıntı aşağıda, "Kullanıcının kaydetmesi". Yazılamayan dosyada (dönüştürülmemiş doc, RTF, xls; düzenlemeye kapalı korumalı belge) etkisizdir: düğme yoktur, tuş tarayıcıya kalır ve `loaded` iletisi `save` taşımaz. Boolean olmayan değer `bad-message` döner.
- `fonts` (dizi, en çok 16 öğe): dosya gösterilmeden önce kaydedilecek fontlar; `BymOffice.registerFont` ile aynı anlam (bkz. kullanım kılavuzu, "Fontlar"). Her öğe: `{ "family": "TeX Gyre Bonum", "data": "<base64>", "as": ["Bookman Old Style"], "weight": "bold", "style": "italic", "prefer": false }`.
  - `family` (zorunlu, en çok 255 karakter; harf, rakam, boşluk, tire dışındakiler atılır, geriye bir şey kalmalı). Font dosyası **ya** `data` (base64, en çok 32 M karakter ≈ 24 MB dosya; boş olamaz) **ya da** `url` ile verilir; ikisi birden ya da hiçbiri `bad-message`.
  - `url`, dosyanın `url` alanıyla **aynı kurala** uyar (sayfanın kendi kökeni, `blob:`, `data:`; iframe içinde hiç kabul edilmez). Font, sayfanın kendisi tarafından getirilir (en çok 24 MB, getirme süresi sınırıyla) ve bayt olarak kaydedilir; sayfanın içerik güvenliği ilkesinde `font-src` açılmaz.
  - `weight`: `"normal"`, `"bold"` ya da 1–1000 arası sayı; `style`: `"normal"` / `"italic"`; `as`: bir ad ya da ad dizisi (en çok 32, her biri en çok 255 karakter); `prefer`: boolean. Aynı yüz (aile adı büyük/küçük harf ayırmadan + ağırlık + stil; `"bold"` ile `700`, `"normal"` ile `400` aynıdır) bir `load` içinde iki kez geçemez (`fonts[1] is the same face as fonts[0] …`). Kurala uymayan her alan bütün `load`'u `bad-message` ile reddeder (hangi öğe olduğu iletide yazar: `fonts[1]: …`) ve hiçbir font kaydedilmez.
  - Fontların yüklenmesi en çok 3 saniye beklenir; dosya o sırada getirilir ve okunur. Yüklenemeyen ya da süresinde gelmeyen font `load`'u **başarısız kılmaz**: dosya gösterilir ve `loaded.warnings` nedeni taşır (`The font "X" could not be loaded: …`, `The font "X" is still loading; …`). Geç gelen font geldiğinde görünüm kendiliğinden yeniden dizilir.
  - Fontlar sayfa genelinde kayıtlı kalır (`close` ya da yeni `load` silmez); aynı aile + ağırlık + stil sonraki bir `load`'da yeniden gönderilirse yenisi yüklendiği anda eskisinin yerini alır (yüklenemezse eskisi kalır). Sonraki `load`'larda yeniden göndermek gerekmez. `load` başarısız olsa ya da başka bir `load` araya girse bile gönderilen fontlar kaydedilir.
- Yeni dosya **önce** okunur ve doğrulanır (base64, biçim, çalışma kitabı, `sheet`, varsa dönüştürme); ancak başarılıysa açık dosya kapatılıp yenisi gösterilir. Başarısız bir `load`, açık dosyayı ve kaydedilmemiş değişikliklerini olduğu gibi bırakır.

### Sayfa → uygulama

| İleti | Alanlar | Ne zaman |
|---|---|---|
| `ready` | `version` | Sayfa yüklendi, ileti alabilir. Bir kez gönderilir |
| `loaded` | `id`?, `format`, `kind`? (`"document"`, yalnız belgede), `sheets: [{ name, state, kind }]` (belgede `[]`), `activeSheet`, `readOnly`? (`true`: dosya yazılamayan bir biçimde ya da korumalı, hep salt okunur), `convertedFrom`?, `warnings`? (`string[]`), yalnız belgede: `protection`? (`"readOnly"` / `"comments"` / `"forms"` / `"trackedChanges"`: belgenin uyguladığı koruma), `protectionExceptions`? (`true`: belgede korumadan muaf aralıklar var; uygulanmazlar), `trackChanges`? (`true`: düzenlemeler izleniyor), `revisions`? (belgedeki izlenen değişiklik sayısı; yalnız sıfırdan büyükse), `save`? (`true`: `load.save` istendi ve dosya yazılabilir; kullanıcı sayfadan kaydedebilir) | `load` başarılı |
| `changed` | `sheets: number[]` (değişen sayfalar; belgede hep `[]`), `canUndo`, `canRedo` | Çalışma kitabı ya da belge her değiştiğinde (geri al/yinele dahil) |
| `saveState` | `dirty` (boolean: dosyanın şimdiki hali, uygulamanın en son sakladığını bildirdiği halden farklı), `saving` (boolean: bir kullanıcı kaydı yanıt bekliyor) | Yalnız `load.save: true` ile açılan dosyada, durum her değiştiğinde. **Kaydedilmemiş değişiklik var mı sorusunun tek doğru kaynağıdır**: uygulama bunu `changed` iletilerini sayarak ya da `canUndo`'ya bakarak çıkarmamalıdır (kayıt yazılırken yapılan değişiklik hiçbir `saved` iletisinde yoktur; kırpılmış bir geri alma geçmişi sonuna dek geri alınsa da dosya açıldığı hali olmayabilir). Kanal yokken yalnız sonuncusu saklanır |
| `saved` | `id`?, `data` (base64), `name`?, `format`? ve `convertedFrom`? (yalnız dönüştürülerek açılmış dosyada, ikisi birlikte) | `save` başarılı |
| `saved` (kullanıcıdan) | `by: "user"`, `saveId` (dizge; `saveResult`'ta geri gönderilir), `data` (base64), `format` (`"docx"` / `"xlsx"`, her zaman), `name`?, `convertedFrom`?; **`id` yok** | Kullanıcı sayfada Kaydet'e ya da Ctrl/Cmd+S'e bastı (`load.save: true`). Uygulama `saveResult` ile yanıtlar |
| `selection` | `sheet`, `range` (`"A1:B2"`, tek hücrede `"A1"`), `active` (`"A1"`) | Seçim değişince; **kare başına en çok bir kez**, yalnız o karedeki son seçim, bir öncekiyle aynıysa gönderilmez. Bekleyen seçim iletisi sayfa listesi değişirse (`sheets`) gösterilen sayfanın **yeni** diziniyle ve o sayfanın seçimiyle gönderilir |
| `revisions` | `count` (belgedeki izlenen değişiklik sayısı; en çok 2.000 sayılır), `trackChanges` (izleme açık mı) | Yalnız belgede: sayı ya da izleme durumu son bildirilenden farklı olunca (yazma, kabul/ret, geri al, `setTrackChanges`, kullanıcının araç çubuğundan yaptığı). Aynı değişikliğin `changed` iletisinden **önce** gelir. Başlangıç değeri `loaded` iletisindedir |
| `sheet` | `index`, `name` | Gösterilen sayfa (sekme) değişince |
| `sheets` | `sheets: [{ name, state, kind, tabColor? }]` (`loaded.sheets` ile aynı biçim, listenin tamamı), `activeSheet` (gösterilen sayfanın dizini) | Sayfa listesi değişince: sayfa eklendi, silindi, taşındı, kopyalandı, gizlendi/gösterildi, adı ya da sekme rengi değişti — komutla, kullanıcının sekme şeridinden ya da bunların geri alınması/yinelenmesiyle. Aynı değişikliğin `sheet` ve `changed` iletilerinden **önce** gelir. Belgede gelmez |
| `done` | `id`, `command`, `readOnly`? (`true`: `setReadOnly(false)` istendi ama dosya salt okunur kalıyor), `revisions`? ve `trackChanges`? (yalnız `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges` yanıtında: komuttan sonraki değişiklik sayısı ve izleme durumu) | `id` taşıyan `setReadOnly`, `setLocale`, `undo`, `redo`, `print`, `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges`, `insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `saveResult`, `close` başarılı (`print`'te: yazdırma penceresi kapandı; `insertSheet` ve `copySheet`'te ek olarak `index`, `name`) |
| `link` | `url` | Belgede bir bağlantı etkinleştirildi (Bölüm 7) |
| `error` | `id`?, `code`, `message` | Aşağıdaki durumlarda |

`sheets[].state`: `visible`, `hidden`, `veryHidden`. `sheets[].kind`: `worksheet`, `chartsheet`. `sheets[].tabColor`: sekme rengi olan sayfada CSS rengi (`"#RRGGBB"`; tema ve indeks renkleri çözülmüş olarak); rengi olmayan sayfada alan yoktur. (`loaded.sheets` de aynı alanı taşır.)

```json
{ "type": "ready", "version": "0.0.1" }
{ "type": "loaded", "id": "1", "format": "xlsx", "sheets": [{ "name": "Veri", "state": "visible", "kind": "worksheet" }], "activeSheet": 0 }
{ "type": "changed", "sheets": [0], "canUndo": true, "canRedo": false }
{ "type": "saved", "id": "2", "data": "UEsDBBQ…", "name": "rapor.xlsx" }
{ "type": "error", "id": "2", "code": "not-loaded", "message": "No file is open" }
{ "type": "loaded", "id": "3", "format": "docx", "kind": "document", "sheets": [], "activeSheet": 0, "convertedFrom": "doc", "warnings": [] }
{ "type": "changed", "sheets": [], "canUndo": true, "canRedo": false }
{ "type": "saved", "id": "5", "data": "UEsDBBQ…", "name": "eski.doc", "format": "docx", "convertedFrom": "doc" }
```

`warnings` girdileri İngilizce, geliştiriciye yönelik metinlerdir (ör. `"2 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over"`, `"Sheet: the content of a chart sheet is not in the model; it is written as an empty worksheet"`). Yalnız okuyucunun modele aldığı ama yeni dosyaya taşınamayan içerik listelenir; okuyucunun hiç okumadığı içerik (kılavuzdaki "Henüz olmayanlar" tablosu) yeni dosyada yoktur ve uyarılmaz. Bu yüzden arşivlik bir dosyanın özgün hali saklanmalıdır.

### Hata kodları

| Kod | Anlam |
|---|---|
| `bad-message` | İleti sözleşmeye uymuyor: bozuk JSON, bilinmeyen `type`, eksik ya da yanlış türde alan, geçersiz base64, izin verilmeyen `url` (iframe içinde her `url`), `data`/`url`/`blank`'ten birden fazlası ya da hiçbiri, `"docx"`/`"xlsx"` dışında bir `blank`, olmayan sayfa dizini. Ayrıca bir komut (`undo` vb.) beklenmedik biçimde başarısız olursa da bu kod döner |
| `not-loaded` | `save`, `undo`, `redo`, `setReadOnly`, `setLocale`, `print`, `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges` ya da bir sayfa komutu (`insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`) gönderildi ama açık dosya yok |
| `read-only` | Sayfa komutu gönderildi ama görünüm salt okunur; ya da `acceptAllRevisions` / `rejectAllRevisions` / `setTrackChanges` gönderildi ama belge salt okunur (öyle açıldı, `setReadOnly(true)` denildi, ya da yazılamayan bir biçimde) |
| `no-document` | `acceptAllRevisions`, `rejectAllRevisions` ya da `setTrackChanges` gönderildi ama açık dosya bir tablo |
| `no-sheet` | Sayfa komutu gönderildi ama açık dosya bir belge, ya da verilen dizinde sayfa yok |
| `sheet-refused` | Sayfa komutunu görünüm reddetti; `message` nedeni görünümün dilinde verir (kullanıcıya gösterilebilir) |
| `print-unavailable` | `print` gönderildi ama yazdırılamıyor: pencerede `print()` yok (ya da sayfayı barındıran uygulamanın verdiği görünüm yazdıramıyor) |
| `print-failed` | Yazdırma başlatılamadı ya da yarıda kaldı: belge ya da tablo sınırların ötesinde uzun, tabloda yazdırılacak bir şey yok, bir sayfa çizilemedi, `print()` hata fırlattı, ya da sayfalar hazırlanırken dosya kapatıldı |
| `unsupported` | Dosya gösterilemeyen bir türde (Word 95 ve öncesi, parola korumalı, Office dosyası olmayan içerik, boş içerik) |
| `corrupt` | Dosya xlsx gibi başlıyor ama okunamıyor (kesik, bozuk) |
| `load-failed` | `url` getirilemedi (ağ hatası, HTTP hata kodu, zaman aşımı, boyut sınırı, içerik güvenliği ilkesi engeli) ya da yükleme bir `close` / yeni `load` ile iptal edildi |
| `save-failed` | Kaydetme sırasında hata; ya da `saved` iletisi uygulamaya teslim edilemeden bekleyen iletiler sınırı aşıldı |

`message` İngilizce, geliştiriciye yönelik bir açıklamadır ve yalnız uygulamaya gider; sayfa onu kullanıcıya göstermez (sayfa kendi yerelleştirilmiş genel başlığını gösterir, Bölüm 6). Kullanıcıya gösterilecek metni uygulama `code`'a göre seçmelidir. Hiçbir hatalı ileti sayfayı çökertmez; sayfa sonraki iletileri işlemeyi sürdürür.

## 4. Platform bağlama örnekleri

Hepsinde akış aynıdır: kanalı kur → sayfayı yükle → `ready`'yi bekle (emin değilseniz `hello` gönderin) → `load` gönder.

**JavaScript dizgi sabiti üzerine not.** İletiyi `receive(…)` çağrısının içine dizgi sabiti olarak yazan platformlarda (Flutter, Android, iOS, React Native) JSON metni bir kez daha JSON'a çevrilir. Bu, iki karakter dışında güvenlidir: U+2028 ve U+2029 JSON'da ham olarak geçerlidir, ama eski JavaScript motorlarında (ES2019 öncesi) dizgi sabitini böler. Karar: aşağıdaki örnekler bu iki karakteri `\u2028` / `\u2029` kaçışına çevirir; taşıma biçimi (JSON metni + `receive`) değiştirilmedi. WebView2'nin `PostWebMessageAsString` çağrısı ve iframe `postMessage` betik metni üretmediği için bu sorundan etkilenmez.

### Flutter (`webview_flutter` 4.x)

Tam örnek: `examples/flutter/` (yeniden kullanılabilir `BymOfficeController` + `BymOfficeView`).

```dart
final controller = WebViewController()
  ..setJavaScriptMode(JavaScriptMode.unrestricted)
  ..addJavaScriptChannel('BymOfficeHost', onMessageReceived: (m) => onMessage(jsonDecode(m.message)))
  ..loadFlutterAsset('assets/bymoffice.webview.html');

Future<void> send(Map<String, dynamic> message) {
  final literal = jsonEncode(jsonEncode(message))
      .replaceAll(String.fromCharCode(0x2028), r'\u2028')
      .replaceAll(String.fromCharCode(0x2029), r'\u2029');
  return controller.runJavaScript('window.BymOfficeBridge.receive($literal);');
}
```

### Android (`WebView`, Kotlin)

```kotlin
class Host(private val onMessage: (String) -> Unit) {
    @JavascriptInterface
    fun postMessage(json: String) = onMessage(json)   // WebView'in kendi iş parçacığında çağrılır
}

webView.settings.javaScriptEnabled = true
webView.addJavascriptInterface(Host { json -> runOnUiThread { handle(JSONObject(json)) } }, "BymOfficeHost")
webView.loadUrl("file:///android_asset/bymoffice.webview.html")

fun send(message: JSONObject) {
    // JSONObject.quote does not escape the two line separators.
    val literal = JSONObject.quote(message.toString())
        .replace(0x2028.toChar().toString(), "\\u2028")
        .replace(0x2029.toChar().toString(), "\\u2029")
    webView.evaluateJavascript("window.BymOfficeBridge.receive($literal);", null)
}
```

`Base64.encodeToString(bytes, Base64.NO_WRAP)` kullanın; satır sonlu çıktı da kabul edilir (`JSONObject.quote` satır sonlarını kaçırır).

### iOS / macOS (`WKWebView`, Swift)

```swift
final class Host: NSObject, WKScriptMessageHandler {
    func userContentController(_ controller: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let json = message.body as? String else { return }
        handle(json)
    }
}

let configuration = WKWebViewConfiguration()
configuration.userContentController.add(Host(), name: "bymoffice")
let webView = WKWebView(frame: .zero, configuration: configuration)
let page = Bundle.main.url(forResource: "bymoffice.webview", withExtension: "html")!
webView.loadFileURL(page, allowingReadAccessTo: page.deletingLastPathComponent())

func send(_ message: [String: Any]) throws {
    let json = String(data: try JSONSerialization.data(withJSONObject: message), encoding: .utf8)!
    // Dizgiyi bir kez daha JSON'a çevirmek geçerli bir JavaScript dizgi sabiti verir.
    let literal = String(String(data: try JSONSerialization.data(withJSONObject: [json]), encoding: .utf8)!.dropFirst().dropLast())
        .replacingOccurrences(of: String(UnicodeScalar(0x2028)!), with: "\\u2028")
        .replacingOccurrences(of: String(UnicodeScalar(0x2029)!), with: "\\u2029")
    webView.evaluateJavaScript("window.BymOfficeBridge.receive(\(literal));")
}
```

`WKUserContentController` işleyiciyi güçlü tutar; görünüm kapanırken `removeScriptMessageHandler(forName: "bymoffice")` çağırın.

### Windows (WebView2, C#)

```csharp
await webView.EnsureCoreWebView2Async();
webView.CoreWebView2.WebMessageReceived += (s, e) => Handle(e.TryGetWebMessageAsString());
// Sayfayı yerel klasörden sanal bir ana makine adıyla sun:
webView.CoreWebView2.SetVirtualHostNameToFolderMapping("bymoffice.example", assetsFolder, CoreWebView2HostResourceAccessKind.Deny);
webView.CoreWebView2.Navigate("https://bymoffice.example/bymoffice.webview.html");

void Send(object message) => webView.CoreWebView2.PostWebMessageAsString(JsonSerializer.Serialize(message));
```

Sayfa WebView2 iletilerini `window.chrome.webview` üzerinden dinler; `ExecuteScriptAsync("window.BymOfficeBridge.receive(…)")` da kullanılabilir.

### React Native (`react-native-webview`)

```jsx
const ref = useRef(null);
const send = (message) => {
  const literal = JSON.stringify(JSON.stringify(message))
    .split(String.fromCharCode(0x2028)).join('\\u2028')
    .split(String.fromCharCode(0x2029)).join('\\u2029');
  ref.current.injectJavaScript(`window.BymOfficeBridge.receive(${literal}); true;`);
};

<WebView
  ref={ref}
  originWhitelist={['*']}
  source={require('./assets/bymoffice.webview.html')}
  onMessage={(event) => handle(JSON.parse(event.nativeEvent.data))}
/>;
```

Android'de `require` ile verilen yerel HTML yayın derlemesinde yüklenmeyebilir; sayfayı `android/app/src/main/assets/` altına koyup `source={{ uri: 'file:///android_asset/bymoffice.webview.html' }}` kullanın.

### `<iframe>` (web)

iframe taşıması **varsayılan olarak kapalıdır**: derlenen sayfa hiçbir üst sayfayla konuşmaz. Açmak için sayfanın içindeki izin listesi doldurulur (Bölüm 5) ve sayfayı sunan sunucu `frame-ancestors` başlığı gönderir.

```sh
BYM_PARENT_ORIGINS="https://app.example.com" pnpm build
```

```html
<iframe id="editor" src="https://office.example.com/bymoffice.webview.html" style="width:100%;height:600px;border:0"></iframe>
<script>
  const frame = document.getElementById('editor');
  const editorOrigin = 'https://office.example.com';
  const send = (message) => frame.contentWindow.postMessage(JSON.stringify(message), editorOrigin);
  window.addEventListener('message', async (event) => {
    if (event.source !== frame.contentWindow || event.origin !== editorOrigin) return;
    const message = JSON.parse(event.data);
    if (message.type === 'ready') send({ type: 'load', id: '1', data: await fetchAsBase64('/files/rapor.xlsx'), name: 'rapor.xlsx' });
  });
  // `ready` dinleyici kurulmadan önce gönderilmiş olabilir; `hello` onu yeniden ister.
  frame.addEventListener('load', () => send({ type: 'hello' }));
</script>
```

Dosyayı üst sayfa kendisi getirir ve `data` ile verir; iframe içinde `url` kabul edilmez. Web uygulamalarında köprü yerine doğrudan `BymOffice.mount()` kullanmak çoğu zaman daha basittir (`docs/kullanim-kilavuzu.md`).

## 5. Güvenlik

- **İçerik güvenliği ilkesi.** Sayfa kendi `Content-Security-Policy` meta etiketini taşır:

  ```
  default-src 'none'; script-src 'sha256-…' 'sha256-…'; style-src 'unsafe-inline';
  img-src data: blob:; connect-src 'self' blob: data:; base-uri 'none'; form-action 'none'
  ```

  - `script-src`: yalnız sayfanın kendi iki satır içi betiği (motor ve başlatıcı), derlemede hesaplanan SHA-256 özetleriyle. `unsafe-inline` ve `unsafe-eval` yoktur; sonradan eklenen hiçbir betik çalışmaz.
  - `style-src 'unsafe-inline'`: motor kendi `<style>` öğesini çalışma zamanında ekler ve içeriği derleme sırasında köprü derlemesine açık değildir, bu yüzden özet kullanılamadı. Stil kod çalıştıramaz; `img-src`/`font-src` kapalı olduğu için stil üzerinden dışarıya istek de atılamaz. (`font-src` tanımlı değildir, yani `default-src 'none'` geçerlidir: adresle font yüklenemez. Sayfaya gömülü font ve `load.fonts` baytlardan `FontFace` ile eklenir; bu bir getirme değildir ve ilkeye takılmaz — Chromium, Firefox ve WebKit'te sınandı.)
  - `img-src data: blob:`: boş sayfa simgesi ve belgedeki görseller (sayfanın dosya baytlarından kendi ürettiği `blob:` adresleri) için.
  - `connect-src 'self' blob: data:`: **uzak adreslere çıkış yoktur.** Karar: köprü sayfası çoğunlukla çevrimdışı bir uygulama varlığıdır ve dosya `data` ile verilir. `url` yalnız sayfanın kendi kökeni (sanal ana makine ya da web sunucusuyla sunulduğunda, en üst pencere olarak), `blob:` ve `data:` adresleri için çalışır. Kod da aynı kuralı uygular: başka kökenden bir adres daha istek atılmadan `bad-message` ile reddedilir; ilke ikinci savunma hattıdır (ör. kendi kökenden başka kökene yönlendirme `load-failed` ile sonuçlanır). Uzak bir dosya gerekiyorsa uygulama onu kendisi indirip `data` ile verir; kimlik doğrulama da böylece uygulamada kalır.
- **iframe içinde kim konuşabilir.** İzinli üst kökenler yalnız sayfanın **içindeki** ayardan okunur: `<meta name="bymoffice-parent-origins" content="">` (boşlukla ayrılmış köken listesi). Boşsa iframe taşıması kapalıdır: sayfa üst pencereden hiçbir ileti kabul etmez ve ona hiçbir ileti göndermez. Liste derlemede `BYM_PARENT_ORIGINS` ortam değişkeniyle ya da sunan tarafın dosyadaki etiketi doldurmasıyla verilir. Sayfanın adresindeki hiçbir parametre bu listeyi etkilemez; çünkü adresi, sayfayı gömen (güvenilmeyen) taraf seçer.
  - Gelen ileti yalnız `event.source === window.parent` ise ve `event.origin` listedeyse kabul edilir; giden iletiler yalnız o kökeni hedef alır, `*` hiç kullanılmaz. Kökenler `new URL(x).origin` biçimine getirilerek karşılaştırılır (sondaki `/`, `:443`, büyük harf fark yaratmaz). `"null"` köken (sandbox'lı iframe, `file:`, `data:`) asla kabul edilmez ve hedef olmaz.
  - Listede birden çok köken varsa hedef, tarayıcının bildirdiği üst kökendir (`location.ancestorOrigins`); bunu bildirmeyen tarayıcıda (Firefox) sayfa, listedeki bir kökenden gelen ilk iletiyle (ör. `hello`) hedefini öğrenir ve `ready`'yi o zaman gönderir.
  - Kardeş iframe'lerden, sayfanın kendisinden ve `window.opener`'dan gelen iletiler yok sayılır. En üst pencere olarak yüklenen sayfa (WebView) hiçbir `window` iletisini komut saymaz.
  - iframe içinde `load.url` kabul edilmez (Bölüm 3).
- **`frame-ancestors` başlığı her durumda şarttır.** Sayfa web'de barındırılıyorsa sunucu `Content-Security-Policy: frame-ancestors https://app.example.com` (kimse gömmeyecekse `frame-ancestors 'none'`) **HTTP başlığını** göndermelidir. Bu yönerge meta etiketinde yok sayıldığı için sayfanın kendi ilkesi bunu sağlayamaz. İzin listesi yalnız mesajlaşmayı sınırlar; başlık olmadan herhangi bir site sayfayı gömüp üzerine sahte arayüz bindirebilir (clickjacking). İkisi birlikte kullanılır.
- **`BymOfficeBridge.receive`** sayfada betik çalıştırabilen herkese açıktır: WebView'de bu uygulamadır; web'de aynı kökendeki sayfalardır (aynı kökendeki bir üst sayfa, izin listesinden bağımsız olarak iframe'in bu işlevini doğrudan çağırabilir — aynı köken zaten tam erişim demektir).
- **Hiçbir ileti içeriği HTML olarak basılmaz.** Sayfanın yazıları `textContent` ile yazılır; motorun hata ayrıntısı sayfada hiç gösterilmez.
- **WebView ayarları.** WebView'i yalnız bu sayfa için kullanın; başka adrese gezinmeyi kapatın. Sayfa dosya sistemine, panoya (kullanıcının kopyala/yapıştır hareketleri dışında) ya da ağa erişmez.
- Makrolar çalıştırılmaz; dosyadaki dış bağlantılar izlenmez.

## 6. Sayfanın görünümü

- Editör bütün sayfayı kaplar; çentik ve ana ekran çizgisi için `env(safe-area-inset-*)` boşlukları bırakılır (`viewport-fit=cover`). Uygulama ayrıca güvenli alan boşluğu eklememelidir.
- `viewport` etiketi `maximum-scale=1` içerir. Gerekçe: hücre editörünün yazı boyutu hücrenin kendi boyutudur (çoğunlukla 16 pikselden küçük) ve iOS, 16 pikselden küçük bir giriş alanı odaklanınca sayfayı kendiliğinden yakınlaştırır; bu, her hücre düzenlemesinde ızgaranın kaymasına yol açar. Bedeli: sayfa iki parmakla yakınlaştırılamaz. **Erişilebilirlik notu:** bu, WCAG 2.x başarı ölçütü 1.4.4'ü (Metni Yeniden Boyutlandırma: metin %200'e kadar büyütülebilmeli) karşılamaz; az gören kullanıcılar ızgaradaki metni büyütemez. Bilinen ve geçici bir eksiktir: ızgaranın kendi yakınlaştırma özelliği editöre eklendiğinde `maximum-scale=1` kaldırılacaktır. O zamana dek uygulama, işletim sisteminin yakınlaştırma (büyüteç) özelliğine güvenir. `user-scalable=no` kullanılmadı.
- Arka plan açık ve koyu temada nötrdür (`prefers-color-scheme`); ızgaranın kendisi her zaman açık renklidir.
- Dosya yokken "Açık dosya yok" / "No file is open", yüklenirken ve (açık dosya yokken) yükleme başarısız olduğunda hata koduna göre kısa, genel bir yazı gösterilir ("Bu dosya türü desteklenmiyor", "Dosya bozuk, açılamadı", "Dosya yüklenemedi"). Motorun İngilizce ayrıntı iletisi kullanıcıya gösterilmez. Dil başlangıçta cihazın dilidir; başarılı bir `load`'un `locale` alanı ve `setLocale` onu değiştirir (Türkçe ya da İngilizce) ve `<html lang>` buna göre güncellenir. Açık bir dosya varken başarısız olan yükleme ekranda bir şey değiştirmez; uygulama `error` iletisine göre kullanıcıyı bilgilendirir.

## 7. Sınırlar

- **xlsx ve docx gösterilir ve düzenlenir; doc, RTF ve xls kendi biçimlerine yazılamadığı için salt okunur gösterilir**, `load.convert` ile docx/xlsx'e çevrilerek düzenlenebilir açılır (Bölüm 3). Parola korumalı dosyalar `unsupported` ile reddedilir.
- **Yazılamayan biçimler (dönüştürülmemiş doc, RTF, xls) için sözleşme:** `loaded` iletisi `readOnly: true` taşır. `setReadOnly` ile düzenleme açılamaz: istek `done` ile yanıtlanır ama yanıt `readOnly: true` taşır ve görünüm değişmez (düzenlemeler kaydedilemezdi). `save` özgün baytları döner; `changed` hiç gelmez.
- **Belge (docx; doc ve RTF) için sözleşme:** `loaded` iletisi `{ format, kind: 'document', sheets: [], activeSheet: 0 }` taşır; `load.sheet` yok sayılır. docx, `load.readOnly` `true` değilse düzenlenebilir açılır. Belge her değiştiğinde (yazma, silme, biçimlendirme, yapıştırma, geri al/yinele) tablodakiyle **aynı** `changed` iletisi gelir; belgenin sayfa (sheet) kavramı olmadığı için `sheets` hep boş dizidir, `canUndo`/`canRedo` belgenin geri alma geçmişini bildirir. `undo`, `redo`, `setReadOnly`, `setLocale` belgede de çalışır. `save`, süren bir giriş yöntemi (IME) kompozisyonunu önce belgeye işler ve değişiklikleri içeren docx baytlarını döner; hiç değiştirilmemiş bir docx özgün baytlarıyla döner. `selection` ve `sheet` iletileri belgede gelmez.
- **Belgede bağlantılar:** sayfa bir bağlantıyı kendisi açmaz, uygulamaya `{ type: 'link', url }` iletisi gönderir (yalnız http, https, mailto, tel). Salt okunur belgede bağlantıya basmak/dokunmak yeter; düzenlenen belgede düz tıklama imleci koyar, bağlantı Ctrl (macOS'ta Cmd) + tıklama ile ya da bağlantı odaktayken Enter ile etkinleşir. Dokunmatik ekranda düzenlenen belgede dokunma imleci koyar; bağlantıların dokunarak izlenmesi gerekiyorsa belge `setReadOnly(true)` ile salt okunur yapılır.
- **Base64 bellek maliyeti.** Dosya iki yönde de base64 dizgisi olarak taşınır: metin dosyadan üçte bir büyüktür ve yükleme/kaydetme sırasında dosyanın birkaç kopyası aynı anda bellekte durur (uygulamadaki baytlar, base64 metni, WebView'e geçirilen betik, sayfadaki dizgi, çözülmüş baytlar). Sayfanın kendi kodlama/çözme işlemi parça parça ve doğrusal sürede çalışır (20 MB için Node'da gidiş-dönüş ≈ 0,4 sn; tarayıcılarda 5 MB'lık dosya için yükleme ≈ 0,1–0,4 sn, kaydetme < 0,1 sn ölçüldü). Asıl maliyet platformun dizgiyi WebView'e geçirmesindedir ve gerçek cihazda ölçülmedi.
- **Gerçek cihazda sınanmadı.** Özellikle şunlar doğrulanmayı bekliyor: iOS'ta odak yakınlaştırmasının gerçekten engellendiği, güvenli alan boşlukları, ekran klavyesi açıldığında yerleşim, çok büyük `runJavaScript` betiklerinin davranışı, Android'in eski WebView sürümleri.
- `load` ve `save` için ileti boyutu sınırı yoktur (`url` ile getirmede 256 MB); bellek sınırı platformundur.
- **Yazdırma** (belge ve tablo) sayfanın `window.print()`'i üzerinden yapılır; WebView'lerde uygulamanın bu işlevi yerel yazdırmaya bağlaması gerekir (Bölüm 3, "Yazdırma"). Sayfalar görüntü olarak basılır (PDF'te metin seçilemez).
- **Kullanıcının kaydetmesi** (`load.save`): sonuç 60 saniye beklenir ve bu süre iletiyle değiştirilemez; kaydedilmemiş değişiklik varken sayfadan ayrılma uyarısı yoktur (uygulamanın işi); `saved` iletisi teslim edilemeden çok büyürse (64 M karakter sınırı) `save-failed` hatasına döner ve kayıt zaman aşımıyla başarısız sayılır. Gerçek WebView'de denenmedi.
- Aynı anda tek dosya açıktır.
