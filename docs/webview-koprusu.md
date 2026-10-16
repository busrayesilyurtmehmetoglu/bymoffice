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
- Durum bildiren iletilerden (`changed`, `selection`, `sheet`, `sheets`, `saveState`, `collabState`) her türden yalnız sonuncusu tutulur; birleştirilen `changed` iletisinin `sheets` alanı, yerine geçtiklerinin sayfalarını da içerir.
- Saklanan iletilerin toplamı 64 milyon karakteri aşarsa en eski `saved` iletileri, aynı `id` ile `error` / `save-failed` iletisine çevrilir (dosya baytları bırakılır, isteğin yanıtı kaybolmaz). Bırakılan ileti **kullanıcının kaydıysa** (`by: "user"`; `id`'si yoktur) `error` iletisi `id` yerine o kaydın **`saveId`**'sini taşır ve kayıt **hemen** başarısız sayılır: sayfa 60 saniyeyi beklemeden "Dosya kaydedilemedi" der, `saveState.saving` `false` olur ve o `saveId` için sonradan gelen `saveResult` `bad-message` alır.

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
| `load` | `data` (base64), `url` **ya da** `blank` (`"docx"` / `"xlsx"`); isteğe bağlı `id`, `name`, `locale`, `readOnly`, `sheet`, `convert`, `toolbar`, `save`, `author`, `trackChanges`, `commentsPanel`, `fonts`, `password`, `passwordPrompt` | Dosyayı açar (ya da boş bir dosyayla başlar) ve gösterir. Yanıt: `loaded` ya da `error` |
| `save` | isteğe bağlı `id` | Düzenlenmiş dosyanın baytlarını ister. Tabloda önce yazılmakta olan uygulanır: açık hücre düzenleyicisi, **not alanı** ve **sekme adı alanı** (`commitPending`). Uygulanamayan bir not ya da ad varsa alan açık kalır ve hiçbir şey kaydedilmez. **Açık kural kutusu** (veri doğrulama, koşullu biçim) kaydetmeyle uygulanmaz: içinde değişiklik yoksa kapatılır ve kaydedilir; değişiklik varsa kutu açık kalır ve hiçbir şey kaydedilmez (`save-failed`; kullanıcıya "Önce açık kural kutusunu onaylayın ya da kapatın" denir). Yanıt: `saved` ya da `error` (`save-failed`) |
| `setReadOnly` | `readOnly` (boolean); isteğe bağlı `id` | Salt okunur kipini açar/kapatır (tablo ve belge). Yazılamayan biçimlerde (dönüştürülmemiş doc, RTF, xls) düzenleme açılamaz: yanıt `readOnly: true` taşır |
| `setLocale` | `locale` (ör. `"tr"`, `"en"`); isteğe bağlı `id` | Sayı/tarih biçimlerinin, görünüm metinlerinin ve sayfanın kendi yazılarının dilini değiştirir. Tabloda içinde değişiklik yapılmış bir kural kutusu açıkken dil değişmez (kutu açık kalır): yanıt `error` / `locale-refused`; değişiklik yapılmamış kutu kapatılır ve komut yapılır |
| `undo`, `redo` | isteğe bağlı `id` | Geri al, yinele (tablo ve belge). Bir şey değişirse `changed` gelir. Tabloda not ya da sekme adı yazılırken gelen komut yazılanı **atmaz**: önce uygular ve o komutta başka bir şey yapmaz (sonraki `undo` onu geri alır); hücre düzenlenirken eskisi gibi düzenlemeyi bırakır. **Kural kutusu açıkken** komut hiçbir şey yapmaz (kutu kiplidir; yanıt yine `done`'dır, `changed` gelmez) |
| `print` | isteğe bağlı `id`, `dpi` (sayı, varsayılan 200), `sheets` (`"active"` varsayılan / `"all"`; yalnız tabloda) | Açık belgeyi ya da tabloyu yazdırır: sayfaları hazırlar ve sayfanın `window.print()`'ini çağırır. Yanıt: yazdırma penceresi kapanınca `done`, yapılamazsa `error` (aşağıda "Yazdırma") |
| `exportPdf` | isteğe bağlı `id`, `dpi` (sayı, 72–300, varsayılan 150), `image` (`"auto"` varsayılan / `"jpeg"` / `"lossless"`), `quality` (0'dan büyük, en çok 1; varsayılan 0,85), `sheets` (`"active"` varsayılan / `"all"`; yalnız tabloda), `text` (`"auto"` varsayılan / `"image"` / `"vector"`: sayfalar nasıl yazılır), `title`, `author` (dizge), `links`, `outline` (boolean, ikisi de varsayılan `true`: köprüler PDF bağlantısı, başlıklar / sayfa adları PDF anahattı olur) | Açık belgenin ya da tablonun **PDF dosyasını** üretir ve `pdf` iletisiyle gönderir; yapılamazsa `error` (aşağıda "PDF") |
| `printPreview` | isteğe bağlı `open` (boolean, varsayılan `true`; `false` kapatır), `sheets` (`"active"` varsayılan / `"all"`), `page` (1'den tam sayı: ilk gösterilecek yaprak), `id` | **Yalnız tabloda.** Tablonun **baskı önizlemesini** açar ya da kapatır: editörün üstünde, yazdırılacak yaprakları yazdırma ve PDF ile aynı dizilişte gösteren katman (yalnız görünüm; dosya değişmez, salt okunur tabloda da çalışır). Açıkken yeniden gönderilirse istenen sayfalara/yaprağa geçer. Kullanıcı katmanı kendi kapatır (×, Escape); bu bildirilmez. Yanıt: `done` (`open`: önizleme şimdi açık mı; açılamadıysa — ör. uygulanamayan bir hücre girişi — `false`) ya da `error` (`not-loaded`; açık dosya bir belgeyse `no-sheet`: belge zaten sayfalı gösterilir; geçersiz alan `bad-message`). `load.name` görünüme dosya adı olarak verilir: önizlemede, araç çubuğundan yazdırmada ve PDF düğmesinde `&F` ve PDF'in adı odur. Ayrıntı: [kullanım kılavuzu](kullanim-kilavuzu.md), "Baskı önizleme (tablo)" |
| `insertSheet` | isteğe bağlı `id`, `at` (0 tabanlı yer; verilmezse sona), `name` (verilmezse görünümün dilinde sıradaki varsayılan ad: `Sayfa2`, `Sheet2`) | Boş bir çalışma sayfası ekler ve gösterir (tek geri alma adımı). Yanıt: `done` (`index` ve `name` ile) ya da `error` |
| `renameSheet` | `index`, `name`; isteğe bağlı `id` | Sayfayı yeniden adlandırır; formüller yeni adı izler. Ad aynıysa hiçbir şey değişmez, yine `done` gelir |
| `deleteSheet` | `index`; isteğe bağlı `id` | Sayfayı siler (onay sorulmaz; `undo` geri getirir) |
| `moveSheet` | `from`, `to` (sayfanın taşındıktan sonraki 0 tabanlı dizini); isteğe bağlı `id` | Sayfayı taşır; formüller değişmez |
| `setSheetState` | `index`, `state` (`"visible"`, `"hidden"`, `"veryHidden"`); isteğe bağlı `id` | Sayfayı gizler ya da gösterir (görünür son sayfa gizlenemez) |
| `copySheet` | `index`; isteğe bağlı `id`, `at` (kopyanın yeri; verilmezse hemen arkası), `name` (verilmezse `Ad (2)`) | Sayfanın kopyasını ekler ve gösterir |
| `setTabColor` | `index`; isteğe bağlı `id`, `color` (`"#RRGGBB"`; verilmezse ya da `null` ise renk kaldırılır) | Sekme rengini değiştirir |
| `refreshPivots` | isteğe bağlı `id` | Tablodaki **bütün pivot tabloları** kaynaklarından yeniler (yenilenebilenleri; tek geri alma adımı). Önce yazılmakta olan uygulanır. Yanıt: `done` (`refreshed`: yenilenen tablo sayısı; `refused`: yenilenmeyen her tablo için `{ name, reason }`, `reason` görünümün dilinde ve kullanıcıya gösterilebilir — tablo adı dosyadan gelir, **düz metin** olarak gösterin). Hiçbir tablo yenilenemese de `done` gelir. Birlikte düzenlemede (`load.collab`) `error` / `pivot-refused` döner (ayrıntı: [kullanım kılavuzu](kullanim-kilavuzu.md), "Pivot tablolar") |
| `insertPivot` | isteğe bağlı `id`, `source` (`'A1:D20'`, `'Sayfa!A1:D20'` ya da tanımlı ad; yoksa seçim — tek hücre: çevresindeki blok), `cell` (`'F2'`: gösterilen sayfada raporun sol üst köşesi; yoksa yeni sayfa), `rows`, `cols`, `filters` (alan numaraları: kaynağın sütunları, 0'dan), `values` (`[{ field, subtotal?, numFmt? }]`), `panel` (`true`: alanlar paneli açılır) | Pivot tablo oluşturur (tek geri alma adımı; verilmeyen alanlar Excel'in önerisiyle). Önce yazılmakta olan uygulanır. Yanıt: `done` (`index`: tablonun `pivots()` sırası, `sheet`: sayfası). Reddedilirse `error` / `pivot-refused` (nedeni görünümün dilinde) |
| `setPivotFields` | `index` ve isteğe bağlı `id`, `rows`, `cols`, `filters` (`[alan]` ya da `[{ field, item? }]`), `values` (`[{ field, subtotal?, numFmt?, name? }]`), `rowGrandTotals`, `colGrandTotals`, `subtotals` (`[{ field, on }]`), `hidden` (`[{ field, items }]`), `sort` (`[{ field, order }]`) | Pivot tablonun alanlarını değiştirir ve tabloyu kaynağından yeniden kurar (tek geri alma adımı; verilmeyen kalır). Yanıt: `done` (`changed`: bir şey değişti mi). Reddedilirse `error` / `pivot-refused` |
| `deletePivot` | `index`, isteğe bağlı `id` | Pivot tabloyu hücreleriyle ve (başka tablonun kullanmadığı) önbelleğiyle **siler** (tek geri alma adımı; soru sorulmaz). Önce yazılmakta olan uygulanır. Yanıt: `done` (`changed`: tablo silindi mi). Reddedilirse (önbelleğini dilimleyici, pivot grafiği ya da başka bir parça anıyor; birlikte düzenleniyor…) `error` / `pivot-refused` (nedeni görünümün dilinde) |
| `pivotFields` | `index`, isteğe bağlı `id` | Yanıt: `done` (`fields`: `{ refusal?, fields: [{ index, name, items?: [{ text, hidden }], subtotal, sort }], rows, cols, filters, values, rowGrandTotals, colGrandTotals }`; alan ve öğe adları dosyadan gelir, **düz metin** olarak gösterin). Salt okunur görünümde de çalışır |
| `setTabStops` | `set`? (en çok 64 `{ pos, align, leader? }`; `pos` twip), `clear`? (en çok 64 konum), `clearAll`? (`true`), isteğe bağlı `id` | Belgede seçimin paragraflarının **sekme duraklarını** koyar ve kaldırır (tek geri alma adımı; her paragrafın stiline göre hesaplanır). Üçü de verilmezse paragrafların kendi durakları kalkar (stilinkiler geçerli olur). Geçersiz durak görünümce reddedilir ve kullanıcıya söylenir. İzleme açıkken izlenen değişikliktir. Yanıt: `done` (`changed`: belge değişti mi, `tabStops`: imlecin paragrafının durakları, `defaultTabStop`) |
| `setPageSetup` | **Tabloda:** isteğe bağlı `index` (sayfa dizini; verilmezse etkin sayfa), `paperSize`? (kâğıt kodu: 1 Letter, 3 Tabloid, 5 Legal, 7 Executive, 8 A3, 9 A4, 11 A5, 12 B4, 13 B5), `orientation`? (`"portrait"`/`"landscape"`), `margins`? (`{ left, right, top, bottom, header, footer }`, **inç**, verilenler), `scale`? (10..400) ya da `fitTo`? (`{ width, height }`, 0 = gerektiği kadar), `horizontalCentered`?, `verticalCentered`?, `gridLines`?, `headings`? (mantıksal), `printArea`? (`"A1:D20"`, birden çok alan `,` ya da `;` ile; `null` ya da `""` kaldırır), `titleRows`? (`"$1:$2"`), `titleCols`? (`"$A:$B"`; `null` kaldırır). **Belgede:** `paper`? (`"letter"`, `"legal"`, `"executive"`, `"a3"`, `"a4"`, `"a5"`, `"b5"`), `orientation`?, `margins`? (`{ top, bottom, left, right, header, footer, gutter }`, **twip**), `scope`? (`"section"` varsayılan: seçimin bölümleri; `"document"`: bütün bölümler). İsteğe bağlı `id` | **Sayfa yapısını** değiştirir (tek geri alma adımı); yalnız verilen alanlar değişir. Tabloda geçersiz değer görünümce reddedilir (`page-setup-refused`, `message` görünümün dilinde); belgede izleme açıkken ve birlikte düzenlemede görünüm reddeder ve kullanıcıya söyler (`changed: false`). Tabloda da birlikte düzenlemede reddedilir. Yanıt: `done` (`changed`, `pageSetup`: tabloda sayfanın şimdiki ayarı — `view.pageSetup()` biçiminde, kenar boşlukları inç; belgede imlecin bölümü — `view.sectionSetup`, twip) |
| `insertPageBreak`, `removePageBreak` | Yalnız tabloda: isteğe bağlı `index`, `row`?, `col`? (verilmezse etkin hücre: üstüne ve soluna), isteğe bağlı `id`. Gösterilmeyen sayfada `row` ya da `col` gerekir | Elle **sayfa sonu** ekler / kaldırır (tek geri alma adımı). Yanıt: `setPageSetup`'ınki gibi `done` (`changed: false`: orada zaten sonu vardı / yoktu) |
| `resetPageBreaks` | Yalnız tabloda: isteğe bağlı `index`, `id` | Sayfanın bütün elle sayfa sonlarını kaldırır. Yanıt: `done` |
| `insertSectionBreak` | Yalnız belgede: `kind`? (`"nextPage"` varsayılan, `"continuous"`, `"evenPage"`, `"oddPage"`), isteğe bağlı `id` | İmlece **bölüm sonu** koyar (gövdede, tablo dışında); izlemede ve birlikte düzenlemede reddedilir ve söylenir (`changed: false`). Yanıt: `done` (`changed`, `pageSetup`) |
| `setDefaultTabStop` | `twips` (1..31680), isteğe bağlı `id` | Belgenin **varsayılan sekme aralığını** değiştirir (tek geri alma adımı). İzleme açıkken ve birlikte düzenlemede reddedilir (söylenir; `changed: false`). Yanıt: `setTabStops`'unki gibi `done` |
| `acceptAllRevisions`, `rejectAllRevisions` | isteğe bağlı `id` | Belgedeki izlenen değişikliklerin tümünü kabul ya da reddeder (tek geri alma adımı). Çözülemeyen türler ve korumalı belge değişmez. Yanıt: `done` (`revisions`: kalan değişiklik sayısı, `trackChanges`: izleme açık mı) ya da `error` |
| `setTrackChanges` | `on` (boolean); isteğe bağlı `id` | Belgenin "değişiklikleri izle" ayarını açar ya da kapatır (geri alınabilir; `save` ile dosyaya yazılır). Kapatılamayan belgede (uygulanan `trackedChanges` koruması, ya da `load.trackChanges: true`) hiçbir şey değişmez; `load.trackChanges: true` ile izlenen belgede `on: true` da dosyanın ayarına dokunmaz (izleme zaten açıktır). Yanıt: `done` (`revisions`, `trackChanges`: **şimdi geçerli olan** durum) ya da `error` |
| `addComment` | `text` (dizge, en çok 32.767 karakter); isteğe bağlı `anchor` ve `focus` (`{ story?, path, offset }`: yorumun ilgili olduğu aralık; yalnız `anchor` verilirse o noktadaki sözcük; ikisi de yoksa **seçim**), `id` | Belgeye yorum ekler (tek geri alma adımı; izleme açıkken de izlenen değişiklik değildir). Yazar `load.author`'dır. Yanıt: `done` (`commentId`: yeni yorumun kimliği) ya da `error` |
| `editComment` | `commentId` (sayı), `text`; isteğe bağlı `id` | Yorumun metnini değiştirir (düz metin; `\n` paragraf ayırır) |
| `replyComment` | `commentId`, `text`; isteğe bağlı `id` | Yorumun zincirine yanıt ekler. Yanıt: `done` (`commentId`: yanıtın kimliği) |
| `resolveComment` | `commentId`; isteğe bağlı `done` (boolean, varsayılan `true`), `id` | Yorumun zincirini çözülmüş (ya da `done: false` ile yeniden açık) işaretler |
| `removeComment` | `commentId`; isteğe bağlı `id` | Yorumu metindeki işaretleriyle siler; zincirin ilk yorumu yanıtlarını da götürür |
| `commentsPanel` | isteğe bağlı `open` (boolean; verilmezse panel açıksa kapanır, kapalıysa açılır), `id` | Sayfaların yanındaki **yorum panelini** açar ya da kapatır (yalnız görünüm; belge değişmez, salt okunur belgede de çalışır). Yanıt: `done` (`open`: panel şimdi açık mı) ya da `error` (`not-loaded`, `no-document`) |
| `close` | isteğe bağlı `id` | Dosyayı kapatır, süren yüklemeyi iptal eder, boş duruma döner. Dosya yokken de hata vermez |
| `saveResult` | `saveId` (o `saved` iletisinin `saveId`'si), `ok` (boolean); isteğe bağlı `message` (kısa metin), `conflict` (boolean; `ok: false` ile: çakışma — sayfa kullanıcıya üzerine yazmayı sorar, aşağıda "Çakışma"), `id` | **Kullanıcının** başlattığı kaydın (`by: "user"` taşıyan `saved` iletisi; aşağıda "Kullanıcının kaydetmesi") sonucunu bildirir: `ok: true` ise sayfa "Kaydedildi" der, `false` ise "Dosya kaydedilemedi". Yalnız **o anda beklenen** kaydın `saveId`'siyle geldiğinde uygulanır; değilse `bad-message`. Sıra beklemez. `id` verilmişse `done` ile onaylanır |
| `password` | `value` (parola dizgesi, en çok 1.024 karakter; `null`: vazgeç); isteğe bağlı `id` | `passwordRequired` iletisinin yanıtı (aşağıda "Parola korumalı dosyalar"). Sırasını beklemez: süren yükleme onu bekliyordur. Yanıt: `done`; parola sorulmuyorken gelirse (yükleme yok, ya da bir önceki parola hâlâ denetleniyor) `error` / `password-not-expected` — ileti sıraya alınmaz. **Değeri hiçbir iletide geri gönderilmez** |
| `hello` | — | Sayfa `ready`'yi yeniden gönderir |

**Görsel komutu yok:** tabloya görsel eklemek, taşımak, silmek, döndürmek, kırpmak, alternatif metnini değiştirmek ya da görseli kopyalayıp yapıştırmak için köprüde ileti yoktur (`insertPicture` gibi bir ileti bilinmeyen tür olarak `error` / `bad-message` ile reddedilir). Köprü sayfasındaki araç çubuğunun **Görsel ekle** düğmesi (sayfanın kendi dosya penceresi), yapıştırma, ızgaraya sürükle-bırak, görselin menüsü (kes, kopyala, yapıştır, döndür, kırp, alternatif metin, sil) ve döndürme/kırpma tutamaçları çalışır; görsel kopyası köprü sayfasının içinde tutulur (sistem panosuna yalnız tarayıcı motoru izin verirse PNG olarak gider); WebView'de dosya penceresinin açılması uygulamanın dosya seçme iznine bağlıdır (ayrıntı: [kullanım kılavuzu](kullanim-kilavuzu.md), "Görseller (xlsx)").

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

PDF (`exportPdf`):

- Belgede ve tabloda çalışır, salt okunur ve dönüştürülmemiş (doc, RTF, xls) dosyada da; `DocumentView.exportPdf()` / `WorkbookView.exportPdf()` ile aynı koddur: yazısının tamamı font dosyası elde olan fontlarla (paketle gelen font; köprüde `load.fonts` ile verilen fontlar) yazılmış sayfa **vektör** yazılır (gerçek metin, çizim, ayrı görseller), öbür sayfalar yazdırmadakiyle aynı çizimin görüntüsüdür ve üzerlerinde aranabilir/kopyalanabilir görünmez metin katmanı vardır. `text: "image"` her sayfayı görüntü yazar; `text: "vector"` vektör yazılamayan sayfa varsa dosya üretmez ve `error` / `pdf-unsupported` döner (`message` sayfa numarasını ve nedeni söyler) (ayrıntı, sınırlar ve dosya boyutları: [kullanım kılavuzu](kullanim-kilavuzu.md), "PDF olarak kaydetme"). `window.print()` gerektirmez; WebView'de PDF almanın yolu budur (sayfadaki indirme düğmesi WebView'lerde çoğunlukla bir şey kaydetmez).
- Yanıt `{ "type": "pdf", "id": …, "data": "<base64>", "name": "Rapor.pdf", "pages": 12, "vector": 10 }` (`vector`: vektör yazılan sayfa sayısı; kalanı görüntüdür) — `saved` iletisindeki gibi **tek iletide** base64; parçalama yoktur. `id` verilmese de `pdf` gönderilir (istenen şey dosyanın kendisidir). **Boyut sınırı:** dosya ham 64 MB'ı aşarsa `pdf` yerine `error` / `pdf-too-large` gelir (ileti tek bir dizgedir: sayfa kurar, kanal kopyalar, uygulama çözer; base64 hali dosyanın 4/3'üdür). Sayfa dosyayı tek parça halinde tutmaz, base64'ü yazıldığı parçalardan üretir. Dosya hazırlandıktan sonra oluşan her hata da (`pdf-failed`) yanıtlanır: uygulama `pdf` ya da `error`'dan birini mutlaka alır. Varsayılan `image: "auto"` sayfa sayfa küçük olanı seçer; 10 sayfalık metin belgesi 150 dpi'de ≈ 2,7 MB, fotoğraflı 10 sayfa ≈ 2,4 MB.
- **Tabloda** `sheets: "active"` (varsayılan) gösterilen sayfayı, `"all"` görünür bütün sayfaları yazar; `load` iletisindeki `name` üstbilgi/altbilgideki `&F` olur. Yazılmakta olan hücre, not ve sekme adı önce uygulanır.
- `exportPdf` **sırayı tutmaz**: sayfalar çizilirken sonraki iletiler işlenir; yanıt dosya hazır olunca gelir. Aynı anda tek dosya üretilir (ikinci `exportPdf` sırasını bekler). Sayfalar çizilirken dosya kapatılır ya da değiştirilirse (`close`, yeni `load`) iş `error` / `pdf-failed` ile biter. Köprüde iptal komutu yoktur.
- Geçersiz `dpi`, `image`, `quality`, `sheets`, `text`, `title`, `author`, `links`, `outline` → `bad-message`; açık dosya yoksa `not-loaded`.
- **Bağlantılar ve anahat:** dosyanın köprüleri PDF bağlantısı (yalnız `http`/`https`/`mailto` adresleri ve dosya içi hedefler; `javascript:`, `file:`, `data:`, `tel:` ve göreli adresler yazılmaz), belgenin başlıkları ya da tablonun sayfa adları PDF anahattı olur; `links: false` / `outline: false` kapatır. Yanıttaki `links` nesnesi ne yazıldığını ve neyin neden yazılmadığını sayar (`refused`: güvenli olmayan adres, `unresolved`: hedefi bulunamayan iç bağlantı, `dropped`: sınırı aşan).
- Metin katmanına yalnız sayfada **görünen** metin girer (komşu hücrenin kestiği metnin görünmeyen kısmı PDF'te yoktur); ayrıntı kullanım kılavuzunda.

Kullanıcının kaydetmesi (`load.save: true`, `saved` + `by: "user"`, `saveResult`):

- Kullanıcı Kaydet düğmesine ya da Ctrl/Cmd+S'e basınca sayfa, uygulama `save` göndermiş gibi dosyayı yazar ve **`saved`** iletisini gönderir: `{ "type": "saved", "by": "user", "saveId": "u1", "data": "<base64>", "format": "docx" | "xlsx", "name"?, "convertedFrom"? }`. Bu iletide **`id` yoktur** (hiçbir isteğin yanıtı değildir), `format` her zaman vardır ve **`saveId`** o kaydın kimliğidir (sayfa açık kaldıkça hiç yinelenmeyen bir dizge; biçimine güvenmeyin, aynen geri gönderin); `by: "user"` onu uygulamanın kendi `save` isteğinin yanıtından ayırır (o yanıt `by` taşımaz ve eskisi gibidir).
- Uygulama dosyayı saklar (diske yazar, sunucuya yollar…) ve sonucu **aynı `saveId` ile** bildirir: `{ "type": "saveResult", "saveId": "u1", "ok": true }` ya da `{ "type": "saveResult", "saveId": "u1", "ok": false, "message": "Sunucuya ulaşılamadı" }`. Sonuç gelene dek düğme "meşgul" görünür (`aria-busy`). `ok: true` → "Kaydedildi", düğme solar. `ok: false` → "Dosya kaydedilemedi (…)"; `message`'ın ilk 200 karakteri tek satır olarak, **metin olarak** gösterilir (HTML olarak basılmaz) ve değişiklikler kaydedilmemiş sayılmaya devam eder.
- **60 saniye** içinde `saveResult` gelmezse kayıt başarısız sayılır ("Dosya kaydedilemedi: zamanında yanıt gelmedi"); sonradan gelen sonuç hiçbir şeyi değiştirmez. **`saveResult` yalnız kendi kaydına uygulanır:** `saveId`'si o anda beklenen kaydınkiyle aynı değilse (zaman aşımına uğramış kaydın geç yanıtı, iki kez gönderilen yanıt, kapatılmış dosyanın kaydı, bekleyen kayıt yokken gelen yanıt) hiçbir şey değişmez ve uygulamaya `error` / `bad-message` ("No save with this "saveId" is waiting for a result") döner — böylece geç gelen bir "tamam", bir sonraki kaydı (ya da başka bir dosyanın kaydını) yanlışlıkla "kaydedildi" yapamaz. `saveId` yoksa ya da metin değilse, `ok` boolean değilse, `message` metin değilse ya da `conflict` boolean değilse de `bad-message` döner (kayıt sonucunu beklemeyi sürdürür).
- **Çakışma (`conflict: true`):** HTTP'yi uygulama yaptığı için dosyanın sunucuda (ya da diskte) o arada değiştiğini uygulama bilir. Üzerine yazmak yerine `{ "type": "saveResult", "saveId": "u1", "ok": false, "conflict": true, "message"?: "…" }` gönderirse sayfa, editörün içinde bir **onay** gösterir (`role="alertdialog"`; dosyanın dilinde): "Dosya, siz açtıktan sonra sunucuda değişmiş (başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir). Kaydederseniz oradaki değişikliklerin üzerine yazılır." — **Üzerine yaz** / **Vazgeç** (odak Vazgeç'te; Escape = Vazgeç). `message` verilmişse onayın altında tek satır, en çok 200 karakter, **düz metin** olarak gösterilir. Görünüm ayrıca "Dosya kaydedilmedi: sunucudaki dosya sizin açtığınızdan farklı…" bildirimini verir ve `saveState` `dirty: true, saving: false` olur.
  - **Üzerine yaz** → sayfa yeni bir kullanıcı kaydı gönderir: `{ "type": "saved", "by": "user", "saveId": "u2", "overwrite": true, "data": …, "format": … }` (yeni `saveId`; dosyanın **o anki** hali). Uygulama bunu **koşulsuz** yazar ve her kayıt gibi `saveResult` ile yanıtlar. `overwrite: true` taşıyan bir kayda yine `conflict: true` yanıtı verilirse sayfa yeniden sormaz; yalnız "kaydedilmedi" der.
  - **Vazgeç** → hiçbir ileti gitmez; dosya kaydedilmemiş kalır (son `saveState` bunu zaten söyler). Sonraki Kaydet, `overwrite` taşımayan olağan bir `saved` gönderir; uygulama yine `conflict` derse onay yeniden gösterilir.
  - Onay açıkken Kaydet düğmesi ve Ctrl/Cmd+S yeni ileti göndermez, ikinci onay açmaz; editör klavye ve fare girdisi almaz. `setReadOnly` (`true`), `close`, yeni bir `load` ve sayfanın yok edilmesi onayı yanıtsız kapatır (vazgeçilmiş sayılır).
  - **Dosya salt okunurken sorulmaz:** `setReadOnly` ile salt okunura alınmış (ya da `readOnly: true` ile yüklenmiş) dosyanın kaydına `conflict: true` yanıtı verilirse onay açılmaz; yalnız "Dosya kaydedilmedi: sunucudaki dosya sizin açtığınızdan farklı…" bildirimi gösterilir ve `saveState` `dirty: true, saving: false` olur.
  - **Yanlışlıkla üzerine yazmaya karşı:** onayda düğmeler arasında yalnız Tab ile gezilir (ok tuşları hiçbir şey yapmaz); "Üzerine yaz" onay açıldıktan sonraki 400 ms boyunca yanıt vermez (`aria-disabled="true"`; Kaydet'e çift dokunmanın ikinci basışı üzerine yazdırmaz), basılı tutulan tuş yanıt sayılmaz; Vazgeç ve Escape hemen çalışır. Uygulamanın `conflict` yanıtını anında vermesi bu yüzden güvenlidir.
  - **Üzerine yazma izni tek kayıt içindir:** `overwrite: true` yalnız kullanıcının "Üzerine yaz" yanıtının doğurduğu `saved` iletisinde gelir; uygulama onu `ok: true` ile yanıtladıktan sonraki kayıtlar yine `overwrite`'sız gelir.
  - Onay, gösterilen görünümün ait olduğu editörün kök öğesinin üstünü kaplar (görünüm `element` alanını veriyorsa; köprü sayfasının görünümleri verir). `createBridge()`'i kendi görünümünüzle kullanıyorsanız `BridgeHost.askOverwrite(question, locale, element)` üçüncü bağımsız değişken olarak bu öğeyi alır (görünüm vermiyorsa `undefined`); `installBridge()` o durumda kapsayıcının kendisini kaplar.
  - `conflict` alanı yoksa (ya da `false` ise) davranış eskisi gibidir: "Dosya kaydedilemedi (…)". `ok: true` ile gelen `conflict` yok sayılır. `load` iletisinde `toolbar: false` olsa da onay çalışır.
- Aynı anda tek kullanıcı kaydı bekler: sonuç beklenirken yeniden basılırsa ve arada bir şey değişmediyse ikinci ileti gönderilmez; değiştiyse sonuç geldikten sonra bir kez daha `saved` gönderilir. Sonuç beklenirken yazılan şey "kaydedildi" sayılmaz.
- Kaydetmeden önce yazılmakta olan uygulanır (`save` ile aynı kural); uygulanamıyorsa hiçbir şey gönderilmez ve neden görünümde söylenir. Kaydedilecek değişiklik yoksa düğme soluktur ve tuş hiçbir şey göndermez.
- Uygulamanın kendi `save` isteği her an çalışır ve kullanıcının "kaydedilmemiş değişiklik" durumunu **değiştirmez** (sayfa, baytların saklandığını yalnız `saveResult` ile öğrenir). `close` ya da yeni bir `load`, bekleyen kullanıcı kaydını bırakır: görünümde hiçbir bildirim çıkmaz, o kaydın `saveResult`'ı artık `bad-message` alır ve uygulamaya durumun bittiği `saveState` (`dirty: false`, `saving: false`) ile söylenir.
- Sayfa, kaydedilmemiş değişiklik varken kapatılmaya karşı kendisi uyarmaz; bunu uygulama **`saveState`** iletisinden izler (son gelen `dirty` / `saving`; `changed` ve `saved` iletilerini sayarak çıkarılamaz — aşağıda `saveState`).

```json
{ "type": "load", "id": "1", "data": "UEsDBBQ…", "name": "rapor.xlsx", "toolbar": true, "save": true }
{ "type": "saved", "by": "user", "saveId": "u1", "data": "UEsDBBQ…", "format": "xlsx", "name": "rapor.xlsx" }
{ "type": "saveResult", "saveId": "u1", "ok": true }
```

Kendiliğinden kaydetme (`load.save: { "auto": … }`, `saved` + `by: "auto"`):

- `save` alanı `true` yerine bir nesne olabilir: `{ "auto": true }` ya da `{ "auto": { "delayMs": 2000, "maxWaitMs": 30000 } }` (`delayMs` 250–600.000, `maxWaitMs` ≥ `delayMs`; `true` = 2 sn / 30 sn). Nesne verilmesi `save: true` demektir (Kaydet düğmesi ve Ctrl/Cmd+S de vardır); `{}` ve `{ "auto": false }` yalnız düğmedir. Sınır dışı değer `bad-message` döner ve dosya açılmaz. `loaded` iletisi `save: true` yanında **`autoSave: true`** taşır.
- Sayfa, son değişiklikten `delayMs` sonra (durmadan yazılıyorsa ilk değişiklikten en geç `maxWaitMs` sonra) dosyayı **kendiliğinden** verir: `{ "type": "saved", "by": "auto", "saveId": "a1", "data": …, "format": …, "name"? , "convertedFrom"? }`. `by` dışında kullanıcı kaydıyla aynıdır; `saveId` aynı sayaçtan gelir (`a1`, `u2`, `a3`… — biçimine güvenmeyin). Uygulama **aynı sözleşmeyle** `saveResult` gönderir; yanıt gelene dek düğme meşguldür ve `saveState.saving` `true`'dur. Aynı anda tek kayıt (kullanıcının ya da kendiliğinden) yanıt bekler.
- **Yazılmakta olan uygulanmaz ve bölünmez:** açık hücre düzenleyicisi, not, sekme adı, kural kutusu, IME bileşimi, sürükleme ya da çakışma onayı varken `by: "auto"` iletisi gitmez; bir `delayMs` sonra yeniden bakılır. (Kullanıcının kaydı ve uygulamanın `save` isteği eskisi gibi önce yazılmakta olanı uygular.)
- **`conflict: true` yanıtı:** onay kendiliğinden **açılmaz**; görünüm çakışmayı söyler, `saveState` `{ dirty: true, saving: false }` olur ve kendiliğinden kayıt **durur** (yeni değişiklikler de başlatmaz). Kullanıcı Kaydet'e basınca sayfa `by: "user"` taşıyan (ve `overwrite` taşımayan) yeni bir `saved` gönderir; uygulama ona da `conflict: true` derse "Üzerine yaz / Vazgeç" onayı o zaman açılır.
- **`ok: false` (çakışmasız) yanıtı** kendiliğinden kaydı durdurur; kullanıcının bir sonraki kaydına dek yeniden denenmez. **Yanıt hiç gelmezse** (60 sn) kayıt başarısız sayılır ve 5 sn, 10 sn, 20 sn… (en çok 5 dk) sonra yeniden denenir; aynı hata bir kez gösterilir. `ok: true` her şeyi sıfırlar.
- Salt okunur dosyada (`setReadOnly`) hiçbir şey verilmez; yeniden düzenlenebilir yapılınca süre baştan başlar. Sayfa gizlenince (`visibilitychange` → `hidden`) bekleyen değişiklik hemen verilir. `close` ve yeni `load` zamanlayıcıyı kaldırır.
- **Çakışma denetimi uygulamadadır.** `mount({ autoSave })`'daki "yalnız koşullu gidebiliyorsa gider" kuralı (ETag / `If-Match`) köprüde **uygulanmaz**: HTTP'yi (ya da diske yazmayı) uygulama yapar, sayfa hiçbir sürüm bilgisi tutmaz. Kendiliğinden gelen dosyayı (`by: "auto"`) saklamadan önce elinizdeki sürümle karşılaştırın ve değiştiyse **yazmadan** `saveResult` ile `conflict: true` deyin; koşulsuz yazarsanız başkasının değişikliği kimseye sorulmadan ezilir. `auto` nesnesindeki `unconditional` alanı kabul edilir ve etkisizdir.
- **"Kaydedildi" her seferinde gösterilmez:** `by: "auto"` kaydına `ok: true` gelince görünümde bildirim çıkmaz (düğme ve `saveState` güncellenir); bir sorun gösterildikten sonraki ilk başarılı kayıt bir kez "Kaydedildi" der. Kullanıcının kaydı her seferinde der.
- Yanıtlanan kaydın baytları hemen bırakılır: `saveResult` gelince 60 saniyelik bekleme zamanlayıcısı iptal edilir (gömen kod `createBridge`'i kendi `timer`'ıyla kuruyorsa, `timer` iptal işlevi döndürebilir; döndürmese de sayfa baytlara başvuru tutmaz).
- Çalışırken açıp kapatan bir ileti yoktur: ayar `load` ile verilir.

```json
{ "type": "load", "id": "1", "data": "UEsDBBQ…", "name": "rapor.xlsx", "toolbar": true, "save": { "auto": true } }
{ "type": "loaded", "id": "1", "format": "xlsx", "sheets": [ … ], "activeSheet": 0, "save": true, "autoSave": true }
{ "type": "saved", "by": "auto", "saveId": "a1", "data": "UEsDBBQ…", "format": "xlsx", "name": "rapor.xlsx" }
{ "type": "saveResult", "saveId": "a1", "ok": true }
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
- `commentsPanel` (`"auto"` / `"open"` / `"closed"`, varsayılan `"closed"`; yalnız belgede): sayfaların yanındaki yorum paneli baştan açık mı (`"auto"`: belgede yorum varsa ve sayfa en az 720 piksel genişse; kap henüz genişliksizse — gizli ya da yerleşmemiş — karar ilk geçerli boyutta bir kez verilir). Kullanıcı paneli görünüm çubuğundaki **Yorumlar** düğmesiyle ya da Ctrl+Alt+Shift+M ile de açıp kapatır; uygulama `commentsPanel` komutuyla.
- **Korumalı belge:** `w:documentProtection` uygulanıyorsa ve türü `readOnly` ya da `forms` ise belge `readOnly` ne olursa olsun **salt okunur** gösterilir; `loaded` iletisi `readOnly: true` ve `protection` taşır, `setReadOnly` ile açılamaz (yanıt `done` + `readOnly: true`), `save` `read-only` hatası verir. Türü `comments` ise **metin düzenlenemez ama yorumlar serbesttir**: `loaded` iletisi `protection: "comments"` ve `commentsOnly: true` taşır, `readOnly` taşımaz; yorum komutları (aşağıda), `undo`/`redo` (yalnız yorum adımları) ve `save` çalışır; `acceptAllRevisions`, `rejectAllRevisions` ve `setTrackChanges` hiçbir şey yapmadan **`read-only` hatasıyla** yanıtlanır (metin ve izlenen değişiklikleri kilitlidir), kullanıcı balondan ve (`toolbar: true` ile) araç çubuğundan yorum ekleyip değiştirebilir; yazdığı metin belgeye girmez. `setReadOnly(true)` yorumları da kapatır. Ayrıntı: kullanım kılavuzunda "Belge koruması ve değişiklik izleme" ve "Belge yorumları".
- **Yorumlar (yalnız belgede):** belgenin yorumları `loaded.comments` ile, sonraki her değişiklikte `comments` iletisiyle bildirilir (yanıtlar `parent`, çözülmüş zincir `done: true`). Komutlar: `addComment`, `editComment`, `replyComment`, `resolveComment`, `removeComment`. Yorumlar sayfada vurgu, kenar işareti ve balonla, istenirse yanda **yorum panelinde** (`load.commentsPanel`, `commentsPanel` komutu) görünür; salt okunur belgede de gösterilir (komutlar `read-only` döner).
- `toolbar` (boolean, varsayılan `false`): `true` ise belge, sayfaların üstünde **biçimlendirme araç çubuğuyla** gösterilir (`DocumentEditor`; belge salt okunurken araç çubuğu gizlenir). `false` iken belge yalnız sayfalar olarak gösterilir (altta sayfa numarası ve yakınlaştırma çubuğu vardır); klavyeyle yazma ve kısayollar iki kipte de çalışır. **Tablo, köprü sayfasında her iki kipte de editör kabuğuyla (araç çubuğu, formül çubuğu, durum çubuğu) gösterilir**; `toolbar` tabloda bir şey değiştirmez. Komutlar ve iletiler iki kipte de aynıdır.
  - **Tabloda:** alan verilmezse ya da `true` ise tablo editör kabuğuyla (araç çubuğu, formül çubuğu, durum çubuğu) gösterilir. Açıkça `toolbar: false` verilirse **araç çubuğu ve formül çubuğu gösterilmez**; ızgara, sayfa sekmeleri ve durum çubuğu kalır (salt okunur bir görüntüleyici için). Bu durumda `save: true` ile gelen Kaydet düğmesi de görünmez; Ctrl/Cmd+S çalışır.
- `save` (boolean ya da nesne, varsayılan `false`): `true` ise **kullanıcı sayfanın içinden kaydedebilir**: araç çubuğunun başında Kaydet düğmesi görünür (belgede `toolbar: true` de gerekir; tablo hep kabukla gösterilir) ve Ctrl/Cmd+S aynı işi yapar (tarayıcının "sayfayı kaydet" penceresi açılmaz). Ayrıntı aşağıda, "Kullanıcının kaydetmesi". Yazılamayan dosyada (dönüştürülmemiş doc, RTF, xls; düzenlemeye kapalı korumalı belge) etkisizdir: düğme yoktur, tuş tarayıcıya kalır ve `loaded` iletisi `save` taşımaz. Nesne verilirse (`{ "auto": true }` ya da `{ "auto": { "delayMs", "maxWaitMs" } }`) aynı şey geçerlidir ve sayfa ayrıca **kendiliğinden kaydeder** (aşağıda, "Kendiliğinden kaydetme"). Boolean ya da nesne olmayan değer `bad-message` döner.
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
| `loaded` | `id`?, `format`, `kind`? (`"document"`, yalnız belgede), `sheets: [{ name, state, kind }]` (belgede `[]`), `activeSheet`, `readOnly`? (`true`: dosya yazılamayan bir biçimde ya da `readOnly`/`forms` korumalı, hep salt okunur; `comments` korumasında **gelmez**), `convertedFrom`?, `warnings`? (`string[]`), yalnız belgede: `protection`? (`"readOnly"` / `"comments"` / `"forms"` / `"trackedChanges"`: belgenin uyguladığı koruma), `protectionExceptions`? (`true`: belgede korumadan muaf aralıklar var; uygulanmazlar), `trackChanges`? (`true`: düzenlemeler izleniyor), `revisions`? (belgedeki izlenen değişiklik sayısı; yalnız sıfırdan büyükse), `commentsOnly`? (`true`: belge `comments` korumalı — metin düzenlenemez, yorum komutları ve kaydetme çalışır), `comments`? (belgenin yorumları, `comments` iletisindeki biçimle; yalnız yorum varsa), `commentsMore`? (`true`: 2.000'den çok yorum var, liste ilk 2.000'i taşır), `save`? (`true`: `load.save` istendi ve dosya yazılabilir; kullanıcı sayfadan kaydedebilir) | `load` başarılı |
| `passwordRequired` | `id`? (o `load`'un `id`'si), `wrong` (boolean: denenen son parola yanlıştı), `attempts` (başarısız deneme sayısı), `format` (`"ooxml"`: docx ya da xlsx; `"doc"`; `"xls"`) | Yüklenen dosya parola korumalı ve parola verilmemiş ya da yanlış. Yükleme `password` iletisini (ya da sayfadaki kutuya yazılanı) bekler; parola yanlış oldukça yinelenir |
| `changed` | `sheets: number[]` (değişen sayfalar; belgede hep `[]`), `canUndo`, `canRedo` | Çalışma kitabı ya da belge her değiştiğinde (geri al/yinele dahil) |
| `saveState` | `dirty` (boolean: dosyanın şimdiki hali, uygulamanın en son sakladığını bildirdiği halden farklı), `saving` (boolean: bir kullanıcı kaydı yanıt bekliyor) | Yalnız `load.save: true` ile açılan dosyada, durum her değiştiğinde. **Kaydedilmemiş değişiklik var mı sorusunun tek doğru kaynağıdır**: uygulama bunu `changed` iletilerini sayarak ya da `canUndo`'ya bakarak çıkarmamalıdır (kayıt yazılırken yapılan değişiklik hiçbir `saved` iletisinde yoktur; kırpılmış bir geri alma geçmişi sonuna dek geri alınsa da dosya açıldığı hali olmayabilir). **Her `loaded`'dan sonra başlangıç durumu gelir:** `save: true` ile açılan (ve yazılabilen) her dosyada `loaded`'ın hemen ardından `{ dirty: false, saving: false }` gönderilir. **`close` ve `save`'siz `load` durumu sıfırlar:** önceki dosya `save` ile açılmışsa, `close` (onayından önce), `save`'siz ya da yazılamayan bir dosyanın `loaded`'ı ve eski dosyayı kapatıp gösterilemeyen dosyanın `error` / `corrupt` iletisi ardından bir kez `{ dirty: false, saving: false }` gelir; uygulama kapanan dosyanın "kaydedilmemiş" ya da "kaydediliyor" durumuna takılı kalmaz. `save` ile hiç dosya açılmadıysa bu ileti hiç gönderilmez. Kanal yokken yalnız sonuncusu saklanır |
| `saved` | `id`?, `data` (base64), `name`?, `format`? ve `convertedFrom`? (yalnız dönüştürülerek açılmış dosyada, ikisi birlikte) | `save` başarılı |
| `saved` (kullanıcıdan) | `by: "user"`, `saveId` (dizge; `saveResult`'ta geri gönderilir), `data` (base64), `format` (`"docx"` / `"xlsx"`, her zaman), `name`?, `convertedFrom`?, `overwrite`? (`true`: kullanıcı bir çakışma onayında "Üzerine yaz" dedi — koşulsuz yazın); **`id` yok** | Kullanıcı sayfada Kaydet'e ya da Ctrl/Cmd+S'e bastı (`load.save: true`). Uygulama `saveResult` ile yanıtlar |
| `saved` (kendiliğinden) | `by: "auto"`, `saveId`, `data`, `format`, `name`?, `convertedFrom`?; **`id` ve `overwrite` yok** | Sayfa kendiliğinden kaydetti (`load.save: { "auto": … }`). Uygulama aynı biçimde `saveResult` ile yanıtlar; `conflict: true` derse onay açılmaz, kendiliğinden kayıt durur |
| `pdf` | `id`?, `data` (base64: PDF dosyası), `name` (yüklenen dosyanın adından: `Rapor.docx` → `Rapor.pdf`; ad yoksa `document.pdf` / `workbook.pdf`), `pages` (sayfa sayısı), `vector` (bunlardan vektör yazılanların sayısı), `links` (`{ external, internal, refused, unresolved, dropped, rotated, outline, outlineSkipped, unread }`: yazılan dış/iç bağlantı, güvenli olmadığı için yazılmayan, hedefi bulunamayan, sınırı aşan, döndürülmüş metindeki bağlantı; yazılan ve yazılmayan anahat girdisi; belgede aranmayan blok sayıları; kullanım kılavuzu, "PDF olarak kaydetme") | `exportPdf` başarılı |
| `selection` | `sheet`, `range` (`"A1:B2"`, tek hücrede `"A1"`), `active` (`"A1"`) | Seçim değişince; **kare başına en çok bir kez**, yalnız o karedeki son seçim, bir öncekiyle aynıysa gönderilmez. Bekleyen seçim iletisi sayfa listesi değişirse (`sheets`) gösterilen sayfanın **yeni** diziniyle ve o sayfanın seçimiyle gönderilir |
| `revisions` | `count` (belgedeki izlenen değişiklik sayısı; en çok 2.000 sayılır), `trackChanges` (izleme açık mı) | Yalnız belgede: sayı ya da izleme durumu son bildirilenden farklı olunca (yazma, kabul/ret, geri al, `setTrackChanges`, kullanıcının araç çubuğundan yaptığı). Aynı değişikliğin `changed` iletisinden **önce** gelir. Başlangıç değeri `loaded` iletisindedir |
| `comments` | `comments: [{ id, author, text, done, date?, parent?, truncated? }]` (**sınırlı** liste — yorum başına `text` en çok 2.000 karakter, kırpıldıysa `truncated: true`; `author` en çok 255 karakter; en çok 2.000 yorum ve toplam ~500.000 karakter, fazlası listede yoktur ve `more: true` gelir. Tam metin için ayrı komut yoktur: kullanıcı balonda tam metni görür, uygulama dosyayı `save` ile alıp okuyabilir. Sıra: metne bağlı yorumlar metin sırasıyla, her biri ardından yanıtlarıyla; `parent` zincirin ilk yorumunun `id`'si; `done`: zincir çözülmüş; `date` dosyadaki ISO metni), `more`? (`true`: 2.000'den çok yorum) | Yalnız belgede: liste son bildirilenden farklı olunca (yorum komutları, kullanıcının balondan/araç çubuğundan yaptığı, geri al/yinele). Aynı değişikliğin `changed` iletisinden **önce** gelir. Başlangıç listesi `loaded` iletisindedir |
| `sheet` | `index`, `name` | Gösterilen sayfa (sekme) değişince |
| `sheets` | `sheets: [{ name, state, kind, tabColor? }]` (`loaded.sheets` ile aynı biçim, listenin tamamı), `activeSheet` (gösterilen sayfanın dizini) | Sayfa listesi değişince: sayfa eklendi, silindi, taşındı, kopyalandı, gizlendi/gösterildi, adı ya da sekme rengi değişti — komutla, kullanıcının sekme şeridinden ya da bunların geri alınması/yinelenmesiyle. Aynı değişikliğin `sheet` ve `changed` iletilerinden **önce** gelir. Belgede gelmez |
| `done` | `id`, `command`, `readOnly`? (`true`: `setReadOnly(false)` istendi ama dosya salt okunur kalıyor), `revisions`? ve `trackChanges`? (yalnız `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges` yanıtında: komuttan sonraki değişiklik sayısı ve izleme durumu), `commentId`? (yalnız `addComment` ve `replyComment` yanıtında: yeni yorumun kimliği), `open`? (yalnız `commentsPanel` ve `printPreview` yanıtında: panel / önizleme açık mı), `changed`?, `tabStops`?, `defaultTabStop`? (yalnız `setTabStops` ve `setDefaultTabStop` yanıtında), `pageSetup`? (yalnız sayfa yapısı komutlarının yanıtında) | `id` taşıyan `setReadOnly`, `setLocale`, `undo`, `redo`, `print`, `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges`, `addComment`, `editComment`, `replyComment`, `resolveComment`, `removeComment`, `commentsPanel`, `printPreview`, `insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `saveResult`, `close` başarılı (`print`'te: yazdırma penceresi kapandı; `insertSheet` ve `copySheet`'te ek olarak `index`, `name`) |
| `link` | `url` | Belgede bir bağlantı etkinleştirildi (Bölüm 7) |
| `error` | `id`?, `code`, `message`; `saveId`? (yalnız bırakılan **kullanıcı** kaydında: `save-failed`, bkz. "Saklananlar için kurallar") | Aşağıdaki durumlarda |

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
{ "type": "comments", "comments": [{ "id": 1, "author": "Ayşe", "text": "Bu sayı doğru mu?", "done": true, "date": "2026-10-01T09:00:00Z" }, { "id": 7, "author": "Ali", "text": "Evet.", "done": true, "parent": 1 }] }
{ "type": "done", "id": "8", "command": "addComment", "commentId": 9 }
```

`warnings` girdileri İngilizce, geliştiriciye yönelik metinlerdir (ör. `"2 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over"`, `"Sheet: the content of a chart sheet is not in the model; it is written as an empty worksheet"`). Yalnız okuyucunun modele aldığı ama yeni dosyaya taşınamayan içerik listelenir; okuyucunun hiç okumadığı içerik (kılavuzdaki "Henüz olmayanlar" tablosu) yeni dosyada yoktur ve uyarılmaz. Bu yüzden arşivlik bir dosyanın özgün hali saklanmalıdır.

### Hata kodları

| Kod | Anlam |
|---|---|
| `bad-message` | İleti sözleşmeye uymuyor: bozuk JSON, bilinmeyen `type`, eksik ya da yanlış türde alan, geçersiz base64, izin verilmeyen `url` (iframe içinde her `url`), `data`/`url`/`blank`'ten birden fazlası ya da hiçbiri, `"docx"`/`"xlsx"` dışında bir `blank`, olmayan sayfa dizini. Ayrıca bir komut (`undo` vb.) beklenmedik biçimde başarısız olursa da bu kod döner |
| `not-loaded` | `save`, `undo`, `redo`, `setReadOnly`, `setLocale`, `print`, `exportPdf`, `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges`, `setTabStops`, `setDefaultTabStop`, bir sayfa yapısı komutu (`setPageSetup`, `insertPageBreak`, `removePageBreak`, `resetPageBreaks`, `insertSectionBreak`) ya da bir sayfa komutu (`insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`), `refreshPivots`, `insertPivot`, `setPivotFields`, `pivotFields`, `deletePivot`, `commentsPanel` ya da `printPreview` gönderildi ama açık dosya yok |
| `read-only` | Sayfa komutu, sayfa yapısı komutu, `refreshPivots`, `insertPivot`, `setPivotFields` ya da `deletePivot` gönderildi ama görünüm (belge) salt okunur; ya da `acceptAllRevisions` / `rejectAllRevisions` / `setTrackChanges` / `setTabStops` / `setDefaultTabStop` gönderildi ama belge salt okunur ya da `comments` korumalı (yalnız yorum alır; öyle açıldı, `setReadOnly(true)` denildi, ya da yazılamayan bir biçimde) |
| `no-document` | `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges`, `setTabStops`, `setDefaultTabStop` ya da bir yorum komutu (`addComment`, `editComment`, `replyComment`, `resolveComment`, `removeComment`, `commentsPanel`) gönderildi ama açık dosya bir tablo |
| `comment-refused` | Bir yorum komutu hiçbir şey değiştirmedi: öyle bir yorum yok, yorum oraya konamıyor (üstbilgi, altbilgi, metin kutusu; seçim yokken `anchor` verilmedi), yorum değiştirilemiyor (işaretleri olduğu gibi saklanan içerikte) ya da değişecek bir şey yok (zaten çözülmüş, aynı metin). Belge değişmez. Yorum komutları salt okunur belgede (`loaded.readOnly`, `load.readOnly: true`, dönüştürülmemiş doc/RTF) `read-only` ile yanıtlanır |
| `no-sheet` | Sayfa komutu, `insertPageBreak` / `removePageBreak` / `resetPageBreaks`, `refreshPivots`, `insertPivot`, `setPivotFields`, `pivotFields`, `deletePivot` ya da `printPreview` gönderildi ama açık dosya bir belge, ya da verilen dizinde sayfa yok (`insertSectionBreak` tabloya gönderilirse `no-document`) |
| `page-setup-refused` | Tabloda sayfa yapısı komutunu görünüm reddetti (geçersiz kâğıt kodu, ölçek, kenar boşluğu; kâğıtta yer bırakmayan kenar boşlukları; grafik sayfası; birlikte düzenleme; yazılmakta olan alan uygulanamadı) ya da belge görünümü sayfa yapısını değiştiremiyor. `message` nedeni görünümün dilinde verir |
| `pivot-refused` | `refreshPivots` bütünüyle reddedildi: birlikte düzenleniyor, yazılmakta olan uygulanamıyor ya da görünüm pivot tablo yenileyemiyor. `message` nedeni görünümün dilinde verir. Tek tek reddedilen tablolar hata değildir: `done` içindeki `refused` listesindedir; `insertPivot` / `setPivotFields` / `pivotFields` / `deletePivot` reddedildi (kaynak ya da hücre okunamıyor, `index` yok, düzen burada kurulmuyor, yer dolu, birlikte düzenleniyor…; `message` nedeni görünümün dilinde verir) |
| `sheet-refused` | Sayfa komutunu görünüm reddetti; `message` nedeni görünümün dilinde verir (kullanıcıya gösterilebilir) |
| `locale-refused` | `setLocale` gönderildi ama görünüm dilini değiştirmedi: tabloda içinde değişiklik yapılmış bir kural kutusu açık (kullanıcı onaylamalı ya da kapatmalı). Hiçbir şeyin dili değişmez; `message` nedeni görünümün dilinde verir |
| `print-unavailable` | `print` gönderildi ama yazdırılamıyor: pencerede `print()` yok (ya da sayfayı barındıran uygulamanın verdiği görünüm yazdıramıyor) |
| `print-failed` | Yazdırma başlatılamadı ya da yarıda kaldı: belge ya da tablo sınırların ötesinde uzun, tabloda yazdırılacak bir şey yok, bir sayfa çizilemedi, `print()` hata fırlattı, ya da sayfalar hazırlanırken dosya kapatıldı |
| `pdf-unavailable` | `exportPdf` gönderildi ama PDF üretilemiyor: sayfada tuval yok (ya da sayfayı barındıran uygulamanın verdiği görünüm PDF üretemiyor) |
| `pdf-unsupported` | `exportPdf` `text: "vector"` ile gönderildi ve bir sayfa vektör (metin + çizim) yazılamıyor: cihaz fontuyla yazı, fontta olmayan karakter, gradyan… `message` sayfa numarasını ve nedeni (İngilizce) söyler; neden dosyadaki bir font adını alıntılayabilir — **düz metin** olarak gösterin (denetim karakterleri temizlenmiştir, en çok 300 karakter). `text` vermeden ya da `"auto"` ile yeniden isteyin: o sayfa görüntü yazılır |
| `pdf-too-large` | PDF üretildi ama bir iletinin taşıdığından büyük: ham **64 MB**'ın (67.108.864 bayt) üstünde. `message` boyutu söyler. Daha düşük `dpi` ile, tabloda `sheets: "active"` ile yeniden isteyin |
| `pdf-failed` | PDF üretilemedi ya da yarıda kaldı: sınırların ötesinde uzun (2.000 sayfa; dosya 768 / dokunmatikte 192 MB), tabloda yazılacak bir şey yok, bir sayfa çizilemedi, tabloda yazılmakta olan alan uygulanamadı, ya da sayfalar çizilirken dosya kapatıldı |
| `unsupported` | Dosya gösterilemeyen bir türde (Word 95 ve öncesi, **okunamayan bir şifrelemeyle** korunan dosya — Extensible, sertifika, `spinCount` sınırın üstünde —, Office dosyası olmayan içerik, boş içerik) |
| `corrupt` | Dosya xlsx gibi başlıyor ama okunamıyor (kesik, bozuk); şifreli dosya bozuk ya da bütünlük denetimini (HMAC) geçmiyor |
| `password-not-expected` | `password` iletisi geldi ama parola sorulmuyor: yüklenen dosya yok, soru yanıtlanmış ve parola denetleniyor, ya da yükleme bitmiş. İleti saklanmaz; her `passwordRequired` bir kez yanıtlanır |
| `password-cancelled` | Parola korumalı dosya için parola verilmedi: uygulama `password` iletisinde `value: null` gönderdi ya da kullanıcı sayfadaki parola kutusundan vazgeçti. Açık dosya (varsa) açık kalır |
| `load-failed` | `url` getirilemedi (ağ hatası, HTTP hata kodu, zaman aşımı, boyut sınırı, içerik güvenliği ilkesi engeli) ya da yükleme bir `close` / yeni `load` ile iptal edildi |
| `save-failed` | Kaydetme sırasında hata; ya da `saved` iletisi uygulamaya teslim edilemeden bekleyen iletiler sınırı aşıldı |

`message` İngilizce, geliştiriciye yönelik bir açıklamadır ve yalnız uygulamaya gider; sayfa onu kullanıcıya göstermez (sayfa kendi yerelleştirilmiş genel başlığını gösterir, Bölüm 6). Kullanıcıya gösterilecek metni uygulama `code`'a göre seçmelidir. Hiçbir hatalı ileti sayfayı çökertmez; sayfa sonraki iletileri işlemeyi sürdürür.

### Parola korumalı dosyalar

Şifreli docx/xlsx (Agile, Standard) ve doc/xls (RC4, RC4 CryptoAPI, XOR) açılır; ayrıntı ve sınırlar: kullanım kılavuzu, "Parola korumalı dosyalar".

```
uygulama → { "type": "load", "id": "1", "data": "…", "password": "Parola1234_" }      // parola biliniyorsa
sayfa    → { "type": "loaded", "id": "1", "format": "xlsx", "encrypted": true, … }

uygulama → { "type": "load", "id": "2", "data": "…" }                                 // bilinmiyorsa
sayfa    → { "type": "passwordRequired", "id": "2", "wrong": false, "attempts": 0, "format": "ooxml" }
uygulama → { "type": "password", "value": "yanlış" }
sayfa    → { "type": "passwordRequired", "id": "2", "wrong": true, "attempts": 1, "format": "ooxml" }
uygulama → { "type": "password", "value": "Parola1234_" }     // ya da { "value": null } → error / password-cancelled
sayfa    → { "type": "loaded", "id": "2", "format": "xlsx", "encrypted": true, … }
```

- **`load.password`** (dizge, en çok 1.024 karakter): dosyanın parolası. Şifreli olmayan dosyada etkisizdir. Yanlışsa `passwordRequired` (`wrong: true`, `attempts: 1`) gelir. Verilen parola **yalnız** denenir (her deneme tek anahtar türetmedir); Office'in herkesçe bilinen parolası yalnız parola hiç verilmediğinde denenir.
- **`loaded.defaultPassword: true`**: dosya o herkesçe bilinen parolayla (`VelvetSweatshop`) şifreli; sorusuz açıldı. `encrypted: true` ile birlikte gelir ve **gizlilik sağlamadığını** söyler. **`loaded.integrity: false`**: dosya bütünlük denetimi (HMAC) taşımıyor; içeriğini parolayı bilmeyen biri değiştirmiş olabilir — uygulama kullanıcıyı uyarmalıdır (alan yalnız bu durumda gelir). Kaydedilen dosya her zaman denetim taşır.
- **İki soran vardır, ilk yanıt geçerlidir:** sayfa `passwordRequired` gönderir **ve** editörün alanında kendi parola kutusunu gösterir (TR/EN, `load.locale`'e göre; Escape vazgeçer). Kullanıcının kutuya yazdığı ile uygulamanın `password` iletisi eşdeğerdir. Uygulamanın `password` iletisi gelince sayfadaki kutu **kapanır** (parola yanlışsa yeni soruyla birlikte yeniden açılır). Kutuda "Vazgeç", parola denetlenirken de çalışır: süren denetimi keser ve yükleme `password-cancelled` ile biter. Böylece iletiyi tanımayan (eski) bir uygulamada da dosya açılır. Kendi kutusunu gösteren uygulama `load` iletisine **`passwordPrompt: false`** ekler: sayfa kutu göstermez, yalnız uygulamaya sorar.
- **Parola beklenirken gelen öbür komutlar** (`exportPdf`, `save`, `undo`, `print`, sayfa ve yorum komutları…) kaybolmaz ve reddedilmez: yüklemenin arkasında sırada bekler, dosya açılınca sırayla yanıtlanır; paroladan vazgeçilirse her biri `not-loaded` alır. Sırasını beklemeyenler yalnız `password`, `close`, `hello` ve `saveResult`'tır. `load.collab` ile istenen oda da ancak dosya açılınca kurulur (`collabState` o zaman gelir).
- **`collab` ile birlikte:** odadaki sürüm, yüklenen (şifreli) baytlardan hesaplanır ve `saved` ile verilen şifreli baytların sürümü bildirilir; uygulama `collabSaved` ile gelen dosyayı yeniden yüklerken parolayı `load.password` ile yeniden verir. Hücre değişiklikleri aktarıcıya şifresiz gider (dosyanın parolası oturumu korumaz). **`exportPdf`'in ürettiği PDF şifresizdir.**
- Bekleme sırasında gelen `close` ya da yeni bir `load` yüklemeyi iptal eder (`error` / `load-failed`, "The load was cancelled") ve kutuyu kaldırır. Uygulamanın kendi istek zaman aşımı varsa (Flutter sarmalayıcısında 60 sn) parola beklenirken de işler: süreyi uzatın ya da zaman aşımında `close` gönderin.
- **`loaded.encrypted: true`** ve **`saved.encrypted: true`**: dosya parolayla açıldı; `save` yanıtı, kullanıcının/kendiliğinden kaydın `saved` iletisi **aynı parolayla şifrelenmiş** dosyayı taşır (bileşik dosya; `format` yine `docx`/`xlsx`'tir — içerik türünü `encrypted`'a bakarak seçin: şifreli dosyalar `application/vnd.openxmlformats-…` yerine genel ikili olarak da saklanabilir, uzantı aynı kalır). Parolasız kaydetme ya da parola değiştirme köprüde **yoktur**.
- `convert: true` ile açılan şifreli doc/xls, aynı parolayla şifreli docx/xlsx olarak kaydedilir. `convert` verilmezse salt okunur gösterilir ve `save` özgün (şifreli) baytları verir.
- **Parola hiçbir iletide geri yansıtılmaz** (ne `done`, ne `error`, ne `passwordRequired`); hata iletileri parolayı ya da anahtarı içermez. Sayfa parolayı bayt olarak tutar ve `close`, yeni `load` ya da köprünün kapanmasıyla sıfırlar; açılıp da gösterilemeyen dosyanın (yükleme araya giren bir `close` / `load` ile geçersizleşti, dönüştürme ya da model okuma hata verdi) parolası da hemen sıfırlanır. Uygulama→sayfa kanalının kendisi (ör. `postMessage`, `runJavaScript`) uygulamanın sorumluluğundadır: parolayı günlüğe yazan bir ileti kaydedicisi kullanmayın.
- Sayfa durumu: parola beklenirken `onState('password')` (köprü sayfası "Dosya açılıyor…" yazısını kaldırır, kutu görünür); parola verilince yeniden `loading`.

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

- **xlsx ve docx gösterilir ve düzenlenir; doc, RTF ve xls kendi biçimlerine yazılamadığı için salt okunur gösterilir**, `load.convert` ile docx/xlsx'e çevrilerek düzenlenebilir açılır (Bölüm 3). Parola korumalı dosyalar parolasıyla açılır ve aynı parolayla şifreli kaydedilir (Bölüm 3, "Parola korumalı dosyalar").
- **Yazılamayan biçimler (dönüştürülmemiş doc, RTF, xls) için sözleşme:** `loaded` iletisi `readOnly: true` taşır. `setReadOnly` ile düzenleme açılamaz: istek `done` ile yanıtlanır ama yanıt `readOnly: true` taşır ve görünüm değişmez (düzenlemeler kaydedilemezdi). `save` özgün baytları döner; `changed` hiç gelmez.
- **Belge (docx; doc ve RTF) için sözleşme:** `loaded` iletisi `{ format, kind: 'document', sheets: [], activeSheet: 0 }` taşır; `load.sheet` yok sayılır. docx, `load.readOnly` `true` değilse düzenlenebilir açılır. Belge her değiştiğinde (yazma, silme, biçimlendirme, yapıştırma, geri al/yinele) tablodakiyle **aynı** `changed` iletisi gelir; belgenin sayfa (sheet) kavramı olmadığı için `sheets` hep boş dizidir, `canUndo`/`canRedo` belgenin geri alma geçmişini bildirir. `undo`, `redo`, `setReadOnly`, `setLocale` belgede de çalışır. `save`, süren bir giriş yöntemi (IME) kompozisyonunu önce belgeye işler ve değişiklikleri içeren docx baytlarını döner; hiç değiştirilmemiş bir docx özgün baytlarıyla döner. `selection` ve `sheet` iletileri belgede gelmez.
- **Belgede bağlantılar:** sayfa bir bağlantıyı kendisi açmaz, uygulamaya `{ type: 'link', url }` iletisi gönderir (yalnız http, https, mailto, tel). Salt okunur belgede bağlantıya basmak/dokunmak yeter; düzenlenen belgede düz tıklama imleci koyar, bağlantı Ctrl (macOS'ta Cmd) + tıklama ile ya da bağlantı odaktayken Enter ile etkinleşir. Dokunmatik ekranda düzenlenen belgede dokunma imleci koyar; bağlantıların dokunarak izlenmesi gerekiyorsa belge `setReadOnly(true)` ile salt okunur yapılır.
- **Base64 bellek maliyeti.** Dosya iki yönde de base64 dizgisi olarak taşınır: metin dosyadan üçte bir büyüktür ve yükleme/kaydetme sırasında dosyanın birkaç kopyası aynı anda bellekte durur (uygulamadaki baytlar, base64 metni, WebView'e geçirilen betik, sayfadaki dizgi, çözülmüş baytlar). Sayfanın kendi kodlama/çözme işlemi parça parça ve doğrusal sürede çalışır (20 MB için Node'da gidiş-dönüş ≈ 0,4 sn; tarayıcılarda 5 MB'lık dosya için yükleme ≈ 0,1–0,4 sn, kaydetme < 0,1 sn ölçüldü). Asıl maliyet platformun dizgiyi WebView'e geçirmesindedir ve gerçek cihazda ölçülmedi.
- **Gerçek cihazda sınanmadı.** Özellikle şunlar doğrulanmayı bekliyor: iOS'ta odak yakınlaştırmasının gerçekten engellendiği, güvenli alan boşlukları, ekran klavyesi açıldığında yerleşim, çok büyük `runJavaScript` betiklerinin davranışı, Android'in eski WebView sürümleri.
- `load` ve `save` için ileti boyutu sınırı yoktur (`url` ile getirmede 256 MB); bellek sınırı platformundur.
- **Yazdırma** (belge ve tablo) sayfanın `window.print()`'i üzerinden yapılır; WebView'lerde uygulamanın bu işlevi yerel yazdırmaya bağlaması gerekir (Bölüm 3, "Yazdırma"). Sayfalar görüntü olarak basılır (yazdırma penceresinden alınan PDF'te metin seçilemez).
- **PDF dosyası** `exportPdf` komutuyla alınır (Bölüm 3, "PDF"): `window.print()` gerekmez, metni aranır ve kopyalanır; font dosyası elde olan sayfalar vektör, öbürleri görüntüdür (`pdf.vector` kaçının vektör olduğunu söyler). Gerçek WebView'de **denenmedi**.
- **Kendiliğinden kaydetme** (`load.save.auto`): çalışırken açılıp kapatılamaz; durduğunu ayrı bir ileti söylemez (son `saveState` `dirty: true, saving: false` kalır); sayfa gizlenirken verilen ileti uygulamaya ulaşmayabilir. Gerçek WebView'de denenmedi.
- **Kullanıcının kaydetmesi** (`load.save`): sonuç 60 saniye beklenir ve bu süre iletiyle değiştirilemez; kaydedilmemiş değişiklik varken sayfadan ayrılma uyarısı yoktur (uygulamanın işi); `saved` iletisi teslim edilemeden çok büyürse (64 M karakter sınırı) `save-failed` hatasına döner ve kayıt zaman aşımıyla başarısız sayılır. Gerçek WebView'de denenmedi.
- Aynı anda tek dosya açıktır.

## Birlikte düzenleme: `load.collab` ve `collabState`

`load` iletisi `collab` alanı taşıyabilir: bir aktarıcının adresi (`"wss://…"`) ya da `{ "url": "wss://…", "room"?: "…", "user"?: { "name"?: "…", "color"?: "#RRGGBB" }, "token"?: "…" }`. Çalışma kitabı (dönüştürülmemiş xlsx) ya da belge (dönüştürülmemiş, değişiklik izlemeyen docx) o odaya katılır; kurallar ve kapsam kılavuzdaki "Birlikte düzenleme" bölümüyle ve [`collab-sozlesmesi.md`](collab-sozlesmesi.md) ile aynıdır. Adres `ws://` ya da `wss://` değilse `load` `bad-message` ile reddedilir. Özel taşıyıcı köprüde yoktur (gömen yerel kod `BridgeHost.createTransport(url)` verebilir).

Sayfa, oturumun durumu ya da öbür katılımcıların listesi her değiştiğinde şunu gönderir:

```json
{ "type": "collabState", "state": "online", "rev": 12, "pending": 0,
  "peers": [{ "id": "c3", "name": "Ayşe", "color": "#D93025", "readOnly": false }] }
```

- `state`: `connecting`, `online`, `offline` (bağlantı yok; düzenleme sürer, değişiklikler birikir), `closed` (oturum bitti). xls, doc, RTF, dönüştürülmüş dosya ya da **değişiklik izleyen belge** (`trackChanges: true` verilmiş ya da belge kendisi istiyor) için `collab` verilirse dosya yine açılır ve bir kez `{ "type": "collabState", "state": "closed", "reason": "unsupported", … }` gönderilir. Belgede `people` öğelerinde `sheet` alanı yoktur; başkasından gelen değişiklikte uygulama olağan `changed` iletisini (`sheets: []`) alır ve `saveState.dirty` `true` olur. Birlikte düzenlenen belgede yorum, izleme ve görsel komutları (`addComment`, `setTrackChanges`, `acceptAllRevisions`…) hiçbir şey değiştirmez (sayfa bildirimle söyler); kapsam kılavuzdaki "Belgeleri birlikte düzenleme" bölümündedir.
- **`reason`** (yalnız `state: "closed"` iken ve oturum kendiliğinden bittiyse): `"version"` — oda başka bir protokol sürümüyle (başka bir paket sürümünü yükleyen sayfayla) düzenleniyor; sayfa kullanıcıya "Bu dosya başka bir sürümle düzenleniyor; sayfayı yenileyin." der, uygulama da sayfayı/köprüyü güncel paketle yeniden yüklemelidir (yeniden `load` aynı paketle yine reddedilir). `"refused"` — aktarıcı başka bir nedenle reddetti ya da odayı kapattı; `"peer"` — odadan bu dosyaya ait olmayan bir şey geldi; `"stale"` — dosya başka yerde değişti ve yeniden yüklenemedi; `"unsupported"` — bu dosya birlikte düzenlenemez (yukarıda). Oturum açıkken ya da odadan isteyerek çıkıldığında `reason` yoktur.
- `peers[].name` öbür katılımcının verdiği **metindir** (`null`: ad vermemiş); uygulama onu metin olarak göstermelidir.
- `changed` iletisindeki `canUndo` / `canRedo`, birlikte düzenlenen kitapta **katılımcının kendi** adımlarını söyler.
- `state` ayrıca `detached` olabilir: oda başka bir sürümde ve bu kitapta kaydedilmemiş değişiklik var; kitap yerinde kalır, uygulama dosyayı kaydetmeli ve güncelini yeniden `load` etmelidir (köprü dosyayı kendisi yeniden yükleyemez).
- **Köprüde tek kaydedici, nöbet ve kanıt YOKTUR.** Köprüde kayıt uygulamadan geçer (`saved` → `saveResult`); sayfa saklanan dosyayı kendisi çekemez. Bu yüzden `mount()`'taki "odada tek katılımcı kendiliğinden kaydeder / kaydeden olmazsa nöbetle kaydet / başkasının kaydını dosyaya bakarak doğrula" kuralları köprüde çalışmaz ve `saveState`'te `autoWaits` alanı yoktur (önceki sürüm notu yanlıştı). Köprüde her sayfa kendi `autoSave`'iyle kendi işini kaydeder; `collabSaved` ile gelen "biri kaydetti" sözünü doğrulamak ve çakışan kayıtları (ör. `saveResult conflict`) çözmek **uygulamanın** işidir. Başkasından gelen değişiklik sayfayı `saveState.dirty: true` yapar ama kendiliğinden kaydı başlatmaz: yalnız başkasının işini taşıyan sayfa, kendisi (kullanıcı ya da uygulama `save` ile) kaydedene dek "kaydedilmemiş" görünür.
- **Çatal sorusu köprüde yoktur.** `mount()`'taki üç yollu soru (saklananı yükle / benimkini kaydet / vazgeç + kopyayı indir) köprü sayfasında açılmaz; köprü `saveResult conflict` için kendi iki düğmeli sorusunu kullanır ve odadan ayrı düşmüş (`detached`) kitabı sunucudaki dosyayla değiştirmek uygulamanın yeni bir `load` göndermesiyle olur.
- **`collabState.people`**: dosyadaki herkes, bu sayfanın katılımcısı başta: `[{ id, name, color, self, readOnly, sheet? }]` (`name` verilen ad ya da sayfanın dilinde "Misafir N"; en çok 200 kişi). `peers` (yalnız öbürleri, `name` verilmediyse `null`) olduğu gibi durur. Adlar katılımcının bildirdiğidir, kimlik kanıtı değildir; metin olarak gösterin.
- **Damga ve dönem (protokol 2 ve sonrası):** sayfa, yüklenen dosyanın içindeki oda damgasını (`BymOfficeCollab` özel özelliği) okuyup odaya o adımdan katılır ve `save` ile verdiği baytlara odanın dönemini ve adımını yazar; uygulama o baytları değiştirmeden saklamalıdır. Aktarıcının unuttuğu oda başka bir odadır: sayfa eski adımına güvenmez (kaydedilmemiş iş yoksa `collabState closed`/yeniden `load` beklenir, varsa `detached`).
- **Bildirimler (`notice`)**: birlikte düzenlemede yeni metinler — geri alma sırasını beklerken bırakıldı / zamanında yerleşmedi (`undo-wait`), gönderilen adım 30 sn onaylanmadı ve yeniden bağlanılıyor (`slow`), bekleyen adım düştü, formül yazılırken hücreler kaydı (yazılan metin bildirimdedir).
- Kayıt: kullanıcının kaydı (`saved by:"user"`/`"auto"`) için uygulama `saveResult ok` dedikten **sonra** sayfa öbür katılımcılara o baytların özetini ve tuttuğu adımı bildirir (uygulamanın `save` komutuyla çektiği baytlar bildirilmez: saklandıklarını sayfa bilemez). Başka bir katılımcının bildirdiği kayıt uygulamaya `{ "type": "collabSaved", "etag": "…" | null, "rev": sayı | null, "versions": […] }` iletisiyle aktarılır; bunun doğru olup olmadığını **uygulama** kendi sunucusuna sorarak doğrular — sayfa hiçbir şeyi kaydedilmiş saymaz. Aktarıcı katılımcıyı salt okunur yaparsa görünüm salt okunur olur.
- **Kısıt:** köprü sayfasının içerik güvenlik ilkesi `connect-src 'self' blob: data:` der; aktarıcı sayfayla **aynı kökende** değilse tarayıcı bağlantıyı engeller. Başka kökendeki aktarıcı için sayfanın kendi kopyanızdaki ilkesine o adresi eklemeniz gerekir.
