# BYM Office

**Word ve Excel dosyalarını tarayıcıda açın, düzenleyin, kaydedin — docx, xlsx, doc, xls ve RTF.**
TypeScript ile sıfırdan yazılmış belge ve tablo motoru: tek betik, sunucu yok, çalışma zamanı bağımlılığı yok.

[Canlı demo](https://busrayesilyurtmehmetoglu.github.io/bymoffice/) · [English](README.md) · [Geliştirici kılavuzu](docs/kullanim-kilavuzu.md) · [WebView köprüsü](docs/webview-koprusu.md) · [Sürüm notları](CHANGELOG.md)

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>
<div id="editor" style="height: 600px"></div>
<script>
  BymOffice.mount(document.getElementById('editor'), {
    file: '/files/rapor.xlsx', // adres, File/Blob ya da baytlar; biçim içerikten tanınır
    saveTo: true,              // Kaydet düğmesi ve Ctrl/Cmd+S dosyayı aynı adrese geri yazar
    locale: 'tr',
  });
</script>
```

## Neden

- **Tamamen tarayıcıda çalışır.** Dosya cihazda okunur, düzenlenir ve yazılır; siz bir yere kaydetmedikçe hiçbir şey gönderilmez (motor yalnız, belge gerektirirse paketle gelen fontu betiğin yanından yükler). Aynı kod sunucu tarafı işler için Node'da da çalışır.
- **Tek dosya, sıfır bağımlılık.** CDN'den tek `<script>` satırı ya da npm. Paket (sıkıştırılmış 4,7 MB) Bookman Old Style ile yazılmış belgeler için bir font da getirir.
- **Kayıpsız.** Motorun anlamadığı parçalar bayt bayt korunur; yalnız değiştirdiğiniz yeniden yazılır. Değişmemiş bir docx ya da xlsx açıldığı haliyle, bayt bayt kaydedilir.
- **Yalnız okuyucu değil, hazır editör.** Araç çubukları, klavye ve dokunmatik, geri al/yinele, pano, bul ve değiştir, yazdırma, PDF — Türkçe ya da İngilizce.
- **Web sayfaları ve uygulamalar.** Her web sayfası ve çatı; mobil ve masaüstü uygulamalarda tek dosyalık WebView sayfasıyla.

## Neler yapar

**Excel (xlsx; xls salt okunur açılır ya da xlsx'e çevrilir)**
- Yalnız görünen alanı çizen Canvas ızgara; dondurulmuş bölmeler, gizli satır/sütun, sayfa sekmeleri.
- Hücre düzenleme, biçim araç çubuğu, sayı biçimleri (Türkçe ve İngilizce yerel ayar), birleştirme, kenarlık, koşullu biçim, veri doğrulama, süzgeç ve sıralama.
- Formüller: 265 fonksiyon, artımlı yeniden hesap, dinamik diziler (taşma, `A1#`), `LET` ve `LAMBDA`.
- Grafikler (çizme, ekleme, düzenleme, kopyalama), pivot tablolar (yenileme, oluşturma, alan değiştirme, silme), görseller ve şekiller, notlar, sayfa koruması.
- Yazdırma, baskı önizleme, sayfa yapısı, PDF.

**Word (docx; doc ve RTF salt okunur açılır ya da docx'e çevrilir)**
- Sayfalı düzen: üstbilgi/altbilgi, dipnot, tablo, liste, görsel, kayan nesne, metin kutusu, grafik.
- İmleç, seçim, IME, biçimlendirme, stiller, listeler, tablolar, görseller, bağlantılar, sekme durakları, sayfa yapısı ile düzenleme.
- Değişiklik izleme (yazma, silme ve biçim izlenir; kabul/ret), yan panelli yorumlar, belge koruması.
- Yazdırma ve seçilebilir metinli, bağlantılı, anahatlı PDF.

**İkisinde de**
- Parola korumalı dosyalar: şifreli docx/xlsx/doc/xls açılır; docx/xlsx aynı parolayla kaydedilir.
- Eşzamanlı düzenleme: aynı çalışma kitabını ya da belgeyi birden çok kişi düzenler. Aktarıcı yalnız iletileri sıraya koyar, her arka uçta yazılabilir; bağımlılıksız bir Node örneği `server/` klasöründe gelir.
- Dosyanın geldiği yere geri kaydetme (`saveTo`, ETag/If-Match, çakışma sorusu) ya da kendiliğinden kaydetme.
- Mobil ve masaüstü uygulamalar için WebView köprüsü (`dist/bymoffice.webview.html`, internet gerektirmez).

**Henüz olmayanlar** kılavuzdaki "Henüz olmayanlar" tablosunda açıkça yazılı. En önemlisi: doc, xls ve RTF kendi biçimlerinde kaydedilemez — docx/xlsx olarak kaydedilir.

## Kurulum

**CDN** — derleme gerekmez, global `BymOffice` tanımlanır:

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>
```

**npm**

```sh
npm i bymoffice
```

```js
import { mount, open } from 'bymoffice';

// Sayfada editör
const view = await mount(document.getElementById('editor'), { file: blobOrUrl, locale: 'tr' });

// Arayüzsüz: oku, değiştir, yaz (tarayıcı ya da Node)
const file = await open(bytes);
const saved = await file.save();
```

Paketleyici kullanıyorsanız `node_modules/bymoffice/dist/fonts/` klasörünü sunduğunuz dosyalara kopyalayıp adresini verin: `mount(el, { fontsUrl: '/assets/bymoffice-fonts/' })`.

Yalnız ES modülü ve `<script>` çıktısı dağıtılır (CommonJS çıktısı yok). TypeScript tipleri pakette.

## Çatılar ve uygulamalar

| Nerede | Nasıl |
|---|---|
| Düz HTML, Laravel/Blade, Django, Rails, PHP … | CDN'den `<script>`, ardından `BymOffice.mount(...)` |
| React, Vue, Svelte, Angular | Bir effect / `onMounted` içinde `import { mount } from 'bymoffice'`; bileşen kalkarken `destroy()` |
| Next.js, Nuxt | Yalnız tarayıcı tarafında: `'use client'` bileşeni ya da SSR kapalı dinamik içe aktarma |
| Flutter, React Native, yerel iOS/Android, Electron, Tauri | `bymoffice.webview.html` bir WebView'e yüklenir, iletilerle konuşulur (`load`, `save`, `changed`, `error` …) — bkz. [köprü sözleşmesi](docs/webview-koprusu.md) |
| Node | `open()`, düzenleme işlemleri ve `save()` ekransız çalışır |

Tablo bir tariftir, test raporu değil: motor bugün bir Laravel sayfasında ve bir Flutter web uygulamasında kullanılıyor; Next.js, React Native, Electron, Tauri ve gerçek mobil WebView'ler henüz denenmedi. Motorun kendisi her değişiklikte Chromium, Firefox, WebKit ve iki mobil profilde (iPhone, Pixel) sınanır. `DecompressionStream` olmayan eski WebView'lerde motor kendi inflate/deflate kodunu kullanır.

## Geliştirme

```sh
pnpm install
pnpm build      # dist/: bymoffice.min.js (global BymOffice), bymoffice.esm.js, tipler, fontlar, WebView sayfası
pnpm site       # demo sitesi site-dist/ klasörüne
pnpm demo       # geliştirici test sayfası: http://127.0.0.1:4318/demo/index.html
pnpm verify     # tip denetimi, birim testleri, tarayıcı testleri, paket denetimi, dosya havuzunda gidiş-dönüş
```

Kılavuz ve kod yorumları motorun uyduğu kuralları anlatır (kayıpsız kaydetme, sessizce kaybetmek yerine reddetme, kötü niyetli dosyanın sayfayı dondurmasını önleyen sınırlar).

## Fontlar ve üçüncü taraf içerik

Paket, Bookman Old Style'ı olmayan cihazlar için ölçüleri ona yakın **TeX Gyre Bonum** fontunu (GUST e-foundry) getirir. Dosyalar değiştirilmeden, GUST Font License ile dağıtılır ve yalnız gerektiğinde yüklenir. Ayrıntı: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Lisans

[MIT](LICENSE)
