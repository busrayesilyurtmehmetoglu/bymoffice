# BYM Office — geliştirici kullanım kılavuzu

Bu kılavuz, BYM Office'i kendi uygulamasına ekleyecek geliştiriciler içindir. Her kod değişikliğinde güncellenir;
burada yazan her şey depodaki sürümde çalışır durumdadır, henüz yapılmamış olanlar "Henüz olmayanlar" başlığında açıkça listelenir.

**Kılavuzun karşılık geldiği sürüm:** 0.0.1 (2026-10-01) — kapsayıcı katmanı.

## 1. Bu sürüm ne yapar

- docx, xlsx, doc ve xls dosyalarını **içeriğinden** tanır (dosya adına bakmaz).
- docx/xlsx içindeki her parçayı (XML, görsel, bilinmeyen ikili parça) okur.
- Bir parçayı değiştirip dosyayı yeniden kaydeder; dokunulmayan parçalar özgün sıkıştırılmış baytlarıyla, **bayt bayt aynı** kalır.
- doc/xls içindeki akışları (stream) okur.
- Tarayıcıda ve Node'da aynı kodla, çalışma zamanı bağımlılığı olmadan çalışır.

### Henüz olmayanlar

| Özellik | Durum |
|---|---|
| Tablo/belge görüntüleme ve düzenleme arayüzü (`BymOffice.mount()`) | Yok (Aşama 1 ve sonrası) |
| Hücre değerleri, stiller, formüller için model | Yok (görev 1.1'den itibaren) |
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

### `open(input): Promise<OfficeFile>`

`input`: `Uint8Array`, `ArrayBuffer` ya da `Blob`/`File`.

```ts
interface OfficeFile {
  format: 'docx' | 'xlsx' | 'doc' | 'xls';
  kind: 'document' | 'spreadsheet';
  container: 'ooxml' | 'cfb';
  package?: OpcPackage; // docx/xlsx
  cfb?: CfbFile;        // doc/xls
  save(): Promise<Uint8Array>;
}
```

Hatalar:

- `UnsupportedFileError`: girdi Word/Excel dosyası değil ya da henüz desteklenmeyen bir tür (RTF, parola korumalı).
- `Error`: dosya bozuk (iletisi `Corrupt ZIP archive: …`, `Corrupt ZIP entry …` vb. ile başlar).

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
