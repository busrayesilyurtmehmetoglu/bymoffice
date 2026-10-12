# BYM Office

Word ve Excel dosyalarını (docx, xlsx, doc, xls, RTF) tarayıcıda açıp düzenlemek için sıfırdan yazılan motor.
Tek bir `.js` dosyası olarak dağıtılır, çalışma zamanı bağımlılığı yoktur; aynı kod Node üzerinde sunucuda da çalışır.

## Kullanım

```html
<script src="https://cdn.example.com/bymoffice.min.js"></script>
<script>
  const file = await BymOffice.open(blobOrBytes); // biçim içerikten algılanır
  file.format;            // 'docx' | 'xlsx' | 'doc' | 'xls'
  const bytes = await file.save();

  // Editör: dosyayı adresinden açar, düzenletir ve Kaydet düğmesi / Ctrl+S ile aynı adrese geri kaydeder.
  await BymOffice.mount(document.getElementById('editor'), { file: '/files/42', saveTo: true });
</script>
```

ES modülü olarak: `import { open } from './bymoffice.esm.js'`.

Ayrıntılı anlatım, API ve test arayüzü: [geliştirici kullanım kılavuzu](docs/kullanim-kilavuzu.md). Sürüm notları: [CHANGELOG.md](CHANGELOG.md).

## Fontlar ve üçüncü taraf içerik

Paket, Bookman Old Style'ı olmayan cihazlarda (macOS, iOS, Android) o fontla yazılmış metni göstermek için **TeX Gyre Bonum** fontunu getirir (`dist/fonts/tex-gyre-bonum/`; B. Jackowski ve J. M. Nowacki, GUST e-foundry). Font dosyaları değiştirilmeden, kendi lisanslarıyla (GUST Font License) dağıtılır ve yalnız gerektiğinde yüklenir; ayrıntı: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) ve kılavuzun "Fontlar" bölümü. Betiği kendi sunucunuzdan sunuyorsanız `dist/fonts/` klasörünü de yanına koyun; ES modülü ya da paketleyici kullanıyorsanız klasörün adresini `mount(el, { fontsUrl })` ile verin.

## Komutlar

| Komut | İş |
|---|---|
| `pnpm verify` | Her şeyi sıfırdan sınar (tip denetimi, birim, derleme, tarayıcı testleri, dosya havuzu) |
| `pnpm demo` | Test arayüzünü sunar: `http://127.0.0.1:4318/demo/index.html` |
| `pnpm test` | Birim testleri |
| `pnpm e2e` | Playwright tarayıcı testleri |
| `pnpm build` | `dist/bymoffice.min.js` (global `BymOffice`), `dist/bymoffice.esm.js`, tip tanımları |
| `pnpm corpus <klasör>` | Gerçek dosyalarda aç → kaydet → yeniden aç denetimi |

`demo/index.html` derlenmiş paketi yükler; dosyaları sürükleyip parçalarını ve içeriklerini görebilir, kayıpsız kaydetme denetimini izleyebilir ve yeniden kaydedebilirsiniz.

## Kayıpsızlık ilkesi

Motorun anlamadığı parçalar hiç ayrıştırılmaz; ZIP içindeki sıkıştırılmış baytlarıyla saklanır ve kaydederken aynen geri yazılır.
Yalnız değiştirilen parçalar yeniden üretilir. Her yeni özellik `pnpm corpus` denetiminden geçmelidir.

## Durum

- [x] Aşama 0: ZIP, XML, OPC (docx/xlsx kapsayıcısı), CFB okuyucu (doc/xls kapsayıcısı), biçim algılama, gidiş-dönüş denetimi
- [x] Aşama 1: xlsx tablo modeli, Canvas ızgara, hücre düzenleme, kayıpsız xlsx yazma, editör arayüzü, WebView köprüsü
- [ ] Aşama 2: formül motoru (253 fonksiyon), biçimlendirme, koşullu biçim gösterimi tamam; süzgeç, sıralama, veri doğrulama sürüyor
- [x] Aşama 3: docx belge modeli, sayfa düzeni motoru, salt okunur belge görünümü
- [ ] Aşama 4: Word düzenleme
- [ ] Aşama 5: xls (BIFF8), doc (Word 97–2003) ve RTF okuyucuları tamam (salt okunur); bu biçimlere yazma yok
- [ ] Aşama 6: grafikler, pivot, PDF çıktısı
- [ ] Aşama 7: eşzamanlı düzenleme

Bilinen eksikler: parola korumalı dosyalar; docx, doc, RTF ve xls düzenleme (salt okunur açılırlar). Ayrıntı: `docs/kullanim-kilavuzu.md`.
