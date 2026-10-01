# BYM Office

Word ve Excel dosyalarını (docx, xlsx, doc, xls) tarayıcıda açıp düzenlemek için sıfırdan yazılan motor.
Tek bir `.js` dosyası olarak dağıtılır, çalışma zamanı bağımlılığı yoktur; aynı kod Node üzerinde sunucuda da çalışır.

## Kullanım

```html
<script src="https://cdn.example.com/bymoffice.min.js"></script>
<script>
  const file = await BymOffice.open(blobOrBytes); // biçim içerikten algılanır
  file.format;            // 'docx' | 'xlsx' | 'doc' | 'xls'
  const bytes = await file.save();
</script>
```

ES modülü olarak: `import { open } from './bymoffice.esm.js'`.

Ayrıntılı anlatım, API ve test arayüzü: [geliştirici kullanım kılavuzu](docs/kullanim-kilavuzu.md).

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
- [ ] Aşama 1: xlsx tablo modeli, Canvas ızgara, hücre düzenleme
- [ ] Aşama 2: formül motoru, biçimler
- [ ] Aşama 3: docx belge modeli, sayfa düzeni motoru
- [ ] Aşama 4: Word düzenleme
- [ ] Aşama 5: xls (BIFF8), doc (Word Binary) ve RTF okuyucuları; CFB yazıcı
- [ ] Aşama 6: grafikler, pivot, koşullu biçim, PDF çıktısı
- [ ] Aşama 7: eşzamanlı düzenleme

Bilinen eksikler: parola korumalı dosyalar, RTF, doc/xls yazma.
