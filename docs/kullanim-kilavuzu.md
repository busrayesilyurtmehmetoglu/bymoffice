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
- **Hücre notları eklenir, düzenlenir ve silinir** (xlsx): hücrenin sağ tık menüsünden ya da Shift+F2 ile; not hücrenin yanındaki alana yazılır, her değişiklik tek geri alma adımıdır ve `save()` ile dosyaya yazılır. Değişmeyen notlar, kutuları ve iş parçacıklı yorumlar (yazışmalar) dosyada bayt bayt korunur. Bkz. "Hücre notları".
- **Sayfa yapısını (yazdırma ayarlarını) modele okur** (xlsx ve xls): `sheet.pageSetup` — kenar boşlukları, kâğıt boyutu, yön, ölçek ve sayfaya sığdırma, üstbilgi/altbilgi metinleri, yatay/dikey ortalama, kılavuz çizgisi ve başlık yazdırma, sayfa sonları. xls → xlsx dönüştürmede yeni dosyaya yazılır. Yazdırmada uygulanır ("Tabloyu yazdırma"); bu ayarları değiştiren bir arayüz yoktur.
- **Word belgelerini yazdırır** (docx, doc, RTF; salt okunur görünümde de): belge çubuğundaki Yazdır düğmesi, Ctrl/Cmd+P ya da `print()` tarayıcının yazdırma penceresini belgenin kâğıt boyutunda açar; "PDF olarak kaydet" de oradan yapılır. Sayfalar görüntü olarak basılır (ayrıntı ve sınırlar: "Belgeyi yazdırma").
- **Tabloları yazdırır** (xlsx, xls; salt okunur görünümde de): araç çubuğundaki Yazdır düğmesi, Ctrl/Cmd+P ya da `print()` etkin sayfayı (istenirse görünür bütün sayfaları) sayfa yapısına göre kâğıda dizer: kâğıt ve yön, kenar boşlukları, ölçek ya da sayfaya sığdırma, yazdırma alanı, yinelenen başlıklar, elle sayfa sonları, üstbilgi/altbilgi, kılavuz çizgisi ve satır/sütun başlığı seçenekleri (ayrıntı ve sınırlar: "Tabloyu yazdırma").
- **Koşullu biçimleri gösterir ve düzenler**: hücre kuralları, formüllü kurallar, renk ölçekleri, veri çubukları, ikon kümeleri, ilk/son N, ortalama üstü/altı, yinelenen/benzersiz, metin ve tarih kuralları. Hücre düzenlenince ve geri alınınca güncellenir. Kural eklenir, değiştirilir, silinir ve sıralanır (araç çubuğunda **Koşullu biçimlendirme** menüsü, kural yöneticisi; aşağıda "Koşullu biçim kuralları"); kaydederken dokunulmayan kurallar dosyada bayt bayt korunur.
- **Veri doğrulama kuralı ekler, değiştirir, siler** (araç çubuğunda ve hücre menüsünde **Veri doğrulama…**; aşağıda "Veri doğrulama kuralları").
- **Formülleri hesaplar**: 225 Excel fonksiyonu; bir hücre değişince yalnız ona bağlı formüller yeniden hesaplanır.
- **WebView köprüsü** ile mobil ve masaüstü uygulamalara gömülür (ayrıntı: [webview-koprusu.md](webview-koprusu.md)).
- doc/xls içindeki akışları (stream) okur.
- Tarayıcıda ve Node'da aynı kodla, çalışma zamanı bağımlılığı olmadan çalışır.

### Henüz olmayanlar

| Özellik | Durum |
|---|---|
| Formüllerde: dinamik dizi taşması, `LET`/`LAMBDA`, tablo başvuruları (`Tablo[Sütun]`), başka dosyaya ve 3B (`Sayfa1:Sayfa3!A1`) başvurular, `;` ile argüman ayırma | Yok. Böyle bir formül içeren hücre dosyadaki son değerini korur ya da `#REF!`/`#NAME?` verir; formül metni kaydederken bozulmaz |
| Tablo ya da form denetimi içeren sayfalarda, ve içinde grafik ya da pivot tablo bulunan çalışma kitaplarında satır/sütun ekleme-silme | Reddedilir (bu içeriklerin başvuruları henüz kaydırılmıyor; sessizce bozulmasınlar diye); hücre düzenleme ve biçimlendirme serbesttir. Resim ve yorum içeren sayfalarda serbesttir |
| Formül yazarken hücre göstermede: formüldeki bütün başvuruların renklendirilmesi (yalnız son eklenen çerçevelenir); çerçeveyi sürükleyerek başvuruyu taşıma/boyutlandırma; dokunmatikte aralık gösterme (tek dokunuş hücre, satır ya da sütun ekler; sürükleme kaydırmadır); köşe düğmesiyle tüm sayfa; Ctrl+tık ile birden çok alan; adlandırılmış aralık ya da tablo adı önerme; formül çubuğundan klavyeyle hücre gösterme ve sayfa değiştirme; başka sayfa gösterilirken hücre içi düzenleyicinin görünmesi | Yok. Gösterilen başvuru her zaman göreli yazılır; `$` F4 ile ya da elle eklenir. F4'ün macOS Excel'deki eşi Cmd+T **yoktur** (tarayıcının yeni sekme kısayolu) |
| Yazı tipi adı seçimi; dikey hizalama düğmeleri | Yok |
| Kural iletişim kutularında (veri doğrulama, koşullu biçim): hücre göstererek aralık/kaynak seçme (aralık ve kaynak elle yazılır: `=$A$1:$A$5`, `A1:B10; D1`); iletişim kutusu açıkken ızgarada gezinme | Yok |
| Kopyala/kes-yapıştır ile veri doğrulama ve koşullu biçim kurallarının taşınması | Yok. Görünümün kendi kopyası değer, biçim ve notları taşır; kurallar kaynakta kalır, hedeftekiler değişmez (Excel taşır) |
| Veri doğrulamada: 2010 uzantı listesindeki kuralın (`<x14:dataValidation>`; listesi başka sayfada olan eski dosyalar) değiştirilmesi/silinmesi; "aynı ayarlı öbür hücrelere de uygula"; uyarı/bilgi türünde onay penceresi | Yok. Uzantı listesindeki kural gösterilir ve uygulanır, bayt bayt korunur; o hücreleri kapsayan `setValidation`/iletişim kutusu **reddedilir** (çift kural oluşmasın diye) |
| Koşullu biçimde: ikon kümesi ve tarih dönemi (`timePeriod`) kuralı **oluşturma/içeriğini değiştirme** (listelenir, aralığı ve "doğruysa durdur"u değişir, taşınır, silinir); 2010 uzantısında karşılığı olan kuralın (Excel 2010+'ın yazdığı veri çubukları: `extId`) **içeriğini** (renk, eşik, biçim) değiştirme — aralığı, sırası ve "doğruysa durdur"u değişir, kural **silinebilir**; yalnız uzantı listesinde duran kuralın (x14 ikon kümeleri vb.) her türlü değişikliği ve böyle bir kural içeren sayfada sıra değiştirme; biçimde kenarlık, sayı biçimi ve gradyan dolgu düzenleme (var olan korunur); veri çubuğunda/renk ölçeğinde eşik türü ve değeri seçimi (yeni kuralda en küçük–en büyük, üç renkte orta nokta %50 yüzdelik; **var olan** kuralın eşikleri, çubuk uzunlukları, `showValue` ve `stdDev` değeri olduğu gibi korunur) | Yok / reddedilir (anlaşılır iletiyle). Uzantı listesinin geri kalanı her durumda bayt bayt korunur |
| Sayfa kopyalamada: koşullu biçimlerin yalnız 2010 uzantı listesinde duran bölümü (uzantıya özgü kurallar, veri çubuğunun kenarlık/eksen/negatif renkleri: `conditionalFormatExtensions`), görseller/şekiller/grafikler, tablolar, pivot tablolar, form denetimleri, yorum zincirleri, dış bağlantılar, mini grafikler (sparkline), özel görünümler, yok sayılan hata kuralları, izlenen hücreler, sıralama durumu, fonetik ayarı, anahat ayarları ve düzeyleri, sayfanın kod adı, üstbilgi/altbilgi görselleri (`legacyDrawingHF`), sayfa arka plan görseli (`picture`), satırların **başka ad alanından** gelen öznitelikleri (Excel'in her satıra yazdığı `x14ac:dyDescent` ipucu sayılmaz; ad alansız öznitelikler — `thickBot`, `thickTop`, `ph` — kopyaya yazılır); başka çalışma kitabına kopyalama | Kopyaya **alınmaz** (dosyadaki haliyle saklanan içerik; dosyada olmayan bir sayfa için yazılamıyor). Kopyalama yapılır, alınmayanların **her biri** türü ve sayısıyla `onNotice` ile bildirilir (`copySheetLosses`). Sayfa koruması ve korumalı aralıklar kopyaya **alınır** (aralıklar, o ana dek eklenen/silinen satır ve sütunlara göre kaydırılmış yerleriyle). Sekme taşımanın klavye kısayolu yok (yalnız sekme menüsü: tarayıcılar Ctrl+Shift+PageUp/PageDown'ı kendine ayırır); grafik sayfasının sekme rengi değiştirilemez; sayfa koruması uygulanmaz ve düzenlenemez (yalnız korunur ve kopyalanır) |
| İçinde grafik, pivot tablo ya da form denetimi bulunan çalışma kitabında sayfa silme ve yeniden adlandırma; tablo içeren sayfayı silme; grafik sayfasını silme ve yeniden adlandırma; adları yalnız harf büyüklüğüyle ayrılan iki sayfası (`IL` / `İL`) olan kitapta her sayfa işlemi (ekleme dahil) | Reddedilir (bu içerikler sayfaları adıyla anar ve henüz yeniden yazılmıyor); menüde öğe devre dışı görünür. Sayfa **ekleme** her zaman serbesttir |
| Çizimlerin (şekil, görsel) "belgede yer" bağlantılarında modelin güncellenmesi | Dosyada güncellenir, modelde değil: çizim ilişkilerindeki `#Sayfa!A1` hedefleri ve sayfanın `Target="#…"` olarak yazılmış hücre bağlantıları yeniden adlandırma, silme ve satır/sütun ekle-silde **kaydedilen dosyada** yeni adı/yeri alır (silinen sayfaya bakan hedef `#REF!` olur, ilişki kalır — UNSURE: Excel'in buraya ne yazdığı doğrulanmadı; geri alınınca parça özgün baytlarına döner). Açık modeldeki `hyperlink.target` dosya yeniden açılana dek eski metni taşır |
| Koşullu biçim/doğrulama gibi "dosyadaki haliyle" tutulan içerikte 3B başvuru (`Sayfa2:Sayfa4!A1`) | UNSURE: ucundaki sayfa silinip ardından komşusunda satır/sütun eklenirse bu içerikteki başvuru, aynı başvurunun hücredeki halinden farklı çıkabilir (hücreler işlem işlem, bu içerik dosya durumundan tek geçişte ayarlanır). Excel bu kurallarda 3B başvuruya izin vermez; pratikte yalnız elle üretilmiş dosyada görülür |
| Renge ve ikona göre süzme, tablo içi süzgeç, yorumlu ya da tablolu sayfada sıralama, uyarı türü doğrulamada onay penceresi, Türkçe hata adları (`#YOK`) | Yok. Uygulanamayan süzgeç ölçütleri dosyada korunur |
| Koşullu biçimde: Excel'in özgün ikon görselleri (basit vektör ikonlar çizilir), kuraldaki yazı tipi adı/boyutu ve desenli/gradyan dolgu | Yok. `calculate: false` iken formüllü kurallar uygulanmaz |
| Word 95 ve öncesi, parola korumalı doc | Yok; `UnsupportedFileError` |
| RTF'te matematik, çerçeve yerleşimi, Word 95 çizim nesnelerinin (`\do`) çizgi ve dolguları (metin kutusu olanların yalnız metni okunur, diğerleri atlanır), eski `\pn` listeleri, tablo stilleri, Uzak Doğu kod sayfaları; RTF'e yazma | Yok. RTF salt okunur açılır, `save()` özgün baytları verir |
| doc düzenleme ve kaydetme; doc'ta EMF/WMF görseller, resim madde imleri, eski otomatik numaralar, yorum aralıkları | Yok. doc salt okunur açılır, `save()` özgün baytları verir |
| Sayfa yapısını (kenar boşluğu, kâğıt, ölçek, üstbilgi/altbilgi, sayfa sonu, yazdırma alanı) düzenleme; sayfa sonu önizlemesi; yazdırma önizlemesi | Yok. Sayfa yapısı dosyadan okunur (`sheet.pageSetup`), yazdırmada uygulanır, dosyada korunur ve xls → xlsx dönüştürmede taşınır |
| Tablo yazdırmada: bir sayfaya sığmayan satır/sütunun ortadan bölünmesi (kendi sayfasına konur, taşan kısmı kenar boşluğunda kesilir); yorumların basılması (`cellComments`: sayfa sonunda ya da göründüğü gibi); hata hücrelerinin boş/tire/`#N/A` basılması (`errors`); siyah-beyaz ve taslak kipleri; üstbilgi/altbilgi görselleri (`&G`), tema renkleri (`&K04+000`), ana hat ve gölge (`&O`, `&H`), dosya yolu (`&Z` boş kalır); grafik sayfaları; kopya sayısı ve yazıcı çözünürlüğü ayarları; yalnız seçili aralığı yazdırma; sağdan sola sayfalar; metni seçilebilir PDF | Yok. Bu ayarlar dosyada korunur, çıktıyı etkilemez |
| Belge yazdırmada: metni seçilebilir/aranabilir PDF (sayfalar görüntü olarak basılır), doğrudan PDF dosyası üretme (yalnız tarayıcının yazdırma penceresi üzerinden), sayfa aralığı/kopya sayısı seçenekleri (tarayıcının penceresinden seçilir), WebView'lerde kendiliğinden yazdırma (`window.print()` olmayan ya da hiçbir şey yapmayan WebView'de uygulama kendi yazdırmasını bağlar; bkz. [webview-koprusu.md](webview-koprusu.md)) | Yok |
| xls düzenleme ve xls olarak kaydetme; xls'te koşullu biçim, veri doğrulama, otomatik süzgeç ölçütleri, yazıcı sürücüsü verisi (PLS), özel görünümlerin (custom view) sayfa yapısı, üstbilgi/altbilgi görselleri, özet (outline) ayarları, sayfa koruması, grafikler, form denetimleri, makrolar, dış çalışma kitabı yolları; parola korumalı xls; Excel 4 ve öncesi | Yok. xls salt okunur açılır, `save()` özgün baytları verir. Bunlar modele okunmadığı için `saveAsXlsx()` ile üretilen xlsx'te de bulunmaz |
| Belge düzenlemede: dipnot/sonnot başvurusu, yorum, denklem, içerik denetimi, izlenen değişiklik, metin kutusu ve metinli şekil, grafik, diyagram, gömülü nesne (OLE), VML resim, `mc:AlternateContent`, birden çok paragrafa yayılan alan (içindekiler tablosu), akıllı etiket ve özel XML'in panoyla taşınması (kopyada bunların yalnız metni gider; bunları içeren seçim **kesilemez**, bildirimle yerinde bırakılır), başka sekmedeki/sayfadaki bir düzenleyiciye ya da başka uygulamaya zengin içerik taşıma (pano dışarıya yalnız metin ve basit HTML verir), dışarıdan gelen HTML'in biçimiyle yapıştırılması (yalnız metni alınır), tablo hücrelerinin üzerine hücre hücre yapıştırma (tablo, iç içe tablo olarak eklenir), değişiklik izleme ve yorum düzenleme, metin kutularının içini düzenleme | Yok. Görsel, bağlantı, alan, sayfa/sütun sonu, tablo, paragraf biçimi, stil ve liste aynı sayfadaki düzenleyiciler arasında taşınır (bkz. Pano) |
| İzlenen değişikliklerde: modelin yalnız XML olarak sakladığı türleri kabul/ret etme (eklenen, silinen ve birleştirilen **hücreler** `w:cellIns`/`w:cellDel`/`w:cellMerge`, tablo ızgarası değişikliği `w:tblGridChange`, liste numarası değişikliği `w:numberingChange`, metin kutusu, `mc:AlternateContent`, yorum ve tanınmayan öğelerin içindeki değişiklikler — listede `other` türüyle görünür, "tümü" bunları atlar ve sayar); bir alanın yalnız bir bölümünü götürecek kabul/ret (reddedilir); belge sayfasında sağ tık / uzun basma menüsü; kabul/ret ve gezinme için klavye kısayolu; kenar çubuğu ve balonlar; silinmiş metnin üzerinde fare ipucu; "yalnız son hali göster" görünümü; **izleme açıkken** biçim (kalın, renk, yazı tipi…), paragraf biçimi ve stil, liste, tablo ekleme ve satır/sütun ekle-sil, görsel ve bağlantı ekleme, blok ekleme (`insertBlock`), biçimli/görselli/tablolu yapıştırma, tabloya girip çıkan ya da tablo kapsayan seçimi silme | Yok. İzleme açık belgede bu işlemler **reddedilir** (bildirimle; belge değişmez), çünkü izlenmeden yapılmaları belgenin ayarına aykırı olurdu. Zengin kopya, metni izlenen ekleme olarak yapıştırılır ve bu bildirilir. Yazma, silme, Enter, sayfa sonu, kesme ve düz metin yapıştırma izlenir (bkz. "Belge koruması ve değişiklik izleme") |
| Belge korumasında: parola doğrulama, korumayı kaldırma ya da ekleme; yorum ekleme ve form alanı doldurma (bu korumalar salt okunur açılır); korumadan muaf aralıkların (`w:permStart`/`w:permEnd`) düzenlenebilmesi; `w:writeProtection` (salt okunur önerisi) ve biçimlendirme kısıtlamasının (`w:formatting`; `w:edit` taşımayan koruma hiçbir şeyi kilitlemez) uygulanması; doc'ta salt okunur korumanın okunması | Yok. Koruma öğesi dosyada bayt bayt korunur; muaf aralık varlığı `protection.exceptions` ile bildirilir, belge yine tümüyle salt okunurdur |
| Fontlarda: Bookman Old Style dışındaki aileler için paketle gelen font (yalnız TeX Gyre Bonum gelir; öbür ailelerde cihazdakiler ve `registerFont` ile verdikleriniz kullanılır), belgeye gömülü fontların (docx `word/fonts/*.odttf`) yüklenmesi, Tahoma/Verdana/Garamond/Palatino Linotype için ölçü uyumlu yedek, Bookman Old Style ve tabloda olmayan öbür aileler için Word'ün satır yüksekliği, `FontFace` olmayan tarayıcıda font kaydı | Yok. Bkz. "Fontlar" |
| Belgede: belge içi bağlantılar, arama, sağdan sola yazı, heceleme, grafik ve vektör görsel (EMF/WMF/SVG) çizimi, belgeye gömülü fontlar | Yok. Çizilemeyen nesnelerin yerinde yer tutucu görünür |
| Belgede şekiller (docx, doc, RTF): dikdörtgen, yuvarlatılmış dikdörtgen, elips ve düz çizgi dışındaki geometriler (oklar, serbest biçimler, akış şeması şekilleri…); ok başları; çift/üçlü (bileşik) çizgi; özel kesik çizgi; gradyan, desen ve resim dolgu; saydamlık; gölge ve diğer efektler; şekil gruplarının ve çizim tuvalinin parçaları; metin kutusunun metniyle birlikte döndürülmesi | Yok. **Tek şekilli çizimde** (grup değil) çizilenler: hazır geometri `rect` (kesik/tek köşesi yuvarlak dikdörtgen türleri ve `flowChartProcess` de düz dikdörtgen olarak), `roundRect`, `ellipse`, `line`, `straightConnector1`; düz renk dolgu; çizginin rengi, kalınlığı (0 = en ince çizgi) ve hazır kesiklik türü; metinsiz şekilde döndürme, çizgide çevirme. **Yaklaşık çizilenler:** ok başlı çizgi başsız, bileşik çizgi tek, özel kesiklik düz çizilir; gölge çizilmez; kısmi saydam renk opak boyanır (tam saydam boyanmaz); gradyan/desen olan *çizgi* hiç çizilmez; metinli şekil döndürülmez. **Yer tutucu olanlar:** metinsiz şekil başka bir geometrideyse (ya da docx'te özel geometri taşıyor ya da geometri belirtmiyorsa) ya da dolgusu gradyan/desen/resimse; metinsiz gruplar, tuvaller, grafikler, diyagramlar. Bu durumlardaki **metin kutusunda** metin gösterilir: tanınmayan geometride docx'te çerçevesiz, doc/RTF'te dikdörtgen çerçeveyle; gösterilemeyen dolguda dolgusuz. **Gruplarda** yalnız parçaların metinleri (alt alta) ve resimleri gösterilir, parçaların geometrisi çizilmez. doc/RTF'te ayrıca: palet/şema/sistem dizinli renkler çözülmez (dolgu boyanmaz, çizgi siyah olur), çizim grubunun varsayılan özellik tablosu okunmaz (şekil kendi belirtmediği özellikte belirtimin varsayılanını alır), ok başı, bileşik çizgi ve gölge okunmaz; resmi bulunamayan resim şekli yer tutucudur. docx'te tüm bunlar dosyada korunur |
| Görünümde: döndürülmüş metin, sığdırmak için küçültme, iki yana yaslama, sağdan sola sayfa, yakınlaştırma, çapraz kenarlık, desenli/gradyan dolgu, başlıkları gizleme, görsel ve grafikler | Yok; içerik düz çizilir ya da atlanır |
| Sayı biçimlerinde: Türkçe ve İngilizce dışındaki diller, `*` ile doldurma, `_` boşluğunun gerçek genişliği, Hicri/Japon takvimleri, `[DBNum]` | Yok; `*` hiçbir şey yazmaz, `_x` tek boşluk olur |
| Model | doc ve RTF `document()` ile docx'le aynı belge modeline, xls `workbook()` ile aynı tablo modeline okunur. docx modeli `BymOffice.docOps` işlemleriyle düzenlenir ve `file.save()` ile dosyaya yazılır; doc ve RTF modeli yalnız `saveAsDocx()` ile docx olarak, xls modeli yalnız `saveAsXlsx()` ile xlsx olarak yazılır |
| doc/xls biçimine yazma | Yok; `save()` özgün baytları döndürür (`saveAsDocx()`/`saveAsXlsx()` yeni biçime çevirir) |
| Parola korumalı dosyalar | Yok; `UnsupportedFileError` fırlatılır |
| Excel'de grafik ve pivot tablo gösterimi (yerinde yer tutucu görünür); görsel/şekil ekleme, taşıma, boyutlandırma; dikdörtgen, elips ve çizgi dışındaki şekil geometrileri; gradyan/desen dolgu, gölge | Yok; dosyada korunur |
| Hücre notlarında: iş parçacıklı yorumu (Excel 365 "yorum"; yazışma) düzenleme, silme ve yanıtlama; yeni yazışma başlatma; not kutusunu taşıma/boyutlandırma ve "hep göster" ayarını değiştirme; notta zengin metin (kalın, renk); yorum zincirinin kopyalanması ve satırıyla sıralanması; editör araç çubuğunda not düğmesi; köprüde not komutu; xls'e yazma | Yok. İş parçacıklı yorumu olan hücrede not komutları nedeniyle birlikte devre dışıdır ve yorum dosyada bayt bayt kalır. Biçimli bir notun metni düzenlenirse not **tek biçimli düz metin** olur (düzenlenmeyen notların biçimi korunur). Yorum zinciri içeren aralık sıralanmaz; kopyalanan/kesilen hücrenin yorum zinciri yerinde kalır. Hücre içeriğini silmek (Delete) notu silmez |
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
| `dist/fonts/tex-gyre-bonum/` | Paketle gelen font (TeX Gyre Bonum, dört `.otf`) ve lisansı; betikle **birlikte** sunulur (pakette `bymoffice/fonts/…`). Bkz. "Fontlar" |

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

ES modülü kendi adresini bilmez; paketle gelen fontu (bkz. "Fontlar", 4. madde) kullanacaksanız klasörünü bir kez bildirin: `fonts.configure({ baseUrl: new URL('./dist/fonts/', import.meta.url).href })` ya da `mount(el, { fontsUrl })`.

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
| `author` | Tabloda: bu görünümde **eklenen notların** yazarı (dosyaya `<author>` olarak yazılır). Verilmezse `"Author"`; 255 karakterden uzunsa kesilir. Düzenlenen not kendi yazarını korur. Belgede: izlenen değişikliklerin yazarı (bkz. "Belge koruması ve değişiklik izleme") |
| `onNotice({ message, kind })` | Kullanıcıya gösterilen hata ve bilgi iletileri (arayüz dilinde) |
| `sheet` | İlk gösterilecek sayfanın dizini. Verilmezse dosyanın etkin sayfası; o gizliyse ilk görünür sayfa |
| `readOnly` | `true` ise düzenleme kapalıdır (editör, yapıştırma, kesme, silme, geri al, satır/sütun komutları çalışmaz); kopyalama çalışır. Varsayılan `false` |
| `onChange({ sheets })` | Çalışma kitabı her değiştiğinde çağrılır (geri al ve yinele dahil); `sheets` etkilenmiş olabilecek sayfa dizinleridir |
| `fonts` | Dosya gösterilmeden önce kaydedilecek fontlar: `registerFont()`'a verilen nesnelerin dizisi (bkz. "Fontlar"). Yüklenmeleri en çok 3 saniye beklenir, böylece dosya baştan bu fontlarla dizilir. Yüklenemeyen font dosyanın açılmasını engellemez: `onNotice` ile bildirilir (`kind: 'error'`); süre dolduğunda hâlâ yüklenen font da bildirilir (`kind: 'info'`) ve geldiğinde görünüm kendiliğinden yeniden dizilir. Fontlar görünüm kaldırılınca kayıtlı kalır. `false`: paketle gelen fontu kapatır (`fontsUrl: null` ile aynı) |
| `fontsUrl` | Paketle gelen fontun (TeX Gyre Bonum; bkz. "Fontlar", 4. madde) klasörünün adresi: paketin `dist/fonts/` klasörü nereye konduysa orası. Verilmezse betiğin yanındaki `fonts/` klasörü kullanılır. Motoru ES modülü olarak (`import`) ya da bir paketleyiciyle (Vite, webpack…) kullanıyorsanız **vermeniz gerekir**. `null` fontu kapatır. Ayar sayfa genelindedir (`BymOffice.fonts.configure({ baseUrl })` ile aynı şey) ve görünüm kaldırılınca kalır |
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
| `clearSelection()` | Seçili aralığın içeriğini siler, biçimi korur (notlar kalır) |
| `setComment(row, col, text)`, `editComment(row?, col?)`, `isEditingComment` | Hücrenin notunu yazar ya da (metin `undefined`/boşsa) siler; not alanını hücrenin yanında açar. Aşağıda "Hücre notları" |
| `setValidation(range, rule \| undefined)`, `validationAt(row, col)`, `openValidationDialog()` | Aralığın veri doğrulama kuralını koyar/kaldırır; hücrenin şu an geçerli kuralını verir; iletişim kutusunu açar. Aşağıda "Veri doğrulama kuralları" |
| `conditionalRules(range?)`, `addConditionalRule(range, rule)`, `updateConditionalRule(index, patch)`, `removeConditionalRule(index)`, `moveConditionalRule(index, to)`, `clearConditionalRules(range?)`, `openConditionalDialog(what?)` | Koşullu biçim kurallarını listeler, ekler, değiştirir, siler, sıralar, temizler; kural iletişim kutusunu ya da yöneticiyi açar. Aşağıda "Koşullu biçim kuralları" |
| `isDialogOpen`, `ruleStrings` | Görünümün üstünde bir kural iletişim kutusu açık mı; iletişim kutusu ve menü metinleri (görünümün dilinde) |
| `undo()`, `redo()`, `history.canUndo`, `history.canRedo` | Geri al ve yinele; seçim ve sayfa da geri döner. Hücre düzenlenirken çağrı yalnız düzenlemeyi **bırakır** (`true`). **Not ya da sekme adı yazılırken** çağrı yazılanı atmaz: önce **uygular** ve o çağrıda başka bir şey yapmaz (`true`; sonraki çağrı onu geri alır); uygulanamıyorsa alan açık kalır ve `false` döner. Araç çubuğundaki Geri al/Yinele bu üç durumda da devre dışıdır (alanın kendi Ctrl+Z'si alanın içinde çalışır) |
| `commitPending()` | Görünümde yazılmakta olan her şeyi uygular (hücre düzenleyicisi, not alanı, sekme adı alanı); hiçbiri açık kalmadıysa `true`. Hücre girişi reddedilirse `commitEditing()` gibi fırlatır; not/ad uygulanamıyorsa bildirir ve `false` döner. **Açık kural kutusu yazılmakta olan alan değildir** (Tamam/Vazgeç'i vardır): içinde hiçbir alan değiştirilmemişse Vazgeç gibi kapatılır (`true`); bir alan değiştirilmişse kutu açık kalır, **hiçbir şey uygulanmaz**, neden kutuda ve `onNotice` ile bildirilir ("Önce açık kural kutusunu onaylayın ya da kapatın") ve `false` döner. `mounted.save()`, yazdırma, köprü ve görünümün çalışma kitabını değiştiren her yöntemi bunu kendisi çağırır |
| `insertSheet(at?, name?)`, `renameSheet(index, name)`, `deleteSheet(index)`, `startRename(index)`, `isRenaming` | Sayfa ekler, yeniden adlandırır, siler; sekmede ad alanını açar. Aşağıda "Sayfa ekleme, silme, yeniden adlandırma" |
| `moveSheet(from, to)`, `setSheetState(index, state)`, `copySheet(index, { at?, name? })`, `setTabColor(index, color?)`, `hiddenSheets` | Sayfayı taşır, gizler/gösterir, kopyalar, sekmesini renklendirir; gizli sayfaların dizinleri. Aşağıda "Sayfa taşıma, gizleme, kopyalama, sekme rengi" |
| `cycleReference()` | F4'ün yaptığı: yazılmakta olan formülde imlecin durduğu (ya da seçimin değdiği) başvuruları bir sonraki hale geçirir (`A1` → `$A$1` → `A$1` → `$A1`). Değiştirecek başvuru yoksa ya da formül yazılmıyorsa `false` döner |
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
  - **F4 — mutlak/göreli başvuru:** formül yazılırken (hücre içi düzenleyicide ve formül çubuğunda) F4, imlecin içinde, başında ya da hemen sonunda durduğu başvuruyu Excel sırasıyla dolaştırır: `A1` → `$A$1` → `A$1` → `$A1` → `A1`. Aralığın (`A1:B2`) iki ucu birlikte değişir (sıra ilk uca göre belirlenir); tüm sütun/satır başvurusunun iki hali vardır (`A:A` ↔ `$A:$A`, `1:1` ↔ `$1:$1`); sayfa adı (`Sayfa2!`, `'Satış 2026'!`) olduğu gibi kalır; başvuru büyük harfle yazılır. Metin seçiliyse seçimin değdiği **bütün** başvurular, her biri kendi halinden bir sonrakine geçer ve seçim korunur. İmleçle basıldığında imleç başvurunun hemen sonunda kalır: yinelenen F4 dolaşmayı sürdürür, yazmaya devam edilince başvurunun arkasına yazılır. Gösterilerek eklenen başvuru F4'ten sonra da "son gösterilen" olarak kalır: sonraki tıklama onun yerine geçer; başka bir başvuru değiştirildiyse iz temizlenir. Metin formül değilse, imleç bir başvuruda değilse (tırnak içindeki metin, fonksiyon adı — `LOG10(` — başvuru değildir) ya da Shift/Ctrl/Alt basılıysa tuş hiçbir şey yapmaz ve tarayıcıya bırakılır. Formül çubuğunda, henüz düzenlenmeyen bir formülde F4 düzenlemeyi başlatır. Tarayıcının metin geri alması (Ctrl+Z) değişikliği geri alır (WebKit yazılanların tamamını tek adım sayar). **macOS:** dizüstü klavyede F4, Fn+F4'tür (ya da Sistem Ayarları'nda "F1, F2… tuşlarını standart işlev tuşları olarak kullan"); Excel'in Cmd+T'si desteklenmez. Dokunmatik ekran klavyesinde F4 yoktur; `view.cycleReference()` aynı işi yapar.
  - **Formül çubuğundan:** formül çubukta yazılırken ızgaraya ya da sekmeye tıklamak odağı çubuktan almaz; başvuru çubuktaki imlecin yerine eklenir.
  - **Düzenleyicide Ctrl/Cmd+Z** metin alanının kendi geri almasıdır: Chromium ve Firefox'ta eklenen başvuruyu tek adımda geri alır (sürüklemede aralık her değiştiğinde bir adım oluşur; aynı hücrenin içindeki hareket adım üretmez ve `onEditingChange` her yazımda bir kez çağrılır); WebKit/Safari bir seferde yazılan her şeyi tek adım saydığından orada başvuruyla birlikte ondan önce yazılanlar da gider (giriş açık kalır). Sürükleme sırasında bir şey yazılırsa sürükleme biter.
- **Biçimlendirme:** araç çubuğu ya da Ctrl/Cmd+B, I, U.
- **Sıralama ve süzgeç:** araç çubuğundaki A→Z, Z→A ve Süzgeç düğmeleri. Tek hücre seçiliyken çevresindeki dolu bölge alınır, ilk satır başlık sayılır. Süzgeç açıkken başlık hücrelerindeki düğme (ya da Alt+Aşağı ok) menüyü açar: sırala, ara, değerleri seç, süzgeci temizle. Durum çubuğu "N kayıttan M tanesi" gösterir. Süzgecin veri satırları **gerçekten hücre taşıyan son satırda** biter (`lastFilledRow`): tam sütun seçilerek açılan süzgeç oraya kadar kesilir; hücreleri temizlenmiş alt satırlar listenin satırı sayılmaz, gizlenmez ve sayılmaz (`sheet.rowCount` temizlemeyle küçülmediği için ölçü alınmaz — aynı komut, çalışma kitabında ve ondan kaydedilip yeniden açılan dosyada aynı satırları gizler). Yalnız biçim taşıyan hücre de hücredir; hücresi olmayan (yalnız yüksekliği ayarlanmış) satır değildir. Sıralama satırları formülleri ve biçimleriyle taşır (satır yükseklikleri taşınmaz); notlar satırlarıyla birlikte gider; birleşik hücre, çok hücreli dizi formülü ya da yorum zinciri içeren aralık ve tablolu sayfa sıralanmaz. Kaydederken değişmeyen süzgeç bayt bayt korunur; değişen süzgeç Excel'in beklediği biçimde yazılır. API: `view.sort(descending)`, `view.toggleFilter()`, `view.setColumnFilter(col, choice)`. Süzgeçli listede temizle, kopyala/kes, biçimlendir ve satır sil yalnız GÖRÜNEN satırlara uygulanır (Excel gibi); yapıştırma gizli satırlara da yazar; tam sütun biçimlendirmesi gizli satırları da etkiler. Ölçüt sütunu silinince o ölçütün gizlediği satırlar geri gelir; başlık satırı silinince süzgeç kalkar. Sıralamada aksanlı harfler taban harfiyle birlikte sıralanır, başka sayfaya giden başvurular kaymaz; tek hücreyle sıralamada başlık satırı türe ya da kalınlığa göre sezilir. Süzgeçte ve listelerde ı/i ayrımı yerele göredir.
- **Veri doğrulama:** dosyadaki kurallar uygulanır. "Dur" türündeki kurala uymayan değer yazılamaz ve kuralın iletisi gösterilir; "uyarı" ve "bilgi" türünde değer yazılır ve bildirim verilir. Liste kurallı hücrede açılır ok (Alt+Aşağı ok) değerleri gösterir; hücre seçilince giriş iletisi balonda görünür. Yapıştırma doğrulanmaz (Excel de doğrulamaz). Kurallar eklenir, değiştirilir ve silinir: aşağıda "Veri doğrulama kuralları". Formüllü kurallar, yazılan değere bağlı formüller hesaplanmış haliyle değerlendirilir; reddedilen giriş hiçbir iz bırakmaz (dosya bayt bayt aynı kalır). Motorun hesaplayamadığı kural (desteklenmeyen fonksiyon) uygulanmaz. **Formül olarak yazılan giriş** (`=B2`) metniyle değil, yazıldığı andaki **sonucuyla** denetlenir: "1–5 tam sayı" kurallı hücreye `=B2` yazılır ve B2 9 ise giriş reddedilir; B2 sonradan değişirse hücre yeniden denetlenmez. Türkçe yazılan formül (`=TOPLA(B2;1)`) saklandığı haliyle hesaplanır. Sonucu burada hesaplanamayan formül (desteklenmeyen fonksiyon, başka dosyaya ya da 3B başvuru, tanımsız ad), sonucu hata değeri olan formül (`#DIV/0!`; Excel'in bunu reddedip reddetmediği doğrulanmadı) ve `calculate: false` ile açılmış görünümdeki her formül denetlenmeden kabul edilir. İki durum daha denetlenmez: **kendi hücresini okuyan formül** (`A1`'e `=A1+1`, `=TOPLA(A:A)`; hücrenin eski değeriyle yargılanmasın diye — Excel bunu döngüsel başvuru sayar) ve o sırada süren hesap bitirilemediyse yazılan formül (görünüm önce hesabı bitirmeyi dener; okunan hücreler bayat değer taşıyabilir). **Hata veren kural hücreyi kilitlemez:** kuralın sınırı (`formula1`/`formula2`) hata değeri veriyorsa (silinen sayfaya bakan sınır `#REF!` olur) kural uygulanmaz ve giriş kabul edilir; kaynağı hata olan liste **boş** sayılır (`list()` → `{ items: [] }`, açılır ok gösterilmez) ve her giriş kabul edilir; özel (custom) formül `#REF!` veriyorsa yine kabul edilir. Özel formülün **başka** bir hata vermesi (`=KAÇINCI(A1;liste;0)` bulamayınca `#YOK`) eskisi gibi rettir: Excel hatayı YANLIŞ sayar. Kurallardaki ve yazılan formüldeki sayfa adları motorun karşılaştırdığı gibi karşılaştırılır (`KASIM!A1`, `Kasım` sayfasının hücresidir).
- **Sağ tık menüsü** (dokunmatikte uzun basma): kes, kopyala, yapıştır, içeriği temizle, satır/sütun ekle-sil. Menüden "Yapıştır" tarayıcı izni ister ve her zaman düz metin yapıştırır; biçim ve formülleriyle yapıştırmak için Ctrl/Cmd+V kullanın.
- **Boyutlandırma:** sütun ya da satır başlığının kenarını sürükleyin (Esc vazgeçer); kenara çift tıklamak sütunu içeriğine sığdırır. Sıfıra sürüklemek gizler.
- **Satır ve sütun:** satır ya da sütun başlığına sağ tıklayın (dokunmatikte uzun basın); açılan menüden ekleyin ya da silin. Tam satır ya da sütun seçiliyken Ctrl/Cmd+`+` ekler, Ctrl/Cmd+`-` siler.
- **Sayfalar:** sekmelerin solundaki **+** düğmesi sona yeni sayfa ekler ve onu gösterir. Sekmeye sağ tıklamak (dokunmatikte uzun basmak; klavyede odaktaki sekmede Shift+F10 ya da Menü tuşu) menüyü açar: "Sayfa ekle" (o sekmenin sağına), "Yeniden adlandır", "Sil". Sekmeye çift tıklamak/çift dokunmak ya da odaktaki sekmede F2 adı yerinde düzenletir: Enter ya da alandan çıkmak uygular, Esc vazgeçer. Silmede onay sorulmaz; Ctrl/Cmd+Z geri getirir. Salt okunur görünümde düğme, menü ve ad alanı yoktur. Ayrıntı aşağıda.
- **Klavye** (önce görünüme tıklayın ya da Tab ile odaklanın): oklar, Shift+oklar, Ctrl/Cmd+oklar (veri bloğunun kenarı), Home, Ctrl+Home, Ctrl+End, PageUp/PageDown, Enter/Shift+Enter, Ctrl+A; sayfa değiştirmek için Alt+PageDown/PageUp (Ctrl+PageDown/PageUp da çalışır, ama bazı tarayıcılar onu kendine ayırır). Tab ve Shift+Tab hücre değiştirmez: odağı ızgaradan sayfadaki sonraki/önceki öğeye taşır, böylece klavye kullanıcıları görünümde hapsolmaz. Sayfa sekmeleri Tab ile odaklanıp Enter ya da Boşluk ile seçilir; Tab sırası: ızgara → **+** düğmesi → sekmeler.

Bilinmesi gerekenler:

- Hücreler dosyadaki sayı biçimiyle gösterilir. `[Red]` gibi renk bölümleri yazı rengini değiştirir; biçimiyle gösterilemeyen değer (ör. negatif tarih) ve sütuna sığmayan sayı `####` olarak çizilir.
- Çizimde hücrenin kendi fontu kullanılır; cihazda yoksa `Carlito`, `Calibri`, `Arial` sırasıyla yedeğe düşülür (Bookman Old Style için önce paketle gelen font yüklenir; bkz. "Fontlar"). Font farkı metin genişliklerini biraz değiştirebilir.
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

- **Satır sonları ve bayt sırası işareti korunur:** sayfa parçası BOM (`EF BB BF`) ya da CRLF satır sonlarıyla yazılmışsa (elle düzenlenmiş, başka araçtan çıkmış dosyalar) kaydetme parçayı baştan sona LF'e çevirmez: yalnız değişen öğenin (eklenen/değişen/silinen `<conditionalFormatting>`, `<dataValidation>`, adı değişen sayfayı anan formül, yeniden yazılan `<sheetData>`…) baytları değişir, öğe içinde değişmeyen bölümler ve öğeler arasındaki her bayt aynı kalır. Not ve sekme rengi yazımı zaten böyleydi.
- **Yalnız değişen yeniden yazılır.** Değiştirmediğiniz sayfalar, görseller, tema, makrolar ve motorun tanımadığı her parça özgün baytlarıyla kalır. Değişen bir sayfada bile dokunmadığınız satırlar özgün XML'iyle, bayt bayt aynen kopyalanır.
- **Hiçbir şey değişmediyse** kaydedilen dosya açılanla bayt bayt aynıdır. **Değiştirilip geri alınan** dosya da öyledir, arada kaydedilmiş olsa bile: "sayfa ekle (ya da sil) → kaydet → geri al → kaydet" özgün dosyayı bayt bayt verir. İçeriği dosyadaki haline dönen her ZIP girdisi özgün girdisiyle (sıkıştırılmış baytları, ek alanları, arşivdeki yeri) yazılır; içerik türleri listesi (`[Content_Types].xml`) dosyadakiyle aynı türleri bildiriyorsa, sırası ne olursa olsun, özgün baytlarıyla yazılır.
- **Etkin sayfa:** görünüm, kullanıcının gösterdiği sayfayı modele yazar (`workbook.activeSheet`); kaydedilen dosya o sayfayla açılır (`activeTab` ve iki sayfanın `tabSelected` işareti güncellenir: eski ve yeni etkin sayfa) — ama **yalnız başka bir değişiklikle birlikte** (hücre, biçim, sayfa ekleme/silme/yeniden adlandırma). Yalnız sekme değiştirip kaydetmek dosyayı değiştirmez. Formül yazarken başvuru göstermek için bakılan sayfa ve `setActiveSheet()` ile gösterilen gizli sayfa etkin sayfa olmaz. **Birden çok penceresi olan dosyada** (`<bookViews>` içinde birden çok `workbookView`) yalnız sekme değişikliği dosyaya **yazılmaz** (her pencerenin kendi etkin sayfası vardır); sayfa eklenip silindiğinde ise bütün pencereler etkin sayfayla açılır. Grafik sayfası etkin sayfa olabilir: parçasının yalnız `<sheetViews>` öğesi düzeltilir, başka hiçbir baytı değişmez (UTF-16 kodlu grafik sayfası parçasına dokunulmaz). `activeSheet` gizli bir sayfayı gösteriyorsa dosya ilk görünür sayfayla açılır. Görünüm olmadan `workbook.activeSheet` alanını kendiniz de değiştirebilirsiniz; aynı kural geçerlidir.
- Sayfanın modelden yazılmayan bölümleri (koşullu biçimler, veri doğrulama, sayfa yapısı/yazdırma ayarları ve yazıcı ayarı parçası, uzantılar) özgün baytlarıyla aynen korunur; satır/sütun eklenip silindiyse içlerindeki hücre başvuruları, resimlerin konumları ve yorumlar birlikte kayar.
- Yeni biçimler stil tablosuna yalnız **eklenir**; var olan stil kayıtları değişmez.
- Hesaplanan formül sonuçları dosyaya yazılır. Yine de dosyaya "açılışta yeniden hesapla" işareti konur ve eski hesap zinciri (`calcChain`) kaldırılır; Excel dosyayı açınca tüm sonuçları kendi motoruyla günceller.
- Düzenlediğiniz metin hücreleri satır içi metin (`inlineStr`) olarak yazılır; dokunulmayan hücreler paylaşılan metin tablosundaki yerlerini korur.
- Aynı çalışma kitabı art arda kaydedilebilir; kaydetmeden sonra geri al/yinele çalışmayı sürdürür.
- **Yazılmakta olan da kaydedilir:** `mounted.save()` (ve Ctrl/Cmd+S, köprünün `save` iletisi) önce `view.commitPending()` çağırır: açık hücre düzenleyicisi, **not alanı** ve **sekme adı alanı** uygulanır, sonra dosya yazılır. Hücre girişi reddedilirse (doğrulama) `save()` o hatayla reddedilir ve düzenleyici açık kalır (eskisi gibi). Not ya da ad uygulanamıyorsa (başka sayfanın adı, geçersiz ad) alan yazılanla açık kalır, neden `onNotice` ve canlı bölgeyle bildirilir, `commitPending()` `false` döner ve `save()` **reddedilir** (hiçbir şey kaydedilmez). **Açık kural kutusu** (veri doğrulama, koşullu biçim, kural yöneticisi) kaydetmeyle **uygulanmaz**: hiçbir alanı değiştirilmemişse kapatılır ve kaydedilir; değiştirilmişse kutu açık kalır, `save()` reddedilir ("… a dialog of rules is open with changes (confirm or close it first)") ve kullanıcıya "Önce açık kural kutusunu onaylayın ya da kapatın" denir — otomatik kaydetme, kullanıcının onaylamadığı kuralı dosyaya yazmaz. `file.save()` / `file.saveAsXlsx()` görünümü bilmez: bunları doğrudan çağırıyorsanız önce `view.commitPending()` çağırın. Yazdırma (`print()`) da aynı yöntemi çağırır; uygulanamayan ad/not varken `'cancelled'` ile reddedilir. Not ve ad alanında Ctrl/Cmd+S artık yukarı bırakılır (önceden yutuluyor, tarayıcının "sayfayı kaydet" penceresi açılıyordu).
- **`sheet` seçeneği:** görünüm `mount({ sheet: 2 })` ile belli bir sayfada açıldıysa o sayfa modele etkin sayfa olarak yazılır (dosyaya yine yalnız başka bir değişiklikle birlikte taşınır). Seçenek verilmediyse modelin etkin sayfasına dokunulmaz.
- **UTF-16 parçalar:** UTF-16 kodlu bir parça (`workbook.xml`, `*.rels`, `[Content_Types].xml`) yeniden yazılınca UTF-8 yazılır ve bildirimindeki kodlama da `UTF-8` olur.
- **`calcChain`:** hesap zinciri yalnız **dosyadaki** bir sayfa değişince kalıcı olarak kaldırılır; yalnız sayfa eklendiyse/silindiyse geçici çıkarılır ve geri alınınca özgün baytlarıyla döner.

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
  - **"eklenmiş, sonra silinmiş" metin** (RTF'te `\revised` + `\deleted`, doc'ta `sprmCFRMark` + `sprmCFRMarkDel` birlikte) docx'teki gibi **eklemenin içinde silme** olarak okunur (`<w:ins Ann>… <w:del Bob>…</w:del> …</w:ins>`; yazar ve tarihler ayrı ayrı): "tümünü reddet" metni çıkarır, eklemeyi kabul edip silmeyi reddetmek geri getirir (önceden yalnız silme olarak okunuyor, ret metni kalıcı yapıyordu). Aynı yazarın aynı andaki, **araya düz metin girmiş** iki değişikliği iki ayrı kayıttır ve biri çözülünce öbürü etkilenmez.
  - **izlenen değişiklikler** yazarı ve tarihiyle taşınır: RTF'te yazar `\revtbl` tablosundan (`\revauthN`, `\revauthdelN`), tarih `\revdttmN`/`\revdttmdelN` DTTM değerinden; doc'ta `SttbfRMark` tablosundan ve `sprmCIbstRMark`/`sprmCDttmRMark` (silme için `…Del`) özelliklerinden okunur. Tarih olmayan bir DTTM (saat 24 ve üstü, ayda olmayan gün…) tarihsiz bırakılır; DTTM yerel saattir ve Word'ün yaptığı gibi `Z` ile yazılır (UNSURE). Silinmiş/eklenmiş **paragraf işareti** (`{\deleted\par}`; doc'ta işaret karakterinin özellikleri) ve silinmiş **alan** (işaretleri `w:del` içinde, kodu `w:delInstrText`) de böyle taşınır. Yeni docx'te her `w:ins`/`w:del` öğesi ve her işaret revizyonu, belgedeki öbür kimliklerle çakışmayan sıralı bir `w:id` alır; yazarı okunamayan revizyon `"Author"` adıyla yazılır ve bu tek satırda sayısıyla bildirilir (`N tracked change(s) were carried over without their authors…`); tarihi olmayan revizyona `w:date` yazılmaz. Okunmayanlar: biçim revizyonları (`\revprop`, `\crauthN`, `\crdateN`; doc'ta `sprmCPropRMark` vb.), silinmiş tablo satırı (hücre işaretinin revizyonu); hem eklenmiş hem silinmiş metin yalnız silinmiş olarak gösterilir.
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
- Seçenekler: `zoom` (sayı, `'fit-width'`, `'fit-page'`), `toolbar` (sayfa göstergesi ve yakınlaştırma çubuğu), `readOnly`, `trackChanges`, `author`, `now` (bkz. "Belge koruması ve değişiklik izleme"), `onDocumentChange()` (belge her değiştiğinde, geri al ve yinele dahil), `onPageChange(page, count)`, `onLinkClick(target)` (`false` dönerse bağlantı açılmaz; yalnız http, https, mailto, tel hedefleri bağlantı olur), `locale`, `onNotice`, `onSave`.
- `DocumentView`: `pageCount`, `currentPage`, `layoutDone`, `zoom`, `setZoom()`, `zoomIn()`, `zoomOut()`, `goToPage(n)`, `selectAll()`, `getSelectedText()`, `setLocale()`, `focus()`, `print(options?)`, `printing`, `destroy()`.
- İzlenen değişiklikler: `revisions()`, `goToRevision()`, `acceptRevision()`, `rejectRevision()`, `acceptAllRevisions()`, `rejectAllRevisions()`, `setTrackChanges()` (bkz. "Belge koruması ve değişiklik izleme").
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
  - **İzlenen değişiklikler kopyada hiç taşınmaz** (aynı belgeye yapıştırmada da): yapıştırılan içerik düz içeriktir. İzlenen eklemenin metni metin olarak gider, izlenen silmenin metni gitmez; paragraf işaretindeki ekleme/silme/taşıma işareti düz işaret olur (silinmiş paragraf işareti kopyada iki paragrafı **ayrı** bırakır); biçim değişikliği kayıtları (`w:pPrChange`, `w:rPrChange`, `w:sectPrChange`, `w:tblPrChange`, `w:tblPrExChange`, `w:trPrChange`, `w:tcPrChange`, `w:tblGridChange`, `w:numberingChange`) ve satır/hücre ekleme-silme-birleştirme işaretleri atılır, biçimin bugünkü hali taşınır. Böylece hedefte yinelenen revizyon kimliği ve kaynağın yazar adı oluşmaz. Kes de bir kopyadır: kesilip geri yapıştırılan paragraf bu kayıtları yitirir (UNSURE: Word izleme kapalıyken korur).
  - **Çok fazla stil/liste adlayan kopya** (256'dan çok tanım) geldiği belgeye yapıştırılır; **başka** belgeye yapıştırılmaz ("Kopyalanan içerik buraya yapıştırılamıyor"): tanımı taşınamayan kimlikler hedefte başka bir stile/listeye bağlanırdı.
  - **İzleme açık belgede** kopya düz metin olarak yapıştırılır; metni olmayan kopya (yalnız görsel) yapıştırılamaz ve bu `onNotice` ile bildirilir ("…bu tür değişiklik henüz izlenemiyor; hiçbir şey değiştirilmedi").
  - **Tek geri alma adımı:** yapıştırma, eklediği medya parçası, ilişki, stil ve liste tanımlarıyla birlikte tek adımdır; geri alınınca hepsi modelden çıkar ve kaydedilen pakette yetim parça kalmaz.
  - **Bellek:** sayfa son 8 kopyayı, toplam 64 MB'a kadar (görsel baytları dahil) tutar; eskiler atılır. Paketten okunan görsel baytları kopyadan hemen sonra eşzamansız okunur; baytlar gelmeden başka belgeye yapıştırılırsa bu bildirilir ve yapıştırma baytlar gelince, **istendiği yere** yapılır: bu arada imleç taşınmış, seçim ya da belge değişmişse yapıştırma yapılmaz ve bildirilir (yeniden yapıştırmak yeter). Sınırı aşıp görsel baytları atılan kopyaya sonradan okunan baytlar da eklenmez. Sınıra sığmayan kopya görsel baytlarını yitirir (yalnız görselin zaten bulunduğu belgeye yapıştırılabilir). 5.000 paragraftan ya da 20.000 öğe/hücreden büyük seçim zengin kopyalanmaz (düz metin olarak gider) ve metin dışında bir şey içeriyorsa kesilemez. Uzun yapıştırma sınırı (5 milyon karakter) zengin kopya için de geçerlidir.
  - Yalnız görsel seçiliyken kopyalanan panoda düz metin boştur; dokunmatik balonundaki Yapıştır düğmesi panonun HTML'ini okuyabildiği tarayıcılarda (`navigator.clipboard.read`) zengin kopyayı tanır, okuyamadığında düz metni yapıştırır ya da klavyeyle yapıştırmayı ister.
  - Salt okunur görünümden kopyalanan seçim de zengin kopyadır. Paket olmayan bir dosyadan (RTF) alınan kopyanın bir docx'e görseli, alanı ve stiliyle yapıştırılıp kaydedilmesi birim testiyle sınanmıştır; ikili doc için ayrı bir test yoktur.
  - Programdan: `BymOffice.docOps.fragmentOf(doc, range)` kopyayı (`DocFragment`), `docOps.insertFragment(doc, konum | aralık, fragment, { taken? })` yapıştırma düzenlemesini (`DocEdit`) verir; `docOps.fragmentNeeds(doc, fragment)` hedef belgenin baytını beklediği görselleri sayar. `insertFragment`'e `moved: true` vermek, kopyanın kesilerek alındığını söyler (yukarıdaki "kesilenin geri konması" kuralı). `sliceOf()` dönüşündeki `DocSlice` artık `fragment` alanını da taşır.
- Çok sayıda bloğu değiştiren işlemde (ör. bütün belgeyi kalın yapmak) sayfalar açılıştaki gibi dilimler halinde yeniden dizilir; arayüz donmaz.
- Tuşlar: Ctrl/Cmd ile `+`, `-`, `0` ve tekerlek yakınlaştırır (dokunmatikte iki parmak); PageUp/PageDown/Boşluk/Home/End/oklar gezdirir; fareyle sürükleyerek ya da Ctrl/Cmd+A ile seçilen metin Ctrl/Cmd+C ile kopyalanır.
- Sayfalar aşamalı dizilir: ilk sayfalar hemen görünür, kalanı arka planda; binlerce sayfalık belgede yalnız ekrandaki sayfaların tuvali tutulur.
- **Fontlar ve sayfalama:** metin cihazdaki fontlarla ölçülür. Belgenin fontu cihazda yoksa ölçüleri uyumlu açık fontlar denenir (Calibri→Carlito, Cambria→Caladea, Arial→Liberation Sans, Times New Roman→Liberation Serif, Courier New→Liberation Mono), sonra genel aile. Bookman Old Style için yedek (TeX Gyre Bonum) paketle gelir ve kendiliğinden yüklenir; öbür fontları sayfanıza `@font-face` ile eklerseniz kendiliğinden kullanılır. Font farklıysa satır ve sayfa kırılmaları Word'dekinden farklı olabilir; Word'le yan yana karşılaştırma yapılmamıştır.
- Alt düzey: `BymOffice.createCanvasMeasurer({ fonts: doc.fonts })` → `new BymOffice.DocLayout(doc, measurer)` → `layout.layout({ pages: 5 })` → `BymOffice.renderPage(ctx, layout.pages[0], { scale })`. Sayfanın öğeleri (`page.items`) `text`, `rect`, `line`, `image`, `shape` (şekil: `geometry`, `fill`, `stroke`, `rotation`) ve `mark` türündedir. Model düzenlendikten sonra `layout.edited({ milliseconds })` yalnız değişen sayfadan itibaren yeniden dizer; `undefined` dönerse ya da sonraki `layout()` çağrılarından sonra `layout.broken` `true` olursa yeni bir `DocLayout` kurulur.

### Belge koruması ve değişiklik izleme

Belgenin kendi ayarları uygulanır; ikisi de dosyadan okunur. Koruma dosyada olduğu gibi kalır. `word/settings.xml` yalnız **izleme ayarı değiştiğinde** yazılır ve o zaman da içinde `<w:trackRevisions/>` öğesinden başka hiçbir bayt değişmez (aşağıda "İzlemeyi açıp kapatma").

**Koruma (`w:documentProtection`).** `document.settings.protection` → `{ edit, enforced, xml? }` (`edit`: `"readOnly"`, `"comments"`, `"trackedChanges"`, `"forms"`, `"none"` ya da dosyanın yazdığı başka bir değer; `xml`: parola karması öznitelikleriyle öğenin kendisi). Motor parolayı doğrulamaz, korumayı kaldırmaz. `w:enforcement` ve `w:edit` değerleri dosyaların yazdığı kadar gevşek okunur: baştaki/sondaki boşluk ve harf büyüklüğü önemsizdir (`TRUE`, `On`, `" 1 "` uygulanıyor demektir; `TrackedChanges` → `trackedChanges`, `None` → kilit yok); tanınmayan bir tür yazıldığı gibi kalır. Ayarlar parçasında birden çok `w:documentProtection` varsa **uygulanan ilk öğe** sayılır (hiçbiri uygulanmıyorsa ilk öğe); öndeki boş bir öğe kilidi düşürmez. `BymOffice.enforcedProtection(document)` uygulanmakta olan türü verir (`w:enforcement` kapalıysa, tür `none` ise ya da `w:edit` hiç yoksa — Word'ün "yalnız biçimlendirme kısıtlaması" — `undefined`; boş olmayan tanınmayan bir tür `"readOnly"` sayılır).

| Uygulanan koruma | Davranış |
|---|---|
| `readOnly`, `comments`, `forms` | Belge **salt okunur** açılır ve öyle kalır: `view.setReadOnly(false)`, `editor.setReadOnly(false)` ve köprünün `setReadOnly` komutu düzenlemeyi açmaz (yorum ekleme ve form doldurma arayüzü olmadığı için kısmi izin verilemez). Düzenleme istenmişse neden `onNotice` ile **bir kez** bildirilir |
| `trackedChanges` | Belge düzenlenir; her değişiklik izlenir ve bu kapatılamaz (`trackChanges: false` etkisizdir) |
| uygulanmıyor (`w:enforcement` yok/kapalı) | Hiçbir kısıtlama yok; öğe yalnız korunur |

`mount()` dönüşünde (belgede) `protection?: { edit, exceptions }` (koruma uygulanıyorsa) ve `trackChanges?: true` (düzenlemeler izleniyorsa) vardır; `exceptions`, belgede korumadan muaf aralık (`w:permStart`) bulunduğunu söyler (bunlar uygulanmaz). `DocumentView`: `protection`, `trackingChanges`. `w:writeProtection` (salt okunur önerisi) `settings.writeProtection` olarak saklanır, uygulanmaz. RTF'te koruma şu kuralla okunur: `\enforceprot1` varsa `\protlevelN`'in (yoksa eski sözcüğün) söylediği tür uygulanır; `\enforceprot0` ise koruma saklanır, uygulanmaz; `\enforceprot` hiç yoksa yalnız eski sözcüklerden biri (`\formprot`, `\revprot`, `\annotprot`, `\readprot`) açıkça varsa uygulanır; tek başına `\protlevelN` koruma değildir. doc'ta (DopBase) üç bitin her biri **kendi başına** uygulanan bir korumadır: `fLockRev` → `trackedChanges`, `fLockAtn` → `comments` (Dop2003'teki `fTreatLockAtnAsReadOnly` açıksa `readOnly`), `fProtEnabled` → `forms`; birden çoğu açıksa bu sırayla ilki sayılır. Komşu bitler (`fMirrorMargins`, `fEmbedFonts`, `fRevMarking`) koruma değildir. UNSURE: ikisi de belirtime göre yazıldı, Word'ün koruduğu dosyalarla karşılaştırılmadı; çakışmadaki sıra ve `fTreatLockAtnAsReadOnly`'nin yeri (Dop2002'nin 594 baytından sonraki ilk bit, yalnız 616 bayt ve üstü yapıda okunur) doğrulanmadı — yanlışsa yorum koruması salt okunur (ya da tersi) adlandırılır, ikisi de burada kilitli açılır. doc/RTF → docx dönüştürmede koruma yeni dosyaya **parolasız** yazılır ve bu `warnings`'te bildirilir; dönüştürülen belge de korumaya göre açılır.

**Değişiklik izleme.** İzleme şu durumlarda açıktır: belgenin ayarı (`w:trackRevisions`, `settings.trackRevisions`), uygulanan `trackedChanges` koruması, ya da uygulamanın isteği (`mount({ trackChanges: true })`). Uygulama, belgenin istediği izlemeyi seçenekle **kapatamaz** (`trackChanges: false` "dosyanın ayarı geçerli" demektir); kullanıcı belgenin ayarını araç çubuğundan açıp kapatabilir (aşağıda "İzlemeyi açıp kapatma"). İzleme açıkken:

- **Yazar ve tarih:** `mount({ author })` (köprüde `load.author`); verilmezse dosyaya `"Author"` yazılır (UNSURE: Word kullanıcının adını yazar; boş yazar adı yazılmaz). **Yazar adı verilmemişse belgedeki hiçbir ekleme "kendi eklemesi" sayılmaz** (aşağıdaki "gerçekten çıkar" kısayolu uygulanmaz): Word, kişisel bilgisi silinmiş belgelerde de yazarı `Author` yazar ve bunlar bizim sayılmamalıdır. Bu yüzden uygulamanın `author` vermesi önerilir. Tarih işlem anıdır, UTC, saniye çözünürlüğünde (`2026-10-05T09:08:07Z`); `mount({ now: () => Date })` sınamalar için saati verir; programdan verilen `date` ISO 8601 değilse `RangeError`. Kimlikler (`w:id`) belgedeki bütün öykülerde (gövde, üstbilgi/altbilgi, notlar, yorumlar) var olan en büyük kimliğin üstünden verilir ve geri alınan bir kimlik yeniden kullanılmaz. Bir düzenleme bir revizyonu ya da paragrafı ikiye böldüğünde ikinci yarı yeni kimlik alır — **izleme açıkken de kapalıyken de**: ortasına yazılan ya da ortasından silinen `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`, `w:rPrChange` taşıyan çalıştırma, bölünen paragrafın `w:pPrChange`'i ve işaretindeki `w:rPrChange` ile ekleme/silme işareti. Böylece yazılan dosyada aynı `w:id` iki öğede bulunmaz ve bellekteki kayıtlar dosyadakilerle aynıdır (kural her düzenlemenin belgeye koyduğu paragraflara uygulanır; tablo içindeki paragraflara yalnız doğrudan düzenlendiklerinde). **Yeni yazılan ya da yapıştırılan metin komşusunun `w:rPrChange` kaydını devralmaz** (biçimini alır; "bu metnin biçimini kim değiştirdi" kaydı o metnin değildir).
- **Ekleme** (yazma, düz metin yapıştırma, Enter, sayfa sonu): metin `w:ins` içine girer; yeni paragraf işareti `w:pPr/w:rPr/w:ins` ile işaretlenir (bölünen paragrafın **ilk** yarısının işareti yenidir). Aynı yazarın bitişik eklemesine yazmak onu uzatır (tarih ilk eklemenin kalır). Bölünen paragrafın özelliklerinde XML olarak saklanan izlenen kayıtlar (liste numarası değişikliği `w:numberingChange`, `w:numPr` içindeki `w:ins`) da iki paragrafta yinelenmez: yeni paragraftaki kopya **yeni kimlik** alır. Başkasının eklemesinin ortasına yazmak onu ikiye böler (ikinci yarı yeni bir kimlik alır); bölünecek eklemeyle yazılan yer arasında bir **içerik denetimi** (`w:sdt`) ya da `w:customXml` varsa (denetim kimliğiyle birlikte çoğalırdı) işlem reddedilir, kendi eklemesinin içindeki denetime birden çok paragraf yapıştırmak da öyle.
- **Silme** (Backspace/Delete, seçim silme, kesme, üzerine yazma): metin belgeden çıkmaz; çalıştırması `w:del` içine alınır, metni `w:delText` (alan kodu `w:delInstrText`) olur. Görsel, dipnot başvurusu, alan ve denklem de bulundukları çalıştırmayla `w:del` içine girer (dipnotun kendisi silinmez). **Aynı (adı verilmiş) yazarın izlenen eklemesinin doğrudan içindeki metin, sekme ve satır sonu** silinince gerçekten çıkar (iz kalmaz); o eklemedeki başka her şey (dipnot/sonnot başvurusu, alan — kodu ve sonucu —, görsel, denklem, simge, sayfa/sütun sonu) `w:ins` içinde `w:del` olur, dipnotun kendisi silinmez. Eklemenin içinde **başka bir sarmalayıcının** (basit alan `w:fldSimple`, bağlantı `w:hyperlink`, içerik denetimi `w:sdt`, akıllı etiket `w:smartTag`, `w:customXml`, `w:dir`/`w:bdo`) içinde duran metin de çıkmaz: sarmalayıcı yerinde kalır, metni onun içinde `w:del` olur (`<w:ins><w:fldSimple><w:del>…`; basit alanda silinen, alanın saklanan sonucudur). Tersine, bir bağlantının ya da denetimin içine yazılmış kendi eklemesi, o sarmalayıcıda başka içerik kalıyorsa gerçekten çıkar; sarmalayıcının tek içeriğiyse `w:del` olarak kalır (sarmalayıcı kaybolmasın diye). **Karar sarmalayıcı başınadır, çalıştırma başına değil:** silme sonunda sarmalayıcı (bağlantı, `w:sdt`, `w:fldSimple`, `w:smartTag`…) boş kalacaksa o sarmalayıcıdaki **bütün** kendi metni `w:del` olur; boş kalmayacaksa hepsi gerçekten çıkar. Sonuç metnin kaç çalıştırmaya bölündüğüne (ör. bir bölümü kalın) bağlı değildir. Tuş tuş silmede (art arda Backspace) bu şu demektir: her tuşta sarmalayıcıda başka kendi harfi kaldığı için harfler gerçekten çıkar, **son harf** ise sarmalayıcıyı boşaltacağı için `w:del` olarak kalır. Yani izleme açıkken belgeden yalnız yazarın kendi yazdığı metin çıkabilir, hiçbir öğe onunla birlikte gitmez. Başka yazarın eklemesindeki metin `w:ins` içinde `w:del` olur. Zaten silinmiş metin yeniden silinmez.
- **Paragraf birleştirme:** paragraf işareti `w:pPr/w:rPr/w:del` ile işaretlenir, **iki paragraf ayrı kalır** (dosyada ve ekranda; Word de böyle yazar, kabul edilince birleşir). Backspace imleci önceki paragrafın sonuna, Delete sonrakinin başına taşır. Aynı yazarın eklediği işaret silinince paragraflar **yalnız ikisinin biçimi aynıysa** (stil, liste, hizalama, girinti, işaretin biçimi, saklanan `w:pPr` içeriği; birincide bölüm sonu yoksa) gerçekten birleşir; biçimleri farklıysa işaret başkasınınki gibi `w:del` ile işaretlenir ve paragraflar ayrı kalır (ikincinin stili, listesi ve hizalaması birleşmeyle kaybolmasın diye; kesip aynı yere yapıştırınca iki paragraf kendi biçiminde kalır).
- **Görünüm:** eklenen metin altı çizili, silinen metin **üstü çizili olarak yerinde görünür**; ikisi de yazara göre renklidir (8 renkli sabit palet, yazar adının karma değeriyle; `revisionColor`). Dosyadaki mevcut `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` de aynı kuralla gösterilir. Silinmiş metin **metnin parçası değildir**: konum (ofset) almaz, imleç içine girmez (üzerinden atlar), seçilemez, kopyalanmaz, `documentText()`/`paragraphText()` çıktısında ve sayfanın yardımcı teknolojiler için tutulan metninde (`.bym-doc-text`) yoktur; düzen öğesinde `deleted: true` ile işaretlidir. Silinmiş görsel ve not işareti gösterilmez. Yalnız silinmiş metinden oluşan bir satırın kendi konumu yoktur: **Yukarı/Aşağı ok, PageUp/PageDown ve Shift'li seçim** böyle satırların üzerinden geçer (imleç hiçbir zaman yerinde saymaz); böyle bir satıra **tıklanınca** imleç o satırda, silinmiş metnin sağ ucunda çizilir (model konumu çevresindeki satırlarla aynıdır; yazılan metin silinmiş metinden sonra gelir) ve oklarla iki yöne de çıkılır. Satırı silinmiş metinle biten bir konumda imleç silinmiş metnin **sağında** durur (yazılan metnin belireceği yerde). Bir alanın yalnız sonucu ve bitiş işareti silinmişse alan orada kapanır, ardındaki metin gösterilir. **Yazdırmada** da eklenen metin altı çizili, silinen metin üstü çizili basılır (ekrandaki sayfanın aynısı; "yalnız son hali" çıktısı yoktur).
- **İzlenmeyen işlemler reddedilir** (yukarıdaki "Henüz olmayanlar" satırı): `docOps` işlevi `RangeError` (`docOps.TRACK_UNSUPPORTED`) fırlatır, görünüm bunu `onNotice` ile bildirir, belge değişmez. Araç çubuğunda paragraf hizalama, liste, girinti, stil, tablo, görsel ve bağlantı düğmeleri devre dışıdır; karakter biçimi düğmeleri imleçte "sonraki yazılacak metin" için çalışır (yeni metnin biçimi izlenen bir değişiklik değildir), seçimde reddedilir.
- **İzleme KAPALIYKEN yapılan düzenleme ve belgedeki izlenen kayıtlar:** izlenmeyen bir düzenleme revizyon kaydı **üretmez** ve dokunmadığı kaydı yok etmez; listeden yalnız içeriği tümüyle silinen kayıtlar çıkar.
  - **Enter, sayfa sonu, tablo ekleme, çok satırlı yapıştırma** (paragrafı bölen her şey): dosyadaki **özgün işaret** paragrafın sonudur ve ikinci yarıda kalır — işaretin revizyon bayrakları (`w:ins`/`w:del`/taşıma), işaretin `w:rPrChange`'i, paragrafın `w:pPrChange`'i ve özelliklerde XML olarak saklanan kayıtlar (`w:numberingChange`) onunla birlikte. Kullanıcının yaptığı **yeni işaret** (ilk yarının sonu) düz işarettir: aynı biçim, hiçbir kayıt yok. Böylece "işareti Ann'in eklediği" bir paragrafı bölmek listede iki kayıt üretmez, "tümünü reddet" kullanıcının Enter'ını geri almaz, "tümünü kabul et" kullanıcının böldüğünü birleştirmez. (Başlık sonunda Enter ile "sonraki stil"e geçen yeni paragraf da özgün işaretin kayıtlarını taşır.)
  - **Birleştirme** (paragraf başında Backspace, sonunda Delete, paragrafları aşan seçimi silme): birleşen paragrafın **biçimi** eskisi gibi düzenleyici kuralıyla belirlenir (ilk paragraf boş değilse ilkinin biçimi), ama işaretin **revizyon bayrakları ve `w:rPrChange` kaydı** her zaman **kalan (ikinci) işaretten** gelir; giden işaretin revizyonu onunla birlikte gider (o işareti kullanıcı sildi). Yani "işareti silinmiş" bir paragrafın sonunda Delete'e basınca silinmiş işaret kalıcı olarak gider (kabul edilecek bir şey kalmaz); işareti izlenen bir paragrafın **başında** Backspace'e basınca o paragrafın işaret kaydı birleşen paragrafta durur ve listede kalır.
  - **Tabloya satır/sütun ekleme:** yeni satırın satır, hücre ve paragraf özellikleri komşusundan **kayıtsız** kopyalanır (satırın `w:ins`/`w:del`'i, `w:trPrChange`, `w:tblPrExChange`, `w:tcPrChange`, `w:cellIns`/`w:cellDel`/`w:cellMerge`, paragraf işaretinin bayrakları ve `…Change` kayıtları alınmaz): yinelenen kimlik ve kullanıcının yapmadığı "değişiklik" oluşmaz. (İzleme açıkken bu işlemler eskisi gibi reddedilir.)
  - Listeden çıkan öğe (boş liste öğesinde Enter, paragraf başında Backspace) dosyada da listeden çıkar: `w:numPr` içinde saklanan bir numara değişikliği kaydı varsa o da numarayla birlikte gider (önceden öğe modelde listeden çıkıyor, dosyada listede kalıyordu).
- **İzleme açıkken farklı davrananlar:** paragraf başında Backspace listeden çıkarmaz/girintiyi geri almaz, önceki paragraf işaretini siler; Enter boş liste öğesini listeden çıkarmaz ve başlıktan sonra "sonraki stil"e geçmez (ikisi de izlenemeyen biçim değişikliği olurdu); kesme, panoya sığmayan içerik için reddedilmez (belgeden yazarın kendi yazdığı metinden başka hiçbir şey çıkmaz; o da panoda metin olarak durur).
- **Tarih biçimi:** `date` seçeneği geçerli bir `xsd:dateTime` olmalıdır: saniye zorunlu (`2026-10-05T10:00` reddedilir), gün aya ve artık yıla göre (31 Şubat reddedilir), yıl en az `0001`, saat dilimi yok, `Z` ya da en çok `±14:00`; `24:00:00` kabul edilmez. Aksi `RangeError`.
- **Revizyon kimlikleri:** yeni kimlik, belgedeki (tek ya da çift tırnakla, `=` çevresinde boşlukla yazılmış olanlar dahil) ve sonradan yapıştırmayla gelen her kimlikten büyüktür; yapıştırma geri alınıp yinelense de.
- **Geri al/yinele:** izlenen her düzenleme tek adımdır (art arda yazılan karakterler eskisi gibi birleşir); hepsi geri alınınca kaydedilen dosya açılanla bayt bayt aynıdır.
- **Programdan:** her metin işlevi son parametre olarak `{ trackChanges?, author?, date? }` alır (`docOps.insertText(doc, konum, 'x', undefined, { author: 'Ada', date: '2026-10-05T09:08:07Z' })`); belge izleme istiyorsa seçenek verilmese de izlenir. `docOps.setTracking(doc, { force?, author?, now? })` bir belge için varsayılanı kurar (görünüm bunu kendisi yapar), `docOps.isTracking(doc)` sorar.
- UNSURE (Word'le karşılaştırılmadı; Word yoktu): üretilen `w:ins`/`w:del` yapısı şema sırasına ve iç içe kurallara göre sınandı ve bağımsız bir okuyucuyla okundu, ama Word'de açılıp kabul/ret denenmedi; silinen bir görselin `w:del` içindeki çalıştırmada kalması; `mc:AlternateContent` içeren silinen çalıştırmanın iç metninin `w:t` kalması.

#### İzlemeyi açıp kapatma

Araç çubuğundaki **Değişiklikleri izle** düğmesi (`aria-pressed`) belgenin ayarını (`settings.trackRevisions`) değiştirir; programdan `view.setTrackChanges(true | false)` (değişiklik yapıldıysa `true`), `docOps.setTrackRevisions(doc, on)` (`DocEdit`). Geri alınabilen bir adımdır; `onDocumentChange` tetiklenir, `view.trackingChanges` güncel durumu verir.

- **Kaydetme:** `word/settings.xml` özgün metninden kesip ekleyerek yazılır: açılırken `<w:trackRevisions/>` şema sırasındaki yerine (`w:revisionView` ve öncesindeki öğelerin ardına) eklenir, kapatılırken dosyadaki her `w:trackRevisions` öğesi (nasıl yazılmış olursa olsun) çıkarılır; dosya ayarı `w:val="false"` ile yazmışsa öğe yerinde değiştirilir. Başka hiçbir bayt değişmez. Ayarlar parçası olmayan belgede parça, ilişkisi ve içerik türüyle birlikte eklenir. "Aç → kapat → kaydet" ve "aç → geri al → kaydet" özgün dosyayı **bayt bayt** verir (eklenen parça ve ilişki de geri çıkar).
- **Bayt düzeyinde:** ayarlar parçası bayt sırası işareti (BOM), `\r\n` satır sonları ya da UTF-16 ile yazılmışsa bunlar da korunur; öğenin dışındaki her bayt aynıdır.
- **Kapatılamayan durumlar:** uygulanan `trackedChanges` koruması; ve uygulamanın isteği (`mount({ trackChanges: true })`, köprüde `load.trackChanges`) — uygulamanın açtığını kullanıcı kapatamaz. Uygulamanın isteğiyle izlenirken **açma** isteği de (`setTrackChanges(true)`, köprüde `setTrackChanges { on: true }`) hiçbir şey yapmaz (`false`): izleme zaten açıktır ve dosyanın ayarına dokunulmaz. İkisinde de düğme basılı ve devre dışıdır (ipucunda neden yazar), `setTrackChanges(false)` `false` döner ve bildirir, `docOps.setTrackRevisions` `RangeError` (`docOps.TRACK_ENFORCED`) fırlatır; `view.trackingEnforced` bunu sorar. Uygulamanın isteğiyle izlenen belgede dosyanın kendi ayarı **değişmez** (ayarlar parçası yazılmaz).
- Salt okunur görünümde `setTrackChanges` hiçbir şey yapmaz (`false`).

#### İzlenen değişiklikleri listeleme, kabul etme ve reddetme

```js
const view = mounted.view;                 // DocumentView
const list = view.revisions();             // DocRevision[]
// { index, kind, author?, date?, story, start, end, text, marks, rows, count, move?, canAccept, canReject }
view.goToRevision('next');                 // seçer, görünüme getirir, ne olduğunu duyurur; sonda başa sarar
view.acceptRevision(0);                    // dizinle; ya da listedeki nesneyle; bağımsız değişkensiz: seçimin değdiği değişiklikler
view.rejectRevision(list[1]);
view.acceptAllRevisions();                 // view.rejectAllRevisions()

// Görünüm olmadan (Node'da da):
const { docOps } = BymOffice;
const history = new docOps.DocHistory(doc);
const [first] = docOps.revisionsOf(doc);
history.applyEdit(docOps.rejectRevision(doc, first));
const all = docOps.acceptAll(doc, (r) => r.author === 'Ayşe');   // { ops, selection, resolved, skipped }
history.applyEdit(all);
```

**Liste** (`docOps.revisionsOf(doc)`, `docOps.listRevisions(doc, { limit })` → `{ list, more }`): okuma sırasıyla gövde, sonra üstbilgi/altbilgiler, sonra dipnot ve sonnotlar. En çok `docOps.REVISION_LIST_LIMIT` (2.000) kayıt döner; fazlası varsa `more: true`. Listeleme belgenin üzerinden tek geçiştir. **Bir kayıt, okuyanın tek düzenleme saydığı birimdir:** araya başka metin girmeden art arda gelen, aynı türde, aynı yazarın, aynı tarihli (ya da aynı kimlikli) parçalar tek kayıttır; paragraf işaretini aşan ekleme tek kayıttır (işaret `text` içinde `¶`, `marks` sayısıyla), eklenen tablo satırları içerikleriyle tek kayıttır (`rows`). Bir silme ile yanındaki aynı yazarın eklemesi ("değiştirme") Word'deki gibi **iki** kayıttır. Bir taşımanın iki ucu iki kayıttır (`moveFrom`, `moveTo`; aynı `move` adını taşırlar). `kind`: `insert`, `delete` (metin, paragraf işareti, tablo satırı), `moveFrom`, `moveTo`, `format` (karakter biçimi: `w:rPrChange`, metinde ya da paragraf işaretinde), `paragraphFormat` (`w:pPrChange`), `tableFormat` (`w:tblPrChange`, `w:tblPrExChange`, `w:trPrChange`, `w:tcPrChange`), `sectionFormat` (`w:sectPrChange`), `other` (modelin yalnız XML olarak sakladığı değişiklikler; `count` kaç öğe olduğunu söyler). `start`/`end` konumdur (`{ story, path, offset }`); silme ve paragraf işareti yer kaplamadığı için ikisi aynıdır. `text` en çok 80 karakterlik önizlemedir. Yazar adı dosyadan geldiği gibi düz metindir; arayüz onu yalnız metin olarak basar.

**Kabul ve ret ne yapar** (Word'ün anlamı):

| Tür | Kabul | Ret |
|---|---|---|
| Ekleme (`w:ins`) | Sarmalayıcı kalkar, içerik düz içerik olur | İçerik belgeden çıkar |
| Silme (`w:del`) | İçerik belgeden çıkar | Sarmalayıcı kalkar; `w:delText` yeniden `w:t`, `w:delInstrText` yeniden `w:instrText` olur |
| Eklenmiş paragraf işareti | İşaret düz olur | Paragraf sonrakiyle birleşir |
| Silinmiş paragraf işareti | Paragraf sonrakiyle birleşir | İşaret düz olur |
| Taşıma (`w:moveFrom` / `w:moveTo`) | Kaynak içerik çıkar, hedef düz kalır | Hedef içerik çıkar, kaynak düz kalır |
| Biçim değişikliği (`w:rPrChange`, `w:pPrChange`, `w:sectPrChange`, `w:tblPrChange`, `w:tblPrExChange`, `w:trPrChange`, `w:tcPrChange`) | Kayıt kalkar, şimdiki biçim kalır | Kayıttaki eski özellikler docx okuyucusuyla okunur ve şimdikilerin yerine geçer; kayıt kalkar |
| Eklenmiş tablo satırı | Satır kalır | Satır çıkar |
| Silinmiş tablo satırı | Satır çıkar | Satır kalır |
| `other` | Yapılamaz | Yapılamaz |

- **Paragraf birleşmesi (Word'ün kuralı):** birleşen paragraf **ikinci (sonraki) paragraftır**; birincinin içeriği onun içeriğinin önüne gelir. Kalan işaret ikincinin işaretidir ve onunla birlikte stili, listesi, hizalaması, `w:pPrChange` kaydı ve bölüm sonu kalır; işareti kaldırılan paragrafın özellikleri işaretiyle birlikte gider. Kural birleşme anında paragrafın boş olup olmadığına bakmaz, bu yüzden **sonuç değişikliklerin çözülme sırasına bağlı değildir**. (Düzenleyicinin öbür birleştirmeleri — Backspace, seçim silme — eskisi gibi `joinedProps` kuralını kullanır: ilk paragraf boş değilse ilkinin biçimi.) **Bölüm sonu:** işareti kaldırılan paragraf bir bölümü bitiriyorsa bölüm sonu da kalkar (Word de siler) ve o metin sonraki bölümün sayfa düzenine geçer; işlem yapılır ve bu **bildirilir**: `acceptRevision`/`rejectRevision`/`acceptAll`/`rejectAll` dönüşündeki `sectionsRemoved` kaç bölüm sonunun kalktığını söyler, görünüm sonucu duyururken ekler ("1 bölüm sonu, paragraf işaretiyle birlikte kaldırıldı."). Çıkması gereken işaretin ardında birleşecek paragraf yoksa (hücrenin son paragrafı, tablodan hemen önceki paragraf): paragraf boşsa ve blok listesi onsuz geçerliyse paragrafla birlikte çıkar; değilse işaret **düz işaret** olarak kalır (Word de son paragraf işaretini çıkaramaz).
- **Taşıma:** aynı adı taşıyan iki uç (`w:moveFromRangeStart`/`w:moveToRangeStart` `w:name`) **birlikte** çözülür — birini kabul etmek ötekini de kabul eder ("tümü"ne verilen süzgeç yalnız bir ucu seçse de) — ve aralık işaretleri kalkar. Adı olmayan (aralık işaretsiz) uç tek başına, ekleme/silme gibi çözülür. Bir taşımanın aralık işaretlerinden biri modelin yalnız sakladığı XML'deyse (metin kutusu, `mc:AlternateContent`, yorum, tablo satırları arasındaki saklanan işaret) o taşıma **çözülemez**: kayıtları `canAccept: false, canReject: false` taşır, tek işlemde `REVISION_UNSUPPORTED`, "tümü"nde iki ucu da atlanır ve sayılır (bir ucu çözüp ötekini eşsiz bırakmamak için).
- **Biçim reddi:** kayıt eski özellikleri taşımıyorsa (boş `w:rPr`) eski biçim "biçimsiz"dir. Kayıttan yalnız **biçim** okunur: şemanın izin vermediği içerik (bir `w:pPrChange` içinde `w:sectPr` ya da işaret `w:rPr`'si; bir `w:rPrChange`/`w:trPrChange` içinde `w:ins`/`w:del`; iç içe `…Change`) uygulanmaz — bozuk ya da kötü niyetli bir dosyada ret, paragrafı bölüm sonu yapamaz ve yeni revizyon üretemez. `w:pPrChange` paragraf işaretinin biçimini ve bölüm sonunu, `w:sectPrChange` üstbilgi/altbilgi başvurularını, `w:trPrChange` satırın ekleme/silme işaretini taşımaz: bunlar şimdiki haliyle kalır. Hücresinde `w:cellIns`/`w:cellDel`/`w:cellMerge` da bulunan bir `w:tcPrChange` **reddedilemez** (`canReject: false`; eski özellikler yazılırken o işaretler kaybolurdu), kabul edilebilir.
- **Kimlikler ve dosyadaki baytlar:** bir düzenlemenin böldüğü her izlenen öğe kendi `w:id`'sini alır, ama yalnız **yeniden yazılacak** (değişen) bloklarda: komutun dokunmadan geri verdiği blok dosyadaki kimliğiyle kalır ve özgün baytlarıyla yazılır (aynı kimliği iki kez taşıyarak **gelen** bir dosyada bile model, yazılıp yeniden okunan dosyayla aynıdır).
- **Boş sarmalayıcının yer tutucusu:** içeriği kalmayan bir içerik denetimi (`w:sdt`) modelde `empty` yer tutucusuyla durur. İki paragrafa bölünmüş bir denetimin yarıları yeniden birleşince (kabul/ret, Backspace/Delete) ya da denetime yeniden içerik gelince yer tutucu kalkar: dolu denetimde bayat yer tutucu kalmaz (dosya zaten doğru yazılıyordu; model yeniden okunanla aynı olur).
- **Tablo:** çıkan satırda başlayan dikey birleşme sonraki satırda başlar; son satırı çıkan tablo çıkar ve gerekiyorsa (liste boş kalacak, tabloyla bitecek ya da iki tablo bitişecekse) yerine boş bir paragraf konur.
- **Çıkan içerikten kalanlar:** yer imi, yorum ve izin aralığı işaretleri ile yorum başvurusu yerinde kalır (her silmenin kuralı; yorumun kendisi silinmez). Son başvurusu çıkan **dipnot/sonnot da çıkar** (reddedilen eklenmiş başvuru, kabul edilen silinmiş başvuru). Tümü boşalan bağlantı çıkar (içindeki yer imleri kalır); içerik denetimi (`w:sdt`), basit alan (`w:fldSimple`), akıllı etiket ve `w:customXml` boş olarak kalır. Görsel, denklem ve çıkan içerikteki tanınmayan öğeler içerikle birlikte çıkar; kalan içerikteki tanınmayan öğe ve öznitelikler yerinde durur.
- **Alanlar:** bütünüyle eklenmiş ya da silinmiş alan bütünüyle çıkar ya da kalır. Bir alanın işaretlerinden (başlangıç, ayırıcı, bitiş) bazılarını götürüp bazılarını bırakacak kabul/ret yapılmaz (`REVISION_UNSUPPORTED`; "tümü"nde atlanır). Alanlar öykü boyunca izlenir: başlangıcı bir paragrafta, bitişi sonraki paragrafta ya da çıkacak bir tablo satırında olan alan da böyle korunur. Çözülen paragraftaki alanın kodu ve sonucu yeniden derlenir (silinmiş kodu reddedilen alan, düzen için de yeniden o alandır).
- **İç içe:** başkasının eklemesinin içindeki silme ayrı bir kayıttır. Dıştaki eklemeyi reddetmek içindeki her şeyi çıkarır; sıra sonucu değiştirmez (birim testi: farklı biçimli paragraflar üzerinde rastgele sırayla tek tek = aynı kararlarla "tümü", dosya ve model aynı).
- **Hiçbir şey yarım yapılmaz:** bir kaydın bir parçası çözülemiyorsa kaydın tamamı (ve taşımada eşi) bırakılır. Tek kayıtta `RangeError` (`docOps.REVISION_UNSUPPORTED`), "tümü"nde atlanır: dönen `skipped` atlanan, `resolved` çözülen kayıt sayısıdır. Liste eskimişse (araya düzenleme girmişse) `RangeError` (`docOps.REVISION_GONE`). Listeden gelen nesnenin **kopyası** da verilebilir (`JSON`, `postMessage`): kopya, tür + yazar + tarih + metin + öykü + dizin **ve başlangıç/bitiş konumu** şimdiki listede aynen bulunan kaydı adlar; tam eşleşme yoksa `REVISION_GONE` (liste kaydıysa başka bir değişiklik çözülmez).
- **Kabul/ret izlenmez** (izleme açıkken de; Word de izlemez) ve tek geri alma adımıdır; geri alınınca kaydedilen dosya açılanla bayt bayt aynıdır. Dokunulmayan paragraflar özgün XML'den kopyalanır.
- **Koruma:** uygulanan her korumada (`trackedChanges` dahil) kabul/ret reddedilir (`docOps.REVISIONS_PROTECTED`; Word'de de yapılamaz); kayıtların `canAccept`/`canReject` alanları `false` olur. Salt okunur görünümde görünümün yöntemleri `false` döner; liste ve gezinme çalışır.
- **doc ve RTF:** `convert: true` ile açılan belgedeki revizyonlar docx'e taşındığı için aynı biçimde listelenir ve çözülür.

**Arayüz.** Araç çubuğunun sonunda "değişiklikler" grubu vardır: **Değişiklikleri izle** (aç/kapat), **Önceki değişiklik**, **Sonraki değişiklik**, **Değişikliği kabul et**, **Değişikliği reddet**, **Tüm değişiklikler** (menü: Tüm değişiklikleri kabul et / reddet). Önceki/sonraki değişikliği **seçer**, görünüme getirir ve canlı bölgede (ve `onNotice` ile) duyurur ("Değişiklik 2 / 5. Silindi (Ayşe), 2026-10-01 09:00: …"); sondan sonra başa, baştan önce sona sarar ve bunu söyler. Kabul/ret, gezinmeyle seçilen kaydı; yoksa imlecin **içinde** durduğu kaydı; seçim birden çok kaydı kapsıyorsa **hepsini** (tek adımda) çözer. İmleç (seçim yokken) her zaman **en çok bir** kayda değer: iki kaydın ortak sınırındaysa (bir eklemenin bittiği, bir silmenin durduğu yer) okuma sırasında **önce gelen** seçilir; düğmelerin ipucu neyin çözüleceğini söyler. Düğmeler sonucu duyurur ("1 değişiklik kabul edildi. 2 değişiklik kaldı.") ve Word'deki gibi **sıradaki değişikliğe geçer** (programdan `acceptRevision(hedef, true)`; varsayılan geçmez). İmleç kabul edilen eklemenin sonunda, öbür durumlarda değişikliğin başladığı yerde kalır; "tümü" imleci yerinde bırakır. Belgede değişiklik yoksa gezinme ve menü, seçim bir kayda değmiyorsa kabul/ret devre dışıdır. İmleç tek bir kayıttayken kabul/ret düğmelerinin ipucu kaydı anlatır; eklenmiş ya da biçimi değişmiş metnin üzerine fareyle gelince sayfanın ipucu da (kim, ne zaman, ne). Düğmeler klavyeyle (araç çubuğu tek Tab durağı, oklarla gezilir) ve dokunmatikte kullanılır; dar ekranda araç çubuğu tek satır kalıp yana kayar.

**2.000'den çok değişiklik:** `revisions()` ilk 2.000 kaydı verir (`view.moreRevisions` fazlası olduğunu söyler); imlecin değdiği kaydı bulma, önceki/sonraki ve kabul/ret bu sınıra **bağlı değildir** (belgenin bütün kayıtlarında aranır; 2.001. değişiklik de seçilir, çözülür, "Sonraki" 2.000'den sonra devam eder). Sayı duyurularda "2000+" olarak söylenir ("Değişiklik 2001 / 2000+"). Programdan: `docOps.listRevisions(doc, { limit: Infinity })`, `docOps.revisionsTouching(liste, aralık)`, `docOps.revisionStep(liste, aralık, 'next' | 'previous', şimdiki)`.

`DocumentView` üyeleri: `moreRevisions`, `revisions()`, `selectedRevisions()`, `describeRevision(kayıt)`, `goToRevision(dizin | 'next' | 'previous')`, `acceptRevision(hedef?, sonrakineGeç?)`, `rejectRevision(…)`, `acceptAllRevisions()`, `rejectAllRevisions()`, `setTrackChanges(on)`, `trackingChanges`, `trackingEnforced`. Hepsi değişiklik yaptıysa `true` döner ve `onDocumentChange` tetiklenir; yapılamayanın nedeni `onNotice` ile bildirilir. `revisions()` görünümden yapılan düzenlemeleri izler; modeli görünümün dışından değiştirdiyseniz listeyi `docOps.revisionsOf(doc)` ile alın.

UNSURE (Word yoktu; karşılaştırılamadı): birleşecek paragrafı olmayan işaretin düz işaret olarak kalması; ortak sınırda imlecin "önce gelen" kayda sayılması (Word hangisini seçer bilinmiyor); kayıtların Word'ün gezinmesiyle bire bir aynı gruplanıp gruplanmadığı (özellikle biçim değişiklikleri ve tablo içi); reddedilen `w:sectPrChange`'te üstbilgi/altbilgi başvurularının şimdiki haliyle kalması; çıkan içerikteki yorum başvurusunun kalması (Word yorumu da siler).

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

### Fontlar

Dosyalar font adlarını taşır, fontların kendisini taşımaz; metin, cihazda bulunan fontla ölçülür ve çizilir. İstenen font cihazda yoksa satır ve sayfa sonları kayar. Motor bunu dört yolla azaltır:

**1. Kod sayfası sonekli adlar.** Eski Windows belgeleri (doc, RTF, xls) aynı fontu kod sayfasına göre adlandırır: `Bookman Old Style Tur`, `Arial CE`, `Times New Roman Cyr`. Böyle bir font hiçbir cihazda yoktur; kastedilen, sonekin önündeki fonttur. Gösterim ve ölçümde ad, taban adıyla birlikte denenir:

| Sonek | Örnek | Taban ad |
|---|---|---|
| ` Tur`, ` CE`, ` Cyr`, ` Greek`, ` Baltic` (büyük/küçük harf fark etmez) | `Arial Tur`, `Arial CYR` | `Arial` |
| ` (Turkish)`, ` (Hebrew)`, ` (Arabic)`, ` (Vietnamese)`, ` (Thai)`, ` (Greek)`, ` (Cyrillic)`, ` (Baltic)`, ` (Central European)`, ` (Western)` | `Tahoma (Thai)` | `Tahoma` |

- Yalnız bu liste, adın **sonunda** ve ayrı bir sözcük olarak sonektir. `Arial Hebrew`, `Noto Sans Thai`, `Arial Turkish` (ayraçsız) kendi adlarıyla kalır; ad yalnız sonekten oluşuyorsa (`Tur`) dokunulmaz; yalnız bir sonek soyulur.
- Font dizgisinde dosyadaki ad **önce** gelir, taban ad hemen ardından: `"SBL Greek"` gibi gerçekten o adı taşıyan bir font cihazda varsa o kullanılır.
- Yerine koyma zinciri, satır yüksekliği tablosu, serif/eş aralıklı sınıflaması ve "cihazda var mı" sınaması (`hasFont`) taban adla çalışır.
- **Modeldeki ve dosyaya yazılan ad değişmez**; soyma yalnız gösterim katmanındadır (belgede `CanvasTextMeasurer`, tabloda `cssFont`). `BymOffice.baseFontFamily('Arial Tur')` → `'Arial'`.

**2. Ölçü uyumlu yedekler.** Belgede bir font adı şu sırayla denenir: `prefer` ile kaydedilmiş fontlar → dosyadaki ad → taban ad → `as` ile kaydedilmiş fontlar → dosyanın font tablosundaki alternatif ad (`w:altName`, RTF `\falt`) → aşağıdaki ölçü uyumlu açık fontlar (paketle gelen font bunların içindedir) → genel aile (`serif` / `sans-serif` / `monospace`).

| İstenen | Ardından denenenler |
|---|---|
| Calibri, Calibri Light | Carlito |
| Cambria, Cambria Math | Caladea |
| Arial (Helvetica) | Liberation Sans, Arimo, Helvetica |
| Arial Narrow | Liberation Sans Narrow |
| Times New Roman (Times) | Liberation Serif, Tinos, Times |
| Courier New (Courier) | Liberation Mono, Cousine, Courier |
| Georgia | Gelasio |
| Bookman Old Style | TeX Gyre Bonum, URW Bookman, URW Bookman L, Bookman |
| Book Antiqua | TeX Gyre Pagella, Palatino Linotype, Palatino, URW Palladio L |
| Century Schoolbook | TeX Gyre Schola, Century Schoolbook L |
| Century Gothic | TeX Gyre Adventor, URW Gothic |

**TeX Gyre Bonum dışında** bu fontlar pakette yoktur; cihazda ya da sayfanızda (`@font-face` ya da `registerFont`) varsa kullanılırlar. TeX Gyre Bonum paketle gelir (4. madde). Son dört satır UNSURE: yedekler, Monotype ailelerinin uyumlu çizildiği PostScript fontlarının (ITC Bookman, Palatino, New Century Schoolbook, ITC Avant Garde Gothic) eşleridir; Monotype dosyalarıyla genişlikleri karşılaştırılamadı. Tahoma, Verdana, Garamond ve Palatino Linotype için ölçü uyumlu açık font bilinmiyor: zincirleri yalnız genel aileyle biter. Tablo (xlsx/xls) görünümünde zincir daha kısadır: ad, taban ad, kaydedilmiş fontlar, (Bookman Old Style için, yüklenmeye başladıysa) paketle gelen font, sonra Carlito/Calibri/Arial.

Satır yüksekliği, yaygın fontlar için font dosyalarından okunmuş sayılardan gelir (Word'ün kullandığı `usWinAscent`/`usWinDescent` ve dış satır aralığı) ve her tarayıcıda aynıdır. Belirleyici olan belgenin **istediği** ailedir (Word belgeyi onunla sayfalamıştır), cihazın çizdiği font değil. Tek istisna paketle gelen fonttur: istenen aile tabloda yoksa ve zincirde cihazın gösterebildiği **ilk** font TeX Gyre Bonum ise onun satırı kullanılır. Bookman Old Style isteyen ve TeX Gyre Bonum ile çizilen metnin satırı, Bonum dosyalarından okunmuş değerlerle (1000 birimlik em'de düz 1069/268, kalın 1153/306, italik 1073/279, kalın italik 1141/328; satır arası boşluk 0) hesaplanır ve **bütün tarayıcılarda aynıdır**. UNSURE: gerçek Bookman Old Style dosyası elde olmadığı için onun değerleri bilinmiyor; bu satır yüksekliği Word'dekinden farklı olabilir ve sayfa sonları kayabilir. Tablodaki öbür satırlar başka aileye **ödünç verilmez**: tabloda olmayan bir aile başka bir fontla çizildiğinde (ör. Book Antiqua → Palatino Linotype, Cambria Math → Caladea, `w:altName`'i Arial olan bilinmeyen aile) satır yüksekliği tarayıcının o an çizdiği fontun kutusudur (tarayıcıdan tarayıcıya bir pikselden az oynar).

**3. Kendi fontunuzu kaydetme: `registerFont`.** Uygulama, cihazda olmayan bir fontu sayfaya ekleyebilir ve belge font adlarının yerine geçirebilir:

```js
await BymOffice.registerFont({
  family: 'TeX Gyre Bonum',                 // fontun kendi aile adı
  source: '/fonts/texgyrebonum-regular.woff2', // adres ya da baytlar (ArrayBuffer, Uint8Array, Blob/File)
  as: ['Bookman Old Style'],                // yerine geçeceği belge font adları (isteğe bağlı)
});
await BymOffice.registerFont({ family: 'TeX Gyre Bonum', source: boldBytes, weight: 'bold', as: ['Bookman Old Style'] });
```

| Alan | İş |
|---|---|
| `family` | Aile adı. Dosyadan gelen adlarla aynı kuraldan geçer: yalnız harf, rakam, boşluk ve tire kalır; geriye bir şey kalmıyorsa reddedilir |
| `source` | Font dosyasının adresi (dizge; olduğu gibi tarayıcıya verilir, CORS ve içerik güvenliği ilkesi sayfanızınkidir) ya da baytları. TrueType, OpenType, WOFF, WOFF2 (tarayıcının desteklediği) |
| `weight` | `'normal'` (varsayılan), `'bold'` ya da 1–1000 arası sayı |
| `style` | `'normal'` (varsayılan) ya da `'italic'` |
| `as` | Bir ad ya da ad dizisi (en çok 32): belgede/tabloda bu adlar (ve sonekli halleri: `Bookman Old Style Tur`) istenince, **cihazda o font yoksa** bu font kullanılır. Cihazda varsa cihazdaki kullanılır |
| `prefer` | `true` ise `as` adları için cihazdaki font yerine **her zaman** bu font kullanılır (her cihazda aynı satır sonları için) |

- Dönüş: `Promise<RegisteredFont>` (`{ family, weight, style, as, prefer, status: 'loaded' }`; `weight` sayıdır: `'normal'` 400, `'bold'` 700); font dosyası okunup kullanılabilir olunca çözülür. Dosya yüklenemez ya da font değilse, kayıt geçersizse, 64 yüz sınırı aşılırsa reddeder. Kalın ve italik için ayrı dosyalar ayrı çağrılarla kaydedilir; kaydedilmeyen yüzü tarayıcı yapay olarak üretir. Aynı yüz (aile adı büyük/küçük harf ayırmadan + ağırlık + stil; `'bold'` ile `700` aynı yüzdür) yeniden kaydedilirse yenisi **yüklendiği anda** eskisinin yerini alır ve görünümler bir kez yeniden dizilir; o ana dek eski yüz kullanılmaya devam eder (`registeredFonts()` ikisini de gösterir: biri `loaded`, biri `loading`). Yeni dosya yüklenemezse çağrı reddeder ve eski yüz kalır. Aynı yüz için yüklenmesi süren bir kayıt varken yenisi gelirse öncekinden vazgeçilir (onun sözü reddeder).
- Yükleme bitince **açık bütün görünümler kendiliğinden yeniden ölçer**: belge baştan dizilir (düzenleyicide imleç ve seçim korunur, görünen sayfa aynı kalır; salt okunur görünümde seçim kalkar), tablo yeniden çizilir (hücreler ve şekil/metin kutusu metinleri). Yeniden dizmeden sonra görünüm **okunan yerde kalır**: en çok görünen sayfa ve o sayfanın içindeki kaydırma konumu korunur (font geldi diye sayfanın başına dönülmez; yeni düzende sayfa sınırları kayarsa aynı sayfanın aynı yüksekliği gösterilir); `view.relayout()` için de böyledir. Belge düzenleyicide bir giriş yöntemi (IME) bileşimi ya da fareyle sürükleyerek seçim sürüyorsa yeniden dizme **o bitene dek ertelenir** (sayfaya `@font-face` ile gelen fontlar için de). Baskı hazırlığı sürerken gelen font, sayfaların baştan hazırlanmasına yol açar. Yazdırma da kayıtlı fontu kullanır. Belge araç çubuğundaki font listesine kayıtlı aileler eklenir.
- `BymOffice.registeredFonts()` kayıtları listeler (yüklenmekte olanlar `status: 'loading'`); `BymOffice.unregisterFont(family)` ailenin bütün yüzlerini kaldırır (görünümler yine yeniden dizilir) ve bir şey kaldırdıysa `true` döner. Üçü `BymOffice.fonts` altında da durur (`BymOffice.fonts.registerFont` …).
- Kayıt sayfa genelindedir (tarayıcının `FontFace` arayüzü ve `document.fonts`); görünüm kaldırılınca silinmez. `FontFace` olmayan ortamda (Node, çok eski WebView) `registerFont` açıklayıcı bir hatayla reddeder; `@font-face` kuralıyla yedek yol **yoktur**. Sayfanıza kendi `@font-face` kuralınızla eklediğiniz fontlar da çalışır: belge görünümü tarayıcının "fontlar yüklendi" olayında yeniden dizilir.
- `mount({ fonts: [...] })` aynı nesneleri alır (yukarıdaki seçenek tablosu); WebView köprüsünde `load.fonts` (bkz. `docs/webview-koprusu.md`).
- Dosyanın içindeki gömülü fontlar (docx `word/fonts/*.odttf`) **yüklenmez**; yalnız dosyada korunur.

**4. Paketle gelen font: TeX Gyre Bonum (Bookman Old Style için).** Bookman Old Style macOS, iOS ve Android'de yoktur. Paket, onunla aynı karakter genişliklerine göre çizilmiş açık fontu, TeX Gyre Bonum'u dört yüzüyle (düz, kalın, italik, kalın italik) getirir: `dist/fonts/tex-gyre-bonum/texgyrebonum-{regular,bold,italic,bolditalic}.otf` (toplam 576 KB) ve lisansı. Dosyalar GUST e-foundry'nin dağıttığı haliyle, **değiştirilmeden** gelir (alt küme çıkarılmaz, biçim çevrilmez, ad değişmez; derleme sağlama toplamlarını denetler). Betiklerin (`bymoffice.min.js`, `bymoffice.esm.js`) içinde font yoktur.

- **Ne zaman yüklenir.** Bir belge ya da tablo görünümü `Bookman Old Style` (ya da sonekli hali: `Bookman Old Style Tur`) ile yazılmış metni ilk kez ölçtüğünde, şu üçü birden doğruysa dört dosya istenir: cihaz o aileyi gösteremiyor; font zincirinde ondan önce gelen hiçbir şey (dosyanın `w:altName` alternatifi, `as` ile kaydettiğiniz font) gösterilemiyor; uygulama o ad için kendi fontunu kaydetmemiş (`registerFont` ile `as: ['Bookman Old Style']`, ya da aile adı `Bookman Old Style` / `TeX Gyre Bonum` olan bir kayıt; yüklenmesi süren kayıt da sayılır). Bu aileyi kullanmayan dosyada **hiçbir istek yapılmaz**. Aynı sayfadaki bütün görünümler için tek yükleme yapılır (dört istek, yüz başına bir); gelince açık görünümler kendiliğinden yeniden dizilir. Cihazda zincirin başka bir yedeği (ör. Linux'ta URW Bookman) bulunsa da TeX Gyre Bonum yüklenir ve zincirde ondan önce gelir: bütün cihazlar aynı fontu gösterir. Font **dört yüzüyle birlikte** kullanılır: biri gelmezse hiçbiri kullanılmaz. Font, yalnız sayfaya eklendikten sonra font dizgilerine girer (belgede ve tabloda aynı kural).
- **Kendi fontunuz önce gelir.** `registerFont` ile `as: ['Bookman Old Style']` verdiğiniz font zincirde paketle gelenden öndedir; paketle gelen font `registeredFonts()` listesinde görünmez ve `unregisterFont` ile kaldırılmaz.
- **Dosyaya yazılan ad değişmez.** Belge araç çubuğunun font listesinde `Bookman Old Style` vardır (font yüklenebildiği sürece); seçilince belgeye `Bookman Old Style` yazılır, ekranda TeX Gyre Bonum ile gösterilir, Word'de gerçek fontla açılır.
- **Dosyalar nereden istenir.** Varsayılan: betiğin kendi adresinin yanındaki `fonts/` klasörü.

  | Kullanım | Çözülen adres |
  |---|---|
  | `<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js">` (unpkg de aynı) | `…/npm/bymoffice/dist/fonts/` — ek ayar gerekmez. Çapraz köken font isteği CORS ister; UNSURE: jsDelivr ve unpkg `Access-Control-Allow-Origin: *` gönderir, paket yayımlanmadığı için denenemedi |
  | `<script src="/vendor/bymoffice.min.js">` (kendi sunucunuz; `async`, `defer`, sonradan eklenen betik ve `?v=…` dahil) | `/vendor/fonts/` — `dist/fonts/` klasörünü betikle **birlikte kopyalayın**. Başka kökenden sunuyorsanız font dosyalarına `Access-Control-Allow-Origin` başlığı verin |
  | ES modülü (`import … from '…/bymoffice.esm.js'`, `<script type="module">`) | **Kendiliğinden çözülmez** (modül kendi adresini bilmez; motor `import.meta` kullanmaz, ES2019 olarak ayrıştırılır). Adresi siz verirsiniz, kendi modülünüzde bir satırla: `fonts.configure({ baseUrl: new URL('./fonts/', import.meta.url).href })` ya da `mount(el, { fontsUrl: '/dist/fonts/' })`. Vermezseniz istek yapılmaz ve bir kez bildirilir |
  | Paketleyici (Vite, webpack, esbuild…) | Çözülemez ya da yanlış çözülür (betiğin adresi paketleyicinin ürettiği dosyadır). `node_modules/bymoffice/dist/fonts` klasörünü sitenizin statik dosyalarına kopyalayın ve adresini verin: `mount(el, { fontsUrl: '/static/bymoffice-fonts/' })`. Dosyalara paket içinden `bymoffice/fonts/tex-gyre-bonum/…` adıyla da ulaşılır. (Paketleyicilerle denenmedi) |
  | Satır içi betik, `blob:`/`data:` modülü, `file:` olmayan bir sayfada `file:` adresi | Çözülemez; adresi verin |
  | WebView köprü sayfası (`bymoffice.webview.html`) | Ağ isteği yok: font sayfanın içindedir (bkz. `docs/webview-koprusu.md`) |

- **Ayar: `BymOffice.fonts.configure(options)`** (sayfa genelinde; yalnız verilen alanlar değişir; bir şey değiştiyse açık görünümler yeniden ölçer ve daha önce yüklenememiş font yeniden denenir):

  | Alan | Anlamı |
  |---|---|
  | `baseUrl` | Font klasörünün (`dist/fonts/`, içinde `tex-gyre-bonum/` olan) adresi; sayfanın adresine göre okunur (`'/static/bym-fonts'`, `'../fonts/'`, tam adres). `null`: paketle gelen font **kapalı**: hiçbir istek yapılmaz; font yüklenmişse sayfadan çıkarılır ve görünümler onsuz yeniden dizilir (okunan dosyalar bellekte kalır, yeniden açınca istek yapılmadan geri gelir). `undefined`: varsayılana (`<script>`'in yanı) döner |
  | `as` | Fontun yerine geçtiği belge font adları; varsayılan `['Bookman Old Style']`. Sonekli adlar taban adla kapsanır. Örn. `['Bookman Old Style', 'Bookman']` |
  | `source` | `(dosyaAdı) => baytlar \| adres \| Promise` — dosyayı kendiniz verirsiniz (köprü sayfası böyle yapar). Bir şey döndürmediği dosya `baseUrl`'den istenir |
  | `hasFamily` | `(aile) => boolean` — "cihazda bu aile var mı" sınamasının yerine geçer (testler için; yalnız tablo görünümünde ve ölçücüsü olmayan yerde sorulur, belgede `createCanvasMeasurer({ hasFamily })` geçerlidir) |

  `mount({ fontsUrl })` `baseUrl` ile, `mount({ fonts: false })` ve `fontsUrl: null` kapatma ile aynıdır. `BymOffice.fonts.embedded()` durumu verir: `{ family: 'TeX Gyre Bonum', as, baseUrl, status }`; `status`: `'off'`, `'idle'` (henüz gerekmedi), `'loading'`, `'loaded'` (dört yüz sayfada), `'failed'`.
- **Yüklenemezse.** Adres çözülemiyorsa ya da istekler başarısızsa dosya yine açılır ve metin genel yedek fontla gösterilir; sayfadaki görünümlerden **biri**, **bir kez** `onNotice` ile bildirir (`kind: 'info'`; "Bu cihazda "Bookman Old Style" fontu yok, yerine geçen font da yüklenemedi… Font klasörünün adresini "fontsUrl" ile verin ya da fontu "fonts: false" ile kapatın"). Yeniden deneme yapılmaz; `configure()` ile ayar değişirse yeniden denenir (yükleme sürerken değişen ayar, eski yükleme başarısız bittiğinde denenir) ve yeni bir başarısızlık yeniden bildirilir. Dört yüzden biri bile gelmezse font kullanılmaz (`status: 'failed'`).
- **Sınırlar.** Gerçek Bookman Old Style ile genişlik karşılaştırması yapılamadı (dosya yok; UNSURE: TeX Gyre Bonum, Bookman Old Style'ın uyumlu çizildiği ITC Bookman'ın eşidir). Satır yüksekliği Bonum'un kendi ölçüsüdür (2. maddenin sonu). `FontFace` olmayan ortamda (Node, çok eski WebView) font yüklenmez ve bildirim de yapılmaz. Uygulamanın kendi kaydı yüklenirken başarısız olursa paketle gelen font o görünüm yeniden ölçene dek istenmez. Paketleyicilerle deneme yapılmadı.
- **Lisans.** TeX Gyre Bonum, GUST Font License (LPPL 1.3c) ile dağıtılır; metni fontların yanındadır (`dist/fonts/tex-gyre-bonum/GUST-FONT-LICENSE.txt`), ayrıntı depo kökündeki `THIRD-PARTY-NOTICES.md` dosyasındadır. Font klasörünü kopyalarken lisans dosyasını da kopyalayın. BYM Office'in MIT lisansı font dosyalarını kapsamaz.

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
- **WebView köprüsü:** köprüde sekme şeridi aynı biçimde çalışır; sayfa işlemlerinin köprü komutları (`insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`) ve liste her değiştiğinde giden `sheets` iletisi `docs/webview-koprusu.md`'de anlatılır.

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
| `Op` türleri `insertSheet`, `deleteSheet`, `renameSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `restoreSheets`; `SheetsSnapshot` | `setTabColor` kendi tersidir; ötekilerin tersi `restoreSheets`'tir (kendi tersidir) ve sayfa listesinin önceki halini `SheetsSnapshot` olarak taşır. Hepsinde `sheet` etkilenen dizindir |
| `moveSheet(workbook, from, to)`, `setSheetState(workbook, index, state)`, `copySheet(workbook, index, { at?, name? })`, `setTabColor(workbook, index, color?)` | İşlemi (`Op[]`) üretir; değişecek bir şey yoksa boş liste; yapılamayacaksa `RangeError` |
| `canMoveSheet`, `canHideSheet`, `canCopySheet`, `copySheetName(workbook, index)`, `copySheetLosses(workbook, index)` | Neden (ya da `undefined`); kopyanın varsayılan adı; kopyaya alınmayacak içerik (`{ what, count }[]`) |

### Sayfa taşıma, gizleme, kopyalama, sekme rengi

```js
view.moveSheet(0, 2);                       // 1. sayfa 3. sıraya; true
view.setSheetState(1, 'hidden');            // gizle; 'visible' gösterir
const copy = view.copySheet(0);             // "Ad (2)", hemen arkasına; kopyanın dizini
view.copySheet(0, { at: 0, name: 'Yedek' });
view.setTabColor(0, '#FF8800');             // undefined rengi kaldırır
```

Her biri **tek geri alma adımıdır**, `workbook.sheetsModified`'ı kurar, `onChange({ sheets })`'i bütün dizinlerle çağırır; salt okunur görünümde hiçbir şey yapmaz (`false` / `-1`). Yazılmakta olan hücre, not ya da sekme adı işlemden önce uygulanır (`commitPending()`); uygulanamıyorsa (çok uzun not, geçersiz sayfa adı) işlem **yapılmaz**, alan açık kalır, neden `onNotice` ile bildirilir ve yöntem `false` / `-1` döner. Aynı kural `insertSheet`, `deleteSheet` ve `renameSheet` için de geçerlidir. İçinde değişiklik yapılmış bir **kural kutusu** açıksa da öyle: işlem yapılmaz, kutu açık kalır (değişiklik yapılmamış kutu kapatılır ve işlem yapılır). Formül yazılırken sekme menüsü açılmaz ve sekme sürüklenmez; sekme sürüklenirken açılan kural kutusu sürüklemeyi bırakır. **Kopyalanmış hücreler** (iç kopya; formülleriyle yapıştırılır) sekme rengi, gizleme/gösterme, taşıma, sayfa ekleme ve sayfa kopyalamadan sonra da geçerlidir; kopya alındığı andaki sayfalardan biri **silinir ya da yeniden adlandırılırsa** unutulur (formülleri artık olmayan adları anardı) ve yapıştırma düz metin olur.

- **Taşıma:** `moveSheet(from, to)`: `to`, sayfanın taşındıktan sonraki dizinidir. Sayfaya yerel tanımlı adlar (`localSheet`) ve etkin sayfa sayfalarını izler; gösterilen sayfa gösterilmeye devam eder. **Hiçbir formül değişmez**: 3B başvuru (`Sayfa1:Sayfa3!A1`) metin olarak aynı kalır ve artık uçlarının arasında duran sayfaları kapsar (Excel de böyle yapar; bu motor 3B başvuru hesaplamaz). **Arayüz:** sekmeyi fareyle sürükleyip bırakın: 6 pikselden kısa hareket tıklamadır; sürüklerken mavi çizgi (`.bym-tab-drop`) bırakma yerini gösterir, şeridin iki ucunda sekmeler kendiliğinden kayar. **Esc** sürüklemeyi bırakır (hiçbir şey taşınmaz); düğme şeridin dışında bırakılırsa, pencere odağını yitirirse ya da sürükleme sırasında geri al/yinele, bir sayfa işlemi, `setReadOnly(true)` ya da `destroy()` gelirse sürükleme taşıma yapmadan biter. **Dokunmatikte** sekmeye uzun basın (menü açılır); parmağı kaldırmadan kaydırırsanız menü kapanır ve sekme sürüklenir, parmağı yerinde kaldırırsanız menü kalır. Uzun basmadan kaydırmak şeridi kaydırır. **Klavye:** sekme menüsünde (Shift+F10) "Sola taşı" / "Sağa taşı" (komşu **görünür** sekmeyle yer değiştirir; odak sekmede kalır). Taşıma canlı bölgede ve `onNotice` ile söylenir.
- **Gizle / göster:** `setSheetState(index, 'hidden' | 'visible' | 'veryHidden')`. Görünür son sayfa gizlenemez (`RangeError`; menüde öğe nedeniyle devre dışı). Gösterilen sayfa gizlenirse sağındaki ilk görünür sayfa (yoksa solundaki) gösterilir ve etkin sayfa olur; gösterilen bir sayfa gösterilir. **Arayüz:** sekme menüsünde "Gizle"; gizli sayfa varsa "Göster…" gizli sayfaların listesini açar (ayrı bir menü; ok tuşlarıyla gezilir). **Çok gizli** (`veryHidden`) sayfalar listede yoktur; yalnız programla (`setSheetState(i, 'visible')`) gösterilir ve dokunulmadıkça dosyada öyle kalır. `view.hiddenSheets` listelenen dizinleri verir.
- **Kopyalama:** `copySheet(index, { at?, name? })` sayfanın modeldeki içeriğini yeni bir sayfaya kopyalar ve kopyayı gösterir. Varsayılan yer hemen arkası, varsayılan ad Excel gibi `Ad (2)` (boş olan en küçük sayı; 31 karakteri aşarsa ad kısaltılır; `Ad (2)`'nin kopyası `Ad (3)`). **Kopyalananlar:** hücreler (değer, biçim, formül ve dosyadaki sonucu), satır/sütun ölçü ve biçimleri, birleşik hücreler, kitap içi bağlantılar, süzgeç, görünüm (dondurma, yakınlaştırma, kılavuz çizgileri), sayfa yapısı (kenar boşlukları, kâğıt, üstbilgi/altbilgi, sayfa sonları), sekme rengi, notlar (metin ve yazar; kutuları yeni notun varsayılan yerinde) ve sayfaya yerel tanımlı adlar (yazdırma alanı, yazdırma başlıkları, süzgeç aralığı). **Formüller:** sayfa adı taşımayan başvurular aynen kalır ve kopyayı okur; kopyalanan sayfayı **adıyla** anan başvurular (ve yerel adlar, bağlantılar) kopyanın adına çevrilir (`Veri!A1` → `'Veri (2)'!A1`); başka sayfalara gidenler değişmez. Hiçbir şey hesaplanmaz. **Kurallar da kopyalanır:** veri doğrulamalar ve koşullu biçimler, o anda geçerli aralık ve formülleriyle (kopyada her biri düzenlenmiş kural sayılır; uzantı listesindeki doğrulama kuralı kopyada ana bölüme yazılır). **Kopyalanmayanlar** (dosyadaki haliyle saklandıkları için dosyada olmayan sayfaya yazılamıyor): koşullu biçimlerin yalnız 2010 uzantı listesinde duran bölümü (`conditionalFormatExtensions`: uzantıya özgü kurallar; veri çubuğu ana bölümdeki haliyle kopyalanır), görsel/şekil/grafik, tablo, pivot tablo, form denetimi, yorum zinciri, dış (kitap dışı) bağlantılar, satırların başka ad alanından gelen öznitelikleri (`rowAttributes`; Excel'in her satıra yazdığı `x14ac:dyDescent` ipucu kayıp **sayılmaz**, ad alansız öznitelikler — `thickBot`, `thickTop`, `ph`, `customFormat`'sız `s` — kopyaya **yazılır**); satır/sütun eklenip silinmiş sayfada sayfa sonları; sayfa parçasında modelin alan tutmadığı her öğe, okuyucunun not ettiği adıyla (`sheet.unmodelled`: `sparklineGroups`, `customSheetViews`, `ignoredErrors`, `cellWatches`, `sortState`, `phoneticPr`, `outlinePr`, `outlineLevels`, `codeName`, üstbilgi/altbilgi görselleri `legacyDrawingHF`, arka plan görseli `picture`, …). **Sayfa koruması** (`<sheetProtection>`, parola karmasıyla) ve korumalı aralıklar (`<protectedRanges>`) kopyaya **yazılır** (`sheet.protection`; korumalı sayfanın kopyası korumalıdır). `sheet.protection` dosyadaki haliyle durur; kopyada açık aralıkların `sqref`'i, kaynak sayfada o ana dek eklenen/silinen satır ve sütunlara göre **kaydırılır** (yazıcının kaynak sayfada yaptığıyla aynı kural; hücreleri tümüyle silinen aralık kopyaya girmez, adı/parola karması/`securityDescriptor`'ı aynen kalır); yorum zinciri kopyaya not olarak dönüşmez, atlanır ve sayılır. Bunlar varsa kopyalama **yine yapılır** ve alınmayanlar sayılarıyla `onNotice` (`kind: 'info'`) ile bildirilir; önceden `copySheetLosses(workbook, index)` ile sorulabilir. Tablo ya da pivot içeren sayfa reddedilmez: hücreleri düz hücre olarak kopyalanır. Grafik sayfası kopyalanamaz.
- **Sekme rengi:** `setTabColor(index, '#RRGGBB')`; `undefined` rengi kaldırır; başka yazım `RangeError`. **Arayüz:** sekme menüsünde "Sekme rengi…" araç çubuğundaki renk paletini (`.bym-popover`, 60 renk + "Renk yok") sekmenin altında açar. Renk sekmenin alt kenarında çizgi olarak görünür.
- **Kaydetme:** taşıma `workbook.xml`'de `<sheets>` sırasını, `activeTab`/`firstSheet`'i, `localSheetId`'leri ve `docProps/app.xml` başlık sırasını günceller; **parça adları ve `sheetId`'ler değişmez**, sayfa parçaları (seçili sekme işareti değişmedikçe) bayt bayt kalır. Gizleme/gösterme `<sheet state>`'i yazar; etkin sayfa gizlendiyse `activeTab` ve iki sayfanın `tabSelected` işareti de yazılır. **Birden çok penceresi olan dosyada** (`<workbookView>` > 1) gizlenen sayfayı gösteren her pencere etkin sayfaya (o da gizliyse ilk görünür sayfaya) çekilir, gizli sayfanın `tabSelected` işaretleri kalkar ve yalnız o pencerelerin görünümünde etkin sayfa seçili işaretlenir; öbür pencereler dokunulmadan kalır. `firstSheet` gizlenen bir sayfayı gösteriyorsa ilk görünür sayfaya çekilir. Rengi dosyadakiyle aynı olan (geri alınmış) sekme için sayfa parçası okunmaz bile. Sekme rengi yalnız o sayfa parçasındaki `<sheetPr><tabColor>`'ı değiştirir (yoksa `<sheetPr>` eklenir, kaldırılınca öğe çıkar; parçanın başka hiçbir baytı değişmez; `fullCalcOnLoad` yazılmaz); dosyadaki tema/indeks/ton renkli `tabColor` değiştirilmedikçe olduğu gibi kalır. Kopya, yeni sayfa gibi kendi parçasına (`sheetN.xml`), notları kendi yorum parçalarına, yerel adları `definedNames`'e yazılır; kaynak sayfanın parçaları bayt bayt kalır. Sıra değişince (taşıma, kopya) ekleme/silmede olduğu gibi `fullCalcOnLoad` yazılır ve `calcChain` çıkarılır. **Her biri için "kaydet → geri al → kaydet" açılan dosyayı bayt bayt geri verir.**
- **Retler:** adları yalnız harf büyüklüğüyle ayrılan sayfaları olan kitapta bunlar da reddedilir. Grafik, pivot tablo ya da form denetimi içeren kitapta taşıma, gizleme ve renk **serbesttir** (sayfa adına dokunmazlar); kopyalama da serbesttir (yukarıdaki kayıplarla).

### Hücre notları

Bir hücrenin notu (Excel 365'te "not", eski sürümlerde "açıklama"; dosyada `xl/commentsN.xml` + `xl/drawings/vmlDrawingN.vml`) görünümden ve programla eklenir, düzenlenir, silinir. Her değişiklik **tek geri alma adımıdır** ve `save()` ile dosyaya yazılır.

```js
const view = mounted.view;                    // WorkbookView
view.setComment(1, 1, 'Kontrol edildi');      // B2'ye not yazar (yoksa ekler, varsa metnini değiştirir)
view.setComment(1, 1, undefined);             // notu siler (boş ya da yalnız boşluk metin de siler; aşağıda "Metin kuralı")
view.editComment();                           // etkin hücrenin not alanını açar; editComment(4, 0) → A5
view.isEditingComment;                        // not alanı açık mı
```

- **Arayüz:** hücrenin sağ tık menüsünde (dokunmatikte uzun basma) notu olmayan hücrede **Not ekle**, olan hücrede **Notu düzenle** ve **Notu sil**; klavyede **Shift+F2** etkin hücrenin notunu ekler ya da düzenler. Not, balonun yerinde açılan çok satırlı alana (`.bym-note-input`) yazılır: **Ctrl+Enter** (macOS'ta Cmd+Enter de) ya da alanın dışına tıklamak/dokunmak **uygular**, **Esc vazgeçer**, Enter yeni satırdır. Alandayken tuşlar ızgaraya gitmez; giriş yöntemi (IME) bileşimi sırasındaki Enter/Esc notu etkilemez. Alan açıkken başka bir sayfa gösterilirse, görünüm salt okunura alınırsa ya da programdan başka bir değişiklik yapılırsa yazılan önce uygulanır; `undo()`/`redo()` çağrısı ise yazılanı bırakır (hücre editöründe olduğu gibi). Metni boşaltıp uygulamak notu siler. **Alan açılıp metin değiştirilmeden kapatılırsa hiçbir şey değişmez** (geri alma adımı oluşmaz; biçimli, satır sonu farklı ya da metni boş bir not olduğu gibi kalır). Köşedeki kırmızı üçgen ve balon hemen güncellenir; silme canlı bölgeden "Not silindi" diye duyurulur.
- **Yazar:** yeni notun yazarı `mount({ author })` / `new WorkbookView(…, { author })` seçeneğidir; verilmezse `"Author"` yazılır (UNSURE: Excel oturum açan kullanıcının adını yazar; var olan notların yazarı yeni nota **verilmez**). Excel'in yeni nota kendiliğinden eklediği kalın "Ad:" satırı eklenmez; metin yazıldığı gibidir, yazar balonda ayrıca gösterilir. Düzenlenen not yazarını korur.
- **Sınırlar ve retler:** metin en çok 32.767 karakterdir (hücre metni sınırı); aşan metin `RangeError` ile reddedilir ve alan açık kalır. **İş parçacıklı yorumu** (yazışma) olan hücrede ekleme, düzenleme ve silme reddedilir (`RangeError`; menüde öğeler devre dışıdır ve nedenini söyler); aynı sayfadaki öbür notlar düzenlenebilir. Salt okunur görünümde (ve dönüştürülmemiş xls'te) not komutları yoktur, yöntemler hiçbir şey yapmaz.
- **Metin kuralı (tek kural):** `undefined` notu siler; notun zaten taşıdığı metin (satır sonu yazımı dışında aynıysa) **değişiklik değildir**; bunların dışında boş ya da yalnız boşluk olan metin notu siler. Yani metni boş bir dosya notuna `''` vermek onu silmez (silmek için `undefined`); notu olmayan hücreye boş metin not eklemez.
- **Birleşik hücre:** birleşik alanın tek notu vardır: sol üst hücresindeki, o yoksa alanın içindeki herhangi bir hücrede duran not (not yazıldıktan sonra birleştirilen hücreler). Alanın hangi hücresi verilirse verilsin o not düzenlenir/silinir; alana ikinci not eklenmez. Yeni not sol üst hücreye konur. (Sol üstte olmayan notun köşe işareti ve balonu gösterilmez; menü ve Shift+F2 onu bulur.)
- **Not işlemlerinin kapalı olduğu sayfalar:** not kutularının çizimi (VML) okunamayan, yorum parçasını başka bir sayfayla paylaşan ya da birden çok yorum parçası olan sayfada notlar gösterilir ama eklenemez, düzenlenemez, silinemez (`sheet.notesBlocked`: `"drawing"` / `"shared"` / `"parts"`; `noteProblem` nedenini verir, menüde öğeler nedeniyle devre dışıdır). Sayfanın geri kalanı düzenlenir ve kaydedilir.
- **Öbür işlemlerle:** satır/sütun eklenip silinince notlar hücreleriyle kayar; hücresi silinen not silinir (geri alınınca döner). Hücre içeriğini silmek (Delete, "İçeriği temizle") notu **silmez**. Sayfa silinince notlarının parçaları da paketten çıkar.
- **Kopyala / kes / yapıştır:** görünümün kendi kopyası (aynı görünümde kopyalanıp yapıştırılan hücreler) hücrelerin **notlarını da taşır** (Excel gibi): metin ve yazar gider; yapıştırılan not yeni bir nottur (düz metin, kutusu yeni notun varsayılan yerinde). Hedef hücrede not varsa kopyadaki not onun **yerini alır**; kopyada notu olmayan hücrenin hedefteki notu **kalır** (Excel "Tümünü yapıştır"da onu da siler — doğrulanamadı, veri kaybetmeyen seçim yapıldı). Kopya daha büyük bir seçime çoğaltılarak yapıştırılırsa notlar da çoğalır. **Kes** kaynak hücrelerin notlarını hücre içeriğiyle birlikte, tek geri alma adımında siler; yapıştırma da tek adımdır. Yorum zinciri kopyalanmaz ve kesilmez; hedef hücrede yorum zinciri varsa ya da sayfanın notları değiştirilemiyorsa (`notesBlocked`) hücreler yine yapıştırılır, notlar atlanır ve kaç notun atlandığı `onNotice` (`kind: 'info'`) ile bildirilir (`PasteResult.skippedNotes`). Sayfadaki not sayısı sınırı (100.000) aşılacaksa yapıştırma tümüyle reddedilir. Notlar **seçimin kendisinden** toplanır: verinin bittiği yerin ötesindeki boş bir hücrede duran not da (hücreler kullanılan alana kırpılsa bile) kopyaya girer, yapıştırılan aralık seçim kadar büyür ve not seçimdeki yerine konur; kes yalnız kopyaya giren notları siler. Sayfa kenarının dışına düşecek not yapıştırılmaz. Süzgeçli listeden yalnız görünen satırların notları kopyalanır (kes de gizli satırların notunu silmez). Başka programdan (ya da sayfa listesi değiştikten sonra) gelen düz metin not taşımaz. Programda: `copyRange` sonucu `notes` alanını taşır, `pasteAt` not işlemini (`setComments`) hücre işlemiyle birlikte döndürür, `cutRange(workbook, sheet, range)` içeriği ve notları birlikte siler, `clearNotes` yalnız notları.
- **Sıralama:** sıralanan aralıktaki her not **satırıyla birlikte** yeni hücresine gider; aralığın dışındaki notlar yerinde kalır; geri alınca notlar eski hücrelerine döner. Dosyadan gelen not taşınınca `<comment>` öğesi biçimleri, yazarı ve tanınmayan öznitelikleriyle korunur, yalnız `ref`'i değişir; kutusu (`x:Row`, `x:Column`, `x:Anchor`) notuyla birlikte kayar (yorum ve VML parçaları, satır/sütun kaydırmada olduğu gibi baştan serileştirilir). Modelde taşınan dosya notu yeni hücresini `row`/`col`/`since` ile, dosyadaki hücresini `from` ile tutar (`SheetComment.from`); konum yine tek işlevle (`commentCellNow`) hesaplanır. **Yorum zinciri** içeren aralık sıralanmaz (zincir satırıyla taşınamıyor); okuyucunun gösteremediği yorumu olan (uyarılı) ya da notları değiştirilemeyen (`notesBlocked`) sayfa da sıralanmaz. "sırala → kaydet → geri al → kaydet" not parçalarını açılan dosyadaki baytlarıyla geri verir.
- **Biçimli notlar:** birden çok biçimli (kalın yazar adı + düz metin gibi) bir notun metni düzenlenirse not tek biçimli düz metin olur; kutusu, yazarı ve "hep göster" ayarı kalır. Düzenlenmeyen notlar biçimleriyle bayt bayt korunur.
- **Dosyaya yazma:** yalnız notları değişen sayfanın not parçaları yeniden yazılır; değişmeyen her `<comment>` öğesi ve her VML şekli (form denetimleri dahil) özgün baytlarıyla, satır sonlarıyla birlikte kalır (parça kes-ekle ile yazılır; satır/sütun kaydırması gereken sayfada yorum parçası baştan serileştirilir). UTF-16 parça UTF-8 olarak yazılır ve bildirimi düzeltilir. İlk notu eklenen sayfaya `xl/commentsN.xml`, `xl/drawings/vmlDrawingN.vml`, ilişkileri ve sayfa XML'inde `<legacyDrawing>` eklenir (içerik türleri parça başına `Override` ile bildirilir); son notu silinen sayfada bunlar kaldırılır (form denetimi taşıyan VML kalır). Yalnız not değiştiren bir düzenleme yeniden hesap istemez (`fullCalcOnLoad` yazılmaz, `calcChain` kalır) ve hücre verisine dokunmaz. Sayfa parçası yalnız ilk notta (`<legacyDrawing>` eklenir) ve son notta (kaldırılır) değişir, başka hiçbir baytı değişmez; çalışma kitabı parçası ve iki sayfanın `tabSelected` işareti yalnız gösterilen sayfa dosyadakinden farklıysa yazılır (her kayıtta olduğu gibi). **"not ekle/düzenle/sil → kaydet → geri al → kaydet" açılan dosyayı bayt bayt geri verir** (arada hücre ya da satır/sütun değişikliği yapılmadıysa). Yeni kutunun yeri ve boyutu sabittir (hücrenin sağ üstü; Excel açınca kendi ölçüsüne göre çizer).

Görünüm olmadan, model üzerinde:

```js
const history = new BymOffice.History(workbook);
history.apply(BymOffice.setComment(workbook, 0, 1, 1, 'Not', { author: 'Ada' }), undefined, undefined);
BymOffice.noteProblem(workbook, 0, 1, 1);          // yapılamıyorsa İngilizce neden, yoksa undefined
BymOffice.commentAtCell(workbook.sheets[0], 1, 1); // { comment, index } | undefined — hücrenin ŞİMDİKİ yerine göre
```

| Üye | İş |
|---|---|
| `setComment(workbook, sheetIndex, row, col, text, { author? })` | İşlemi (`Op[]`) üretir, modeli değiştirmez; değişiklik yoksa boş dizi. Metin kuralı yukarıdadır. Satır sonları `\n` yapılır |
| `noteProblem(workbook, sheetIndex, row, col)`, `commentAtCell(sheet, row, col)` | Ret nedeni; hücrenin (birleşikse birleşik alanın) şimdiki notu |
| `MAX_NOTE_TEXT`, `DEFAULT_NOTE_AUTHOR` | 32767; `"Author"` |
| `Op` türü `setComments` | Sayfanın yorum listesini değiştirir (`comments: SheetComment[] \| undefined`); tersi yine `setComments`'tir. Liste değiştirilmez, yenisiyle değiştirilir |
| `commentCellNow(comment, edits)`, `commentEdits(comment, edits)` | Bir yorumun şimdi durduğu hücre (hücresi silindiyse `undefined`); yorumu kaydıran yapısal düzenlemeler |

**Modelde konum kuralı.** `sheet.comments` dosyadaki yorumları **dosyadaki** hücreleriyle tutar; sonradan eklenen/silinen satır ve sütunlar (`sheet.structuralEdits`) gösterirken ve yazarken uygulanır. Sonradan **eklenen** not ise konduğu andaki hücresini ve o ana dek yapılmış yapısal düzenleme sayısını (`comment.since`) taşır; onu yalnız daha sonraki düzenlemeler kaydırır. Düzenlenen dosya notu yerinde kalır, yalnız metni değişir. Görünüm, `writeXlsx` ve `buildXlsx` aynı işlevi (`commentCellNow`) kullanır. `sheet.commentsModified` notların değiştiğini, `sheet.modified`'dan bağımsız olarak söyler.

### Veri doğrulama kuralları

```js
const view = mounted.view;                                   // WorkbookView
const R = BymOffice.parseRange;
view.setValidation(R('B2:B20'), { type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorTitle: 'Adet', error: '1-10 arası' });
view.setValidation(R('C2:C20'), { type: 'list', formula1: '"Evet,Hayır"' });      // öğeler dosyada HER ZAMAN virgülle ayrılır
view.setValidation(R('D2:D20'), { type: 'list', formula1: 'Listeler!$A$1:$A$5' }); // hücrelerden liste
view.setValidation(R('B5'), undefined);                       // B5'in kuralını kaldırır
view.validationAt(1, 1);                                      // B2'nin şu an geçerli kuralı (aralık ve formüller şimdiki konumda)
view.openValidationDialog();                                  // seçimin iletişim kutusu
```

- **Türler** (`type`): `none` (herhangi bir değer), `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`. **İşleçler** (`operator`; sayı/tarih/saat/uzunlukta): `between` (varsayılan), `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`. `between`/`notBetween` iki değer ister (`formula1`, `formula2`), öbürleri bir. Öbür alanlar: `allowBlank` (boşu yok say; varsayılan `true`), `hideDropDown` (liste oku gizli), `showInputMessage` + `promptTitle`/`prompt`, `showErrorMessage` + `errorStyle` (`stop`/`warning`/`information`) + `errorTitle`/`error`.
- **Formüller dosyadaki gibi yazılır** (API'de): İngilizce fonksiyon adı, `,` ayırıcı, `.` ondalık; baştaki `=` isteğe bağlıdır. Tarih ve saat **seri numarasıdır** (`45657`, `0.5`). Göreli başvurular kuralın **ilk aralığının sol üst hücresine** göredir. İletişim kutusunda ise her şey görünümün dilinde yazılır: `1,5`, `31.12.2024`, `=TOPLA($A$1:$A$3; 0,5)`, liste öğeleri Türkçede `;` ile (`evet; hayır`), İngilizcede `,` ile. Dosya öğeleri virgülle ayırdığı için **virgül içeren liste öğesi** (Türkçede `1,5`) reddedilir.
- **Excel sınırları:** başlık 32, ileti 255, kurala yazılan liste 255 karakter; sayfa başına 65.534 kural. Aşılırsa iletişim kutusu hatayı ilgili alanda gösterir ve açık kalır; `setValidation` `RangeError` fırlatır. Okunamayan formül (`=SUM(A1`) kaydedilmez.
- **Çakışma (Excel davranışı):** kural konan hücreler, daha önce içinde bulundukları kurallardan **çıkarılır** (eski kuralın `sqref`'i dikdörtgenlere bölünür, formülleri yeni ilk hücresine göre yeniden yazılır); eski kural öbür hücreleri için kalır, hücresi kalmayan kural silinir. Seçim birden çok kural (ya da kurallı ve kuralsız hücreler) içeriyorsa iletişim kutusu bunu söyler; Tamam hepsinin yerine yeni kuralı koyar. **Tümünü temizle** seçimdeki kuralları kaldırır ve kutuyu kapatır (Excel'de bu düğme yalnız alanları sıfırlar). "Herhangi bir değer" + gösterilecek giriş iletisi kural olarak **kalır** (Excel gibi); ileti yoksa kural kaldırılır.
- **İletişim kutusu:** modaldır (`role="dialog"`, `aria-modal`), düzenleyicinin tamamını örter; odak içinde kalır (Tab döner), **Esc** vazgeçer, **Enter** uygular (çok satırlı alanda Ctrl/Cmd+Enter), kapanınca odak ızgaraya döner. Açılmadan önce yazılmakta olan hücre/not/sekme adı uygulanır (uygulanamıyorsa kutu açılmaz); **formül yazılırken açılmaz** ("Önce formülü bitirin"). Salt okunur görünümde açılmaz. **Kutu, yazılmakta olan bir alan değildir** (Tamam/Vazgeç'i olan kipli bir kutudur; içindekini onaylamak kullanıcının işidir). `commitPending()` (kaydetme, yazdırma, sayfa işlemleri, köprü): kutuda **hiçbir alan değiştirilmemişse** kutu Vazgeç gibi kapatılır ve `true` döner (kural yöneticisinden açılmış kural kutusu yöneticiye **dönmez**); **bir alan değiştirilmişse** kutu açık kalır, **hiçbir şey uygulanmaz**, kutunun uyarı satırında ve `onNotice` ile "Önce açık kural kutusunu onaylayın ya da kapatın" denir ve `false` döner (kaydedilmez, yazdırılmaz, sayfa işlemi yapılmaz). Bir alanı değiştirip eski değerine döndürmek değişiklik değildir. Yöneticide "değişiklik", bir kuralın **uygulandığı yer** alanına yazılmış ama henüz uygulanmamış (Enter'a basılmamış, alandan çıkılmamış) metindir; yoksa yönetici kapatılır. (Önceki davranış — açık kutuyu Tamam'a basılmış gibi uygulamak — otomatik kaydetmede kullanıcının onaylamadığı kuralı ekliyordu.) **Geri al / yinele** (`undo()`, `redo()`, köprü) kutu açıkken **hiçbir şey yapmaz** ve `false` döner. **Çalışma kitabını değiştiren öbür çağrılar** (`setCellInput`, `paste`, `insertRows`/`deleteRows`, `sort`, `toggleFilter`, `format`, `setValidation`, `addConditionalRule`, …) kutunun açıldığı aralığı ve kural sırasını bayatlatırdı: değişiklik yapılmamış kutu kapatılır ve işlem yapılır; değişiklik yapılmış kutu varken işlem **yapılmaz** (aynı bildirimle), kutu açık kalır. Kutunun yerine başka bir şeyin geçtiği çağrılar kutuyu **kapatır ve içindeki değişikliği atar** (Esc gibi): `setReadOnly(true)`, `setLocale()` (yazılan değer başka dilin kurallarıyla okunmasın), gösterilen sayfayı değiştirme, hücre/not/sekme adı düzenlemeye başlama, bir menü açma. Dosyadan gelen metin yalnız alan değeri/`textContent` olarak basılır.
- **Bakıp çıkmak değişiklik değildir:** kutu açılıp hiçbir alan değiştirilmeden Tamam'a basılırsa işlem üretilmez, dosya bayt bayt aynı kalır. Bir alan değiştirilince yalnız o alan değişir: dokunulmayan değer alanı kuralın dosyadaki yazımını korur (`"a, b,c"` listesi, formüldeki boşluklar), kural aynı hücrelerde düzenleniyorsa **aynı öğe olarak yerinde kalır** (sırası, `imeMode`/`xr:uid` gibi tanınmayan öznitelikleri). Programda da öyle: `setValidation` kuralın tam hücrelerine aynı ayarlarla çağrılırsa boş dizi döner.
- **Modele alınmayan öğeler:** okuyucunun kural saymadığı `<dataValidation>` / `<conditionalFormatting>` öğeleri (okunamayan aralık, kuralsız öğe, 65.534 sınırının ötesi) başka kurallar düzenlenirken de **olduğu gibi** kalır; `count` gerçekten yazılan öğe sayısıdır.
- **Geri al/kaydet:** her kural değişikliği tek geri alma adımıdır ve `onChange` (köprüde `changed`) tetikler. Kaydederken yalnız `<dataValidations>` öğesi güncellenir: dokunulmayan her `<dataValidation>` özgün baytlarıyla kalır, değişen kural **kendi yerinde** yeniden yazılır (tanınmayan öznitelikleri korunarak), yeni kurallar sona eklenir, `count` güncellenir; öğe yoksa şema sırasındaki yerine konur. Hiçbir hücre değişmediği için `fullCalcOnLoad` yazılmaz; "kaydet → geri al → kaydet" özgün dosyayı bayt bayt verir.
- **Öbür işlemlerle:** satır/sütun ekle-sil kuralı hücreleriyle taşır (tümü silinen kural kalkar, geri al getirir); sayfa yeniden adlandırma/silme kural formüllerine yansır (silinen sayfa `#REF!`); **sıralama ve yapıştırma kuralları taşımaz**; sayfa kopyası kuralları alır. Göreli başvuru, kuralın sol üst hücresine göre "şu kadar ötedeki hücre" demektir: kural yerinde kalırken araya sütun girerse `B1` yine `B1`'dir (dosyadaki kurallarla aynı davranış; UNSURE: Excel'le doğrulanmadı).
- **Programda (görünümsüz):** `BymOffice.setValidation(workbook, sheetIndex, range | ranges, rule | undefined)` `Op[]` üretir (`History.apply` ile uygulanır; değişiklik yoksa boş dizi); `validationsIn(workbook, sheetIndex, ranges)`, `validationProblem(…)`, `dataValidationNow(rule, ruleSpace(workbook, sheet))`. İşlem türü `setDataValidations` listeyi değiştirir (tersi yine odur). Modelde düzenlenen/eklenen kural `live: true` taşır: aralığı ve formülleri **şimdiki** satır/sütun ve sayfa adlarıyladır ve işlemler onu güncel tutar; dosyadaki kural dosyadaki haliyle durur (`src`: parçadaki sırası) ve `structuralEdits`/ad eşlemesiyle okunur. İkisi için tek işlev: `dataValidationNow()`.

### Koşullu biçim kuralları

```js
const R = BymOffice.parseRange;
view.addConditionalRule(R('A1:A20'), { type: 'cellIs', operator: 'greaterThan', formulas: ['100'], format: { fill: '#FFC7CE', color: '#9C0006' } });
view.addConditionalRule(R('B1:B20'), { type: 'expression', formulas: ['=MOD(ROW(),2)=0'], format: { bold: true } });
view.addConditionalRule(R('C1:C20'), { type: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] });
view.conditionalRules();                       // [{ index, ranges, rule, format, editable, removable, movable }, …] uygulanma sırasıyla
view.updateConditionalRule(0, { ranges: R('C1:C50'), stopIfTrue: true });
view.moveConditionalRule(2, 0);                // 3. kuralı en öne alır
view.removeConditionalRule(1);
view.clearConditionalRules(R('A1:B5'));        // bu hücreleri her kuraldan çıkarır; aralıksız: sayfanın bütün kuralları
view.openConditionalDialog('manage');          // kural yöneticisi; 'greaterThan' | 'between' | 'containsText' | 'top10' | 'colorScale' | … yeni kural; sayı: o kuralı düzenle
```

- **Oluşturulan/düzenlenen türler** (`type`): `cellIs` (8 işleç), `containsText`, `notContainsText`, `beginsWith`, `endsWith` (`text`), `duplicateValues`, `uniqueValues`, `top10` (`rank`, `percent`, `bottom`), `aboveAverage` (`aboveAverage: false` = altı, `equalAverage`), `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`, `expression` (`formulas: [formül]`), `colorScale` (`colors`: 2 ya da 3 renk; en küçük → [orta: %50 yüzdelik] → en büyük), `dataBar` (`colors: [renk]`; en küçük–en büyük, uzantısız temel biçim). Metin, boş ve hata kurallarının formülü Excel'in yazdığı kalıpla üretilir (`NOT(ISERROR(SEARCH("x",A1)))` …). **İkon kümesi ve tarih dönemi** oluşturulmaz, içeriği değiştirilmez; dosyadakiler listelenir, aralığı/"doğruysa durdur"u değiştirilebilir, taşınır, silinir.
- **Bakıp çıkmak değişiklik değildir; değişen yalnız değişendir:** kural kutusu açılıp hiçbir alan değiştirilmeden Tamam'a basılırsa işlem üretilmez. Bir alan değişince kuralın geri kalanı — kutunun gösteremediği her şey dahil — olduğu gibi kalır: renk ölçeğinin eşik türleri ve değerleri (`num 0`, `formula $H$2`, `percent 90`), veri çubuğunun `minLength`/`maxLength`/`showValue`'su ve eşikleri, `stdDev`, formüllerin dosyadaki yazımı, dokunulmayan renk (tema rengi olarak kalır). Metin kuralının formülü yalnız aranan metin değişince yeniden üretilir ve eski formülün andığı hücre yazımıyla (`$A1`) üretilir; formülü Excel'in kalıbına uymayan metin kuralı kutuda **Formül** olarak gösterilir. Programda `updateConditionalRule(index, { rule })` da bir yamadır: tür aynı kaldıkça verilmeyen ya da kuraldakiyle aynı verilen her şey kuraldaki gibi kalır; hiçbir şey değişmiyorsa boş dizi döner. `{ format }` yalnız biçimi değiştirir; `format: null` biçimi kaldırır, verilmezse biçim kalır.
- **Biçim** (`format`): `color` (yazı rengi), `fill` (dolgu) — `#RRGGBB`; `bold`, `italic`, `underline`, `strike`. İletişim kutusunda üç hazır biçim (açık kırmızı dolgu/koyu kırmızı metin, sarı/koyu sarı, yeşil/koyu yeşil), **Özel biçim**, **Biçim yok** ve — biçimi hazır üçlüden biri olmayan kuralda, seçili gelen — **Var olan biçim** (yalnız kenarlık, sayı biçimi ya da gradyan dolgudan oluşan biçim de budur; başka bir biçim seçilmedikçe kayıt aynen kalır). Özel biçimde değiştirilen özellik değişir, öbürleri dosyadaki haliyle kalır (tema rengi, açıkça kapalı yazılmış kalın `<b val="0"/>`; işaretsiz bırakılan kutu "değişmedi" demektir, işaretlenip kaldırılan özellik kayıttan çıkar). Bir kitapta en çok 65.536 biçim kaydı olabilir; aşacak kural reddedilir. Biçim `styles.xml`'in `<dxfs>` tablosuna **eklenir** (var olan kayıtlar yerinde kalır); aynı biçim tek kayıttır, dosyadaki eş kayıt yeniden kullanılır. Bir kuralın biçimi değiştirilince kayıttaki öbür her şey (kenarlık, sayı biçimi, yazı boyutu, hizalama) **korunur**; kenarlık ve sayı biçimi düzenlenmez. Kayıtlar her kayıtta yeniden hesaplanır: kural geri alınınca kayıt da çıkar ve `styles.xml` özgün baytlarına döner (dosyada `styles.xml`/`<dxfs>` yoksa oluşturulur, geri alınca kaldırılır).
- **Sıra:** yeni kural **en öne** gelir (Excel gibi), öbürleri birer kayar; her değişiklikte `priority` değerleri sayfa genelinde 1..n olarak yeniden numaralanır. Kaydederken sırası değişen ama başka bir şeyi değişmeyen `<cfRule>` öğesinde **yalnız `priority` özniteliği** yeniden yazılır; dokunulmayan öğeler bayt bayt kalır. Bir kural değişince kendi `<conditionalFormatting>` öğesine taşınır (aynı öğedeki öbür kurallar yerinde kalır; tanınmayan öznitelikleri ve `<extLst>`'si korunur). **Aralık değişince** kuralın formülleri yeni sol üst hücreye göre kaydırılır (her hücre için anlamı aynı kalır; UNSURE: Excel "Uygulandığı yer" değişince formül metnini değiştirmez).
- **2010 uzantı listesi (x14):** **Uzantıda karşılığı olan kural** (Excel 2010+'ın veri çubukları; ana bölümdeki `<cfRule>` ile `extLst`'teki `x14:cfRule` bir kimlikle bağlıdır) silinebilir, aralığı ve sırası değiştirilebilir, hücreleri temizlenebilir: yazıcı eşini de birlikte keser ya da `xm:sqref`'ini yeni hücrelerle yazar (boş kalan `x14:conditionalFormattings`/`ext` sarmalayıcısı ve boş `extLst` kaldırılır; eş `priority` taşıyorsa o da güncellenir); içeriği (renk, eşik) değiştirilemez (`editable: false`, `RangeError`). Ana `<cfRule>` kaynak metniyle yazılır. **Yalnız uzantıda duran kural** (x14 ikon kümeleri vb.) hiç değiştirilmez (`editable`/`removable` `false`); böyle bir kuralı olan sayfada **sıra değiştirilemez** (`movable: false`) ve yeni kural **sona** eklenir (o kuralın numarası uzantıda durduğu için öbürleri yeniden numaralanmaz). `extLst`'in geri kalanı (mini grafikler, öbür x14 kuralları) bayt bayt kalır; her şey geri alınınca özgün baytlar. Yöneticide kapalı düğmelerin nedeni ipucunda yazar.
- **Yönetici** (`openConditionalDialog('manage')`, araç çubuğu menüsünde "Kuralları yönet…"): geçerli seçimin ya da sayfanın kuralları uygulanma sırasıyla; her satırda kuralın özeti, **uygulandığı yer** alanı (`A1:B10; D1:D5` — `;`, `,` ya da boşlukla; `A:C`, `2:5`, `$` kabul edilir; Enter ya da alandan çıkınca uygulanır, geçersizse alanda hata), **Doğruysa durdur**, ↑ ↓ Düzenle Sil; "Yeni kural…". **Her değişiklik anında uygulanır ve kendi geri alma adımıdır** (Excel'deki "Uygula"ya kadar bekleme yok). Liste kuralları **200'er 200'er** gösterir (`MAX_MANAGER_ROWS`): fazlası varsa altta hangilerinin gösterildiği yazar ("450 kuralın 201–400. olanları") ve **Önceki 200** / **Sonraki N** düğmeleriyle öbürlerine geçilir (`data-bym-action="previous-page"` / `"next-page"`); böylece 201. ve sonraki kurallar da düzenlenir, silinir, taşınır (bir bölümün ilk kuralı ↑ ile önceki bölümün son kuralının üstüne geçer). Son bölümün kuralları silinince liste bir önceki bölümü gösterir. Düzenle/Yeni kural kural iletişim kutusunu açar; o kapanınca yöneticiye dönülür. İletişim kutularının klavye/odak/`commitPending` davranışı veri doğrulamadakiyle aynıdır.
- **Araç çubuğu menüsü:** Büyüktür…, Küçüktür…, Arasında…, Eşittir…, Metin içerir…, Yinelenen değerler…, İlk 10 öğe…, Renk ölçekleri…, Veri çubukları…, Seçili hücrelerden kuralları temizle, Bu sayfadan kuralları temizle, Kuralları yönet….
- **İletişim kutusundaki değerler** görünümün dilinde yazılır: sayı (`1,5`), tarih (seri numarasına çevrilir), `=` ile formül (`=TOPLA($C$1:$C$2; 0,5)`); başka her şey **metin** sayılır ve tırnak içinde saklanır (`bitti` → `"bitti"`).
- **Öbür işlemlerle:** veri doğrulamadakiyle aynı (satır/sütun, sayfa adı/silme, sıralama ve yapıştırma taşımaz, sayfa kopyası alır). Eşik formülleri (`cfvo`) hücre formülleri gibi izler.
- **Programda:** `BymOffice.addConditionalRule(workbook, sheetIndex, ranges, settings)`, `updateConditionalRule(workbook, sheetIndex, index, { ranges?, stopIfTrue?, rule? })`, `removeConditionalRule`, `moveConditionalRule`, `clearConditionalRules`, `conditionalRules(workbook, sheetIndex, ranges?)`, `conditionalFormatNow(format, ruleSpace(workbook, sheet))`. İşlem türü `setConditionalFormats`. Modelde düzenlenen/eklenen biçim `live: true` taşır ve biçimini kuralın kendisinde tutar (`rule.dxf`; varsa başladığı kayıt `rule.dxfId`); dosyadaki kurallar `src` ile parçadaki yerlerini bilir. `buildXlsx()` (ve xls → xlsx) kuralları ve biçim kayıtlarını modelden yazar.

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
| `activeSheet` | Açılışta gösterilen sayfanın dizini. Tablo görünümü, kullanıcı başka bir (görünür) sayfayı gösterince günceller; xlsx `save()` başka bir değişiklik de varsa dosyaya yazar |

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
- Paket ES2019 sözdizimine indirilmiş olarak derlenir. İki çıktıda da `import.meta` yoktur.
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
- xlsx dosyalarında **tablo görünümü**: gerçek ızgara (kaydırma, seçim, klavye, sayfa sekmeleri, dondurulmuş bölmeler) ve **hücre düzenleme** (yazma, geri al/yinele, kopyala/yapıştır, satır/sütun ekle-sil; üstteki düğmeler ve "Salt okunur" seçimi); **hücre notları** (sağ tık → Not ekle / Notu düzenle / Notu sil, ya da Shift+F2); araç çubuğundaki **Yazdır** düğmesi (ya da Ctrl/Cmd+P) etkin sayfayı sayfa yapısına göre yazdırır; altında seçili aralık, etkin hücrenin görünen değeri, ham değeri ve formülü yazar. "Yeniden kaydet ve indir" düzenlemelerinizi dosyaya yazar; indirilen dosyayı Excel'de açıp denetleyebilirsiniz. "Sayı ve tarih biçimi dili" seçimiyle Türkçe ve İngilizce gösterim arasında geçilir.
- docx dosyalarında **belge görünümü**: sayfalar ve belge düzenleyici (yazma, seçim, biçimlendirme araç çubuğu, geri al/yinele, pano); doc ve RTF salt okunur gösterilir. Korumalı belge salt okunur açılır; izleme isteyen belgede düzenlemeler **"Yazar adı"** alanındaki adla izlenir (yazılan altı çizili, silinen üstü çizili) ve görünümün üstünde koruma/izleme durumu yazar. Belge çubuğundaki **Yazdır** düğmesi (ya da Ctrl/Cmd+P) tarayıcının yazdırma penceresini açar.
- "Yeniden kaydet ve indir" ile kaydedilen dosyayı indirip Word/Excel'de açmak.
- Bir xls dosyası seçiliyken "xlsx olarak indir" ile `saveAsXlsx()` çıktısını indirmek.
- "Örnek xlsx üret ve aç" ile elinizde dosya olmadan denemek.
- "Boş docx oluştur ve aç" / "Boş xlsx oluştur ve aç" ile `createBlank()` çıktısını açmak.
- "Eski cihaz kipi" ile tarayıcının yerleşik sıkıştırmasını kapatıp motorun kendi kodunu sınamak.
- **"Font dosyası"** ile kendi font dosyanızı (ttf, otf, woff, woff2) yanındaki alana yazdığınız belge font adının (varsayılan `Bookman Old Style`) yerine kaydetmek (`registerFont`); açık belge ve tablo kendiliğinden yeniden dizilir. "Fontları kaldır" kayıtları siler.
- **"Paketle gelen font"** kutusu TeX Gyre Bonum'u açıp kapatır (`BymOffice.fonts.configure({ baseUrl: null })`); yanındaki satır durumunu ve dosyaların isteneceği adresi gösterir. Bookman Old Style kullanan bir belge açınca "yüklendi" olur.

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
