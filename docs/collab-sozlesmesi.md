# Birlikte düzenleme sözleşmesi

Bu belge, bir çalışma kitabını (xlsx) birlikte düzenleyen tarayıcılar ile aralarındaki **aktarıcı** (relay) arasındaki iletileri tanımlar. Aktarıcıyı kendi arka ucunuzda yazacaksanız (Laravel Reverb, Pusher, bir Node süreci, başka bir şey) gereken her şey buradadır. Depodaki hazır örnek: `server/collab-server.mjs` (bağımlılıksız Node; **bellekte** çalışır, hiçbir şeyi diske yazmaz). Aynı kuralların sayfa içi sürümü: `BymOffice.collab.memoryServer()`.

Aşama 4a yalnız **tabloları** kapsar. Belgeler (docx) birlikte düzenlenmez.

## Güvenlik: önce bunu okuyun

**Odaya giren herkes dosyanın içeriğini okur ve değiştirebilir.** Aktarıcı kimin girebileceğine karar veren tek yerdir: üretimde `authorize` (ve `origins`) **zorunludur**; örnek sunucu ikisi de verilmeden başlatılırsa açık bir uyarı yazar. Motor, odadaki öbür katılımcılara ve aktarıcıya **güvenmez**: gelen her adım şemaya göre okunur, dosyaya yalnız metin olarak girebilecek değerler kabul edilir, bir katılımcının "kaydedildi" demesi sunucuya sorulmadan benimsenmez. Ama odaya girmiş biri, kendisine verilen yetkiyle hücreleri **değiştirebilir** — bu, yetkilendirmenin işidir, protokolün değil.

## Fikir

- **Aktarıcı aptaldır.** İşlemlerin içine bakmaz, dosyayı tutmaz, biçimi bilmez. Tek işi sıra vermektir: oda başına bir sayaç (`rev`) ve bir günlük.
- **Dosya aktarıcıda durmaz.** Herkes dosyayı kendi adresinden açar (`file`); günlük, dosyanın açıldığı sürümden **sonraki** işlemlerdir. Bu yüzden oda, dosyanın bir sürümüne bağlıdır: katılan, dosyasının neyle tanındığını söyler (`v`: ETag ve/veya içerik özeti).
- **Birleştirme istemcidedir.** İstemci kendi değişikliğini hemen uygular ve kuyruğa koyar; aynı anda en çok bir adım yoldadır. Aktarıcı bir adımı ancak **odanın o anki `rev`'i üzerine** yapılmışsa kabul eder. Böylece her adım, yazarının uyguladığı kitabın aynısına, herkeste aynı sırayla uygulanır: yakınsama için başka hiçbir şey gerekmez. Araya başkasının adımı girdiyse istemci bekleyen adımlarını geri alır, gelen adımı uygular, kendi adımlarını yeni duruma göre **yeniden temellendirir** (aşağıdaki tablo) ve yeniden gönderir; artık uygulanamayan adım düşer ve kullanıcıya söylenir.

## Dört kural (aktarıcı)

1. **Oda** bir `rev` (0'dan başlar), bir `floor` (günlükteki en eski adımdan önceki sürüm) ve bir günlük tutar.
2. **`op`:** `base === rev` ise kabul: `rev++`, günlüğe yaz, odadaki **herkese (gönderen dahil)** yayınla. Değilse yalnız gönderene `nack`.
3. **`join`:** katılana `hello` ve bildiği sürümden sonraki günlük gönderilir; günlük o sürümü kapsamıyorsa `reset` (istemci dosyayı yeniden yükleyip yeniden katılır).
4. **Hafif iletiler** (`presence`, `saved`) günlüğe girmez, sıra numarası almaz, öbürlerine aktarılır.

Ek kurallar (aktarıcı):

5. **Sürüm (`pv`):** `join` protokol sürümünü taşır (şimdi `1`). Odanın sürümü ilk katılanla sabitlenir; başka sürümle gelen `error version` alır. (Yeni bir sürümün tanıdığı işlem türü eskilerde atlanırdı; aynı odada iki sürüm ıraksama demektir.)
6. **Kimlik öneki (`key`):** `join` bir `key` taşır (1–40 karakter, `[A-Za-z0-9_-]`); o bağlantının gönderdiği her `op.id` `"<key>:"` ile başlamalıdır, değilse `nack reason:"id"`. Aynı odada o an bağlı başka bir bağlantının `key`'i ile gelen `error key-taken` alır (istemci birkaç kez yeniden dener: aktarıcı kopan bağlantıyı henüz fark etmemiş olabilir). İstemci ayrıca kendi adımını yalnız adından değil **içeriğinden** de tanır: aynı adla gelen başka içerik başkasının adımıdır.
7. **Bağlantı başına tek oda:** katılmış bağlantıdan ikinci `join` → `bad-message`. `reset` almış bağlantı yalnız aynı odaya yeniden katılabilir. Sırası bozuk adım gören istemci yeni bağlantı kurup `since` ile katılır (dosya yeniden yüklenmez).
8. **Sürüm eşlemesi değişmez:** bir sürüm dizgesinin hangi adımı tuttuğu bir kez yazılır (ilk yazan kazanır; odanın kuruluş sürümleri 0'a kilitlidir). `saved` yalnız yazma yetkili üyeden kabul edilir ve `rev` yalnız gönderenin gerçekten gördüğü aralıkta (`katıldığı adım ≤ rev ≤ odanın adımı`).
9. **`sync`:** biri sürümü bilinmediği için katılamazsa (`reset`), aktarıcı odadakilere `{ "t":"sync" }` der (en çok 5 sn'de bir). Yazma yetkili üyelerden **biri** (aktarıcının verdiği kimliği en küçük olan) dosyayı kaydeder ve `saved { v, rev }` duyurur; yeni gelen dosyayı yeniden yükleyip katılır. Kimse kaydedemiyorsa yeni gelen artan aralıklarla altı kez dener ("şu an katılınamıyor: odadakilerden biri kaydetmeli" bildirimi), sonra oturumu kapanır.
10. **Sınırlar (örnek sunucu):** `op.ops` 1–10.001 öğe; bağlantı başına işlem hızı (jeton kovası; aşan `nack reason:"rate"`, istemci bir saniye sonra yeniden dener) ve varlık hızı; toplam bağlantı, IP başına bağlantı, bütün odaların günlükleri için toplam bellek.

Günlük sınırlıdır (örnek sunucuda oda başına 20.000 adım / 32 MiB); aşınca en eski adımlar atılır ve `floor` ilerler. **Kaydedilmeden** biriken günlük budanırsa, eski dosyayla sonradan gelen katılamaz (`reset` alır, dosya hâlâ eskidir, oturum kapanır). Sınırı, iki kayıt arasındaki düzenleme hacmine göre seçin.

## İletiler

Her ileti tek bir JSON nesnesidir (WebSocket'te bir metin çerçevesi). En büyük ileti 1 MiB'dır. Tanınmayan ya da alanı yanlış ileti: aktarıcı `error bad-message` gönderip bağlantıyı kapatır; istemci oturumu kapatır.

### İstemci → aktarıcı

| İleti | Alanlar |
|---|---|
| `join` | `{ "t":"join", "pv":1, "key":dizge, "v":[dizge…], "since"?:tamsayı, "user"?:{ "name"?:dizge, "color"?:"#RRGGBB" }, "token"?:dizge, "ro"?:true, "room"?:dizge }` |
| `op` | `{ "t":"op", "base":tamsayı, "id":dizge, "ops":[…] }` |
| `presence` | `{ "t":"presence", "data":{ "s":sayfa, "a":[satır,sütun], "r":[üst,sol,alt,sağ] } \| null }` |
| `saved` | `{ "t":"saved", "v":[dizge…], "rev"?:tamsayı }` |
| `ping` | `{ "t":"ping" }` |

- `join` bağlantının **ilk** iletisidir. `v`: dosyanın neyle tanındığı — 0–8 dizge, her biri 1–200 karakter. Motor şunları gönderir: `e:<ETag>` (dosya bir adresten ETag ile geldiyse) ve `h:<uzunluk>:<16 onaltılık>` (baytların özeti; gizli değildir, kanıt değildir). `since`: istemcinin elindeki son `rev` (yeniden bağlanırken). `token`: yalnız bu iletide gider; motor onu hiçbir bildirime ya da günlüğe yazmaz. `ro`: yalnız izler.
- `op.id`: istemcinin kendi adımını tanıdığı ad (`<oturum anahtarı>:<sıra>`, en çok 100 karakter). `ops`: aktarıcı için **opak** bir JSON dizisi (1–10.001 öğe); biçimi aşağıda.
- `saved.rev`: kaydedilen dosyanın tuttuğu adım. Verilmezse aktarıcı yalnız iletiyi aktarır (öbürleri yeni ETag'i öğrenir), sürüm tablosuna yazmaz.

### Aktarıcı → istemci

| İleti | Alanlar |
|---|---|
| `hello` | `{ "t":"hello", "pv":1, "rev":tamsayı, "from":tamsayı, "you":dizge, "ro":mantıksal, "v":[dizge…], "peers":[{ "id", "user":{name?,color?}, "ro", "presence"? }…] }` |
| `op` | `{ "t":"op", "rev":tamsayı, "id":dizge, "from":dizge, "ops":[…] }` |
| `nack` | `{ "t":"nack", "id":dizge, "rev":tamsayı, "reason"?:"read-only"\|"id"\|"rate" }` |
| `sync` | `{ "t":"sync" }` — biri katılamıyor: kaydedebilen kaydetsin |
| `reset` | `{ "t":"reset", "rev":tamsayı, "v":[dizge…] }` |
| `presence` | `{ "t":"presence", "from":dizge, "data":… }` |
| `saved` | `{ "t":"saved", "from":dizge, "v":[dizge…], "rev"?:tamsayı }` |
| `peer` / `left` | `{ "t":"peer", "id", "user", "ro" }` / `{ "t":"left", "id" }` |
| `pong` | `{ "t":"pong" }` |
| `error` | `{ "t":"error", "code":dizge }` — ardından bağlantı kapanır (`bad-message`, `unauthorized`, `room-full`, `server-full`, `join-timeout`, `version`, `key-taken`) |

`hello.from`: günlüğün oynatılmaya başlandığı sürüm; `hello`'dan hemen sonra `rev > from` olan bütün adımlar sırayla gelir.

### `join` nasıl yanıtlanır (sırayla)

a. `since` verilmiş ve `floor ≤ since ≤ rev` → `hello` (`from: since`) + sonraki adımlar. (Taze odada `since: 0` bu kurala girmez; aşağıdakilere bakılır.)
b. `v`'den biri odanın sürüm tablosunda ve o sürümün adımı `r ≥ floor` → `hello` (`from: r`) + sonraki adımlar (birden çok eşleşmede en büyük `r`).
c. Oda **taze** (`rev === 0` ve başka katılmış bağlantı yok) → tablo temizlenir, `v`'nin her dizgesi 0 olarak yazılır → `hello` (`from: 0`).
d. Hiçbiri → `reset`. Bağlantı açık kalır ama odada sayılmaz; istemci dosyasını yeniden yükleyip aynı bağlantıdan yeniden `join` gönderir.

`saved { v, rev }` gelince (`floor ≤ rev ≤ oda rev'i`) `v`'nin her dizgesi tabloya o `rev` ile yazılır (en çok 16 dizge; en eskisi atılır).

### Oda anahtarı, yetki

Örnek sunucuda oda = isteğin **yol adı** (sorgu hariç) + varsa `join.room`. Yani `wss://example.com/collab/42` adresini açan herkes aynı odadadır. Yetki sizdedir: örnek sunucunun `authorize({ room, token, origin, request })` kancası `true` / `false` / `{ readOnly: true }` döndürür; `origins` seçeneği izinli `Origin` listesidir. Kendi aktarıcınızda `token`'ı (ya da çerezinizi) doğrulayın: **odaya giren herkes dosyanın içeriğini öğrenir ve değiştirebilir.**

## `ops`: bir adımın içeriği

`ops`, tel işlemlerinin dizisidir; hücre biçimi taşıyan adımda **ilk öğe** biçim tablosudur: `{ "t":"st", "f":[biçim…] }`. İşlemlerdeki biçim alanları (`s`) bu tablonun dizinleridir. Biçim: `{ "n": sayı biçimi (yerleşik biçimin numarası 0–163 ya da kod dizgesi), "f": font, "l": dolgu, "b": kenarlık, "x"?: adlandırılmış stil dizini, "a"?: hizalama, "p"?: koruma, "q"?: true }` (alanlar modeldeki `Font`, `Fill`, `Border`, `Alignment`, `Protection` ile aynıdır; geçişli dolgu `{ "type":"gradient", "xml":"…" }`).

Bütün konumlar 0 tabanlıdır ve **şimdiki** durumdadır (dosyadaki değil). `s` sayfa dizinidir. Aralık `[üst, sol, alt, sağ]`.

| `t` | Alanlar | Anlam |
|---|---|---|
| `cells` | `c: [[satır, sütun, yama]…]` | Yama `0`: hücre gider. Değilse `{ "v"?: içerik, "s"?: biçim }`: verilmeyen yan olduğu gibi kalır (biçim ile değer ayrı ayrı değişir). İçerik: `{ "k": "blank"\|"number"\|"string"\|"boolean"\|"error"\|"date", "v": değer, "f"?: { "x": formül metni ("=" yok), "t"?: "array"\|"dataTable", "r"?: aralık, "a"?: [[ad,değer]…] }, "r"?: zengin metin parçaları, "ph"?: true }`. Ortak (shared) formül, hücresi için açık yazılmış gider. |
| `rows` | `r: [[satır, { "h"?: sayı\|null, "ch"?: mantıksal, "hid"?: mantıksal, "s"?: biçim\|null }]…]` | Yalnız adı geçen özellik değişir; `null` yüksekliği/biçimi kaldırır. |
| `cols` | `p: [[ilk, son, ayar \| 0]…]` | O sütunların ayarının tamamı: `{ "w"?: genişlik, "cw", "hid", "s"?: biçim, "ol": 0–7, "col" }`; `0`: ayarsız. |
| `merges` | `add: [aralık…], del: [aralık…]` | Birleştir / ayır. Var olan bir birleşimi kesen `add` reddedilir. |
| `filter` | `f: süzgeç \| 0, show?: [satır…], hide?: [satır…]` | Süzgecin model alanları (`range`, `columns`, `sortState`, `held`); gösterilecek/gizlenecek satırları **gönderen** hesaplar (yerel ayardan ve hesaplanan değerlerden bağımsız uygulanır). Dosyanın modelde olmayan içeriği (`raw`, `attrs`, `extra`) telde gitmez; alan aynıysa herkes kendi dosyasındakini korur. |
| `sort` | `g: aralık, o: [dizin…]` | `g.üst + i` satırı, `g.üst + o[i]` satırının hücrelerini alır (`o` bir permütasyondur). |
| `ins` / `del` | `x: "row"\|"col", at, n` | Satır/sütun ekle, sil. |
| `notes` | `n: [[satır, sütun, not \| 0]…]` | Not: `{ "x": metin, "a"?: yazar, "r"?: parçalar, "vis"?: true }`; `0` siler. Notlar şimdiki hücreleriyle adlandırılır. |
| `sheet+` | `name` | `s` dizinine boş sayfa. |
| `sheet-` | | Sayfa sil. |
| `name` | `name` | Yeniden adlandır. |
| `move` | `to` | Taşı. |
| `state` | `state: "visible"\|"hidden"\|"veryHidden"` | Göster/gizle. |
| `tab` | `color: renk \| 0` | Sekme rengi. |
| `cmd` | `name, a: [bağımsız değişkenler]` | Kural ve grafik **komutu**; her katılımcı kendi kitabında çalıştırır: `setValidation`, `addConditionalRule`, `updateConditionalRule`, `removeConditionalRule`, `moveConditionalRule`, `clearConditionalRules`, `addChart`, `updateChart`, `moveChart`, `removeChart` (bağımsız değişkenler `src/ops/rules.ts` ve `src/ops/charts.ts` işlevlerinin kitap ve sayfadan sonrakileridir). |

Kurallar ve grafikler neden komut olarak gider: modelde bu listeler, her katılımcının **kendi okuduğu dosyanın** öğelerine (kaynak öğe dizini, dosyada korunan içerik) başvurur; sonradan kaydedilmiş dosyayla katılan biri için bunlar başka sayılardır. Komut herkeste aynı anlamı taşır.

Gelen her adım **güvenilmezdir**: her alan türüne ve aralığına bakılarak yeni bir nesneye okunur, uygulanırken modelin kendi denetimlerinden geçer (sayfa sınırı, dizi formülü, birleşik hücre, sayfa adı, komutların reddettikleri). Şemadan geçmeyen ya da kitabın kabul etmediği adım **herkeste aynı biçimde atlanır** (doğrulama yalnız adımın kendisine ve o adımdaki kitaba bakar; günlüğü sonradan oynatan da atlar), oda yaşamaya devam eder ve kullanıcıya bir kez söylenir. Dürüst bir istemcinin adımı atlanmaz: yazarı onu aynı kitaba uygulamıştır.

Şemanın kabul etmedikleri (dosyaya metin dışında bir şey olarak girebilecek her şey):

- XML'de yer alamayan karakter taşıyan **her** dizge (sekme ve satır sonu dışındaki denetim karakterleri, yarım vekil çift, `U+FFFE/FFFF`) — hücre metni, sayfa adı, not, yazar, font adı, sayı biçimi kodu, süzgeç değeri, komut bağımsız değişkeni.
- Formülde `t` (dizi, veri tablosu, ortak) ve `r`; formül özniteliklerinden yalnız `aca`, `ca`, `bx` (değeri `1/0/true/false`) kabul edilir. Dizi formülü ve veri tablosu **telde taşınmaz**: böyle bir hücrenin içeriğini değiştiren yerel adım yapılmaz ve söylenir (biçimi değişebilir); geri alması böyle bir hücreyi geri getirecek adım geri alınamaz.
- Dolgu deseni ve kenarlık stili şemanın adlarından biri olmalıdır; **geçişli dolgu taşınmaz** (dosyada ham öğe olarak durur): geçişli dolgulu hücrenin biçimini değiştiren yerel adım yapılmaz ve söylenir.
- `date` değeri ISO 8601, `error` değeri bilinen hata sabitlerinden biri; süzgeç işleci altı karşılaştırmadan biri.
- Adım başına en çok 2.000 birleştirme (ekle ve sil ayrı ayrı), 5.000 not, 1.024 sütun koşusu, 256 biçim; bir adım en çok 10.000 işlem ve 900 KB. Reddedilen adım stil tablosuna hiçbir şey eklemez. (Stil tablosu için oda genelinde bütçe **yoktur**: katılımcılar farklı dosya sürümlerinden başlayabildiği için herkeste aynı sayılamıyor.)

## Yeniden temellendirme kuralları

Bekleyen (henüz sıraya girmemiş) bir adım, araya giren adımın içinden okunur. "Kayar": konum, eklenen/silinen satır-sütunla ve sayfa diziniyle birlikte yer değiştirir; formül metinleri aynı kurallarla yeniden yazılır (`adjustFormula`, `renameSheetsInFormula`; silinen sayfa `#REF!`). Bir adım bütün olarak kalır ya da bütün olarak düşer (düşen adım kullanıcıya söylenir); yalnız hücre/not/satır/sütun listelerinde hedefi giden öğeler çıkarılıp kalanı uygulanır (bu da söylenir).

| İşlem | Araya giren: satır/sütun ekle-sil (aynı sayfa) | Araya giren: sayfa ekle/sil/taşı | Araya giren: aynı şeye dokunan |
|---|---|---|---|
| `cells` | Hücreler kayar; silinen satır/sütundakiler çıkar (hepsi çıkarsa adım düşer). Formüller yeniden yazılır (başka sayfadaki düzenleme için de, adla anıyorsa) | Dizin kayar; sayfa silindiyse düşer. Yeniden adlandırmada formüllerdeki ad değişir | Aynı hücre: **sonra gelen kazanır** (bekleyen adım sonra uygulanır). Biçim ile değer ayrı yamalardır: ikisi de kalır. Sıralanan aralıktaki hücre satırıyla birlikte gider |
| `rows` / `cols` | Satırlar/sütunlar kayar; gidenler çıkar | Dizin kayar; silindiyse düşer | Sonra gelen kazanır |
| `merges` | Aralıklar kayar/daralır; tek hücreye inen çıkar | Dizin kayar; silindiyse düşer | Var olan birleşimi kesen ekleme düşer |
| `filter` | Aralık ve ölçütler `shiftAutoFilter` ile kayar; başlık satırı gittiyse düşer; satır listeleri kayar | Dizin kayar; silindiyse düşer | Aynı sayfada iki süzgeç değişikliği: sonraki **düşer** (satır listeleri eski duruma göre hesaplanmıştır) |
| `sort` | Aralığın üstündeki satırlar aralığı kaydırır; aralığın içine giren/çıkan satır varsa **düşer**; sütunlar aralığı genişletir/daraltır | Dizin kayar; silindiyse düşer | Kesişen iki sıralama: sonraki düşer |
| (her biri) × araya giren `sort` | `cells`, `notes`: sıralanan sütunlardaki hücre satırıyla gider. `del` (satır): silinecek satırlar sıralamadan sonra durdukları yerlerden silinir (gerekirse birkaç parça halinde); sıralama sayfanın kullanılan bütün sütunlarını kapsamıyorsa (satırın yarısı taşınmışsa) **düşer**. `ins`, `rows`, `cols`, `merges`, `filter`: yerinde kalır | — | — |
| `ins` | Yer kayar; aynı yere iki ekleme: ikisi de eklenir (sonra gelenin satırları üstte); başkasının sildiği aralığın **içine** yapılan ekleme silinen yerin başına iner (eklenen satırlar kalır) | Dizin kayar; silindiyse düşer | — |
| `del` | Aralık kayar; başkasının sildiğiyle örtüşen kısım düşülür, hepsi silinmişse adım **sessizce** biter (istenen zaten olmuş); aralığın **içine** başkasının eklediği satırlar aralıkla birlikte gider (o ekleme, sonraki bekleyen adımlar için yok sayılır) | Dizin kayar; silindiyse sessizce biter | — |
| `notes` | Hücreler kayar; gidenler çıkar | Dizin kayar; silindiyse düşer | Aynı hücre: sonra gelen kazanır |
| `sheet+` | — | Yer kayar | Aynı ad: sonraki düşer (ad alınmış) |
| `sheet-` | — | Dizin kayar; zaten silinmişse sessizce biter | — |
| `name`, `state`, `tab` | — | Dizin kayar; silindiyse düşer | Sonra gelen kazanır (ad alınmışsa düşer) |
| `move` | — | İki dizin de kayar; silindiyse düşer | Sonra gelen uygulanır |
| `cmd` `setValidation`, `addConditionalRule`, `clearConditionalRules` | Aralıklar kayar, formüller yeniden yazılır; aralığın hepsi gittiyse düşer | Dizin kayar; silindiyse düşer | Sonra gelen uygulanır |
| `cmd` `update/remove/moveConditionalRule` (kuralı **sırasıyla** adlandırır) | Aynı sayfada satır/sütun değiştiyse **düşer** (bir kural gitmiş, sıralar kaymış olabilir) | Dizin kayar; silindiyse düşer | Aynı sayfada başka bir koşullu kural komutu: **düşer** |
| `cmd` `addChart`, `updateChart`, `moveChart` | Kitabın **herhangi bir** sayfasında satır/sütun değiştiyse **düşer** (başvurular ve çapalar yeniden yazılmıyor) | Sayfa silme/adlandırmada düşer; ekleme/taşımada dizin kayar | `update/moveChart`: aynı sayfada başka grafik komutu varsa düşer |
| `cmd` `removeChart` | Değişmez | Dizin kayar; silindiyse düşer | Aynı sayfada başka grafik komutu: düşer (sıra kaymış olabilir) |

İki yön tutarlıdır: bekleyen adım araya girenin içinden okunurken, araya giren de (sonraki bekleyen adımlar için) bekleyenin içinden okunur — bekleyen silme araya giren eklemeyi yuttuysa ekleme sonrakiler için yoktur; bekleyen ekleme araya giren silmenin içindeyse silme, eklenen satırların iki yanında iki parça olarak okunur; araya giren sıralama, bekleyen satır ekleme/silmenin ardından kalan satırların gittiği yerler olarak okunur. Sütunlar aynı kuraldır. Bu kurallar, yeniden temellendirme kodundan bağımsız bir başvuru modeliyle sınanır (`test/collab-semantic.test.ts`: satır ve sütunlar işaretlenir, beklenen sonuç yalnız aktarıcının sırasından hesaplanır).

Bu tablo **yakınsamayı etkilemez** (o, aktarıcının sırasından gelir); yalnız bir değişikliğin, araya başkası girdiğinde nereye düştüğünü belirler.

## Geri al / yinele

Yalnız katılımcının **kendi** adımları geri alınır. Geri alma, o adımı geri çeviren **yeni bir adımdır** (herkese gider); sonradan gelen adımlara göre yeniden temellendirilir. Başkası aynı hücreyi/notu/sayfa adını sonradan değiştirdiyse o kısım geri alınmaz; hiçbir şey kalmadıysa geri alma reddedilir ve söylenir.

| Adım | Geri alınması |
|---|---|
| Hücre, biçim, satır/sütun boyutu-gizleme, birleştirme, süzgeç, not | Önceki durum geri yazılır |
| Sıralama | Ters permütasyon + taşınan formüllerin özgün metinleri (tek kullanıcılı geçmişle aynı sonuç) |
| Satır/sütun ekleme | O satır/sütunlar silinir (başkası oraya yazdıysa o da gider) |
| Satır/sütun silme | Satır/sütunlar yeniden eklenir; hücreleri, biçimleri, boyutları, birleşimleri ve silmenin bozduğu formüller geri yazılır. Onlarla giden **notlar, bağlantılar, kurallar, süzgeç, adlar geri gelmez** (kullanıcıya söylenir). 200.000 hücreden fazlasını geri getirmesi gereken silme geri alınamaz |
| Sayfa ekleme / adlandırma / taşıma / gizleme / sekme rengi | Tersi; adlandırmada sayfayı anan formüller yazıldıkları biçimiyle geri gelir. Yeni ad (ya da eklenen sayfanın adı) kitapta **zaten bir formülde anılıyorsa** adım geri alınamaz (geri alma o formülü de değiştirirdi; söylenir) |
| **Sayfa silme** | **Geri alınamaz** (söylenir) |
| **Kural ve grafik komutları** | **Geri alınamaz** (söylenir) |

Değişmez: araya başkasının adımı girmeden yapılan "işlem → geri al", tek kullanıcılı geçmişin bıraktığı kitabı bırakır; bırakamıyorsa kullanıcıya söylenir (`test/collab-undo.test.ts`).

## Kayıt (`saveTo`) ile birlikte

- Kaydeden istemci, bekleyen adımı kalmayınca (en çok 3 sn bekler) dosyayı yazar; yazarken gelen iletiler bekletilir (dosya tek bir anın kitabıdır). Kayıt başarılı olunca `saved { v: ["e:<yeni ETag>", "h:…"], rev }` gönderir.
- Öbürleri bu iddiayı **sunucuya sorar** (`saveTo` adresine `HEAD`; desteklenmiyorsa `If-None-Match` ile `GET` — gövde indirilmez): sunucunun güçlü ETag'i söylenenle aynıysa ETag benimsenir (sonraki kayıt 412 almaz) ve kitap tam o adımdaysa, bekleyeni yoksa "kaydedilmiş" olur. Doğrulanamazsa hiçbir şey kaydedilmiş sayılmaz; söylenen ETag yalnız bir ipucudur: kendi kaydı 412 alırsa **koşullu olarak** (o ETag ile `If-Match`) bir kez yeniden denenir. Kaydeden, sonradan katılana son `saved`'ini yeniden söyler. İşlev olarak verilen `saveTo`'da ve köprüde doğrulama uygulamanındır.
- Dosya yazılırken yerelde bir şey değiştirildiyse o kayıt hiçbir adımın dosyası değildir: `saved` `rev` olmadan gider.
- Kendiliğinden kaydetme açıkken başkasının değişikliğini odada **tek** katılımcı kendiliğinden kaydeder (yazma yetkililerden kimliği en küçük olan); öbürleri `saved`'i doğrulayıp temizlenir.
- İki istemci aynı anda kaydederse biri 412 alır; o arada öbürünün `saved` iletisi geldiyse kayıt yeni ETag ile kendiliğinden yeniden gönderilir; gelmediyse olağan çakışma sorusu açılır ve `saved` gelince kapanır.
- Başkasının değişikliği dosyayı "kaydedilmemiş" yapar ama **kendiliğinden kaydetmeyi başlatmaz**: değişikliği yapan kaydeder.
- Çevrimdışıyken yapılan kayıt bildirilemez; bağlanınca adımlar gönderilir, öbürlerinin kaydı gerekirse çakışma yolundan geçer.

## Oda başka bir sürümdeyse: `detached`

Aktarıcı `reset` derse (dosya başka yerde değişmiş, günlük budanmış) ya da oda istemcinin gerisindeyse (aktarıcı yeniden başlamış, boş oda silinmiş):

- Kitap **temizse** (kaydedilmemiş hiçbir şey yok) dosya sessizce yeniden yüklenir ve odaya katılınır.
- Kitapta **hiçbir dosyada olmayan değişiklik varsa** (bekleyen ya da yerleşmiş ama kaydedilmemiş adım) kitap **asla** sessizce atılmaz: oturum `detached` olur, kitap yerinde kalır, değişiklikler kimseye gitmez, kullanıcıya açıkça söylenir. Kullanıcı kaydedince (çakışma varsa olağan 412 yolu) oturum, kaydedilen dosyanın sürümüyle kendiliğinden yeniden katılır; ya da uygulama `mounted.collab.reload()` ile sunucudaki dosyaya dönmeyi seçer (değişiklikler kaybolur).

## Uçucu fonksiyonlar

Formüller her istemcide yerel hesaplanır; hesaplanan değerler işlem olarak **yayınlanmaz**. `NOW`, `TODAY`, `RAND` gibi fonksiyonlar istemciler arasında farklı görünebilir; dosyaya kimin kaydı yazılırsa onun değerleri girer.

## Kendi taşıyıcınız

`collab: { transport }` ile kendi kanalınızı verebilirsiniz (Pusher, Reverb, bir ileti kuyruğu). Taşıyıcı dört yöntemdir:

```js
const transport = {
  send(text) { /* aktarıcıya gönder */ },
  onMessage(cb) { /* gelen her metni cb(text) ile ver */ },
  onState(cb) { /* isteğe bağlı: cb('open') / cb('closed') */ },
  close() {},
};
```

Bu durumda "dört kural"ı uygulayan taraf sizin arka ucunuzdur: yayın kanalları sıra garantisi vermediği için `op` kabulünü (tek sayaç, `base === rev`) sunucuda yapın ve iletileri bağlantı başına **sırayla** teslim edin. İstemci sırası bozuk gelen (`rev` atlayan) adımı görürse günlüğü `since` ile yeniden ister.
