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

1. **Oda** bir `rev` (0'dan başlar), bir `floor` (günlükteki en eski adımdan önceki sürüm), bir günlük ve bir **dönem adı** (`ep`) tutar. Dönem adı oda **kurulurken** aktarıcının ürettiği rastgele bir dizgedir (`[A-Za-z0-9_-]{1,64}`, gizli değildir) ve her `hello`'da verilir. Oda unutulup yeniden kurulursa (aktarıcı yeniden başladı, boş oda silindi, sayfa içi aktarıcının sayfası yenilendi) ya da boş ve `rev 0` iken **başka bir dosyayla** yeniden kurulursa dönem adı değişir: yeni oda adımlarını baştan sayar ve eskisinin "3. adımı" onda hiçbir şey demek değildir.
2. **`op`:** `base === rev` ise kabul: `rev++`, günlüğe yaz, odadaki **herkese (gönderen dahil)** yayınla. Değilse yalnız gönderene `nack`.
3. **`join`:** katılana `hello` ve durduğu adımdan sonraki günlük gönderilir. Durduğu adım iki yoldan bilinir: (i) `since` + **bu odanın** `ep`'i (odaya geri dönen istemci, ya da dosyasının içinde o odanın damgası olan istemci — aşağıda "Dosyadaki damga"); (ii) odanın **kurulduğu** dosyayla gelen, 0'dan. Başka her durumda `reset` (istemci dosyayı yeniden yükleyip yeniden katılır). `ep`'siz ya da başka `ep`'li `since` **yok sayılır**.
4. **Hafif iletiler** (`presence`, `saved`) günlüğe girmez, sıra numarası almaz, öbürlerine aktarılır.

Ek kurallar (aktarıcı):

5. **Sürüm (`pv`):** `join` protokol sürümünü taşır (şimdi **`2`**; `1`: dönem adı ve damga yoktu, `since` tek başına kabul ediliyordu — sessiz ıraksamaya yol açtığı için `1` konuşan istemci `2` odasına alınmaz). Odanın sürümü ilk katılanla sabitlenir; başka sürümle gelen `error version` alır. (Yeni bir sürümün tanıdığı işlem türü eskilerde atlanırdı; aynı odada iki sürüm ıraksama demektir.)
6. **Gizli anahtar (`key`):** `join` bir `key` taşır (1–64 karakter, `[A-Za-z0-9_-]`). Anahtar katılımcı ile aktarıcı arasında **sırdır**: hiçbir iletide, günlükte ya da oda listesinde yer almaz. Aynı odaya aynı anahtarla gelen bağlantı o katılımcının kendisidir (kopup dönen): eski bağlantı `error replaced` alır ve kapanır, yenisi yerine geçer. `op.id` herhangi bir dizgedir (1–100 karakter) ve anahtarla ilgisi yoktur; istemci adımlarını anahtardan türetilmeyen, herkesin gördüğü ayrı bir önekle adlandırır ve kendi adımını yalnız adından değil **içeriğinden** de tanır (aynı adla gelen başka içerik başkasının adımıdır). İçerik karşılaştırması anahtar sırasından bağımsızdır; yine de aktarıcı `ops`'u **aldığı metinle** aktarmalıdır (örnek sunucu öyle yapar). Gönderilen adım 30 sn içinde ne onay ne ret alırsa istemci bunu söyler ve yeniden bağlanır.
7. **Bağlantı başına tek oda:** katılmış bağlantıdan ikinci `join` → `bad-message`. `reset` almış bağlantı yalnız aynı odaya yeniden katılabilir. Sırası bozuk adım gören istemci yeni bağlantı kurup `since` ile katılır (dosya yeniden yüklenmez).
8. **Sürüm eşlemesi değişmez:** bir sürüm dizgesinin hangi adımı tuttuğu bir kez yazılır (ilk yazan kazanır; odanın kuruluş sürümleri 0'a kilitlidir). `saved` yalnız yazma yetkili üyeden kabul edilir ve `rev` yalnız gönderenin gerçekten gördüğü aralıkta (`katıldığı adım ≤ rev ≤ odanın adımı`). Tablo sınırlıdır (16 sürüm); **kuruluş sürümleri ve son kaydedilen sürüm asla atılmaz**, atılan bir sürüm **bir daha yazılamaz** (atılanlar ayrıca tutulur: en çok 4.096; dolunca oda yeni sürüm yazmaz). Var olan ya da atılmış bir sürüme başka adım söyleyen `saved`, `rev`'i olmadan aktarılır. Yeni sürüm yazan `saved` bağlantı başına hız sınırlıdır (örnek sunucu: dakikada 6).
9. **`sync`:** **yazma yetkili** biri sürümü bilinmediği için katılamazsa (`reset`), aktarıcı odadaki **yazma yetkili** üyelere `{ "t":"sync" }` der (oda başına dakikada en çok 2; salt okunur katılımcı ne tetikler ne alır). Yazma yetkili üyelerden **biri** (aktarıcının verdiği kimliği en küçük olan) dosyayı, **kendiliğinden kaydetme kurallarıyla** kaydeder (yazılmakta olan alan uygulanmaz, bekleyen adımı varsa ya da salt okunursa kaydetmez, çakışmada üzerine yazmaz) ve `saved { v, rev }` duyurur; yeni gelen dosyayı yeniden yükleyip katılır. Kimse kaydedemiyorsa yeni gelen artan aralıklarla altı kez dener ("şu an katılınamıyor: odadakilerden biri kaydetmeli" bildirimi), sonra oturumu kapanır. Kaydedemeyen lider yerine sıradaki adayın geçmesi YOKTUR.
10. **Sınırlar (örnek sunucu):** `op.ops` 1–10.001 öğe, en çok 256 düzey iç içe; bağlantı başına işlem hızı (jeton kovası; aşan `nack reason:"rate"`, istemci bir saniye sonra yeniden dener), varlık hızı ve **yeni hücre biçimi hızı** (adımların `st` tablolarındaki biçim sayısı; varsayılan dakikada 2.000, aşan `nack reason:"rate"`); toplam bağlantı, IP başına bağlantı, bütün odaların günlükleri için toplam bellek. Oda yolu yüzde-çözülmüş haliyle anahtardır (`/p/%61` ile `/p/a` aynı odadır); çözülünce `/`, `\`, `#` ya da denetim karakteri içeren yol reddedilir (400).

Günlük sınırlıdır (örnek sunucuda oda başına 20.000 adım / 32 MiB); aşınca en eski adımlar atılır ve `floor` ilerler. **Kaydedilmeden** biriken günlük budanırsa, eski dosyayla sonradan gelen katılamaz (`reset` alır, dosya hâlâ eskidir, oturum kapanır). Sınırı, iki kayıt arasındaki düzenleme hacmine göre seçin.

## İletiler

Her ileti tek bir JSON nesnesidir (WebSocket'te bir metin çerçevesi). En büyük ileti 1 MiB'dır. Tanınmayan ya da alanı yanlış ileti: aktarıcı `error bad-message` gönderip bağlantıyı kapatır; istemci oturumu kapatır.

### İstemci → aktarıcı

| İleti | Alanlar |
|---|---|
| `join` | `{ "t":"join", "pv":2, "key":dizge, "v":[dizge…], "since"?:tamsayı, "ep"?:dizge, "user"?:{ "name"?:dizge, "color"?:"#RRGGBB" }, "token"?:dizge, "ro"?:true, "room"?:dizge }` |
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
| `hello` | `{ "t":"hello", "pv":2, "ep":dizge, "rev":tamsayı, "from":tamsayı, "you":dizge, "ro":mantıksal, "v":[dizge…], "peers":[{ "id", "user":{name?,color?}, "ro", "presence"? }…] }` |
| `op` | `{ "t":"op", "rev":tamsayı, "id":dizge, "from":dizge, "ops":[…] }` |
| `nack` | `{ "t":"nack", "id":dizge, "rev":tamsayı, "reason"?:"read-only"\|"rate" }` |
| `sync` | `{ "t":"sync" }` — biri katılamıyor: kaydedebilen kaydetsin |
| `reset` | `{ "t":"reset", "rev":tamsayı, "v":[dizge…] }` |
| `presence` | `{ "t":"presence", "from":dizge, "data":… }` |
| `saved` | `{ "t":"saved", "from":dizge, "v":[dizge…], "rev"?:tamsayı }` |
| `peer` / `left` | `{ "t":"peer", "id", "user", "ro" }` / `{ "t":"left", "id" }` |
| `pong` | `{ "t":"pong" }` |
| `error` | `{ "t":"error", "code":dizge }` — ardından bağlantı kapanır (`bad-message`, `unauthorized`, `room-full`, `server-full`, `join-timeout`, `version`, `replaced` — aynı anahtarla yeni bağlantı geldi; istemci bunu hata saymaz) |

`hello.from`: günlüğün oynatılmaya başlandığı sürüm; `hello`'dan hemen sonra `rev > from` olan bütün adımlar sırayla gelir.

### `join` nasıl yanıtlanır (sırayla)

a. `since` verilmiş, `ep` **bu odanın dönem adı** ve `floor ≤ since ≤ rev` → `hello` (`from: since`) + sonraki adımlar. `hello.v`, katılanın `v`'sinden odanın **duyduğu** (kuruluş ya da `saved` tablosundaki) dizgeleri taşır.
b. Oda **taze** (`rev === 0` ve başka katılmış bağlantı yok) → oda (yeniden) kurulur: tablo temizlenir, `v`'nin her dizgesi **kuruluş sürümü** olarak 0'a yazılır → `hello` (`from: 0`). `v`'den biri odanın önceki kuruluş sürümüyse dönem adı **korunur** (hiçbir şey yazılmamış aynı oda: içindekiler geri dönebilir); değilse yeni dönem adı üretilir.
c. `v`'den biri odanın **kuruluş sürümü** ve `floor === 0` → `hello` (`from: 0`) + bütün adımlar.
d. Hiçbiri → `reset`. Bağlantı açık kalır ama odada sayılmaz; istemci dosyasını yeniden yükleyip aynı bağlantıdan yeniden `join` gönderir.

`saved { v, rev }` gelince (`floor ≤ rev ≤ oda rev'i`) `v`'nin her dizgesi tabloya o `rev` ile yazılır (en çok 16 dizge; en eskisi atılır). **Tablo kimseyi içeri almaz** (sürüm 1'de alıyordu): bir dizgenin hangi adımı tuttuğu yalnız birinin SÖZÜDÜR — aynı kitabı tutan herkes aynı baytları yazacağı için yazma yetkili biri dürüst kaydın özetini önceden hesaplayıp yanlış adıma yazdırabilir. Tablo şunlara yarar: aktarılan `saved` iletisinin `rev`'i, `reset.v`, ve (a) ile katılana "oda bu dosyayı duydu" demek (`hello.v`).

### Dosyadaki damga

Birlikte düzenlenen kitap kaydedilirken motor dosyanın **içine** odanın dönem adını ve dosyanın tuttuğu adımı yazar: `docProps/custom.xml` içinde tek bir özel belge özelliği, `BymOfficeCollab` = `"<ep>:<rev>"` (parça yoksa ilişkisi ve içerik türüyle birlikte oluşturulur; dosyanın öbür özel özelliklerine dokunulmaz; `collab` olmadan açılan kitapta **hiç yazılmaz**). Dosya hiçbir adımın dosyası değilse (bekleyen adım vardı, ya da yazılırken bir şey değiştirildi — motor iki kez yeniden dener) ya da oturum odadan ayrıysa (`detached`, `closed`) damga **silinir**.

Dosyayı açan istemci damgayı okur ve `join`'e `since: <rev>`, `ep: <ep>` koyar. `hello` geldiğinde: dönem adı damgadakiyle aynıysa `from` damgadaki adım olmalı **ve** `hello.v` boş olmamalıdır (oda bu baytları duymuş olmalı: damgayı taşıdığı halde oda dışında — Excel'de, `collab`'sız — değiştirilmiş dosya başka bir dosyadır); tutmuyorsa istemci damgayı bırakıp damgasız katılır (kuruluş dosyası değilse `reset` → `sync` → odanın kaydedicisi kaydeder → yeniden yükleme). Dönem adı başkaysa yalnız `from: 0` kabul edilir. Odaya **geri dönen** istemci (`since` = onaylanmış adımı, `ep` = içinde olduğu oda) `hello.ep` başka çıkarsa kendi adımına güvenmez: kaydedilmemiş işi yoksa dosyayı yeniden yükler, varsa `detached` olur (iş atılmaz).

Damga, **dosyanın durduğu yere yazabilen** biri tarafından sahtelenebilir; o kişi dosyanın kendisini de değiştirebilir. Aktarıcıdaki bir katılımcının sözüyle sahtelenemez. Damgayı yazıp okuyan modül: `src/formats/xlsx/collab-stamp.ts` (`readCollabStamp`, `writeCollabStamp`, `writeSharedFile`). Kendi kaydeden uygulama (`collab.saved(bytes, etag)`, köprüde `save`) `mounted.save()`'in verdiği baytları saklarsa damga içindedir.

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
| `merges` | Aynı sayfada araya giren satır/sütun ekleme-silmede **düşer** | Dizin kayar; silindiyse düşer | Var olan birleşimi kesen ekleme düşer; araya girenin yazdığı hücre/not birleşen aralıktaysa düşer |
| `filter` | Aralık ve ölçütler `shiftAutoFilter` ile kayar; başlık satırı gittiyse düşer; satır listeleri kayar | Dizin kayar; silindiyse düşer | Aynı sayfada iki süzgeç değişikliği: sonraki **düşer** (satır listeleri eski duruma göre hesaplanmıştır) |
| `sort` | Aynı sayfada araya giren **her** satır/sütun ekleme-silmede **düşer** (kanıtlanmış bir kural yok) | Dizin kayar; silindiyse düşer | Kesişen iki sıralama: sonraki düşer |
| (her biri) × araya giren `sort` | `cells`, `notes`: sıralanan sütunlardaki hücre satırıyla gider. Aynı sayfadaki bekleyen **satır/sütun ekleme-silme düşer ve ondan sonraki bütün bekleyen adımlar da düşer** (söylenir). `rows`, `cols`, `merges`, `filter`: yerinde kalır | — | — |
| `ins` | Yer kayar; aynı yere iki ekleme: ikisi de eklenir (sonra gelenin satırları üstte); başkasının sildiği aralığın **içine** yapılan ekleme silinen yerin başına iner (eklenen satırlar kalır) | Dizin kayar; silindiyse düşer | — |
| `del` | Aralık kayar; başkasının sildiğiyle örtüşen kısım düşülür, hepsi silinmişse adım **sessizce** biter (istenen zaten olmuş). **Silme yalnız yazarının GÖRDÜĞÜ satırları siler:** aralığın içine başkasının eklediği satırlar kalır, silme onların iki yanında iki parça olur (iki sırada da aynı sonuç) | Dizin kayar; silindiyse sessizce biter | — |
| `notes` | Hücreler kayar; gidenler çıkar | Dizin kayar; silindiyse düşer | Aynı hücre: sonra gelen kazanır |
| `sheet+` | — | Yer kayar | Aynı ad: sonraki düşer (ad alınmış) |
| `sheet-` | — | Dizin kayar; zaten silinmişse sessizce biter | — |
| `name`, `state`, `tab` | — | Dizin kayar; silindiyse düşer | Sonra gelen kazanır (ad alınmışsa düşer) |
| `move` | — | **Düşer** (aşağıdaki kurala bakın) | **Düşer** |
| `cmd` `setValidation`, `addConditionalRule`, `clearConditionalRules` | Aralıklar kayar, formüller yeniden yazılır; aralığın hepsi gittiyse düşer | Dizin kayar; silindiyse düşer | Sonra gelen uygulanır |
| `cmd` `update/remove/moveConditionalRule` (kuralı **sırasıyla** adlandırır) | Aynı sayfada satır/sütun değiştiyse **düşer** (bir kural gitmiş, sıralar kaymış olabilir) | Dizin kayar; silindiyse düşer | Aynı sayfada başka bir koşullu kural komutu: **düşer** |
| `cmd` `addChart`, `updateChart`, `moveChart` | Kitabın **herhangi bir** sayfasında satır/sütun değiştiyse **düşer** (başvurular ve çapalar yeniden yazılmıyor) | Sayfa silme/adlandırmada düşer; ekleme/taşımada dizin kayar | `update/moveChart`: aynı sayfada başka grafik komutu varsa düşer |
| `cmd` `removeChart` | Değişmez | Dizin kayar; silindiyse düşer | Aynı sayfada başka grafik komutu: düşer (sıra kaymış olabilir) |

**Kanıtlanmamış birleşim düşer.** Bir (bekleyen × araya giren) birleşiminin doğru yere taşındığı bağımsız bir başvuru modeliyle (`test/collab-semantic*.test.ts`, `test/collab-sheets.test.ts`, `test/collab-undo2.test.ts`) gösterilmemişse bekleyen adım düşer ve söylenir. Şimdi böyle olanlar: (1) **sayfa taşıma** (`move`) bekleyen ya da araya giren taraftaysa, yanındaki sayfa dizinli her bekleyen adım ve **ondan sonraki bütün bekleyen adımlar** düşer; (2) aynı sayfada **sıralama × satır/sütun ekleme-silme** (hangisi bekliyor olursa olsun) ve sonraki bekleyenler; (3) aynı sayfada **birleştirme × satır/sütun ekleme-silme**; (4) bekleyen bir **geri al/yinele** adımı araya girenle herhangi bir biçimde çakışırsa ya da ancak kısmen uygulanabilecekse bütünüyle düşer. Bir adımın uygulanması yerel stil tablosunu taşıracaksa (65.430 biçim) adım uygulanmadan **önce** oturum `detached` olur (kitap yerinde kalır, söylenir); başkasının adımı kitaba uygulanırken beklenmeyen bir hata çıkarsa da öyle — `session` dışına yakalanmamış hata çıkmaz.

İki yön tutarlıdır: bekleyen adım araya girenin içinden okunurken, araya giren de (sonraki bekleyen adımlar için) bekleyenin içinden okunur — bekleyen silme araya giren eklemeyi yuttuysa ekleme sonrakiler için yoktur; bekleyen ekleme araya giren silmenin içindeyse silme, eklenen satırların iki yanında iki parça olarak okunur; araya giren sıralama, bekleyen satır ekleme/silmenin ardından kalan satırların gittiği yerler olarak okunur. Sütunlar aynı kuraldır. Bu kurallar, yeniden temellendirme kodundan bağımsız bir başvuru modeliyle sınanır (`test/collab-semantic.test.ts`: satır ve sütunlar işaretlenir, beklenen sonuç yalnız aktarıcının sırasından hesaplanır).

Bu tablo **yakınsamayı etkilemez** (o, aktarıcının sırasından gelir); yalnız bir değişikliğin, araya başkası girdiğinde nereye düştüğünü belirler.

## Geri al / yinele

Yalnız katılımcının **kendi** adımları geri alınır. Geri alma, o adımı geri çeviren **yeni bir adımdır** (herkese gider). Geri alma yığını **onaylanmış** sıraya göre tutulur ve **LIFO**'dur: bir adım aktarıcı onaylayınca yığına girer ve yığındaki kayıt katılımcının KENDİ sonraki adımlarından **geçirilmez** (o adımlar yığında üstündedir ve onlar geri alınmadan alttakine ulaşılmaz; kayıt yazıldığı kitap için geçerlidir). Yalnız BAŞKASININ adımları yığından geçirilir — her kayda, üstündeki kayıtların geri alınmış halinden okunarak. Geri alınamayan bir adım (sayfa silme gibi) satır/sütun/sayfa kaydırdıysa altındaki kayıtlar ondan bir kez, kalıcı geçirilir; geri alınamaz hale gelen (`undo-gone`) bir adım satır/sütun/sayfa kaydırdıysa altındaki kayıtlar da düşer (söylenir). Değişmez (`test/collab-alone.test.ts`): odada başkası yokken — ya da başkası yalnız başka bir sayfaya yazarken — her işlem/geri al/yinele dizisi `collab`'sız geçmişle aynı kitabı bırakır ve hiçbir şey söylenmez; tek istisna geri getirilemeyen sayfa (eklemesi geri alınan sayfa yinelenemez: `undo-none`). katılımcının bekleyen adımı varken istenen geri alma, bekleyenler yerleşene dek **bekler** (3 sn'de yerleşmezse "şimdi geri alınamıyor" denir, `undo-wait`) ve sonra uygulanır. Geri alma **bütündür ya da hiç yoktur**: başkası aynı hücreyi/notu/sayfa adını sonradan değiştirdiyse, ya da geri alınacak satır/sütun ekleme-silmesinin sayfasına başkası sonradan satır/sütun ekleyip sildiyse, ya da adımın kendisi yerleşirken başkasının satır/sütun adımının yanından geçtiyse o adım **geri alınamaz** ve söylenir (`undo-gone`). Bekleyen geri alma isteği, sırası gelmeden yeni bir değişiklik yapılırsa bırakılır ve söylenir (`undo-wait`). Başkasının sildiği bir aralığın içine ya da hemen arkasına yapılan bekleyen ekleme, o silme geri alınıp satırlar geri gelince aralıktaki eski yerine (aynı uzaklığa) konur.

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
- Öbürleri bu iddiayı **sunucuya sorar** (`saveTo` adresine `HEAD`; desteklenmiyorsa `If-None-Match` ile `GET` — gövde indirilmez): sunucunun güçlü ETag'i söylenenle aynıysa ETag benimsenir (sonraki kayıt 412 almaz) ve kitap "kaydedilmiş" sayılır — ama yalnız şu koşulların **hepsi** tutuyorsa: sunucunun şimdiki ETag'i söylenenle aynı, söylenen `rev` bu istemcinin onaylanmış `rev`'i, bekleyen adımı yok, söylenen ETag bu istemcinin zaten bildiği ETag değil (bildiği ETag'den sonra adım uygulandıysa o dosya onları tutmaz) ve **bu istemcinin kendi yaptığı, kendi kaydetmediği değişiklik yok** (kendi işinin dosyada olduğuna başkasının sözüyle inanılmaz; yanlış "kirli" bir kayıt daha demektir, yanlış "temiz" veri kaybı). Doğrulanamazsa hiçbir şey kaydedilmiş sayılmaz; söylenen ETag yalnız bir ipucudur: kendi kaydı 412 alırsa **koşullu olarak** (o ETag ile `If-Match`) bir kez yeniden denenir. Kaydeden, sonradan katılana son `saved`'ini yeniden söyler. İşlev olarak verilen `saveTo`'da ve köprüde doğrulama uygulamanındır.
- Dosya yazılırken yerelde bir şey değiştirildiyse o kayıt hiçbir adımın dosyası değildir: `saved` `rev` olmadan gider.
- **Kendiliğinden kaydetmeyi odada TEK katılımcı yapar** (odanın kaydedicisi: yazma yetkililerden aktarıcının verdiği kimliği en küçük olan), herkesin değişikliği için. Öbürlerinin kendiliğinden kaydı durur (elle Kaydet herkeste çalışır); kaydedici ayrılırsa sıradaki devralır; aktarıcıya ulaşamayan katılımcı (kimin kaydedici olduğunu bilemez) kendiliğinden kaydetmez; odadan ayrı (`detached`, `closed`) kitap her kitap gibi kaydedilir. Katılımcının **kendi** değişiklikleri yalnız **kaydedicinin** doğrulanmış `saved`'iyle "kaydedilmiş" sayılır; başkasının sözüyle sayılmaz.
- İki istemci aynı anda kaydederse biri 412 alır; öbürünün `saved` iletisi o arada geldiyse ya da **2,5 sn içinde gelip doğrulanırsa** (ETag sunucuda öyle) çakışma sayılmaz: kayıt yeni ETag'i koşul alarak bir kez yeniden gönderilir (doğrulanmayan söz bunu tetiklemez; hiçbir zaman koşulsuz yazılmaz); gelmediyse olağan çakışma sorusu açılır. Açık bir çakışma sorusunu `saved` **kapatmaz** (doğrulanmış olsa da): üzerine yazma kararı kullanıcınındır. Aktarıcı ya da oda değişince (yeniden yükleme, `detached`) son `saved` unutulur.
- Başkasının değişikliği dosyayı "kaydedilmemiş" yapar ama **kendiliğinden kaydetmeyi başlatmaz**: değişikliği yapan kaydeder.
- Çevrimdışıyken yapılan kayıt bildirilemez; bağlanınca adımlar gönderilir, öbürlerinin kaydı gerekirse çakışma yolundan geçer.

## Oda başka bir sürümdeyse: `detached`

Aktarıcı `reset` derse (dosya başka yerde değişmiş, günlük budanmış) ya da oda istemcinin gerisindeyse (aktarıcı yeniden başlamış, boş oda silinmiş):

- Kitap **temizse** (kaydedilmemiş hiçbir şey yok) dosya sessizce yeniden yüklenir ve odaya katılınır.
- Kitapta **hiçbir dosyada olmayan değişiklik varsa** (bekleyen ya da yerleşmiş ama kaydedilmemiş adım) kitap **asla** sessizce atılmaz: oturum `detached` olur, kitap yerinde kalır, değişiklikler kimseye gitmez, kullanıcıya açıkça söylenir. Kullanıcı kaydedince (çakışma varsa olağan 412 yolu) oturum, kaydedilen dosyanın sürümüyle kendiliğinden yeniden katılır; ya da uygulama `mounted.collab.reload()` ile sunucudaki dosyaya dönmeyi seçer (değişiklikler kaybolur).
- **`detached` + kaydet bir çatallanmadır.** Akış: kayıt bildiği ETag ile `If-Match` gönderir; sunucudaki dosya başkaysa 412 gelir ve kullanıcıya sorulur. "Üzerine yaz" = *benimkini kaydet* (odadakilerin o ana dek kaydettiği dosya ezilir; onların sonraki kaydı 412 alır ve onlara sorulur). Vazgeçip `reload()` = *sunucudakini yükle* — uygulama önce `mounted.save()` ile kitabı indirmeyi önermelidir. Görünümde bu iki seçeneği yan yana sunan hazır bir kutu YOKTUR (çakışma sorusu + `reload()` API'si vardır). `saveTo`'suz (ETag'siz) kayıtta bu koruma yoktur: uygulama kendi denetler.
- **Yeniden yükleme sürerken** görünüm düzenlemeye kapalıdır ("güncelleniyor" bildirimi) ve oturum adım kabul etmez (`submit` hata fırlatır); yükleme bitince açılır.

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
