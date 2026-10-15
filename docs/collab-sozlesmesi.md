# Birlikte düzenleme sözleşmesi

Bu belge, bir çalışma kitabını (xlsx) ya da bir belgeyi (docx) birlikte düzenleyen tarayıcılar ile aralarındaki **aktarıcı** (relay) arasındaki iletileri tanımlar. Aktarıcıyı kendi arka ucunuzda yazacaksanız (Laravel Reverb, Pusher, bir Node süreci, başka bir şey) gereken her şey buradadır. Depodaki hazır örnek: `server/collab-server.mjs` (bağımlılıksız Node; **bellekte** çalışır, hiçbir şeyi diske yazmaz). Aynı kuralların sayfa içi sürümü: `BymOffice.collab.memoryServer()`.

Aktarıcı **her iki tür için aynıdır** (biçimden habersizdir): oda, sıra, günlük, dönem adı, damga, `saved`, `sync`, `presence` kuralları tablolar ve belgeler için birdir. Türe göre değişen yalnız `op.ops`'un ve `presence.data`'nın **içeriğidir** (aktarıcı ikisine de bakmaz): tablolar için aşağıdaki "`ops`: bir adımın içeriği", belgeler için en sondaki "Belgeler (docx)" bölümü. Bir odada tek tür dosya düzenlenir (oda bir dosyanın sürümüne bağlıdır).

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

5. **Sürüm (`pv`):** `join` protokol sürümünü taşır (şimdi **`6`**; `5`: belgede yorum komutları — `cmt` — yoktu; `4`: `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `EXPAND`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `TEXTSPLIT` hesaplanmıyordu (`#NAME?`), `TEXTBEFORE`/`TEXTAFTER` ayırıcı dizisini tek liste saymıyordu; `3`: formülün dinamik dizi işareti — hücre formülünde `d: 1` — yoktu; `2`: tablo `pane` adımı ve belgede sekme durakları — `para.f.tabs`, `put` paragraflarında `tabs` listesi — yoktu; `1`: dönem adı ve damga yoktu, `since` tek başına kabul ediliyordu). Odanın sürümü ilk katılanla sabitlenir; başka sürümle gelen `error version` alır. `hello` odanın sürümünü `pv` ile söyler; istemci kendi sürümünden farklı `pv` görürse de katılmaz. `pv` bir sayıdır (1–1000 tamsayı); istemci `hello.pv`'de yalnız rakamlardan oluşan dizgeyi de (`"3"`; ör. PHP arka uçları) kabul eder, başka biçim (`"3.0"`, `"x"`) iletiyi geçersiz kılar. Örnek aktarıcılar `join.pv`'de dizge kabul etmez. İstemci `version` hatasını genel retten (`refused`) ayırır: oturum `closed` olur, `reason` `version`'dır, kullanıcıya "Bu dosya başka bir sürümle düzenleniyor; sayfayı yenileyin." denir. (Yeni bir sürümün tanıdığı işlem türü eskilerde atlanırdı; aynı odada iki sürüm ıraksama demektir.) **Sürüm ne zaman artar:** tel adımına yeni tür ya da alan eklenince, bir adımın anlamı değişince **ve formüllerin hesap sonucu değişince** (yeni fonksiyon, bir fonksiyonun başka sonuç vermesi): her katılımcı odanın formüllerini kendi motoruyla hesaplar, aynı adımları alan eski sürüm fonksiyonu tanımaz (`#NAME?`) ya da eski yoldan hesaplar ve ıraksar. Tel parmak izi (`test/collab-protocol-version.test.ts`) yalnız tel biçimini görür; hesap değişikliğinde sürüm elle artırılır.

   **Sürüm kuralı:** tel adımına (tablo `ops` ya da belge komutu) yeni bir tür ya da alan eklendiğinde, ya da var olan bir adımın anlamı değiştiğinde `PROTOCOL_VERSION` (`src/collab/room.ts`) bir artırılır. Bunu `test/collab-protocol-version.test.ts` hatırlatır: tablo ve belge çözücülerinin kabul ettiği adım türlerinden ve her türde adıyla okudukları alanlardan bir "parmak izi" çıkarır (örnek adımlar, okunan her anahtarı not eden bir vekil (proxy) içinden çözülür); parmak izi değişip sürüm aynı kalırsa test düşer ve "pv'yi artır, parmak izini güncelle" der. Parmak izinin görmediği: komut bağımsız değişkenlerinin içi (`cmd.a`), alanların alabildiği değerler (sıralı listelerin üyeleri) ve aktarıcı iletileri (`join`, `hello`…) — bunlarda değişiklik yine sürüm ister, ama elle.
6. **Gizli anahtar (`key`):** `join` bir `key` taşır (1–64 karakter, `[A-Za-z0-9_-]`). Anahtar katılımcı ile aktarıcı arasında **sırdır**: hiçbir iletide, günlükte ya da oda listesinde yer almaz. Aynı odaya aynı anahtarla gelen bağlantı o katılımcının kendisidir (kopup dönen): eski bağlantı `error replaced` alır ve kapanır, yenisi yerine geçer. `op.id` herhangi bir dizgedir (1–100 karakter) ve anahtarla ilgisi yoktur; istemci adımlarını anahtardan türetilmeyen, herkesin gördüğü ayrı bir önekle adlandırır ve kendi adımını yalnız adından değil **içeriğinden** de tanır (aynı adla gelen başka içerik başkasının adımıdır). İçerik karşılaştırması anahtar sırasından bağımsızdır; yine de aktarıcı `ops`'u **aldığı metinle** aktarmalıdır (örnek sunucu öyle yapar). Gönderilen adım 30 sn içinde ne onay ne ret alırsa istemci bunu söyler ve yeniden bağlanır.
7. **Bağlantı başına tek oda:** katılmış bağlantıdan ikinci `join` → `bad-message`. `reset` almış bağlantı yalnız aynı odaya yeniden katılabilir. Sırası bozuk adım gören istemci yeni bağlantı kurup `since` ile katılır (dosya yeniden yüklenmez).
8. **Sürüm tablosu:** bir sürüm dizgesinin hangi adımı tuttuğu bir kez yazılır (ilk yazan kazanır; odanın kuruluş sürümleri 0'a kilitlidir). `saved` yalnız yazma yetkili üyeden kabul edilir ve `rev` yalnız gönderenin gerçekten gördüğü aralıkta (`katıldığı adım ≤ rev ≤ odanın adımı`). Tablo sınırlıdır (16 sürüm); kuruluş sürümleri ve son kaydedilen sürüm atılmaz. Atılan sürümlerin **son** 4.096'sı hatırlanır ve hatırlandıkça yeniden yazılamaz (kayan pencere; oda yazmayı hiçbir zaman bırakmaz — tablo kimseyi içeri almadığı için unutulan bir sürümün yeniden yazılması zararsızdır, yazmayı bırakan oda ise kaydettiği dosyayla dönen herkesi dışarıda bırakırdı). Tabloyla çelişen ya da hatırlanan bir sürümü anan `saved`, `rev`'i olmadan aktarılır. Yeni sürüm yazan `saved` hız sınırlıdır (örnek sunucu: bağlantı başına dakikada **40**, oda başına **120**; istemcinin kendiliğinden kaydı en sık 2 sn'de birdir = dakikada 30, sınır bunun üstünde olmalıdır: sınırı aşan kayıt tabloya yazılmaz ve o dosyayla dönen istemci odaya ancak yeni bir kayıttan sonra girer).
9. **`sync`:** **yazma yetkili** biri sürümü bilinmediği için katılamazsa (`reset`), aktarıcı odadaki **yazma yetkili** üyelere `{ "t":"sync" }` der (oda başına dakikada en çok 2; salt okunur katılımcı ne tetikler ne alır). Yazma yetkili üyelerden **biri** (aktarıcının verdiği kimliği en küçük olan) dosyayı, **kendiliğinden kaydetme kurallarıyla** kaydeder (yazılmakta olan alan uygulanmaz, bekleyen adımı varsa ya da salt okunursa kaydetmez, çakışmada üzerine yazmaz) ve `saved { v, rev }` duyurur; yeni gelen dosyayı yeniden yükleyip katılır. Kimse kaydedemiyorsa yeni gelen artan aralıklarla altı kez dener ("şu an katılınamıyor: odadakilerden biri kaydetmeli" bildirimi), sonra oturumu kapanır. Kaydedemeyen lider yerine sıradaki adayın geçmesi YOKTUR.
10. **Sınırlar (örnek sunucu):** `op.ops` 1–10.001 öğe, en çok 256 düzey iç içe; bağlantı başına işlem hızı (jeton kovası; aşan `nack reason:"rate"`, istemci bir saniye sonra yeniden dener), varlık hızı ve **yeni hücre biçimi hızı** (adımların `st` tablolarındaki biçim sayısı; varsayılan dakikada 2.000, aşan `nack reason:"rate"`); toplam bağlantı, IP başına bağlantı, bütün odaların günlükleri için toplam bellek. Oda yolu yüzde-çözülmüş haliyle anahtardır (`/p/%61` ile `/p/a` aynı odadır); çözülünce `/`, `\`, `#` ya da denetim karakteri içeren yol reddedilir (400).

Günlük sınırlıdır (örnek sunucuda oda başına 20.000 adım / 32 MiB); aşınca en eski adımlar atılır ve `floor` ilerler. **Kaydedilmeden** biriken günlük budanırsa, eski dosyayla sonradan gelen katılamaz (`reset` alır, dosya hâlâ eskidir, oturum kapanır). Sınırı, iki kayıt arasındaki düzenleme hacmine göre seçin.

## İletiler

Her ileti tek bir JSON nesnesidir (WebSocket'te bir metin çerçevesi). En büyük ileti 1 MiB'dır. Tanınmayan ya da alanı yanlış ileti: aktarıcı `error bad-message` gönderip bağlantıyı kapatır; istemci oturumu kapatır.

### İstemci → aktarıcı

| İleti | Alanlar |
|---|---|
| `join` | `{ "t":"join", "pv":3, "key":dizge, "v":[dizge…], "since"?:tamsayı, "ep"?:dizge, "user"?:{ "name"?:dizge, "color"?:"#RRGGBB" }, "token"?:dizge, "ro"?:true, "room"?:dizge }` |
| `op` | `{ "t":"op", "base":tamsayı, "id":dizge, "ops":[…] }` |
| `presence` | `{ "t":"presence", "data":{ "s":sayfa, "a":[satır,sütun], "r":[üst,sol,alt,sağ] } \| null }` |
| `saved` | `{ "t":"saved", "v":[dizge…], "rev"?:tamsayı }` |
| `ping` | `{ "t":"ping" }` |
| (hız) | Bir bağlantı dakikada en çok 10 kez `join` gönderebilir (örnek sunucu: `joinsPerMinute`); aşan bağlantı `error join-rate` ile kapatılır |
| `leave` | `{ "t":"leave" }` — odadan çıkar, bağlantı açık kalır; ardından yalnız **aynı odaya** yeniden `join` gönderilebilir (`reset` sonrası gibi). Yeni bağlantı kuramayan istemci (uygulama taşıyıcısı: `suspend`/`resume` yok) yeniden katılması gerektiğinde bunu kullanır; katılmış bağlantıdan `leave`'siz ikinci `join` `bad-message`'dır |

- `join` bağlantının **ilk** iletisidir. `v`: dosyanın neyle tanındığı — 0–8 dizge, her biri 1–200 karakter. Motor şunları gönderir: `e:<ETag>` (dosya bir adresten ETag ile geldiyse) ve `h:<uzunluk>:<16 onaltılık>` (baytların özeti; gizli değildir, kanıt değildir). `since`: istemcinin elindeki son `rev` (yeniden bağlanırken). `token`: yalnız bu iletide gider; motor onu hiçbir bildirime ya da günlüğe yazmaz. `ro`: yalnız izler.
- `op.id`: istemcinin kendi adımını tanıdığı ad (`<oturum anahtarı>:<sıra>`, en çok 100 karakter). `ops`: aktarıcı için **opak** bir JSON dizisi (1–10.001 öğe); biçimi aşağıda.
- `saved.rev`: kaydedilen dosyanın tuttuğu adım. Verilmezse aktarıcı yalnız iletiyi aktarır (öbürleri yeni ETag'i öğrenir), sürüm tablosuna yazmaz.

### Aktarıcı → istemci

| İleti | Alanlar |
|---|---|
| `hello` | `{ "t":"hello", "pv":3, "ep":dizge, "rev":tamsayı, "from":tamsayı, "you":dizge, "ro":mantıksal, "v":[dizge…], "peers":[{ "id", "user":{name?,color?}, "ro", "presence"? }…] }` |
| `op` | `{ "t":"op", "rev":tamsayı, "id":dizge, "from":dizge, "ops":[…] }` |
| `nack` | `{ "t":"nack", "id":dizge, "rev":tamsayı, "reason"?:"read-only"\|"rate" }` |
| `sync` | `{ "t":"sync" }` — biri katılamıyor: kaydedebilen kaydetsin |
| `reset` | `{ "t":"reset", "rev":tamsayı, "v":[dizge…] }` |
| `presence` | `{ "t":"presence", "from":dizge, "data":… }` |
| `saved` | `{ "t":"saved", "from":dizge, "v":[dizge…], "rev"?:tamsayı }` |
| `peer` / `left` | `{ "t":"peer", "id", "user", "ro" }` / `{ "t":"left", "id" }` |
| `pong` | `{ "t":"pong" }` |
| `error` | `{ "t":"error", "code":dizge }` — ardından bağlantı kapanır (`bad-message`, `unauthorized`, `room-full`, `server-full`, `join-timeout`, `version` — oda başka bir protokol sürümünde; istemci bunu genel retten ayrı söyler —, `replaced` — aynı anahtarla yeni bağlantı geldi; istemci bunu hata saymaz) |

`hello.user` (isteğe bağlı): katılımcının odada nasıl göründüğü (`{ name?, color? }`); aktarıcı adı kendisi verdiyse (örnek sunucuda `authorize` → `{ user }`) istemci kendini o adla gösterir. `hello.from`: günlüğün oynatılmaya başlandığı sürüm; `hello`'dan hemen sonra `rev > from` olan bütün adımlar sırayla gelir.

### `join` nasıl yanıtlanır (sırayla)

a. `since` verilmiş, `ep` **bu odanın dönem adı** ve `floor ≤ since ≤ rev` → `hello` (`from: since`) + sonraki adımlar. `hello.v`, katılanın `v`'sinden odanın **duyduğu** (kuruluş ya da `saved` tablosundaki) dizgeleri taşır.
b. Oda **taze** (`rev === 0` ve başka katılmış bağlantı yok) → oda (yeniden) kurulur: tablo temizlenir, `v`'nin her dizgesi **kuruluş sürümü** olarak 0'a yazılır → `hello` (`from: 0`). `v`'den biri odanın önceki kuruluş sürümüyse dönem adı **korunur** (hiçbir şey yazılmamış aynı oda: içindekiler geri dönebilir); değilse yeni dönem adı üretilir.
c. `v`'den biri odanın **kuruluş sürümü** ve `floor === 0` → `hello` (`from: 0`) + bütün adımlar.
d. Hiçbiri → `reset`. Bağlantı açık kalır ama odada sayılmaz; istemci dosyasını yeniden yükleyip aynı bağlantıdan yeniden `join` gönderir.

`saved { v, rev }` gelince (`floor ≤ rev ≤ oda rev'i`) `v`'nin her dizgesi tabloya o `rev` ile yazılır (en çok 16 dizge; en eskisi atılır). **Tablo kimseyi içeri almaz** (sürüm 1'de alıyordu): bir dizgenin hangi adımı tuttuğu yalnız birinin SÖZÜDÜR — aynı kitabı tutan herkes aynı baytları yazacağı için yazma yetkili biri dürüst kaydın özetini önceden hesaplayıp yanlış adıma yazdırabilir. Tablo şunlara yarar: aktarılan `saved` iletisinin `rev`'i, `reset.v`, ve (a) ile katılana "oda bu dosyayı duydu" demek (`hello.v`).

### Dosyadaki damga

Birlikte düzenlenen kitap kaydedilirken motor dosyanın **içine** odanın dönem adını ve dosyanın tuttuğu adımı yazar: `docProps/custom.xml` içinde tek bir özel belge özelliği, `BymOfficeCollab` = `"<ep>:<rev>"` (parça yoksa ilişkisi ve içerik türüyle birlikte oluşturulur; dosyanın öbür özel özelliklerine dokunulmaz; `collab` olmadan açılan kitapta **hiç yazılmaz**). Dosya hiçbir adımın dosyası değilse (bekleyen adım vardı, ya da yazılırken bir şey değiştirildi — motor iki kez yeniden dener) ya da oturum odadan ayrıysa (`detached`, `closed`) damga **silinir**.

Dosyayı açan istemci damgayı okur ve `join`'e `since: <rev>`, `ep: <ep>` koyar. `hello` geldiğinde: dönem adı damgadakiyle aynıysa `from` damgadaki adım olmalı **ve** `hello.v` boş olmamalıdır (oda bu baytları duymuş olmalı: damgayı taşıdığı halde oda dışında — Excel'de, `collab`'sız — değiştirilmiş dosya başka bir dosyadır); tutmuyorsa istemci damgayı bırakıp damgasız yeniden katılır (yeni bağlantıyla, ya da taşıyıcı bağlantıyı yeniden kuramıyorsa `leave` + `join` ile) (kuruluş dosyası değilse `reset` → `sync` → odanın kaydedicisi kaydeder → yeniden yükleme). Dönem adı başkaysa yalnız `from: 0` kabul edilir. Odaya **geri dönen** istemci (`since` = onaylanmış adımı, `ep` = içinde olduğu oda) `hello.ep` başka çıkarsa kendi adımına güvenmez: kaydedilmemiş işi yoksa dosyayı yeniden yükler, varsa `detached` olur (iş atılmaz).

Dosyanın özel özellikler parçası okunamıyorsa (bozuk XML, DTD, 1 MiB'tan büyük) parçaya **dokunulmaz**: dosya damgasız kaydedilir ("hiçbir adımın dosyası değil" yolu) ve bir kez söylenir; öyle bir parçadan damga da okunmaz. Aynı adlı iki damga özelliği damga **yok** sayılır (yazarken tekine indirilir, silerken hepsi çıkar). Yeni özelliğin `pid`'i 2'den başlayarak boş olan en küçük sayıdır.

Damga, **dosyanın durduğu yere yazabilen** biri tarafından sahtelenebilir; o kişi dosyanın kendisini de değiştirebilir. Aktarıcıdaki bir katılımcının sözüyle sahtelenemez. Damgayı yazıp okuyan modül: `src/formats/xlsx/collab-stamp.ts` (`readCollabStamp`, `writeCollabStamp`, `writeSharedFile`). Kendi kaydeden uygulama (`collab.saved(bytes, etag)`, köprüde `save`) `mounted.save()`'in verdiği baytları saklarsa damga içindedir.

### Oda anahtarı, yetki

Örnek sunucuda oda = isteğin **yol adı** (sorgu hariç) + varsa `join.room`. Yani `wss://example.com/collab/42` adresini açan herkes aynı odadadır. Yetki sizdedir: örnek sunucunun `authorize({ room, token, origin, request })` kancası `true` / `false` / `{ readOnly: true }` döndürür; `origins` seçeneği izinli `Origin` listesidir. Kendi aktarıcınızda `token`'ı (ya da çerezinizi) doğrulayın: **odaya giren herkes dosyanın içeriğini öğrenir ve değiştirebilir.**

## `ops`: bir adımın içeriği

`ops`, tel işlemlerinin dizisidir; hücre biçimi taşıyan adımda **ilk öğe** biçim tablosudur: `{ "t":"st", "f":[biçim…] }`. İşlemlerdeki biçim alanları (`s`) bu tablonun dizinleridir. Biçim: `{ "n": sayı biçimi (yerleşik biçimin numarası 0–163 ya da kod dizgesi), "f": font, "l": dolgu, "b": kenarlık, "x"?: adlandırılmış stil dizini, "a"?: hizalama, "p"?: koruma, "q"?: true }` (alanlar modeldeki `Font`, `Fill`, `Border`, `Alignment`, `Protection` ile aynıdır; geçişli dolgu `{ "type":"gradient", "xml":"…" }`).

Bütün konumlar 0 tabanlıdır ve **şimdiki** durumdadır (dosyadaki değil). `s` sayfa dizinidir. Aralık `[üst, sol, alt, sağ]`.

| `t` | Alanlar | Anlam |
|---|---|---|
| `cells` | `c: [[satır, sütun, yama]…]` | Yama `0`: hücre gider. Değilse `{ "v"?: içerik, "s"?: biçim }`: verilmeyen yan olduğu gibi kalır (biçim ile değer ayrı ayrı değişir). İçerik: `{ "k": "blank"\|"number"\|"string"\|"boolean"\|"error"\|"date", "v": değer, "f"?: { "x": formül metni ("=" yok), "t"?: "array"\|"dataTable", "r"?: aralık, "a"?: [[ad,değer]…], "d"?: 1 }, "r"?: zengin metin parçaları, "ph"?: true }`. Ortak (shared) formül, hücresi için açık yazılmış gider. `d: 1` dinamik dizi formülüdür (Excel 365: sonucu taşar); nereye taşacağı tele girmez, her katılımcı kendi motoruyla aynı sonucu hesaplar. |
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
- Öbürleri bu iddiayı **sunucuya sorar** (`saveTo` adresine `HEAD`; desteklenmiyorsa `If-None-Match` ile `GET` — gövde indirilmez): sunucunun güçlü ETag'i söylenenle aynıysa ETag benimsenir (sonraki kayıt 412 almaz) ve kitap "kaydedilmiş" sayılır — ama yalnız şu koşulların **hepsi** tutuyorsa: sunucunun şimdiki ETag'i söylenenle aynı, söylenen `rev` bu istemcinin onaylanmış `rev`'i, bekleyen adımı yok, söylenen ETag bu istemcinin zaten bildiği ETag değil (bildiği ETag'den sonra adım uygulandıysa o dosya onları tutmaz) ve **bu istemcinin kendi yaptığı, kendi kaydetmediği değişiklik yok** (kendi işinin dosyada olduğuna başkasının sözüyle inanılmaz; yanlış "kirli" bir kayıt daha demektir, yanlış "temiz" veri kaybı). Doğrulanamazsa hiçbir şey kaydedilmiş sayılmaz; sunucuya **ulaşılamadıysa** söylenen ETag yalnız bir ipucudur (sunucu dosyayı başka adla anıyorsa ipucu bile değildir; ipucu dosyanın bilinen adı ilerleyince silinir, eski bir ada asla geri dönülmez; doğrulanamayan söz bekleyen bir kaydın beklemesini erken bitirmez): kendi kaydı 412 alırsa **koşullu olarak** (o ETag ile `If-Match`) bir kez yeniden denenir. Kaydeden, sonradan katılana son `saved`'ini yeniden söyler. İşlev olarak verilen `saveTo`'da ve köprüde doğrulama uygulamanındır.
- Dosya yazılırken yerelde bir şey değiştirildiyse o kayıt hiçbir adımın dosyası değildir: `saved` `rev` olmadan gider.
- **Kendiliğinden kaydetmeyi odada olağan durumda TEK katılımcı yapar** (odanın kaydedicisi: yazma yetkililerden aktarıcının verdiği kimliği en küçük olan), herkesin değişikliği için. **Kaydedicilik yetki değildir ve hiçbir güven taşımaz** (kimlik yeniden bağlanarak değiştirilebilir): yalnız kimin önce kaydedeceğini belirler.
- **Nöbet:** kaydedici olmayan katılımcıda kaydedilmemiş iş (kendi işi ya da başkasınınki) bir süre boyunca hiçbir kayda girmediyse — kaydedicinin `autoSave`'i kapalı, `saveTo`'su yok, kaydı reddediliyor, sekmesi donmuş… — o katılımcı **kendisi** kaydeder (koşullu, `If-Match`; kendiliğinden kaydetmenin bütün kurallarıyla). Süre: `autoSave` gecikmesinin 5 katı, en az 15 sn, üstüne rastgele bir pay (en az sürenin %25'i; odadaki kişi sayısıyla büyür: kişi başına 3 sn — bir kayıt ve haberi, sıradaki nöbetçinin sırası gelmeden öbürlerini temizlesin). Pay yalnız **yazabilen** katılımcı sayısıyla büyür ve en çok 30 sn'dir. Aktarıcıya **uzun** bir kesintiden (nöbet süresi kadar ya da daha uzun) sonra yeniden ulaşılınca bekleyenler yalnız pay kadar (herkes kendi anında) daha bekler; kısa kopmalarda süre kaldığı yerden sürer (bağlantısı sık kopan katılımcının işi kaydedilmeden kalmaz). Kendiliğinden kayıt odadaki kayıtlarla art arda çakışırsa bu çakışma sayılmaz: "oda yoğun" denir, bekler ve yeniden dener (kalıcı durmaz); `collab.saveWatchMs` ile değişir. Bu arada `SaveState.autoWaits: 'room'` görünür. Aktarıcıya ulaşılamıyorken kendiliğinden kayıt **yapılmaz** (`autoWaits: 'offline'`): kopukken yazılan dosya hiçbir adımın dosyası değildir (damga silinir) ve dönüşte çatallanma olurdu; elle Kaydet çalışır, sayfadan ayrılırken sorulur.
- **"Kaydedilmiş" sayılma KANITLA olur, sözle değil.** Kendi kaydedilmemiş değişikliği olan katılımcı, kim söylerse söylesin başkasının `saved`'iyle temiz sayılmaz. Söz geldiğinde (ETag sunucuda doğrulanmış, söylenen `rev` kendi onaylanmış `rev`'i, bekleyeni yok) 0,4 sn sonra — arada yeni söz gelirse yalnız sonuncusu için — dosyayı `GET` ile çeker (32 MiB'tan büyükse çekmez) ve **kendi kitabının aynı adımda üreteceği dosyayla** (paketin bir kopyasından yazılır; dosyanın kendi paketi değişmez) karşılaştırır: önce yalnız dizin okunur — parça adları ve bildirilen açılmış boyutlar tutmuyorsa hiçbir şey açılmadan "başka dosya" denir — paketin parça adları kümesi ve her parçanın **açılmış** içeriği birebir aynı olmalıdır (sıkıştırma motordan motora değişir; sıra ve ZIP üstverisi sayılmaz). Yanıtın ETag'i söylenenle aynı ve içerik aynıysa temizdir ve o andan sonra katılımcı "kendi işi olmayan" gibi davranır (yeni bir şey yazana dek dosyayı yeniden çekmez); değilse kirli kalır ve nöbet kuralı işler. (Parola korumalı dosyada baytlar hiçbir zaman aynı çıkmaz: orada her katılımcı işini nöbetle kendi kaydeder.) Kendi değişikliği **olmayan** katılımcı için doğrulanmış söz + aynı adım yeter.
- İki istemci aynı anda kaydederse biri 412 alır; öbürünün `saved` iletisi o arada geldiyse ya da **2,5 sn içinde gelip doğrulanırsa** (ETag sunucuda öyle) çakışma sayılmaz: kayıt yeni ETag'i koşul alarak yeniden gönderilir (art arda en çok altı kez: odada birkaç kişi peş peşe kaydedebilir) (doğrulanmayan söz bunu tetiklemez; hiçbir zaman koşulsuz yazılmaz); gelmediyse olağan çakışma sorusu açılır. Açık bir çakışma sorusunu `saved` **kapatmaz** (doğrulanmış olsa da): üzerine yazma kararı kullanıcınındır. Aktarıcı ya da oda değişince (yeniden yükleme, `detached`) son `saved` unutulur.
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

## Belgeler (docx)

Aşama 4b. Bir belge odasında aktarıcı, oda kuralları, damga (`docProps/custom.xml` içinde `BymOfficeCollab = "<ep>:<rev>"`; aynı modül, aynı kayıpsızlık kuralları), `saved` / `sync` / kayıt kuralları tablolardakinin **aynısıdır**. Farklı olan adımın içeriği, yeniden temellendirme kuralları, geri alma ve varlık iletisidir. Kod: `src/collab/doc/` (`wire.ts` tel + doğrulama, `rebase.ts` dönüşüm, `session.ts` oturum, `trace.ts` + `history.ts` görünümün komutlarını adıma çeviren katman). Başvuru modelleri: `test/collab-doc-model.ts` ve ondan habersiz, bağımsız incelemede yazılmış ikincisi `test/collab-doc-model2.ts` (ikisi de dönüşüm kodunu ve belge komutlarını içe aktarmaz).

**Kapsam dışı (oda kurulmaz, söylenir):** docx olmayan ya da burada dönüştürülmüş dosya; **değişiklik izleyen ya da izlenmiş değişiklik taşıyan belge** (`w:trackRevisions`, izlemeyi zorlayan koruma, `mount({ trackChanges: true })`, ya da izleme kapalıyken bile içinde `w:ins` / `w:del` / `w:rPrChange`… bulunan belge): izlenen bir değişiklik yazarını, tarihini ve kendi `w:id`'sini taşır; bölünen bir değişiklik yeni numarasını katılımcının kendi sayacından alır, iki katılımcıda aynı değişiklik olmaz. Koruması yalnız okumaya / yalnız yoruma izin veren belge odaya **izleyici** (`ro`) olarak katılır.

### Konum

Bir konum `[öykü, yol, ofset]` ya da `[öykü, yol, ofset, 1]` dizisidir. Öykü: `"body"`, `"header:<rId>"`, `"footer:<rId>"`, `"footnote:<id>"`, `"endnote:<id>"`. Yol paragrafın yoludur: gövdede `[blok]`, tablo hücresinde `[tablo, satır, hücre, blok]` (iç içe tablolarda uzar; en çok 96 sayı). Ofset, paragrafın görünen metnindeki UTF-16 birimidir (silinmiş metin ve alan kodları ofset almaz; görsel gibi genişliksiz öğeler iki ofset arasında durur). Dördüncü öğe `1`: konum o ofsetteki genişliksiz öğenin **önündedir**. Bütün konumlar **şimdiki** belgedir.

### `ops`: bir adımın içeriği

1–64 öğelik dizi. Her öğe, yazarının çalıştırdığı **komuttur**; her katılımcı aynı komutu kendi belgesinde çalıştırır (aktarıcı, komutun yazarının çalıştırdığı belgenin aynısına uygulanmasını sağlar).

| `t` | Alanlar | Komut |
|---|---|---|
| `text` | `at`, `to?`, `x` (metin; `\n` paragraf sonu, sekme sekme, `\u000b` satır sonu), `f?` (karakter biçimi yaması) | Yazma / düz metin yapıştırma; `to` varsa aralığın yerine |
| `del` | `at`, `to` | Aralığı sil (Backspace / Delete de, yalnız bir aralık sildiklerinde böyle gider) |
| `split` | `at` | Enter |
| `run` | `at`, `to`, `f` | Karakter biçimi (`bold`, `italic`, `underline`, `strike`, `font`, `size`, `color`, `highlight`, `vertAlign`, `style`) |
| `unrun` | `at`, `to`, `k` (`bold` \| `italic` \| `underline` \| `strike`) | Aralığın **tamamında açık** olan özelliği kapat (artık tamamında açık değilse uygulanmaz) |
| `para` | `at`, `to`, `f` | Paragraf biçimi (`align`, `indent`, `spacing`, `style`, `numbering`, `keepNext`…) |
| `cmd` | `name`, `at`, `to?`, `a` | Adlı komut: `toggleList` `[tür]`, `indentList` `[±n]`, `insertPageBreak` `[]`, `insertTable` `[satır, sütun]`, `insertRow` `["above"\|"below"]`, `deleteRow` `[]`, `insertColumn` `["left"\|"right"]`, `deleteColumn` `[]`, `insertHyperlink` `[adres, ilişki kimliği?]`, `deleteBackward` / `deleteForward` `["char"\|"word"]` (tuşun aralık silmekten başka bir şey yaptığı durumlar: listeden çıkarma, girintiyi geri alma) |
| `put` | `at` (ofseti 0), `n`, `x`, `h` | `at`'taki paragraftan başlayan `n` bloğun yerine `x`'teki düz paragrafları koy (geri almanın taşıdığı biçim; aşağıda). `h`: yerine konan blokların özeti (8 onaltılık hane); alıcıdaki bloklar bu özeti vermiyorsa adım uygulanmaz |
| `cmt` | `k` ve türüne göre alanlar | Yorum komutu (aşağıda "Yorumlar") |

Telde **taşınmayanlar** (yerelde yapılmaz, kullanıcıya söylenir): görsel ekleme / boyutlama, izlenen değişikliği kabul / ret, izlemeyi açıp kapatma, zengin yapıştırma (görsel, alan, dipnot taşıyan parça — metni, biçimi ve tablosu düz yapıştırma yoluyla gider).

### Yorumlar (`cmt`)

Bir yorum komutu, `src/ops/doc/comments.ts` komutlarının (`addComment`, `replyToComment`, `editComment`, `resolveComment`, `removeComment`) aynısıdır; her katılımcı kendi belgesinde çalıştırır. Kod: `src/collab/doc/comments.ts`.

| `k` | Alanlar | Komut |
|---|---|---|
| `add` | `at`, `to?`, `x`, `u`, `i`, `d`, `q` | `at`..`to` aralığına (ya da `to` yoksa imleçteki kelimeye) yorum |
| `reply` | `c`, `x`, `u`, `i`, `d`, `q` | `c`'nin zincirine yanıt |
| `edit` | `c`, `x`, `h?` | Yorumun metni |
| `done` | `c`, `v` (`1` çöz, `0` yeniden aç), `h?` | Zinciri çöz / aç |
| `del` | `c`, `h?` | Yorumu sil (ilk yorum yanıtlarıyla gider) |

- `x`: yorum metni, paragraflar `\n` ile ayrılır, en çok 32.767 karakter ve 1.000 paragraf, XML'de yer alabilen karakterler, `\r` yok (komutun temizlediği haliyle gider). `u`: yazar (1–255 karakter, tek satır), `i`: baş harfler (0–32 karakter, boş: yazılmaz), `d`: tarih (ISO 8601). **Yazar ve tarih yazarın tarafında belirlenir** (`mount({ author })`, yoksa odadaki ad; yazarın saati) ve olduğu gibi uygulanır: alıcı kendi adını ya da saatini koymaz.
- **Yorumun adı (`c`)**: yorumun bilinen paragraf kimliği (`w14:paraId`, 8 büyük onaltılık hane, `00000000` değil, `80000000`'den küçük) ya da kimliği olmayan dosya yorumu için `[numara, özet]` — özet, yorumun yazar + tarihinin 8 küçük onaltılık haneli özetidir (FNV-1a); numarası kalıp başka bir yoruma geçmiş olanı (silinip yeni eklenen) ondan ayırır. Bulunamayan yorum: adım uygulanmaz.
- **Yeni yorumun kimliği (`q`)**: yazar rastgele seçer; belgede (dosyada tutulacağı haliyle: gövde, üstbilgi/altbilgi, notlar, yorumlar — yorumun kimliği son paragrafında yazılır) aynı `w14:paraId` varsa adım herkeste reddedilir.
- **Kimliği olmayan dosya yorumu** bir yanıt ya da çözme ile kimlik alınca, kimlik telde gitmez: herkes yorumdan aynı kuralla türetir — `FNV-1a(stable(["paraId", numara, yazar+tarih özeti, k]))` & `0x7FFFFFFF` (8 büyük onaltılık hane), `k` 0'dan başlayarak belgede (yukarıdaki okumayla) ya da aynı komutta verilmiş olmayan ilk değer. Böylece aynı anda yanıt veren ya da çözen iki üye yoruma aynı kimliği verir ve ardında bekleyen komutlar yorumu bulur.
- **Numara**: telde **yoktur**. Herkes belgeden aynı kuralla hesaplar: dosyada tutulacak en büyük yorum numarasının (`w:comment`, `w:commentRangeStart`, `w:commentRangeEnd`, `w:commentReference` üzerindeki `w:id`; sahipsiz işaretler dahil) bir fazlası; en büyük numara `0x7FFFFFFF` ise yeni yorum ve yanıt herkeste reddedilir (tek kişilik düzenleyici de o belgeye yorum ekleyemez). Sayım, `put` kimlik sayımıyla aynı okumadır (kaydedilecek dosyanın XML'i), sonradan katılan aynı numarayı verir. Katılımcının kendi sayacına (geri alınan yorumların numaraları) bakılmaz.
- **`h`** (yalnız geri alma üretir): yorumun geri alma anındaki halinin özeti; alıcıdaki yorum bu özeti vermiyorsa (o arada biri yanıtladı, düzenledi, çözdü) adım uygulanmaz. Silmede özet, gidecek her yorumu (ilk yorum yanıtlarıyla) kapsar; düzenlemede metni; çözmede zincirin üyelerini ve çözüm durumlarını.
- **Bir adımda en çok 8 yorum komutu** (her biri belgeyi baştan okur).
- **Bütçe**: odada bir belgenin bütün yorumları en çok 4.000.000 karakter ve 50.000 paragraf tutar; yorum ekleyen, yanıtlayan ya da büyüten bir komut bu sınırı aşacaksa herkeste reddedilir (karar belgenin kendisinden; gelen adım atlanır ve `bad-step` söylenir, yerel komut reddedilir). Sınırın üstündeki bir dosyada odada yeni yorum metni eklenemez.
- **Yeniden temellendirme**: yalnız `add` bir yer anar. Aralığı yazarının gördüğü metni kapsar, araya başkasının yazdığı metin içinde kalır (aralık büyür; iki parçaya bölünmez). Aralığın bir kısmı önce silindiyse kalan metin hakkındadır (`partial` söylenir); **tamamı** silindiyse ya da imleçteki yorumun yeri silinmişse **düşer** (`dropped`). Öbür komutlar yer anmaz, yorumu adıyla bulur.
- **Ne yaptığı**: işaretlerin girdiği / çıktığı paragraflar "biçimi değişti, hiçbir şey kaymadı" (`tch`) sayılır; bloklar arasında duran bir işaretin silinmesi o bloğun gitmesidir.
- **Geri al**: `add` / `reply` → `del` (+ `h`); `edit` → eski metinle `edit` (yalnız düz metin yorumda); `done` → tersi (yalnız zincirin hepsi öbür durumdaydıysa); `del` geri alınamaz (`undo-none`).
- **Sınır**: kaydedilmiş dosyadan sonradan katılanın yazdığı dosya üç yerde farklı olabilir: yazıcının kayıtta oluşturduğu parçaların ilişki kimlikleri, kayıt anında oluşmuş ve artık bir şey söylemeyen `commentsExtended.xml`'in korunması, ve (başlangıç dosyasında kimliksiz yorumlar varsa) `commentsExtended.xml` girdilerinin sırası (kılavuz: "Yorumlar odada"). Aynı dosyadan başlayan üyeler aynı dosyayı yazar.

Gelen adım **güvenilmezdir**: her alan türüne ve aralığına bakılarak yeni bir nesneye okunur (öykü adı kalıbı, yol uzunluğu, tamsayılar, metinde XML'de yer alamayan karakter yok, yamalar düz JSON ve derinlik / boyut sınırlı, komut adı listeden, bağımsız değişkenler komutun kendi denetiminden). Uygulanırken belge komutlarının kendi denetimlerinden geçer (konum var mı, stil belgede var mı, bağlantı adresi izinli şemada mı…). Okunamayan ya da belgenin kabul etmediği adım **bütünüyle** geri alınır ve **herkeste aynı biçimde atlanır**, bir kez söylenir. Bir adım en çok 900 KB'tır.

**Bütçeler** (aşan adım okunurken ya da uygulanırken **herkeste aynı biçimde** reddedilir; başkasından geldiyse atlanır ve bir kez `bad-step` söylenir): adımda en çok 64 komut; bir `text` en çok 800.000 karakter ve 5.000 paragraf sonu; `insertTable` en çok 100 satır × 63 sütun; bir paragraf en çok 1.000.000, bir öykü (gövde, üstbilgi…) en çok 20.000.000 karakter; bir öyküde en çok 200.000 blok ve 100.000 tablo hücresi.

**`put` paragrafı** `{ p?, m?, a?, r }`. Değerleri, biçim değiştiren komutların (`para`, `run`) okunduğu kuralla, **aynı işlevlerle** (`patchParaProps`, `patchRunProps`) okunur; bu kuralın almadığı hiçbir değer yazılamaz:
- `p`: paragraf biçimi, `para` komutunun yaması gibi (`align`, `indent`, `spacing`, `style`, `numbering`, `keepNext`, `keepLines`, `pageBreakBefore`, `widowControl`, `contextualSpacing`, `outlineLevel`). Stil belgenin paragraf stili, liste belgenin listesi olmalıdır.
- `m`: paragraf işaretinin biçimi, bir çalıştırma biçimi gibi (aşağıda). Karakter stili belgede olmalıdır.
- `a`: öznitelikler; yalnız `w:rsid*`, `w14:paraId`, `w14:textId`, değerleri onaltılık. Ön ek parçada bildirilmiş olmalıdır.
- `r` satır içi öğeleri:
  - `{ p?, a?, c }` bir çalıştırma. `p` biçimdir: `f` `run` komutunun yaması gibi (stil belgenin karakter stili olmalı); `cs` karmaşık yazı ayrı durduğunda `{ bold?, italic?, size? }`, aynı kuralla; `fonts` dört yuvanın adları ve `hint` (`default`/`eastAsia`/`cs`); `lang` dil etiketleri. `c` metin parçalarıdır (`"metin"` ya da `["metin", "kaynak yazımı"]`) ve sayıyla öğeler (`0` sekme, `1` satır sonu, `2` sayfa sonu, `3` sütun sonu, `4` bölünmez tire, `5` yumuşak tire, `6` son çizilen sayfa sonu).
  - `{ e }` yazım işareti; yalnız `spellStart`, `spellEnd`, `gramStart`, `gramEnd`.
  - `{ b: [0|1, kimlik, ad?] }` yer imi başlangıcı / sonu.
  - `{ k: [0|1, numara] }` yorum aralığının başlangıcı / sonu (numara rakam dizgesi, 0–2³¹−1); çalıştırmada `c` öğesi olarak `{ r: numara }` yorum başvurusu. Bir `put`'un getirdiği yorum işaretleri (başlangıç, son, başvuru; numaralarıyla) yerine koyduğu blokların taşıdıklarıyla **birebir aynı ve aynı sırada** olmalıdır (belge sırasıyla; hangi paragrafta durdukları sorulmaz — bir Enter'ın ya da birleştirmenin geri alınması aynı işaretleri başka paragraflarda getirir), değilse adım herkeste reddedilir: ters dönmüş aralık, yer değiştirmiş numaralar, kesişen aralıklar gelemez; yerine konan blokların özeti (`h`) işaretleri de kapsar. Böylece yorum işareti taşıyan paragraftaki silme / biçim / Enter geri alınabilir; o arada yoruma dokunan (silen, aynı paragrafa yorum ekleyen, yanıt veren) adım geri almayı düşürür (`undo-gone`), silinmiş bir yorumun işareti geri gelmez.

Değerler belgeyle birlikte denetlenir (gönderen ve alıcı aynı işlevle). Sayım, belgenin **kaydedilecek dosyada yazılacağı haliyle** yapılır: her blok docx yazıcının onu yazdığı (ya da değişmemiş bloğun özgün baytlarını kopyaladığı) XML'den, blokların çevresi (kök, son bölüm, saklanan öğeler, arka plan; not ve yorum parçalarının çevresi) yazıcının yazdığı haliyle, yazılmayacak (değişmemiş) parçalar özgün kaynaklarıyla okunur; herhangi bir öğedeki `w14:paraId`, `w:bookmarkStart` (kimlik, ad) ve `w:bookmarkEnd` (kimlik) sayılır. `mc:AlternateContent`'in iki dalı da okunur, aynı değer iki dalda geçiyorsa bir kez sayılır. Test denetleyicisi (`test/tools/package_check.py`, `named()`) kaydedilen dosyayı aynı tanımla okur. Okunamayan (XML olmayan, çok büyük, çok derin) belge "boş" sayılmaz: `put` reddedilir. Bir bloğun sayımı blok nesnesi yaşadıkça saklanır (bloklar değiştirilmez, yenisi yapılır). `put`'tan sonra her paragraf kimliği, yer imi başlangıcı ve sonu (kimlik **sayı olarak**: `09` ile `9` aynı kimliktir) ve yer imi adı (harf büyüklüğü gözetilmeden) belgede **en çok bir kez** geçer — ya da `put`'tan önce kaç kez geçiyorsa en çok o kadar (dosyanın kendi yinelemesi kalabilir, artamaz); `put`'un dokunduğu her yer iminin `put`'tan sonra belge sırasında önce tam bir başı, sonra tam bir sonu vardır (ya da öncekiyle aynı dizilişi). `put`'taki yer imi başlangıcının adı zorunludur, kimliği öne sıfırsız 0–2147483647'dir. rsid tam 8 onaltılık hane, `w14:paraId` / `w14:textId` tam 8 hane ve 80000000'ın altında olmalıdır; yerine konan blokların zaten taşıdığı değerler (öznitelik adı=değer) bu biçim kuralından muaftır. Bir adımda en çok 16 `put` bulunur ve bunların çıkardığı ile koyduğu toplam en çok 4.000 paragraftır (ayrık 17 paragrafa dokunan tek bir geri alma kaydı bu yüzden geri alınamaz); aşan adım geri almada `undo-none` olarak söylenir. Aynı anahtarla birleşmiş bir kaydın (IME bileşimi gibi) art arda aynı paragraflara dokunan `put`'ları geri alınırken tek `put`'ta birleşir.

Yerine konacak bloklar da yalnız böyle düz paragraflar ya da yalnız onlardan oluşan tablolar olabilir. Bu biçimle anlatılamayan özellik taşıyan paragraf (kenarlık, gölgeleme, sekme durakları, `cnfStyle`, renkli alt çizgi…) yolculuk etmez: o paragrafı değiştiren adım birlikte düzenlenirken geri alınamaz (`undo-none`).

### Yeniden temellendirme

Bir blok listesi (gövde, bir üstbilgi, bir hücrenin içi) tek bir uzun metin olarak okunur: paragraf bir satırdır, paragraf sonu herhangi bir karakter gibidir. Yazma, silme, Enter ve iki paragrafı birleştirme o metnin ekleme ve silmeleridir. Her komut uygulandıktan sonra ne yaptığı **belgenin kendisinden okunur**: aralığın önündeki ve arkasındaki metin yerinde duruyor ve arası yalnız paragrafsa "şuraya şu kadar metin geldi / şu aralık gitti" denir; bu kesin görülemiyorsa "şu bloklar değişti", o da görülemiyorsa "bu öyküde bir şey değişti" denir.

| Kural | |
|---|---|
| Aynı yere iki ekleme | İkisi de kalır; aktarıcının **sonra** sıraya koyduğu önde durur |
| Silme × araya giren ekleme | Silme yalnız yazarının **gördüğü** metni siler; aralığın içine başkasının eklediği metin kalır (silme onun iki yanında iki aralık olur) |
| Ekleme × araya giren silme | Silinen aralığın içine yazılan metin kalır, aralığın olduğu yerde |
| İki silme | Örtüşen kısım bir kez silinir; hiçbir şey kalmayan silme **sessizce** biter (istenen zaten olmuş) |
| Aralığın yerine yazma (`text` + `to`) | Aralıktan kalanı siler; metin aralığın başladığı yere, oraya başkasının eklediğinin **önüne** girer |
| Karakter biçimi | Silme gibi: yazarının gördüğü metni kapsar. Aynı metne iki biçim: sonra gelen kazanır |
| Paragraf biçimi, liste | Paragrafı (aralığın değdiği paragrafları) kapsar; imlecin iki yanındaki metin başkasınca **paragraf birleşerek** silindiyse **düşer** (kastedilen paragraf artık yok) |
| Enter | Paragraf sonu eklemesidir |
| Aynı tabloda yapı değişikliği (satır / sütun ekle-sil) × o tablonun içindeki her değişiklik | Sonra gelen **düşer** |
| "Şu bloklar değişti" olarak bilinen her komut (tablo ekleme, bir tabloyu kesen silme, tuş olarak giden Backspace) × aynı bloklara değen değişiklik | Sonra gelen **düşer**; başka bloklardaki değişiklikler yalnız kayar |
| "Bu öyküde bir şey değişti" × aynı öyküdeki her değişiklik | Sonra gelen **düşer** |

**Kanıtlanmamış birleşim düşer**: düşen her adım kullanıcıya söylenir (`dropped`). **Düşen adımın ardında bekleyen adımlar da düşer** (düşenin üzerine yapılmışlardır; nereye gideceklerini kanıtlayamayız) ve sayıya katılıp söylenir. Tek ayrık: düşen adım hiçbir şeyi kaydırmamışsa (yalnız karakter biçimi), ardındakiler yerinde kalır.

**Kısmen yapılamayan değişiklik söylenir** (`partial`): bekleyen bir silme, aralığın yerine yazma ya da karakter biçimi, kapsadığı metnin bir bölümü (ya da tamamı) başkasınca aynı anda silindiği için ancak kalanında yapılabildiyse (ya da hiç yapılamadıysa). Öbürü silmesini geri alırsa o metin bu değişiklik **olmadan** geri gelir: silme de, biçim de orada yapılmamıştır. Söylenmeden kaybolmaz.

Başvuru modellerinin doğruladığı aile: `text`, `del`, `split`, aralığın yerine yazma, `run`, imleçteki `para`, geri al / yinele (2–5 katılımcı, uzun kuyruklar, kopan bağlantılar, rastgele teslim; birinci model her iletiden sonra her katılımcının gördüğü metni ve her söylenen `dropped` / `partial` sayısını, ikinci model her adımda yazarın kastettiğini — önce / sonra — yerleşen belgede arar). Öbür aileler (`cmd`, tablo içi) için sınanan: yakınsama, düşenin söylenmesi, dosyanın sağlamlığı.

### Geri al / yinele

Geri alma **sıradan bir adımdır**: katılımcı, kendi yığınındaki girdinin tersini **komut olarak** üretir ve gönderir; alan taraf geçmiş tutmaz, sonradan katılan da onu her adım gibi uygular. Ters iki biçimdedir:

- Yalnız metin eklemiş bir adımın (yazma, Enter, yapıştırma) tersi `del`: eklenen aralığı sil.
- Geri kalan her şeyin (silme, biçim, paragraf biçimi, birleştirme, satır sonu) tersi `put`: adımın değiştirdiği paragrafları, adımdan önceki halleriyle (çalıştırmalar, biçim, yer imleri) yerine koy. Herkes — gönderen de — aynı tel verisini uygular.

Tersi komutla anlatılamayan adım (liste yapma / kaldırma, bağlantı, tablo yapısı, tablo ekleme sonrası yineleme, görsel / alan / dipnot taşıyan ya da `put` biçimiyle anlatılamayan paragraf) geri alınamaz ve söylenir (`undo-none`); altındaki adımlar, o adım hiçbir şeyi kaydırmadıysa geri alınabilir kalır.

Geri alma ya da yineleme yoldayken gelen adım yığınlara **her yığının kendi koordinatıyla** okunur: geri almanın çıktığı yığın, geri almadan SONRAKİ belgeye göre; geri almanın oluşturduğu girdinin altındakiler (öbür yığın) ondan ÖNCEKİ belgeye göre.

**Bağlantının ilişkisi** telde gider: `insertHyperlink` `[adres, ilişki kimliği]` (`rId` + 1–32 harf/rakam; yazarı rastgele seçer). Kimlik alıcının belgesinde varsa adım herkeste reddedilir. Kimliksiz gelirse komutun kendisinden türetilir. Böylece dosyadan sonradan katılan, bağlantıya herkesle aynı ilişkiyi verir; yazılan bir ilişki parçası ilişkilerini kimliklerine göre sıralar (hangi dosyadan açılmış olursa olsun aynı dosya).

**İmleç**: başkasının `put` ile yerine koyduğu paragraflarda imleç düz metne göre eşlenir: eski ve yeni metnin ortak önekindeyse aynı ofset, ortak sonekindeyse sondan aynı uzaklık, arada ise değişen bölümün sonu.

Yığın **LIFO**'dur. Girdi katılımcının kendi sonraki adımlarından geçirilmez (önce onlar geri alınır); **başkalarının** adımlarından geçirilir. Başkasının o arada değiştirdiği yer geri alınmaz: `put`'un paragraflarına başkası dokunduysa ya da `del`'in aralığına başkası yazdıysa girdi düşer ve söylenir (`undo-gone`); düşen girdi bir şeyi kaydırmışsa altındakiler de gider (onlar onsuz bir belgeye göre saklanmıştı). Yoldaki geri alma düşerse (`undo-gone`) aynı kural uygulanır. Katılımcının adımı yoldayken istenen geri alma, adımlar yerleşene dek bekler (`undo-wait`). Ardışık yazılan harfler (aynı paragrafta, bir öncekinin bittiği yerde) aktarıcı yanıt vermeden önce **tek adımda** birleşir: bekleyen adım sayısı yazma hızıyla büyümez.

- Değişmez (`test/collab-doc-alone.test.ts`): odada başkası bir şey yapmıyorken (geç yanıt veren aktarıcı, yanında okuyan biri, araya karışan yazma / geri al / yinele dahil) her işlem / geri al / yinele dizisi `collab`'sız düzenleyicinin bıraktığı belgeyi blok blok bırakır ve hiçbir şey söylenmez — `undo-none` olanlar dışında (yukarıdaki aile; testin üreteci onları üretmez). Yazılan dosya **belge olarak** aynıdır; geri almanın koyduğu paragraflar modelden yazıldığından bayt bayt aynı olmayabilir.
- Oda yeniden yükleme döngüsüne giremez: art arda yüklemeler arasında hiçbir adım işlenmediyse sayılır, altıncıdan sonra `stale` söylenir ve oturum `detached` olur (düzenlenen kalır).

### `presence.data` (belge)

`{ "st": öykü, "a": [yol, ofset], "f": [yol, ofset] }` (seçimin iki ucu; imleçte aynı yer) ya da `null`. Alan öbürlerinin sayfasında imleç çizgisi + ad etiketi ve seçim boyası olarak gösterilir; araya giren değişikliklerle birlikte kayar.

### Dosya ve kayıt

Kayıt kuralları tablolardakiyle aynıdır (tek kaydedici, nöbet, kanıt). Kanıt, iki katılımcının aynı adımda **aynı dosyayı** yazmasına dayanır: aynı dosyadan başlayıp aynı adımları tutan iki katılımcı parça parça (ve ölçülen koşularda bayt bayt) aynı docx'i yazar. Odaya **kaydedilmiş dosyadan** katılan katılımcının yazdığı dosya çoğu zaman aynıdır ama garanti değildir (dosyadan okunan belge ile düzenlenerek oluşan belge blok blok aynı nesneler değildir); tutmadığında o katılımcı işini nöbetle kendi kaydeder.
