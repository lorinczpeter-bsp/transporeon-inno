# Hartmann → InnoTMS

## Futtatás Streamlit Community Cloudon

1. Másold a csomag teljes tartalmát egy GitHub-repozitóriumba: `app.py`, `requirements.txt`, `.streamlit/config.toml` és az `assets` mappa. A logó és a betűtípusok az assets mappában vannak.
2. A Streamlit Community Cloudon válaszd ki a repozitóriumot, fő fájlként az `app.py`-t, Python 3.12-t vagy újabbat.
3. A futó alkalmazásban töltsd fel a Transporeon-exportot.
4. Töltsd fel a **Hartmann fuvardíjak** Google-táblázatból letöltött XLSX-et, vagy használd a Google-kapcsolatot.
5. A „Feltöltési időszak” részen igény szerint kapcsold be az odaítélési dátumszűrést, és/vagy válassz kirakodási dátumtartományt vagy teljesítési heteket.
6. Rendezd a jelzett dátumeltéréseket és útvonalválasztásokat. Ellenőrizd a listát, vedd ki a nem importálandó fuvarokat, majd töltsd le az InnoTMS-importot.

A Final External ratesheet **Routes Overview (Road)** adatai és az InnoTMS importsablon be vannak építve az egyetlen `app.py` fájlba. Ezek az oldalsávban új XLSX feltöltésével lecserélhetők. Az eredeti 47 mintafuvar nincs a beépített sablonban.

## Google Sheets, zárt megosztással

A ChatGPT Google-hozzáférése nem kerül át a Streamlit alkalmazásba. Két lehetőség van:

- **Feltöltés:** a Google Sheetsben Fájl → Letöltés → Microsoft Excel (.xlsx), majd feltöltés az alkalmazásban. Ez megőrzi a cellák teljes pontosságú értékét.
- **Közvetlen olvasás:** Google-szolgáltatásfiók kulcsát add meg a Streamlit Cloud **Secrets** felületén `google_service_account` néven. Engedélyezd a Google Sheets API-t, és a díjtáblázatot oszd meg a szolgáltatásfiók `client_email` címével olvasásra. Az alkalmazás csak olvasási jogosultságot kér. A kulcsot ne tedd az app.py-ba vagy GitHubra. A szokásos szolgáltatásfiók-JSON összes mezője kerüljön az azonos nevű Secrets-szekcióba.

Szolgáltatásfiók nélkül a közvetlen olvasás csak a már bejelentkezés nélkül olvasható táblázatnál működik. Az alkalmazás nem módosítja a táblázat megosztását. A frissítés kézi gombnyomásra történik; hiba esetén nincs korábbi díjra történő automatikus visszaállás.

## Rögzített szabályok

- Az odaítélési dátumszűrés az „Odaítélés dátuma (Fuvar szintje)” mezőre vonatkozik, mindkét szélső nappal együtt. Az időpont órája nem zárja ki a kiválasztott nap fuvarjait.
- A teljesítési időszak alapja az eredeti „Kirakodás dátuma (eddig)” mező. Dátumtartomány vagy egy/több ISO-hét választható; a hétjelölés tartalmazza az ISO-évet, hétfőtől vasárnapig. Az odaítélési és teljesítési feltétel együtt is alkalmazható, ilyenkor mindkettőnek teljesülnie kell ugyanannál a szállítmánysornál.
- A szűrés fuvarszinten működik: ha legalább egy szállítmány minden aktív feltételnek megfelel, a teljes fuvar minden szállítmánya megmarad, változatlan összesített súllyal és cbm-mel. Részleges dátumegyezésnél ezt külön jelzés mutatja. Az aktív szűréshez hiányzó vagy hibás dátumot szintén jelzi az app.
- A szűrés a forrás dátumaihoz kötődik, a későbbi kézi dátumválasztástól nem ugranak ki/be fuvarok. Csak a szűrésben látható és kijelölt fuvarok kerülhetnek az exportba. Üres találati lista vagy befejezetlen időszakválasztás esetén nincs letöltés.
- A fuvarszám az összevonás kulcsa. A különböző szállítmányok súlya, cbm-je és rakodási egysége összeadódik, akkor is, ha a mennyiségek véletlenül azonosak. Az ismétlődő szállítmányazonosítót hiba jelzi.
- Eltérő berakodási vagy kirakodási időintervallumnál fuvaronként dátumot kell választani. Az eredeti szállítmánysorok az ellenőrzésnél megtekinthetők.
- Egyéb mezőeltérésnél ki kell választani, melyik szállítmány törzsadatai érvényesek. Automatikusan nem dobunk el eltérő adatot.
- Egyébként a forrás „ettől” napja kerül át. Mindkét dátum szerkeszthető a fuvarlistában. A kirakodás nem előzheti meg a berakodást.
- Az útvonalat a Routes Overview (Road) irányítószám-szabályai és a Rates táblában szereplő felrakó-/lerakóországok alapján keressük. Csak a díjtáblában szereplő útvonalakat ajánljuk fel. Több találat vagy hiányzó szabály esetén a felhasználó választ az adott ország árazott útvonalai közül. Azonos ország/irányítószám esetén a döntés több fuvarra is vonatkozik.
- 0 < cbm ≤ 70: az összevont mennyiség szorozva az adott sáv EUR/cbm árával. 70 cbm felett fix díj.
- A sávok folytonosak: például 10 < cbm ≤ 20 a „10,1–20” sorhoz tartozik. A cbm-et nem kerekítjük tizedesre. Ez az alkalmazásban használt határérték-szabály.
- A Rates utolsó kitöltött, `Price` kezdetű fejlécű oszlopát használjuk az egész táblában, akkor is, ha annak dátuma korábbi a fuvar dátumánál. A használt oszlop látható. Egy hiányzó cellát nem pótolunk korábbi havi díjjal.
- A teljes pontosságú díjjal számolunk; csak a fuvardíj végösszege kerekül 2 tizedesre, hagyományos kerekítéssel.
- Hollandiában a 7030-as irányítószám az NL161B útvonalhoz tartozik; ez elsőbbséget kap az NL161 általános „All ZIP codes” szabályával szemben. A javított Google-táblában a 285–291. sorok a 70 cbm-ig érvényes NL161B egységárakat, a 292. sor a 70 cbm feletti fix díjat tartalmazza. Ha az NL161B tarifa hiányzik, nem helyettesítjük az általános NL161 díjjal. Esetleges jövőbeli azonos útvonalkódú többszörös díjblokkokat `[1]`, `[2]` jelöléssel és az eredeti leírással különítünk el.
- Hiányzó tarifa esetén indoklással pozitív kézi fuvardíj adható meg fuvaronként. A kézi ár elsőbbséget kap a számított árral szemben, és külön jelzés jelenik meg. A megadott díjtáblában nincs németországi díj; három mintafuvarhoz kézi ár vagy az importból kizárás szükséges. A kézi díj indoka a munkamenetben marad, az eredeti importformátumba nem adunk hozzá új oszlopot.
- Nincs automatikus további üzemanyag- vagy egyéb felár. A kiválasztott ároszlop díja a számítás alapja.

## InnoTMS-kimenet

Az import az eredeti 30 oszlopot, azok sorrendjét, a két „Súly” oszlopot és a rejtett TP-Meta munkalapot őrzi meg. Mindkét súlymezőbe az összevont kg kerül; a „Szállítmányok” az összevont szállítmánysorok száma. A fuvardíj EUR-ban értendő, az importban numerikus érték.

A felület a jutalékelszámolási alkalmazás Bábolna Sped arculatát használja: sötétkék (#2A3756), sárga (#F2E47D), világos háttér, eredeti logó, Fira Sans és Montserrat betűtípusok. Az importfájlba nem kerül arculati elem. Az eredeti importsablon fejléc- és adatsorformátumait, számozási/dátumformátumait, oszlopszélességeit és munkalap-beállításait használjuk; nincs hozzáadott szűrő, rögzített fejléc vagy új színezés.

A „Megbízó további számai” eredete a megadott forrásból nem azonosítható: kezdetben üres, kézzel kitölthető. A rendszám a forrásból jön; az azonosítatlan céloldali státuszok és az „Utoljára módosítva” üresen maradnak. Az odaítélés időpontját nem nevezzük át módosítási időpontnak.

Az export csak akkor érhető el, ha minden **kijelölt** fuvarnak van megfelelő dátuma, útvonala, súlya, térfogata és ára. A hibás tételek kijelölése kivehető; nincs néma kihagyás. Az alkalmazás nem tudja, mely fuvarok szerepelnek már az InnoTMS-ben: ezt a kijelöléskor ellenőrizni kell.

Az alkalmazás nem tölt fel automatikusan az InnoTMS-be. A kimenet tényleges elfogadását először néhány fuvarral ellenőrizd, különösen a forrásból nem azonosított célmezők miatt. A dátum- és útvonaldöntések a böngésző aktuális munkamenetében élnek; tartós archívum és korábbi importokat követő adatbázis ebben a verzióban nincs.

## Helyi futtatás

`pip install -r requirements.txt`

`streamlit run app.py`

A csomag nem tartalmaz BAT-indítót vagy helyi környezethez kötött fájlokat.
