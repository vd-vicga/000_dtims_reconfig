# Belægninger 2026, linear search og dTIMS

Noter fra arbejdet 6. november 2026. Gælder projektmappen `000_dtims_reconfig` og databasen `udv_dTIMS`.

## Opgaven

Udgangspunktet var `posted_type.py`, som henter skiltet vejtype og geometri fra Vejdirektoratets linear search og skriver en geopackage.

Samme kald skulle bruges på `Nye belægninger på kørebaner i 2026.xlsx`. Nøglerne i arket er vejnummer, vejdel, vejside og fra-/til-kilometrering. Geometrien skulle tilbage i arket og ind i databasen. Derefter skulle hver belægning have measure fra `dtims.base`, og der skulle ligge hændelsestabeller, som kan bruges på både `dtims.base` og `dtims.pms_analysis`.

`pms_analysis` er delt i sektioner. Samler man sektioner med samme vej og side, får man de lange ruter, som `base` består af.

## To referencesystemer

Kilometrering og measure er ikke det samme. Det blev aftalt eksplicit, og scriptet blander dem ikke.

- Kilometrering er `FRAKMT`/`TILKMT` fra vejman og linear search. Den bruges kun til at hente og klippe geometrien.
- `dtims.base.from` / `to` er samme stationering som `dtims.pms_analysis.from` / `to`. For vej `3-0` starter begge ved 801.
- `dtims.base.frommeasure` / `tomeasure` er målet langs det enkelte base-element. Det starter ved 0. For alle 1 044 base-rækker er `to - from` lig `tomeasure - frommeasure`.

En hændelse har derfor begge tal: `from`/`to` i base- og sektionsstationeringen, og `measure_from`/`measure_to` læst på base-linjen.

Linjens længde i meter behøver ikke være lig `TILKMT - FRAKMT`.

## Lane

Samme koder i belægninger, `pms_analysis` og Vaisala-eksporten:

| Vejside | Kode i sektionsnavn | lane |
|---|---|---|
| Midt | B | 0 |
| Højre | R | 1 |
| Venstre | L | 2 |

I linear search er vejside et tal: `-1` Venstre, `0` Midt, `1` Højre.

`base` har ingen lane og ingen side. Match mod base sker på vejnavn (`9-2`) og geometri. Side bruges bagefter til at vælge sektioner i `pms_analysis`.

Sektionsnavne ser ud som `3-0_L_801`: vej, side, fra-stationering.

## Linear search

Ét endpoint, ingen header:

`GET https://linearsearch.vd.dk/linear-search-api/v4/execute?query=...`

Timeout mindst 180 sekunder. Svaret er GeoJSON i EPSG:25832, koordinater `[x, y, 0]`.

Tre kald, gemt som skills i `C:\Users\b307007\.cursor\skills\`:

- `vd-linear-search-attributes` henter vejnummer, vejdel, vejside og kilometrering uden geometri.
- `vd-linear-search-geometry` er samme kald med `geometry`. Tjenesten returnerer vejdelens referencestykker. Et `where` på fra-/til-kilometrering klipper ikke et delinterval ud. Scriptet henter vejdelen og klipper lokalt.
- `vd-linear-search-theme` henter et tema, for eksempel `vm:VejtypeSkiltet`, for et filtreret net. Det er mønstret i `posted_type.py`.

`Bestyrer = 0` er statsveje. Et kald uden `where` returnerer hele nettet og må ikke sendes. Et ukendt `vm:`-felt giver HTTP 400 med forslag. Der er ikke et katalog-endpoint. `geometry` er linjen. `vm:Geometri` er et andet felt.

## Base, som den faktisk ser ud

`dtims.base` har 1 044 rækker og 370 veje. Vejnavnet i `name` matcher `pms_analysis.roadname` for alle 370.

Relevante kolonner:

- `gid` er rækkens nøgle. `id` er en uuid og er ikke unik: samme id kan ligge på flere stykker af samme vej.
- `name`, for eksempel `3-0` eller `332-1`.
- `from`, `to`, `length`.
- `frommeasure`, `tomeasure`.
- `geom` er `geometry(LineString, 4326)`.
- `validto` og `endedon` var tomme på alle rækker, da data blev læst.

En vej kan have flere base-elementer, når der er huller. `3-0` har tre: 801–11287, 12840–18282 og 19627–26500.

## Hvad scriptet gør

`belaegning_2026.py` læser regnearket, kalder linear search og skriver både Excel og database. Svarene caches i `_linear_cache\`. Mappen er ignoreret i git.

For hver udlægning:

1. Hent referencestykker for vejnummer, vejdel og vejside.
2. Klip dem til fra-/til-kilometrering og saml dem. Huller bliver en MultiLineString.
3. Find base-elementer med samme vejnøgle, hvis belægningens linje ligger inden for 15 m.
4. Læs measure på base-linjen mellem de punkter, der rammer elementet. Linjen antages at løbe fra `frommeasure` mod `tomeasure`.
5. Omsæt det mål til `from`/`to` i base-stationeringen.
6. Del hændelsen på de `pms_analysis`-sektioner med samme vej og lane, som intervallet overlapper.

Hændelsens geometri er udsnittet af base-linjen, transformeret til EPSG:25832. Geometrien i hovedtabellen og i Excel er linear search-linjen.

Scriptet sletter og genopretter de tre tabeller, når det køres igen.

## Tabeller

- `dtims.nye_belaegninger_2026` er én række pr. udlægning.
- `dtims.nye_belaegninger_2026_base` er hændelser på base-elementer. `from`/`to` er base-stationering. `measure_from`/`measure_to` er mål på elementet.
- `dtims.nye_belaegninger_2026_pms` er samme hændelser delt på sektioner. `name` er sektionsnavnet.

## Hvad der ligger i databasen nu

Først blev alle 560 udlægninger indlæst. 560 havde geometri. 405 blev stedfæstet. 178 belægninger dækkede 370 sektioner. Ingen belægning ramte mere end ét base-element.

Derefter blev rækker med vejdel over 2, og rækker med bogstav i vejdelen, slettet. Det var 151 rækker. De havde ingen rækker i de to hændelsestabeller.

Tilbage:

- 409 udlægninger, kun vejdel 0, 1 og 2 (369, 20 og 20).
- 405 med mål fra base.
- 405 hændelser på base.
- 370 hændelser på sektioner.
- 4 udlægninger, hvis vej findes i base, men som geometrien ikke kunne stedfæste.

Veje som ikke findes i base, og som blev fjernet, er overvejende vejdel 3–12 samt vejdele som `4A`, `5A` og `3A`. Det blev accepteret, fordi `base` ikke indeholder de vejdele.

Excel-filen blev ikke filtreret. Den har stadig alle 560 rækker, også vejdel over 2.

## Excel

`Nye belægninger på kørebaner i 2026.xlsx` har de oprindelige kolonner plus:

`GEOMETRY`, `ROADNAME`, `lane`, `BASE_ID`, `BASE_NAME`, `MEASURE_FROM`, `MEASURE_TO`, `PMS_SEKTIONER`.

`lane` i det ark kommer fra vejside: Midt 318, Højre 122, Venstre 120. Fem lange geometrier er forenklet i cellen, fordi Excel ikke rummer den fulde tekst. Den fulde geometri ligger i databasen.

`pms_analysis_vaisala.xlsx` fik også en `lane`-kolonne, som kolonne J. Den er udledt af sidekoden i sektionsnavnet, ikke slået op i databasen, fordi forbindelsen var nede da det blev gjort. Fordelingen er lane 0: 1 029, lane 1: 991, lane 2: 959. Det passer med antallet af B-, R- og L-sektioner. Den blev ikke krydstjekket mod `pms_analysis.lane`.

## Forbindelse til databasen

Postgres ligger på `10.98.82.57:5432`, database `udv_dTIMS`, bruger `postgres`. Adgangskoden ligger i `exp.env`, som er ignoreret i git. Serveren understøtter ikke SSL, så linjen i `pg_hba.conf` skal være `host`, ikke `hostssl`.

Maskinens IP har skiftet undervejs:

- `10.175.32.199` blev afvist, derefter tilføjet, og `pg_reload_conf()` gav true. Indlæsningen kørte på den adresse.
- Senere var adressen `10.175.47.252`, og den blev afvist. Vaisala-arket kunne derfor ikke joines mod databasen.

`SHOW hba_file;` viser, hvilken fil serveren bruger. Postgres bruger den første linje, der matcher.

## Forbehold

Retningen på base-linjen er antaget, ikke bevist. Længden af de klippede hændelsesgeometrier lå tæt på målintervallet (gennemsnit 1,00, spænd 0,85–1,09), men det siger ikke, om linjen er vendt. Det bør ses efter på et par veje.

Op til 7 belægninger kan ligge på samme sektion. Det er flere udlægninger på samme sted, ikke en fejl i opdelingen.

`base.id` kan gå igen på flere gid-rækker. Hændelserne peger på den id, der blev ramt, sammen med det konkrete from/to-interval.
