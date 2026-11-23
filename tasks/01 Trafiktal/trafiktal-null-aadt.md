# Trafiktal til PMS-sektioner uden AADT

Skrevet 23. november 2026. Gælder `udv_dTIMS` på `10.98.82.57` og projektmappen `000_dtims_reconfig`.

## Formål

`dtims.pms_analysis` havde 84 sektioner med `aadt IS NULL` (samlet længde 43.537 m). Målet var at hente årsdøgntrafik (trafiktaltype 2) fra Vejdirektoratets linear search og skrive værdierne ind i `aadt`.

## Forudsætninger

- Databaseforbindelse fra `exp.env` (`POSTGRES_*` / `DATABASE_URL`).
- Linear search: `GET https://linearsearch.vd.dk/linear-search-api/v4/execute`
- Script: `_fetch_trafiktal_null_aadt.py` i projektroden.
- Cursor-skills: `vd-linear-search-attributes` / `vd-linear-search-theme`.

## Proces

### 1. Subset af null-AADT

```sql
CREATE TABLE dtims.pms_analysis_aadt_null AS
SELECT * FROM dtims.pms_analysis WHERE aadt IS NULL;
```

Resultat: 84 sektioner, 43.537 m.

### 2. Linear search-kald

Select og filter-mønster (trafiktaltype 2 = ÅDT):

```text
vm:admvejnr, vm:admvejdel, vm:TPE_vejside, vm:frakmt, vm:tilkmt,
vm:Trafik_trafiktal, vm:Trafik_trafiktaltype, vm:Trafik_tælleår
where
vm:admvejnr = <vejnr> AND vm:admvejdel = '<vejdel>' AND vm:Trafik_trafiktaltype = 2
```

Kaldet kørte pr. unik `roadname` (admvejnr + admvejdel) for de 84 sektioner. Svar caches under `_linear_cache/trafiktal/` (gitignored).

`TPE_vejside`: `-1` venstre, `0` midt, `1` højre.

### 3. Kilometrering: brug kmstring, ikke PMS from/to

PMS-stationering (`from`/`to`) kan **ikke** sammenholdes direkte med Vejman `Frakmt`/`Tilkmt`. Offsettet varierer fra vej til vej.

I stedet bruges `kmstringfrom` / `kmstringto`. Konvertering: fjern skråstregen.

| kmstring | Vejman-kmt |
|---|---|
| `58/0637` | `580637` |
| `3/1022` | `31022` |
| `6/0595` | `60595` |

### 4. Matchregel

For hver sektion:

1. Map `lane` → `TPE_vejside`: lane 0→0 (midt), 1→1 (højre), 2→-1 (venstre).
2. Find trafikstykker på samme vej/del/side med overlap mellem sektionens kmstring-interval og `Frakmt`/`Tilkmt`.
3. Vælg trafiktal: længdevægtet; hvis én værdi dækker ≥80 % af overlappet, bruges den direkte.
4. **Midt-fallback** (`TPE_vejside = 0`) blev prøvet for R/L uden side-match. Den gav **0** ekstra hits.

Mange manglende R/L-sektioner har trafik på **vejdel 0**, ikke på vejdel 1/2. Det blev **ikke** brugt som fallback (bevidst valg efter midt-testen).

### 5. Hjælpetabel

Resultatet ligger i `dtims.pms_analysis_aadt_null_trafiktal` med bl.a. `trafiktal`, `taelleaar`, `match_source` (`side` / `midt_fallback`), kmstring og overlap.

### 6. Indskrivning i PMS

23. november 2026:

```sql
UPDATE dtims.pms_analysis AS p
SET aadt = t.trafiktal,
    modifiedon = CURRENT_TIMESTAMP,
    modifiedby = 'trafiktal_linear_search_2026-11-23'
FROM dtims.pms_analysis_aadt_null_trafiktal AS t
WHERE p.gid = t.gid
  AND p.aadt IS NULL;
```

Samme værdier blev også skrevet til `dtims.pms_analysis_aadt_null.aadt`.

## Resultat

| | Antal | Længde |
|---|---|---|
| Null-AADT i udgangspunktet | 84 | 43.537 m |
| Opdateret med linear search | **51** | **35.261 m** |
| Stadig uden AADT | **33** | **8.276 m** |

På de 51 matchede: ÅDT fra ca. 1.980 til 77.099 (gns. ca. 23.000). Tælleår: 20×2025, 19 ældre, 12 uden år.

### Stadig uden match

Typisk:

- Ingen type-2-trafik i API for vejdelen (`143`, `544`, `700`, `200518`).
- R/L på vejdel 1/2, hvor tallene kun findes på vejdel 0 (fx dele af `140`, `142`, `145`, `415`, `505`, `501`).

## Genkørsel

```text
python _fetch_trafiktal_null_aadt.py
```

Scriptet genskaber `dtims.pms_analysis_aadt_null_trafiktal`. Selve `UPDATE` af `pms_analysis.aadt` køres manuelt (som ovenfor), så man ikke utilsigtet overskriver eksisterende AADT.

## Relaterede objekter

| Objekt | Rolle |
|---|---|
| `dtims.pms_analysis` | Måltabel; `aadt` opdateret for 51 gid |
| `dtims.pms_analysis_aadt_null` | Snapshot af oprindelige null-rækker (nu delvist udfyldt) |
| `dtims.pms_analysis_aadt_null_trafiktal` | Matchresultat fra linear search |
| `_fetch_trafiktal_null_aadt.py` | Hent + match |
