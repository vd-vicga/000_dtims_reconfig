"""Indlæs den opdaterede slidlagsliste, stedfæst den på base og pms_analysis og regn overlap.

Rækkefølge:
1. De gamle tabeller (nye_belaegninger_2026*, pms_analysis_belaegning_overlap) slettes,
   og listen i mappen nye_slidlag indlæses i dtims.nye_slidlag_2026.
   1b. Identiske højre/venstre-par på veje hvor pms kun har lane 0 lægges sammen til én
   midt-række (lane 0).
2. Geometri hentes fra Vejdirektoratets linear search (kilometrering fra listen) og
   skrives på rækkerne.
3. Hver række stedfæstes på dtims.base, og kommer som hændelse på base og på
   dtims.pms_analysis. Kilometrering og measure blandes ikke.
4. Overlap pr. pms-sektion beregnes i dtims.pms_analysis_slidlag_overlap.
"""

import collections
import re
from datetime import date, datetime
from pathlib import Path

import belaegning_2026 as b
from openpyxl import load_workbook
from psycopg2.extras import execute_values

ROOT = Path(__file__).resolve().parent
SOURCE_DIR = ROOT / "nye_slidlag"
FRESH_CACHE = ROOT / "_linear_cache" / "slidlag"

PARENT = "dtims.nye_slidlag_2026"
BASE_EVENTS = "dtims.nye_slidlag_2026_base"
PMS_EVENTS = "dtims.nye_slidlag_2026_pms"
OVERLAP = "dtims.pms_analysis_slidlag_overlap"
OLD_TABLES = (
    "dtims.nye_belaegninger_2026_pms",
    "dtims.nye_belaegninger_2026_base",
    "dtims.nye_belaegninger_2026",
    "dtims.pms_analysis_belaegning_overlap",
)

REQUIRED_HEADERS = (
    "Bestyrer", "Admvejnr", "Vejdel", "Frakmt", "Tilkmt", "Vejnavn", "VejStiKlasse",
    "Distrikt", "vejside", "Vejside", "tværprofilelement", "Tværprofilelement",
    "belægningstype", "Belægningstype", "Tykkelse", "Udlægningsdato", "entreprenør",
    "Entreprenør", "øverste lag", "Øverste lag", "bemærkning", "Bruger", "Senest rettet",
)
MERGE_EXTRA = {("450-0", 140654.0, 140890.0)}
SIDE_CODES = {"Højre": "H", "Venstre": "V", "Midt": "M"}


def to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        return datetime.fromisoformat(value.strip()[:10]).date()
    return None


def text(value):
    if value is None:
        return None
    cleaned = str(value).strip()
    return cleaned or None


def find_workbook():
    files = sorted(
        path for path in SOURCE_DIR.glob("*.xlsx") if not path.name.startswith("~$")
    )
    if len(files) != 1:
        raise RuntimeError(f"Forventede én xlsx-fil i {SOURCE_DIR}, fandt {len(files)}.")
    return files[0]


def read_rows():
    path = find_workbook()
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["Sheet0"] if "Sheet0" in workbook.sheetnames else workbook.worksheets[0]
    index = None
    rows = []
    skipped = collections.Counter()
    for number, cells in enumerate(sheet.iter_rows(values_only=True), start=1):
        if index is None:
            if cells and cells[0] == "Bestyrer":
                index = {}
                for position, name in enumerate(cells):
                    name = None if name is None else str(name).strip()
                    if name and name not in index:
                        index[name] = position
                missing = [name for name in REQUIRED_HEADERS if name not in index]
                if missing:
                    raise RuntimeError("Mangler kolonner: " + ", ".join(missing))
            continue
        if not cells or cells[index["Admvejnr"]] is None:
            continue

        def cell(name):
            position = index[name]
            return cells[position] if position < len(cells) else None

        vejdel = str(cell("Vejdel")).strip()
        if not re.fullmatch(r"[A-Za-z0-9]+", vejdel):
            raise RuntimeError(f"Uventet vejdel {vejdel!r} i række {number}.")
        if vejdel not in b.KEEP_VEJDEL:
            skipped[vejdel] += 1
            continue
        side_name = str(cell("Vejside")).strip()
        side = b.SIDES.get(side_name)
        if side is None:
            raise RuntimeError(f"Ukendt vejside {side_name!r} i række {number}.")
        code = str(cell("vejside")).strip()
        if code != SIDE_CODES[side_name]:
            raise RuntimeError(f"Vejside {side_name!r} og kode {code!r} passer ikke i række {number}.")
        admvejnr = int(cell("Admvejnr"))
        rows.append(
            {
                "excel_row": number,
                "bestyrer": int(cell("Bestyrer")),
                "admvejnr": admvejnr,
                "admvejdel": vejdel,
                "frakmt": float(cell("Frakmt")),
                "tilkmt": float(cell("Tilkmt")),
                "vejnavn": text(cell("Vejnavn")),
                "vejstiklasse": text(cell("VejStiKlasse")),
                "distrikt": text(cell("Distrikt")),
                "vejside": side_name,
                "vejside_kode": code,
                "tpe_vejside": side["tpe"],
                "tvaerprofilelement": text(cell("Tværprofilelement")),
                "belaegningstype": text(cell("Belægningstype")),
                "belaegningstype_kode": text(cell("belægningstype")),
                "tykkelse": None if cell("Tykkelse") is None else int(cell("Tykkelse")),
                "udlaegningsdato": to_date(cell("Udlægningsdato")),
                "entreprenoer_nr": text(cell("entreprenør")),
                "entreprenoer": text(cell("Entreprenør")),
                "oeverste_lag": text(cell("Øverste lag")),
                "bemaerkning": text(cell("bemærkning")),
                "bruger": text(cell("Bruger")),
                "senest_rettet": to_date(cell("Senest rettet")),
                "roadname": f"{admvejnr}-{vejdel}",
                "lane": side["lane"],
                "side_code": side["code"],
            }
        )
    workbook.close()
    print(f"Læste {len(rows)} rækker fra {path.name}", flush=True)
    if skipped:
        detail = ", ".join(f"{key}: {count}" for key, count in sorted(skipped.items()))
        print(f"Sprang {sum(skipped.values())} rækker over (vejdel uden for 0, 1, 2): {detail}", flush=True)
    return rows


def step_load(connection, rows):
    with connection:
        with connection.cursor() as cursor:
            for table in OLD_TABLES:
                cursor.execute(f"DROP TABLE IF EXISTS {table}")
            cursor.execute(f"DROP TABLE IF EXISTS {PMS_EVENTS}")
            cursor.execute(f"DROP TABLE IF EXISTS {BASE_EVENTS}")
            cursor.execute(f"DROP TABLE IF EXISTS {PARENT}")
            cursor.execute(
                f"""
                CREATE TABLE {PARENT} (
                    excel_row integer PRIMARY KEY,
                    bestyrer smallint,
                    admvejnr integer NOT NULL,
                    admvejdel text NOT NULL,
                    frakmt double precision NOT NULL,
                    tilkmt double precision NOT NULL,
                    vejnavn text,
                    vejstiklasse text,
                    distrikt text,
                    vejside text NOT NULL,
                    vejside_kode text NOT NULL,
                    tpe_vejside smallint NOT NULL,
                    tvaerprofilelement text,
                    belaegningstype text,
                    belaegningstype_kode text,
                    tykkelse integer,
                    udlaegningsdato date,
                    entreprenoer_nr text,
                    entreprenoer text,
                    oeverste_lag text,
                    bemaerkning text,
                    bruger text,
                    senest_rettet date,
                    roadname text NOT NULL,
                    lane smallint NOT NULL,
                    side_code text NOT NULL,
                    vejside_original text,
                    sammenlagt_excel_rows text,
                    geom geometry(Geometry, 25832),
                    base_id text,
                    base_name text,
                    chain_from double precision,
                    chain_to double precision,
                    measure_from double precision,
                    measure_to double precision,
                    match_distance_m double precision,
                    base_element_count integer NOT NULL DEFAULT 0,
                    pms_section_count integer NOT NULL DEFAULT 0,
                    pms_lane smallint NOT NULL,
                    pms_match text
                )
                """
            )
            execute_values(
                cursor,
                f"""
                INSERT INTO {PARENT} (
                    excel_row, bestyrer, admvejnr, admvejdel, frakmt, tilkmt, vejnavn,
                    vejstiklasse, distrikt, vejside, vejside_kode, tpe_vejside,
                    tvaerprofilelement, belaegningstype, belaegningstype_kode, tykkelse,
                    udlaegningsdato, entreprenoer_nr, entreprenoer, oeverste_lag,
                    bemaerkning, bruger, senest_rettet, roadname, lane, side_code, pms_lane
                ) VALUES %s
                """,
                [
                    (
                        row["excel_row"], row["bestyrer"], row["admvejnr"], row["admvejdel"],
                        row["frakmt"], row["tilkmt"], row["vejnavn"], row["vejstiklasse"],
                        row["distrikt"], row["vejside"], row["vejside_kode"],
                        row["tpe_vejside"], row["tvaerprofilelement"], row["belaegningstype"],
                        row["belaegningstype_kode"], row["tykkelse"], row["udlaegningsdato"],
                        row["entreprenoer_nr"], row["entreprenoer"], row["oeverste_lag"],
                        row["bemaerkning"], row["bruger"], row["senest_rettet"],
                        row["roadname"], row["lane"], row["side_code"], row["lane"],
                    )
                    for row in rows
                ],
                page_size=200,
            )
            cursor.execute(f"SELECT count(*) FROM {PARENT}")
            loaded = cursor.fetchone()[0]
    print(f"Trin 1: gamle tabeller slettet, {loaded} rækker indlæst i {PARENT}", flush=True)


def step_merge(connection, rows):
    """Slår identiske højre- og venstre-slidlag sammen til ét midt-slidlag.

    Gælder veje hvor pms_analysis udelukkende har lane 0, samt de strækninger der er
    nævnt i MERGE_EXTRA. To rækker regnes for ens, når vej, frakmt, tilkmt,
    belægningstype og udlægningsdato er identiske.
    """
    with connection.cursor() as cursor:
        cursor.execute(
            """
            SELECT roadname, array_agg(DISTINCT lane::int ORDER BY lane::int)
            FROM dtims.pms_analysis GROUP BY roadname
            """
        )
        pms_lanes = {name: lanes for name, lanes in cursor.fetchall()}

    groups = collections.defaultdict(lambda: {"H": [], "V": []})
    candidates = 0
    for row in rows:
        only_middle = pms_lanes.get(row["roadname"]) == [0]
        listed = (row["roadname"], row["frakmt"], row["tilkmt"]) in MERGE_EXTRA
        if row["vejside_kode"] in ("H", "V") and (only_middle or listed):
            candidates += 1
            key = (
                row["roadname"], row["frakmt"], row["tilkmt"],
                row["belaegningstype"], row["udlaegningsdato"],
            )
            groups[key][row["vejside_kode"]].append(row)

    dropped = set()
    kept = []
    attribute_differences = 0
    for sides in groups.values():
        if len(sides["H"]) != 1 or len(sides["V"]) != 1:
            continue
        first, second = sorted((sides["H"][0], sides["V"][0]), key=lambda r: r["excel_row"])
        for name in ("tykkelse", "entreprenoer", "tvaerprofilelement", "oeverste_lag"):
            if first[name] != second[name]:
                attribute_differences += 1
        first.update(
            vejside="Midt", vejside_kode="M", tpe_vejside=0, lane=0, side_code="B",
            vejside_original="Højre og venstre",
            sammenlagt_excel_rows=f"{first['excel_row']},{second['excel_row']}",
        )
        kept.append(first)
        dropped.add(second["excel_row"])

    with connection:
        with connection.cursor() as cursor:
            for row in kept:
                cursor.execute(
                    f"""
                    UPDATE {PARENT} SET vejside = %s, vejside_kode = %s, tpe_vejside = %s,
                        lane = %s, pms_lane = %s, side_code = %s,
                        vejside_original = %s, sammenlagt_excel_rows = %s
                    WHERE excel_row = %s
                    """,
                    (
                        row["vejside"], row["vejside_kode"], row["tpe_vejside"], row["lane"],
                        row["lane"], row["side_code"], row["vejside_original"],
                        row["sammenlagt_excel_rows"], row["excel_row"],
                    ),
                )
            if dropped:
                cursor.execute(
                    f"DELETE FROM {PARENT} WHERE excel_row = ANY(%s)", (sorted(dropped),)
                )
            cursor.execute(f"SELECT count(*) FROM {PARENT}")
            remaining = cursor.fetchone()[0]
    print(
        f"Trin 1b: {len(kept)} identiske højre/venstre-par lagt sammen til midt "
        f"({len(dropped)} rækker fjernet, {candidates - 2 * len(kept)} H/V-rækker uparret, "
        f"{attribute_differences} par med forskel i tykkelse/entreprenør m.m.). "
        f"{remaining} rækker i {PARENT}.",
        flush=True,
    )
    return [row for row in rows if row["excel_row"] not in dropped]


def step_geometry(connection, rows):
    b.CACHE = FRESH_CACHE
    network = b.load_network(rows)
    records = b.build_records(rows, network)
    with connection:
        with connection.cursor() as cursor:
            for record in records:
                if record["geom"] is None:
                    continue
                cursor.execute(
                    f"""
                    UPDATE {PARENT}
                    SET geom = ST_Multi(ST_CollectionExtract(
                        ST_MakeValid(ST_SetSRID(ST_GeomFromText(%s), 25832)), 2))
                    WHERE excel_row = %s
                    """,
                    (record["geom"].wkt, record["excel_row"]),
                )
            cursor.execute(f"SELECT count(*), count(geom) FROM {PARENT}")
            total, with_geom = cursor.fetchone()
    print(f"Trin 2: geometri fra linear search på {with_geom} af {total} rækker", flush=True)
    return records


def create_event_tables(cursor):
    cursor.execute(
        f"""
        CREATE TABLE {BASE_EVENTS} (
            id integer GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            excel_row integer NOT NULL REFERENCES {PARENT} (excel_row),
            base_id text NOT NULL,
            base_name text,
            roadname text NOT NULL,
            lane smallint NOT NULL,
            "from" double precision NOT NULL,
            "to" double precision NOT NULL,
            measure_from double precision NOT NULL,
            measure_to double precision NOT NULL,
            length_m double precision NOT NULL,
            belaegningstype text,
            udlaegningsdato date,
            vejside text NOT NULL,
            admvejnr integer NOT NULL,
            admvejdel text NOT NULL,
            geom geometry(Geometry, 25832),
            CHECK ("to" > "from")
        );
        CREATE TABLE {PMS_EVENTS} (
            id integer GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            excel_row integer NOT NULL REFERENCES {PARENT} (excel_row),
            base_id text NOT NULL,
            pms_gid bigint,
            name text NOT NULL,
            roadname text NOT NULL,
            lane smallint NOT NULL,
            "from" double precision NOT NULL,
            "to" double precision NOT NULL,
            measure_from double precision NOT NULL,
            measure_to double precision NOT NULL,
            length_m double precision NOT NULL,
            belaegningstype text,
            udlaegningsdato date,
            vejside text NOT NULL,
            admvejnr integer NOT NULL,
            admvejdel text NOT NULL,
            geom geometry(Geometry, 25832),
            CHECK ("to" > "from")
        );
        CREATE INDEX nye_slidlag_2026_geom_idx ON {PARENT} USING gist (geom);
        CREATE INDEX nye_slidlag_2026_base_geom_idx ON {BASE_EVENTS} USING gist (geom);
        CREATE INDEX nye_slidlag_2026_pms_geom_idx ON {PMS_EVENTS} USING gist (geom);
        CREATE INDEX nye_slidlag_2026_base_idx
            ON {BASE_EVENTS} (base_id, measure_from, measure_to);
        CREATE INDEX nye_slidlag_2026_pms_name_idx ON {PMS_EVENTS} (name, "from", "to");
        COMMENT ON TABLE {PARENT} IS
            'Nye slidlag 2026 (Vejman-udtræk 12. november). Én række pr. udlægning. geom er linear search-linjen ud fra kilometrering (frakmt/tilkmt). measure_from/measure_to er mål på dtims.base. chain_from/chain_to er samme stationering som base.from/to og pms_analysis.from/to. lane er vejsidens lane, pms_lane den lane der blev brugt mod pms_analysis, pms_match er vejside eller vejdel. Identiske højre/venstre-par på veje med kun pms-lane 0 er lagt sammen til én midt-række (lane 0); vejside_original og sammenlagt_excel_rows viser det.';
        COMMENT ON TABLE {BASE_EVENTS} IS
            'Slidlaget som hændelser på dtims.base. from/to er base-stationeringen. measure_from/measure_to er mål på base-elementet. geom er base-linjens udsnit.';
        COMMENT ON TABLE {PMS_EVENTS} IS
            'Slidlaget delt op på dtims.pms_analysis-sektionerne. from/to og measure_from/measure_to er klippet til sektionen. lane er pms-lanen. geom er base-linjens udsnit.';
        """
    )


def insert_events(cursor, records):
    base_rows = []
    pms_rows = []
    for record in records:
        for event in record["base_events"]:
            base_rows.append(
                (
                    record["excel_row"], event["base_id"], event["base_name"],
                    record["roadname"], record["lane"], event["from"], event["to"],
                    event["measure_from"], event["measure_to"], event["to"] - event["from"],
                    record["belaegningstype"], record["udlaegningsdato"], record["vejside"],
                    record["admvejnr"], record["admvejdel"],
                    None if event["geom"] is None else event["geom"].wkt,
                )
            )
        for event in record["pms_events"]:
            pms_rows.append(
                (
                    record["excel_row"], event["base_id"],
                    int(event["gid"]) if event["gid"] is not None else None,
                    event["name"], record["roadname"], record["pms_lane"],
                    event["from"], event["to"], event["measure_from"], event["measure_to"],
                    event["to"] - event["from"], record["belaegningstype"],
                    record["udlaegningsdato"], record["vejside"], record["admvejnr"],
                    record["admvejdel"],
                    None if event["geom"] is None else event["geom"].wkt,
                )
            )
    geom_sql = "ST_Multi(ST_CollectionExtract(ST_MakeValid(ST_SetSRID(ST_GeomFromText(%s), 25832)), 2))"
    if base_rows:
        execute_values(
            cursor,
            f"""
            INSERT INTO {BASE_EVENTS} (
                excel_row, base_id, base_name, roadname, lane, "from", "to",
                measure_from, measure_to, length_m, belaegningstype,
                udlaegningsdato, vejside, admvejnr, admvejdel, geom
            ) VALUES %s
            """,
            base_rows,
            template="(" + ", ".join(["%s"] * 15) + ", " + geom_sql + ")",
            page_size=100,
        )
    if pms_rows:
        execute_values(
            cursor,
            f"""
            INSERT INTO {PMS_EVENTS} (
                excel_row, base_id, pms_gid, name, roadname, lane, "from", "to",
                measure_from, measure_to, length_m, belaegningstype,
                udlaegningsdato, vejside, admvejnr, admvejdel, geom
            ) VALUES %s
            """,
            pms_rows,
            template="(" + ", ".join(["%s"] * 16) + ", " + geom_sql + ")",
            page_size=100,
        )


def update_parent(cursor, records):
    values = [
        (
            record["excel_row"], record["base_id"], record["base_name"],
            record["chain_from"], record["chain_to"], record["measure_from"],
            record["measure_to"], record["match_distance_m"],
            len(record["base_events"]), len(record["pms_events"]),
            record["pms_lane"], record["pms_match"],
        )
        for record in records
    ]
    execute_values(
        cursor,
        f"""
        UPDATE {PARENT} AS p SET
            base_id = v.base_id::text,
            base_name = v.base_name::text,
            chain_from = v.chain_from::double precision,
            chain_to = v.chain_to::double precision,
            measure_from = v.measure_from::double precision,
            measure_to = v.measure_to::double precision,
            match_distance_m = v.match_distance_m::double precision,
            base_element_count = v.base_element_count::integer,
            pms_section_count = v.pms_section_count::integer,
            pms_lane = v.pms_lane::smallint,
            pms_match = v.pms_match::text
        FROM (VALUES %s) AS v (
            excel_row, base_id, base_name, chain_from, chain_to, measure_from,
            measure_to, match_distance_m, base_element_count, pms_section_count,
            pms_lane, pms_match
        )
        WHERE p.excel_row = v.excel_row::integer
        """,
        values,
        page_size=100,
    )


def step_reference(connection, records):
    with connection:
        with connection.cursor() as cursor:
            base_features = b.load_base(cursor)
            sections = b.load_sections(cursor)
            print(
                f"Læste {len(base_features)} base-elementer og "
                f"{sum(len(items) for items in sections.values())} sektioner",
                flush=True,
            )
            b.assign_measures(records, base_features, sections, True)
            create_event_tables(cursor)
            insert_events(cursor, records)
            update_parent(cursor, records)
            cursor.execute(
                f"""
                SELECT count(*), count(measure_from),
                       count(*) FILTER (WHERE pms_section_count > 0),
                       count(*) FILTER (WHERE pms_match = 'vejdel')
                FROM {PARENT}
                """
            )
            parents = cursor.fetchone()
            cursor.execute(f"SELECT count(*) FROM {BASE_EVENTS}")
            base_events = cursor.fetchone()[0]
            cursor.execute(f"SELECT count(*) FROM {PMS_EVENTS}")
            pms_events = cursor.fetchone()[0]
    print(
        f"Trin 3: {parents[1]} af {parents[0]} stedfæstet på base, "
        f"{parents[2]} har pms-sektioner ({parents[3]} via vejdel). "
        f"Hændelser: {base_events} på base, {pms_events} på pms_analysis.",
        flush=True,
    )


OVERLAP_SQL = f"""
DROP TABLE IF EXISTS {OVERLAP};
CREATE TABLE {OVERLAP} (
    gid integer PRIMARY KEY REFERENCES dtims.pms_analysis (gid),
    name text NOT NULL,
    roadname text,
    lane double precision,
    section_length_m double precision NOT NULL,
    overlap_m double precision NOT NULL,
    overlap_pct double precision NOT NULL,
    belaegning_count integer NOT NULL
);
INSERT INTO {OVERLAP} (
    gid, name, roadname, lane, section_length_m, overlap_m, overlap_pct, belaegning_count
)
SELECT
    s.gid, s.name, s.roadname, s.lane,
    s."to" - s."from",
    LEAST(s."to" - s."from", COALESCE(hit.overlap_m, 0)),
    CASE
        WHEN s."to" = s."from" THEN 0
        ELSE round((
            100 * LEAST(s."to" - s."from", COALESCE(hit.overlap_m, 0))
            / (s."to" - s."from")
        )::numeric, 2)::double precision
    END,
    COALESCE(hit.belaegning_count, 0)
FROM dtims.pms_analysis AS s
LEFT JOIN LATERAL (
    SELECT
        count(p.excel_row)::integer AS belaegning_count,
        ST_Length(
            ST_Intersection(
                ST_Transform(s.geom, 25832),
                ST_Union(ST_Buffer(p.geom, 15))
            )
        ) AS overlap_m
    FROM {PARENT} AS p
    WHERE p.roadname = s.roadname
      AND p.pms_lane::double precision = s.lane
      AND p.geom IS NOT NULL
      AND ST_DWithin(p.geom, ST_Transform(s.geom, 25832), 15)
) AS hit ON true;
CREATE INDEX pms_analysis_slidlag_overlap_name_idx ON {OVERLAP} (name);
COMMENT ON TABLE {OVERLAP} IS
    'Andel af hver pms_analysis-sektion der ligger inden for 15 m af et nyt slidlag med samme vej og pms-lane (nye_slidlag_2026.pms_lane). overlap_pct er 0-100 af sektionens stationering. Slidlag oven i hinanden tælles én gang.';
"""


def step_overlap(connection):
    with connection:
        with connection.cursor() as cursor:
            cursor.execute(OVERLAP_SQL)
            cursor.execute(
                f"""
                SELECT count(*), count(*) FILTER (WHERE overlap_pct > 0),
                       count(*) FILTER (WHERE overlap_pct = 100),
                       count(*) FILTER (WHERE overlap_pct > 100),
                       round(max(overlap_pct)::numeric, 2)
                FROM {OVERLAP}
                """
            )
            summary = cursor.fetchone()
    print(
        f"Trin 4: overlap beregnet for {summary[0]} sektioner. "
        f"{summary[1]} har overlap, {summary[2]} har 100 %, {summary[3]} over 100 %, "
        f"max {summary[4]} %.",
        flush=True,
    )


def main():
    rows = read_rows()
    connection = b.connect()
    try:
        step_load(connection, rows)
        rows = step_merge(connection, rows)
        records = step_geometry(connection, rows)
        step_reference(connection, records)
        step_overlap(connection)
    finally:
        connection.close()


if __name__ == "__main__":
    main()
