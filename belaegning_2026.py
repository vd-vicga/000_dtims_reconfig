"""Hent geometri for nye kørebane-belægninger og gem dem som hændelser.

1. Rækkerne i regnearket slås op i Vejdirektoratets linear search med vejnummer,
   vejdel, vejside og fra-/til-kilometrering. Kilometreringen bruges kun til at
   hente geometrien. Geometrien skrives tilbage i regnearket.
2. Hver belægning stedfæstes på dtims.base ved geometri. Målene (measure) læses
   fra base.frommeasure/tomeasure. Kilometrering og measure er to forskellige
   referencesystemer, så kilometrering bruges ikke til at finde mål.
3. Belægningen lægges som hændelse på base (én række pr. base-element) og deles
   op på dtims.pms_analysis-sektionerne, den dækker. Geometrien for hændelserne
   er klippet ud af base-linjen.
"""

import json
import re
import site
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

sys.path = [path for path in sys.path if path != site.getusersitepackages()]

import psycopg2
import requests
from openpyxl import load_workbook
from psycopg2.extras import execute_values
from shapely import from_wkt, to_wkt
from shapely.geometry import LineString, MultiLineString, Point
from shapely.ops import linemerge, substring

URL = "https://linearsearch.vd.dk/linear-search-api/v4/execute"
ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "Nye belægninger på kørebaner i 2026.xlsx"
CACHE = ROOT / "_linear_cache"
EXCEL_WKT_LIMIT = 32000
MATCH_DISTANCE_M = 15.0

SIDES = {
    "Højre": {"tpe": 1, "lane": 1, "code": "R"},
    "Venstre": {"tpe": -1, "lane": 2, "code": "L"},
    "Midt": {"tpe": 0, "lane": 0, "code": "B"},
}
HEADER_ROW_MARK = "ADMVEJNR"
EXTRA_HEADERS = (
    "GEOMETRY",
    "ROADNAME",
    "lane",
    "BASE_ID",
    "BASE_NAME",
    "MEASURE_FROM",
    "MEASURE_TO",
    "PMS_SEKTIONER",
)


def distance(a, b):
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5


def clean_coords(coords):
    points = []
    for coord in coords:
        point = (float(coord[0]), float(coord[1]))
        if points and distance(points[-1], point) < 0.01:
            continue
        points.append(point)
    return points


def geometry_parts(geometry):
    kind = geometry.get("type")
    coords = geometry.get("coordinates") or []
    if kind == "LineString":
        return [coords]
    if kind == "MultiLineString":
        return list(coords)
    return []


def clip_line(coords, measure_from, measure_to, want_from, want_to):
    start = max(want_from, measure_from)
    end = min(want_to, measure_to)
    span = measure_to - measure_from
    if end - start <= 0.05 or span <= 0:
        return None
    points = clean_coords(coords)
    if len(points) < 2:
        return None
    line = LineString(points)
    if line.length <= 0:
        return None
    part = substring(
        line,
        (start - measure_from) / span * line.length,
        (end - measure_from) / span * line.length,
    )
    if part.is_empty or part.geom_type != "LineString" or len(part.coords) < 2:
        return None
    return part


def stitch(lines):
    if not lines:
        return None
    chunks = [list(lines[0].coords)]
    for line in lines[1:]:
        coords = list(line.coords)
        if distance(chunks[-1][-1], coords[0]) <= 0.5:
            chunks[-1].extend(coords[1:])
        elif distance(chunks[-1][-1], coords[-1]) <= 0.5:
            chunks[-1].extend(list(reversed(coords))[1:])
        else:
            chunks.append(coords)
    geoms = [LineString(chunk) for chunk in chunks if len(chunk) >= 2]
    if not geoms:
        return None
    if len(geoms) == 1:
        return geoms[0]
    return MultiLineString(geoms)


def dedupe_segments(features):
    unique = []
    seen = set()
    for feature in features:
        props = feature.get("properties") or {}
        measure_from = props.get("Frakmt")
        measure_to = props.get("Tilkmt")
        if measure_from is None or measure_to is None:
            continue
        for part in geometry_parts(feature.get("geometry") or {}):
            coords = clean_coords(part)
            if len(coords) < 2:
                continue
            key = (
                round(float(measure_from), 3),
                round(float(measure_to), 3),
                tuple((round(x, 2), round(y, 2)) for x, y in coords),
            )
            if key in seen:
                continue
            seen.add(key)
            unique.append(
                {
                    "frakmt": float(measure_from),
                    "tilkmt": float(measure_to),
                    "coords": coords,
                    "line": LineString(coords),
                }
            )

    unique.sort(key=lambda item: (item["frakmt"], -(item["tilkmt"] - item["frakmt"])))
    kept = []
    for segment in unique:
        drop = False
        for other in kept:
            overlap = min(segment["tilkmt"], other["tilkmt"]) - max(
                segment["frakmt"], other["frakmt"]
            )
            if overlap > 1 and segment["line"].hausdorff_distance(other["line"]) < 2:
                drop = True
                break
        if not drop:
            kept.append(segment)
    return kept


def geometry_for_interval(segments, want_from, want_to):
    ordered = sorted(
        (
            segment
            for segment in segments
            if segment["tilkmt"] > want_from + 0.05
            and segment["frakmt"] < want_to - 0.05
        ),
        key=lambda segment: (segment["frakmt"], segment["tilkmt"]),
    )
    lines = []
    previous_end = None
    for segment in ordered:
        coords = list(segment["coords"])
        if previous_end is not None:
            start_distance = distance(coords[0], previous_end)
            end_distance = distance(coords[-1], previous_end)
            if end_distance < start_distance and end_distance < 1:
                coords = list(reversed(coords))
        previous_end = coords[-1]
        part = clip_line(
            coords, segment["frakmt"], segment["tilkmt"], want_from, want_to
        )
        if part is not None:
            lines.append(part)
    return stitch(lines)


def excel_wkt(geom):
    if geom is None:
        return None, False
    text = to_wkt(geom, rounding_precision=2)
    if len(text) <= EXCEL_WKT_LIMIT:
        return text, False
    simplified = geom
    tolerance = 0.5
    while tolerance <= 50 and len(text) > EXCEL_WKT_LIMIT:
        simplified = simplified.simplify(tolerance, preserve_topology=True)
        text = to_wkt(simplified, rounding_precision=2)
        tolerance *= 2
    if len(text) > EXCEL_WKT_LIMIT:
        return None, True
    return text, True


def read_rows():
    workbook = load_workbook(WORKBOOK)
    sheet = workbook.active
    header_row = None
    for cells in sheet.iter_rows(min_row=1, max_row=15):
        if cells[0].value == HEADER_ROW_MARK:
            header_row = cells[0].row
            break
    if header_row is None:
        raise RuntimeError("Fandt ikke kolonneoverskriften ADMVEJNR.")

    rows = []
    for cells in sheet.iter_rows(min_row=header_row + 1):
        if cells[0].value is None:
            continue
        side_name = str(cells[6].value or "").strip()
        side = SIDES.get(side_name)
        if side is None:
            raise RuntimeError(f"Ukendt vejside {side_name!r} i række {cells[0].row}.")
        laid = cells[5].value
        if isinstance(laid, datetime):
            laid = laid.date()
        elif isinstance(laid, str) and laid:
            laid = datetime.fromisoformat(laid[:10]).date()
        elif not isinstance(laid, date):
            laid = None
        vejdel = str(cells[1].value).strip()
        if not re.fullmatch(r"[A-Za-z0-9]+", vejdel):
            raise RuntimeError(f"Uventet vejdel {vejdel!r} i række {cells[0].row}.")
        admvejnr = int(cells[0].value)
        rows.append(
            {
                "excel_row": cells[0].row,
                "admvejnr": admvejnr,
                "admvejdel": vejdel,
                "frakmt": float(cells[2].value),
                "tilkmt": float(cells[3].value),
                "belaegningstype": cells[4].value,
                "udlaegningsdato": laid,
                "vejside": side_name,
                "tpe_vejside": side["tpe"],
                "lane": side["lane"],
                "side_code": side["code"],
                "plade_tpe": cells[7].value,
                "oeverste_lag": cells[8].value,
                "roadname": f"{admvejnr}-{vejdel}",
            }
        )
    return workbook, sheet, header_row, rows


def fetch_segments(vejnr, vejdel, tpe):
    cache = CACHE / f"{vejnr}_{vejdel}_{tpe}.json"
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))

    select = (
        "vm:admvejnr, vm:admvejdel, vm:TPE_vejside, vm:frakmt, vm:tilkmt, geometry"
    )
    queries = [
        (
            f"{select}\nwhere\nvm:admvejnr = '{vejnr}' AND vm:admvejdel = '{vejdel}' "
            f"AND vm:TPE_vejside = {tpe} AND Bestyrer = 0"
        ),
        (
            f"{select}\nwhere\nvm:admvejnr = '{vejnr}' AND vm:admvejdel = '{vejdel}' "
            f"AND vm:TPE_vejside = {tpe}"
        ),
    ]
    features = []
    for query in queries:
        response = None
        for attempt in range(3):
            try:
                response = requests.get(URL, params={"query": query}, timeout=180)
                response.raise_for_status()
                break
            except requests.RequestException:
                if attempt == 2:
                    raise
                time.sleep(2 * (attempt + 1))
        features = response.json().get("features") or []
        if features:
            break
    CACHE.mkdir(exist_ok=True)
    cache.write_text(json.dumps({"features": features}), encoding="utf-8")
    return {"features": features}


def load_network(rows):
    keys = sorted({(row["admvejnr"], row["admvejdel"], row["tpe_vejside"]) for row in rows})
    network = {}
    print(f"Henter geometri for {len(keys)} vejside-strækninger", flush=True)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {
            pool.submit(fetch_segments, vejnr, vejdel, tpe): (vejnr, vejdel, tpe)
            for vejnr, vejdel, tpe in keys
        }
        done = 0
        for future in as_completed(futures):
            key = futures[future]
            payload = future.result()
            network[key] = dedupe_segments(payload.get("features") or [])
            done += 1
            if done % 50 == 0 or done == len(keys):
                print(f"  {done}/{len(keys)}", flush=True)
    return network


def build_records(rows, network):
    records = []
    missing = []
    simplified = 0
    for row in rows:
        segments = network[(row["admvejnr"], row["admvejdel"], row["tpe_vejside"])]
        geom = geometry_for_interval(segments, row["frakmt"], row["tilkmt"])
        record = dict(row)
        record["geom"] = geom
        record["wkt"], was_simplified = excel_wkt(geom)
        simplified += int(was_simplified)
        record["simplified"] = was_simplified
        if geom is None:
            missing.append(
                f"{row['roadname']} {row['vejside']} {row['frakmt']:.0f}-{row['tilkmt']:.0f}"
            )
        records.append(record)
    with_geom = sum(record["geom"] is not None for record in records)
    print(f"Geometri på {with_geom} af {len(records)} rækker", flush=True)
    if simplified:
        print(f"{simplified} geometrier er forenklet i Excel, fulde geometrier gemmes i databasen.")
    if missing:
        print(f"Uden geometri: {len(missing)}")
        for item in missing[:20]:
            print(" ", item)
    return records


def ensure_column(sheet, header_row, title):
    for cell in sheet[header_row]:
        if cell.value == title:
            return cell.column
    column = sheet.max_column + 1
    sheet.cell(header_row, column, title)
    return column


def write_excel(workbook, sheet, header_row, records):
    columns = {
        title: ensure_column(sheet, header_row, title) for title in EXTRA_HEADERS
    }
    by_row = {record["excel_row"]: record for record in records}
    for cells in sheet.iter_rows(min_row=header_row + 1):
        record = by_row.get(cells[0].row)
        if record is None:
            continue
        values = {
            "GEOMETRY": record.get("wkt"),
            "ROADNAME": record["roadname"],
            "lane": record["lane"],
            "BASE_ID": record.get("base_id"),
            "BASE_NAME": record.get("base_name"),
            "MEASURE_FROM": record.get("measure_from"),
            "MEASURE_TO": record.get("measure_to"),
            "PMS_SEKTIONER": record.get("pms_sections"),
        }
        for title, value in values.items():
            sheet.cell(cells[0].row, columns[title], value)
    try:
        workbook.save(WORKBOOK)
        print(f"Skrev i {WORKBOOK.name}", flush=True)
    except PermissionError:
        alternative = WORKBOOK.with_name(WORKBOOK.stem + " med geometri.xlsx")
        workbook.save(alternative)
        print(
            f"{WORKBOOK.name} er åben i Excel. Resultatet er gemt i {alternative.name}.",
            flush=True,
        )


def env_settings():
    settings = {}
    for line in (ROOT / "exp.env").read_text(encoding="utf-8").splitlines():
        if not line or line.lstrip().startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        settings[key.strip()] = value.strip().strip('"').strip("'")
    return settings


def connect():
    settings = env_settings()
    return psycopg2.connect(
        host=settings["POSTGRES_HOST"],
        port=settings["POSTGRES_PORT"],
        user=settings["POSTGRES_USER"],
        password=settings["POSTGRES_PASSWORD"],
        dbname=settings["POSTGRES_DB"],
        connect_timeout=15,
        sslmode="disable",
    )


def parse_route(name):
    if name is None:
        return None
    match = re.fullmatch(r"(\d+)-(\d+)(?:_([BLR]))?(?:_\d+)?", str(name).strip())
    if match is None:
        return None
    return f"{match.group(1)}-{match.group(2)}", match.group(3)


def prepare_line(geom):
    if geom is None or geom.is_empty or geom.length <= 0:
        return None
    if geom.geom_type == "LineString":
        return geom
    if geom.geom_type != "MultiLineString":
        return None
    merged = linemerge(geom)
    if merged.is_empty or merged.length <= 0:
        return None
    return merged


def line_parts(geom):
    if geom.geom_type == "LineString":
        return [geom]
    return [part for part in geom.geoms if part.length > 0]


def point_at_fraction(geom, fraction):
    parts = line_parts(geom)
    total = sum(part.length for part in parts)
    if not parts or total <= 0:
        return None
    target = total * min(1.0, max(0.0, fraction))
    walked = 0.0
    for part in parts:
        if walked + part.length >= target or part is parts[-1]:
            return part.interpolate(max(0.0, target - walked))
        walked += part.length
    return Point(parts[-1].coords[-1])


def mean_distance(paving, base_line):
    if paving.length <= 0:
        return paving.distance(base_line)
    steps = max(2, int(paving.length / 200) + 1)
    total = 0.0
    for index in range(steps + 1):
        point = point_at_fraction(paving, index / steps)
        total += point.distance(base_line)
    return total / (steps + 1)


def locate(line, point, measure_from, measure_to):
    """Mål for det punkt på base-linjen, der er tættest på point."""
    parts = line_parts(line)
    total = sum(part.length for part in parts)
    if total <= 0:
        return measure_from
    best_gap = None
    walked = 0.0
    along = 0.0
    for part in parts:
        gap = part.distance(point)
        projected = walked + (part.project(point) if part.length else 0.0)
        if best_gap is None or gap < best_gap:
            best_gap = gap
            along = projected
        walked += part.length
    fraction = min(1.0, max(0.0, along / total))
    return measure_from + fraction * (measure_to - measure_from)


def measures_against(base_line, measure_from, measure_to, paving):
    """Mål på base for den del af belægningen, der ligger langs base-elementet."""
    steps = max(2, int(paving.length / 50) + 1)
    hits = []
    for index in range(steps + 1):
        point = point_at_fraction(paving, index / steps)
        if point is not None and point.distance(base_line) <= MATCH_DISTANCE_M:
            hits.append(point)
    if len(hits) < 2:
        return None
    first = locate(base_line, hits[0], measure_from, measure_to)
    last = locate(base_line, hits[-1], measure_from, measure_to)
    low, high = sorted((first, last))
    floor = min(measure_from, measure_to)
    ceiling = max(measure_from, measure_to)
    low = min(max(low, floor), ceiling)
    high = min(max(high, floor), ceiling)
    if high - low <= 0.5:
        return None
    return low, high


def load_base(cursor):
    cursor.execute(
        """
        SELECT id::text, name, "from", "to", frommeasure, tomeasure,
               ST_AsText(ST_Force2D(ST_Transform(geom, 25832)))
        FROM dtims.base
        WHERE validto IS NULL AND endedon IS NULL
          AND "from" IS NOT NULL AND "to" IS NOT NULL
          AND frommeasure IS NOT NULL AND tomeasure IS NOT NULL
        """
    )
    features = []
    for base_id, name, chain_from, chain_to, meas_from, meas_to, wkt in cursor.fetchall():
        parsed = parse_route(name)
        roadname = parsed[0] if parsed else str(name)
        features.append(
            {
                "id": base_id,
                "name": name,
                "roadname": roadname,
                "chain_from": min(float(chain_from), float(chain_to)),
                "chain_to": max(float(chain_from), float(chain_to)),
                "meas_from": min(float(meas_from), float(meas_to)),
                "meas_to": max(float(meas_from), float(meas_to)),
                "line": prepare_line(from_wkt(wkt)) if wkt else None,
            }
        )
    return features


def load_sections(cursor):
    cursor.execute(
        """
        SELECT gid::text, roadname, name, lane, "from", "to"
        FROM dtims.pms_analysis
        WHERE "from" IS NOT NULL AND "to" IS NOT NULL
        """
    )
    grouped = {}
    for gid, roadname, name, lane, measure_from, measure_to in cursor.fetchall():
        grouped.setdefault((roadname, int(lane)), []).append(
            {
                "gid": gid,
                "name": name,
                "from": float(measure_from),
                "to": float(measure_to),
            }
        )
    return grouped


def to_chain(feature, measure):
    span = feature["meas_to"] - feature["meas_from"]
    if span <= 0:
        return feature["chain_from"]
    return feature["chain_from"] + (measure - feature["meas_from"]) * (
        (feature["chain_to"] - feature["chain_from"]) / span
    )


def base_piece(feature, low_m, high_m):
    """Del af base-linjen mellem to mål."""
    line = feature["line"]
    if line is None or line.geom_type != "LineString" or line.length <= 0:
        return None
    span = feature["meas_to"] - feature["meas_from"]
    if span <= 0:
        return None
    f0 = min(1.0, max(0.0, (low_m - feature["meas_from"]) / span))
    f1 = min(1.0, max(0.0, (high_m - feature["meas_from"]) / span))
    if f1 - f0 <= 0:
        return None
    piece = substring(line, f0 * line.length, f1 * line.length)
    if piece.is_empty or piece.geom_type != "LineString":
        return None
    return piece


def assign_measures(records, base_features, sections, has_geometry):
    grouped = {}
    for feature in base_features:
        grouped.setdefault(feature["roadname"], []).append(feature)

    matched = 0
    for record in records:
        record.update(
            {
                "base_events": [],
                "pms_events": [],
                "base_id": None,
                "base_name": None,
                "measure_from": None,
                "measure_to": None,
                "chain_from": None,
                "chain_to": None,
                "match_distance_m": None,
                "pms_sections": None,
            }
        )
        paving = record["geom"]
        candidates = grouped.get(record["roadname"], [])
        scored = []
        if paving is not None and has_geometry:
            for feature in candidates:
                if feature["line"] is None:
                    continue
                gap = mean_distance(paving, feature["line"])
                if gap <= MATCH_DISTANCE_M:
                    scored.append((gap, feature))
        scored.sort(key=lambda item: item[0])

        accepted = []
        for gap, feature in scored:
            located = measures_against(
                feature["line"], feature["meas_from"], feature["meas_to"], paving
            )
            if located is None:
                continue
            low_m, high_m = located
            low_c, high_c = sorted((to_chain(feature, low_m), to_chain(feature, high_m)))
            overlaps = any(
                min(high_c, item["chain_to"]) - max(low_c, item["chain_from"]) > 1
                for item in accepted
            )
            if overlaps:
                continue
            accepted.append(
                {
                    "gap": gap,
                    "feature": feature,
                    "meas_from": low_m,
                    "meas_to": high_m,
                    "chain_from": low_c,
                    "chain_to": high_c,
                }
            )

        if not accepted:
            continue
        matched += 1
        record["match_distance_m"] = min(item["gap"] for item in accepted)
        record["measure_from"] = min(item["meas_from"] for item in accepted)
        record["measure_to"] = max(item["meas_to"] for item in accepted)
        record["chain_from"] = min(item["chain_from"] for item in accepted)
        record["chain_to"] = max(item["chain_to"] for item in accepted)
        names = sorted({item["feature"]["name"] for item in accepted})
        if len(accepted) == 1:
            record["base_id"] = accepted[0]["feature"]["id"]
        record["base_name"] = ", ".join(names)

        for item in accepted:
            feature = item["feature"]
            record["base_events"].append(
                {
                    "feature": feature,
                    "base_id": feature["id"],
                    "base_name": feature["name"],
                    "from": item["chain_from"],
                    "to": item["chain_to"],
                    "measure_from": item["meas_from"],
                    "measure_to": item["meas_to"],
                    "geom": base_piece(feature, item["meas_from"], item["meas_to"]),
                }
            )

        route_sections = sections.get((record["roadname"], record["lane"]), [])
        pms_names = []
        for event in record["base_events"]:
            chain_span = event["to"] - event["from"]
            measure_span = event["measure_to"] - event["measure_from"]
            scale = measure_span / chain_span if chain_span > 0 else 1.0
            feature = event["feature"]
            for section in route_sections:
                low = max(event["from"], section["from"])
                high = min(event["to"], section["to"])
                if high - low <= 0.5:
                    continue
                low_m = event["measure_from"] + (low - event["from"]) * scale
                high_m = event["measure_from"] + (high - event["from"]) * scale
                pms_names.append(section["name"])
                record["pms_events"].append(
                    {
                        "base_id": event["base_id"],
                        "gid": section["gid"],
                        "name": section["name"],
                        "from": low,
                        "to": high,
                        "measure_from": low_m,
                        "measure_to": high_m,
                        "geom": base_piece(feature, low_m, high_m),
                    }
                )
        record["pms_sections"] = ", ".join(dict.fromkeys(pms_names)) or None

    print(f"Stedfæstet {matched} af {len(records)} belægninger på base", flush=True)
    pms_rows = sum(len(record["pms_events"]) for record in records)
    covered = sum(bool(record["pms_events"]) for record in records)
    print(f"{covered} belægninger dækker {pms_rows} sektioner i pms_analysis", flush=True)


def create_tables(cursor):
    cursor.execute(
        """
        DROP TABLE IF EXISTS dtims.nye_belaegninger_2026_pms;
        DROP TABLE IF EXISTS dtims.nye_belaegninger_2026_base;
        DROP TABLE IF EXISTS dtims.nye_belaegninger_2026;
        """
    )
    cursor.execute(
        """
        CREATE TABLE dtims.nye_belaegninger_2026 (
            excel_row integer PRIMARY KEY,
            admvejnr integer NOT NULL,
            admvejdel text NOT NULL,
            frakmt double precision NOT NULL,
            tilkmt double precision NOT NULL,
            belaegningstype text,
            udlaegningsdato date,
            vejside text NOT NULL,
            tpe_vejside smallint NOT NULL,
            plade_tpe text,
            oeverste_lag text,
            roadname text NOT NULL,
            lane smallint NOT NULL,
            side_code text NOT NULL,
            geom geometry(Geometry, 25832),
            base_id text,
            base_name text,
            chain_from double precision,
            chain_to double precision,
            measure_from double precision,
            measure_to double precision,
            match_distance_m double precision,
            base_element_count integer NOT NULL,
            pms_section_count integer NOT NULL
        );
        CREATE TABLE dtims.nye_belaegninger_2026_base (
            id integer GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            excel_row integer NOT NULL
                REFERENCES dtims.nye_belaegninger_2026 (excel_row),
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
        CREATE TABLE dtims.nye_belaegninger_2026_pms (
            id integer GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            excel_row integer NOT NULL
                REFERENCES dtims.nye_belaegninger_2026 (excel_row),
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
        CREATE INDEX nye_belaegninger_2026_geom_idx
            ON dtims.nye_belaegninger_2026 USING gist (geom);
        CREATE INDEX nye_belaegninger_2026_base_geom_idx
            ON dtims.nye_belaegninger_2026_base USING gist (geom);
        CREATE INDEX nye_belaegninger_2026_pms_geom_idx
            ON dtims.nye_belaegninger_2026_pms USING gist (geom);
        CREATE INDEX nye_belaegninger_2026_base_idx
            ON dtims.nye_belaegninger_2026_base (base_id, measure_from, measure_to);
        CREATE INDEX nye_belaegninger_2026_pms_name_idx
            ON dtims.nye_belaegninger_2026_pms (name, "from", "to");
        COMMENT ON TABLE dtims.nye_belaegninger_2026 IS
            'Nye kørebane-belægninger 2026. Én række pr. udlægning. measure_from/measure_to er mål fra dtims.base.frommeasure/tomeasure. chain_from/chain_to er samme referencesystem som base.from/to og pms_analysis.from/to.';
        COMMENT ON TABLE dtims.nye_belaegninger_2026_base IS
            'Belægningen som hændelser på dtims.base. from/to er base-referencen. measure_from/measure_to er mål på base-elementet. geom er base-linjens udsnit.';
        COMMENT ON TABLE dtims.nye_belaegninger_2026_pms IS
            'Belægningen delt op på dtims.pms_analysis-sektionerne. from/to og measure_from/measure_to er klippet til sektionen. geom er base-linjens udsnit.';
        """
    )


def insert_records(cursor, records):
    parent_rows = [
        (
            record["excel_row"],
            record["admvejnr"],
            record["admvejdel"],
            record["frakmt"],
            record["tilkmt"],
            record["belaegningstype"],
            record["udlaegningsdato"],
            record["vejside"],
            record["tpe_vejside"],
            record["plade_tpe"],
            record["oeverste_lag"],
            record["roadname"],
            record["lane"],
            record["side_code"],
            None if record["geom"] is None else record["geom"].wkt,
            record["base_id"],
            record["base_name"],
            record["chain_from"],
            record["chain_to"],
            record["measure_from"],
            record["measure_to"],
            record["match_distance_m"],
            len(record["base_events"]),
            len(record["pms_events"]),
        )
        for record in records
    ]
    execute_values(
        cursor,
        """
        INSERT INTO dtims.nye_belaegninger_2026 (
            excel_row, admvejnr, admvejdel, frakmt, tilkmt, belaegningstype,
            udlaegningsdato, vejside, tpe_vejside, plade_tpe, oeverste_lag,
            roadname, lane, side_code, geom, base_id, base_name,
            chain_from, chain_to, measure_from, measure_to, match_distance_m,
            base_element_count, pms_section_count
        ) VALUES %s
        """,
        parent_rows,
        template="""(
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
            ST_Multi(ST_CollectionExtract(ST_MakeValid(ST_SetSRID(ST_GeomFromText(%s), 25832)), 2)),
            %s, %s, %s, %s, %s, %s, %s, %s, %s
        )""",
        page_size=40,
    )

    base_rows = []
    pms_rows = []
    for record in records:
        for event in record["base_events"]:
            base_rows.append(
                (
                    record["excel_row"],
                    event["base_id"],
                    event["base_name"],
                    record["roadname"],
                    record["lane"],
                    event["from"],
                    event["to"],
                    event["measure_from"],
                    event["measure_to"],
                    event["to"] - event["from"],
                    record["belaegningstype"],
                    record["udlaegningsdato"],
                    record["vejside"],
                    record["admvejnr"],
                    record["admvejdel"],
                    None if event["geom"] is None else event["geom"].wkt,
                )
            )
        for event in record["pms_events"]:
            pms_rows.append(
                (
                    record["excel_row"],
                    event["base_id"],
                    int(event["gid"]) if event["gid"] is not None else None,
                    event["name"],
                    record["roadname"],
                    record["lane"],
                    event["from"],
                    event["to"],
                    event["measure_from"],
                    event["measure_to"],
                    event["to"] - event["from"],
                    record["belaegningstype"],
                    record["udlaegningsdato"],
                    record["vejside"],
                    record["admvejnr"],
                    record["admvejdel"],
                    None if event["geom"] is None else event["geom"].wkt,
                )
            )
    if base_rows:
        execute_values(
            cursor,
            """
            INSERT INTO dtims.nye_belaegninger_2026_base (
                excel_row, base_id, base_name, roadname, lane, "from", "to",
                measure_from, measure_to, length_m, belaegningstype,
                udlaegningsdato, vejside, admvejnr, admvejdel, geom
            ) VALUES %s
            """,
            base_rows,
            template="""(
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                ST_Multi(ST_CollectionExtract(ST_MakeValid(ST_SetSRID(ST_GeomFromText(%s), 25832)), 2))
            )""",
            page_size=100,
        )
    if pms_rows:
        execute_values(
            cursor,
            """
            INSERT INTO dtims.nye_belaegninger_2026_pms (
                excel_row, base_id, pms_gid, name, roadname, lane, "from", "to",
                measure_from, measure_to, length_m, belaegningstype,
                udlaegningsdato, vejside, admvejnr, admvejdel, geom
            ) VALUES %s
            """,
            pms_rows,
            template="""(
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                ST_Multi(ST_CollectionExtract(ST_MakeValid(ST_SetSRID(ST_GeomFromText(%s), 25832)), 2))
            )""",
            page_size=100,
        )


def load_database(records):
    connection = connect()
    try:
        with connection:
            with connection.cursor() as cursor:
                base_features = load_base(cursor)
                sections = load_sections(cursor)
                print(
                    f"Læste {len(base_features)} base-elementer og "
                    f"{sum(len(items) for items in sections.values())} sektioner",
                    flush=True,
                )
                assign_measures(records, base_features, sections, True)
                create_tables(cursor)
                insert_records(cursor, records)
                cursor.execute(
                    """
                    SELECT count(*), count(geom), count(measure_from),
                           count(*) FILTER (WHERE base_element_count > 1)
                    FROM dtims.nye_belaegninger_2026
                    """
                )
                summary = cursor.fetchone()
                cursor.execute("SELECT count(*) FROM dtims.nye_belaegninger_2026_base")
                base_events = cursor.fetchone()[0]
                cursor.execute("SELECT count(*) FROM dtims.nye_belaegninger_2026_pms")
                pms_events = cursor.fetchone()[0]
        print(
            "Indlæst i dtims.nye_belaegninger_2026: "
            f"{summary[0]} rækker, {summary[1]} med geometri, "
            f"{summary[2]} med mål fra base, {summary[3]} fordelt på flere base-elementer.",
            flush=True,
        )
        print(f"Hændelser: {base_events} på base, {pms_events} på pms_analysis.", flush=True)
    finally:
        connection.close()


def main():
    workbook, sheet, header_row, rows = read_rows()
    network = load_network(rows)
    records = build_records(rows, network)
    try:
        load_database(records)
    except psycopg2.OperationalError as exc:
        print("Databasen afviste forbindelsen, så mål og hændelser er ikke skrevet.")
        print(str(exc).strip())
    write_excel(workbook, sheet, header_row, records)


if __name__ == "__main__":
    main()
