"""Hent Trafik_trafiktal (type 2) via linear search for pms_analysis_aadt_null.

Matcher på Vejman-kilometrering fra kmstringfrom/kmstringto (ikke PMS from/to).
"""

from __future__ import annotations

import json
import os
import re
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import psycopg2
import requests
from dotenv import load_dotenv

URL = "https://linearsearch.vd.dk/linear-search-api/v4/execute"
ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "_linear_cache" / "trafiktal"
LANE_TO_TPE = {0: 0, 1: 1, 2: -1}  # midt, højre, venstre
KMSTRING_RE = re.compile(r"^(\d+)/(\d+)$")

load_dotenv(ROOT / "exp.env")


def connect():
    return psycopg2.connect(
        host=os.environ["POSTGRES_HOST"],
        port=os.environ["POSTGRES_PORT"],
        user=os.environ["POSTGRES_USER"],
        password=os.environ["POSTGRES_PASSWORD"],
        dbname=os.environ["POSTGRES_DB"],
    )


def kmstring_to_kmt(value: str | None) -> float | None:
    """Konverter '58/0637' -> 580637.0 (fjern skråstreg)."""
    if value is None:
        return None
    text = str(value).strip()
    match = KMSTRING_RE.match(text)
    if not match:
        raise ValueError(f"Ukendt kmstring-format: {value!r}")
    return float(f"{match.group(1)}{match.group(2)}")


def fetch_trafiktal(admvejnr: int, admvejdel: str) -> list[dict]:
    CACHE.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE / f"{admvejnr}_{admvejdel}_type2.json"
    if cache_path.exists():
        return json.loads(cache_path.read_text(encoding="utf-8"))

    query = (
        "vm:admvejnr, vm:admvejdel, vm:TPE_vejside, vm:frakmt, vm:tilkmt, "
        "vm:Trafik_trafiktal, vm:Trafik_trafiktaltype, vm:Trafik_tælleår\n"
        "where\n"
        f"vm:admvejnr = {admvejnr} AND vm:admvejdel = '{admvejdel}' "
        "AND vm:Trafik_trafiktaltype = 2"
    )
    last_err = None
    for attempt in range(3):
        try:
            response = requests.get(URL, params={"query": query}, timeout=180)
            response.raise_for_status()
            features = response.json().get("features") or []
            rows = []
            seen = set()
            for feature in features:
                props = feature.get("properties") or {}
                row = {
                    "admvejnr": props.get("Admvejnr"),
                    "admvejdel": props.get("Admvejdel"),
                    "tpe_vejside": props.get("TPE_Vejside"),
                    "frakmt": props.get("Frakmt"),
                    "tilkmt": props.get("Tilkmt"),
                    "trafiktal": props.get("Trafik_trafiktal"),
                    "trafiktaltype": props.get("Trafik_trafiktaltype"),
                    "taelleaar": props.get("Trafik_tælleår"),
                }
                key = (
                    row["tpe_vejside"],
                    row["frakmt"],
                    row["tilkmt"],
                    row["trafiktal"],
                    row["taelleaar"],
                )
                if key in seen:
                    continue
                seen.add(key)
                rows.append(row)
            cache_path.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
            return rows
        except requests.RequestException as exc:
            last_err = exc
            time.sleep(2 * (attempt + 1))
    raise last_err


def overlap_len(a0: float, a1: float, b0: float, b1: float) -> float:
    return max(0.0, min(a1, b1) - max(a0, b0))


def _candidates_for_tpe(section: dict, segments: list[dict], tpe: int) -> list[tuple[float, dict]]:
    candidates = []
    for seg in segments:
        if seg["tpe_vejside"] != tpe:
            continue
        if seg["frakmt"] is None or seg["tilkmt"] is None or seg["trafiktal"] is None:
            continue
        ov = overlap_len(
            section["frakmt"],
            section["tilkmt"],
            float(seg["frakmt"]),
            float(seg["tilkmt"]),
        )
        if ov <= 0:
            continue
        candidates.append((ov, seg))
    return candidates


def _choose_from_candidates(candidates: list[tuple[float, dict]], tpe: int, match_source: str) -> dict:
    by_value: dict[float, float] = defaultdict(float)
    year_by_value: dict[float, float] = {}
    for ov, seg in candidates:
        val = float(seg["trafiktal"])
        by_value[val] += ov
        year = seg["taelleaar"]
        if year is not None:
            prev = year_by_value.get(val)
            if prev is None or float(year) > prev:
                year_by_value[val] = float(year)

    total_ov = sum(by_value.values())
    weighted = sum(v * ov for v, ov in by_value.items()) / total_ov
    dominant = max(by_value.items(), key=lambda x: x[1])
    if dominant[1] / total_ov >= 0.8:
        chosen = dominant[0]
    else:
        chosen = round(weighted)
    return {
        "trafiktal": chosen,
        "taelleaar": year_by_value.get(chosen)
        or max(
            (s["taelleaar"] for _, s in candidates if s["taelleaar"] is not None),
            default=None,
        ),
        "overlap_m": round(total_ov, 2),
        "n_segments": len(candidates),
        "tpe_vejside": tpe,
        "match_source": match_source,
    }


def pick_trafiktal(section: dict, segments: list[dict]) -> dict | None:
    tpe = LANE_TO_TPE.get(int(section["lane"]) if section["lane"] is not None else -999)
    if tpe is None:
        return None

    candidates = _candidates_for_tpe(section, segments, tpe)
    if candidates:
        return _choose_from_candidates(candidates, tpe, "side")

    # Fallback: midt (TPE 0), hvis sektionen selv ikke allerede er midt
    if tpe != 0:
        midt = _candidates_for_tpe(section, segments, 0)
        if midt:
            return _choose_from_candidates(midt, 0, "midt_fallback")
    return None


def main():
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT gid, roadname, name, lane, length,
                       "from" AS pms_from, "to" AS pms_to,
                       kmstringfrom, kmstringto
                FROM dtims.pms_analysis_aadt_null
                ORDER BY roadname, lane, "from"
                """
            )
            sections = []
            for r in cur.fetchall():
                sections.append(
                    {
                        "gid": r[0],
                        "roadname": r[1],
                        "name": r[2],
                        "lane": r[3],
                        "length": r[4],
                        "pms_from": float(r[5]),
                        "pms_to": float(r[6]),
                        "kmstringfrom": r[7],
                        "kmstringto": r[8],
                        "frakmt": kmstring_to_kmt(r[7]),
                        "tilkmt": kmstring_to_kmt(r[8]),
                    }
                )

        roads = sorted(
            {
                (
                    int(s["roadname"].rsplit("-", 1)[0]),
                    s["roadname"].rsplit("-", 1)[1],
                )
                for s in sections
            }
        )
        print(f"{len(sections)} sektioner, {len(roads)} vejdele", flush=True)
        print(
            f"Eksempel kmstring: {sections[0]['kmstringfrom']} -> {sections[0]['frakmt']:.0f}",
            flush=True,
        )

        traffic_by_road: dict[tuple[int, str], list[dict]] = {}
        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {
                pool.submit(fetch_trafiktal, vejnr, vejdel): (vejnr, vejdel)
                for vejnr, vejdel in roads
            }
            for fut in as_completed(futures):
                key = futures[fut]
                try:
                    rows = fut.result()
                    traffic_by_road[key] = rows
                    print(f"  {key[0]}-{key[1]}: {len(rows)} trafikstykker", flush=True)
                except Exception as exc:
                    traffic_by_road[key] = []
                    print(f"  {key[0]}-{key[1]}: FEJL {exc}", flush=True)

        matches = []
        missing = []
        for section in sections:
            vejnr = int(section["roadname"].rsplit("-", 1)[0])
            vejdel = section["roadname"].rsplit("-", 1)[1]
            hit = pick_trafiktal(section, traffic_by_road.get((vejnr, vejdel), []))
            if hit is None:
                missing.append(section)
                continue
            matches.append({**section, **hit, "admvejnr": vejnr, "admvejdel": vejdel})

        with conn.cursor() as cur:
            cur.execute("DROP TABLE IF EXISTS dtims.pms_analysis_aadt_null_trafiktal")
            cur.execute(
                """
                CREATE TABLE dtims.pms_analysis_aadt_null_trafiktal (
                    gid integer PRIMARY KEY,
                    roadname text,
                    name text,
                    lane double precision,
                    tpe_vejside integer,
                    pms_from double precision,
                    pms_to double precision,
                    kmstringfrom text,
                    kmstringto text,
                    frakmt double precision,
                    tilkmt double precision,
                    length double precision,
                    trafiktal double precision,
                    taelleaar double precision,
                    overlap_m double precision,
                    n_segments integer,
                    match_source text
                )
                """
            )
            for m in matches:
                cur.execute(
                    """
                    INSERT INTO dtims.pms_analysis_aadt_null_trafiktal
                    (gid, roadname, name, lane, tpe_vejside, pms_from, pms_to,
                     kmstringfrom, kmstringto, frakmt, tilkmt, length,
                     trafiktal, taelleaar, overlap_m, n_segments, match_source)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """,
                    (
                        m["gid"],
                        m["roadname"],
                        m["name"],
                        m["lane"],
                        m["tpe_vejside"],
                        m["pms_from"],
                        m["pms_to"],
                        m["kmstringfrom"],
                        m["kmstringto"],
                        m["frakmt"],
                        m["tilkmt"],
                        m["length"],
                        m["trafiktal"],
                        m["taelleaar"],
                        m["overlap_m"],
                        m["n_segments"],
                        m["match_source"],
                    ),
                )
            cur.execute(
                """
                COMMENT ON TABLE dtims.pms_analysis_aadt_null_trafiktal IS
                'Linear search Trafik_trafiktal (type 2) matchet på pms_analysis_aadt_null via vej/del/side og kmstringfrom/kmstringto (Vejman-kmt). match_source=side eller midt_fallback.';
                """
            )
            cur.execute(
                """
                SELECT match_source,
                       COUNT(*) AS n,
                       ROUND(SUM(length)::numeric, 1) AS length_m
                FROM dtims.pms_analysis_aadt_null_trafiktal
                GROUP BY match_source
                ORDER BY match_source
                """
            )
            by_source = cur.fetchall()
            cur.execute(
                """
                SELECT COUNT(*) AS matched,
                       ROUND(SUM(length)::numeric, 1) AS matched_m
                FROM dtims.pms_analysis_aadt_null_trafiktal
                """
            )
            stats = cur.fetchone()
            cur.execute(
                """
                SELECT name, lane, kmstringfrom, kmstringto,
                       trafiktal, taelleaar, match_source
                FROM dtims.pms_analysis_aadt_null_trafiktal
                WHERE match_source = 'midt_fallback'
                ORDER BY roadname, pms_from
                """
            )
            fallback_rows = cur.fetchall()
        conn.commit()

    print()
    print(f"Matchet: {stats[0]} / {len(sections)} sektioner ({stats[1]} m)")
    for source, n, length_m in by_source:
        print(f"  {source}: {n} sektioner ({length_m} m)")
    print(f"Uden match: {len(missing)}")
    if fallback_rows:
        print(f"Nye via midt_fallback ({len(fallback_rows)}):")
        for s in fallback_rows:
            print(f"  {s}")
    if missing:
        print("Mangler stadig:")
        for s in missing:
            print(
                f"  {s['name']} lane={s['lane']} "
                f"{s['kmstringfrom']}-{s['kmstringto']} len={s['length']}"
            )


if __name__ == "__main__":
    main()
