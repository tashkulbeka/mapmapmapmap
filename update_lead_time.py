# -*- coding: utf-8 -*-
# Lead-time updater for the EMEA map.
# Reads the raw shipments file, recalculates lane stats, writes them back into
# the same Excel (new sheets) and injects the numbers into the website file.
#
# Lane stats per country+city+service:
#   lead time  = mode (most frequent value)
#   range      = min .. frequency-weighted average (so outliers don't inflate it)
#   confidence = 3 if >=30 shipments, 2 if >=10, else 1
#
# First run reads the whole Excel and caches the per-lane frequency tables in
# _leadtime_cache.pkl. Later runs skip the raw file completely if it hasn't
# changed. Delete the cache to force a full rebuild.
#
# Run:  python update_lead_times.py

import json
import math
import os
import pickle
import re
import shutil
import sys
import unicodedata
import urllib.parse
import urllib.request
from collections import Counter
from difflib import get_close_matches
from pathlib import Path

import pandas as pd

# script lives in the "Updating" folder, the website one level up
BASE_DIR = Path(__file__).resolve().parent
RAW_XLSX = str(BASE_DIR / "Lead times raw data.xlsx")
HTML_FILE = r"C:\Users\betas\Pandora\LP X CR - General\JEWL\Lead times\Lead times map\Carrier lead times.html"
CACHE_FILE = str(BASE_DIR / "_leadtime_cache.pkl")
GEOCODE_CACHE = str(BASE_DIR / "_geocode_cache.json")
COUNTRIES_JSON = str(BASE_DIR / "countries.json")

RAW_SHEET = 0
CITY_SHEET = "City Lead Times"
REVIEW_SHEET = "Unmatched Cities"

SERVICE_MAP = {
    ("DHL Express", "BBX Break Bulk Express (doc)"): "dhl_bbx",
    ("DHL Express", "WPX Express Worldwide (non-doc)"): "dhl_exp",
    ("DHL Express", "ECX Express Worldwide (eu)"): "dhl_exp",
    ("DHL Express", "XPD Express Envelope (doc)"): "dhl_exp",
    ("DHL Express", "ESU Economy Select (doc)"): "dhl_std",
    ("UPS", "UPS Saver®"): "ups_exp",
    ("UPS", "UPS Expedited®"): "ups_exp",
    ("UPS", "UPS® Standard"): "ups_std",
    ("KUEHNE + NAGEL", "KN Road"): "kn_road",
}

# K&N reports use full country names instead of ISO codes
COUNTRY_MAP = {
    "ARMENIA": "AM", "AUSTRIA": "AT", "BELGIUM": "BE",
    "BOSNIA AND HERZEGOVINA": "BA", "BULGARIA": "BG", "CROATIA": "HR",
    "CZECH REPUBLIC": "CZ", "DENMARK": "DK", "ESTONIA": "EE",
    "FINLAND": "FI", "FRANCE": "FR", "GEORGIA": "GE", "GERMANY": "DE",
    "GREECE": "GR", "HUNGARY": "HU", "IRELAND": "IE", "ITALY": "IT",
    "LITHUANIA": "LT", "LUXEMBOURG": "LU", "MALTA": "MT",
    "MOLDOVA, REPUBLIC OF": "MD", "NETHERLANDS": "NL",
    "NORTH MACEDONIA": "MK", "NORWAY": "NO", "POLAND": "PL",
    "PORTUGAL": "PT", "ROMANIA": "RO", "SERBIA": "RS", "SLOVAKIA": "SK",
    "SLOVENIA": "SI", "SPAIN": "ES", "SWEDEN": "SE", "SWITZERLAND": "CH",
    "TURKEY": "TR", "UNITED KINGDOM": "GB",
}

# fixes for raw city names; targets must match names used in the website
CITY_ALIASES = {
    "ALTINOVA /ANTALYA": "ANTALYA",
    "ALANYA /ANTALYA": "ANTALYA",
    "KONYAALTI ANTALYA LIMAN": "ANTALYA",
    "KONYAALTI ANTALYA OTHERS": "ANTALYA",
    "WIEN": "VIENNA",
    "BEOGRAD": "BELGRADE",
    "KOBENHAVN K": "COPENHAGEN",
    "KOBENHAVN V": "COPENHAGEN",
    "KOEBENHAVN": "COPENHAGEN",
    "GOTEBORG": "GOTHENBURG",
    "MUENCHEN": "MUNCHEN",
    "MAILAND": "MILANO",
    "MESTRE VENEZIA": "VENEZIA",
    "FRIENZE": "FIRENZE",
    "PRAHA": "PRAGUE",
    "BRUXELLES": "BRUSSELS",
    "ATHINA": "ATHENS",
    "ARRECIFE DE LANZAROTE - LAS PALMAS": "LANZAROTE",
    "ARRECIFE DE LANZAROTE": "LANZAROTE",
    "DONOSTIA - SAN SEBASTIAN": "SAN SEBASTIAN",
    "QUPEYE": "OUPEYE",
    "GIUGLIANO IN C.": "GIUGLIANO IN CAMPANIA",
    "PLZEN": "PILSEN",
    "PLZEN 8": "PILSEN",
    "CAPAIORE-LUCCA": "CAPANNORI LUCCA",
    "SAN G. VALDARNO": "SAN GIOVANNI VALDARNO",
    "S. M. CAPUA VETERE": "SANTA MARIA CAPUA VETERE",
    "BARCELLONA POZZO G.": "BARCELLONA POZZO DI GOTTO",
    "CESKE BUDEJOVICE": "BUDEJOVICE",
    "LUDWIGSHAFEN": "LUDWIGSHAFEN AM RHEIN",
    "VALENA A": "VALENCA",
    "ZURICH-WALLISELLEN": "WALLISELLEN",
    "BUCURESTI": "BUCHAREST",
    "BOLOGNE": "BOLOGNA",
    "PAFOS": "PAPHOS",
    "SETAOBAL": "SETUBAL",
    "NEUSIEDL AM SEE": "NEUSIEDL/SEE",
    "LIA GE": "LIEGE",
    "DIDIMOTOICHO": "DIDYMOTEICHO",
    "A VORA": "EVORA",
    "GA TTINGEN": "GOTTINGEN",
    "SALZGITTER BAD": "SALZGITTER",
    "NASR CITY": "NASR CITY, CAIRO",
    "KUSADASI AYDIN TURKMEN": "KUSADASI/AYDIN",
    "ANDORRA": "ANDORRA LA VELLA",
    "ALICANTE": "ALACANT/ALICANTE",
    "PISA": "PIZA",
    "VELIZY-VILLACOUBLAY": "VELIZY",
    "FIUMICINO": "FIUMICINO AEROPORTO",
    "LAS PALMAS": "LAS PALMAS DE GRAN CANARIA",
    "TERRASA": "TERRASA (BARCELONA)",
    "SAN EUGENIO - ADEJE": "COSTA ADEJE",
    "DESSAU": "DESSAU-ROSSLAU",
    "SELCUKLU KONYA": "KONYA",
    "YENIMAHALLE ANKARA": "ANKARA",
    "SEHITKAMIL GAZIANTEP OTHERS": "GAZIANTEP",
    "ORTAHISAR TRABZON KALKINMA": "TRABZON",
    "YENISEHIR MERSIN": "MERSIN",
    "ONIKISUBAT KAHRAMANMARAS SAZI BEY": "KAHRAMANMARAS",
    "SAMSUN E.A. HAST.": "SAMSUN",
}

# new cities you actually want on the map. The review sheet gives ready-made
# lines to paste here - nothing gets on the map without being approved.
APPROVED_CITIES = {
    "ES|POLA DE SIERO": {"name": "Pola De Siero", "lat": 43.388, "lon": -5.6648},
    "ES|ARROYO DE LA MIEL": {"name": "Arroyo De La Miel", "lat": 36.1125, "lon": -5.4968},
    "FR|PORTET S/GARONNE": {"name": "Portet S/Garonne", "lat": 43.5222, "lon": 1.4076},
    "IT|ACILIA (RM)": {"name": "Acilia (Rm)", "lat": 41.7827, "lon": 12.3653},
    "ES|SANTA PONSA (CALVIA)": {"name": "Santa Ponsa (Calvia)", "lat": 39.5162, "lon": 2.4832},
    "IT|MIRANO": {"name": "Mirano", "lat": 45.4878, "lon": 12.0859},
    "ES|BLANCA DONA (EIVISSA)": {"name": "Blanca Dona (Eivissa)", "lat": 38.9223, "lon": 1.4263},
    "GB|CRAMLINGTON": {"name": "Cramlington", "lat": 55.0856, "lon": -1.5907},
    "RO|GLINA": {"name": "Glina", "lat": 44.3862, "lon": 26.2452},
    "ES|COLL DA EN RABASSA (ISLAS BALEARES)": {"name": "Coll Da En Rabassa (Islas Baleares)", "lat": 39.5508, "lon": 2.6963},
    "ES|FUENGIROLA": {"name": "Fuengirola", "lat": 36.5573, "lon": -4.6209},
    "PT|MOREIRAA -A MAIA": {"name": "Moreiraa -A Maia", "lat": 41.2371, "lon": -8.6539},
}

CONF_RULES = [(30, 3), (10, 2), (0, 1)]
FUZZY_CUTOFF = 0.86

TRANSLIT = str.maketrans({"Ł": "L", "ł": "l", "Ø": "O", "ø": "o", "Đ": "D",
                          "đ": "d", "ß": "ss", "Æ": "AE", "æ": "ae",
                          "Œ": "OE", "œ": "oe"})
GARBAGE = re.compile(r"[§¶‰“”\"'_]|X0*0?0?D", re.I)
PAREN = re.compile(r"\([^)]*\)")
CITIES_RE = re.compile(r"(const CITIES\s*=\s*)(\[.*\])(;\s*(?://.*)?$)", re.M)
AGG_RE = re.compile(r"(const AGG\s*=\s*)(\{.*\})(;\s*(?://.*)?$)", re.M)
SHAPES_RE = re.compile(r"(const SHAPES\s*=\s*)(\{.*\})(;\s*(?://.*)?$)", re.M)

# same Miller-variant projection the website uses to place countries
MAP_W, MAP_H = 3000, 1695
_SY, _DY, _CROP = -840.9495796213082, 1762.4811355830254, [140, 450, 3520, 2360]


def overview_xy(lat, lon):
    miller = 1.25 * math.log(math.tan(math.pi / 4 + 0.4 * lat * math.pi / 180))
    x = (14.765 * lon + 1998.5 - _CROP[0]) * MAP_W / (_CROP[2] - _CROP[0])
    y = (_SY * miller + _DY - _CROP[1]) * MAP_H / (_CROP[3] - _CROP[1])
    return round(x, 1), round(y, 1)


def norm(name):
    s = str(name).translate(TRANSLIT)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = GARBAGE.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip().upper()


def mode_min_wmax(counter):
    n = sum(counter.values())
    top = max(counter.values())
    mode = min(v for v, f in counter.items() if f == top)
    wavg = math.ceil(sum(v * f for v, f in counter.items()) / n)
    return mode, min(counter), max(mode, min(wavg, max(counter))), n


def conf_of(n):
    for threshold, level in CONF_RULES:
        if n >= threshold:
            return level
    return 1


def cache_path(xlsx_path, name):
    return os.path.join(os.path.dirname(os.path.abspath(xlsx_path)), name)


def save_agg_cache(xlsx_path, counters):
    sig = (os.path.getsize(xlsx_path), os.path.getmtime(xlsx_path))
    with open(cache_path(xlsx_path, CACHE_FILE), "wb") as f:
        pickle.dump({"sig": sig, "counters": counters}, f)


def load_counters(xlsx_path):
    path = cache_path(xlsx_path, CACHE_FILE)
    reg_mtime = os.path.getmtime(COUNTRIES_JSON) if os.path.exists(COUNTRIES_JSON) else 0
    sig = (os.path.getsize(xlsx_path), os.path.getmtime(xlsx_path), reg_mtime)
    if os.path.exists(path):
        try:
            cached = pickle.load(open(path, "rb"))
            if cached.get("sig") == sig:
                print("Excel unchanged, using cache (raw file not re-read).")
                return cached["counters"]
        except Exception:
            pass

    print("Reading raw Excel...")
    df = pd.read_excel(xlsx_path, sheet_name=RAW_SHEET,
                       usecols=["Carrier", "Service", "Receiver City",
                                "Receiver Country", "Date Shipped", "Delivery Date"])
    df["Date Shipped"] = pd.to_datetime(df["Date Shipped"], errors="coerce")
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], errors="coerce")
    lead = (df["Delivery Date"] - df["Date Shipped"]).dt.days

    negative = int((lead < 0).sum())
    df = df[lead.notna() & (lead >= 0)].copy()
    df["lead"] = lead[lead.notna() & (lead >= 0)].astype(int)

    df["skey"] = df.apply(lambda r: SERVICE_MAP.get((r["Carrier"], r["Service"])), axis=1)
    unmapped = df[df["skey"].isna()]
    if not unmapped.empty:
        print("WARNING, unmapped carrier/service combos (add to SERVICE_MAP):")
        print(unmapped.groupby(["Carrier", "Service"]).size().to_string())
    df = df[df["skey"].notna()]

    name_map = dict(COUNTRY_MAP)
    if os.path.exists(COUNTRIES_JSON):
        reg = json.load(open(COUNTRIES_JSON, encoding="utf-8"))
        for cc, info in reg.items():   # full country names from the registry
            name_map.setdefault(info["n"].upper(), cc)
    df["cc"] = (df["Receiver Country"].astype(str).str.strip().str.upper()
                .map(lambda c: name_map.get(c, c)))
    unknown_cc = sorted(set(df["cc"]) - set(name_map.values()) -
                        {c for c in df["cc"] if len(c) == 2})
    if unknown_cc:
        print("WARNING, unknown country names (add to COUNTRY_MAP):", unknown_cc)

    counters = {}
    for key, g in df.groupby(["cc", df["Receiver City"].map(norm), "skey"]):
        counters[key] = Counter(g["lead"])

    save_agg_cache(xlsx_path, counters)
    print(f"Processed {len(df):,} rows ({negative} negative-lead rows excluded).")
    return counters


def load_website(html_path):
    html = open(html_path, encoding="utf-8").read()
    m1, m2, m3 = CITIES_RE.search(html), AGG_RE.search(html), SHAPES_RE.search(html)
    if not (m1 and m2 and m3):
        sys.exit("Could not find CITIES / AGG / SHAPES blocks in the HTML.")
    return html, json.loads(m1.group(2)), json.loads(m2.group(2)), json.loads(m3.group(2))


def load_registry():
    if not os.path.exists(COUNTRIES_JSON):
        print("NOTE: countries.json not found, new countries cannot be added.")
        return {}
    return json.load(open(COUNTRIES_JSON, encoding="utf-8"))


def add_country(cc, agg, shapes, registry):
    info = registry[cc]
    x, y = overview_xy(info["cap"]["lat"], info["cap"]["lon"])
    agg[cc] = {"n": info["n"], "x": x, "y": y,
               "lat": info["cap"]["lat"], "lon": info["cap"]["lon"],
               "cities": 0, "ship": 0, "reg": info["reg"], "svc": {}}
    flat = [[v for lon, lat in ring for v in
             (int(round(lon * 100)), int(round(lat * 100)))] for ring in info["rings"]]
    if flat:
        xs = [v for ring in flat for v in ring[0::2]]
        ys = [v for ring in flat for v in ring[1::2]]
        shapes[cc] = {"n": info["n"],
                      "b": [min(xs), min(ys), max(xs), max(ys)], "p": flat}
    print(f"  new country added: {cc} ({info['n']}, {info['reg']})")


def geocode_city(raw_city, cc, country_hint=""):
    path = cache_path(RAW_XLSX, GEOCODE_CACHE)
    cache = {}
    if os.path.exists(path):
        try:
            cache = json.load(open(path, encoding="utf-8"))
        except Exception:
            pass
    key = f"{cc}|{raw_city}"
    if key in cache:
        return tuple(cache[key]) if cache[key] else None
    result = None
    try:
        url = "https://photon.komoot.io/api/?" + urllib.parse.urlencode(
            {"q": f"{raw_city.title()}, {country_hint or cc}", "limit": 1})
        req = urllib.request.Request(url, headers={"User-Agent": "leadtime-map/1.0"})
        feats = json.loads(urllib.request.urlopen(req, timeout=30).read())["features"]
        if feats:
            lon, lat = feats[0]["geometry"]["coordinates"]
            result = (round(lat, 4), round(lon, 4))
    except Exception as e:
        print(f"  geocode failed for {key}: {e}")
    cache[key] = list(result) if result else None
    json.dump(cache, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return result


def subset_match(tokens, city_idx):
    t = set(tokens.split())
    hits = [rec for name, rec in city_idx.items()
            if t and (t <= set(name.split()) or set(name.split()) <= t)]
    return hits[0] if len(hits) == 1 else None


def resolve_city(raw_city, cc, city_idx, full_idx):
    target = CITY_ALIASES.get(raw_city, raw_city)
    if target in city_idx:
        return city_idx[target]
    variants = {PAREN.sub("", target).strip(), target.split(",")[0].strip()}
    variants |= {p.strip() for p in target.split("-")}
    for v in variants:
        if v and v in city_idx:
            return city_idx[v]
    close = get_close_matches(target, list(city_idx), n=1, cutoff=FUZZY_CUTOFF)
    if close:
        return city_idx[close[0]]
    rec = subset_match(target, city_idx)
    if rec:
        return rec
    for other_cc, other_idx in full_idx.items():   # wrong country in raw data
        if other_cc != cc and target in other_idx:
            print(f"  note: '{raw_city}' labelled {cc} but matched in {other_cc}")
            return other_idx[target]
    return None


def merge(counters, cities, agg, shapes, registry):
    idx = {}
    for c in cities:
        idx.setdefault(c["cc"], {})[norm(c["c"])] = c
    unmatched, suggestions, combined = Counter(), {}, {}
    new_ccs = set()

    for (cc, raw_city, skey), cnt in counters.items():
        if cc not in agg:
            if cc in registry:
                add_country(cc, agg, shapes, registry)
                new_ccs.add(cc)
            else:
                unmatched[f"{cc} (country not on map, not in registry)"] += sum(cnt.values())
                continue
        city_idx = idx.get(cc, {})
        rec = resolve_city(raw_city, cc, city_idx, idx)
        if rec is None and f"{cc}|{raw_city}" in APPROVED_CITIES:
            ap = APPROVED_CITIES[f"{cc}|{raw_city}"]
            rec = {"cc": cc, "c": ap["name"], "z": "",
                   "la": ap["lat"], "lo": ap["lon"], "s": {}}
            cities.append(rec)
            city_idx[raw_city] = rec
            idx.setdefault(cc, {})[raw_city] = rec
        if rec is None:
            coords = geocode_city(raw_city, cc, agg[cc].get("n", cc))
            if cc in new_ccs and coords:
                rec = {"cc": cc, "c": raw_city.title(), "z": "",
                       "la": coords[0], "lo": coords[1], "s": {}}
                if norm(rec["c"]) == norm(registry[cc]["cap"]["name"]):
                    rec["cap"] = 1
                cities.append(rec)
                idx.setdefault(cc, {})[norm(rec["c"])] = rec
                print(f"  new city added: {rec['c']}, {cc} (geocoded)")
            else:
                unmatched[f"{cc} | {raw_city}"] += sum(cnt.values())
                suggestions[f"{cc} | {raw_city}"] = coords
                continue
        combined.setdefault((id(rec), skey), [rec, Counter()])[1].update(cnt)

    for (rid, skey), (rec, cnt) in combined.items():
        mode, mn, mx, n = mode_min_wmax(cnt)
        rec["s"][skey] = [mode, mn, mx, n, conf_of(n)]
    return unmatched, suggestions


def rebuild_agg(cities, agg):
    # overview cards take their lead times from the capital city's lanes
    for cc, a in agg.items():
        cc_cities = [c for c in cities if c["cc"] == cc]
        a["cities"] = len(cc_cities)
        if not cc_cities:
            continue  # e.g. Kosovo - keep the manual country-level services
        a["ship"] = sum(v[3] for c in cc_cities for v in c["s"].values())
        caps = [c for c in cc_cities if c.get("cap")] or cc_cities
        svc = {}
        for k in {k for c in cc_cities for k in c["s"]}:
            total = sum(c["s"][k][3] for c in cc_cities if k in c["s"])
            cap_entry = next((c["s"][k] for c in caps if k in c["s"]), None)
            if cap_entry:
                svc[k] = {"a": cap_entry[0], "mn": cap_entry[1],
                          "mx": cap_entry[2], "s": total}
            else:
                lanes = [c["s"][k] for c in cc_cities if k in c["s"]]
                w = sum(l[3] for l in lanes) or 1
                svc[k] = {"a": round(sum(l[0] * l[3] for l in lanes) / w),
                          "mn": min(l[1] for l in lanes),
                          "mx": max(l[2] for l in lanes), "s": total}
        a["svc"] = svc


def dump(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def save_website(html_path, html, cities, agg, shapes):
    bak = str(BASE_DIR / (Path(html_path).name + ".bak"))  # backup stays in Updating, not SharePoint
    if not os.path.exists(bak):
        shutil.copy2(html_path, bak)
    cities.sort(key=lambda c: (c["cc"], c["c"]))
    html = CITIES_RE.sub(lambda m: m.group(1) + dump(cities) + m.group(3), html, count=1)
    html = AGG_RE.sub(lambda m: m.group(1) + dump(agg) + m.group(3), html, count=1)
    html = SHAPES_RE.sub(lambda m: m.group(1) + dump(shapes) + m.group(3), html, count=1)
    open(html_path, "w", encoding="utf-8").write(html)
    print(f"Website updated: {html_path}")


def save_excel(xlsx_path, cities):
    rows = [{"Country": c["cc"], "City": c["c"], "Service": k,
             "Lead time (mode)": v[0], "Range min": v[1], "Range max (w.avg)": v[2],
             "Shipments": v[3], "Confidence": v[4]}
            for c in cities for k, v in c["s"].items() if v[3] > 0]
    lanes = pd.DataFrame(rows).sort_values(["Country", "City", "Service"])
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as w:
        lanes.to_excel(w, sheet_name=CITY_SHEET, index=False)
    print(f"Excel updated: '{CITY_SHEET}' -> {len(lanes)} lanes.")


def save_review(xlsx_path, unmatched, suggestions):
    if not unmatched:
        return
    rows = []
    for k, v in unmatched.most_common():
        cc, raw = k.split(" | ", 1) if " | " in k else (k, "")
        sug = suggestions.get(k)
        approve = (f'"{cc}|{raw}": {{"name": "{raw.title()}", '
                   f'"lat": {sug[0]}, "lon": {sug[1]}}},') if sug else ""
        rows.append({"Unmatched raw city": k, "Shipments": v,
                     "Suggested lat": sug[0] if sug else "",
                     "Suggested lon": sug[1] if sug else "",
                     "Paste into APPROVED_CITIES to add to map": approve})
    rev = pd.DataFrame(rows)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as w:
        rev.to_excel(w, sheet_name=REVIEW_SHEET, index=False)
    print(f"NOTE: {len(unmatched)} unmatched cities NOT added to map -> '{REVIEW_SHEET}'.")


def main():
    counters = load_counters(RAW_XLSX)
    html, cities, agg, shapes = load_website(HTML_FILE)
    registry = load_registry()
    unmatched, suggestions = merge(counters, cities, agg, shapes, registry)
    rebuild_agg(cities, agg)
    save_website(HTML_FILE, html, cities, agg, shapes)
    save_excel(RAW_XLSX, cities)
    save_review(RAW_XLSX, unmatched, suggestions)
    save_agg_cache(RAW_XLSX, counters)  # re-stamp, sheets changed the file
    print("Done.")


if __name__ == "__main__":
    main()
