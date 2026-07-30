# -*- coding: utf-8 -*-
"""
Build an updated self-contained Delivery Map HTML:
  * reads the current MapData sheet (fresh data, all 14 columns)
  * patches the existing viewer (delivery map text.txt) to add:
      - POOL status filter button (PLANNEDORNO = POOL)
      - Load Group filter buttons (column N)
      - Delivery-date dual-range slider
      - POOL color on pins / legend / popup header
      - all 14 requested columns shown in every popup
"""
import os, sys, json, shutil, tempfile
from datetime import datetime, date

import openpyxl

# ---------------------------------------------------------------------------
# CONFIG  -- everything lives in THIS folder, so the kit runs on any machine.
# Point SRC_XLSX at YOUR team's data sheet (same columns as DATA-FORMAT.md).
# Your ship-from origins are defined in DEFAULT_PICKS lower in this file.
# ---------------------------------------------------------------------------
BASE = os.path.dirname(os.path.abspath(__file__))
SRC_XLSX   = os.path.join(BASE, "MapData sheet.xlsx")                        # <-- drop your data sheet here (or change this path)
SRC_HTML   = os.path.join(BASE, "delivery map text.txt")                    # base viewer template (do not rename)
OUT_HTML   = os.path.join(BASE, "DeliveryMap (all columns + filters).html")  # the map this builds
MAKER_HTML = os.path.join(BASE, "Map Maker.html")                           # optional; not needed (SheetJS is bundled)
SHEETJS_CACHE = os.path.join(BASE, "sheetjs.js")                            # bundled SheetJS reader (keep this file)
ZIP_COORDS    = os.path.join(BASE, "zip_coords.json")                       # US+CA ZIP/postal → lat,lng lookup table
STOP_LOCS     = os.path.join(BASE, "stop_locations.json")                   # enabled stop locations → lat,lng lookup


_zip_db = None
_stop_db = None

def _load_stop_db():
    global _stop_db
    if _stop_db is not None:
        return _stop_db
    if not os.path.exists(STOP_LOCS):
        print("  NOTE: stop_locations.json not found — will fall back to ZIP geocoding.")
        _stop_db = {"by_ref": {}, "by_namezip": {}}
        return _stop_db
    with open(STOP_LOCS, "r", encoding="utf-8") as f:
        _stop_db = json.load(f)
    return _stop_db

def _stop_to_coords(locref, name, postal):
    """Return (lat, lng) using stop locations DB, or None.
    Tries: 1) exact location reference, 2) name+postal combined key."""
    db = _load_stop_db()
    if locref:
        entry = db["by_ref"].get(str(locref).strip())
        if entry:
            return (entry[0], entry[1])
    if name and postal:
        n = str(name).strip().upper()
        p = str(postal).strip().split("-")[0].strip()
        p = p.zfill(5) if p[:1].isdigit() else p.upper()
        entry = db["by_namezip"].get(n + "|" + p)
        if entry:
            return (entry[0], entry[1])
    return None
def _load_zip_db():
    global _zip_db
    if _zip_db is not None:
        return _zip_db
    if not os.path.exists(ZIP_COORDS):
        print("  WARNING: zip_coords.json not found — orders without lat/lng will be skipped.")
        _zip_db = {}
        return _zip_db
    with open(ZIP_COORDS, "r", encoding="utf-8") as f:
        _zip_db = json.load(f)
    return _zip_db

def _zip_to_coords(z):
    """Return (lat, lng) for a ZIP/postal code, or None if not found."""
    if not z:
        return None
    db = _load_zip_db()
    # Try exact match first (handles Canadian "A1B 2C3" format)
    key = str(z).strip()
    if key in db:
        c = db[key]; return (c[0], c[1])
    # US: normalise to 5-digit (strip ZIP+4 suffix)
    key5 = key.split("-")[0].strip().zfill(5)
    if key5 in db:
        c = db[key5]; return (c[0], c[1])
    return None


RATES_XLSX = os.path.join(BASE, "DH_CURRENT_RATES_7.30.26.xlsx")

# Canadian province code → 2-letter abbreviation for state matching
_CA_POSTAL_PROV = {
    'A':'NL','B':'NS','C':'PE','E':'NB','G':'QC','H':'QC','J':'QC',
    'K':'ON','L':'ON','M':'ON','N':'ON','P':'ON','R':'MB','S':'SK',
    'T':'AB','V':'BC','X':'NT','Y':'YT',
}
# US ZIP first-digit prefix → state abbreviations (first 3 digits → state for precision)
# We build a compact ZIP3→state table for the ~161 unique drop ZIPs we have
_US_ZIP3_STATE = {}

def _zip_to_state(z):
    """Return 2-letter state/province code for a ZIP/postal code, or None."""
    z = str(z or '').strip()
    if not z:
        return None
    if not z[:1].isdigit():
        return _CA_POSTAL_PROV.get(z[:1].upper())
    # US ZIP: use 3-digit prefix lookup (built on first call)
    if not _US_ZIP3_STATE:
        _build_zip3_state()
    return _US_ZIP3_STATE.get(z[:3].zfill(3))

# Build ZIP3→state from known US ZIP ranges
def _build_zip3_state():
    ranges = [
        ('005','005','MA'),('006','009','PR'),('010','027','MA'),('028','029','RI'),
        ('030','038','NH'),('039','039','ME'),('040','049','ME'),('050','059','VT'),
        ('060','069','CT'),('070','089','NJ'),('090','098','AE'),('100','149','NY'),
        ('150','196','PA'),('197','199','DE'),('200','205','DC'),('206','212','MD'),
        ('214','219','MD'),('220','246','VA'),('247','268','WV'),('270','289','NC'),
        ('290','299','SC'),('300','319','GA'),('320','349','FL'),('350','369','AL'),
        ('370','385','TN'),('386','397','MS'),('398','399','GA'),('400','427','KY'),
        ('430','458','OH'),('460','479','IN'),('480','499','MI'),('500','528','IA'),
        ('530','549','WI'),('550','567','MN'),('570','577','SD'),('580','588','ND'),
        ('590','599','MT'),('600','620','IL'),('622','631','IL'),('633','641','MO'),
        ('644','658','MO'),('660','679','KS'),('680','693','NE'),('700','714','LA'),
        ('716','729','AR'),('730','749','OK'),('750','799','TX'),('800','816','CO'),
        ('820','831','WY'),('832','838','ID'),('840','847','UT'),('850','865','AZ'),
        ('870','884','NM'),('889','898','NV'),('900','961','CA'),('970','979','OR'),
        ('980','994','WA'),('995','999','AK'),
    ]
    for lo, hi, st in ranges:
        for p in range(int(lo), int(hi)+1):
            _US_ZIP3_STATE[str(p).zfill(3)] = st


def build_rates_lookup():
    """Parse TL rates file; return compact lookup for JS embedding.

    Returns dict:
      { orig_city_upper: { dest_state_upper: [ {c,b,r,m,cy}, ... ] } }
    Only keeps rates where origin matches a DEFAULT_PICKS city.
    Carriers sorted cheapest-first by CPM rate (FLT carriers shown as separate entry).
    """
    if not os.path.exists(RATES_XLSX):
        print(f"  NOTE: rates file not found at {RATES_XLSX} — rate display disabled.")
        return {}

    import openpyxl as _xl
    wb = _xl.load_workbook(RATES_XLSX)
    ws = wb['TL CURRENT RATES']
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    pick_cities = {p["match"][0].upper() for p in DEFAULT_PICKS}
    pick_cities.add("ST PAUL")

    from collections import defaultdict

    # ---- TL rates (CPM / FLT) ----
    tl_raw = defaultdict(lambda: defaultdict(list))
    for r in rows:
        orig_city = str(r[18] or '').strip().upper()
        if orig_city not in pick_cities:
            continue
        carrier   = str(r[3]  or '').strip().upper()
        basis     = str(r[10] or '').strip().upper()
        rate_val  = float(r[11]) if r[11] is not None else 0.0
        min_cost  = float(r[12]) if r[12] is not None else 0.0
        currency  = str(r[14] or 'USD').strip().upper()
        dest_city = str(r[25] or '').strip().upper()
        dest_st   = str(r[26] or '').strip().upper()
        if not dest_city or not dest_st:
            continue
        if basis not in ('CPM', 'FLT'):
            continue
        tl_raw[orig_city][dest_st].append({
            'mo': 'TL', 'c': carrier, 'b': basis,
            'r': rate_val, 'm': min_cost, 'cy': currency, 'dc': dest_city,
        })

    # ---- LTL rates (CWT / PLT with min_cost as the floor charge) ----
    ws_ltl = wb['LTL CURRENT RATES']
    ltl_rows = list(ws_ltl.iter_rows(min_row=4, values_only=True))
    ltl_raw = defaultdict(lambda: defaultdict(list))
    for r in ltl_rows:
        orig_city = str(r[20] or '').strip().upper()
        if orig_city not in pick_cities:
            continue
        carrier   = str(r[3]  or '').strip().upper()
        basis     = str(r[11] or '').strip().upper()   # CWT, PLT, ECH
        rate_val  = float(r[12]) if r[12] is not None else 0.0
        min_cost  = float(r[13]) if r[13] is not None else 0.0
        currency  = str(r[15] or 'USD').strip().upper()
        dest_city = str(r[27] or '').strip().upper()
        dest_st   = str(r[28] or '').strip().upper()
        if not dest_city or not dest_st:
            continue
        if min_cost <= 0 and rate_val <= 0:
            continue   # no usable rate data
        ltl_raw[orig_city][dest_st].append({
            'mo': 'LTL', 'c': carrier, 'b': basis,
            'r': rate_val, 'm': min_cost, 'cy': currency, 'dc': dest_city,
        })

    # ---- Merge into one lookup: orig → state → {tl: [...], ltl: [...]} ----
    all_origs = set(tl_raw) | set(ltl_raw)
    all_states = set()
    for d in (tl_raw, ltl_raw):
        for v in d.values():
            all_states.update(v.keys())

    result = {}
    for orig in all_origs:
        result[orig] = {}
        states = set(tl_raw.get(orig, {}).keys()) | set(ltl_raw.get(orig, {}).keys())
        for st in states:
            # TL: dedupe by carrier, keep cheapest CPM rate (FLT kept separately)
            tl_entries = tl_raw.get(orig, {}).get(st, [])
            tl_best = {}
            for e in tl_entries:
                k = e['c']
                if k not in tl_best or e['r'] < tl_best[k]['r']:
                    tl_best[k] = e
            # LTL: dedupe by carrier, keep lowest min_cost
            ltl_entries = ltl_raw.get(orig, {}).get(st, [])
            ltl_best = {}
            for e in ltl_entries:
                k = e['c']
                if k not in ltl_best or e['m'] < ltl_best[k]['m']:
                    ltl_best[k] = e
            tl_list  = sorted(tl_best.values(),  key=lambda x: x['r'])[:5]
            ltl_list = sorted(ltl_best.values(), key=lambda x: x['m'])[:5]
            if tl_list or ltl_list:
                result[orig][st] = {'tl': tl_list, 'ltl': ltl_list}

    total_lanes = sum(len(v) for v in result.values())
    tl_origs  = len([o for o in result if any(v['tl']  for v in result[o].values())])
    ltl_origs = len([o for o in result if any(v['ltl'] for v in result[o].values())])
    print(f"  Rates loaded: {total_lanes} origin→state lanes  (TL: {tl_origs} origins, LTL: {ltl_origs} origins)")
    return result


def extract_sheetjs():
    """Return the inlined SheetJS (xlsx) library.

    Order of sources (so the build no longer depends on Map Maker.html being present):
      1. a local cache file,
      2. Map Maker.html (first <script>),
      3. the previously-built DeliveryMap HTML (the SheetJS script before the added CSS).
    The result is cached for next time.
    """
    if os.path.exists(SHEETJS_CACHE):
        with open(SHEETJS_CACHE, "r", encoding="utf-8") as f:
            cached = f.read()
        if "XLSX" in cached[:500] or "SheetJS" in cached[:500]:
            return cached
        os.remove(SHEETJS_CACHE)  # bad cache (e.g. data blob written by mistake) — discard it

    js = None
    if os.path.exists(MAKER_HTML):
        with open(MAKER_HTML, "r", encoding="utf-8") as f:
            m = f.read()
        s = m.index("<script>"); e = m.index("</script>", s)
        js = m[s + len("<script>"):e]
    elif os.path.exists(OUT_HTML):
        with open(OUT_HTML, "r", encoding="utf-8") as f:
            h = f.read()
        marker = "/* ---- added: Load Group buttons + delivery-date slider + POOL color ---- */"
        i = h.find(marker)
        if i != -1:
            style_open = h.rfind("<style>", 0, i)
            end = h.rfind("</script>", 0, style_open)
            start = h.rfind("<script>", 0, end)
            if start != -1 and end != -1:
                js = h[start + len("<script>"):end]

    # Sanity-check: extracted content must look like SheetJS (not the embedded data blob)
    if js and ("XLSX" not in js[:500] and "SheetJS" not in js[:500]):
        js = None

    if not js:
        print("  WARNING: SheetJS not found — in-browser 'Load Data' button will be disabled.")
        print("           (The map still works fully with its embedded data.)")
        print("           To re-enable: place sheetjs.js in this folder and re-run.")
        return ""  # map still works; only in-browser file picker is disabled

    with open(SHEETJS_CACHE, "w", encoding="utf-8") as f:
        f.write(js)
    return js

DEFAULT_PICKS = [
    # ---- Dessert Holdings Operating Facilities (alphabetical) ----
    {"key": "AURORA",          "name": "Aurora, CO",              "lat": 39.7393,  "lng": -104.863,  "match": ["AURORA"],                "group": "dh"},
    {"key": "DELTA",           "name": "Delta, BC",               "lat": 49.0847,  "lng": -123.0587, "match": ["DELTA"],                 "group": "dh"},
    {"key": "HUMBLE",          "name": "Humble, TX",              "lat": 29.9988,  "lng": -95.2622,  "match": ["HUMBLE"],                "group": "dh"},
    {"key": "KENNESAW",        "name": "Kennesaw, GA",            "lat": 34.0236,  "lng": -84.5956,  "match": ["KENNESAW"],              "group": "dh"},
    {"key": "LE_CENTER",       "name": "Le Center, MN",          "lat": 44.3875,  "lng": -93.7302,  "match": ["LE CENTER"],             "group": "dh"},
    {"key": "LONDON",          "name": "London, ON",              "lat": 42.9849,  "lng": -81.2453,  "match": ["LONDON"],                "group": "dh"},
    {"key": "NEWBURYPORT",     "name": "Newburyport, MA",         "lat": 42.8104,  "lng": -70.8893,  "match": ["NEWBURYPORT"],           "group": "dh"},
    {"key": "PEMBROKE",        "name": "Pembroke, NC",            "lat": 34.6802,  "lng": -79.1950,  "match": ["PEMBROKE"],              "group": "dh"},
    {"key": "ST_PAUL",         "name": "St Paul, MN",            "lat": 44.9537,  "lng": -93.0900,  "match": ["ST PAUL", "SAINT PAUL"], "group": "dh"},
    # ---- 3PL Facilities (alphabetical) ----
    {"key": "ATLANTA",         "name": "Atlanta, GA",             "lat": 33.7490,  "lng": -84.3880,  "match": ["ATLANTA"],               "group": "3pl"},
    {"key": "BETHLEHEM",       "name": "Bethlehem, PA",           "lat": 40.5968,  "lng": -75.3762,  "match": ["BETHLEHEM"],             "group": "3pl"},
    {"key": "BOLINGBROOK",     "name": "Bolingbrook, IL",         "lat": 41.6986,  "lng": -88.0684,  "match": ["BOLINGBROOK"],           "group": "3pl"},
    {"key": "BRIGHTON",        "name": "Brighton, CO",            "lat": 39.9727,  "lng": -104.8104, "match": ["BRIGHTON"],              "group": "3pl"},
    {"key": "CALGARY",         "name": "Calgary, AB",             "lat": 51.0447,  "lng": -114.0719, "match": ["CALGARY"],               "group": "3pl"},
    {"key": "CARTHAGE",        "name": "Carthage, MO",            "lat": 37.1786,  "lng": -94.3140,  "match": ["CARTHAGE"],              "group": "3pl"},
    {"key": "FRANKLIN",        "name": "Franklin, IN",            "lat": 39.4806,  "lng": -86.0544,  "match": ["FRANKLIN"],              "group": "3pl"},
    {"key": "GOLDEN_VALLEY",   "name": "Golden Valley, MN",       "lat": 44.9886,  "lng": -93.3694,  "match": ["GOLDEN VALLEY"],         "group": "3pl"},
    {"key": "INGERSOLL",       "name": "Ingersoll, ON",           "lat": 43.0393,  "lng": -80.8847,  "match": ["INGERSOLL", "INGERSOL"], "group": "3pl"},
    {"key": "SMYRNA",          "name": "Smyrna, GA",              "lat": 33.8840,  "lng": -84.5144,  "match": ["SMYRNA"],                "group": "3pl"},
    {"key": "SURREY",          "name": "Surrey, BC",              "lat": 49.1913,  "lng": -122.8490, "match": ["SURREY"],                "group": "3pl"},
]


def match_pick(value):
    if value is None or value == "":
        return None
    s = " ".join(str(value).upper().split())
    for p in DEFAULT_PICKS:
        for kw in p.get("match", [p["key"]]):
            if str(kw).upper() in s:
                return p["key"]
    return None

def iso_date(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return None

def fmt(value, header=None):
    if value is None:
        return ""
    # identifier-style columns: no thousands separators
    if header in ("TMS ID", "Shipment Reference Numbers", "Order Number", "Last Drop Postal Code", "Drop Location Postal Code"):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        if isinstance(value, int):
            return str(value)
        return str(value)
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%b %d, %Y")
        return value.strftime("%b %d, %Y %I:%M %p")
    if isinstance(value, date):
        return value.strftime("%b %d, %Y")
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float):
        return f"{int(value):,}" if value.is_integer() else f"{value:,.2f}"
    if isinstance(value, int):
        return f"{value:,}"
    return str(value)

def to_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None

def open_wb(path):
    tmp = os.path.join(tempfile.gettempdir(), "_dm_build_copy.xlsx")
    shutil.copy2(path, tmp)
    return openpyxl.load_workbook(tmp, data_only=True, read_only=True)

# ---------------------------------------------------------------------------
# EXCLUSION RULES — edit these to control what gets filtered out of the map
# ---------------------------------------------------------------------------

# Carrier SCACs whose loads are excluded entirely
EXCLUDE_SCACS = {"2ACC", "1TOC"}

# Southeastern US ZIP ranges (by state) — used for Walmart exclusion rule 2
_SE_ZIP_RANGES = [
    (35000, 36999),   # AL
    (71600, 72999),   # AR
    (32000, 34999),   # FL
    (30000, 31999),   # GA
    (40000, 42799),   # KY
    (70000, 71599),   # LA
    (38600, 39799),   # MS
    (27000, 28999),   # NC
    (29000, 29999),   # SC
    (37000, 38599),   # TN
    (20100, 24699),   # VA
    (24700, 26899),   # WV
]

def _is_walmart(name):
    """True for US Walmart DCs (excludes Walmart Canada)."""
    up = str(name or "").upper()
    return ("WAL-MART" in up or "WAL MART" in up or "WALMART" in up) and "CANADA" not in up

def _is_southeast_us(zip_str):
    """True if the ZIP code falls in a southeastern US state."""
    try:
        z = int(str(zip_str or "").strip().split("-")[0][:5].zfill(5))
        return any(lo <= z <= hi for lo, hi in _SE_ZIP_RANGES)
    except (ValueError, TypeError):
        return False


def build_payload():
    wb = open_wb(SRC_XLSX)
    ws = wb["Orders"] if "Orders" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])
    H = {str(h).strip().lower() if h is not None else None: i for i, h in enumerate(headers)}

    # col() tries each alias in order and returns the first match.
    ALIASES = {
        "last drop plan date start":  ["last drop plan date start", "leg drop plan start date"],
        "last drop plan date end":    ["last drop plan date end",   "leg drop plan end date"],
        "last drop name":             ["last drop name",            "drop location name"],
        "last drop postal code":      ["last drop postal code",     "drop location postal code"],
        "drop location reference":    ["drop location reference number", "drop location ref", "drop locref"],
        "pallet spaces":              ["pallet spaces",             "shipment pallet spaces"],
        "plannedorno":                ["plannedorno",               "load status"],
        "first pick appt date start": ["first pick appt date start","load pick appt start date"],
        "last drop appt date start":  ["last drop appt date start", "load drop appt start date"],
        "shipment reference numbers": ["shipment reference numbers","order number"],
        "first pick city":            ["first pick city",           "pick location city"],
        "lat":                        ["lat", "latitude"],
        "lng":                        ["lng", "lon", "long", "longitude"],
    }
    def col(*names):
        for name in names:
            for alias in ALIASES.get(name.lower(), [name.lower()]):
                idx = H.get(alias)
                if idx is not None:
                    return idx
        return None

    i_ps    = col("Last Drop Plan Date Start")
    i_pe    = col("Last Drop Plan Date End")
    i_name  = col("Last Drop Name")
    i_zip   = col("Last Drop Postal Code")
    i_wt    = col("Weight (lb)")
    i_pal   = col("Pallet Spaces")
    i_plt   = col("Pallets")
    i_stat  = col("PLANNEDORNO")
    i_pick  = col("First Pick City")
    i_tms   = col("TMS ID")
    i_lg    = col("Load Group")
    i_pkdt  = col("First Pick Appt Date Start")
    i_lat    = col("lat")
    i_lng    = col("lng")
    i_dropref = col("Drop Location Reference")
    i_scac  = col("SCAC", "Carrier SCAC", "Carrier Code", "SCAC Code", "Carrier")
    # which header indices to EXCLUDE from the popup field table (lat/lng only)
    skip_fields = {i_lat, i_lng}

    has_coords = (i_lat is not None and i_lng is not None)
    has_stop_db = os.path.exists(STOP_LOCS)
    if not has_coords:
        if has_stop_db:
            print("  NOTE: no lat/lng columns — geocoding from stop_locations.json (with ZIP fallback)")
        else:
            print("  NOTE: no lat/lng columns — will geocode from ZIP codes using zip_coords.json")
    if i_scac is not None:
        print(f"  NOTE: SCAC column found — will exclude carriers: {', '.join(sorted(EXCLUDE_SCACS))}")
    else:
        print(f"  NOTE: no SCAC column found — SCAC exclusions ({', '.join(sorted(EXCLUDE_SCACS))}) will apply when column is added")

    orders = []
    skipped = 0
    excluded = 0
    for row in rows[1:]:
        if not row or all(c is None or c == "" for c in row):
            continue
        lat = lng = None
        if has_coords:
            try:
                lat = float(row[i_lat]); lng = float(row[i_lng])
            except (TypeError, ValueError):
                pass
        if lat is None or lng is None:
            # Try stop locations DB first (by reference number, then name+postal)
            locref = str(row[i_dropref]).strip() if i_dropref is not None and row[i_dropref] is not None else ""
            name_v = str(row[i_name]).strip() if i_name is not None and row[i_name] is not None else ""
            zip_v  = str(row[i_zip]).strip()  if i_zip  is not None and row[i_zip]  is not None else ""
            coords = _stop_to_coords(locref, name_v, zip_v)
            if coords:
                lat, lng = coords
            else:
                # Final fallback: ZIP geocoding
                coords = _zip_to_coords(zip_v)
                if coords:
                    lat, lng = coords
                else:
                    skipped += 1
                    continue

        # --- Exclusion rules ---
        dest_name = str(row[i_name]).strip() if i_name is not None and row[i_name] is not None else ""
        dest_zip  = str(row[i_zip]).strip()  if i_zip  is not None and row[i_zip]  is not None else ""
        pick_city = str(row[i_pick]).strip().upper() if i_pick is not None and row[i_pick] is not None else ""

        # Rule 1: Atlanta → any US Walmart
        if pick_city == "ATLANTA" and _is_walmart(dest_name):
            excluded += 1; continue

        # Rule 2: Any US Walmart in a southeastern US state (any origin)
        if _is_walmart(dest_name) and _is_southeast_us(dest_zip):
            excluded += 1; continue

        # Rule 3: Carrier SCAC on the exclude list (applies when SCAC column is present)
        if i_scac is not None and row[i_scac] is not None:
            if str(row[i_scac]).strip().upper() in EXCLUDE_SCACS:
                excluded += 1; continue

        status = ""
        if i_stat is not None and row[i_stat] is not None:
            status = str(row[i_stat]).strip().upper()

        # TMS on every row that has one (any status) so shared TMS IDs draw a route line
        tms = None
        if i_tms is not None:
            raw = row[i_tms]
            if raw not in (None, ""):
                tms = str(int(raw)) if isinstance(raw, float) and raw.is_integer() else str(raw).strip()

        load_group = ""
        if i_lg is not None and row[i_lg] is not None:
            load_group = str(row[i_lg]).strip()

        fields = []
        for i, h in enumerate(headers):
            if h is None or i in skip_fields:
                continue
            fields.append([str(h), fmt(row[i], str(h))])

        orders.append({
            "lat": lat, "lng": lng, "status": status,
            "name": str(row[i_name]) if i_name is not None and row[i_name] is not None else "",
            "zip": str(row[i_zip]) if i_zip is not None and row[i_zip] is not None else "",
            "palletSpaces": to_number(row[i_pal]) if i_pal is not None else None,
            "pallets": to_number(row[i_plt]) if i_plt is not None else None,
            "weight": to_number(row[i_wt]) if i_wt is not None else None,
            "tms": tms,
            "pick": match_pick(row[i_pick]) if i_pick is not None else None,
            "loadGroup": load_group,
            "planStart": iso_date(row[i_ps]) if i_ps is not None else None,
            "planEnd": iso_date(row[i_pe]) if i_pe is not None else None,
            "pickDate": iso_date(row[i_pkdt]) if i_pkdt is not None else None,
            "fields": fields,
        })

    mtime = os.path.getmtime(SRC_XLSX)
    rates = build_rates_lookup()
    payload = {
        "updatedAt": datetime.fromtimestamp(mtime).strftime("%b %d, %Y %I:%M:%S %p"),
        "stale": False,
        "picks": [{"key": p["key"], "name": p["name"], "lat": p["lat"], "lng": p["lng"], "group": p.get("group", "dh")} for p in DEFAULT_PICKS],
        "palletThreshold": 16,
        "pollSeconds": 45,
        "tmsAvailable": i_tms is not None,
        "pickAvailable": i_pick is not None,
        "orders": orders,
        "rates": rates,
    }
    print(f"  built {len(orders)} orders ({skipped} skipped for missing coords, {excluded} excluded by rules)")
    return payload

# ---------------------------------------------------------------- HTML patches
NEW_CSS = """<style>
  /* ---- added: Load Group buttons + delivery-date slider + POOL color ---- */
  .lg-tabs { display:flex; flex-wrap:wrap; gap:6px; }
  .lg-tabs button.lgbtn { flex:0 1 auto; padding:6px 9px; border:1px solid var(--line); background:#fff;
    border-radius:8px; font-size:11.5px; cursor:pointer; color:var(--muted); transition:all .12s;
    display:inline-flex; align-items:center; gap:6px; }
  .lg-tabs button.lgbtn:hover { border-color:#c3ccd6; }
  .lg-tabs button.lgbtn.active { background:var(--ink); color:#fff; border-color:var(--ink); }
  .lg-tabs .pcount { font-size:10.5px; color:var(--muted); background:var(--bg); border-radius:10px; padding:1px 6px; }
  .lg-tabs button.lgbtn.active .pcount { background:rgba(255,255,255,.22); color:#fff; }

  .daterange { padding-top:2px; }
  .dr-labels { display:flex; justify-content:space-between; font-size:12.5px; font-weight:700; color:var(--ink); margin-bottom:8px; }
  .dr-slider { position:relative; height:26px; }
  .dr-track { position:absolute; left:0; right:0; top:11px; height:4px; background:var(--line); border-radius:2px; }
  .dr-fill  { position:absolute; top:11px; height:4px; background:var(--accent); border-radius:2px; }
  .dr-slider input[type=range]{ position:absolute; left:0; top:0; width:100%; height:26px; margin:0;
    background:none; -webkit-appearance:none; appearance:none; pointer-events:none; }
  .dr-slider input[type=range]:focus{ outline:none; }
  .dr-slider input[type=range]::-webkit-slider-runnable-track{ background:none; height:26px; }
  .dr-slider input[type=range]::-moz-range-track{ background:none; }
  .dr-slider input[type=range]::-webkit-slider-thumb{ -webkit-appearance:none; appearance:none; pointer-events:auto;
    margin-top:5px; width:16px; height:16px; border-radius:50%; background:var(--accent); border:2px solid #fff;
    box-shadow:0 1px 3px rgba(0,0,0,.3); cursor:pointer; }
  .dr-slider input[type=range]::-moz-range-thumb{ pointer-events:auto; width:16px; height:16px; border-radius:50%;
    background:var(--accent); border:2px solid #fff; box-shadow:0 1px 3px rgba(0,0,0,.3); cursor:pointer; }
  .dr-note { font-size:11.5px; color:#9aa6b2; margin-top:8px; line-height:1.4; }
  .badge.pool { background:var(--pool); }
  .badge2.ACCEPTED,.badge2.accepted { background:var(--accepted); }
  .badge2.ACTIVE,.badge2.active     { background:var(--active); }
  .badge2.TENDERED,.badge2.tendered { background:var(--tendered); }
  .badge2.PENDING,.badge2.pending   { background:var(--pending); }
  .badge2.POOL,.badge2.pool         { background:var(--pool); }
  .consol-origin { margin:12px 0; padding-bottom:10px; border-bottom:1px dashed #eef1f5; }
  .consol-lbl { font-weight:600; color:var(--ink); font-size:13px; margin-bottom:6px; }
  .consol-list { display:flex; flex-wrap:wrap; gap:10px; margin-bottom:6px; }
  .consol-box { position:relative; width:210px; border:1px solid var(--line); border-radius:9px; padding:9px; background:#fafbfc; }
  .consol-box .consol-name { width:100%; padding:6px 8px; border:1px solid var(--line); border-radius:6px; font-size:13px; margin-bottom:5px; }
  .consol-box .consol-zip { width:100%; padding:5px 8px; border:1px solid var(--line); border-radius:6px; font-size:12.5px; color:var(--muted); }
  .consol-box .consol-del { position:absolute; top:1px; right:5px; border:none; background:none; color:var(--muted); font-size:16px; line-height:1; cursor:pointer; }
  .consol-box .consol-del:hover { color:var(--red); }
  .consol-add { margin-top:2px; }
  .loadrow .lname { color:var(--ink); font-weight:600; }
  .topdate { margin-left:auto; font-size:12px; color:var(--muted); white-space:nowrap; }
  .topdate b { color:var(--ink); font-weight:700; }
  #loads-controls .lf-row { flex-basis:100%; display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-top:6px; padding-top:8px; border-top:1px solid var(--line); }
  #loads-controls .lf-in { padding:6px 9px; border:1px solid var(--line); border-radius:7px; font-size:13px; }
  #loads-controls #lf-search { width:240px; }
  #loads-controls .lf-sm { width:84px; }
  #loads-controls .lf-xs { width:64px; }
  table.lt th.lh { cursor:pointer; user-select:none; white-space:nowrap; }
  table.lt th.lh:hover { background:#e8eef7; }
  .dc-filter { display:flex; flex-wrap:wrap; gap:6px; margin:8px 0 2px; }
  .dcf-btn { padding:5px 12px; border:1px solid var(--line); background:#fff; border-radius:16px; font-size:12.5px; cursor:pointer; color:var(--ink); }
  .dcf-btn:hover { background:#f0f4f9; }
  .dcf-btn.on { background:var(--accent); border-color:var(--accent); color:#fff; font-weight:600; }
  .loadrow .tms-chip { color:var(--accent); font-weight:600; cursor:pointer; white-space:nowrap; }
  .loadrow .tms-chip:hover { text-decoration:underline; }
  .loadrow.tms-hl { background:#fff3cd; box-shadow:inset 3px 0 0 #f0a000; border-radius:4px; }
  .tms-banner { margin:8px 0 2px; padding:7px 12px; background:#fff8e1; border:1px solid #f0c000; border-left:4px solid #f0a000; border-radius:6px; font-size:13px; color:#5a4a00; }
  .tms-banner .tms-clear { margin-left:10px; color:var(--accent); cursor:pointer; font-weight:600; white-space:nowrap; }
  .tms-banner .tms-clear:hover { text-decoration:underline; }
  .tms-banner .tms-findmatch { margin-left:10px; color:#92400e; background:#fef3c7; border:1px solid #f59e0b; border-radius:5px; padding:1px 8px; cursor:pointer; font-weight:600; white-space:nowrap; font-size:12px; }
  .tms-banner .tms-findmatch:hover { background:#fde68a; }
  .tms-banner .tms-findmatch.active { background:#f59e0b; color:#fff; border-color:#d97706; }
  .match-results { margin-top:8px; padding-top:8px; border-top:1px dashed #e0b000; }
  .match-results table { width:100%; font-size:12px; border-collapse:collapse; }
  .match-results th { color:#78350f; font-weight:600; font-size:11px; padding:2px 4px; text-align:left; }
  .match-results td { padding:3px 4px; border-top:1px solid #fde68a; vertical-align:middle; }
  .match-results .fill-bar { display:inline-block; height:8px; border-radius:4px; background:var(--accent); }
  .match-map-ring { pointer-events:none; }
  /* ---- load row selection + export ---- */
  .loadrow input.load-chk { cursor:pointer; width:14px; height:14px; margin:0; flex-shrink:0; accent-color:var(--accent); }
  .loadrow.sel-row { background:#eff6ff; box-shadow:inset 3px 0 0 var(--accent); border-radius:4px; }
  .export-bar { position:sticky; bottom:0; background:#1e3a5f; color:#fff; padding:9px 16px; display:none; align-items:center; gap:12px; border-top:2px solid var(--accent); z-index:60; font-size:13px; }
  .export-bar.show { display:flex; }
  .export-bar .exp-count { font-weight:700; }
  .export-bar .exp-btn { padding:6px 14px; background:#fff; color:var(--accent); border:none; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; }
  .export-bar .exp-btn:hover { background:#dbeafe; }
  .export-bar .exp-clear { background:transparent; color:#93c5fd; font-size:12px; border:1px solid #93c5fd; border-radius:6px; padding:4px 10px; cursor:pointer; }
  .export-bar .exp-clear:hover { background:rgba(255,255,255,.1); }
  .tms-banner .tms-warn { margin-top:5px; padding-top:5px; border-top:1px dashed #e0b000; color:#9a3b00; font-size:12.5px; line-height:1.7; }
  .tms-banner .tms-warn .wgrp { white-space:nowrap; background:#ffe0cc; border-radius:3px; padding:1px 5px; }
  .tms-banner .tms-warn i { font-style:normal; font-weight:700; margin:0 4px; }
  .tms-banner .tms-locs { margin-top:5px; padding-top:5px; border-top:1px dashed #e0b000; }
  .tms-banner .tms-locs .locrow { font-size:12.5px; color:#5a4a00; line-height:1.8; }
  /* rate badge in consolidation panel */
  .rate-badge { display:inline-block; background:#f0fdf4; border:1px solid #bbf7d0; color:#15803d;
    border-radius:5px; padding:2px 7px; font-size:11.5px; font-weight:600; cursor:help; white-space:nowrap; }
  /* status filter buttons: wrap on narrow panels */
  .filters { flex-wrap:wrap !important; }
  .filters button { flex:0 1 auto; min-width:60px; }
  /* pick-location dropdown */
  .pick-dropdown { position:relative; display:inline-block; max-width:100%; }
  .pick-dd-btn { display:flex; align-items:center; justify-content:space-between; gap:8px; padding:7px 12px; background:#fff; border:1px solid var(--line); border-radius:8px; font-size:13px; cursor:pointer; color:var(--ink); min-width:200px; max-width:100%; white-space:nowrap; }
  .pick-dd-btn:hover { border-color:var(--accent); }
  .pick-dd-btn.open { border-color:var(--accent); box-shadow:0 0 0 2px rgba(59,130,246,.15); }
  .dd-arrow { font-size:10px; color:var(--muted); margin-left:4px; }
  .pick-dd-panel { display:none; position:absolute; top:calc(100% + 4px); left:0; z-index:2000; background:#fff; border:1px solid var(--line); border-radius:10px; box-shadow:0 6px 20px rgba(0,0,0,.13); min-width:250px; max-height:400px; overflow-y:auto; padding:6px 0; }
  .pick-dd-item { display:flex; align-items:center; gap:7px; padding:6px 14px; cursor:pointer; user-select:none; font-size:13px; }
  .pick-dd-item:hover { background:#f0f4f9; }
  .pick-dd-item input[type=checkbox] { flex-shrink:0; cursor:pointer; width:14px; height:14px; }
  .pick-dd-all { font-weight:600; border-bottom:1px solid var(--line); margin-bottom:2px; padding-bottom:8px; }
  .pick-group-label { padding:8px 14px 3px; font-size:10.5px; font-weight:700; letter-spacing:.05em; text-transform:uppercase; color:var(--muted); }
</style>
</head>"""

PATCHES = [
    # --- CSS: POOL color variable + dot/popup-header colors ---
    ("root pool var",
     "--green:#2e7d32; --red:#c62828; --orange:#ef6c00;",
     "--green:#2e7d32; --red:#c62828; --orange:#ef6c00; --pool:#1565c0; --accepted:#2e7d32; --active:#1565c0; --tendered:#ef6c00; --pending:#c62828;"),
    ("legend dot pool",
     ".dot.green{background:var(--green)} .dot.red{background:var(--red)} .dot.orange{background:var(--orange)}",
     ".dot.green{background:var(--green)} .dot.red{background:var(--red)} .dot.orange{background:var(--orange)} .dot.pool{background:var(--pool)} .dot.accepted{background:var(--accepted)} .dot.active{background:var(--active)} .dot.tendered{background:var(--tendered)} .dot.pending{background:var(--pending)}"),
    ("popup head pool",
     ".pop-head.green{background:var(--green)} .pop-head.red{background:var(--red)} .pop-head.orange{background:var(--orange)}",
     ".pop-head.green{background:var(--green)} .pop-head.red{background:var(--red)} .pop-head.orange{background:var(--orange)} .pop-head.pool{background:var(--pool)} .pop-head.accepted{background:var(--accepted)} .pop-head.active{background:var(--active)} .pop-head.tendered{background:var(--tendered)} .pop-head.pending{background:var(--pending)}"),

    # --- Map: replace single OSM tile layer with layer control (OSM + Google Satellite + Hybrid) ---
    ('tile layer control',
     '''L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom:19, attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
}).addTo(map);''',
     '''const _osmLayer = L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  maxZoom:19, attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
});
const _gSat = L.tileLayer("https://mt{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}", {
  subdomains:["0","1","2","3"], maxZoom:21,
  attribution:'&copy; <a href="https://maps.google.com">Google</a>'
});
const _gHybrid = L.tileLayer("https://mt{s}.google.com/vt/lyrs=y&x={x}&y={y}&z={z}", {
  subdomains:["0","1","2","3"], maxZoom:21,
  attribution:'&copy; <a href="https://maps.google.com">Google</a>'
});
_osmLayer.addTo(map);
L.control.layers({"Street map (OSM)":_osmLayer,"Satellite (Google)":_gSat,"Hybrid (Google)":_gHybrid},{},{position:"topright",collapsed:false}).addTo(map);'''),

    # (CSS + SheetJS are injected at </head> in main(), not here)

    # --- body: replace picktabs with a dropdown ---
    ("picktabs dropdown",
     """  <div class="section">
    <h2>Pick Location (ship-from)</h2>
    <div id="picktabs"></div>
  </div>""",
     """  <div class="section">
    <h2>Pick Location (ship-from)</h2>
    <div class="pick-dropdown" id="pick-dropdown">
      <button class="pick-dd-btn" id="pick-dd-btn" type="button">Select pick locations <span class="dd-arrow">&#9660;</span></button>
      <div class="pick-dd-panel" id="pick-dd-panel"></div>
    </div>
  </div>"""),

    # --- body: status filters get id + POOL button; add Load Group + Date sections ---
    ("search + filters block",
     """  <div class="section">
    <input id="search" type="search" placeholder="Search name, ZIP, reference #, TMS..." autocomplete="off" />
    <div class="filters">
      <button data-f="ALL" class="active">All</button>
      <button data-f="PLANNED">Planned</button>
      <button data-f="UNPLANNED">Unplanned</button>
    </div>
  </div>""",
     """  <div class="section">
    <input id="search" type="search" placeholder="Search name, ZIP, reference #, TMS, load group..." autocomplete="off" />
    <div class="filters" id="status-filters">
      <button data-f="ALL" class="active">All</button>
      <button data-f="ACCEPTED">Accepted</button>
      <button data-f="ACTIVE">Active</button>
      <button data-f="TENDERED">Tendered</button>
      <button data-f="PENDING">Pending</button>
      <button data-f="PLANNED">Planned</button>
      <button data-f="UNPLANNED">Unplanned</button>
      <button data-f="POOL">Pool</button>
    </div>
  </div>

  <div class="section">
    <h2>Load Group</h2>
    <div class="lg-tabs" id="lg-tabs"></div>
  </div>

  <div class="section">
    <h2>Last Drop Plan Date</h2>
    <div class="daterange">
      <div class="dr-labels"><span id="dr-from">&mdash;</span><span id="dr-to">&mdash;</span></div>
      <div class="dr-slider">
        <div class="dr-track"></div>
        <div class="dr-fill" id="dr-fill"></div>
        <input type="range" id="dr-min" min="0" max="100" value="0" step="1" aria-label="From date">
        <input type="range" id="dr-max" min="0" max="100" value="100" step="1" aria-label="To date">
      </div>
      <div class="dr-note">Drag the handles to show only loads whose Last Drop Plan Date Start falls in that window. Loads with no date stay visible.</div>
    </div>
  </div>"""),

    # --- legend: add POOL row ---
    ("legend pool row",
     '      <div class="row"><span class="dot orange"></span> Mixed</div>',
     '      <div class="row"><span class="dot orange"></span> Mixed</div>\n      <div class="row"><span class="dot accepted"></span> Accepted</div>\n      <div class="row"><span class="dot active"></span> Active</div>\n      <div class="row"><span class="dot tendered"></span> Tendered</div>\n      <div class="row"><span class="dot pending"></span> Pending</div>\n      <div class="row"><span class="dot pool"></span> Pool</div>'),

    # --- JS: COLORS add pool ---
    ('COLORS pool',
     'const COLORS = { green:"#2e7d32", red:"#c62828", orange:"#ef6c00" };',
     'const COLORS = { green:"#2e7d32", red:"#c62828", orange:"#ef6c00", pool:"#1565c0", accepted:"#2e7d32", active:"#1565c0", tendered:"#ef6c00", pending:"#c62828" };'),

    # --- JS: state vars ---
    ('state vars',
     'let activePick = "ALL", activeStatus = "ALL", query = "", smallOnly = false, threshold = 16;',
     'let activePick = new Set(), activeStatus = "ALL", query = "", smallOnly = false, threshold = 16;\n'  # activePick: null=all, Set=specific keys (empty Set = nothing selected)
     'let activeLoadGroup = "ALL", loadGroups = [], lgCounts = {};\n'
     'let dateMin = null, dateMax = null, dateFrom = null, dateTo = null, sliderInit = false;\n'
     'let loadsSortKey = "", loadsSortDir = 1, loadsDC = "";\n'
     'let selectedTms = new Set();'),

    # --- JS: colorKeyForOrders handles POOL ---
    ('colorKeyForOrders',
     '''function colorKeyForOrders(list){
  const hasP = list.some(o => o.status==="PLANNED");
  const hasU = list.some(o => o.status==="UNPLANNED");
  if (hasP && hasU) return "orange";
  if (hasP) return "green";
  return "red";
}''',
     '''const STATUS_COLOR = { PLANNED:"green", UNPLANNED:"red", POOL:"pool", ACCEPTED:"accepted", ACTIVE:"active", TENDERED:"tendered", PENDING:"pending" };
function colorKeyForOrders(list){
  const set = new Set(list.map(o => o.status).filter(Boolean));
  if (set.size === 1){ const s=[...set][0]; return STATUS_COLOR[s] || "red"; }
  return set.size === 0 ? "red" : "orange";
}'''),

    # --- JS: orderVisible adds load group + date window ---
    ('orderVisible',
     '''function orderVisible(o){
  if (activePick !== "ALL" && o.pick !== activePick) return false;
  if (activeStatus !== "ALL" && o.status !== activeStatus) return false;
  if (smallOnly && !(o.palletSpaces!==null && o.palletSpaces < threshold)) return false;
  return true;
}''',
     '''function orderVisible(o){
  if (activePick !== null && !activePick.has(o.pick||"")) return false;
  if (activeStatus !== "ALL" && o.status !== activeStatus) return false;
  if (activeLoadGroup !== "ALL" && (o.loadGroup||"") !== activeLoadGroup) return false;
  if (smallOnly && !(o.palletSpaces!==null && o.palletSpaces < threshold)) return false;
  if (dateFrom !== null){
    const d = dayNum(o.planStart);
    if (d!=null && (d < dateFrom || d > dateTo)) return false;
  }
  return true;
}'''),

    # --- JS: buildTabs becomes a multi-select pick filter (only show DCs with orders, show count) ---
    ('buildTabs multiselect',
     '''function buildTabs(){
  const wrap = document.getElementById("picktabs");
  wrap.innerHTML = "";
  const tabs = [{key:"ALL", name:"All"}].concat(picks.map(p => ({key:p.key, name:p.name.split(",")[0]})));
  for (const t of tabs){
    const btn = document.createElement("button");
    btn.className = "ptab" + (t.key===activePick ? " active":"");
    const col = t.key==="ALL" ? "#334155" : pickColor(t.key);
    if (t.key===activePick) btn.style.background = col;
    const cnt = pickCounts[t.key]!==undefined ? pickCounts[t.key] : 0;
    btn.innerHTML = (t.key==="ALL" ? "" : '<span class="pdot" style="background:'+col+'"></span>') +
                    esc(t.name) + ' <span class="pcount">'+cnt+'</span>';
    btn.addEventListener("click", () => { activePick = t.key; buildTabs(); applyFilters(true); });
    wrap.appendChild(btn);
  }
}''',
     '''function buildTabs(){
  const btn = document.getElementById("pick-dd-btn");
  const panel = document.getElementById("pick-dd-panel");
  if (!btn || !panel) return;

  const withOrders = picks.filter(x=>(pickCounts[x.key]||0)>0).map(x=>x.key);

  // Wire open/close once
  if (!btn._ddWired){
    btn._ddWired = true;
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const open = panel.style.display === "block";
      panel.style.display = open ? "none" : "block";
      btn.classList.toggle("open", !open);
    });
    document.addEventListener("click", e => {
      if (!e.target.closest("#pick-dropdown")){
        panel.style.display = "none"; btn.classList.remove("open");
      }
    });
  }

  // Update button label
  if (activePick === null){
    btn.innerHTML = 'All pick locations <span class="dd-arrow">&#9660;</span>';
  } else if (activePick.size === 0){
    btn.innerHTML = 'Select pick locations <span class="dd-arrow">&#9660;</span>';
  } else {
    const names = [...activePick].map(k=>(pickByKey.get(k)||{name:k}).name.split(",")[0]);
    const label = names.length <= 2 ? names.join(", ") : names.length + " locations selected";
    btn.innerHTML = esc(label) + ' <span class="dd-arrow">&#9660;</span>';
  }

  // Toggle logic
  function _togglePick(key){
    if (activePick === null){
      activePick = new Set(withOrders.filter(k=>k!==key));
    } else if (activePick.has(key)){
      activePick = new Set([...activePick].filter(k=>k!==key));
    } else {
      activePick = new Set([...activePick, key]);
      if (withOrders.every(k=>activePick.has(k))) activePick = null;
    }
    buildTabs(); applyFilters(true);
  }

  // Rebuild panel
  panel.innerHTML = "";

  // "All" row
  const allItem = document.createElement("label");
  allItem.className = "pick-dd-item pick-dd-all";
  const allChk = document.createElement("input");
  allChk.type = "checkbox";
  allChk.checked = (activePick === null);
  allChk.indeterminate = (activePick !== null && activePick.size > 0);
  allChk.addEventListener("change", () => {
    activePick = allChk.checked ? null : new Set();
    buildTabs(); applyFilters(true);
  });
  const allLbl = document.createElement("span"); allLbl.textContent = "All locations";
  allItem.appendChild(allChk); allItem.appendChild(allLbl);
  panel.appendChild(allItem);

  // One labelled group of checkboxes
  function renderGroup(label, groupPicks){
    if (!groupPicks.length) return;
    const hdr = document.createElement("div");
    hdr.className = "pick-group-label"; hdr.textContent = label;
    panel.appendChild(hdr);
    for (const p of groupPicks){
      const cnt = pickCounts[p.key] || 0;
      const on = (activePick === null) || activePick.has(p.key);
      const item = document.createElement("label");
      item.className = "pick-dd-item";
      const chk = document.createElement("input"); chk.type = "checkbox"; chk.checked = on;
      const dot = document.createElement("span");
      dot.className = "pdot"; dot.style.cssText = "background:"+pickColor(p.key)+";flex-shrink:0";
      const nm = document.createElement("span"); nm.textContent = p.name.split(",")[0];
      const ct = document.createElement("span"); ct.className = "pcount"; ct.textContent = cnt;
      item.appendChild(chk); item.appendChild(dot); item.appendChild(nm); item.appendChild(ct);
      const key = p.key;
      chk.addEventListener("change", ()=>_togglePick(key));
      panel.appendChild(item);
    }
  }

  renderGroup("Dessert Holdings Operating Facilities",
    picks.filter(p=>(p.group==="dh"||!p.group)&&(pickCounts[p.key]||0)>0));
  renderGroup("3PL Facilities",
    picks.filter(p=>p.group==="3pl"&&(pickCounts[p.key]||0)>0));
}'''),

    # --- JS: result count supports multi-select pick ---
    ('result count pick label',
     'layers.length+" of "+markerRecs.length+" drop locations" + (activePick!=="ALL"?" · "+(pickByKey.get(activePick)||{}).name:"");',
     '''layers.length+" of "+markerRecs.length+" drop locations" +
    (activePick!==null && activePick.size>0 ? " · " + [...activePick].map(k=>(pickByKey.get(k)||{}).name||k).join(", ") : "");'''),

    # --- JS: drawOriginMarkers supports multi-select pick + click-to-filter ---
    ('drawOriginMarkers multiselect',
     '''function drawOriginMarkers(keys){
  originMarkers.clearLayers();
  const show = (activePick==="ALL") ? keys : new Set([activePick]);
  show.forEach(k => {
    const p = pickByKey.get(k); if (!p) return;
    const mk = L.marker([p.lat,p.lng], { icon:originPinIcon(pickColor(k)), zIndexOffset:1000 });
    mk.bindPopup('<div class="pop"><div class="pop-head" style="background:'+pickColor(k)+'">'+
      '<div class="pop-title">'+esc(p.name)+'</div><div class="pop-meta">Pick location &middot; ships '+
      (pickCounts[k]||0)+' orders</div></div></div>', {maxWidth:300});
    mk.bindTooltip(p.name, {direction:"top"});
    originMarkers.addLayer(mk);
  });
}''',
     '''function drawOriginMarkers(keys){
  originMarkers.clearLayers();
  const show = (activePick === null) ? keys : new Set([...activePick].filter(k => keys.has(k)));
  keys.forEach(k => {
    const p = pickByKey.get(k); if (!p) return;
    const isActive = activePick === null || activePick.has(k);
    const col = pickColor(k);
    const mk = L.marker([p.lat,p.lng], { icon:originPinIcon(col), zIndexOffset:1000, opacity: isActive ? 1 : 0.35 });
    const cnt = pickCounts[k]||0;
    mk.bindPopup(
      '<div class="pop"><div class="pop-head" style="background:'+col+'">' +
      '<div class="pop-title">'+esc(p.name)+'</div>' +
      '<div class="pop-meta">Pick location &middot; ships '+cnt+' orders</div></div>' +
      '<div style="padding:10px 14px 6px">' +
      '<button onclick="javascript:_pickPinClick(this.dataset.k)" data-k="'+k+'" style="width:100%;padding:7px;background:'+col+';color:#fff;border:none;border-radius:7px;cursor:pointer;font-weight:600;font-size:13px">' +
      (isActive && activePick !== null ? '&#10005; Deselect location' : '&#10003; Filter to this location') +
      '</button></div></div>',
      {maxWidth:300}
    );
    mk.bindTooltip(p.name + ' — click to filter', {direction:"top"});
    originMarkers.addLayer(mk);
  });
}
function _pickPinClick(k){
  if (activePick === null){
    // currently showing all — select only this one
    activePick = new Set([k]);
  } else if (activePick.has(k) && activePick.size === 1){
    // only one selected and it's this one — deselect (show all)
    activePick = null;
  } else {
    // toggle this key in the set
    activePick = new Set([...activePick]);
    activePick.has(k) ? activePick.delete(k) : activePick.add(k);
    if (activePick.size === 0) activePick = new Set(); // keep as "nothing"
  }
  buildTabs(); applyFilters(true);
}'''),

    # --- JS: searchText include loadGroup ---
    ('searchText loadGroup',
     '[o.name, o.zip, o.tms||"", o.pick||"", ...o.fields.filter(f=>/reference/i.test(f[0])).map(f=>f[1])].join(" ")',
     '[o.name, o.zip, o.tms||"", o.pick||"", o.loadGroup||"", ...o.fields.filter(f=>/reference/i.test(f[0])).map(f=>f[1])].join(" ")'),

    # --- JS: remove duplicate "Load (TMS)" popup row (TMS ID now shown in columns) ---
    ('remove dup tms row',
     'if (o.tms) html += \'<tr><td class="k">Load (TMS)</td><td class="v">\'+esc(o.tms)+\'</td></tr>\';',
     '/* TMS ID is shown above as one of the order columns */'),

    # --- JS: status filter handler scoped to #status-filters (so Load Group buttons don't collide) ---
    ('status handler scope',
     '''document.querySelectorAll(".filters button").forEach(btn => btn.addEventListener("click", () => {
  document.querySelectorAll(".filters button").forEach(b=>b.classList.remove("active"));
  btn.classList.add("active"); activeStatus=btn.dataset.f; applyFilters(true);
}));''',
     '''document.querySelectorAll("#status-filters button").forEach(btn => btn.addEventListener("click", () => {
  document.querySelectorAll("#status-filters button").forEach(b=>b.classList.remove("active"));
  btn.classList.add("active"); activeStatus=btn.dataset.f; applyFilters(true);
}));'''),

    # --- JS: collapse same-TMS + same-name orders into one combined stop (run first in ingest) ---
    ('ingest merge stops',
     '''function ingest(data){
  DATA = data;
  threshold = (document.getElementById("thresh").value|0) || data.palletThreshold || 16;''',
     '''function ingest(data){
  DATA = data;
  if(!data._raw) data._raw = data.orders;        // keep the un-merged rows so re-grouping can re-run
  data.orders = mergeStops(data._raw);
  threshold = (document.getElementById("thresh").value|0) || data.palletThreshold || 16;'''),

    # --- JS: per-order status badge handles POOL ---
    ('badge pool color',
     'const badge = o.status==="PLANNED" ? "green" : (o.status==="UNPLANNED" ? "red" : "");',
     'const badge = STATUS_COLOR[o.status] || "";'),

    # --- JS: show how many shipments were combined into a grouped order ---
    ('combined count badge',
     "if (pickName) html += '<span class=\"badge pick\">'+esc(pickName)+'</span>';",
     "if (pickName) html += '<span class=\"badge pick\">'+esc(pickName)+'</span>';\n    if (o._count>1) html += '<span class=\"badge\" style=\"background:#475569\">'+o._count+' shipments combined</span>';"),

    # --- Loads by DC: orders sharing a TMS ID are ALREADY one consolidated load (don't offer to re-combine) ---
    ('consolidation count by TMS',
     '''  const fits = weight<maxW && pallets<maxP;
  const multi = orders.length>=2;''',
     '''  const fits = weight<maxW && pallets<maxP;
  const loadKeys = new Set(orders.map((o,i)=> o.tms ? ("T:"+o.tms) : ("U:"+i)));   // same TMS = already one load
  const nLoads = loadKeys.size;
  const multi = nLoads>=2;'''),
    ('consolidation single-load text',
     'if(!multi) statusText="single load";',
     'if(!multi) statusText = orders.some(o=>o.tms) ? "already 1 load" : "single load";'),
    ('consolidation loads count',
     'n:orders.length, weight:weight,',
     'n:nLoads, weight:weight,'),

    # --- Settings: clarify the negative-End rule in the delivery-window note ---
    ('delivery window note',
     'Shift the start and/or end by <b>+</b> or <b>&minus;</b> days below (e.g. start &minus;1, end +2). A load never ships before its pick date.',
     'Shift the start and/or end by <b>+</b> or <b>&minus;</b> days below. A <b>negative End</b> (e.g. end &minus;5) means the customer can take delivery that many days <b>before the End date</b> &mdash; the End stays the latest and the Start date is ignored. A load never ships before its pick date.'),

    # --- Settings: add a "same-location customers" (aliases) list ---
    ('default settings aliases',
     '''  customers:[{match:"AMAZON",startAdj:0,endAdj:0,biz:false},{match:"MCLANE",startAdj:0,endAdj:0,biz:false}] };''',
     '''  customers:[{match:"AMAZON",startAdj:0,endAdj:0,biz:false},{match:"MCLANE",startAdj:0,endAdj:0,biz:false}], aliases:[], consolidated:{}, sameLoc:[] };'''),
    ('normSettings aliases',
     '''    customers:Array.isArray(p.customers)?p.customers:JSON.parse(JSON.stringify(DEFAULT_SETTINGS.customers)) };''',
     '''    customers:Array.isArray(p.customers)?p.customers:JSON.parse(JSON.stringify(DEFAULT_SETTINGS.customers)),
    aliases:Array.isArray(p.aliases)?p.aliases:[],
    consolidated:(p.consolidated && typeof p.consolidated==="object")?p.consolidated:{},
    sameLoc:Array.isArray(p.sameLoc)?p.sameLoc:[] };'''),

    # --- Settings page: add the Same-location customers section after Per-customer overrides ---
    ('settings aliases section',
     '''      <div id="cust-rows"></div>
      <button id="cust-add" class="cp-btn">+ Add customer</button>
    </div>''',
     '''      <div id="cust-rows"></div>
      <button id="cust-add" class="cp-btn">+ Add customer</button>
    </div>
    <div class="set-sec">
      <h3>Same&#8209;location customers</h3>
      <div class="set-note">Tell the map which customer names are the <b>same delivery location</b>. Put names (or parts of names) that mean the same place in one row, separated by commas &mdash; loads on the same TMS that match any name in the row are treated as one customer when grouping. Example: <b>WAL-MART, WAL MART</b>.</div>
      <div id="alias-rows"></div>
      <button id="alias-add" class="cp-btn">+ Add same&#8209;location group</button>
    </div>
    <div class="set-sec">
      <h3>Same location (merge ZIPs)</h3>
      <div class="set-note">For a delivery point that shows up under <b>more than one ZIP</b> (same name, different ZIPs &mdash; e.g. <b>ADUSA DISTRIBUTION, LLC</b> at 18017 and 18020), add the name and each ZIP here. On the <b>Loads by DC</b> load summary those will be counted and totaled as <b>one location</b>. Leave ZIPs blank to merge <b>all</b> ZIPs for that name.</div>
      <div id="sameloc-rows"></div>
      <button id="sameloc-add" class="cp-btn">+ Add same&#8209;location group</button>
    </div>
    <div class="set-sec">
      <h3>Consolidated Locations</h3>
      <div class="set-note">For each ship&#8209;from origin, add delivery locations that consolidate well out of that shipper &mdash; one per box (location name on top, ZIP underneath), and use <b>+ Add location</b> for more. Locations under the same origin are treated as a lane that can ship together. In <b>Loads by DC</b>, set <b>Same destination = good consolidation lanes</b> to group them &mdash; truck limits and delivery&#8209;window rules still apply.</div>
      <div id="consol-rows"></div>
    </div>'''),

    # --- renderSettings: render the alias rows ---
    ('renderSettings aliases',
     '''    c.appendChild(row);
  });
}''',
     '''    c.appendChild(row);
  });
  const ar=document.getElementById("alias-rows");
  if(ar){ ar.innerHTML="";
    (SETTINGS.aliases||[]).forEach((val,i)=>{
      const row=document.createElement("div"); row.className="win-row";
      row.innerHTML='<input class="al-terms cu-match" style="width:320px" data-i="'+i+'" value="'+esc(val||"")+'" placeholder="e.g. WAL-MART, WAL MART"> <button class="al-del cu-del" data-i="'+i+'" title="Remove">&times;</button>';
      ar.appendChild(row);
    });
  }
  const cl=document.getElementById("consol-rows");
  if(cl){ cl.innerHTML="";
    const ORIG=(typeof picks!=="undefined" && picks.length) ? picks : [{key:"JONESTOWN",name:"Jonestown, PA"},{key:"GOODYEAR",name:"Goodyear, AZ"},{key:"DEKALB",name:"DeKalb, IL"},{key:"GRAND PRAIRIE",name:"Grand Prairie, TX"},{key:"MCDONOUGH",name:"McDonough, GA"}];
    SETTINGS.consolidated=SETTINGS.consolidated||{};
    ORIG.forEach(p=>{
      let arr=SETTINGS.consolidated[p.key];
      if(typeof arr==="string") arr=arr.split(/[\\n,]/).map(s=>({name:s.trim(),zip:""})).filter(e=>e.name);  // migrate old text
      if(!Array.isArray(arr)) arr=[];
      arr=arr.map(e=> (typeof e==="string") ? {name:e,zip:""} : {name:(e&&e.name)||"",zip:(e&&e.zip)||""});  // migrate old string rows
      while(arr.length<3) arr.push({name:"",zip:""});                                                        // always show at least 3 boxes
      SETTINGS.consolidated[p.key]=arr;
      const sec=document.createElement("div"); sec.className="consol-origin";
      let boxes="";
      arr.forEach((e,i)=>{ boxes+='<div class="consol-box"><button class="consol-del" data-key="'+esc(p.key)+'" data-i="'+i+'" title="Remove">&times;</button>'+
        '<input class="consol-name" data-key="'+esc(p.key)+'" data-i="'+i+'" value="'+esc(e.name||"")+'" placeholder="Location name">'+
        '<input class="consol-zip" data-key="'+esc(p.key)+'" data-i="'+i+'" value="'+esc(e.zip||"")+'" placeholder="ZIP">'+
        '</div>'; });
      sec.innerHTML='<div class="consol-lbl">'+esc(p.name)+'</div><div class="consol-list">'+boxes+'</div><button class="consol-add cp-btn" data-key="'+esc(p.key)+'">+ Add location</button>';
      cl.appendChild(sec);
    });
  }
  const sl=document.getElementById("sameloc-rows");
  if(sl){ sl.innerHTML="";
    SETTINGS.sameLoc=Array.isArray(SETTINGS.sameLoc)?SETTINGS.sameLoc:[];
    SETTINGS.sameLoc.forEach((e,i)=>{
      if(!e||typeof e!=="object"){ e={name:String(e||""),zips:[]}; SETTINGS.sameLoc[i]=e; }
      if(!Array.isArray(e.zips)) e.zips=[];
      let zb=""; e.zips.forEach((z,j)=>{ zb+='<div class="consol-box" style="width:118px"><button class="sl-delzip consol-del" data-i="'+i+'" data-j="'+j+'" title="Remove ZIP">&times;</button><input class="sl-zip consol-zip" data-i="'+i+'" data-j="'+j+'" value="'+esc(z||"")+'" placeholder="ZIP"></div>'; });
      const sec=document.createElement("div"); sec.className="consol-origin";
      sec.innerHTML='<div class="consol-lbl" style="display:flex;gap:8px;align-items:center"><input class="sl-name" data-i="'+i+'" value="'+esc(e.name||"")+'" placeholder="Location name (e.g. ADUSA DISTRIBUTION, LLC)" style="flex:1;padding:6px 8px;border:1px solid var(--line);border-radius:6px;font-size:13px"><button class="sl-delgrp consol-del" data-i="'+i+'" title="Remove this location" style="position:static;font-size:18px">&times;</button></div><div class="consol-list">'+zb+'</div><button class="sl-addzip cp-btn" data-i="'+i+'">+ Add ZIP</button>';
      sl.appendChild(sec);
    });
  }
}'''),

    # --- settings handlers: alias add/edit/delete + live re-grouping ---
    ('settings aliases handlers',
     '''  const add=document.getElementById("cust-add"); if(add) add.addEventListener("click", ()=>{ SETTINGS.customers.push({match:"",startAdj:0,endAdj:0,biz:false}); saveSettings(); renderSettings(); });
  loadSettings(); renderSettings();''',
     '''  const add=document.getElementById("cust-add"); if(add) add.addEventListener("click", ()=>{ SETTINGS.customers.push({match:"",startAdj:0,endAdj:0,biz:false}); saveSettings(); renderSettings(); });
  let _regroupT=null;
  function regroup(){ clearTimeout(_regroupT); _regroupT=setTimeout(()=>{ if(DATA) ingest(DATA); }, 400); }
  const al=document.getElementById("alias-rows");
  if(al){
    al.addEventListener("input", e=>{ if(!e.target.classList.contains("al-terms")) return; const i=+e.target.dataset.i; if(isNaN(i)||!Array.isArray(SETTINGS.aliases)) return; SETTINGS.aliases[i]=e.target.value; saveSettings(); regroup(); });
    al.addEventListener("click", e=>{ const d=e.target.closest(".al-del"); if(!d) return; SETTINGS.aliases.splice(+d.dataset.i,1); saveSettings(); renderSettings(); regroup(); });
  }
  const aadd=document.getElementById("alias-add"); if(aadd) aadd.addEventListener("click", ()=>{ SETTINGS.aliases=SETTINGS.aliases||[]; SETTINGS.aliases.push(""); saveSettings(); renderSettings(); });
  const cl2=document.getElementById("consol-rows");
  if(cl2){
    cl2.addEventListener("input", e=>{ const t=e.target; const isN=t.classList.contains("consol-name"), isZ=t.classList.contains("consol-zip"); if(!isN&&!isZ) return; const k=t.dataset.key, i=+t.dataset.i; SETTINGS.consolidated=SETTINGS.consolidated||{}; if(!Array.isArray(SETTINGS.consolidated[k])) SETTINGS.consolidated[k]=[]; let row=SETTINGS.consolidated[k][i]; if(!row||typeof row!=="object") row=SETTINGS.consolidated[k][i]={name:"",zip:""}; row[isN?"name":"zip"]=t.value; saveSettings(); if(typeof renderLoads==="function") renderLoads(); });
    cl2.addEventListener("click", e=>{
      const add=e.target.closest(".consol-add");
      if(add){ const k=add.dataset.key; SETTINGS.consolidated=SETTINGS.consolidated||{}; if(!Array.isArray(SETTINGS.consolidated[k])) SETTINGS.consolidated[k]=[]; SETTINGS.consolidated[k].push({name:"",zip:""}); saveSettings(); renderSettings(); return; }
      const del=e.target.closest(".consol-del");
      if(del){ const k=del.dataset.key, i=+del.dataset.i; if(Array.isArray(SETTINGS.consolidated[k])){ SETTINGS.consolidated[k].splice(i,1); saveSettings(); renderSettings(); if(typeof renderLoads==="function") renderLoads(); } }
    });
  }
  const slc=document.getElementById("sameloc-rows");
  if(slc){
    slc.addEventListener("input", e=>{ const t=e.target; SETTINGS.sameLoc=SETTINGS.sameLoc||[]; if(t.classList.contains("sl-name")){ const i=+t.dataset.i; if(SETTINGS.sameLoc[i]) SETTINGS.sameLoc[i].name=t.value; saveSettings(); } else if(t.classList.contains("sl-zip")){ const i=+t.dataset.i,j=+t.dataset.j; if(SETTINGS.sameLoc[i]&&Array.isArray(SETTINGS.sameLoc[i].zips)) SETTINGS.sameLoc[i].zips[j]=t.value; saveSettings(); } });
    slc.addEventListener("click", e=>{
      const az=e.target.closest(".sl-addzip"); if(az){ const i=+az.dataset.i; if(SETTINGS.sameLoc[i]){ SETTINGS.sameLoc[i].zips=SETTINGS.sameLoc[i].zips||[]; SETTINGS.sameLoc[i].zips.push(""); saveSettings(); renderSettings(); } return; }
      const dz=e.target.closest(".sl-delzip"); if(dz){ const i=+dz.dataset.i,j=+dz.dataset.j; if(SETTINGS.sameLoc[i]&&Array.isArray(SETTINGS.sameLoc[i].zips)){ SETTINGS.sameLoc[i].zips.splice(j,1); saveSettings(); renderSettings(); } return; }
      const dg=e.target.closest(".sl-delgrp"); if(dg){ const i=+dg.dataset.i; SETTINGS.sameLoc.splice(i,1); saveSettings(); renderSettings(); return; }
    });
  }
  const sladd=document.getElementById("sameloc-add"); if(sladd) sladd.addEventListener("click", ()=>{ SETTINGS.sameLoc=SETTINGS.sameLoc||[]; SETTINGS.sameLoc.push({name:"",zips:["",""]}); saveSettings(); renderSettings(); });
  loadSettings(); renderSettings();'''),

    # --- Shareable copy: ship rules as a SEED (editable + saved on each recipient's computer) ---
    ('settings shareable seed',
     '''const LOCKED = !!window.__SETTINGS__;   // true when rules were baked into this file
function saveSettings(){ if(LOCKED) return; try{ localStorage.setItem("dm_settings", JSON.stringify(SETTINGS)); }catch(e){} }
function loadSettings(){
  if(LOCKED){ SETTINGS=normSettings(window.__SETTINGS__); return; }       // locked copy: use the baked-in rules
  try{ const s=localStorage.getItem("dm_settings"); if(s){ SETTINGS=normSettings(JSON.parse(s)); return; } }catch(e){}
  SETTINGS=JSON.parse(JSON.stringify(DEFAULT_SETTINGS)); saveSettings();
}''',
     '''function saveSettings(){ try{ localStorage.setItem("dm_settings", JSON.stringify(SETTINGS)); }catch(e){} }
function loadSettings(){
  try{ const s=localStorage.getItem("dm_settings"); if(s){ SETTINGS=normSettings(JSON.parse(s)); return; } }catch(e){}   // this computer's own saved rules win
  if(window.__SETTINGS__){ SETTINGS=normSettings(window.__SETTINGS__); saveSettings(); return; }   // first open of a shared copy: seed from the baked-in rules, then editable + saved here
  SETTINGS=JSON.parse(JSON.stringify(DEFAULT_SETTINGS)); saveSettings();
}'''),

    # --- Export: clear the alias rows in the cloned copy (renderSettings rebuilds them on open) ---
    ('export clear alias rows',
     '["map","loads-body","cust-rows","routes-chips","loads-summary"].forEach',
     '["map","loads-body","cust-rows","alias-rows","routes-chips","loads-summary"].forEach'),

    # --- Export: bake the CURRENTLY-loaded loads into the shared copy (raw rows, so recipient rules re-group) ---
    ('export embed current data',
     r'''    let html = "<!DOCTYPE html>\n" + clone.outerHTML;
    html = html.replace(/<script>\s*window\.__SETTINGS__[\s\S]*?<\/script>\s*/g, "");''',
     r'''    let html = "<!DOCTYPE html>\n" + clone.outerHTML;
    if (DATA){   // bake the currently-loaded loads in so the shared copy shows them on any computer
      const snap = Object.assign({}, DATA); snap.orders = DATA._raw || DATA.orders; delete snap._raw;
      const dataJs = JSON.stringify(snap).replace(/<\//g,"<\\/");
      html = html.replace(/<script>\s*window\.__EMBEDDED_DATA__[\s\S]*?<\/script>/, "<script>window.__EMBEDDED_DATA__ = "+dataJs+";<\/script>");
    }
    html = html.replace(/<script>\s*window\.__SETTINGS__[\s\S]*?<\/script>\s*/g, "");'''),

    # --- Settings page: reword the share section ---
    ('share section heading',
     '<h3>Lock these settings into the map you email</h3>',
     '<h3>Share a copy of this map</h3>'),
    ('share section note',
     'Saves a copy of this map with the limits and customer windows above <b>baked in</b> &mdash; whoever opens that copy sees exactly these rules.',
     'Saves a complete copy &mdash; the current loads <b>and</b> all the rules above are baked in. Email it; the other person just saves it and double-clicks to open (works offline except the map background). They start with your rules and can adjust them on their own computer.'),
    ('share section button',
     '>Download a locked copy to email</button>',
     '>Download a shareable copy</button>'),

    # --- Loads by DC: "good consolidation lanes" grouping option ---
    ('lanes dropdown option',
     '      <option value="loc">exact location</option>',
     '      <option value="loc">exact location</option>\n      <option value="lanes">good consolidation lanes</option>'),
    ('lanes destkey',
     '''  function destKey(o){
    const z=nz(o.zip), loc=o.lat.toFixed(5)+","+o.lng.toFixed(5);
    if(grp==="loc") return "L:"+loc;''',
     '''  function destKey(o){
    const z=nz(o.zip), loc=o.lat.toFixed(5)+","+o.lng.toFixed(5);
    if(grp==="lanes"){
      let entries=(SETTINGS.consolidated && SETTINGS.consolidated[o.pick]) || [];
      if(typeof entries==="string") entries=entries.split(/[\\n,]/).map(s=>({name:s.trim(),zip:""}));
      if(!Array.isArray(entries)) entries=[];
      const up=(o.name||"").toUpperCase();
      const oz=String(o.zip||"").replace(/[^0-9]/g,"").slice(0,5);
      for(const e of entries){
        const en=(typeof e==="string"?e:((e&&e.name)||"")).trim().toUpperCase();
        const ez=(typeof e==="object"&&e?String(e.zip||""):"").replace(/[^0-9]/g,"").slice(0,5);
        if(!en && !ez) continue;
        const nameOk = en ? (up.indexOf(en)>=0) : true;   // if name given it must match
        const zipOk  = ez ? (oz===ez) : true;             // if ZIP given it must match
        if(nameOk && zipOk) return "LANE";                // both must match when both are filled in
      }
      return "L:"+loc;                                                    // not in a lane -> its own destination
    }
    if(grp==="loc") return "L:"+loc;'''),
    # --- Loads by DC: when a destination group holds more than one customer (e.g. ZIP-only),
    #     show ALL their names instead of just the most-common one ---
    ('group title lists all customers',
     '''  let name="", best=-1; const counts={};
  orders.forEach(o => { if(o.name){ counts[o.name]=(counts[o.name]||0)+1; if(counts[o.name]>best){best=counts[o.name]; name=o.name;} } });''',
     '''  const _nm=[]; orders.forEach(o => { if(o.name && _nm.indexOf(o.name)<0) _nm.push(o.name); });
  let name = _nm.length ? (_nm.slice(0,3).join(" + ") + (_nm.length>3 ? " +"+(_nm.length-3)+" more" : "")) : "";'''),

    # --- Delivery window: a NEGATIVE "End" adjustment means "deliver up to N days BEFORE the
    #     Last Drop Plan End date." End stays the latest; the Start date is ignored. ---
    ('end-anchored delivery window',
     '''    const r=ruleForOrDefault(o.name);                                 // standard or per-customer window
    let ws=addDays(ps, r.startAdj||0, r.biz), we=addDays(pe, r.endAdj||0, r.biz);
    const pk=dayNum(o.pickDate); if(pk!=null) ws=Math.max(ws,pk);     // can't ship before pick date''',
     '''    const r=ruleForOrDefault(o.name);                                 // standard or per-customer window
    let ws, we;
    if((r.endAdj||0) < 0){                                            // "N days before the End date": End is the latest, Start is ignored
      we = pe; ws = addDays(pe, r.endAdj, r.biz);
    } else {
      ws = addDays(ps, r.startAdj||0, r.biz); we = addDays(pe, r.endAdj||0, r.biz);
    }
    const pk=dayNum(o.pickDate); if(pk!=null) ws=Math.max(ws,pk);     // can't ship before pick date'''),

    ('lanes group label',
     '    let groups=Object.keys(byDC[dc]).map(lk=>analyzeGroup(byDC[dc][lk], maxW, maxP));',
     '    let groups=[];\n    Object.keys(byDC[dc]).forEach(lk=>{ clusterByWindow(byDC[dc][lk]).forEach(cl=>{ const g=analyzeGroup(cl, maxW, maxP); if(lk==="LANE") g.name="\\u2605 Consolidation lane"; groups.push(g); }); });'),

    # --- Loads by DC: richer load rows (location name, pick appt + delivery appt w/ time, clickable TMS) ---
    ('loadList rich rows',
     '''function loadList(orders){
  let h='<div class="loadlist">';
  orders.forEach(o => {
    const win = o.planStart ? (o.planStart===o.planEnd ? fmtDayISO(o.planStart) : fmtDayISO(o.planStart)+" &ndash; "+fmtDayISO(o.planEnd)) : "no date";
    h+='<div class="loadrow"><span class="ref">'+esc(refOf(o))+'</span>'+
       (o.status?'<span class="badge2 '+esc(o.status)+'">'+esc(o.status)+'</span>':'')+
       '<span class="sub">'+(o.weight!=null?Math.round(o.weight).toLocaleString()+" lb":"")+
       (o.palletSpaces!=null?" &middot; "+(Math.round(o.palletSpaces*10)/10)+" pallet sp":"")+
       " &middot; deliver "+win+(o.pickDate?" &middot; picks "+fmtDayISO(o.pickDate):"")+'</span></div>';
  });
  return h+'</div>';
}''',
     '''let _hlTms=null;
function _clearTmsBanners(){ document.querySelectorAll("#loads-body .dc-banner").forEach(b=>{ b.style.display="none"; b.innerHTML=""; }); }
function _sameLocKey(o){                                              // user-defined "same location" merge rules (Rules & Limits page)
  const up=(o.name||"").toUpperCase(), z5=String(o.zip||"").replace(/[^0-9]/g,"").slice(0,5);
  const rules=(typeof SETTINGS!=="undefined" && Array.isArray(SETTINGS.sameLoc))?SETTINGS.sameLoc:[];
  for(let i=0;i<rules.length;i++){ const r=rules[i]||{}; const rn=String(r.name||"").toUpperCase().trim(); if(!rn) continue;
    const zs=(Array.isArray(r.zips)?r.zips:[]).map(z=>String(z||"").replace(/[^0-9]/g,"").slice(0,5)).filter(Boolean);
    if(up.indexOf(rn)>=0 && (zs.length===0 || zs.indexOf(z5)>=0)) return "SAME#"+i; }   // matched a rule -> one canonical location
  return up.trim()+"|"+z5;
}
function _locLabel(ls){ const nm=ls[0].name||""; const zs=[...new Set(ls.map(o=>String(o.zip||"").trim()).filter(Boolean))]; return esc(nm)+(zs.length?' ('+zs.map(esc).join(", ")+')':''); }
function highlightTms(t){
  document.querySelectorAll("#loads-body .loadrow.tms-hl").forEach(el=>el.classList.remove("tms-hl"));
  _clearTmsBanners();
  if(_hlTms===t){ _hlTms=null; return; }                              // click same TMS again to clear
  _hlTms=t;
  const sel='#loads-body .loadrow[data-tms="'+((window.CSS&&CSS.escape)?CSS.escape(t):t)+'"]';
  const rows=document.querySelectorAll(sel);
  rows.forEach(r=>{ r.classList.add("tms-hl"); const det=r.closest("tr.detail"); if(det && det.style.display==="none"){ det.style.display=""; const grp=det.previousElementSibling; if(grp) grp.classList.add("open"); } });
  const mates=(DATA._raw||DATA.orders||[]).filter(o=>String(o.tms)===String(t));
  if(mates.length && rows[0]){
    const wt=mates.reduce((s,o)=>s+(o.weight||0),0), sp=mates.reduce((s,o)=>s+(o.palletSpaces||0),0);
    const locMap={}; mates.forEach(o=>{ const k=_sameLocKey(o); (locMap[k]=locMap[k]||[]).push(o); });
    const locKeys=Object.keys(locMap);
    const days=[...new Set(mates.map(o=>_apptStr(o,"Last Drop Appt Date Start").replace(/\\s+\\d{1,2}:\\d{2}.*$/,"").trim()).filter(Boolean))];
    const warns=[];                                                  // same location, but plan delivery windows do not all overlap
    locKeys.forEach(k=>{
      const list=locMap[k].filter(o=>o.planStart&&o.planEnd).map(o=>({ref:refOf(o), ws:dayNum(o.planStart), we:dayNum(o.planEnd)}));
      if(list.length<2) return;
      list.sort((a,b)=>a.ws-b.ws);
      const cl=[];
      list.forEach(it=>{ let placed=false; for(const c of cl){ const nws=Math.max(c.ws,it.ws), nwe=Math.min(c.we,it.we); if(nws<=nwe){ c.ws=nws; c.we=nwe; c.refs.push(it.ref); placed=true; break; } } if(!placed) cl.push({ws:it.ws,we:it.we,refs:[it.ref]}); });
      if(cl.length>1){
        const parts=cl.map(c=>'<span class="wgrp">'+esc(fmtRange(c.ws,c.we))+': '+c.refs.map(esc).join(", ")+'</span>');
        warns.push('<b>'+_locLabel(locMap[k])+'</b> '+parts.join(' <i>vs</i> '));
      }
    });
    let breakdown='';
    if(locKeys.length>1){
      breakdown='<div class="tms-locs">'+locKeys.map(k=>{ const ls=locMap[k]; const lw=ls.reduce((s,o)=>s+(o.weight||0),0), lp=ls.reduce((s,o)=>s+(o.palletSpaces||0),0); return '<div class="locrow">&#8226; <b>'+_locLabel(ls)+'</b> &mdash; '+ls.length+' shipment'+(ls.length>1?'s':'')+' &middot; '+Math.round(lw).toLocaleString()+' lb &middot; '+(Math.round(lp*10)/10)+' pallet sp</div>'; }).join('')+'</div>';
    }
    const html='<b>Load '+esc(t)+'</b> &mdash; '+mates.length+' shipment'+(mates.length>1?'s':'')+' &middot; '+locKeys.length+' location'+(locKeys.length>1?'s':'')+' &middot; '+Math.round(wt).toLocaleString()+' lb &middot; '+(Math.round(sp*10)/10)+' pallet sp'+(days.length?' &middot; deliver '+esc(days.join(", ")):'')+' <span class="tms-clear">clear &times;</span><span class="tms-findmatch" data-tms="'+esc(t)+'" title="Find loads that could ship with this one">🔍 Find consolidation matches</span>'+breakdown+(warns.length?'<div class="tms-warn">&#9888;&#65039; Same location, delivery windows do not overlap &mdash; '+warns.join('; ')+'</div>':'');
    const sec=rows[0].closest(".dc-sec"); const bn=sec?sec.querySelector(".dc-banner"):null;
    if(bn){ bn.innerHTML=html; bn.style.display=""; bn.scrollIntoView({block:"nearest"}); }
  }
}
function _apptStr(o, label){ const f=(o.fields||[]).find(x=>x[0]===label); const v=f?String(f[1]):""; return v.replace(/,? \\d{4}/,""); }
function loadList(orders){
  let h='<div class="loadlist">';
  orders.forEach(o => {
    const del=_apptStr(o,"Last Drop Appt Date Start");
    const pik=_apptStr(o,"First Pick Appt Date Start");
    const win = o.planStart ? (o.planStart===o.planEnd ? fmtDayISO(o.planStart) : fmtDayISO(o.planStart)+" &ndash; "+fmtDayISO(o.planEnd)) : "";
    const chkd = o.tms && selectedTms.has(String(o.tms));
    h+='<div class="loadrow'+(chkd?' sel-row':'')+'"'+(o.tms?' data-tms="'+esc(o.tms)+'"':'')+'>'+
       (o.tms?'<input type="checkbox" class="load-chk" data-tms="'+esc(o.tms)+'"'+(chkd?' checked':'')+' title="Select for export">':'<span style="width:14px;flex-shrink:0"></span>')+
       '<span class="ref">'+esc(refOf(o))+'</span>'+
       (o.tms?'<span class="tms-chip" data-tms="'+esc(o.tms)+'" title="Click to highlight everything on this load">'+esc(o.tms)+'</span>':'')+
       (o.status?'<span class="badge2 '+esc(o.status)+'">'+esc(o.status)+'</span>':'')+
       (o.name?'<span class="lname">'+esc(o.name)+'</span>':'')+
       '<span class="sub">'+(o.weight!=null?Math.round(o.weight).toLocaleString()+" lb":"")+
       (o.palletSpaces!=null?" &middot; "+(Math.round(o.palletSpaces*10)/10)+" pallet sp":"")+
       (pik?" &middot; pick appt "+esc(pik):"")+(del?" &middot; delivery appt "+esc(del):"")+(win?" &middot; delivery window "+win:"")+'</span>'+
       '</div>';
  });
  return h+'</div>';
}'''),

    # --- Top bar: "Data as of ..." stamp ---
    ('topbar data stamp',
     '''  <div class="pagenav">
    <button id="nav-map" class="pagebtn active">Map</button>
    <button id="nav-loads" class="pagebtn">Loads by DC</button>
    <button id="nav-settings" class="pagebtn">Rules &amp; Limits</button>
  </div>''',
     '''  <div class="pagenav">
    <button id="nav-map" class="pagebtn active">Map</button>
    <button id="nav-loads" class="pagebtn">Loads by DC</button>
    <button id="nav-settings" class="pagebtn">Rules &amp; Limits</button>
  </div>
  <div id="topdate" class="topdate">Data as of &mdash;</div>'''),

    # --- Loads page: add export bar at bottom ---
    ('loads page export bar',
     '<div id="loads-summary"></div>',
     '<div id="loads-summary"></div>\n  <div id="export-bar" class="export-bar"><span class="exp-count" id="exp-count">0 loads selected</span><button class="exp-btn" id="exp-btn">⬇ Export to Excel (CSV)</button><button class="exp-clear" id="exp-clear">Clear selection</button></div>'),
    ('topbar data stamp fill',
     'document.getElementById("stale-banner").style.display=data.stale?"block":"none"; }',
     'document.getElementById("stale-banner").style.display=data.stale?"block":"none"; var td=document.getElementById("topdate"); if(td) td.innerHTML="Data as of <b>"+esc(data.updatedAt||"")+"</b>"; }'),

    # --- Loads by DC: "same ZIP + delivery date" grouping ---
    #     groups same-ZIP loads that share the same Last Drop Plan window, so an off-date
    #     load doesn't poison the whole ZIP. This is the "ship together if same zip + window" view.
    ('zipwin dropdown option',
     '      <option value="zip">same ZIP only</option>',
     '      <option value="zipwin">same ZIP + delivery date</option>'),
    ('zipwin destkey',
     '    if(grp==="zip") return z ? "Z:"+z : "L:"+loc;',
     '    if(grp==="zipwin"){ const w=(o.planStart||"?")+".."+(o.planEnd||"?"); return z ? "ZW:"+z+"|"+w : "L:"+loc; }\n    if(grp==="zip") return z ? "Z:"+z : "L:"+loc;'),

    # --- Loads by DC: keyword search + per-column filters + sortable headers ---
    ('loads search controls',
     '''    <label><input type="checkbox" id="c-greenonly"> only combinable (green)</label>
    <span class="ctrl-note">Truck limits &amp; delivery&#8209;window rules are on the <a id="goto-settings" class="lnk">Rules &amp; Limits</a> tab.</span>''',
     '''    <label><input type="checkbox" id="c-greenonly"> only combinable (green)</label>
    <span class="ctrl-note">Truck limits &amp; delivery&#8209;window rules are on the <a id="goto-settings" class="lnk">Rules &amp; Limits</a> tab.</span>
    <div id="dc-filter" class="dc-filter"></div>
    <div class="lf-row">
      <input id="lf-search" type="search" class="lf-in" placeholder="Search destination, ZIP, reference #..." />
      <label>ZIP <input id="lf-zip" class="lf-in lf-sm" placeholder="any" /></label>
      <label>Min loads <input id="lf-minloads" type="number" min="0" step="1" class="lf-in lf-xs" /></label>
      <label>Min pallet sp <input id="lf-minpal" type="number" min="0" step="0.5" class="lf-in lf-xs" /></label>
      <button id="lf-clear" class="cp-btn" type="button">Clear filters</button>
      <span class="ctrl-note">Click a column header to sort.</span>
    </div>'''),

    ('loads filter+sort helpers',
     'function renderLoads(){',
     '''function _loadsCtl(id){ const e=document.getElementById(id); return e?e.value:""; }
function applyLoadsFilters(groups){
  const q=_loadsCtl("lf-search").trim().toLowerCase();
  const zq=_loadsCtl("lf-zip").trim().toLowerCase();
  const minL=parseFloat(_loadsCtl("lf-minloads")), minP=parseFloat(_loadsCtl("lf-minpal"));
  return groups.filter(g=>{
    if(zq && String(g.zip||"").toLowerCase().indexOf(zq)<0) return false;
    if(isFinite(minL) && g.n < minL) return false;
    if(isFinite(minP) && g.pallets < minP) return false;
    if(q){
      let hay=(g.name||"")+" "+(g.zip||"");
      for(const o of g.orders){ hay+=" "+(o.name||""); for(const f of (o.fields||[])){ if(/reference/i.test(f[0])) hay+=" "+f[1]; } }
      if(hay.toLowerCase().indexOf(q)<0) return false;
    }
    return true;
  });
}
function sortLoadGroups(groups){
  if(!loadsSortKey){ groups.sort((a,b)=> (b.green-a.green) || (b.n-a.n) || (b.weight-a.weight)); return; }
  const v=(g)=>{ switch(loadsSortKey){
    case "name": return (g.name||"").toLowerCase();
    case "zip": return g.zip||"";
    case "n": return g.n; case "weight": return g.weight; case "pallets": return g.pallets; case "fill": return g.fill;
    case "window": return (g.windowText||"").toLowerCase();
    case "st": return g.green?1:0;
    default: return 0; } };
  groups.sort((a,b)=>{ const x=v(a), y=v(b); if(x<y) return -loadsSortDir; if(x>y) return loadsSortDir; return 0; });
}
function loadsHeader(){
  const cols=[["name","Destination",""],["zip","ZIP",""],["n","Loads","num"],["weight","Weight (lb)","num"],["pallets","Pallet Spaces","num"],["fill","Truck %","num"],["window","Delivery window",""],["st","Ship together?",""]];
  let h="<tr>";
  for(let i=0;i<cols.length;i++){ const c=cols[i]; const ar=loadsSortKey===c[0]?(loadsSortDir>0?" \\u25B2":" \\u25BC"):""; h+='<th class="lh '+c[2]+'" data-col="'+c[0]+'">'+c[1]+ar+'</th>'; }
  return h+"</tr>";
}
function orderWindow(o){
  const ps=dayNum(o.planStart), pe=dayNum(o.planEnd);
  if(ps==null||pe==null) return null;                              // no dates -> compatible with anything
  const r=ruleForOrDefault(o.name);
  let ws, we;
  if((r.endAdj||0)<0){ we=pe; ws=addDays(pe, r.endAdj, r.biz); }   // "N days before End" rule (Start ignored)
  else { ws=addDays(ps, r.startAdj||0, r.biz); we=addDays(pe, r.endAdj||0, r.biz); }
  const pk=dayNum(o.pickDate); if(pk!=null) ws=Math.max(ws,pk);    // never before pick date
  return {ws:ws, we:we};
}
function clusterByWindow(orders){
  // split a destination's loads into sets that SHARE a common delivery day (rule-adjusted),
  // so one off-date load can't block the rest from showing as combinable.
  const arr=orders.map(o=>({o:o, w:orderWindow(o)}));
  arr.sort((a,b)=> ((b.w?b.w.ws:-Infinity) - (a.w?a.w.ws:-Infinity)));   // latest-starting first
  const clusters=[];
  for(const it of arr){
    let placed=false;
    if(it.w==null){ if(clusters.length){ clusters[0].items.push(it.o); placed=true; } }
    else { for(const c of clusters){ const nWs=Math.max(c.ws,it.w.ws), nWe=Math.min(c.we,it.w.we); if(nWs<=nWe){ c.items.push(it.o); c.ws=nWs; c.we=nWe; placed=true; break; } } }
    if(!placed) clusters.push({ws: it.w?it.w.ws:-Infinity, we: it.w?it.w.we:Infinity, items:[it.o]});
  }
  return clusters.map(c=>c.items);
}
function renderLoads(){'''),

    # --- Loads by DC: analyze the RAW shipments (true per-shipment windows), not the
    #     merged-stop span, so a multi-window TMS load can't create a false overlap ---
    ('loads use raw shipments',
     'DATA.orders.forEach(o => { const dc=o.pick||"__none"; const k=destKey(o);',
     '(DATA._raw||DATA.orders).forEach(o => { const dc=o.pick||"__none"; const k=destKey(o);'),

    ('loads filter+sort apply',
     '    groups.sort((a,b)=> (b.green-a.green) || (b.n-a.n) || (b.weight-a.weight));',
     '    groups = applyLoadsFilters(groups);\n    sortLoadGroups(groups);'),

    # --- Loads by DC: per-DC filter buttons above the search row ---
    ('loads dc filter buttons',
     '''  let html="", totalGreen=0, totalGreenLoads=0;
  order.forEach(dc => {
    if(!byDC[dc]) return;''',
     '''  { const dcfEl=document.getElementById("dc-filter");
    if(dcfEl){ let bh='<button class="dcf-btn'+(loadsDC===""?" on":"")+'" data-dc="">All DCs</button>';
      order.forEach(dc=>{ if(!byDC[dc]) return; const nm=dc==="__none"?"No DC / unmatched":((pickByKey.get(dc)||{}).name||dc); bh+='<button class="dcf-btn'+(loadsDC===dc?" on":"")+'" data-dc="'+esc(dc)+'">'+esc(nm)+'</button>'; });
      dcfEl.innerHTML=bh; } }
  let html="", totalGreen=0, totalGreenLoads=0;
  order.forEach(dc => {
    if(!byDC[dc]) return;
    if(loadsDC && dc!==loadsDC) return;'''),

    ('loads sortable header',
     '''      '<table class="lt"><thead><tr><th>Destination</th><th>ZIP</th><th class="num">Loads</th>'+
      '<th class="num">Weight (lb)</th><th class="num">Pallet Spaces</th><th class="num">Truck %</th><th>Delivery window</th><th>Ship together?</th></tr></thead><tbody>';''',
     '''      '<div class="dc-banner tms-banner" style="display:none"></div>'+
      '<table class="lt"><thead>'+loadsHeader()+'</thead><tbody>';'''),

    # --- Loads by DC: add Best Rate column to header and destination rows ---
    ('loads rate column header',
     'const cols=[["name","Destination",""],["zip","ZIP",""],["n","Loads","num"],["weight","Weight (lb)","num"],["pallets","Pallet Spaces","num"],["fill","Truck %","num"],["window","Delivery window",""],["st","Ship together?",""]];',
     'const cols=[["name","Destination",""],["zip","ZIP",""],["n","Loads","num"],["weight","Weight (lb)","num"],["pallets","Pallet Spaces","num"],["fill","Truck %","num"],["window","Delivery window",""],["st","Ship together?",""],["rate","Best rate",""]];'),

    ('loads rate column row',
     '''html+='<tr class="grp '+(g.green?"green":"")+'"><td><span class="caret">&#9656;</span>'+esc(g.name)+'</td><td>'+esc(g.zip)+'</td>'+
        '<td class="num">'+g.n+'</td><td class="num">'+Math.round(g.weight).toLocaleString()+'</td>'+
        '<td class="num">'+(Math.round(g.pallets*10)/10).toLocaleString()+'</td>'+
        '<td class="num">'+g.fill+'%</td>'+
        '<td>'+esc(g.windowText)+'</td><td class="st '+g.st+'">'+esc(g.statusText)+'</td></tr>';
      html+='<tr class="detail" style="display:none"><td colspan="8">'+loadList(g.orders)+'</td></tr>';''',
     '''const _dLat = g.orders.length ? g.orders[0].lat : null;
      const _dLng = g.orders.length ? g.orders[0].lng : null;
      const _rateCell = (DATA.rates && Object.keys(DATA.rates).length && dc!=="__none")
        ? '<td style="font-size:11.5px;white-space:nowrap">'+ratesBadge(dc, g.zip, _dLat, _dLng)+'</td>' : '<td></td>';
      html+='<tr class="grp '+(g.green?"green":"")+'"><td><span class="caret">&#9656;</span>'+esc(g.name)+'</td><td>'+esc(g.zip)+'</td>'+
        '<td class="num">'+g.n+'</td><td class="num">'+Math.round(g.weight).toLocaleString()+'</td>'+
        '<td class="num">'+(Math.round(g.pallets*10)/10).toLocaleString()+'</td>'+
        '<td class="num">'+g.fill+'%</td>'+
        '<td>'+esc(g.windowText)+'</td><td class="st '+g.st+'">'+esc(g.statusText)+'</td>'+_rateCell+'</tr>';
      html+='<tr class="detail" style="display:none"><td colspan="9">'+loadList(g.orders)+'</td></tr>';'''),

    ('loads controls wiring',
     '''["c-greenonly","c-group"].forEach(id => {
  const el = document.getElementById(id); if (el) el.addEventListener("input", renderLoads);
});''',
     '''["c-greenonly","c-group","lf-search","lf-zip","lf-minloads","lf-minpal"].forEach(id => {
  const el = document.getElementById(id); if (el) el.addEventListener("input", renderLoads);
});
{ const lc=document.getElementById("lf-clear"); if(lc) lc.addEventListener("click", ()=>{ ["lf-search","lf-zip","lf-minloads","lf-minpal"].forEach(id=>{ const e=document.getElementById(id); if(e) e.value=""; }); renderLoads(); }); }
{ const dcf=document.getElementById("dc-filter"); if(dcf) dcf.addEventListener("click", e=>{ const b=e.target.closest(".dcf-btn"); if(!b) return; loadsDC=b.dataset.dc||""; renderLoads(); }); }'''),

    ('loads header click sort',
     '''document.getElementById("loads-body").addEventListener("click", e => {
  const tr=e.target.closest("tr.grp"); if(!tr) return;''',
     '''document.getElementById("loads-body").addEventListener("click", e => {
  const clr=e.target.closest(".tms-clear");
  if(clr){ if(_hlTms) highlightTms(_hlTms); clearMatchHighlight(); return; }
  const fm=e.target.closest(".tms-findmatch");
  if(fm){ findConsolidationMatches(fm.dataset.tms); return; }
  const chip=e.target.closest(".tms-chip");
  if(chip){ highlightTms(chip.dataset.tms); return; }
  const th=e.target.closest("th.lh");
  if(th){ const c=th.dataset.col; if(loadsSortKey===c) loadsSortDir=-loadsSortDir; else { loadsSortKey=c; loadsSortDir=1; } renderLoads(); return; }
  const tr=e.target.closest("tr.grp"); if(!tr) return;'''),
]

NEW_FUNCS = '''
// ---- Rate lookup helpers ----
const _CA_PROV = {A:'NL',B:'NS',C:'PE',E:'NB',G:'QC',H:'QC',J:'QC',K:'ON',L:'ON',M:'ON',N:'ON',P:'ON',R:'MB',S:'SK',T:'AB',V:'BC',X:'NT',Y:'YT'};
const _ZIP3ST = (()=>{
  const t={};
  [['005','005','MA'],['010','027','MA'],['028','029','RI'],['030','038','NH'],['039','049','ME'],
   ['050','059','VT'],['060','069','CT'],['070','089','NJ'],['100','149','NY'],['150','196','PA'],
   ['197','199','DE'],['200','205','DC'],['206','219','MD'],['220','246','VA'],['247','268','WV'],
   ['270','289','NC'],['290','299','SC'],['300','319','GA'],['320','349','FL'],['350','369','AL'],
   ['370','385','TN'],['386','397','MS'],['398','399','GA'],['400','427','KY'],['430','458','OH'],
   ['460','479','IN'],['480','499','MI'],['500','528','IA'],['530','549','WI'],['550','567','MN'],
   ['570','577','SD'],['580','588','ND'],['590','599','MT'],['600','631','IL'],['633','641','MO'],
   ['644','658','MO'],['660','679','KS'],['680','693','NE'],['700','714','LA'],['716','729','AR'],
   ['730','749','OK'],['750','799','TX'],['800','816','CO'],['820','831','WY'],['832','838','ID'],
   ['840','847','UT'],['850','865','AZ'],['870','884','NM'],['889','898','NV'],['900','961','CA'],
   ['970','979','OR'],['980','994','WA'],['995','999','AK']
  ].forEach(([lo,hi,st])=>{ for(let i=+lo;i<=+hi;i++) t[String(i).padStart(3,'0')]=st; });
  return t;
})();
function zipToState(z){
  if(!z) return null;
  z = String(z).trim();
  if(!/^\\d/.test(z)) return _CA_PROV[z[0].toUpperCase()]||null;
  return _ZIP3ST[z.slice(0,3).padStart(3,'0')]||null;
}
function _haverMi(lat1,lng1,lat2,lng2){
  const R=3958.8, r=Math.PI/180;
  const dLat=(lat2-lat1)*r, dLng=(lng2-lng1)*r;
  const a=Math.sin(dLat/2)**2+Math.cos(lat1*r)*Math.cos(lat2*r)*Math.sin(dLng/2)**2;
  return R*2*Math.asin(Math.sqrt(a));
}
function getBestRates(pickKey, dropZip, dropLat, dropLng){
  const ratesDB = (DATA.rates)||{};
  const pick = pickByKey.get(pickKey); if(!pick) return [];
  const origCity = pick.name.split(',')[0].trim().toUpperCase();
  const origAlt  = origCity==='ST PAUL'?'SAINT PAUL':(origCity==='SAINT PAUL'?'ST PAUL':null);
  const laneSt   = ratesDB[origCity] || (origAlt&&ratesDB[origAlt]) || {};
  const st = zipToState(dropZip);
  if(!st) return [];
  const lane = laneSt[st] || {};

  // Estimate road miles (haversine × 1.3 road factor) if we have coords
  const haveDist = (dropLat!=null && dropLng!=null);
  const estMi    = haveDist ? _haverMi(pick.lat, pick.lng, dropLat, dropLng)*1.3 : null;

  const all = [];

  // TL entries
  for(const e of (lane.tl||[])){
    let estCost = null;
    if(e.b==='FLT'){
      estCost = e.r;
    } else if(e.b==='CPM' && estMi!=null){
      estCost = Math.max(e.r * estMi, e.m||0);
    }
    const cur = e.cy==='CAD'?'C$':'$';
    let display;
    if(e.b==='FLT') display = cur+e.r.toLocaleString(undefined,{maximumFractionDigits:0})+' flat';
    else if(estCost!=null) display = cur+Math.round(estCost).toLocaleString()+' est.';
    else display = cur+e.r.toFixed(2)+'/mi';
    all.push({ carrier:e.c, mode:'TL', basis:e.b, estCost, display, currency:e.cy,
               sortKey: estCost!=null ? estCost : (e.b==='CPM' ? e.r*500 : e.r) });
  }

  // LTL entries — min_cost is the floor charge; label as "from $X"
  for(const e of (lane.ltl||[])){
    if(!e.m || e.m<=0) continue;
    const cur = e.cy==='CAD'?'C$':'$';
    const display = cur+Math.round(e.m).toLocaleString()+' min';
    all.push({ carrier:e.c, mode:'LTL', basis:e.b, estCost:e.m, display, currency:e.cy,
               sortKey: e.m });
  }

  // Sort by sortKey ascending; dedup carrier+mode (keep cheapest)
  const seen = new Map();
  for(const e of all){
    const k = e.carrier+'|'+e.mode;
    if(!seen.has(k) || e.sortKey < seen.get(k).sortKey) seen.set(k, e);
  }
  return [...seen.values()].sort((a,b)=>a.sortKey-b.sortKey).slice(0,5);
}
function ratesBadge(pickKey, dropZip, dropLat, dropLng){
  const rates = getBestRates(pickKey, dropZip, dropLat, dropLng);
  if(!rates.length) return '<span style="color:#aaa;font-size:11px">—</span>';
  const best = rates[0];
  const modeColor = best.mode==='LTL' ? '#7c3aed' : '#1d4ed8';
  const modeBadge = '<span style="background:'+modeColor+';color:#fff;border-radius:3px;padding:1px 5px;font-size:10px;font-weight:700;margin-right:4px">'+best.mode+'</span>';
  const tip = rates.map(r=>'['+r.mode+'] '+r.carrier+': '+r.display+(r.mode==='LTL'?' (min charge)':r.basis==='CPM'?' (est. road mi × rate)':'')).join('\\n');
  const note = best.mode==='LTL' ? ' (min)' : (best.basis==='CPM' ? ' (est.)' : '');
  return '<span class="rate-badge" title="'+esc(tip)+'">'+modeBadge+esc(best.carrier)+' '+esc(best.display)+esc(note)+'</span>';
}
function buildLoadGroups(){
  const m = new Map();
  for (const o of (DATA.orders||[])){ const g=(o.loadGroup||"").trim(); if(!g) continue; m.set(g,(m.get(g)||0)+1); }
  loadGroups = [...m.keys()].sort();
  lgCounts = { ALL: (DATA.orders||[]).length };
  m.forEach((v,k)=>{ lgCounts[k]=v; });
  if (activeLoadGroup!=="ALL" && loadGroups.indexOf(activeLoadGroup)<0) activeLoadGroup="ALL";
}
function buildLoadGroupTabs(){
  const wrap = document.getElementById("lg-tabs"); if(!wrap) return;
  wrap.innerHTML = "";
  const tabs = [{key:"ALL", name:"All"}].concat(loadGroups.map(g=>({key:g, name:g})));
  for (const t of tabs){
    const btn = document.createElement("button");
    btn.className = "lgbtn" + (t.key===activeLoadGroup ? " active":"");
    const cnt = lgCounts[t.key]!==undefined ? lgCounts[t.key] : 0;
    btn.innerHTML = esc(t.name) + ' <span class="pcount">'+cnt+'</span>';
    btn.addEventListener("click", () => { activeLoadGroup = t.key; buildLoadGroupTabs(); applyFilters(false); });
    wrap.appendChild(btn);
  }
}
function setupDateBounds(){
  let lo=Infinity, hi=-Infinity;
  for (const o of (DATA.orders||[])){
    const d=dayNum(o.planStart);
    if(d!=null){ lo=Math.min(lo,d); hi=Math.max(hi,d); }
  }
  if(isFinite(lo) && isFinite(hi)){ dateMin=lo; dateMax=hi; initDateSlider(); }
}
function initDateSlider(){
  const elMin=document.getElementById("dr-min"), elMax=document.getElementById("dr-max");
  if(!elMin||!elMax) return;
  const span = Math.max(0, dateMax-dateMin);
  elMin.max=span; elMax.max=span; elMin.value=0; elMax.value=span;
  dateFrom=dateMin; dateTo=dateMax;
  if(!sliderInit){
    sliderInit=true;
    const h=()=>{ syncSlider(); applyFilters(false); };
    elMin.addEventListener("input", h); elMax.addEventListener("input", h);
    // don't let scrolling the sidebar accidentally drag the date handles
    const noWheel=e=>e.preventDefault();
    elMin.addEventListener("wheel", noWheel, {passive:false});
    elMax.addEventListener("wheel", noWheel, {passive:false});
  }
  syncSlider();
}
function syncSlider(){
  const elMin=document.getElementById("dr-min"), elMax=document.getElementById("dr-max");
  const fill=document.getElementById("dr-fill"); if(!elMin||!elMax) return;
  const span = Math.max(1, dateMax-dateMin);
  const lo=Math.min(+elMin.value,+elMax.value), hi=Math.max(+elMin.value,+elMax.value);
  dateFrom = dateMin + lo; dateTo = dateMin + hi;
  if(fill){ const pa=lo/span*100, pb=hi/span*100; fill.style.left=pa+"%"; fill.style.width=(pb-pa)+"%"; }
  const f=document.getElementById("dr-from"), t=document.getElementById("dr-to");
  if(f) f.textContent=fmtDayY(dateFrom); if(t) t.textContent=fmtDayY(dateTo);
}
function fmtDayY(dn){ const d=new Date(dn*86400000);
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getUTCMonth()]+" "+d.getUTCDate()+", "+d.getUTCFullYear(); }

/* ---- group orders into stops: same TMS ID + same Last Drop Name = one combined order ---- */
function _fmtNum(n){ if(n==null) return ""; return Number.isInteger(n)? n.toLocaleString() : n.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }
function combineStop(arr){
  const base=Object.assign({}, arr[0]);
  base._count=arr.length;
  base.pallets=arr.reduce((s,o)=>s+(o.pallets||0),0);
  base.weight=arr.reduce((s,o)=>s+(o.weight||0),0);
  base.palletSpaces=arr.reduce((s,o)=>s+(o.palletSpaces||0),0);
  // a merged load's delivery window spans ALL its rows (earliest start -> latest end),
  // so a multi-window TMS load isn't understated to just the first row's dates
  const _ps=arr.map(o=>o.planStart).filter(Boolean); if(_ps.length) base.planStart=_ps.reduce((a,b)=>a<b?a:b);
  const _pe=arr.map(o=>o.planEnd).filter(Boolean); if(_pe.length) base.planEnd=_pe.reduce((a,b)=>a>b?a:b);
  const labels=arr[0].fields.map(f=>f[0]);
  base.fields=labels.map((lab,i)=>{
    if(lab==="Pallets") return [lab,_fmtNum(base.pallets)];
    if(lab==="Weight (lb)") return [lab,_fmtNum(base.weight)];
    if(lab==="Pallet Spaces") return [lab,_fmtNum(base.palletSpaces)];
    if(lab==="Shipment Reference Numbers"){
      const refs=arr.map(o=>(o.fields[i]||["",""])[1]).filter(v=>v!=="" && v!=null);
      return [lab, refs.join(",")];
    }
    return [lab,(arr[0].fields[i]||["",""])[1]];   // any other differing column: first row's value
  });
  return base;
}
function _banner(name){
  const up=String(name||"").toUpperCase();
  // user-defined "same-location" groups from the Rules & Limits page: any name containing
  // a term in a group is treated as that one customer.
  if(typeof SETTINGS!=="undefined" && SETTINGS && Array.isArray(SETTINGS.aliases)){
    for(let gi=0; gi<SETTINGS.aliases.length; gi++){
      const terms=String(SETTINGS.aliases[gi]||"").split(",").map(t=>t.trim().toUpperCase()).filter(Boolean);
      for(const t of terms){ if(up.indexOf(t)>=0) return "ALIAS#"+gi; }
    }
  }
  // default: first word, ignoring hyphens/apostrophes/spacing so "WAL-MART", "WAL MART" and
  // "WALMART" all match (also CORE-MARK, SHAW'S, etc.)
  const s=up.replace(/\\bWAL[\\s.-]*MART\\b/g,"WALMART");
  const w=s.trim().split(/\\s+/);
  return (w[0]||"").replace(/[^A-Z0-9]/g,"");
}
function _zip5(z){ const d=String(z==null?"":z).replace(/[^0-9]/g,""); return d?parseInt(d.slice(0,5),10):NaN; }
function _apptVal(o){ const f=(o.fields||[]).find(x=>String(x[0]).toLowerCase()==="last drop appt date start"); return f?String(f[1]):""; }
function mergeStops(orders){
  // Orders sharing a TMS ID combine into ONE stop only when they are the same delivery:
  //   - same customer/banner  (first word of Last Drop Name, e.g. "WAL-MART")
  //   - same Last Drop Appt Date Start
  //   - ZIP codes within 1 of each other
  // Otherwise they stay separate; a TMS hitting separate stops gets a connecting line.
  // Orders without a TMS are never merged.
  const byTms=new Map(), loose=[];
  for(const o of (orders||[])){
    if(o.tms){ if(!byTms.has(o.tms)) byTms.set(o.tms,[]); byTms.get(o.tms).push(o); }
    else loose.push(o);
  }
  const result=[];
  byTms.forEach(arr=>{
    let clusters=[];
    for(const o of arr){
      const bn=_banner(o.name), ap=_apptVal(o), zp=_zip5(o.zip);
      const hit=clusters.filter(c=> c.banner===bn && bn!=="" && c.appt===ap &&
        c.zips.some(z=>Number.isFinite(z)&&Number.isFinite(zp)&&Math.abs(z-zp)<=1));
      let c;
      if(hit.length===0){ c={banner:bn,appt:ap,zips:[],items:[]}; clusters.push(c); }
      else { c=hit[0];                                   // bridge any clusters this order links
        for(let i=1;i<hit.length;i++){ const m=hit[i]; for(const z of m.zips)c.zips.push(z); for(const it of m.items)c.items.push(it); }
        clusters=clusters.filter(x=>x===c||hit.indexOf(x)===-1);
      }
      c.zips.push(zp); c.items.push(o);
    }
    for(const c of clusters) result.push(c.items.length>1 ? combineStop(c.items) : c.items[0]);
  });
  for(const o of loose) result.push(o);
  return result;
}

/* ---- in-page Excel loading: Choose File -> Load Data (reads the sheet in the browser) ---- */
const PICKS_CFG=[
 {key:"JONESTOWN",name:"Jonestown, PA",lat:40.4115,lng:-76.4814,match:["JONESTOWN"]},
 {key:"GOODYEAR",name:"Goodyear, AZ",lat:33.4353,lng:-112.3576,match:["GOODYEAR"]},
 {key:"DEKALB",name:"DeKalb, IL",lat:41.9294,lng:-88.7504,match:["DEKALB","DE KALB"]},
 {key:"GRAND PRAIRIE",name:"Grand Prairie, TX",lat:32.7459,lng:-96.9978,match:["GRAND PRAIRIE","PRAIRIE"]},
 {key:"MCDONOUGH",name:"McDonough, GA",lat:33.4473,lng:-84.1469,match:["MCDONOUGH","MCDONUGH","MCDON"]}
];
function _normZip(z){ if(z==null) return ""; const d=String(z).replace(/[^0-9]/g,""); return d?d.slice(0,5).padStart(5,"0"):""; }
function _matchPick(v){ if(v==null||v==="") return null; const s=String(v).toUpperCase().replace(/\\s+/g," ").trim();
  for(const p of PICKS_CFG){ for(const kw of (p.match||[p.key])){ if(s.indexOf(String(kw).toUpperCase())>=0) return p.key; } } return null; }
function _num(x){ if(x==null||x==="") return null; const n=parseFloat(String(x).replace(/,/g,"")); return isFinite(n)?n:null; }
const _MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function _serialToDate(n){ return new Date(Math.round((Number(n)-25569)*86400000)); }   // Excel serial -> JS Date (UTC)
function _fmtSerial(n){ const d=_serialToDate(n); if(isNaN(d)) return String(n);
  const base=_MON[d.getUTCMonth()]+" "+String(d.getUTCDate()).padStart(2,"0")+", "+d.getUTCFullYear();
  if(d.getUTCHours()===0&&d.getUTCMinutes()===0&&d.getUTCSeconds()===0) return base;
  let h=d.getUTCHours(), ap=h>=12?"PM":"AM"; h=h%12||12;
  return base+" "+String(h).padStart(2,"0")+":"+String(d.getUTCMinutes()).padStart(2,"0")+" "+ap; }
function _serialToISO(n){ const d=_serialToDate(n); if(isNaN(d)) return null;
  return d.getUTCFullYear()+"-"+String(d.getUTCMonth()+1).padStart(2,"0")+"-"+String(d.getUTCDate()).padStart(2,"0"); }
function _fmtCell(v,h){ if(v==null||v==="") return "";
  const hl=String(h).toLowerCase();
  if(typeof v==="number" && /date/i.test(hl)) return _fmtSerial(v);
  if(hl==="tms id"||hl==="shipment reference numbers"||hl==="last drop postal code") return String(v);
  if(typeof v==="number") return Number.isInteger(v)? v.toLocaleString() : v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});
  if(v instanceof Date && !isNaN(v)) return _MON[v.getUTCMonth()]+" "+String(v.getUTCDate()).padStart(2,"0")+", "+v.getUTCFullYear();
  return String(v); }
const _HKEYS={lat:["lat","latitude"],lng:["lng","long","lon","longitude"],
  status:["plannedorno","planned","planned or no","status"],name:["last drop name"],
  zip:["last drop postal code","postal code","zip","zip code"],pallet:["pallet spaces"],pallets:["pallets"],
  tms:["tms id","tmsid","tms","tms number","tms #"],
  pick:["first pick city","firstpickcity","first pick","pick city","origin city","first pick location","pick location"],
  weight:["weight (lb)","weight","weight(lb)","weight lb","weight lbs","total weight"],
  planstart:["last drop plan date start","plan date start","delivery date start","drop date start","deliver start"],
  planend:["last drop plan date end","plan date end","delivery date end","drop date end","deliver end"],
  pickdate:["first pick appt date start","first pick date","pick plan date","pick date","first pick appt","pick appt date"],
  loadgroup:["load group","loadgroup"]};
function buildPayloadFromWB(wb){
  const zipMap=new Map(); const zs=wb.Sheets["Zip"];
  if(zs){ const zr=XLSX.utils.sheet_to_json(zs,{header:1,raw:true,defval:null});
    for(let i=1;i<zr.length;i++){ const r=zr[i]; if(!r||r[0]==null) continue; const z=_normZip(r[0]);
      const la=parseFloat(r[1]),lo=parseFloat(r[2]); if(z&&isFinite(la)&&isFinite(lo)) zipMap.set(z,[la,lo]); } }
  const os=wb.Sheets["Orders"]||wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(os,{header:1,raw:true,defval:null});
  const headers=rows[0]||[]; const idx={};
  headers.forEach((h,i)=>{ const hl=String(h==null?"":h).trim().toLowerCase();
    for(const k in _HKEYS){ if(idx[k]==null && _HKEYS[k].indexOf(hl)>=0){ idx[k]=i; break; } } });
  const orders=[]; let skipped=0;
  for(let ri=1;ri<rows.length;ri++){ const row=rows[ri];
    if(!row || row.every(c=>c==null||c==="")) continue;
    let lat=idx.lat!=null?parseFloat(row[idx.lat]):NaN, lng=idx.lng!=null?parseFloat(row[idx.lng]):NaN;
    if(!(isFinite(lat)&&isFinite(lng))){ const z=idx.zip!=null?_normZip(row[idx.zip]):""; const c=zipMap.get(z);
      if(c){ lat=c[0]; lng=c[1]; } else { skipped++; continue; } }
    const status=idx.status!=null && row[idx.status]!=null ? String(row[idx.status]).trim().toUpperCase() : "";
    let tms=null; if(idx.tms!=null && row[idx.tms]!=null && row[idx.tms]!==""){ const t=row[idx.tms];
      tms=(typeof t==="number"&&Number.isInteger(t))?String(t):String(t).trim(); }
    const loadGroup=idx.loadgroup!=null && row[idx.loadgroup]!=null ? String(row[idx.loadgroup]).trim() : "";
    const fields=[];
    headers.forEach((h,i)=>{ if(h==null||h===""||i===idx.lat||i===idx.lng) return; fields.push([String(h), _fmtCell(row[i], h)]); });
    orders.push({ lat:lat, lng:lng, status:status,
      name: idx.name!=null && row[idx.name]!=null ? String(row[idx.name]) : "",
      zip: idx.zip!=null && row[idx.zip]!=null ? String(row[idx.zip]) : "",
      palletSpaces: idx.pallet!=null ? _num(row[idx.pallet]) : null,
      pallets: idx.pallets!=null ? _num(row[idx.pallets]) : null,
      weight: idx.weight!=null ? _num(row[idx.weight]) : null,
      tms: tms, pick: idx.pick!=null ? _matchPick(row[idx.pick]) : null, loadGroup: loadGroup,
      planStart: idx.planstart!=null && typeof row[idx.planstart]==="number" ? _serialToISO(row[idx.planstart]) : null,
      planEnd: idx.planend!=null && typeof row[idx.planend]==="number" ? _serialToISO(row[idx.planend]) : null,
      pickDate: idx.pickdate!=null && typeof row[idx.pickdate]==="number" ? _serialToISO(row[idx.pickdate]) : null,
      fields: fields });
  }
  return { updatedAt:new Date().toLocaleString(), stale:false,
    picks:PICKS_CFG.map(p=>({key:p.key,name:p.name,lat:p.lat,lng:p.lng})),
    palletThreshold:16, pollSeconds:45, tmsAvailable:idx.tms!=null, pickAvailable:idx.pick!=null,
    orders:orders, _skipped:skipped };
}
(function(){
  const fi=document.getElementById("fileInput"), btn=document.getElementById("loadDataBtn"), st=document.getElementById("uploadStatus");
  if(!fi||!btn) return;
  function up(m,c){ if(!st) return; st.textContent=m; st.style.fontSize="12px"; st.style.marginTop="6px"; st.style.lineHeight="1.4";
    st.style.color=(c==="err"?"#c62828":(c==="ok"?"#2e7d32":"#647280")); st.style.fontWeight=((c==="ok"||c==="err")?"600":"400"); }
  async function load(){
    const f=fi.files&&fi.files[0]; if(!f){ up("Choose the MapData Excel file first.","err"); return; }
    if(typeof XLSX==="undefined"){ up("Spreadsheet reader did not load.","err"); return; }
    up("Reading "+f.name+" ...");
    try{
      const buf=await f.arrayBuffer();
      const wb=XLSX.read(new Uint8Array(buf),{type:"array"});
      const payload=buildPayloadFromWB(wb);
      if(!payload.orders.length){ up("No delivery rows found - is this the MapData sheet (with an 'Orders' tab)?","err"); return; }
      window.__EMBEDDED_DATA__=payload; lastUpdatedAt=payload.updatedAt;
      didInitialFit=false; ingest(payload); didInitialFit=true;
      let msg="Loaded "+payload.orders.length+" loads from "+f.name+".";
      if(payload._skipped) msg+=" ("+payload._skipped+" rows had no location and were skipped.)";
      up(msg,"ok");
    }catch(err){ up("Could not read that file: "+((err&&err.message)||err),"err"); }
  }
  btn.addEventListener("click", load);
  fi.addEventListener("change", ()=>{ if(fi.files&&fi.files[0]) up('Ready: "'+fi.files[0].name+'". Click "Load Data".'); });
})();
/* ---- load row selection + CSV export ---- */
function updateExportBar(){
  const bar = document.getElementById("export-bar");
  const cnt = document.getElementById("exp-count");
  if (!bar) return;
  if (selectedTms.size === 0){ bar.classList.remove("show"); return; }
  bar.classList.add("show");
  // Count raw shipments (not merged stops) so the count matches what will export
  const raw = DATA._raw || DATA.orders;
  const n = raw.filter(o => o.tms && selectedTms.has(String(o.tms))).length;
  if (cnt) cnt.textContent = n + " shipment" + (n!==1?"s":"") + " selected (" + selectedTms.size + " load" + (selectedTms.size!==1?"s":"") + ")";
}
function exportSelectedLoads(){
  if (!selectedTms.size) return;
  const raw = DATA._raw || DATA.orders;
  const rows = raw.filter(o => o.tms && selectedTms.has(String(o.tms)));
  if (!rows.length) return;
  // Build header row from the fields of the first record + extra structured columns
  const extraHdrs = ["TMS ID","Pick DC","Status","Load Group","Plan Start","Plan End","Pick Date","Weight (lb)","Pallet Spaces","Pallets"];
  const fieldHdrs = (rows[0].fields||[]).map(f=>f[0]).filter(h => !extraHdrs.map(x=>x.toLowerCase()).includes(h.toLowerCase()));
  const allHdrs = [...extraHdrs, ...fieldHdrs];
  const fieldIdx = Object.fromEntries((rows[0].fields||[]).map((f,i)=>[f[0].toLowerCase(),i]));
  function fv(o, label){ const f=(o.fields||[]).find(x=>x[0].toLowerCase()===label.toLowerCase()); return f?f[1]:""; }
  const csvRows = [allHdrs];
  for (const o of rows){
    const extra = [
      o.tms||"",
      o.pick ? ((pickByKey.get(o.pick)||{}).name||o.pick) : "",
      o.status||"", o.loadGroup||"",
      o.planStart||"", o.planEnd||"", o.pickDate||"",
      o.weight!=null?o.weight:"", o.palletSpaces!=null?o.palletSpaces:"", o.pallets!=null?o.pallets:""
    ];
    const fieldVals = fieldHdrs.map(h=>fv(o,h));
    csvRows.push([...extra, ...fieldVals]);
  }
  const csv = csvRows.map(r=>r.map(v=>'"'+String(v==null?"":v).replace(/"/g,'""')+'"').join(",")).join("\\r\\n");
  const blob = new Blob(["﻿"+csv], {type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href=url; a.download="SelectedLoads.csv"; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}
(function wireExportBar(){
  document.addEventListener("change", e=>{
    const chk = e.target.closest(".load-chk"); if (!chk) return;
    const t = String(chk.dataset.tms||""); if(!t) return;
    const row = chk.closest(".loadrow");
    if (chk.checked){ selectedTms.add(t); if(row) row.classList.add("sel-row"); }
    else { selectedTms.delete(t); if(row) row.classList.remove("sel-row"); }
    updateExportBar();
  });
  document.addEventListener("click", e=>{
    const b=e.target.closest("#exp-btn"); if(b){ exportSelectedLoads(); return; }
    const c=e.target.closest("#exp-clear"); if(c){ selectedTms.clear(); updateExportBar(); renderLoads(); }
  });
})();
/* ---- consolidation match finder ---- */
let _matchLayer = null, _matchTms = null;
function clearMatchHighlight(){
  if (_matchLayer){ _matchLayer.clearLayers(); }
  document.querySelectorAll('.match-results').forEach(el=>el.remove());
  document.querySelectorAll('.tms-findmatch.active').forEach(el=>el.classList.remove('active'));
  _matchTms = null;
}
function findConsolidationMatches(tms){
  const btn = document.querySelector('.tms-findmatch[data-tms="'+tms+'"]');
  if (_matchTms === String(tms)){ clearMatchHighlight(); return; }
  clearMatchHighlight();
  _matchTms = String(tms);
  if (btn) btn.classList.add('active');

  const raw = DATA._raw || DATA.orders;
  const maxW = (SETTINGS && SETTINGS.maxWeight) || 44000;
  const maxP = (SETTINGS && SETTINGS.maxPallets) || 26;

  const mine = raw.filter(o => String(o.tms) === _matchTms);
  if (!mine.length){ clearMatchHighlight(); return; }

  const myWt  = mine.reduce((s,o)=>s+(o.weight||0),0);
  const myPal = mine.reduce((s,o)=>s+(o.palletSpaces||0),0);
  const myPick = mine[0].pick;

  // Effective window: latest start → earliest end across all stops on this load
  const myWins = mine.map(o=>orderWindow(o)).filter(Boolean);
  const myWs = myWins.length ? Math.max(...myWins.map(w=>w.ws)) : null;
  const myWe = myWins.length ? Math.min(...myWins.map(w=>w.we)) : null;

  // Group all other orders by TMS
  const byTms = new Map();
  for (const o of raw){
    if (!o.tms || String(o.tms)===_matchTms) continue;
    if (!byTms.has(String(o.tms))) byTms.set(String(o.tms),[]);
    byTms.get(String(o.tms)).push(o);
  }

  const matches = [];
  byTms.forEach((orders, otherTms)=>{
    const otherPick = orders[0].pick;
    if (myPick && otherPick && myPick !== otherPick) return;   // different DC → skip

    const oWt  = orders.reduce((s,o)=>s+(o.weight||0),0);
    const oPal = orders.reduce((s,o)=>s+(o.palletSpaces||0),0);
    if (myWt+oWt > maxW || myPal+oPal > maxP) return;          // over truck limit → skip

    // Delivery window overlap
    if (myWs !== null && myWe !== null){
      const oWins = orders.map(o=>orderWindow(o)).filter(Boolean);
      if (oWins.length){
        const oWs = Math.max(...oWins.map(w=>w.ws));
        const oWe = Math.min(...oWins.map(w=>w.we));
        if (Math.max(myWs,oWs) > Math.min(myWe,oWe)) return;  // no overlap → skip
      }
    }

    const oWins2 = orders.map(o=>orderWindow(o)).filter(Boolean);
    const oWs2 = oWins2.length ? Math.max(...oWins2.map(w=>w.ws)) : null;
    const oWe2 = oWins2.length ? Math.min(...oWins2.map(w=>w.we)) : null;
    const cWs = (myWs!==null&&oWs2!==null) ? Math.max(myWs,oWs2) : (myWs??oWs2);
    const cWe = (myWe!==null&&oWe2!==null) ? Math.min(myWe,oWe2) : (myWe??oWe2);
    const oName = [...new Set(orders.map(o=>o.name).filter(Boolean))].slice(0,2).join(' + ');
    const oPkNm = otherPick ? ((pickByKey.get(otherPick)||{}).name||otherPick) : '—';
    matches.push({ tms:otherTms, orders, wt:oWt, pal:oPal, name:oName, pickName:oPkNm,
      cWt:myWt+oWt, cPal:myPal+oPal, cWs, cWe });
  });

  // Sort: highest combined pallet utilisation first
  matches.sort((a,b)=>b.cPal-a.cPal);

  // Ring matching destination markers on the map with gold dashes
  if (!_matchLayer) _matchLayer = L.layerGroup().addTo(map);
  const matchSet = new Set(matches.map(m=>m.tms));
  for (const rec of (markerRecs||[])){
    if ((rec.group.orders||[]).some(o=>o.tms && matchSet.has(String(o.tms)))){
      L.circleMarker([rec.latlng[0],rec.latlng[1]],{
        radius:20, color:'#f59e0b', weight:3, fill:false, opacity:0.9,
        dashArray:'6,4', className:'match-map-ring'
      }).addTo(_matchLayer);
    }
  }

  // Build results panel
  const _fd = dn=>{ if(dn==null) return '—'; const d=new Date(dn*86400000); return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getUTCMonth()]+" "+d.getUTCDate()+", "+d.getUTCFullYear(); };
  let rh = '<div class="match-results">';
  if (!matches.length){
    rh += '<div style="color:#92400e;font-size:12.5px;">No other loads from the same DC fit on the same truck with an overlapping delivery window.</div>';
  } else {
    rh += '<div style="font-weight:700;font-size:13px;color:#92400e;margin-bottom:6px;">✅ '+matches.length+' load'+(matches.length>1?'s':'')+' could ship with this one (gold rings on map):</div>';
    const hasRates = !!(DATA.rates && Object.keys(DATA.rates).length);
    rh += '<table><tr class="match-hdr"><th>TMS #</th><th>Destination</th><th style="text-align:right">Wt (lb)</th><th style="text-align:right">Pal sp</th><th>Ship window</th>'+(hasRates?'<th>Cheapest carrier</th>':'')+'<th style="text-align:right">Truck fill</th></tr>';
    for (const m of matches){
      const wFill  = Math.round(m.cWt/maxW*100);
      const pFill  = Math.round(m.cPal/maxP*100);
      const fill   = Math.max(wFill,pFill);
      const fillCl = fill>90?'#c62828':fill>75?'#ef6c00':'#2e7d32';
      const winStr = (m.cWs!=null&&m.cWe!=null) ? _fd(m.cWs)+(m.cWs!==m.cWe?' – '+_fd(m.cWe):'') : '—';
      // Get representative ZIP from this match's orders for rate lookup
      const repOrder = m.orders.find(o=>o.zip) || mine.find(o=>o.zip) || {};
      const repZip = repOrder.zip || null;
      const repLat = repOrder.lat || null; const repLng = repOrder.lng || null;
      const rateCl = hasRates ? '<td style="white-space:nowrap;font-size:11.5px">'+ratesBadge(myPick, repZip, repLat, repLng)+'</td>' : '';
      rh += '<tr><td><span class="tms-chip" data-tms="'+esc(m.tms)+'">'+esc(m.tms)+'</span></td>'+
            '<td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(m.name)+'">'+esc(m.name)+'</td>'+
            '<td style="text-align:right;white-space:nowrap">'+Math.round(m.wt).toLocaleString()+' lb</td>'+
            '<td style="text-align:right;white-space:nowrap">'+(Math.round(m.pal*10)/10)+' sp</td>'+
            '<td style="white-space:nowrap">'+esc(winStr)+'</td>'+
            rateCl+
            '<td style="text-align:right;font-weight:700;color:'+fillCl+';white-space:nowrap">'+
              '<div class="fill-bar" style="width:'+Math.min(fill,100)+'px;background:'+fillCl+'"></div> '+fill+'%'+
            '</td></tr>';
    }
    rh += '</table>';
    if(hasRates) rh += '<div style="font-size:11px;color:#92400e;margin-top:4px">&#128176; Rate = cheapest TL carrier for this origin → destination state. Hover for top 3. CPM rates shown as $/mi (actual cost depends on mileage).</div>';
    rh += '<div style="font-size:11px;color:#92400e;margin-top:4px">Truck limits: '+maxW.toLocaleString()+' lb · '+maxP+' pallet sp · Same pick DC only · Delivery windows must overlap</div>';
  }
  rh += '</div>';

  // Append to the visible TMS banner
  const banner = document.querySelector('#loads-body .dc-banner[style*="display"]:not([style*="display: none"]):not([style*="display:none"])');
  if (banner){ banner.insertAdjacentHTML('beforeend', rh); }
}
function buildMarkers(){'''

def main():
    print("Reading workbook...")
    payload = build_payload()

    print("Reading viewer template...")
    with open(SRC_HTML, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("\r\n", "\n")

    # combined ingest patch (insert load-group / date setup) + add new functions before buildMarkers
    ingest_old = ("  buildTabs();\n  buildMarkers();\n  applyFilters(!didInitialFit);")
    ingest_new = ("  buildLoadGroups();\n  buildTabs();\n  buildLoadGroupTabs();\n"
                  "  setupDateBounds();\n  buildMarkers();\n  applyFilters(!didInitialFit);")

    failures = []
    for name, old, new in PATCHES:
        if old not in html:
            failures.append(name)
            continue
        html = html.replace(old, new, 1)

    # ingest patch
    if ingest_old in html:
        html = html.replace(ingest_old, ingest_new, 1)
    else:
        failures.append("ingest setup")

    # inject new JS functions before buildMarkers
    if "function buildMarkers(){" in html:
        html = html.replace("function buildMarkers(){", NEW_FUNCS, 1)
    else:
        failures.append("buildMarkers anchor")

    # replace embedded data line
    data_js = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    lines = html.split("\n")
    replaced = False
    for i, ln in enumerate(lines):
        if "window.__EMBEDDED_DATA__" in ln and "<script>" in ln:
            lines[i] = "<script>window.__EMBEDDED_DATA__ = " + data_js + ";</script>"
            replaced = True
            break
    if not replaced:
        failures.append("embedded data line")
    html = "\n".join(lines)

    # inject the SheetJS reader + new CSS at </head> (done last so the library can't shadow other anchors)
    if "</head>" in html:
        print("Injecting SheetJS + CSS...")
        sheetjs_js = extract_sheetjs()
        sheetjs_tag = ("<script>" + sheetjs_js + "</script>\n") if sheetjs_js else ""
        head_inject = sheetjs_tag + NEW_CSS
        html = html.replace("</head>", head_inject, 1)
    else:
        failures.append("</head> for SheetJS+CSS")

    if failures:
        print("\n  *** PATCH FAILURES (anchors not found): ***")
        for fname in failures:
            print("     -", fname)
        print("  Aborting; output NOT written.")
        sys.exit(2)

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n  OK -> {OUT_HTML}")
    print(f"  size: {len(html):,} chars, orders: {len(payload['orders'])}")

    # also refresh the zip (best for sending in Teams/email -- recipients download, unzip, double-click)
    import zipfile
    zip_path = os.path.join(os.path.dirname(OUT_HTML), "DeliveryMap.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(OUT_HTML, arcname=os.path.basename(OUT_HTML))
    print(f"  zipped -> {zip_path}  ({os.path.getsize(zip_path)/1048576:.1f} MB)")

if __name__ == "__main__":
    main()
