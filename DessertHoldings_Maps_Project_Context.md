# Dessert Holdings — Delivery Map Project Context
### Load this file into Claude Code to work on the Maps project

---

## What This Project Is

An interactive HTML delivery map for Dessert Holdings logistics operations. A Python script (`Rebuild Delivery Map.py`) reads daily order data from an Excel file and generates a self-contained HTML map with carrier rate comparisons, pick location filters, and route alerts.

**GitHub Repo:** `https://github.com/RMillerWiseTech/Optimization-Test-Environment`
**Working Branch:** `claude/stoic-albattani-hv679u`

---

## How to Get Started

1. Open [claude.ai/code](https://claude.ai/code)
2. Clone the repo: `https://github.com/RMillerWiseTech/Optimization-Test-Environment`
3. Check out the working branch: `claude/stoic-albattani-hv679u`
4. To run an update: upload a new `Claude_Optimization.xlsx` and say "remove all old order data and update with the attached file"

---

## Key Files

| File | Purpose |
|---|---|
| `Rebuild Delivery Map.py` | Main generator — all logic lives here |
| `MapData sheet.xlsx` | Current order data (replaced on each update) |
| `DH_CURRENT_RATES_7.30.26.xlsx` | TL and LTL carrier rates |
| `stop_locations.json` | Geocoding DB for drop locations (7,464 entries) |
| `zip_coords.json` | ZIP/postal code coordinate fallback (42,411 entries) |
| `delivery map text.txt` | Base HTML template — do not edit |
| `DeliveryMap (all columns + filters).html` | Latest built map |
| `DeliveryMap.zip` | Latest map zipped for sharing |

---

## How the Generator Works

- Reads order xlsx → geocodes each drop location via `stop_locations.json` (by location reference number, then name+postal fallback), then `zip_coords.json` as final fallback
- Applies exclusion rules → builds JSON payload → injects into HTML template via string PATCHES
- Outputs a self-contained HTML file + zip

### Geocoding Priority
1. `stop_locations.json` by Location Reference Number
2. `stop_locations.json` by Name + Postal Code
3. `zip_coords.json` by ZIP/postal code (42,411 US + Canadian entries)

---

## Exclusion Rules (Baked In)

- **Excluded carrier SCACs:** `2ACC`, `1TOC`
- **Atlanta → any Walmart:** excluded
- **Any Walmart in southeastern US states:** excluded (FL, GA, AL, MS, SC, TN, AR, LA, NC, VA, WV, KY)
- **LTL eligibility:** shipments with **more than 10 pallets** are ineligible for LTL

---

## Pick Locations

**Dessert Holdings (DH) Facilities** (group: `dh`):
- Aurora CO, Delta BC, Humble TX, Kennesaw GA, Le Center MN, London ON, Newburyport MA, Pembroke NC, St Paul MN

**3PL Facilities** (group: `3pl`):
- Atlanta GA, Bethlehem PA, Bolingbrook IL, Brighton CO, Calgary AB, Carthage MO, Franklin IN, Golden Valley MN, Ingersoll ON, Smyrna GA, Surrey BC

---

## Rate Logic

**Rates file:** `DH_CURRENT_RATES_7.30.26.xlsx`

- **TL rates** — sheet "TL CURRENT RATES": CPM (cost-per-mile) or FLT (flat) basis. Estimated cost = haversine distance × 1.3 road factor × CPM rate, floored at min_cost. Shown as "(est.)"
- **LTL rates** — sheet "LTL CURRENT RATES": data starts row 4. Stores minimum floor charges only (cost_per = 0). Requires exact destination city match. Shown as "(min)"
- Map shows: cheapest LTL carrier + SCAC + rate (purple badge) and cheapest TL carrier + SCAC + rate (blue badge)

---

## Map Features

- **Leaflet.js** map with OSM + Google Satellite tile layers
- **Clustered destination pins** color-coded by status (green=Planned, red=Unplanned, orange=mixed)
- **Pick location pins** — click to filter orders by origin
- **Dropdown filters** — DH facilities, 3PL facilities, status, load group, date range
- **Loads by DC tab** — consolidation analysis, sort by any column, export to Excel
- **Rate badge** per destination — LTL and TL best rates
- **Route alert button** — calls OpenRouteService API (free, requires API key in `NEW_FUNCS` constant `ORS_API_KEY`)
- **Consolidation matching** — finds unplanned loads that could be added to planned loads

---

## Order File Column Mapping

The generator uses aliases to handle different column name formats:

| Field | Accepted Column Names |
|---|---|
| Drop name | "Last Drop Name", "Drop Location Name" |
| Drop postal | "Last Drop Postal Code", "Drop Location Postal Code" |
| Drop city | "Drop Location City" |
| Drop locref | "Drop Location Reference Number" |
| Status | "PLANNEDORNO", "Load Status" |
| Pallets | "Pallets" |
| Pallet spaces | "Pallet Spaces", "Shipment Pallet Spaces" |
| TMS ID | "TMS ID" |
| Pick city | "First Pick City", "Pick Location City" |
| Weight | "Weight (lb)" |

---

## Order Data Source

Daily order file comes from SharePoint:
- **Site:** `wisetechglobal.sharepoint.com/sites/w-e2openlaas`
- **Folder:** `DHDocuments/Logistics Analyst/DATA DUMPS`

---

## Common Tasks

### Update with new order data
Upload the new xlsx and say:
> "Remove all order information from before and update the order information with the attached file"

### Change LTL pallet limit
In `Rebuild Delivery Map.py`, find:
```javascript
const ltlEligible = !(totalPallets > 10);
```
Change `10` to the new limit.

### Add a new pick location
In `Rebuild Delivery Map.py`, find `DEFAULT_PICKS` list and add an entry:
```python
{"key": "CITYNAME", "name": "City, ST", "lat": 00.0000, "lng": -00.0000,
 "match": ["KEYWORD"], "group": "dh"},  # or "3pl"
```
Keep alphabetical order within each group.

### Add missing Canadian postal codes
If orders are skipped for missing coords, add to `zip_coords.json`:
```python
{"POSTAL": [lat, lng]}
```

### Add ORS route alert API key
In `Rebuild Delivery Map.py`, find in `NEW_FUNCS`:
```javascript
const ORS_API_KEY = "YOUR_ORS_KEY_HERE";
```
Replace with a free key from [openrouteservice.org](https://openrouteservice.org).

---

## After Any Change

Always rebuild and push:
```bash
python3 "Rebuild Delivery Map.py"
git add -A
git commit -m "Description of change"
git push -u origin claude/stoic-albattani-hv679u
```

---

## Contacts

- **Project owner:** Rebecca Miller — `rebecca.miller@wisetechglobal.com`
- **Backup:** Daniel Bolema — `daniel.bolema@wisetechglobal.com`
