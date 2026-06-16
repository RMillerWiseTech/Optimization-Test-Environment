# Delivery Map — Project Reference

A self-contained, shareable HTML "Delivery Map" generated from a SharePoint Excel sheet. It has three views — a **map**, a **Loads by DC** consolidation view, and a **Rules & Limits** page — used to find which freight loads can ship together on one truck.

---

## TL;DR — how to update it

1. Double-click **`Update Delivery Map.bat`** on the Desktop. (It runs the Python generator.)
2. When it finishes, two files are refreshed on the Desktop:
   - **`DeliveryMap (all columns + filters).html`** ← open this to view
   - **`DeliveryMap.zip`** ← **send THIS** in Teams/email (recipients download + open in a browser)
3. Already had the map open? **Close the tab and reopen** (or Ctrl+F5) — an open tab won't auto-refresh.

> The generator reads the live data straight from SharePoint, so just re-running it picks up the latest sheet.

---

## Key files (all on the Desktop unless noted)

| File | Role |
|------|------|
| `Rebuild Delivery Map.py` | **Source of truth.** Reads the xlsx, patches a base HTML viewer, writes the final HTML + zip. All changes are made here, then re-run. |
| `Update Delivery Map.bat` | One double-click runner for the generator. |
| `DeliveryMap (all columns + filters).html` | Generated viewer (open this). |
| `DeliveryMap.zip` | Generated shareable copy (send this). |
| `delivery map text.txt` | The base viewer template the generator patches (read-only reference). |

**Data source (live, SharePoint-synced):**
`...\LS\Billy\Jonestown Map\MAP JSON\MapData sheet.xlsx`

---

## How the generator works (for making changes)

`Rebuild Delivery Map.py` does NOT hand-write HTML. It:
1. Reads the xlsx with `openpyxl` and builds the embedded data.
2. Takes the base viewer (`delivery map text.txt`) and applies a list of **string-replacement PATCHES** — each a `(name, old, new)` tuple.
3. Injects extra CSS (`NEW_CSS`) and JS (`NEW_FUNCS`), inlines SheetJS + Leaflet for offline use, and writes the HTML, then zips it.

**To change behavior:** edit the relevant PATCH / `NEW_FUNCS` / `NEW_CSS` in the `.py`, then re-run. Each patch's `old` string must match the base/current text exactly.

SheetJS is cached at `%LOCALAPPDATA%\dm_sheetjs_cache.js` (the generator no longer needs `Map Maker.html`).

---

## The three views

### Map
Leaflet map of delivery locations, colored by pick (ship-from) origin, with optional route lines.

### Loads by DC
Destinations grouped by pick city (e.g. *Jonestown, PA*), each expandable to its loads. Highlights destinations that are **combinable** (green).

**Load rows show:** `reference → TMS# → status → location name → weight · pallet sp · pick appt (with time) · delivery appt (with time) · delivery window (plan dates)`

- The **TMS number** (right after the reference) is a clickable blue chip.
- Click it → highlights **every shipment on that load** (across destinations) and shows a **summary banner inside that pick-city's section** (so you don't scroll to the top). Click again / "clear ×" to dismiss.
- The banner shows totals (shipments · locations · weight · pallet sp · delivery day) and, when a load hits more than one location, a **per-location breakdown** (weight + pallet sp each).
- The banner **flags same-location orders whose delivery windows don't overlap**, listing exactly which orders fall in each window.

Filters: a **per-DC button row** (above the search line) to show only one pick city's loads ("All DCs" resets), plus keyword search, ZIP, min loads, min pallet sp, "only combinable," and sortable column headers.

### Rules & Limits
- **Truck limits:** max weight, max pallet spaces.
- **Delivery window rules** (per customer, matched by name "contains"). A **negative "End"** value means *"deliver up to N days BEFORE the Last Drop Plan End date"* — the End date is the latest deliverable day (included), Start is ignored; supports business vs. calendar days; floored by the pick date.
- **Per-customer overrides**, **Same-location customers** (name aliases).
- **Same location (merge ZIPs):** manually treat one delivery point that uses multiple ZIPs (same name, different ZIPs — e.g. ADUSA at 18017 + 18020) as a **single location** in the load summary. Add the name + each ZIP; leave ZIPs blank to merge all ZIPs for that name.
- **Consolidated Locations:** per origin, define delivery locations that consolidate well into lanes.

---

## "Combinable" (green) logic

A destination group turns green when **all** hold:
- **2+ distinct TMS loads** (one TMS = already one consolidated load).
- **Total weight < max weight** and **total pallet spaces < max pallets**.
- A **common delivery day exists** across the rule-adjusted delivery windows of its shipments.

Consolidation analyzes **raw shipments** (true per-shipment windows), and clusters a destination's loads by **common-delivery-day window** so one off-date load can't block the rest. Delivery windows are rule-adjusted (see negative-End rule above) and floored by pick date.

---

## Sharing notes

- Send the **`.zip`** (or the `.html`). Recipients must **download and open in a browser** — opening it inside Teams/SharePoint preview shows no interactivity (no JS).
- Shared copies seed their rules from the baked-in defaults, then each recipient's edits save locally in their own browser.

---

## Gotchas

- An already-open browser tab does **not** refresh — close + reopen or Ctrl+F5.
- If data looks stale, confirm the generator's `SRC_XLSX` still points at the SharePoint sheet path above.
- `.claude/launch.json` intentionally contains only the `load-weather-map` config — leave it unless asked. (Temporary `dm-preview` configs used during development are added via a `.bak` backup and restored afterward.)
