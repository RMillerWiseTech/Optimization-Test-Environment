# Data sheet format

The generator reads an Excel workbook (`.xlsx`). Put it in this folder as **`MapData sheet.xlsx`**
(or repoint `SRC_XLSX` in `Rebuild Delivery Map.py`).

- The data must be on a tab named **`Orders`** (if not present, the **first** tab is used).
- **Row 1 = headers.** Each row below = one shipment.
- Column matching is **by header name, case-insensitive.** Order of columns doesn't matter.
- Any columns **not** in the list below are still shown automatically in the map's popup detail.

## Required

| Column | Why |
|--------|-----|
| `lat`  | Delivery latitude (decimal). **Rows without a valid `lat`/`lng` are skipped.** |
| `lng`  | Delivery longitude (decimal). |

> Tip: geocode each delivery's ZIP/address to lat/lng in the sheet before building.

## Used (recommended — features degrade gracefully if missing)

| Column | Used for |
|--------|----------|
| `Last Drop Name` | Delivery location name (grouping, labels) |
| `Last Drop Postal Code` | ZIP (grouping, "same location" rules) |
| `Last Drop Plan Date Start` | Delivery window start (date) |
| `Last Drop Plan Date End` | Delivery window end (date); drives the "−N days before End" rule |
| `Last Drop Appt Date Start` | Delivery appointment (date **+ time**) shown on load rows |
| `First Pick Appt Date Start` | Pick appointment (date **+ time**); also the "can't ship before pick" floor |
| `First Pick City` | Matched against your `DEFAULT_PICKS` to assign the ship-from origin |
| `TMS ID` | The load/booking ID — orders sharing one ID are treated as one consolidated load |
| `Weight (lb)` | Load weight (totals, truck-limit checks) |
| `Pallet Spaces` | Trailer positions (totals, truck-limit checks) |
| `Pallets` | Pallet count (shown in detail) |
| `PLANNEDORNO` | Status badge / filter — values like `PLANNED`, `UNPLANNED`, `POOL` |
| `Load Group` | Optional grouping buttons on the map |
| `Shipment Reference Numbers` | Order reference shown on each load row |

## Dates & numbers

- **Dates** should be real Excel dates (not text). Date-only columns show as a day; the appointment
  columns keep the **time** when present.
- **Weight / Pallet Spaces / Pallets** should be numbers.
- `TMS ID`, `Shipment Reference Numbers`, and `Last Drop Postal Code` are treated as identifiers
  (no thousands separators).
