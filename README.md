# Delivery Map — Build Kit

A self-contained tool that turns an outbound-shipment spreadsheet into a shareable,
offline **Delivery Map** (map view + "Loads by DC" consolidation view + Rules & Limits).
This kit is **data-free** — drop in your own sheet and build your own map.

---

## Setup (one time)

1. **Install Python 3** from <https://www.python.org/downloads/> — during install, tick **"Add Python to PATH."**
2. **Install the one dependency** — open a terminal/Command Prompt and run:
   ```
   pip install openpyxl
   ```
3. **Add your data sheet.** Put your Excel file in this folder named **`MapData sheet.xlsx`**
   (tab named **`Orders`**), with the columns listed in **`DATA-FORMAT.md`**.
   *(Or edit `SRC_XLSX` near the top of `Rebuild Delivery Map.py` to point anywhere.)*
4. **Set your ship-from origins.** In `Rebuild Delivery Map.py`, edit **`DEFAULT_PICKS`**
   (near the top) to your own distribution centers — `key`, `name`, `lat`, `lng`, and the
   `match` keywords that identify them in your "First Pick City" column.

## Build the map

- **Double-click `Update Delivery Map.bat`** (or run `python "Rebuild Delivery Map.py"`).
- It produces two files in this folder:
  - **`DeliveryMap (all columns + filters).html`** ← open this in a web browser
  - **`DeliveryMap.zip`** ← this is the shareable copy (recipients download + open in a browser)

Re-run any time your sheet changes. If the map is already open in a tab, **close and reopen** (or Ctrl+F5) to see the update.

---

## Customizing for your team

- **Truck limits & delivery-window rules** are edited **inside the map** on the **Rules & Limits** tab
  (max weight, max pallet spaces, per-customer delivery windows, same-location ZIP merges, consolidation lanes).
  These save in each user's browser.
- **Deeper changes** (columns, labels, logic) are made in `Rebuild Delivery Map.py`, which patches the
  base viewer `delivery map text.txt`. See **`ONBOARDING.md`** for how the generator and patches work.

## What's in this kit

| File | Role |
|------|------|
| `Rebuild Delivery Map.py` | The generator (reads your sheet, builds the HTML + zip). **Source of truth.** |
| `delivery map text.txt` | Base viewer template the generator patches. Do not rename. |
| `sheetjs.js` | Bundled Excel-reader library (lets the map load sheets in-browser, and the build run offline). Keep it. |
| `Update Delivery Map.bat` | One double-click runner for the generator. |
| `DATA-FORMAT.md` | The exact columns your sheet needs. |
| `ONBOARDING.md` | Deeper reference: architecture, features, and how to modify. |
| `README.md` | This file. |

> Add your own `MapData sheet.xlsx` — the kit ships without data on purpose.

---

## Notes / gotchas

- **Open in a real browser** (Chrome/Edge). Opening the HTML inside a Teams/SharePoint preview shows no
  interactivity (it needs JavaScript). For sharing, send the **`.zip`**.
- An already-open browser tab won't auto-refresh — close + reopen or **Ctrl+F5**.
- If you see "SheetJS not found," make sure **`sheetjs.js`** is still in this folder.
- If you see a `ModuleNotFoundError: openpyxl`, run `pip install openpyxl` again (step 2).
