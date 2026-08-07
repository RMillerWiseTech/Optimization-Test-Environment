# Delivery Map — Onboarding Guide for Daniel Bolema

Welcome! This guide will get you up to speed on the Dessert Holdings Delivery Map project so you can run updates while Rebecca is out.

---

## What This Project Does

Generates a self-contained interactive HTML map (`DeliveryMap.zip`) showing all outbound shipment orders by destination. The map includes:
- Drop location pins color-coded by status (Planned / Unplanned)
- Pick location filters (DH facilities + 3PL locations)
- LTL and TL rate comparison per lane (cheapest carrier, SCAC, and rate)
- Route alert check button powered by OpenRouteService
- Loads by DC tab for consolidation analysis
- Export to Excel

---

## How to Run an Update

1. **Open this project in Claude Code**
   - Go to [claude.ai/code](https://claude.ai/code)
   - Clone or open the repo: `https://github.com/RMillerWiseTech/Optimization-Test-Environment`

2. **Get the latest order file**
   - The file comes from SharePoint: `wisetechglobal.sharepoint.com/sites/w-e2openlaas` → `DHDocuments/Logistics Analyst/DATA DUMPS`
   - It's an `.xlsx` file named something like `Claude_Optimization.xlsx`

3. **Upload the file and ask Claude to update**
   - Drag the xlsx file into the Claude Code chat
   - Type: `remove all order information from before and update the order information with the attached file`
   - Claude will rebuild the map automatically (takes ~30 seconds)

4. **Download the result**
   - Claude will deliver `DeliveryMap.zip`
   - Extract it and open the HTML file in any browser

---

## Key Files in This Repo

| File | What it is |
|---|---|
| `Rebuild Delivery Map.py` | The map generator — all logic lives here |
| `MapData sheet.xlsx` | Current order data (replaced on each update) |
| `DH_CURRENT_RATES_7.30.26.xlsx` | TL and LTL carrier rates |
| `stop_locations.json` | Geocoding DB for drop locations (7,464 entries) |
| `zip_coords.json` | ZIP/postal code coordinate fallback (42,411 entries) |
| `delivery map text.txt` | Base HTML template (do not edit) |
| `DeliveryMap (all columns + filters).html` | Latest built map |
| `DeliveryMap.zip` | Latest map zipped for sharing |

---

## Current Rules Baked Into the Map

- **Excluded carriers:** 2ACC, 1TOC (any lane)
- **Excluded lanes:** Atlanta origin → any Walmart; any Walmart in southeastern US states
- **LTL eligibility:** Shipments with **more than 10 pallets** are ineligible for LTL (shown as "LTL: >10 pallets")

---

## Pick Locations

**Dessert Holdings (DH) Facilities:**
Aurora CO, Delta BC, Humble TX, Kennesaw GA, Le Center MN, London ON, Newburyport MA, Pembroke NC, St Paul MN

**3PL Facilities:**
Atlanta GA, Bethlehem PA, Bolingbrook IL, Brighton CO, Calgary AB, Carthage MO, Franklin IN, Golden Valley MN, Ingersoll ON, Smyrna GA, Surrey BC

---

## Git Branch

All work goes to: `claude/stoic-albattani-hv679u`

After Claude rebuilds the map, it will commit and push automatically. You don't need to do anything with git manually.

---

## Questions?

Reach out to Rebecca Miller (`rebecca.miller@wisetechglobal.com`) or ask Claude directly — the full project history is in the conversation context.
