# Houspire Budget Generator — Claude Code Handoff

## What You Are Building

An internal Houspire web app. Staff enters client details + uploads room renders → Claude API generates a detailed BOQ and a real local vendor list → downloads two Excel files OR a branded DOCX + PDF.

**Single API key required: Anthropic only. No other services.**

---

## Decisions Already Made — Do Not Re-debate

| Decision | Answer |
|----------|--------|
| Framework | Streamlit (Python) |
| BOQ generation | Claude API with rate library in system prompt — no database |
| Vendor generation | Claude API with `web_search` tool — live Google search |
| Branded output | Node.js `docx` npm → DOCX, LibreOffice → PDF |
| Tiers | **Mid-tier and Premium only** — never economy/standard/luxury |
| City multipliers | Applied silently to rates — never shown as a row in the BOQ |
| Excel format | Plain Calibri 12, zero formatting, formula `=D{row}*E{row}` in Amount |
| Footer rows | None — no subtotal, GST, contingency, grand total |
| Vendor radius | 10 km primary, 15 km fallback, 20 km hard limit |
| Vendors per category | 3–5, sorted nearest-first |
| BOQ granularity | Sample-level (matches vendor site-quote — each point, each piece separately) |
| Deploy target | Replit |

---

## Project Structure

```
houspire_budget_generator/
├── app.py                    # Streamlit UI
├── config.py                 # Cities, multipliers, tier list
├── requirements.txt
├── .env.example
├── recalc.py                 # Formula verifier — must return 0 errors before delivery
├── modules/
│   ├── __init__.py
│   ├── room_analyzer.py      # Claude Vision → room type + sqft + design elements
│   ├── boq_generator.py      # Claude API → BOQ rows + rate sources
│   ├── vendor_finder.py      # Claude API + web_search → vendor rows + notes
│   ├── budget_excel.py       # openpyxl plain BOQ Excel
│   ├── vendor_excel.py       # openpyxl plain Vendor Excel
│   ├── docx_generator.py     # Node.js docx npm → branded DOCX
│   └── pdf_converter.py      # LibreOffice → PDF
└── tests/
    ├── test_room_analyzer.py
    ├── test_boq_generator.py
    ├── test_budget_excel.py
    ├── test_vendor_excel.py
    └── test_docx_generator.py
```

---

## Environment Setup

```bash
pip install streamlit anthropic openpyxl python-dotenv Pillow --break-system-packages
npm install -g docx
apt-get install -y libreoffice   # or: snap install libreoffice
```

`.env`:
```
ANTHROPIC_API_KEY=sk-ant-...
```

---

## config.py

```python
TIERS = ["Mid-tier", "Premium"]

CITIES_WITH_MULTIPLIERS = {
    "Hyderabad": 1.00, "Davanagere": 0.90, "Kathua": 0.88,
    "Jaipur": 0.92, "Sangli": 0.93, "Bhopal": 0.93, "Hubli": 0.92,
    "Bodeli": 0.88, "Gandhinagar": 0.96, "Jorhat": 0.95,
    "Nashik": 0.95, "Trivandrum": 0.95, "Vadodara": 0.98,
    "Visakhapatnam": 0.98, "Kolkata": 1.02, "Pune": 1.08,
    "Chennai": 1.10, "Bangalore": 1.12, "Noida": 1.18, "Delhi": 1.18,
    "Thane": 1.20, "Mumbai": 1.25, "Other": None,
}
CITIES = list(CITIES_WITH_MULTIPLIERS.keys())

ROOM_TYPES = [
    "Living Room", "Master Bedroom", "Bedroom", "Kitchen",
    "Bathroom", "Study / Home Office", "Dining Room",
    "Foyer / Entrance", "Balcony", "Unknown",
]

# Excel column widths — never change these
BOQ_COL_WIDTHS    = {"A": 15.83, "B": 25.83, "C": 10.83, "D": 10.83, "E": 10.83, "F": 15.83}
VENDOR_COL_WIDTHS = {"A": 40, "B": 44, "C": 54, "D": 36, "E": 22, "F": 15, "G": 19}
```

---

## recalc.py

```python
import sys, json, openpyxl

def verify_boq(path, max_rows=500):
    wb = openpyxl.load_workbook(path)   # NOT data_only=True — that strips formulas
    ws = wb.active
    errors = []
    for r in range(2, min(ws.max_row + 1, max_rows + 2)):
        qty = ws.cell(r, 4).value
        if qty is None:
            continue
        expected = f"=D{r}*E{r}"
        found = ws.cell(r, 6).value
        if found != expected:
            errors.append({"row": r, "expected": expected, "found": str(found)})
    result = {"total_errors": len(errors), "errors": errors}
    print(json.dumps(result, indent=2))
    return result

if __name__ == "__main__":
    r = verify_boq(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 500)
    sys.exit(0 if r["total_errors"] == 0 else 1)
```

Run: `python3 recalc.py path/to/boq.xlsx 500` → must return `"total_errors": 0`

---

## modules/room_analyzer.py

```python
import os, json, base64
from dataclasses import dataclass
import anthropic
from dotenv import load_dotenv

load_dotenv()
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

VISION_PROMPT = """Analyse this interior design render carefully. Return a JSON object with:
- room_type: one of [Living Room, Master Bedroom, Bedroom, Kitchen, Bathroom, Study / Home Office, Dining Room, Foyer / Entrance, Balcony, Unknown]
- estimated_sqft: integer estimate of floor area in sqft (bedroom 120-200, living 200-400, kitchen 80-150, bathroom 40-80)
- confidence: "high" | "medium" | "low"
- design_elements: one detailed sentence listing EVERY element visible — floor material, ceiling treatment, every carpentry item (wardrobe/TV unit/headboard/desk/shelves), wall treatment, every light fixture, AC, fans, furniture, soft furnishings, decor. Include material specs and brands where visible (e.g. "cream marble large-format floor, cove-lit tray ceiling with recessed COB downlights, 3-door full-height wardrobe with cream laminate and Hettich soft-close, Atomberg BLDC fan, split AC on feature wall").
Return ONLY valid JSON, no other text."""

@dataclass
class RoomAnalysis:
    room_type: str
    estimated_sqft: int
    confidence: str
    design_elements: str
    image_filename: str = ""

def analyze_render(image_bytes: bytes, media_type: str, filename: str = "") -> RoomAnalysis:
    b64 = base64.standard_b64encode(image_bytes).decode()
    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=512,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                {"type": "text", "text": VISION_PROMPT},
            ]}],
        )
        text = resp.content[0].text.strip()
        if text.startswith("```"):
            text = text.split("```")[1].lstrip("json").strip()
        data = json.loads(text)
        return RoomAnalysis(
            room_type=data.get("room_type", "Unknown"),
            estimated_sqft=int(data.get("estimated_sqft", 120)),
            confidence=data.get("confidence", "medium"),
            design_elements=data.get("design_elements", ""),
            image_filename=filename,
        )
    except Exception:
        return RoomAnalysis("Unknown", 120, "low",
            "Could not analyse image — describe the room manually.", filename)

def analyze_all_renders(images: list[tuple[bytes, str, str]]) -> list[RoomAnalysis]:
    return [analyze_render(b, mt, fn) for b, mt, fn in images]
```

---

## modules/boq_generator.py

### System Prompt (the complete rate library — never truncate)

```python
BOQ_SYSTEM_PROMPT = """You are the Houspire BOQ Generator. Produce client-deliverable Excel BOQs for Indian residential interior projects.

TEMPLATE RULES (non-negotiable):
- Two tiers only: Mid-tier or Premium. Never economy/standard/luxury.
- City multipliers are SILENT — bake into each line-item rate, never show as a row.
- Amount = Quantity × Rate (the Excel formula handles it — do not compute).
- Description must end with " - <Room Name>".
- No footer rows — no subtotal, GST, contingency, grand total.
- TV/electronics and kitchen appliances are client scope — exclude unless asked.

GRANULARITY — always produce sample-level detail (every line item matches a vendor site-quote):
- ELECTRICAL: Individual points — 6A switch (Legrand Arteor, count per room); 2-way/master switch; 16A socket; dedicated 16A AC point + isolator; USB-C outlet; premium dimmer; ceiling fan drop + slab reinforcement; TV-wall MS reinforcement; concealed conduit + FR-LSH wiring lump (Polycab/Havells, scaled 4500-11000 by room size).
- LIGHTING: Cove LED strip (SMD 14W/m, high-CRI 3000K) by rft; COB downlights 5W trimless by nos; BLDC fan (Atomberg/Anemos) by nos; internal shelf LED strip + driver as lump.
- HARDWARE: Hettich Sensys 110° soft-close hinges (pair); Hettich Quadro full-extension drawer slides (pair); Hafele Magic Corner/Le Mans (nos); profile handles (nos); MS bracket + wall anchor lump per floating piece.
- AC: Unit on one line; install kit on the NEXT separate line. Always two rows.
- CARPENTRY: Each piece separately — wardrobe, loft, desk, shelves, bedside, TV unit. Never bundle.
- Brands required: Century BWR ply, Greenlam laminate, Kajaria/Somany tiles, Mikasa engineered wood, Legrand Arteor switches, Polycab/Havells FR-LSH cable, Atomberg fans, Asian Paints Royale Luxury Emulsion, Marshalls/Excel wallpaper, Hettich/Hafele hardware.

RATE LIBRARY — Hyderabad baseline ×1.00. Apply city multiplier silently:

CARPENTRY / CEILING:
Gypsum FC cove+paint: 165/sft | Gypsum FC magnetic-track: 168/sft | Coffered FC: 180/sft
Wood-veneer slat ceiling: 920/sft | Fluted wall paneling (Century veneer+PU): 1400/sft
Bed-back panel (mixed): 1500-1700/sft | Built-in TV unit: 1800-2000/sft
Wardrobe premium: 2200/sft | Wardrobe loft: 1750/sft | Display cabinet reeded glass+LED: 2100/sft
Study desk + drawer + cabinet: 1334/sft | Open display shelves with LED: 5980/nos
Modular kitchen base: 2100/sft | wall: 1900/sft | tall: 2200/sft | island: 2000/sft
Quartz countertop 20mm: 800/sft | Crockery unit: 2050/sft

FLOORING / WALLS:
Premium vitrified (Kajaria/Somany large-format): 179/sft | Engineered wood (Mikasa): 430/sft
Wall emulsion Royale Luxury (putty+primer+2 coats): 38/sft
Wallpaper supply+install (Marshalls/Excel): 140-152/sft | Classical wall mouldings: 140/sft
Bath wall tile (porcelain): 260/sft | Bath floor tile (anti-skid): 230/sft
TV-wall marble cladding (Statuario 15mm): 550/sft

LIGHTING:
COB downlights 5W trimless: 650/nos | LED cove strip high-CRI 14W/m 3000K: 110/rft
Magnetic track rail+driver: 950/rft | Magnetic track spot COB 12W: 2200/nos
BLDC fan with light kit (Atomberg/Anemos): 14000/nos
Bedside lamp brass+linen: 4140/nos | Desk task lamp: 3220/nos | Shelf LED strip+driver: 1656/lump

ELECTRICAL (Legrand Arteor throughout):
6A switch: 1150/nos | 2-way/master: 1300/nos | 16A socket: 1450/nos
Dedicated 16A AC+isolator: 1500/nos | Dedicated 25A appliance: 2200/nos
AV/Cat6+coax: 1250/nos | USB-C: 1900/nos | Premium dimmer: 2600/nos
Ceiling fan drop+slab: 1200/nos | TV-wall MS reinforcement: 2200/lump
FR-LSH wiring lump (Polycab/Havells): 4500-11000/lump

HVAC:
Split AC 1.5T 5-star unit (Daikin/Mitsubishi Heavy): 51000/nos
Split AC 1.5T install kit: 11000/lump
Split AC 2.0T 5-star unit: 62000/nos | Split AC 2.0T install kit: 12000/lump

HARDWARE:
Hettich Sensys 110° hinge: 350/pair | Hettich Quadro drawer slide: 1400/pair
Hettich Senso lift-up: 1200/nos | Hafele Magic Corner/Le Mans: 14000/nos
Pull-out wire basket: 4500/nos | Profile handle: 380-580/nos
MS bracket+anchor lump: 3500-5500/lump

TEXTILES / DECOR:
Curtains sheer+drape+track: 650-750/rft | Area rug premium: 7820/nos
Framed art set of 2 + accessories: 5980/lump

BATHROOM:
Waterproof FC: 145/sft | Waterproofing crystalline+polymer: 95/sft
Wall-hung WC+Geberit cistern+flush plate: 28000-36000/nos
Rain shower head+arm: 9500-12000/nos | Hand shower+hose+bracket: 4500-5500/nos
Concealed 3-way diverter: 9500-13000/nos | Basin mixer: 7500-10000/nos | Health faucet: 3500/nos
Fluted vanity+stone counter: 14000-22000/nos | Backlit LED mirror anti-fog: 8000-9500/nos
Concealed plumbing CPVC+labour per bath: 18000-24000/lump"""
```

### User Prompt Builder

```python
def _build_user_prompt(rooms, city, pincode, tier):
    room_summary = "\n".join(
        f"- {r.room_type} (~{r.estimated_sqft} sft): {r.design_elements}" for r in rooms
    )
    return f"""Generate a DETAILED, sample-level BOQ. Every line item = vendor site-quote granularity.

PROJECT: City: {city} | Pincode: {pincode} | Tier: {tier}

ROOMS (with design elements from renders):
{room_summary}

INSTRUCTIONS:
1. Apply the {city} city multiplier silently to all rates.
2. Generate a line item for EVERY visible and implied element. Never bundle:
   - Carpentry: each piece separately (wardrobe, loft, desk, shelves, bedside, TV unit)
   - Electrical: each point type separately (switch, socket, AC point, fan drop, wiring lump)
   - Lighting: cove strip by rft, downlights by nos, fan by nos
   - AC: always two rows — unit + install kit
   - Hardware: hinges by pair, slides by pair, handles by nos
3. Brand names REQUIRED in every description.
4. Description format: "[full spec + brand] - [Room Name]"
5. Rate Sources: cite real brand website for every rate (centuryply.com, hettich.com, kajariaceramics.com, atomberg.com, asianpaints.com, daikinindia.com, legrand.co.in, etc.)

OUTPUT — exactly these two sections, comma-separated:

=== BOQ TABLE ===
Category,Description,Unit,Quantity,Rate,Amount
[one line per item — Amount = Quantity*Rate as integer]

=== RATE SOURCES ===
Category / Item,Rate Basis ({city} {tier} 2025-26),Source
[real URL for every rate]
Note,{city} multiplier applied silently to all base rates.,-
Note,Room sizes estimated from renders — confirm before ordering.,-
Note,TV/electronics and kitchen appliances are client scope — excluded.,-"""
```

### API Call

```python
def generate_boq(rooms, city, pincode, tier):
    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=4096,
        system=BOQ_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": _build_user_prompt(rooms, city, pincode, tier)}],
    )
    return _parse_boq_response(resp.content[0].text)
```

### Parser

```python
import re
from dataclasses import dataclass

@dataclass
class BOQRow:
    category: str; description: str; unit: str; qty: float; rate: float

@dataclass
class RateSource:
    item: str; basis: str; source: str

def _parse_boq_response(text):
    rows, sources = [], []
    boq_m = re.search(r"=== BOQ TABLE ===\n(.*?)\n(?:=== RATE SOURCES ===|$)", text, re.DOTALL)
    src_m = re.search(r"=== RATE SOURCES ===\n(.*?)$", text, re.DOTALL)
    if boq_m:
        for line in boq_m.group(1).strip().splitlines():
            p = [x.strip() for x in line.split(",")]
            if len(p) < 5 or p[0] == "Category": continue
            try:
                rows.append(BOQRow(p[0], p[1], p[2], float(p[3]), float(p[4])))
            except (ValueError, IndexError):
                continue
    if src_m:
        for line in src_m.group(1).strip().splitlines():
            p = [x.strip() for x in line.split(",", 2)]
            if len(p) < 2 or not p[0]: continue
            sources.append(RateSource(p[0], p[1] if len(p)>1 else "", p[2] if len(p)>2 else ""))
    return rows, sources
```

---

## modules/vendor_finder.py

### System Prompt

```python
VENDOR_SYSTEM_PROMPT = """You are the Houspire Vendor Research specialist. Find real, Google-verified local vendors for interior projects.

NON-NEGOTIABLE RULES:
- NEVER fabricate vendor names, phones, ratings, review counts, or addresses.
- Only include vendors verified via live web/Google Maps search.
- No published phone → write: NA - visit showroom
- NEVER include competitors: Livspace, Artifex, Fabulous Decor, Hipcouch, Urban Company.
- Sort nearest-first within each category.
- Read reviews — flag every verified negative with specific detail.
- Cannot find vendors for a category → say so. Do not invent.

RADIUS RULES:
- PRIMARY: 10 km from pincode centroid — find 3-5 vendors here first.
- FALLBACK: If < 3 within 10 km, extend to 15 km. Label those "far-belt fallback (~X km)".
- HARD LIMIT: Never include vendors beyond 20 km unless zero options exist closer.
- VENDORS PER CATEGORY: exactly 3-5. Never fewer than 3 (extend radius). Never more than 5.

DATA FORMAT:
- Rating: "4.8 (132)" — exact Google count in brackets. Never approximate.
- Phone: "+91 XXXXX XXXXX". If missing from Google → NA - visit showroom.
- Area: "Sub-locality, Locality (PINCODE)"
- Distance: decimal to 1 place e.g. 9.7

OUTPUT — exactly two sections, pipe-separated:

=== VENDOR TABLE ===
Category|Vendor|Specialty / Brands|Area|lat|lng|Rating (count)|Phone

=== NOTES ===
[structured notes — see format below]

Notes structure (mandatory):
1. Pincode reference: sub-areas covered, centroid lat/lng
2. Vendor cluster strategy: belt structure with distances and trades per belt
3. City multiplier: locked (state value) or not locked (suggest band)
4. Trades covered: one bullet per category matched to render elements
5. Vendor highlights: spec-critical sole-sources + strongest picks with rationale
6. Open items/flags: every verified complaint — specific, not softened
7. Exclusions: competitors dropped, low-review vendors dropped, why"""
```

### User Prompt Builder

```python
def _build_vendor_prompt(rooms, city, pincode, tier, categories):
    room_desc = "\n".join(f"- {r.room_type}: {r.design_elements}" for r in rooms)
    cats = ", ".join(categories) if categories else "auto-detect from renders"
    return f"""Find real, Google-verified local vendors for this Houspire project.

PROJECT: City: {city} | Pincode: {pincode} | Tier: {tier}

ROOM DESCRIPTIONS:
{room_desc}

CATEGORIES NEEDED: {cats}

STEP 1 — Centroid
Search: "{pincode} pincode {city} area locality"
Find sub-areas covered and centroid (lat, lng). State it before proceeding.

STEP 2 — Search per category
For EACH category search Google Maps:
  "[trade] [nearest locality] {city}"
  "[brand] dealer [locality] {city}"
Primary zone: 10 km. Extend to 15 km only if < 3 results within 10 km.

STEP 3 — Select 3-5 vendors, nearest-first
Verify each on Google Maps. Record: name, specialty/brands, full locality+pincode, lat, lng, rating+count, phone.

STEP 4 — Review check
Check Google reviews for: warranty disputes | rude behavior | delivery failures | GST fraud | fake products.
If found → include vendor BUT add one-sentence specific flag after " - " in Specialty column.
Example: "Wardrobe + carpentry - flagged: 3 verified complaints about non-delivery after full advance, lock payment milestones in writing"

STEP 5 — Distance
dist_km = sqrt((Δlat×111)² + (Δlng×111×cos(centroid_lat_rad))²), rounded to 1 decimal.

Return VENDOR TABLE and NOTES as specified."""
```

### API Call (with web_search tool)

```python
def generate_vendors(rooms, city, pincode, tier, categories=None):
    if categories is None:
        categories = []
    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8192,
        system=VENDOR_SYSTEM_PROMPT,
        tools=[{"type": "web_search", "name": "web_search"}],
        messages=[{"role": "user",
                   "content": _build_vendor_prompt(rooms, city, pincode, tier, categories)}],
    )
    # Extract final text block after tool use rounds
    final_text = next((b.text for b in resp.content if hasattr(b, "text")), "")
    return _parse_vendor_response(final_text)
```

---

## modules/budget_excel.py — Exact Plain Template

```python
import openpyxl
from openpyxl.styles import Font
from config import BOQ_COL_WIDTHS

BOLD  = Font(name="Calibri", size=12, bold=True)
PLAIN = Font(name="Calibri", size=12, bold=False)

BOQ_HEADERS = ["Category","Description","Unit","Quantity","Rate","Amount (Auto-calculated)"]
SRC_HEADERS = ["Category / Item","Rate Basis","Source"]

def write_budget_excel(client_name, city, pincode, tier, rows, sources, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Sheet 1: BOQ Template
    ws = wb.create_sheet("BOQ Template")
    for c, h in enumerate(BOQ_HEADERS, 1):
        ws.cell(1, c, h).font = BOLD
    for i, row in enumerate(rows, 2):
        ws.cell(i, 1, row.category).font = PLAIN
        ws.cell(i, 2, row.description).font = PLAIN
        ws.cell(i, 3, row.unit).font = PLAIN
        ws.cell(i, 4, row.qty).font = PLAIN
        ws.cell(i, 5, row.rate).font = PLAIN
        ws.cell(i, 6, f"=D{i}*E{i}").font = PLAIN   # ← formula, NEVER hard-coded
    for col, w in BOQ_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    # Sheet 2: Rate Sources
    ws2 = wb.create_sheet("Rate Sources")
    for c, h in enumerate(SRC_HEADERS, 1):
        ws2.cell(1, c, h).font = BOLD
    for i, src in enumerate(sources, 2):
        ws2.cell(i, 1, src.item).font = PLAIN
        ws2.cell(i, 2, src.basis).font = PLAIN
        ws2.cell(i, 3, src.source).font = PLAIN
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 45
    ws2.column_dimensions["C"].width = 45

    wb.save(output_path)
    return output_path
```

**File name:** `{ClientName}_BOQ_{City}.xlsx`

---

## modules/vendor_excel.py — Exact Plain Template (from build_template.py)

```python
import math, openpyxl
from openpyxl.styles import Font
from config import VENDOR_COL_WIDTHS

BOLD  = Font(name="Calibri", size=12, bold=True)
PLAIN = Font(name="Calibri", size=12, bold=False)

def _km(lat, lng, lat0, lng0):
    dlat = (lat - lat0) * 111.0
    dlng = (lng - lng0) * 111.0 * math.cos(math.radians(lat0))
    return round(math.hypot(dlat, dlng), 1)

def write_vendor_excel(project_name, city, pincode, centroid, vendors, notes_text, output_path):
    lat0, lng0 = centroid
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendors"

    headers = ["Category","Vendor","Specialty / Brands","Area",
               f"Approx. km from {pincode}","Rating (count)","Phone"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h).font = BOLD

    ws.freeze_panes = "A2"
    for col, w in VENDOR_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    r = 2
    for v in vendors:
        dist = _km(v.lat, v.lng, lat0, lng0) if v.lat and v.lng else ""
        for c, val in enumerate(
            [v.category, v.vendor, v.specialty, v.area, dist, v.rating, v.phone], 1
        ):
            ws.cell(r, c, val).font = PLAIN
        r += 1

    ns = wb.create_sheet("Notes")
    ns.column_dimensions["A"].width = 120
    for i, line in enumerate(notes_text.splitlines(), 1):
        ns.cell(i, 1, line)

    wb.save(output_path)
    return output_path
```

**File name:** `{City}_{RoomType}_Vendors_{PIN}.xlsx` (single room) or `{Project}_Vendors_{City}_{PIN}.xlsx` (full home)

---

## modules/docx_generator.py — Branded Output

**Brand system (never deviate):**
```
BRAND_PRIMARY = "1B4D3E"   # Deep Forest Green
BRAND_ACCENT  = "D4AF37"   # Gold
BRAND_LIGHT   = "F5F0E8"   # Cream
ROW_ALT       = "F7F5F0"   # Alternating rows
Font          = "Arial" throughout (not Calibri)
```

**CRITICAL:** Never mention city multipliers, rate adjustment factors, or internal pricing mechanics in branded client documents. Instead say: "All rates researched for [City] market (2025-26)".

**Workflow:**
1. Read the plain Excel with `openpyxl` (data_only=True)
2. Extract data into a Python dict
3. Serialize to JSON
4. Generate a self-contained Node.js script that embeds the JSON and uses `require("docx")`
5. Run the script with `subprocess.run(["node", script_path])`
6. Script writes the `.docx` to the output path

**Document structure (both BOQ and Vendor):**
```
1. Cover page — HOUSPIRE wordmark (size 88, characterSpacing 300), tagline, gold divider,
   document title, client/city subtitle, stats line, cream info box
2. Detail pages — one section per category/room, gold section number + green title,
   professional table (green header, alternating rows), section subtotal (BOQ)
3. Rate Sources table (BOQ only)
4. Disclaimer page — Houspire flat-fee model, no commissions, rates are indicative
5. Header: right-aligned italic grey "Houspire [Type] | [Client] | Confidential"
6. Footer: "www.houspire.ai | Page X"
```

**Page setup:** A4 (11906×16838), margins 1000 DXA (dense BOQ) or 1200 DXA (vendor directory)

**Node.js requires:**
```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageBreak, PageNumber
} = require("docx");
```

**File names:**
- `Houspire_BOQ_{ClientName}_{City}.docx`
- `Houspire_Vendors_{ClientName}_{City}.docx`

---

## modules/pdf_converter.py

```python
import os, subprocess

def docx_to_pdf(docx_path, output_dir=None):
    if output_dir is None:
        output_dir = os.path.dirname(docx_path)
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice failed:\n{result.stderr}")
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, stem + ".pdf")
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    return pdf_path
```

---

## app.py — Streamlit UI Flow

```
Page title: "🏠 Houspire Budget Generator"
Caption: "Internal tool — generates BOQ + Vendor Excel from client renders"

Step 1 — Client Details (4 columns)
  client_name | city (selectbox from CITIES) | pincode | tier (Mid-tier / Premium)
  Show city multiplier caption if known. Prompt for estimate if city = "Other".

Step 2 — Upload Room Renders
  st.file_uploader, multiple=True, jpg/jpeg/png

Step 3 — Floor Plan (Optional)
  st.file_uploader, single, jpg/jpeg/png

Step 4 — Analyse Rooms button
  Disabled until renders + client_name + pincode are filled.
  On click: analyze_all_renders() → store in session_state.analyses

Step 5 — Review Detected Rooms
  Expandable per render. Editable: room_type (selectbox), sqft (number_input), design_elements (text_area).
  Show confidence badge.

Step 6 — Output Format
  st.radio: "📊 Plain Excel (internal use)" | "🎨 Branded DOCX + PDF (client delivery)"

Step 7 — Generate
  Two buttons: "📊 Generate Budget BOQ" | "📍 Find Local Vendors"

  BOQ button flow:
    → generate_boq() → write_budget_excel() → [if branded] generate_branded_boq_docx() → docx_to_pdf()
    → store paths in session_state

  Vendor button flow:
    → generate_vendors() → write_vendor_excel() → [if branded] generate_branded_vendor_docx() → docx_to_pdf()
    → store paths in session_state

Downloads section:
  Plain: Budget Excel + Vendor Excel
  Branded adds: Budget DOCX + Vendor DOCX + Budget PDF + Vendor PDF
```

**Session state keys:**
```python
["analyses", "boq_rows", "boq_sources", "vendors", "notes", "centroid",
 "budget_path", "vendor_path",
 "budget_docx_path", "vendor_docx_path",
 "budget_pdf_path", "vendor_pdf_path"]
```

---

## Testing

Each module has a test file in `tests/`. Run all with:
```bash
python -m pytest tests/ -v
```

Key assertions:
- `test_budget_excel.py`: Amount column = `=D2*E2` formula (not a number); no footer rows; header row bold; data rows not bold; two sheets (BOQ Template + Rate Sources)
- `test_vendor_excel.py`: Distance column is float (not lat/lng string); header bold; data plain; Notes sheet present; row count = vendor count
- `test_docx_generator.py`: Output file exists; size > 5000 bytes (not empty)

After generating any BOQ Excel, always run:
```bash
python3 recalc.py path/to/boq.xlsx 500
```
Must return `"total_errors": 0` before the file is delivered.

---

## Replit Deployment

```
Run command: streamlit run app.py --server.port 8080 --server.address 0.0.0.0
Secret:      ANTHROPIC_API_KEY
```

Install LibreOffice on Replit if needed:
```bash
apt-get install -y libreoffice
```

Replit natively supports Python + Node.js — no extra configuration needed for `docx` npm.

---

## What Claude Must Never Do

- Invent vendor names, phone numbers, ratings, or review counts
- Mention city multipliers or rate adjustment factors in any client-facing document
- Hard-code the Amount column in Excel — always `=D{row}*E{row}`
- Add footer rows (subtotal/GST/contingency) to the plain Excel
- Use "economy", "standard", or "luxury" as tier names
- Bundle carpentry items (wardrobe + desk = one line) — each piece is its own row
- Bundle AC unit + install kit — always two separate rows
- Skip the Notes tab in vendor Excel
- Skip the Rate Sources sheet in budget Excel
- Include turnkey competitor firms in vendor lists (Livspace, Artifex, Fabulous Decor, Hipcouch, Urban Company)
- Use Calibri font in branded DOCX — it must be Arial throughout
- Skip the `recalc.py` verification before delivering a BOQ Excel

---

## Files Reference

| File | Location in this project |
|------|--------------------------|
| Implementation plan (full task breakdown with code) | `houspire-budget-generator-plan.md` |
| Sample budget Excel (exact template to match) | Uploaded: `Bedroom_BOQ_Jaipur.xlsx` |
| Sample vendor Excel (exact template to match) | Uploaded: `Jaipur_Bedroom_Vendors_303905.xlsx` |
| Vendor build template (Python + openpyxl) | Uploaded: `build_template.py` |
| Budget skill (full rate library + workflow) | Uploaded: `SKILL (2).md` |
| Vendor skill (full search workflow) | Uploaded: `SKILL (1).md` |
| Branding skill (DOCX brand system) | Uploaded: `houspire-branding-docx-SKILL.md` |
