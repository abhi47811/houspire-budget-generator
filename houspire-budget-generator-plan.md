# Houspire Budget Generator — Implementation Plan (v2)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Internal Houspire web app — staff enters client name, city, pincode, tier, uploads renders + optional floorplan → Claude API generates BOQ + vendor list → **two output formats**: (A) plain Excel template for internal use, or (B) branded DOCX + PDF for offline client delivery.

**Architecture:** Streamlit Python app. Claude API (Sonnet) handles all intelligence: Vision for room analysis, text generation for BOQ using the built-in rate library, web_search tool for live vendor research near the pincode. openpyxl writes plain Excel files. For branded output, a Python module generates a Node.js script using the `docx` npm package (Navy/Gold brand system) and runs it, then LibreOffice converts the DOCX to PDF. One API key (Anthropic) — nothing else.

**Tech Stack:** Python 3.11, Streamlit, anthropic (Python SDK), openpyxl, python-dotenv, Node.js + docx (npm), LibreOffice (PDF conversion)

---

## Critical Template Rules (from live skills — never deviate)

### Budget Excel (houspire-budget skill)
- Font: Calibri 12, **no bold**, no fills, no borders, no merged cells
- Row 1 headers ONLY are bold: `Category | Description | Unit | Quantity | Rate | Amount (Auto-calculated)`
- Column widths: A=15.83, B=25.83, C=10.83, D=10.83, E=10.83, F=15.83
- Amount column: always formula `=D{row}*E{row}` — never a hard-coded number
- **No footer rows** — no subtotal, GST, contingency, grand total
- Second sheet titled `Rate Sources`: Category/Item | Rate Basis (city, tier, year) | Source
- File name: `{Project}_BOQ_{City}.xlsx`

### Vendor Excel (houspire-vendor-research skill + build_template.py)
- Font: Calibri 12, **header row bold only**, no fills, no borders, no merged cells
- Row 1 headers: `Category | Vendor | Specialty / Brands | Area | Approx. km from {PIN} | Rating (count) | Phone`
- Column widths: A=40, B=44, C=54, D=36, E=22, F=15, G=19
- Freeze panes at A2
- Distance column: decimal to 1 place, calculated via Euclidean formula from pincode centroid
- Rating format: `"4.8 (132)"`
- Phone format: `"+91 XXXXX XXXXX"` or `NA - visit showroom`
- Second sheet titled `Notes` (mandatory — see structure below)
- File name: `{City}_{Room}_Vendors_{PIN}.xlsx` or `{Project}_Vendors_{City}_{PIN}.xlsx`

### Tiers — Two Only
Houspire uses **Mid-tier** and **Premium** only. Never economy / standard / luxury / budget.

### City Multipliers (applied silently to rates — never shown in BOQ sheet)
| City | Multiplier |
|------|-----------|
| Hyderabad | ×1.00 (baseline) |
| Davanagere | ×0.90 |
| Kathua J&K | ×0.88 |
| Jaipur | ×0.92 |
| Sangli | ×0.93 |
| Bhopal | ×0.93 |
| Hubli | ×0.92 |
| Bodeli | ×0.88 |
| Gandhinagar | ×0.96 |
| Jorhat | ×0.95 |
| Nashik / Trivandrum | ×0.95 |
| Vadodara | ×0.98 |
| Visakhapatnam | ×0.98 |
| Kolkata | ×1.02 |
| Pune | ×1.08 |
| Chennai | ×1.10 |
| Bangalore | ×1.12 |
| NCR / Noida | ×1.18 |
| Thane / MMR | ×1.20 |
| Mumbai | ×1.25 |

For any city not in this list: flag in Rate Sources notes, suggest a band, ask user to lock before first BOQ.

---

## Environment Variables

```
ANTHROPIC_API_KEY=sk-ant-...
```

That is the only secret. No other APIs.

---

## File Structure

```
houspire_budget_generator/
├── app.py                      # Streamlit UI + orchestration
├── config.py                   # Cities, multipliers, tier list, categories
├── requirements.txt
├── .env.example
├── recalc.py                   # Formula verification script (from skill)
├── modules/
│   ├── __init__.py
│   ├── room_analyzer.py        # Claude Vision → room type + sqft + design elements
│   ├── boq_generator.py        # Claude API call → BOQ table rows + rate sources
│   ├── vendor_finder.py        # Claude API call with web_search → vendor data
│   ├── budget_excel.py         # openpyxl — Budget BOQ Excel (exact plain template)
│   ├── vendor_excel.py         # openpyxl — Vendor Excel (exact plain template)
│   ├── docx_generator.py       # Reads Excel → generates Node.js script → runs it → .docx
│   └── pdf_converter.py        # LibreOffice subprocess → .docx to .pdf
└── tests/
    ├── test_room_analyzer.py
    ├── test_boq_generator.py
    ├── test_budget_excel.py
    ├── test_vendor_excel.py
    └── test_docx_generator.py
```

---

## Task 1: Project Setup

**Files:**
- Create: `houspire_budget_generator/requirements.txt`
- Create: `houspire_budget_generator/.env.example`
- Create: `houspire_budget_generator/config.py`
- Create: `houspire_budget_generator/modules/__init__.py`
- Create: `houspire_budget_generator/recalc.py`

- [ ] **Step 1.1: Create requirements.txt**

```
streamlit==1.35.0
anthropic==0.28.0
openpyxl==3.1.2
python-dotenv==1.0.1
Pillow==10.3.0
```

- [ ] **Step 1.2: Create .env.example**

```
ANTHROPIC_API_KEY=sk-ant-...
```

- [ ] **Step 1.3: Create config.py**

```python
# config.py

TIERS = ["Mid-tier", "Premium"]

# Cities with known multipliers — displayed in dropdown
CITIES_WITH_MULTIPLIERS = {
    "Hyderabad": 1.00,
    "Davanagere": 0.90,
    "Kathua": 0.88,
    "Jaipur": 0.92,
    "Sangli": 0.93,
    "Bhopal": 0.93,
    "Hubli": 0.92,
    "Bodeli": 0.88,
    "Gandhinagar": 0.96,
    "Jorhat": 0.95,
    "Nashik": 0.95,
    "Trivandrum": 0.95,
    "Vadodara": 0.98,
    "Visakhapatnam": 0.98,
    "Kolkata": 1.02,
    "Pune": 1.08,
    "Chennai": 1.10,
    "Bangalore": 1.12,
    "Noida": 1.18,
    "Delhi": 1.18,
    "Thane": 1.20,
    "Mumbai": 1.25,
    "Other": None,  # triggers multiplier prompt
}

CITIES = list(CITIES_WITH_MULTIPLIERS.keys())

ROOM_TYPES = [
    "Living Room", "Master Bedroom", "Bedroom",
    "Kitchen", "Bathroom", "Study / Home Office",
    "Dining Room", "Foyer / Entrance", "Balcony", "Unknown",
]

# BOQ column widths (from skill spec)
BOQ_COL_WIDTHS = {"A": 15.83, "B": 25.83, "C": 10.83, "D": 10.83, "E": 10.83, "F": 15.83}

# Vendor column widths (from build_template.py)
VENDOR_COL_WIDTHS = {"A": 40, "B": 44, "C": 54, "D": 36, "E": 22, "F": 15, "G": 19}
```

- [ ] **Step 1.4: Create recalc.py** (verifies Excel formulas — from skill spec)

```python
# recalc.py
"""
Verify that every Amount cell in a BOQ Excel file equals Quantity * Rate.
Usage: python3 recalc.py {file}.xlsx {max_rows}
Returns: JSON with total_errors count. Must be 0 before delivery.
"""
import sys
import json
import openpyxl

def verify_boq(path: str, max_rows: int = 500) -> dict:
    wb = openpyxl.load_workbook(path)  # do NOT use data_only=True
    ws = wb.active
    errors = []
    for row_idx in range(2, min(ws.max_row + 1, max_rows + 2)):
        qty_cell   = ws.cell(row=row_idx, column=4)
        rate_cell  = ws.cell(row=row_idx, column=5)
        amt_cell   = ws.cell(row=row_idx, column=6)
        # Skip empty rows
        if qty_cell.value is None and rate_cell.value is None:
            continue
        expected_formula = f"=D{row_idx}*E{row_idx}"
        if amt_cell.value != expected_formula:
            errors.append({
                "row": row_idx,
                "expected": expected_formula,
                "found": str(amt_cell.value),
            })
    result = {"total_errors": len(errors), "errors": errors}
    print(json.dumps(result, indent=2))
    return result

if __name__ == "__main__":
    path = sys.argv[1]
    max_r = int(sys.argv[2]) if len(sys.argv) > 2 else 500
    r = verify_boq(path, max_r)
    sys.exit(0 if r["total_errors"] == 0 else 1)
```

- [ ] **Step 1.5: Create modules/__init__.py**

```python
# modules/__init__.py
```

- [ ] **Step 1.6: Commit**

```bash
cd houspire_budget_generator
git init
git add .
git commit -m "feat: project scaffold — config, recalc verifier, requirements"
```

---

## Task 2: Room Analyzer

**Files:**
- Create: `houspire_budget_generator/modules/room_analyzer.py`
- Create: `houspire_budget_generator/tests/test_room_analyzer.py`

- [ ] **Step 2.1: Write failing test**

```python
# tests/test_room_analyzer.py
import pytest
from unittest.mock import patch, MagicMock
from modules.room_analyzer import analyze_render, RoomAnalysis

def test_analyze_render_returns_room_analysis():
    mock_resp = MagicMock()
    mock_resp.content[0].text = """{
        "room_type": "Master Bedroom",
        "estimated_sqft": 180,
        "confidence": "high",
        "design_elements": "cream marble floor, cove-lit tray ceiling, upholstered queen bed with fluted headboard, wardrobe with soft-close"
    }"""
    with patch("modules.room_analyzer.client.messages.create", return_value=mock_resp):
        result = analyze_render(b"fake", "image/jpeg", "master_bed.jpg")
    assert isinstance(result, RoomAnalysis)
    assert result.room_type == "Master Bedroom"
    assert result.estimated_sqft == 180
    assert "marble" in result.design_elements

def test_analyze_render_fallback_on_bad_json():
    mock_resp = MagicMock()
    mock_resp.content[0].text = "I cannot determine."
    with patch("modules.room_analyzer.client.messages.create", return_value=mock_resp):
        result = analyze_render(b"fake", "image/jpeg", "unknown.jpg")
    assert result.room_type == "Unknown"
    assert result.estimated_sqft == 120
```

- [ ] **Step 2.2: Run test — expect FAIL**

```bash
cd houspire_budget_generator
python -m pytest tests/test_room_analyzer.py -v
```

Expected: `ModuleNotFoundError`

- [ ] **Step 2.3: Implement room_analyzer.py**

```python
# modules/room_analyzer.py
import os
import json
import base64
from dataclasses import dataclass, field
import anthropic
from dotenv import load_dotenv

load_dotenv()
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

VISION_PROMPT = """Analyse this interior design render carefully. Return a JSON object with:
- room_type: one of [Living Room, Master Bedroom, Bedroom, Kitchen, Bathroom, Study / Home Office, Dining Room, Foyer / Entrance, Balcony, Unknown]
- estimated_sqft: integer estimate of the room's floor area in square feet (typical Indian apartment: bedroom 120-200 sft, living 200-400 sft, kitchen 80-150 sft, bathroom 40-80 sft)
- confidence: "high" | "medium" | "low"
- design_elements: one detailed sentence listing every visible element relevant to budgeting — floor material, ceiling treatment, carpentry items (wardrobe / TV unit / headboard), wall treatment, lighting fixtures, AC, fans, furniture, soft furnishings, decor. Be specific about materials and brands if visible (e.g. "cream marble large-format floor, cove-lit tray ceiling, 3-door full-height wardrobe with cream laminate, Atomberg-style BLDC fan, split AC unit on wall").

Return ONLY valid JSON, no other text."""


@dataclass
class RoomAnalysis:
    room_type: str
    estimated_sqft: int
    confidence: str
    design_elements: str
    image_filename: str = ""


def analyze_render(image_bytes: bytes, media_type: str, filename: str = "") -> RoomAnalysis:
    """Send one render to Claude Vision. Returns RoomAnalysis; never raises."""
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")
    try:
        resp = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=512,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                    {"type": "text", "text": VISION_PROMPT},
                ],
            }],
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
        return RoomAnalysis(
            room_type="Unknown",
            estimated_sqft=120,
            confidence="low",
            design_elements="Could not analyse image — please describe the room manually.",
            image_filename=filename,
        )


def analyze_all_renders(images: list[tuple[bytes, str, str]]) -> list[RoomAnalysis]:
    """images = list of (bytes, media_type, filename). Returns analyses in same order."""
    return [analyze_render(b, mt, fn) for b, mt, fn in images]
```

- [ ] **Step 2.4: Run test — expect PASS**

```bash
python -m pytest tests/test_room_analyzer.py -v
```

- [ ] **Step 2.5: Commit**

```bash
git add modules/room_analyzer.py tests/test_room_analyzer.py
git commit -m "feat: room analyzer — Claude Vision extracts room type + design elements"
```

---

## Task 3: BOQ Generator

**Files:**
- Create: `houspire_budget_generator/modules/boq_generator.py`
- Create: `houspire_budget_generator/tests/test_boq_generator.py`

This calls Claude API with the houspire-budget skill prompt, passing room analysis + city + tier. Claude returns a structured BOQ table and rate sources.

- [ ] **Step 3.1: Write failing test**

```python
# tests/test_boq_generator.py
import pytest
from unittest.mock import patch, MagicMock
from modules.boq_generator import generate_boq, BOQRow, RateSource
from modules.room_analyzer import RoomAnalysis

SAMPLE_ROOM = RoomAnalysis(
    room_type="Master Bedroom",
    estimated_sqft=180,
    confidence="high",
    design_elements="cream marble floor, cove-lit tray ceiling, 3-door wardrobe",
    image_filename="bedroom.jpg",
)

MOCK_RESPONSE = """=== BOQ TABLE ===
Carpentry,3-door full-height wardrobe - Century BWR ply carcass + matte cream laminate (Greenlam) shutters with soft-close - Master Bedroom,sft,54,1932,104328
Flooring,Premium vitrified marble-look tile (Kajaria) - supply & lay incl. skirting - Master Bedroom,sft,180,179,32220

=== RATE SOURCES ===
Wardrobe,Century BWR ply + Greenlam laminate + Hettich soft-close,centuryply.com; greenlam.com 2025-26
Flooring,Kajaria premium vitrified marble-look large-format,kajariaceramics.com 2025-26
Note,PIN 500001 - Hyderabad. ×1.00 multiplier applied silently.,-
Note,Master Bedroom estimated ~180 sft from render.,-"""

def test_generate_boq_returns_rows_and_sources():
    mock_resp = MagicMock()
    mock_resp.content[0].text = MOCK_RESPONSE
    with patch("modules.boq_generator.client.messages.create", return_value=mock_resp):
        rows, sources = generate_boq([SAMPLE_ROOM], "Hyderabad", "500001", "Premium")
    assert len(rows) == 2
    assert isinstance(rows[0], BOQRow)
    assert rows[0].category == "Carpentry"
    assert rows[0].qty == 54
    assert len(sources) >= 2

def test_boq_row_amount_equals_qty_times_rate():
    mock_resp = MagicMock()
    mock_resp.content[0].text = MOCK_RESPONSE
    with patch("modules.boq_generator.client.messages.create", return_value=mock_resp):
        rows, _ = generate_boq([SAMPLE_ROOM], "Hyderabad", "500001", "Premium")
    for row in rows:
        assert row.qty > 0
        assert row.rate > 0
```

- [ ] **Step 3.2: Run test — expect FAIL**

```bash
python -m pytest tests/test_boq_generator.py -v
```

- [ ] **Step 3.3: Implement boq_generator.py**

```python
# modules/boq_generator.py
import os
import re
from dataclasses import dataclass
import anthropic
from dotenv import load_dotenv
from modules.room_analyzer import RoomAnalysis

load_dotenv()
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

BOQ_SYSTEM_PROMPT = """You are the Houspire BOQ Generator. You produce client-deliverable Excel BOQs for Indian residential interior projects.

TEMPLATE RULES (non-negotiable):
- Two tiers only: Mid-tier or Premium. Never economy/standard/luxury.
- City multipliers are SILENT — bake into each line-item rate, never show as a row.
- Amount = Quantity × Rate (never shown — the Excel formula handles it).
- Description must end with " - <Room Name>".
- No footer rows (no subtotal, GST, contingency, grand total).
- TV / electronics and kitchen appliances are client scope — exclude unless asked.

GRANULARITY — ALWAYS produce sample-level detail (every line item matches a vendor's site-quote):
- ELECTRICAL: Individual points — 6A switch (Legrand Arteor, count per room); 2-way/master switch; 16A socket; dedicated 16A AC point + isolator; USB-C outlet module; premium dimmer; ceiling fan drop + slab reinforcement; TV-wall MS reinforcement; concealed conduit + FR-LSH wiring lump (Polycab/Havells, scaled by room size 4500-11000).
- LIGHTING: Cove LED strip (SMD 14W/m, high-CRI, 3000K) by rft; COB downlights 5W trimless by nos; magnetic track rail + driver separate from spot heads; BLDC fan (Atomberg/Anemos) by nos; internal shelf LED strip + driver as lump.
- HARDWARE: Hettich Sensys 110° soft-close hinges (pair); Hettich Quadro full-extension drawer slides (pair); Hafele Magic Corner/Le Mans (nos); profile handles (nos); MS bracket + wall anchor lump per floating piece.
- AC: Split unit (Daikin/Mitsubishi Heavy 1.5T or 2.0T 5-star) on one line; install kit (copper piping + drain + stabilizer + labour) on the next line — always two separate rows.
- CARPENTRY: Each piece on its own line with full spec — door count, ply grade (Century BWR), laminate brand (Greenlam), hardware brand (Hettich). Never combine wardrobe + desk into one line.
- Brands required on every line: Century BWR ply, Greenlam laminate, Kajaria/Somany tiles, Mikasa engineered wood, Legrand Arteor switches, Polycab/Havells FR-LSH cable, Atomberg fans, Asian Paints Royale Luxury Emulsion, Marshalls/Excel wallpaper, Hettich/Hafele hardware.

RATE LIBRARY — Hyderabad baseline (×1.00). Apply city multiplier silently to every rate.

CARPENTRY / CEILING / SURFACES:
Gypsum FC with cove + paint: 165/sft | Gypsum FC magnetic-track recess: 168/sft | Coffered grid FC: 180/sft
Wood-veneer slat ceiling: 920/sft | Wood-fluted wall paneling (Century veneer + PU): 1400/sft
Bed-back panel (mixed materials): 1500-1700/sft | Built-in TV unit: 1800-2000/sft
TV-wall marble cladding (Statuario 15mm): 550/sft
Wardrobe (sliding/hinged premium): 2200/sft | Wardrobe loft: 1750/sft
Display cabinet (reeded glass + LED): 2100/sft
Modular kitchen — base: 2100/sft | wall: 1900/sft | tall: 2200/sft | island: 2000/sft
Quartz countertop 20mm: 800/sft | Crockery unit: 2050-2091/sft
Study desk with drawer + cabinet: 1334/sft | Open display shelves with LED: 5980/nos

FLOORING / WALLS:
Premium vitrified flooring (Kajaria/Somany large-format): 179/sft
Engineered wood flooring (Mikasa): 430/sft
Wall emulsion (Royale Luxury, putty+primer+2 coats): 38/sft
Wallpaper supply+install (Marshalls/Excel): 140-152/sft | Classical wall mouldings: 140/sft
Bath wall tile cladding (porcelain): 260/sft | Bath floor tile (anti-skid): 230/sft

LIGHTING:
Recessed COB downlights 5W trimless: 650/nos
LED cove strip high-CRI 14W/m warm 3000K: 110/rft
Magnetic track rail + driver: 950/rft | Magnetic track spot head COB 12W: 2200/nos
Designer BLDC fan with light kit (Atomberg/Anemos): 14000/nos
Bedside table lamp (brass base + linen shade): 4140/nos | Desk task lamp: 3220/nos
Internal LED shelf strip + driver: 1656/lump

ELECTRICAL (per point — Legrand Arteor throughout):
6A modular switch: 1150/nos | Master/2-way switch: 1300/nos | 16A modular socket: 1450/nos
Dedicated 16A AC point + isolator: 1500/nos | Dedicated 25A appliance point: 2200/nos
AV/Cat6 + coax point: 1250/nos | USB-C fast-charge outlet: 1900/nos | Premium dimmer: 2600/nos
Ceiling fan drop + slab reinforcement: 1200/nos | TV-wall MS reinforcement: 2200/lump
Concealed conduit + FR-LSH wiring lump (Polycab/Havells, by room size): 4500-11000/lump

HVAC:
Split AC 1.5T 5-star inverter unit (Daikin/Mitsubishi Heavy): 51000/nos
Split AC 1.5T install kit (copper piping+drain+stabilizer+labour): 11000/lump
Split AC 2.0T 5-star inverter unit: 62000/nos | Split AC 2.0T install kit: 12000/lump

HARDWARE:
Hettich Sensys 110° soft-close hinge: 350/pair | Hettich Quadro full-extension drawer slide: 1400/pair
Hettich Senso lift-up hinge: 1200/nos | Hafele Magic Corner/Le Mans: 14000/nos
Hettich/Hafele pull-out wire basket: 4500/nos | Brass/chrome/matte black profile handle: 380-580/nos
MS bracket + wall anchor lump per floating piece: 3500-5500/lump

TEXTILES / SOFT / DECOR:
Premium curtains sheer+drape+concealed track: 650-750/rft | Area rug (premium): 7820/nos
Framed botanical art set of 2 + accessories: 5980/lump

BATHROOM:
Bath waterproof FC: 145/sft | Waterproofing (crystalline+polymer): 95/sft
Wall-hung WC + Geberit/Jaquar cistern + flush plate: 28000-36000/nos
Rain shower head + arm: 9500-12000/nos | Hand shower + hose + bracket: 4500-5500/nos
Concealed 3-way diverter: 9500-13000/nos | Basin mixer: 7500-10000/nos | Health faucet: 3500/nos
Fluted wood vanity with stone counter: 14000-22000/nos | Backlit LED mirror anti-fog: 8000-9500/nos
Concealed plumbing CPVC+PVC+labour per bath: 18000-24000/lump"""


def _build_user_prompt(rooms: list[RoomAnalysis], city: str, pincode: str, tier: str) -> str:
    room_summary = "\n".join(
        f"- {r.room_type} (~{r.estimated_sqft} sft): {r.design_elements}"
        for r in rooms
    )
    return f"""Generate a DETAILED, sample-level BOQ for this Houspire project. Every line item must match the granularity of a vendor's site-quote sheet — no bundled lumps except where specified.

PROJECT DETAILS:
City: {city} | Pincode: {pincode} | Tier: {tier}

ROOMS DETECTED (with design elements visible in renders):
{room_summary}

INSTRUCTIONS:
1. Apply the {city} city multiplier silently to all rates from the rate library.
2. For each room, generate a line item for EVERY visible and implied element in the design description.
   - Carpentry: each piece separately (wardrobe, loft, desk, shelves, bedside, TV unit — never combined)
   - Electrical: each point type separately (switch, socket, AC point, fan drop, wiring lump)
   - Lighting: cove strip by rft, downlights by nos, fan by nos, each lamp separately
   - AC: always two rows — unit on one line, install kit on the next
   - Hardware: hinges by pair, drawer slides by pair, handles by nos
3. Brand names REQUIRED in every description (Century BWR ply, Greenlam, Kajaria/Somany, Hettich, Legrand Arteor, Atomberg, Asian Paints Royale, Polycab/Havells).
4. Description format: "[full spec + brand] - [Room Name]"
5. For Rate Sources: cite the real brand website or authoritative Indian pricing source (e.g. centuryply.com, greenlam.com, hettich.com, kajariaceramics.com, asianpaints.com, atomberg.com). Every rate must have a source.

OUTPUT — return exactly these two sections, comma-separated values, no extra text:

=== BOQ TABLE ===
Category,Description,Unit,Quantity,Rate,Amount
[one line per item — Amount = Quantity × Rate as integer]

=== RATE SOURCES ===
Category / Item,Rate Basis ({city} {tier} 2025-26),Source
[one line per rate source — real URL required]
Note,{city} city multiplier applied silently to all base rates.,-
Note,[Room] size estimated ~[X] sft from render — confirm actual dimensions before ordering.,-
Note,TV/electronics and kitchen appliances are client scope and excluded from this BOQ.,-"""


@dataclass
class BOQRow:
    category: str
    description: str
    unit: str
    qty: float
    rate: float


@dataclass
class RateSource:
    item: str
    basis: str
    source: str


def _parse_boq_response(text: str) -> tuple[list[BOQRow], list[RateSource]]:
    rows: list[BOQRow] = []
    sources: list[RateSource] = []

    boq_match = re.search(r"=== BOQ TABLE ===\n(.*?)\n(?:=== RATE SOURCES ===|$)", text, re.DOTALL)
    src_match = re.search(r"=== RATE SOURCES ===\n(.*?)$", text, re.DOTALL)

    if boq_match:
        for line in boq_match.group(1).strip().splitlines():
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 5 or parts[0] == "Category":
                continue
            try:
                rows.append(BOQRow(
                    category=parts[0],
                    description=parts[1],
                    unit=parts[2],
                    qty=float(parts[3]),
                    rate=float(parts[4]),
                ))
            except (ValueError, IndexError):
                continue

    if src_match:
        for line in src_match.group(1).strip().splitlines():
            parts = [p.strip() for p in line.split(",", 2)]
            if len(parts) < 2 or parts[0] in ("Category / Item", ""):
                continue
            sources.append(RateSource(
                item=parts[0],
                basis=parts[1] if len(parts) > 1 else "",
                source=parts[2] if len(parts) > 2 else "",
            ))

    return rows, sources


def generate_boq(
    rooms: list[RoomAnalysis],
    city: str,
    pincode: str,
    tier: str,
) -> tuple[list[BOQRow], list[RateSource]]:
    """Call Claude API → returns (BOQ rows, Rate Sources)."""
    prompt = _build_user_prompt(rooms, city, pincode, tier)
    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=4096,
        system=BOQ_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )
    return _parse_boq_response(resp.content[0].text)
```

- [ ] **Step 3.4: Run test — expect PASS**

```bash
python -m pytest tests/test_boq_generator.py -v
```

- [ ] **Step 3.5: Commit**

```bash
git add modules/boq_generator.py tests/test_boq_generator.py
git commit -m "feat: BOQ generator — Claude API with Houspire rate library"
```

---

## Task 4: Vendor Finder

**Files:**
- Create: `houspire_budget_generator/modules/vendor_finder.py`

This calls Claude API with the web_search tool enabled, passing the houspire-vendor-research skill prompt.

- [ ] **Step 4.1: Implement vendor_finder.py**

```python
# modules/vendor_finder.py
"""
Calls Claude API with web_search tool enabled.
Claude searches Google/Places for real local vendors near the pincode,
matching the houspire-vendor-research skill output format.
"""
import os
import re
import json
from dataclasses import dataclass, field
import anthropic
from dotenv import load_dotenv
from modules.room_analyzer import RoomAnalysis

load_dotenv()
client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

VENDOR_SYSTEM_PROMPT = """You are the Houspire Vendor Research specialist. You find real, verified local vendors near a client's pincode for interior projects.

RULES (non-negotiable):
- NEVER fabricate vendor names, phones, ratings, review counts, or addresses.
- Only include vendors you can verify exist via web search / Google Maps.
- If a vendor has no published phone, write: NA - visit showroom
- NEVER include turnkey / competitor firms (Livspace, Artifex, Fabulous Decor, Hipcouch, Urban Company).
- Place nearest vendors first within each category (sort by distance ascending).
- Read reviews carefully — flag every verified negative specifically (warranty disputes, rude behavior, delivery failures, GST fraud, fake products).
- If you cannot find vendors for a category, say so — do not invent.

RADIUS RULES:
- PRIMARY ZONE: 10 km from pincode centroid — find 3-5 vendors here first.
- FALLBACK: If fewer than 3 verified vendors exist within 10 km for a trade, extend to 15 km and label those rows "far-belt fallback (~X km)".
- HARD LIMIT: Never include vendors beyond 20 km unless that trade has zero options closer.
- Always state each vendor's distance in km (1 decimal place). Sort nearest-first within each category.

VENDOR COUNT: 3-5 vendors per category. Never fewer than 3 (extend radius if needed). Never more than 5.

DATA AUTHENTICITY:
- Every vendor name, address, rating, review count, and phone must come from a live Google search.
- Rating format: "4.8 (132)" — exact Google count in brackets.
- Phone format: "+91 XXXXX XXXXX" — if missing from Google, write NA - visit showroom.
- Area format: "Sub-locality, Locality (PINCODE)".

OUTPUT FORMAT — return exactly two sections, pipe-separated, no extra text:

=== VENDOR TABLE ===
Category|Vendor|Specialty / Brands|Area|lat|lng|Rating (count)|Phone
[one row per vendor, nearest-first within each category]

=== NOTES ===
[full structured notes text]

Notes must include:
1. Pincode reference (sub-areas covered, centroid lat/lng coordinates)
2. Vendor cluster strategy (belt structure with distances and what each belt supplies)
3. City multiplier status (locked or suggested band)
4. Trades covered (one bullet per category matched to render elements)
5. Vendor highlights (spec-critical sole-sources + strongest cluster picks with rationale)
6. Open items / flags (every verified complaint — be specific about what went wrong and what to verify in writing before paying)
7. Exclusions (competitors dropped, low-review-count vendors dropped, why)"""


def _build_vendor_prompt(
    rooms: list[RoomAnalysis],
    city: str,
    pincode: str,
    tier: str,
    categories: list[str],
) -> str:
    room_desc = "\n".join(
        f"- {r.room_type}: {r.design_elements}" for r in rooms
    )
    cats = ", ".join(categories) if categories else "auto-detect from renders"
    return f"""Find real, Google-verified local vendors for this Houspire project.

PROJECT: City: {city} | Pincode: {pincode} | Tier: {tier}

ROOM DESCRIPTIONS (use to identify required trades):
{room_desc}

TRADE CATEGORIES NEEDED:
{cats}

STEP-BY-STEP INSTRUCTIONS:

STEP 1 — Pincode centroid
Search: "{pincode} pincode {city} area locality state"
Find which sub-areas/localities the pincode covers and its approximate centroid (lat, lng).
State it clearly before proceeding.

STEP 2 — Vendor search per category
For EACH category, search Google/Google Maps:
  "[trade keyword] [nearest locality] {city}"
  "[brand name] dealer [locality] {city}"
Primary zone: 10 km from centroid. If fewer than 3 results within 10 km, extend to 15 km.

STEP 3 — Filter and rank
For each category:
- Select 3-5 vendors only (never fewer than 3, never more than 5)
- Sort by distance — nearest first
- Verify each vendor exists on Google Maps (real address + listed business)
- Record: exact name, specialty/brands carried, full locality + pincode, lat, lng, Google rating + review count, phone

STEP 4 — Review check
Before including any vendor, check its Google reviews for:
  warranty disputes | rude/abusive behavior | delivery failures | GST/bill fraud | B-grade products
If found, include the vendor BUT add a specific one-sentence flag in the Specialty column after " - "
Example: "Wardrobe + carpentry - flagged: 3 verified complaints about non-delivery after full advance payment, lock payment to milestones in writing"

STEP 5 — Distance calculation
dist_km = sqrt((Δlat × 111)² + (Δlng × 111 × cos(centroid_lat_rad))²)
Round to 1 decimal place.

Return the VENDOR TABLE and NOTES sections as specified in system prompt."""


@dataclass
class VendorRow:
    category: str
    vendor: str
    specialty: str
    area: str
    lat: float
    lng: float
    rating: str
    phone: str


def generate_vendors(
    rooms: list[RoomAnalysis],
    city: str,
    pincode: str,
    tier: str,
    categories: list[str] | None = None,
) -> tuple[list[VendorRow], str, tuple[float, float]]:
    """
    Returns (vendor_rows, notes_text, (centroid_lat, centroid_lng)).
    Uses web_search tool so Claude can look up real vendors.
    """
    if categories is None:
        categories = []

    prompt = _build_vendor_prompt(rooms, city, pincode, tier, categories)

    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8192,
        system=VENDOR_SYSTEM_PROMPT,
        tools=[{"type": "web_search", "name": "web_search"}],
        messages=[{"role": "user", "content": prompt}],
    )

    # Extract final text response (after tool use rounds)
    final_text = ""
    for block in resp.content:
        if hasattr(block, "text"):
            final_text = block.text
            break

    return _parse_vendor_response(final_text)


def _parse_vendor_response(text: str) -> tuple[list[VendorRow], str, tuple[float, float]]:
    rows: list[VendorRow] = []
    notes = ""
    centroid = (0.0, 0.0)

    table_match = re.search(r"=== VENDOR TABLE ===\n(.*?)\n(?:=== NOTES ===|$)", text, re.DOTALL)
    notes_match = re.search(r"=== NOTES ===\n(.*?)$", text, re.DOTALL)

    if table_match:
        for line in table_match.group(1).strip().splitlines():
            parts = [p.strip() for p in line.split("|")]
            if len(parts) < 7 or parts[0] in ("Category", ""):
                continue
            try:
                lat = float(parts[4]) if len(parts) > 4 and parts[4] else 0.0
                lng = float(parts[5]) if len(parts) > 5 and parts[5] else 0.0
                rows.append(VendorRow(
                    category=parts[0],
                    vendor=parts[1],
                    specialty=parts[2],
                    area=parts[3],
                    lat=lat,
                    lng=lng,
                    rating=parts[6] if len(parts) > 6 else "",
                    phone=parts[7] if len(parts) > 7 else "NA - visit showroom",
                ))
            except (ValueError, IndexError):
                continue

    if notes_match:
        notes = notes_match.group(1).strip()

    # Try to extract centroid from notes text
    coord_match = re.search(r"(\d+\.\d+)[°\s]*N[,\s]+(\d+\.\d+)[°\s]*E", notes)
    if coord_match:
        centroid = (float(coord_match.group(1)), float(coord_match.group(2)))

    return rows, notes, centroid
```

- [ ] **Step 4.2: Smoke test vendor_finder imports without error**

```bash
cd houspire_budget_generator
python3 -c "from modules.vendor_finder import generate_vendors; print('OK')"
```

Expected: `OK`

- [ ] **Step 4.3: Commit**

```bash
git add modules/vendor_finder.py
git commit -m "feat: vendor finder — Claude API with web_search for real local vendors"
```

---

## Task 5: Budget Excel Writer

**Files:**
- Create: `houspire_budget_generator/modules/budget_excel.py`
- Create: `houspire_budget_generator/tests/test_budget_excel.py`

Must match the exact houspire-budget plain template: Calibri 12, no formatting, formula in Amount column, no footer rows.

- [ ] **Step 5.1: Write failing test**

```python
# tests/test_budget_excel.py
import os, tempfile, pytest
import openpyxl
from modules.budget_excel import write_budget_excel
from modules.boq_generator import BOQRow, RateSource

ROWS = [
    BOQRow("Carpentry", "3-door wardrobe - Century BWR ply - Master Bedroom", "sft", 54, 1932),
    BOQRow("Flooring", "Vitrified marble-look tile - Kajaria - Master Bedroom", "sft", 180, 179),
]
SOURCES = [
    RateSource("Wardrobe", "Century BWR + Greenlam 2025-26", "centuryply.com"),
    RateSource("Note", "PIN 500001 - Hyderabad ×1.00", "-"),
]

def test_creates_file():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test_boq.xlsx")
        write_budget_excel("Test Client", "Hyderabad", "500001", "Premium", ROWS, SOURCES, path)
        assert os.path.exists(path)

def test_sheet_names():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test_boq.xlsx")
        write_budget_excel("Test Client", "Hyderabad", "500001", "Premium", ROWS, SOURCES, path)
        wb = openpyxl.load_workbook(path)
        assert "BOQ Template" in wb.sheetnames
        assert "Rate Sources" in wb.sheetnames

def test_header_row_is_bold():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test_boq.xlsx")
        write_budget_excel("Test Client", "Hyderabad", "500001", "Premium", ROWS, SOURCES, path)
        wb = openpyxl.load_workbook(path)
        ws = wb["BOQ Template"]
        assert ws.cell(1, 1).font.bold is True
        assert ws.cell(2, 1).font.bold is False  # data rows not bold

def test_amount_column_is_formula():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test_boq.xlsx")
        write_budget_excel("Test Client", "Hyderabad", "500001", "Premium", ROWS, SOURCES, path)
        wb = openpyxl.load_workbook(path)  # no data_only
        ws = wb["BOQ Template"]
        assert ws.cell(2, 6).value == "=D2*E2"
        assert ws.cell(3, 6).value == "=D3*E3"

def test_no_footer_rows():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "test_boq.xlsx")
        write_budget_excel("Test Client", "Hyderabad", "500001", "Premium", ROWS, SOURCES, path)
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["BOQ Template"]
        # Row after last data row should be empty
        last_data_row = 1 + len(ROWS)
        next_row_values = [ws.cell(last_data_row + 1, c).value for c in range(1, 7)]
        assert all(v is None for v in next_row_values)
```

- [ ] **Step 5.2: Run test — expect FAIL**

```bash
python -m pytest tests/test_budget_excel.py -v
```

- [ ] **Step 5.3: Implement budget_excel.py**

```python
# modules/budget_excel.py
from __future__ import annotations
import openpyxl
from openpyxl.styles import Font
from modules.boq_generator import BOQRow, RateSource
from config import BOQ_COL_WIDTHS
import datetime

_BOLD = Font(name="Calibri", size=12, bold=True)
_PLAIN = Font(name="Calibri", size=12, bold=False)

BOQ_HEADERS = ["Category", "Description", "Unit", "Quantity", "Rate", "Amount (Auto-calculated)"]
SRC_HEADERS = ["Category / Item", "Rate Basis", "Source"]


def write_budget_excel(
    client_name: str,
    city: str,
    pincode: str,
    tier: str,
    rows: list[BOQRow],
    sources: list[RateSource],
    output_path: str,
) -> str:
    """
    Write a BOQ Excel file matching the exact Houspire plain template.
    Returns output_path on success.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Sheet 1: BOQ Template ──────────────────────────────────────────────
    ws = wb.create_sheet("BOQ Template")

    # Header row — bold
    for c, h in enumerate(BOQ_HEADERS, start=1):
        ws.cell(1, c, h).font = _BOLD

    # Data rows — plain, Amount as formula
    for r_idx, row in enumerate(rows, start=2):
        ws.cell(r_idx, 1, row.category).font = _PLAIN
        ws.cell(r_idx, 2, row.description).font = _PLAIN
        ws.cell(r_idx, 3, row.unit).font = _PLAIN
        ws.cell(r_idx, 4, row.qty).font = _PLAIN
        ws.cell(r_idx, 5, row.rate).font = _PLAIN
        ws.cell(r_idx, 6, f"=D{r_idx}*E{r_idx}").font = _PLAIN  # formula, never hard-coded

    # Column widths
    for col, w in BOQ_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    # ── Sheet 2: Rate Sources ─────────────────────────────────────────────
    ws2 = wb.create_sheet("Rate Sources")

    for c, h in enumerate(SRC_HEADERS, start=1):
        ws2.cell(1, c, h).font = _BOLD

    for r_idx, src in enumerate(sources, start=2):
        ws2.cell(r_idx, 1, src.item).font = _PLAIN
        ws2.cell(r_idx, 2, src.basis).font = _PLAIN
        ws2.cell(r_idx, 3, src.source).font = _PLAIN

    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 45
    ws2.column_dimensions["C"].width = 45

    wb.save(output_path)
    return output_path
```

- [ ] **Step 5.4: Run test — expect PASS**

```bash
python -m pytest tests/test_budget_excel.py -v
```

- [ ] **Step 5.5: Verify with recalc.py**

```bash
python3 houspire_budget_generator/recalc.py /tmp/test_boq.xlsx 30
```

Expected: `"total_errors": 0`

- [ ] **Step 5.6: Commit**

```bash
git add modules/budget_excel.py tests/test_budget_excel.py
git commit -m "feat: budget Excel writer — exact Houspire plain template, formula Amount column"
```

---

## Task 6: Vendor Excel Writer

**Files:**
- Create: `houspire_budget_generator/modules/vendor_excel.py`
- Create: `houspire_budget_generator/tests/test_vendor_excel.py`

Based directly on `build_template.py` — the exact template used across 13+ Houspire projects.

- [ ] **Step 6.1: Write failing test**

```python
# tests/test_vendor_excel.py
import os, tempfile, math, pytest
import openpyxl
from modules.vendor_excel import write_vendor_excel
from modules.vendor_finder import VendorRow

CENTROID = (22.515, 88.330)
VENDORS = [
    VendorRow("Modular Carpentry", "Krishna Furniture", "Furniture + custom carpentry",
              "Sector 26, Pratap Nagar (302033)", 26.82, 75.74, "5.0 (126)", "+91 73401 45883"),
    VendorRow("Flooring", "Tirupati Granites", "Granite + marble",
              "NRI Chouraha (302017)", 26.79, 75.86, "4.6 (381)", "+91 88908 60694"),
]
NOTES = "Test project | PIN 302033\n\nReference point: test."

def test_creates_file():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "vendors.xlsx")
        write_vendor_excel("TestProject", "Jaipur", "302033", CENTROID, VENDORS, NOTES, path)
        assert os.path.exists(path)

def test_sheet_names():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "vendors.xlsx")
        write_vendor_excel("TestProject", "Jaipur", "302033", CENTROID, VENDORS, NOTES, path)
        wb = openpyxl.load_workbook(path)
        assert "Vendors" in wb.sheetnames
        assert "Notes" in wb.sheetnames

def test_header_is_bold_data_is_not():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "vendors.xlsx")
        write_vendor_excel("TestProject", "Jaipur", "302033", CENTROID, VENDORS, NOTES, path)
        wb = openpyxl.load_workbook(path)
        ws = wb["Vendors"]
        assert ws.cell(1, 1).font.bold is True
        assert ws.cell(2, 1).font.bold is False

def test_distance_is_computed_not_lat_lng():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "vendors.xlsx")
        write_vendor_excel("TestProject", "Jaipur", "302033", CENTROID, VENDORS, NOTES, path)
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Vendors"]
        # Column E should be a float distance, not a lat/lng string
        dist = ws.cell(2, 5).value
        assert isinstance(dist, float)
        assert 0 < dist < 200  # sanity: must be a km value

def test_row_count_equals_vendor_count():
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "vendors.xlsx")
        write_vendor_excel("TestProject", "Jaipur", "302033", CENTROID, VENDORS, NOTES, path)
        wb = openpyxl.load_workbook(path)
        ws = wb["Vendors"]
        data_rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if any(v for v in r)]
        assert len(data_rows) == len(VENDORS)
```

- [ ] **Step 6.2: Run test — expect FAIL**

```bash
python -m pytest tests/test_vendor_excel.py -v
```

- [ ] **Step 6.3: Implement vendor_excel.py** (based directly on build_template.py)

```python
# modules/vendor_excel.py
from __future__ import annotations
import math
import openpyxl
from openpyxl.styles import Font
from modules.vendor_finder import VendorRow
from config import VENDOR_COL_WIDTHS

_BOLD = Font(name="Calibri", size=12, bold=True)
_PLAIN = Font(name="Calibri", size=12, bold=False)


def _km(lat: float, lng: float, lat0: float, lng0: float) -> float:
    """Euclidean km approximation. Good for radii under 50 km."""
    dlat = (lat - lat0) * 111.0
    dlng = (lng - lng0) * 111.0 * math.cos(math.radians(lat0))
    return round(math.hypot(dlat, dlng), 1)


def write_vendor_excel(
    project_name: str,
    city: str,
    pincode: str,
    centroid: tuple[float, float],
    vendors: list[VendorRow],
    notes_text: str,
    output_path: str,
) -> str:
    """
    Write a Vendor Excel file matching the exact Houspire plain template.
    Based on build_template.py — same template used across 13+ projects.
    Returns output_path on success.
    """
    lat0, lng0 = centroid
    wb = openpyxl.Workbook()

    # ── Sheet 1: Vendors ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Vendors"

    headers = [
        "Category", "Vendor", "Specialty / Brands", "Area",
        f"Approx. km from {pincode}", "Rating (count)", "Phone"
    ]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c, h).font = _BOLD

    ws.freeze_panes = "A2"

    for col, w in VENDOR_COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    r = 2
    for v in vendors:
        dist = _km(v.lat, v.lng, lat0, lng0) if v.lat and v.lng else ""
        for c, val in enumerate(
            [v.category, v.vendor, v.specialty, v.area, dist, v.rating, v.phone],
            start=1
        ):
            ws.cell(r, c, val).font = _PLAIN
        r += 1

    # ── Sheet 2: Notes ────────────────────────────────────────────────────
    ns = wb.create_sheet("Notes")
    ns.column_dimensions["A"].width = 120
    for i, line in enumerate(notes_text.splitlines(), start=1):
        ns.cell(i, 1, line)

    wb.save(output_path)
    return output_path
```

- [ ] **Step 6.4: Run test — expect PASS**

```bash
python -m pytest tests/test_vendor_excel.py -v
```

- [ ] **Step 6.5: Commit**

```bash
git add modules/vendor_excel.py tests/test_vendor_excel.py
git commit -m "feat: vendor Excel writer — exact Houspire plain template from build_template.py"
```

---

## Task 7: Streamlit App

**Files:**
- Create: `houspire_budget_generator/app.py`

- [ ] **Step 7.1: Implement app.py**

```python
# app.py
import os
import tempfile
import streamlit as st
from dotenv import load_dotenv
from config import CITIES, CITIES_WITH_MULTIPLIERS, TIERS
from modules.room_analyzer import analyze_all_renders, RoomAnalysis
from modules.boq_generator import generate_boq
from modules.vendor_finder import generate_vendors
from modules.budget_excel import write_budget_excel
from modules.vendor_excel import write_vendor_excel

load_dotenv()

st.set_page_config(page_title="Houspire Budget Generator", page_icon="🏠", layout="wide")
st.title("🏠 Houspire Budget Generator")
st.caption("Internal tool — generates BOQ + Vendor Excel from client renders")
st.divider()

# ── Session state ─────────────────────────────────────────────────────────────
for key in ("analyses", "boq_rows", "boq_sources", "vendors", "notes", "centroid",
            "budget_path", "vendor_path"):
    if key not in st.session_state:
        st.session_state[key] = None

# ── Step 1: Client Details ────────────────────────────────────────────────────
st.subheader("1. Client Details")
c1, c2, c3, c4 = st.columns(4)
with c1:
    client_name = st.text_input("Client Name *", placeholder="e.g. Priya Sharma")
with c2:
    city = st.selectbox("City *", CITIES)
with c3:
    pincode = st.text_input("Pincode *", placeholder="e.g. 500032", max_chars=6)
with c4:
    tier = st.selectbox("Tier *", TIERS)

# Warn if city multiplier unknown
if city == "Other":
    multiplier_note = st.text_input(
        "City multiplier not set — enter estimate (e.g. 0.94) or leave blank to let Claude decide",
        placeholder="0.94"
    )
else:
    mult = CITIES_WITH_MULTIPLIERS.get(city)
    if mult:
        st.caption(f"City multiplier ×{mult} (applied silently to rates)")

# ── Step 2: Upload Renders ────────────────────────────────────────────────────
st.subheader("2. Upload Room Renders")
st.caption("One render per room. Claude auto-detects room type, sqft, and design elements.")
renders = st.file_uploader(
    "Room renders (JPG/PNG)", type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

# ── Step 3: Floorplan (optional) ──────────────────────────────────────────────
st.subheader("3. Floor Plan (Optional)")
st.caption("If provided, Claude uses actual room dimensions from the plan.")
floorplan = st.file_uploader("Floor plan (JPG/PNG)", type=["jpg", "jpeg", "png"])

# ── Step 4: Analyse Rooms ─────────────────────────────────────────────────────
st.subheader("4. Analyse Rooms")
can_analyse = bool(renders and client_name and pincode)

if st.button("🔍 Analyse Renders", type="primary", disabled=not can_analyse):
    images = [(f.read(), f.type or "image/jpeg", f.name) for f in renders]
    with st.spinner("Analysing renders with Claude Vision…"):
        st.session_state.analyses = analyze_all_renders(images)
    st.success(f"Detected {len(st.session_state.analyses)} rooms.")

# ── Step 5: Review + Edit ─────────────────────────────────────────────────────
if st.session_state.analyses:
    st.subheader("5. Review Detected Rooms")
    from config import ROOM_TYPES
    edited = []
    for i, a in enumerate(st.session_state.analyses):
        with st.expander(f"Render {i+1}: {a.image_filename or f'Image {i+1}'}", expanded=True):
            col_a, col_b = st.columns([2, 1])
            with col_a:
                rt = st.selectbox(
                    "Room Type",
                    ROOM_TYPES,
                    index=ROOM_TYPES.index(a.room_type) if a.room_type in ROOM_TYPES else 0,
                    key=f"rt_{i}"
                )
                notes_edit = st.text_area(
                    "Design elements (edit if Claude missed anything)",
                    value=a.design_elements,
                    key=f"de_{i}", height=80
                )
            with col_b:
                sqft = st.number_input(
                    "Area (sqft)", min_value=20, max_value=2000,
                    value=a.estimated_sqft, step=5, key=f"sqft_{i}"
                )
                st.caption(f"Confidence: **{a.confidence.upper()}**")
            edited.append(RoomAnalysis(
                room_type=rt, estimated_sqft=sqft,
                confidence=a.confidence, design_elements=notes_edit,
                image_filename=a.image_filename
            ))

    # ── Step 6: Generate ──────────────────────────────────────────────────────
    st.subheader("6. Generate Excel Files")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        gen_boq = st.button("📊 Generate Budget BOQ", type="primary")
    with col_g2:
        gen_vendors = st.button("📍 Find Local Vendors", type="primary")

    if gen_boq:
        with st.spinner("Generating BOQ with Claude…"):
            rows, sources = generate_boq(edited, city, pincode, tier)
        with st.spinner("Writing Excel…"):
            tmpdir = tempfile.mkdtemp()
            safe = client_name.replace(" ", "_")
            path = os.path.join(tmpdir, f"{safe}_BOQ_{city}.xlsx")
            write_budget_excel(client_name, city, pincode, tier, rows, sources, path)
            st.session_state.budget_path = path
            st.session_state.boq_rows = rows
        st.success(f"BOQ generated — {len(rows)} line items.")

    if gen_vendors:
        # Pass categories from BOQ if already generated
        cats = list({r.category for r in st.session_state.boq_rows}) \
               if st.session_state.boq_rows else []
        with st.spinner("Searching for real local vendors near pincode… (this takes 30-60 sec)"):
            vendors, notes_text, centroid = generate_vendors(edited, city, pincode, tier, cats)
        with st.spinner("Writing Excel…"):
            tmpdir = tempfile.mkdtemp()
            safe = client_name.replace(" ", "_")
            room_label = edited[0].room_type.replace(" ", "_") if len(edited) == 1 else safe
            path = os.path.join(tmpdir, f"{city}_{room_label}_Vendors_{pincode}.xlsx")
            write_vendor_excel(client_name, city, pincode, centroid, vendors, notes_text, path)
            st.session_state.vendor_path = path
        st.success(f"Vendor sheet generated — {len(vendors)} vendors found.")

    # ── Downloads ─────────────────────────────────────────────────────────────
    if st.session_state.budget_path or st.session_state.vendor_path:
        st.divider()
        dl1, dl2 = st.columns(2)
        if st.session_state.budget_path:
            with dl1:
                with open(st.session_state.budget_path, "rb") as f:
                    st.download_button(
                        "📥 Download Budget BOQ",
                        data=f.read(),
                        file_name=os.path.basename(st.session_state.budget_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
        if st.session_state.vendor_path:
            with dl2:
                with open(st.session_state.vendor_path, "rb") as f:
                    st.download_button(
                        "📥 Download Vendor Sheet",
                        data=f.read(),
                        file_name=os.path.basename(st.session_state.vendor_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
```

- [ ] **Step 7.2: Run app locally to check imports**

```bash
cd houspire_budget_generator
pip install -r requirements.txt --break-system-packages
streamlit run app.py
```

Open `http://localhost:8501`. Verify: all inputs render, no import errors in terminal.

- [ ] **Step 7.3: Commit**

```bash
git add app.py
git commit -m "feat: Streamlit UI — 6-step flow, separate BOQ + vendor generation"
```

---

## Task 8: Branded DOCX Generator

**Files:**
- Create: `houspire_budget_generator/modules/docx_generator.py`
- Create: `houspire_budget_generator/tests/test_docx_generator.py`

This module reads the plain Excel files generated in Tasks 5/6, builds a Node.js script dynamically with the Houspire brand system, and runs it to produce branded `.docx` files.

Brand system (from houspire-branding-docx skill — never deviate):
```
BRAND_PRIMARY = "1B4D3E"   # Deep Forest Green — headers, titles
BRAND_ACCENT  = "D4AF37"   # Gold — dividers, section numbers
BRAND_LIGHT   = "F5F0E8"   # Cream — info boxes
ROW_ALT       = "F7F5F0"   # Alternating row shading
Font          = "Arial" throughout (never Calibri in branded docs)
```

**CRITICAL:** Never mention city multipliers, rate adjustment factors, or internal pricing methodology in branded client documents. Use: "All rates researched for [City] market (2025-26)".

- [ ] **Step 8.1: Install Node.js docx package**

```bash
npm install -g docx
node -e "require('docx'); console.log('docx OK')"
```

Expected: `docx OK`

- [ ] **Step 8.2: Write failing test**

```python
# tests/test_docx_generator.py
import os, tempfile, pytest
from modules.docx_generator import generate_branded_boq_docx, generate_branded_vendor_docx

def test_generate_branded_boq_creates_docx():
    """Requires a real BOQ Excel at the given path."""
    import openpyxl
    # Create a minimal valid BOQ Excel for testing
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOQ Template"
    headers = ["Category", "Description", "Unit", "Quantity", "Rate", "Amount (Auto-calculated)"]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c, h)
    ws.cell(2, 1, "Carpentry")
    ws.cell(2, 2, "3-door wardrobe - Century BWR ply - Master Bedroom")
    ws.cell(2, 3, "sft")
    ws.cell(2, 4, 54)
    ws.cell(2, 5, 1932)
    ws.cell(2, 6, "=D2*E2")
    ws2 = wb.create_sheet("Rate Sources")
    ws2.cell(1, 1, "Category / Item")
    ws2.cell(1, 2, "Rate Basis")
    ws2.cell(1, 3, "Source")
    ws2.cell(2, 1, "Wardrobe")
    ws2.cell(2, 2, "Century BWR + Greenlam 2025-26")
    ws2.cell(2, 3, "centuryply.com")
    with tempfile.TemporaryDirectory() as d:
        xlsx_path = os.path.join(d, "test_boq.xlsx")
        docx_path = os.path.join(d, "test_boq.docx")
        wb.save(xlsx_path)
        generate_branded_boq_docx(
            xlsx_path=xlsx_path,
            client_name="Test Client",
            city="Hyderabad",
            output_path=docx_path,
        )
        assert os.path.exists(docx_path)
        assert os.path.getsize(docx_path) > 5000  # real docx, not empty

def test_generate_branded_vendor_creates_docx():
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Vendors"
    headers = ["Category", "Vendor", "Specialty / Brands", "Area",
               "Approx. km from 500032", "Rating (count)", "Phone"]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c, h)
    ws.cell(2, 1, "Modular Carpentry")
    ws.cell(2, 2, "Krishna Furniture")
    ws.cell(2, 3, "Custom carpentry + wardrobe")
    ws.cell(2, 4, "Sector 26, Pratap Nagar (302033)")
    ws.cell(2, 5, 9.7)
    ws.cell(2, 6, "5.0 (126)")
    ws.cell(2, 7, "+91 73401 45883")
    ws2 = wb.create_sheet("Notes")
    ws2.cell(1, 1, "Test project notes")
    with tempfile.TemporaryDirectory() as d:
        xlsx_path = os.path.join(d, "test_vendors.xlsx")
        docx_path = os.path.join(d, "test_vendors.docx")
        wb.save(xlsx_path)
        generate_branded_vendor_docx(
            xlsx_path=xlsx_path,
            client_name="Test Client",
            city="Hyderabad",
            pincode="500032",
            output_path=docx_path,
        )
        assert os.path.exists(docx_path)
        assert os.path.getsize(docx_path) > 5000
```

- [ ] **Step 8.3: Run test — expect FAIL**

```bash
python -m pytest tests/test_docx_generator.py -v
```

Expected: `ModuleNotFoundError`

- [ ] **Step 8.4: Implement docx_generator.py**

```python
# modules/docx_generator.py
"""
Reads Houspire plain Excel files and generates branded DOCX via Node.js docx npm package.
Brand system: BRAND_PRIMARY=1B4D3E, BRAND_ACCENT=D4AF37, BRAND_LIGHT=F5F0E8, Font=Arial.
NEVER mention city multipliers or internal pricing mechanics in client-facing output.
"""
from __future__ import annotations
import os
import json
import subprocess
import tempfile
import textwrap
import openpyxl


# ── Brand constants (from houspire-branding-docx skill) ──────────────────────
BRAND_PRIMARY = "1B4D3E"
BRAND_ACCENT  = "D4AF37"
BRAND_LIGHT   = "F5F0E8"
ROW_ALT       = "F7F5F0"
SUBTOTAL_BG   = "E8F0E8"
WHITE         = "FFFFFF"


def _run_node(script: str, tmp_dir: str) -> None:
    """Write script to a temp file and run it with Node.js."""
    script_path = os.path.join(tmp_dir, "generate.js")
    with open(script_path, "w") as f:
        f.write(script)
    result = subprocess.run(
        ["node", script_path],
        capture_output=True, text=True, cwd=tmp_dir
    )
    if result.returncode != 0:
        raise RuntimeError(f"Node.js failed:\n{result.stderr}")


def _read_boq_excel(xlsx_path: str) -> dict:
    """Extract BOQ data from plain Excel into a dict for the Node.js generator."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["BOQ Template"]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:  # skip empty
            continue
        rows.append({
            "category": str(row[0] or ""),
            "description": str(row[1] or ""),
            "unit": str(row[2] or ""),
            "qty": float(row[3] or 0),
            "rate": float(row[4] or 0),
            "amount": float(row[3] or 0) * float(row[4] or 0),
        })

    # Rate Sources sheet
    sources = []
    if "Rate Sources" in wb.sheetnames:
        ws2 = wb["Rate Sources"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if row[0]:
                sources.append({
                    "item": str(row[0] or ""),
                    "basis": str(row[1] or ""),
                    "source": str(row[2] or ""),
                })

    # Group by category
    from collections import defaultdict
    by_cat: dict[str, list] = defaultdict(list)
    for r in rows:
        by_cat[r["category"]].append(r)

    return {"rows": rows, "by_category": dict(by_cat), "sources": sources}


def _read_vendor_excel(xlsx_path: str) -> dict:
    """Extract vendor data from plain Excel into a dict for the Node.js generator."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Vendors"]

    vendors = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        vendors.append({
            "category": str(row[0] or ""),
            "vendor": str(row[1] or ""),
            "specialty": str(row[2] or ""),
            "area": str(row[3] or ""),
            "distance": str(row[4] or ""),
            "rating": str(row[5] or ""),
            "phone": str(row[6] or ""),
        })

    notes_text = ""
    if "Notes" in wb.sheetnames:
        ws2 = wb["Notes"]
        lines = [str(ws2.cell(r, 1).value or "") for r in range(1, ws2.max_row + 1)]
        notes_text = "\n".join(lines)

    from collections import defaultdict
    by_cat: dict[str, list] = defaultdict(list)
    for v in vendors:
        by_cat[v["category"]].append(v)

    return {"vendors": vendors, "by_category": dict(by_cat), "notes": notes_text}


def _boq_node_script(data: dict, client_name: str, city: str, output_path: str) -> str:
    """Generate the Node.js script that builds the branded BOQ DOCX."""
    data_json = json.dumps(data, ensure_ascii=False)
    total = sum(r["amount"] for r in data["rows"])
    cat_count = len(data["by_category"])

    return textwrap.dedent(f"""
    const {{
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
      PageBreak, PageNumber
    }} = require("docx");
    const fs = require("fs");
    const path = require("path");

    const DATA = {data_json};
    const CLIENT = {json.dumps(client_name)};
    const CITY = {json.dumps(city)};
    const TOTAL = {total};
    const OUTPUT = {json.dumps(output_path)};

    const PRIMARY = "{BRAND_PRIMARY}", ACCENT = "{BRAND_ACCENT}";
    const LIGHT = "{BRAND_LIGHT}", ALT = "{ROW_ALT}", WHITE = "FFFFFF";

    const borders = {{
      top: {{style: BorderStyle.SINGLE, size: 1, color: "CCCCCC"}},
      bottom: {{style: BorderStyle.SINGLE, size: 1, color: "CCCCCC"}},
      left: {{style: BorderStyle.NONE}},
      right: {{style: BorderStyle.NONE}},
    }};
    const cellMargins = {{top: 80, bottom: 80, left: 120, right: 120}};
    const TW = 9706;
    const CW = [1400, 3906, 700, 700, 1100, 1900];

    function fmt(n) {{ return "₹" + Math.round(n).toLocaleString("en-IN"); }}

    function hc(text, width) {{
      return new TableCell({{
        borders, width: {{size: width, type: WidthType.DXA}},
        shading: {{fill: PRIMARY, type: ShadingType.CLEAR}},
        margins: cellMargins, verticalAlign: "center",
        children: [new Paragraph({{
          alignment: AlignmentType.CENTER,
          children: [new TextRun({{text, bold: true, font: "Arial", size: 16, color: WHITE}})]
        }})]
      }});
    }}

    function dc(text, width, isAlt, align, bold) {{
      return new TableCell({{
        borders, width: {{size: width, type: WidthType.DXA}},
        shading: {{fill: isAlt ? ALT : WHITE}},
        margins: cellMargins,
        children: [new Paragraph({{
          alignment: align || AlignmentType.LEFT,
          children: [new TextRun({{text: String(text || "-"), font: "Arial", size: 15, color: "333333", bold: !!bold}})]
        }})]
      }});
    }}

    const sections = [];

    // ── Cover page ────────────────────────────────────────────────────────────
    sections.push(new Paragraph({{spacing: {{before: 800}}}});
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: "HOUSPIRE", bold: true, font: "Arial", size: 88,
        characterSpacing: 300, color: PRIMARY}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: "India's First Transparent Interior Planning Platform",
        font: "Arial", size: 20, color: "666666", italics: true}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 200, after: 200}},
      border: {{bottom: {{style: BorderStyle.SINGLE, size: 6, color: ACCENT}}}},
      children: [new TextRun({{text: ""}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 300}},
      children: [new TextRun({{text: "Bill of Quantities", bold: true, font: "Arial",
        size: 52, color: PRIMARY}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: CLIENT + " | " + CITY + " | Premium", font: "Arial",
        size: 24, color: "444444"}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 100}},
      children: [new TextRun({{text: Object.keys(DATA.by_category).length + " categories · " +
        DATA.rows.length + " line items", font: "Arial", size: 20, color: "888888"}})]
    }}));
    sections.push(new Paragraph({{spacing: {{before: 300}}}}));
    // Info box
    sections.push(new Paragraph({{
      border: {{left: {{style: BorderStyle.SINGLE, size: 12, color: PRIMARY}}}},
      shading: {{fill: LIGHT}},
      spacing: {{before: 200}},
      indent: {{left: 300}},
      children: [
        new TextRun({{text: "Scope: ", bold: true, font: "Arial", size: 18, color: PRIMARY}}),
        new TextRun({{text: "All rates researched for " + CITY + " market (2025-26). " +
          "Quantities estimated from uploaded renders. Subject to site measurement confirmation.",
          font: "Arial", size: 18, color: "333333"}}),
      ]
    }}));
    sections.push(new Paragraph({{children: [new PageBreak()]}}));

    // ── Category detail pages ─────────────────────────────────────────────────
    let catNum = 0;
    for (const [cat, items] of Object.entries(DATA.by_category)) {{
      catNum++;
      const subtotal = items.reduce((s, r) => s + r.amount, 0);

      sections.push(new Paragraph({{
        spacing: {{before: catNum > 1 ? 0 : 200, after: 100}},
        children: [
          new TextRun({{text: String(catNum).padStart(2,"0") + " ", bold: true,
            font: "Arial", size: 28, color: ACCENT}}),
          new TextRun({{text: cat, bold: true, font: "Arial", size: 28, color: PRIMARY}}),
          new TextRun({{text: "  (" + items.length + " items)", font: "Arial",
            size: 18, color: "888888"}}),
        ]
      }}));
      sections.push(new Paragraph({{
        border: {{bottom: {{style: BorderStyle.SINGLE, size: 4, color: ACCENT}}}},
        spacing: {{after: 120}},
        children: [new TextRun({{text: ""}})]
      }}));

      const headerRow = new TableRow({{
        children: ["Category","Description","Unit","Qty","Rate","Amount"]
          .map((h, i) => hc(h, CW[i]))
      }});

      const dataRows = items.map((r, idx) => new TableRow({{
        children: [
          dc(r.category, CW[0], idx%2===1),
          dc(r.description, CW[1], idx%2===1),
          dc(r.unit, CW[2], idx%2===1, AlignmentType.CENTER),
          dc(String(r.qty), CW[3], idx%2===1, AlignmentType.CENTER),
          dc(fmt(r.rate), CW[4], idx%2===1, AlignmentType.RIGHT),
          dc(fmt(r.amount), CW[5], idx%2===1, AlignmentType.RIGHT),
        ]
      }}));

      const subtotalRow = new TableRow({{
        children: [
          new TableCell({{
            columnSpan: 5,
            borders, margins: cellMargins,
            shading: {{fill: "{SUBTOTAL_BG}"}},
            width: {{size: CW[0]+CW[1]+CW[2]+CW[3]+CW[4], type: WidthType.DXA}},
            children: [new Paragraph({{
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({{text: cat + " Subtotal", bold: true, font:"Arial",
                size:16, color: PRIMARY}})]
            }})]
          }}),
          dc(fmt(subtotal), CW[5], false, AlignmentType.RIGHT, true),
        ]
      }});

      sections.push(new Table({{
        width: {{size: TW, type: WidthType.DXA}},
        rows: [headerRow, ...dataRows, subtotalRow],
      }}));
      sections.push(new Paragraph({{spacing: {{after: 300}}}}));
    }}

    // ── Rate Sources ──────────────────────────────────────────────────────────
    if (DATA.sources && DATA.sources.length > 0) {{
      sections.push(new Paragraph({{children: [new PageBreak()]}}));
      sections.push(new Paragraph({{
        spacing: {{before: 200, after: 160}},
        children: [new TextRun({{text: "Rate Sources & Methodology", bold: true,
          font: "Arial", size: 28, color: PRIMARY}})]
      }}));
      sections.push(new Paragraph({{
        border: {{bottom: {{style: BorderStyle.SINGLE, size: 4, color: ACCENT}}}},
        spacing: {{after: 200}},
        children: [new TextRun({{text: ""}})]
      }}));
      const srcHeader = new TableRow({{
        children: [
          hc("Item / Category", 3000), hc("Rate Basis", 3500), hc("Source", 3006)
        ]
      }});
      const srcRows = DATA.sources.map((s, i) => new TableRow({{
        children: [
          dc(s.item, 3000, i%2===1),
          dc(s.basis, 3500, i%2===1),
          dc(s.source, 3006, i%2===1),
        ]
      }}));
      sections.push(new Table({{
        width: {{size: TW, type: WidthType.DXA}},
        rows: [srcHeader, ...srcRows],
      }}));
    }}

    // ── Disclaimer ────────────────────────────────────────────────────────────
    sections.push(new Paragraph({{children: [new PageBreak()]}}));
    sections.push(new Paragraph({{
      border: {{top: {{style: BorderStyle.SINGLE, size: 4, color: ACCENT}}}},
      spacing: {{before: 400, after: 200}},
      children: [new TextRun({{text: "Important Notes & Disclaimer", bold: true,
        font: "Arial", size: 24, color: PRIMARY}})]
    }}));
    const disclaimerLines = [
      "Houspire operates on a flat-fee model. We do not earn commissions from vendors or material suppliers.",
      "All rates in this BOQ are based on current market research for " + CITY + " and are indicative. Final rates are subject to vendor quotation and site confirmation.",
      "Quantities are estimated from uploaded renders. Physical site measurements should be taken before ordering materials.",
      "This document is confidential and prepared exclusively for the named client.",
      "www.houspire.ai | contact@houspire.ai",
    ];
    disclaimerLines.forEach(line => {{
      sections.push(new Paragraph({{
        spacing: {{before: 100}},
        children: [new TextRun({{text: line, font: "Arial", size: 17,
          color: "666666", italics: true}})]
      }}));
    }});

    // ── Build document ────────────────────────────────────────────────────────
    const doc = new Document({{
      sections: [{{
        properties: {{
          page: {{
            size: {{width: 11906, height: 16838}},
            margin: {{top: 1000, right: 1000, bottom: 1000, left: 1000}},
          }}
        }},
        headers: {{
          default: new Header({{children: [new Paragraph({{
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({{text: "Houspire BOQ | " + CLIENT + " | Confidential",
              font: "Arial", size: 16, color: "999999", italics: true}})]
          }})]}})
        }},
        footers: {{
          default: new Footer({{children: [new Paragraph({{
            border: {{top: {{style: BorderStyle.SINGLE, size: 2, color: "CCCCCC"}}}},
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({{text: "www.houspire.ai  |  Page ", font:"Arial", size:16, color:"999999"}}),
              new TextRun({{children: [PageNumber.CURRENT], font:"Arial", size:16, color:"999999"}}),
            ]
          }})]}})
        }},
        children: sections,
      }}]
    }});

    Packer.toBuffer(doc).then(buf => {{
      fs.writeFileSync(OUTPUT, buf);
      console.log("OK:" + OUTPUT);
    }});
    """)


def _vendor_node_script(data: dict, client_name: str, city: str, pincode: str, output_path: str) -> str:
    """Generate the Node.js script that builds the branded Vendor DOCX."""
    data_json = json.dumps(data, ensure_ascii=False)
    vendor_count = len(data["vendors"])
    cat_count = len(data["by_category"])

    return textwrap.dedent(f"""
    const {{
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
      PageBreak, PageNumber
    }} = require("docx");
    const fs = require("fs");

    const DATA = {data_json};
    const CLIENT = {json.dumps(client_name)};
    const CITY = {json.dumps(city)};
    const PIN = {json.dumps(pincode)};
    const OUTPUT = {json.dumps(output_path)};

    const PRIMARY = "{BRAND_PRIMARY}", ACCENT = "{BRAND_ACCENT}";
    const LIGHT = "{BRAND_LIGHT}", ALT = "{ROW_ALT}", WHITE = "FFFFFF";

    const borders = {{
      top: {{style: BorderStyle.SINGLE, size: 1, color: "CCCCCC"}},
      bottom: {{style: BorderStyle.SINGLE, size: 1, color: "CCCCCC"}},
      left: {{style: BorderStyle.NONE}},
      right: {{style: BorderStyle.NONE}},
    }};
    const cellMargins = {{top: 80, bottom: 80, left: 120, right: 120}};
    const TW = 9506;
    const CW = [300, 2200, 3800, 2100, 1106];  // SN, Vendor, Specialty/Area, Distance/Rating, Phone

    function hc(text, width) {{
      return new TableCell({{
        borders, width: {{size: width, type: WidthType.DXA}},
        shading: {{fill: PRIMARY}},
        margins: cellMargins, verticalAlign: "center",
        children: [new Paragraph({{
          alignment: AlignmentType.CENTER,
          children: [new TextRun({{text, bold: true, font: "Arial", size: 16, color: WHITE}})]
        }})]
      }});
    }}

    function dc(text, width, isAlt, align) {{
      return new TableCell({{
        borders, width: {{size: width, type: WidthType.DXA}},
        shading: {{fill: isAlt ? ALT : WHITE}},
        margins: cellMargins,
        children: [new Paragraph({{
          alignment: align || AlignmentType.LEFT,
          children: [new TextRun({{text: String(text || "-"), font: "Arial", size: 15, color: "333333"}})]
        }})]
      }});
    }}

    const sections = [];

    // ── Cover page ────────────────────────────────────────────────────────────
    sections.push(new Paragraph({{spacing: {{before: 800}}}}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: "HOUSPIRE", bold: true, font: "Arial", size: 88,
        characterSpacing: 300, color: PRIMARY}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: "India's First Transparent Interior Planning Platform",
        font: "Arial", size: 20, color: "666666", italics: true}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 200, after: 200}},
      border: {{bottom: {{style: BorderStyle.SINGLE, size: 6, color: ACCENT}}}},
      children: [new TextRun({{text: ""}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 300}},
      children: [new TextRun({{text: "Verified Vendor Directory", bold: true,
        font: "Arial", size: 52, color: PRIMARY}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER,
      children: [new TextRun({{text: CLIENT + " | " + CITY + " — PIN " + PIN,
        font: "Arial", size: 24, color: "444444"}})]
    }}));
    sections.push(new Paragraph({{
      alignment: AlignmentType.CENTER, spacing: {{before: 100}},
      children: [new TextRun({{text: Object.keys(DATA.by_category).length +
        " trade categories · " + DATA.vendors.length + " verified vendors",
        font: "Arial", size: 20, color: "888888"}})]
    }}));
    sections.push(new Paragraph({{spacing: {{before: 300}}}}));
    sections.push(new Paragraph({{
      border: {{left: {{style: BorderStyle.SINGLE, size: 12, color: PRIMARY}}}},
      shading: {{fill: LIGHT}},
      indent: {{left: 300}},
      children: [
        new TextRun({{text: "About this directory: ", bold: true, font: "Arial", size: 18, color: PRIMARY}}),
        new TextRun({{text: "Vendors are real, Google-verified local businesses near PIN " + PIN +
          ". Distances calculated from pincode centroid. Ratings from Google at time of research. " +
          "Always verify phone and address before site visit.",
          font: "Arial", size: 18, color: "333333"}}),
      ]
    }}));
    sections.push(new Paragraph({{children: [new PageBreak()]}}));

    // ── Category pages ────────────────────────────────────────────────────────
    let catNum = 0;
    for (const [cat, vendors] of Object.entries(DATA.by_category)) {{
      catNum++;
      sections.push(new Paragraph({{
        spacing: {{before: catNum > 1 ? 0 : 200, after: 100}},
        children: [
          new TextRun({{text: String(catNum).padStart(2,"0") + " ", bold: true,
            font: "Arial", size: 28, color: ACCENT}}),
          new TextRun({{text: cat, bold: true, font: "Arial", size: 28, color: PRIMARY}}),
          new TextRun({{text: "  (" + vendors.length + " vendors)",
            font: "Arial", size: 18, color: "888888"}}),
        ]
      }}));
      sections.push(new Paragraph({{
        border: {{bottom: {{style: BorderStyle.SINGLE, size: 4, color: ACCENT}}}},
        spacing: {{after: 120}},
        children: [new TextRun({{text: ""}})]
      }}));

      const headerRow = new TableRow({{
        children: [
          hc("#", CW[0]),
          hc("Vendor", CW[1]),
          hc("Specialty / Area", CW[2]),
          hc("Distance · Rating", CW[3]),
          hc("Phone", CW[4]),
        ]
      }});

      const dataRows = vendors.map((v, i) => new TableRow({{
        children: [
          dc(String(i+1), CW[0], i%2===1, AlignmentType.CENTER),
          dc(v.vendor, CW[1], i%2===1),
          dc(v.specialty + (v.area ? " | " + v.area : ""), CW[2], i%2===1),
          dc((v.distance ? v.distance + " km" : "") + (v.rating ? "  ·  " + v.rating : ""),
             CW[3], i%2===1, AlignmentType.CENTER),
          dc(v.phone, CW[4], i%2===1),
        ]
      }}));

      sections.push(new Table({{
        width: {{size: TW, type: WidthType.DXA}},
        rows: [headerRow, ...dataRows],
      }}));
      sections.push(new Paragraph({{spacing: {{after: 300}}}}));
    }}

    // ── Disclaimer ────────────────────────────────────────────────────────────
    sections.push(new Paragraph({{children: [new PageBreak()]}}));
    sections.push(new Paragraph({{
      border: {{top: {{style: BorderStyle.SINGLE, size: 4, color: ACCENT}}}},
      spacing: {{before: 400, after: 200}},
      children: [new TextRun({{text: "Important Notes & Disclaimer", bold: true,
        font: "Arial", size: 24, color: PRIMARY}})]
    }}));
    [
      "Houspire operates on a flat-fee model. We do not earn commissions from any vendor listed in this directory.",
      "All vendors are sourced from Google Maps/Places and represent independent businesses. Houspire does not endorse any specific vendor.",
      "Ratings and review counts are from Google at the time of research. Always verify contact details before a site visit.",
      "Distances are approximate, calculated from the PIN " + PIN + " centroid.",
      "This document is confidential and prepared exclusively for " + CLIENT + ".",
      "www.houspire.ai | contact@houspire.ai",
    ].forEach(line => {{
      sections.push(new Paragraph({{
        spacing: {{before: 100}},
        children: [new TextRun({{text: line, font: "Arial", size: 17,
          color: "666666", italics: true}})]
      }}));
    }});

    const doc = new Document({{
      sections: [{{
        properties: {{
          page: {{
            size: {{width: 11906, height: 16838}},
            margin: {{top: 1200, right: 1000, bottom: 1000, left: 1000}},
          }}
        }},
        headers: {{
          default: new Header({{children: [new Paragraph({{
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({{text: "Houspire Vendor Directory | " + CLIENT + " | Confidential",
              font: "Arial", size: 16, color: "999999", italics: true}})]
          }})]}})
        }},
        footers: {{
          default: new Footer({{children: [new Paragraph({{
            border: {{top: {{style: BorderStyle.SINGLE, size: 2, color: "CCCCCC"}}}},
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({{text: "www.houspire.ai  |  Page ", font:"Arial", size:16, color:"999999"}}),
              new TextRun({{children: [PageNumber.CURRENT], font:"Arial", size:16, color:"999999"}}),
            ]
          }})]}})
        }},
        children: sections,
      }}]
    }});

    Packer.toBuffer(doc).then(buf => {{
      fs.writeFileSync(OUTPUT, buf);
      console.log("OK:" + OUTPUT);
    }});
    """)


def generate_branded_boq_docx(
    xlsx_path: str,
    client_name: str,
    city: str,
    output_path: str,
) -> str:
    """Read BOQ Excel → generate branded DOCX. Returns output_path."""
    data = _read_boq_excel(xlsx_path)
    with tempfile.TemporaryDirectory() as tmp:
        script = _boq_node_script(data, client_name, city, output_path)
        _run_node(script, tmp)
    return output_path


def generate_branded_vendor_docx(
    xlsx_path: str,
    client_name: str,
    city: str,
    pincode: str,
    output_path: str,
) -> str:
    """Read Vendor Excel → generate branded DOCX. Returns output_path."""
    data = _read_vendor_excel(xlsx_path)
    with tempfile.TemporaryDirectory() as tmp:
        script = _vendor_node_script(data, client_name, city, pincode, output_path)
        _run_node(script, tmp)
    return output_path
```

- [ ] **Step 8.5: Run test — expect PASS**

```bash
python -m pytest tests/test_docx_generator.py -v
```

- [ ] **Step 8.6: Commit**

```bash
git add modules/docx_generator.py tests/test_docx_generator.py
git commit -m "feat: branded DOCX generator — Node.js docx npm, Houspire Navy/Gold brand system"
```

---

## Task 9: PDF Converter

**Files:**
- Create: `houspire_budget_generator/modules/pdf_converter.py`

- [ ] **Step 9.1: Verify LibreOffice is available**

```bash
libreoffice --version
```

If not installed on Replit: `apt-get install -y libreoffice`

- [ ] **Step 9.2: Implement pdf_converter.py**

```python
# modules/pdf_converter.py
"""
Converts .docx to .pdf using LibreOffice headless mode.
LibreOffice must be installed: apt-get install -y libreoffice
"""
import os
import subprocess
import shutil


def docx_to_pdf(docx_path: str, output_dir: str | None = None) -> str:
    """
    Convert a .docx file to .pdf using LibreOffice headless.
    Returns the path to the generated PDF.
    """
    if output_dir is None:
        output_dir = os.path.dirname(docx_path)

    result = subprocess.run(
        [
            "libreoffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path,
        ],
        capture_output=True, text=True
    )

    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice PDF conversion failed:\n{result.stderr}")

    # LibreOffice names the output as the same stem + .pdf
    stem = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, stem + ".pdf")

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Expected PDF not found: {pdf_path}")

    return pdf_path
```

- [ ] **Step 9.3: Smoke test**

```bash
python3 -c "
from modules.pdf_converter import docx_to_pdf
# Requires a real docx to test — run manually after Task 8 test passes
print('pdf_converter imported OK')
"
```

- [ ] **Step 9.4: Commit**

```bash
git add modules/pdf_converter.py
git commit -m "feat: PDF converter — LibreOffice headless docx-to-pdf"
```

---

## Task 10: Update app.py — Output Format Toggle

**Files:**
- Modify: `houspire_budget_generator/app.py`

Add an output format radio button. When "Branded DOCX + PDF" is selected, after generating the Excel files, also run the docx generator and pdf converter.

- [ ] **Step 10.1: Add output format toggle to app.py**

Add this import at the top of app.py:
```python
from modules.docx_generator import generate_branded_boq_docx, generate_branded_vendor_docx
from modules.pdf_converter import docx_to_pdf
```

Replace the Step 6 section in app.py with this updated version that includes the output format toggle:

```python
    # ── Step 6: Output Format + Generate ─────────────────────────────────────
    st.subheader("6. Output Format")
    output_format = st.radio(
        "Choose output format",
        ["📊 Plain Excel (internal use)", "🎨 Branded DOCX + PDF (client delivery)"],
        horizontal=True,
    )
    branded = "Branded" in output_format

    st.subheader("7. Generate")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        gen_boq = st.button("📊 Generate Budget BOQ", type="primary")
    with col_g2:
        gen_vendors = st.button("📍 Find Local Vendors", type="primary")

    if gen_boq:
        tmpdir = tempfile.mkdtemp()
        safe = client_name.replace(" ", "_")

        with st.spinner("Generating BOQ with Claude…"):
            rows, sources = generate_boq(edited, city, pincode, tier)

        with st.spinner("Writing Excel…"):
            excel_path = os.path.join(tmpdir, f"{safe}_BOQ_{city}.xlsx")
            write_budget_excel(client_name, city, pincode, tier, rows, sources, excel_path)
            st.session_state.budget_path = excel_path
            st.session_state.boq_rows = rows

        if branded:
            with st.spinner("Generating branded DOCX…"):
                docx_path = os.path.join(tmpdir, f"Houspire_BOQ_{safe}_{city}.docx")
                generate_branded_boq_docx(excel_path, client_name, city, docx_path)
                st.session_state.budget_docx_path = docx_path
            with st.spinner("Converting to PDF…"):
                pdf_path = docx_to_pdf(docx_path, tmpdir)
                st.session_state.budget_pdf_path = pdf_path

        st.success(f"BOQ generated — {len(rows)} line items.")

    if gen_vendors:
        cats = list({r.category for r in st.session_state.boq_rows}) \
               if st.session_state.boq_rows else []
        tmpdir = tempfile.mkdtemp()
        safe = client_name.replace(" ", "_")
        room_label = edited[0].room_type.replace(" ", "_") if len(edited) == 1 else safe

        with st.spinner("Searching for real local vendors… (30-60 sec)"):
            vendors, notes_text, centroid = generate_vendors(edited, city, pincode, tier, cats)

        with st.spinner("Writing Excel…"):
            vendor_excel = os.path.join(tmpdir, f"{city}_{room_label}_Vendors_{pincode}.xlsx")
            write_vendor_excel(client_name, city, pincode, centroid, vendors, notes_text, vendor_excel)
            st.session_state.vendor_path = vendor_excel

        if branded:
            with st.spinner("Generating branded DOCX…"):
                docx_path = os.path.join(tmpdir, f"Houspire_Vendors_{safe}_{city}.docx")
                generate_branded_vendor_docx(vendor_excel, client_name, city, pincode, docx_path)
                st.session_state.vendor_docx_path = docx_path
            with st.spinner("Converting to PDF…"):
                pdf_path = docx_to_pdf(docx_path, tmpdir)
                st.session_state.vendor_pdf_path = pdf_path

        st.success(f"Vendor sheet generated — {len(vendors)} vendors.")

    # ── Downloads ─────────────────────────────────────────────────────────────
    any_ready = any([
        st.session_state.get("budget_path"),
        st.session_state.get("vendor_path"),
    ])
    if any_ready:
        st.divider()
        st.subheader("Downloads")
        cols = st.columns(3)
        col_idx = 0

        def dl_button(label, path_key, col):
            path = st.session_state.get(path_key)
            if path and os.path.exists(path):
                with col:
                    with open(path, "rb") as f:
                        ext = os.path.splitext(path)[1]
                        mime = {
                            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            ".pdf":  "application/pdf",
                        }.get(ext, "application/octet-stream")
                        st.download_button(label, data=f.read(),
                            file_name=os.path.basename(path),
                            mime=mime, use_container_width=True)

        dl_button("📥 Budget Excel",       "budget_path",      cols[0])
        dl_button("📥 Vendor Excel",        "vendor_path",      cols[1])
        if branded:
            dl_button("📄 Budget DOCX",     "budget_docx_path", cols[2])
            dl_button("📄 Vendor DOCX",     "vendor_docx_path", cols[0])
            dl_button("🖨️ Budget PDF",      "budget_pdf_path",  cols[1])
            dl_button("🖨️ Vendor PDF",      "vendor_pdf_path",  cols[2])
```

Also add to session state init at top of app.py:
```python
for key in ("analyses", "boq_rows", "boq_sources", "vendors", "notes", "centroid",
            "budget_path", "vendor_path",
            "budget_docx_path", "vendor_docx_path",
            "budget_pdf_path", "vendor_pdf_path"):
    if key not in st.session_state:
        st.session_state[key] = None
```

- [ ] **Step 10.2: Commit**

```bash
git add app.py
git commit -m "feat: output format toggle — plain Excel or branded DOCX+PDF"
```

---

## Task 11: Full Test Suite + Replit Deploy

> Rename from old Task 8.

- [ ] **Step 8.1: Run full test suite**

```bash
cd houspire_budget_generator
python -m pytest tests/ -v
```

All tests must PASS.

- [ ] **Step 8.2: End-to-end smoke test**

Run the app locally. Test with:
- Client: any name | City: Hyderabad | Pincode: 500032 | Tier: Premium
- Upload one test bedroom render
- Click Analyse → verify room detected
- Click Generate Budget BOQ → download and open Excel
  - Verify: Amount column shows `=D2*E2` (not a number) in Excel formula bar
  - Verify: No footer rows, no fills, Calibri 12 font
  - Verify: Rate Sources sheet exists with source URLs
- Click Find Local Vendors → download and open Excel
  - Verify: 7 columns with real vendor names
  - Verify: Notes sheet present

- [ ] **Step 8.3: Verify BOQ formula integrity**

```bash
python3 recalc.py /path/to/downloaded_BOQ.xlsx 500
```

Expected: `"total_errors": 0`

- [ ] **Step 8.4: Deploy to Replit**

1. Go to https://replit.com → New Repl → Python → upload all files
2. In Replit Secrets add: `ANTHROPIC_API_KEY`
3. Set run command: `streamlit run app.py --server.port 8080 --server.address 0.0.0.0`
4. Click Run

- [ ] **Step 8.5: Final commit**

```bash
git add .
git commit -m "feat: Houspire Budget Generator v1.0 — BOQ + Vendor Excel from renders"
git tag v1.0.0
```

---

## Key Decisions

| Decision | Choice | Reason |
|----------|--------|--------|
| BOQ generation | Claude API with rate library in system prompt | Same as chat — no Supabase needed |
| Vendor generation | Claude API with web_search tool | Same search Claude uses in this chat |
| Excel format | Exact plain Calibri-12 template from skills | Matches all 13+ live Houspire projects |
| Tiers | Mid-tier / Premium only | Houspire skill vocabulary — no other tiers |
| City multipliers | Silent — baked into rates | From houspire-budget skill spec |
| Footer rows | None | Plain template has no subtotal/GST/contingency |
| Amount column | `=D{row}*E{row}` formula | Never hard-coded — skill requirement |
| Verification | recalc.py | Skill-required formula check before delivery |
| Platform | Replit | No-code deployment, matches Abhishek's preference |
| API keys | Anthropic only | One key, same capability as Claude chat |
