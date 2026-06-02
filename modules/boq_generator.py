import os, re
from dataclasses import dataclass
import anthropic
from dotenv import load_dotenv

load_dotenv()

def _client():
    try:
        import streamlit as st
        key = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        key = os.environ.get("ANTHROPIC_API_KEY", "")
    return anthropic.Anthropic(api_key=key)

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


def _parse_boq_response(text):
    rows, sources = [], []
    boq_m = re.search(r"=== BOQ TABLE ===\n(.*?)\n(?:=== RATE SOURCES ===|$)", text, re.DOTALL)
    src_m = re.search(r"=== RATE SOURCES ===\n(.*?)$", text, re.DOTALL)
    if boq_m:
        for line in boq_m.group(1).strip().splitlines():
            parts = [x.strip() for x in line.split(",")]
            if len(parts) < 5 or parts[0] == "Category":
                continue
            try:
                # description may contain commas — anchor from both ends
                category = parts[0]
                unit     = parts[-4] if len(parts) >= 5 else parts[2]
                qty      = parts[-3]
                rate     = parts[-2]
                # amount = parts[-1]  # ignored — formula in Excel
                description = ",".join(parts[1:-4]).strip() if len(parts) > 5 else parts[1]
                rows.append(BOQRow(category, description, unit, float(qty), float(rate)))
            except (ValueError, IndexError):
                continue
    if src_m:
        for line in src_m.group(1).strip().splitlines():
            p = [x.strip() for x in line.split(",", 2)]
            if len(p) < 2 or not p[0]:
                continue
            sources.append(RateSource(p[0], p[1] if len(p) > 1 else "", p[2] if len(p) > 2 else ""))
    return rows, sources


def generate_boq(rooms, city, pincode, tier):
    resp = _client().messages.create(
        model="claude-sonnet-4-5",
        max_tokens=4096,
        system=BOQ_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": _build_user_prompt(rooms, city, pincode, tier)}],
    )
    raw = resp.content[0].text
    rows, sources = _parse_boq_response(raw)
    return rows, sources, raw  # raw exposed for debug when rows == 0
