import os, re
from dataclasses import dataclass, field
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


@dataclass
class VendorRow:
    category: str
    vendor: str
    specialty: str
    area: str
    lat: float = 0.0
    lng: float = 0.0
    rating: str = ""
    phone: str = ""


def _parse_vendor_response(text):
    vendors = []
    notes = ""
    tbl_m = re.search(r"=== VENDOR TABLE ===\n(.*?)\n(?:=== NOTES ===|$)", text, re.DOTALL)
    notes_m = re.search(r"=== NOTES ===\n(.*?)$", text, re.DOTALL)
    if tbl_m:
        for line in tbl_m.group(1).strip().splitlines():
            p = [x.strip() for x in line.split("|")]
            if len(p) < 7 or p[0] == "Category":
                continue
            try:
                vendors.append(VendorRow(
                    category=p[0], vendor=p[1], specialty=p[2], area=p[3],
                    lat=float(p[4]) if p[4] else 0.0,
                    lng=float(p[5]) if p[5] else 0.0,
                    rating=p[6] if len(p) > 6 else "",
                    phone=p[7] if len(p) > 7 else "",
                ))
            except (ValueError, IndexError):
                continue
    if notes_m:
        notes = notes_m.group(1).strip()
    return vendors, notes


def generate_vendors(rooms, city, pincode, tier, categories=None):
    if categories is None:
        categories = []
    resp = _client().messages.create(
        model="claude-sonnet-4-5",
        max_tokens=8192,
        system=VENDOR_SYSTEM_PROMPT,
        tools=[{"type": "web_search_20250305", "name": "web_search"}],
        messages=[{"role": "user",
                   "content": _build_vendor_prompt(rooms, city, pincode, tier, categories)}],
    )
    final_text = next((b.text for b in resp.content if hasattr(b, "text")), "")
    return _parse_vendor_response(final_text)
