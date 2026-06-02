import os, json, base64
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
        resp = _client().messages.create(
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
