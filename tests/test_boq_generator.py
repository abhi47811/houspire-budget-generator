import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from modules.boq_generator import _parse_boq_response, BOQRow, RateSource


SAMPLE_RESPONSE = """=== BOQ TABLE ===
Category,Description,Unit,Quantity,Rate,Amount
Flooring,Premium vitrified Kajaria - Living Room,sft,200,179,35800
Electrical,6A switch Legrand Arteor - Living Room,nos,4,1150,4600
HVAC,Split AC 1.5T Daikin - Bedroom,nos,1,51000,51000
HVAC,Split AC 1.5T install kit - Bedroom,lump,1,11000,11000

=== RATE SOURCES ===
Category / Item,Rate Basis (Hyderabad Mid-tier 2025-26),Source
Flooring,Kajaria 2025-26,kajariaceramics.com
Note,Hyderabad multiplier applied silently.,-
"""


def test_parse_boq_rows():
    rows, sources = _parse_boq_response(SAMPLE_RESPONSE)
    assert len(rows) == 4
    assert all(isinstance(r, BOQRow) for r in rows)


def test_parse_skips_header():
    rows, _ = _parse_boq_response(SAMPLE_RESPONSE)
    assert rows[0].category == "Flooring"


def test_parse_sources():
    _, sources = _parse_boq_response(SAMPLE_RESPONSE)
    assert len(sources) >= 1
    assert all(isinstance(s, RateSource) for s in sources)


def test_qty_rate_are_floats():
    rows, _ = _parse_boq_response(SAMPLE_RESPONSE)
    for r in rows:
        assert isinstance(r.qty, float)
        assert isinstance(r.rate, float)
