import sys, os, tempfile, subprocess
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import pytest
from modules.docx_generator import _node_script_boq, _node_script_vendor


def _node_available():
    try:
        r = subprocess.run(["node", "--version"], capture_output=True)
        return r.returncode == 0
    except FileNotFoundError:
        return False


def test_boq_script_is_string():
    data = {
        "client": "Test", "city": "Hyderabad", "tier": "Mid-tier",
        "rows": [{"category": "Flooring", "description": "Tile", "unit": "sft", "qty": 100, "rate": 179}],
        "sources": [],
        "output_path": "/tmp/test.docx",
    }
    script = _node_script_boq(data)
    assert isinstance(script, str)
    assert 'require("docx")' in script
    assert "Packer.toBuffer" in script


def test_vendor_script_is_string():
    data = {
        "client": "Test", "city": "Hyderabad",
        "vendors": [{"category": "Flooring", "vendor": "Kajaria", "specialty": "Tiles",
                     "area": "Banjara Hills", "dist": "2.1", "rating": "4.7 (100)", "phone": "+91 99999 00000"}],
        "output_path": "/tmp/test_v.docx",
    }
    script = _node_script_vendor(data)
    assert isinstance(script, str)
    assert 'require("docx")' in script


def test_boq_script_no_calibri():
    data = {
        "client": "Test", "city": "Hyderabad", "tier": "Premium",
        "rows": [], "sources": [], "output_path": "/tmp/x.docx",
    }
    script = _node_script_boq(data)
    assert "Calibri" not in script, "Branded DOCX must use Arial, not Calibri"


@pytest.mark.skipif(not _node_available(), reason="node not available")
def test_boq_docx_written():
    """Integration test — requires node + docx npm installed."""
    from modules.budget_excel import write_budget_excel
    from modules.boq_generator import BOQRow, RateSource
    from modules.docx_generator import generate_branded_boq_docx

    rows = [BOQRow("Flooring", "Kajaria vitrified - Living Room", "sft", 200.0, 179.0)]
    sources = [RateSource("Flooring", "Kajaria 2025", "kajariaceramics.com")]

    with tempfile.TemporaryDirectory() as d:
        xlsx = os.path.join(d, "boq.xlsx")
        docx = os.path.join(d, "out.docx")
        write_budget_excel("Test", "Hyderabad", "500032", "Mid-tier", rows, sources, xlsx)
        generate_branded_boq_docx("Test", "Hyderabad", "Mid-tier", xlsx, docx)
        assert os.path.exists(docx)
        assert os.path.getsize(docx) > 5000
