import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from modules.room_analyzer import RoomAnalysis
from config import ROOM_TYPES


def test_room_analysis_fields():
    a = RoomAnalysis(
        room_type="Living Room",
        estimated_sqft=250,
        confidence="high",
        design_elements="marble floor, cove ceiling",
        image_filename="render1.jpg",
    )
    assert a.room_type in ROOM_TYPES
    assert isinstance(a.estimated_sqft, int)
    assert a.confidence in ("high", "medium", "low")
    assert len(a.design_elements) > 0


def test_default_filename():
    a = RoomAnalysis("Unknown", 120, "low", "could not analyse")
    assert a.image_filename == ""
