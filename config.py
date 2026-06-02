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

BOQ_COL_WIDTHS    = {"A": 15.83, "B": 25.83, "C": 10.83, "D": 10.83, "E": 10.83, "F": 15.83}
VENDOR_COL_WIDTHS = {"A": 40, "B": 44, "C": 54, "D": 36, "E": 22, "F": 15, "G": 19}
