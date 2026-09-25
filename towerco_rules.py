# towercо_rules.py
"""
Hardcoded permit rules per tower company.

The site's TOWERCO (from Excel column N) determines:
  - work_permit       → "RAAWA" + towerco-specific requirement
  - access_requirement → the standard access template for that towerco
"""

from __future__ import annotations


# Map towercо (uppercase key match) → permit composition
TOWERCO_PERMITS = {
    "PHILTOWER": {
        "work_permit": "RAAWA, HSWP and iTower Account",
        "access_requirement": "RAAWA, LILO, Approved Ticket",
    },
    "FTAP": {
        "work_permit": "RAAWA, iAMS account",
        "access_requirement": "RAAWA, iAMS, Approved Ticket",
    },
    "EDOTCO": {
        "work_permit": "RAAWA, Approved TAP",
        "access_requirement": "RAAWA, Approved TAP",
    },
    # Add more tower companies as needed
}


def get_permits_for_towerco(towercо: str) -> dict:
    """
    Return {"work_permit": ..., "access_requirement": ...} for the given towercо.
    Falls back to generic RAAWA if towercо is unknown but contains "RAAWA".

    If the towercо doesn't match anything, returns empty strings.
    """
    if not towercо:
        return {"work_permit": "", "access_requirement": ""}

    tc_upper = str(towercо).strip().upper()

    # Exact key match
    if tc_upper in TOWERCO_PERMITS:
        return dict(TOWERCO_PERMITS[tc_upper])

    # Substring match (e.g. "PHILTOWER INC" matches "PHILTOWER")
    for key, value in TOWERCO_PERMITS.items():
        if key in tc_upper:
            return dict(value)

    # Unknown towercо — return empty
    return {"work_permit": "", "access_requirement": ""}
