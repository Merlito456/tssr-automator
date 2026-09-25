# permit_rules.py
"""
Hardcoded permit rules per tower company.
"""

from __future__ import annotations


TOWERCO_PERMITS = {
    "PHILTOWER": {
        "work_permit":        "RAAWA, HSWP and iTower Account",
        "access_requirement": "RAAWA, LILO, Approved Ticket",
    },
    "FTAP": {
        "work_permit":        "RAAWA, iAMS account",
        "access_requirement": "RAAWA, iAMS, Approved Ticket",
    },
    "EDOTCO": {
        "work_permit":        "RAAWA, Approved TAP",
        "access_requirement": "RAAWA, Approved TAP",
    },
}


def get_permits_for_towerco(towerco: str) -> dict:
    """Return permit strings for the given towercо (Latin-only identifier)."""
    if not towerco:
        return {"work_permit": "", "access_requirement": ""}

    tc_upper = str(towerco).strip().upper()

    if tc_upper in TOWERCO_PERMITS:
        return dict(TOWERCO_PERMITS[tc_upper])

    for key, value in TOWERCO_PERMITS.items():
        if key in tc_upper:
            return dict(value)

    return {"work_permit": "", "access_requirement": ""}
