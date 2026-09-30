# towerco_resolver.py
"""
Resolve the final Site Owner / Towerco value by comparing the
Excel masterlist (Column N: TOWERCO) against the Ericsson TSSR
(Site Owner checkbox on Page 3).

Decision rules (as specified by the user):
    1. If Excel and TSSR match                -> no action, follow TSSR.
    2. If Excel = GT and TSSR = Towerco       -> notify verify,
                                                 still follow TSSR Towerco.
    3. If both are Towercos but different     -> CALL ROC.
    4. Any other mismatch                     -> notify verify,
                                                 follow TSSR.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


# =============================================================
# Tokens
# =============================================================

GT_TOKENS = {"GT", "GLOBE", "GLOBE TELECOM", "GLOBE TELECOMS"}

TOWERCO_TOKENS = {
    "TCO", "TOWERCO", "TOWER CO", "TOWER COMPANY",
    "FTAP", "PHILTOWER", "PHIL TOWER",
    "EDGEPOINT", "EDGE POINT",
    "SBA", "SBA TOWERS",
    "MIESCOR", "MIES COR",
    "ISOC", "ISOC EDGE",
    "ALTUS", "ALTUS TOWER",
    "ABOITIZ", "ABOITIZ TOWER",
    "CREATIVE TOWER", "CREATIVE TOWERS",
    "LCS", "LCS TOWER",
    "TEC TOWER",
}


class Action(Enum):
    MATCH           = "match"             # no action needed
    VERIFY_FOLLOW   = "verify_follow"     # notify user, follow TSSR
    CALL_ROC        = "call_roc"          # instruct user to call ROC


@dataclass
class TowercoResolution:
    excel_value:   str
    tssr_value:    str
    excel_class:   str          # "GT" | "TOWERCO" | "OTHER"
    tssr_class:    str          # "GT" | "TOWERCO" | "OTHER"
    action:        Action
    final_owner:   str          # value to write into site_owner
    message:       str          # human-readable notification
    needs_review:  bool = False


# =============================================================
# Classification helpers
# =============================================================

def classify_towerco(value: str) -> str:
    """Return 'GT', 'TOWERCO', or 'OTHER' for a given raw string."""
    v = (value or "").strip().upper()
    if not v:
        return "OTHER"
    if v in GT_TOKENS:
        return "GT"
    if v in TOWERCO_TOKENS:
        return "TOWERCO"
    # Substring fallback
    for tok in GT_TOKENS:
        if tok in v:
            return "GT"
    for tok in TOWERCO_TOKENS:
        if tok in v:
            return "TOWERCO"
    return "OTHER"


# =============================================================
# Resolver
# =============================================================

def resolve_site_owner(excel_value: str, tssr_value: str) -> TowercoResolution:
    """
    Apply the decision tree and return a TowercoResolution.

    excel_value : raw value from Excel Column N (TOWERCO)
    tssr_value  : raw value from TSSR Page 3 Site Owner
                  (expected "TCO" when the Tower CO box is checked)
    """
    excel_class = classify_towerco(excel_value)
    tssr_class  = classify_towerco(tssr_value)

    excel_norm = (excel_value or "").strip().upper()
    tssr_norm  = (tssr_value or "").strip().upper()

    # --- Rule 1: exact match ------------------------------------
    if excel_norm == tssr_norm and excel_norm != "":
        return TowercoResolution(
            excel_value=excel_value,
            tssr_value=tssr_value,
            excel_class=excel_class,
            tssr_class=tssr_class,
            action=Action.MATCH,
            final_owner=tssr_value or excel_value,
            message="Excel and TSSR agree. No action needed.",
            needs_review=False,
        )

    # --- Rule 3: both Towercos but different --------------------
    if excel_class == "TOWERCO" and tssr_class == "TOWERCO":
        return TowercoResolution(
            excel_value=excel_value,
            tssr_value=tssr_value,
            excel_class=excel_class,
            tssr_class=tssr_class,
            action=Action.CALL_ROC,
            final_owner=tssr_value,   # still follow TSSR until ROC overrides
            message=(
                f"Both sources are Towercos but different: "
                f"Excel='{excel_value}' vs TSSR='{tssr_value}'. "
                f"PLEASE CALL ROC."
            ),
            needs_review=True,
        )

    # --- Rule 2: Excel = GT, TSSR = Towerco ---------------------
    if excel_class == "GT" and tssr_class == "TOWERCO":
        return TowercoResolution(
            excel_value=excel_value,
            tssr_value=tssr_value,
            excel_class=excel_class,
            tssr_class=tssr_class,
            action=Action.VERIFY_FOLLOW,
            final_owner=tssr_value,   # follow TSSR Towerco
            message=(
                f"Excel says GT but TSSR says Towerco "
                f"(Excel='{excel_value}' vs TSSR='{tssr_value}'). "
                f"Please verify. Following TSSR Towerco."
            ),
            needs_review=True,
        )

    # --- Rule 4: any other mismatch -----------------------------
    return TowercoResolution(
        excel_value=excel_value,
        tssr_value=tssr_value,
        excel_class=excel_class,
        tssr_class=tssr_class,
        action=Action.VERIFY_FOLLOW,
        final_owner=tssr_value or excel_value,
        message=(
            f"Mismatch: Excel='{excel_value}' vs TSSR='{tssr_value}'. "
            f"Please verify. Following TSSR."
        ),
        needs_review=True,
    )


# =============================================================
# CLI test
# =============================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python towerco_resolver.py <excel_value> <tssr_value>")
        sys.exit(1)

    res = resolve_site_owner(sys.argv[1], sys.argv[2])
    print(f"Excel      : {res.excel_value!r}  ({res.excel_class})")
    print(f"TSSR       : {res.tssr_value!r}  ({res.tssr_class})")
    print(f"Action     : {res.action.value}")
    print(f"Final owner: {res.final_owner!r}")
    print(f"Message    : {res.message}")
    print(f"Review?    : {res.needs_review}")
