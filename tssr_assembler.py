# tssr_assembler.py
"""
Assemble the final site dict by merging:
  - Excel masterlist row  (via SiteMasterlist)
  - Raw TSSR JSON         (produced by the TSSR extractor)
  - Towerco resolution    (via towerco_resolver)

Usage:
    ml  = SiteMasterlist("path/to/masterlist.xlsx")
    raw = json.load(open("min447_tssr.json"))
    final = assemble_site("MIN447", ml, raw)
"""

from __future__ import annotations

import json
from typing import Optional

from excel_loader import SiteMasterlist
from towerco_resolver import resolve_site_owner, TowercoResolution


def assemble_site(
    plaid: str,
    masterlist: SiteMasterlist,
    raw_tssr: dict,
    verbose: bool = True,
) -> dict:
    """
    Build the final site dict.

    raw_tssr must contain at least:
        raw_tssr["site_owner"]   -> "TCO" / "GT" / etc.
    Everything else is optional and merged when present.
    """
    plaid = str(plaid).strip().upper()

    # 1. Pull the Excel row
    excel_site = masterlist.get_site(plaid)
    if excel_site is None:
        raise ValueError(f"PLAID '{plaid}' not found in masterlist.")

    # 2. Resolve Towerco
    excel_towerco = excel_site.get("towerco", "")
    tssr_owner    = raw_tssr.get("site_owner", "")

    resolution: TowercoResolution = resolve_site_owner(excel_towerco, tssr_owner)

    if verbose:
        print(f"\n--- Towerco resolution for {plaid} ---")
        print(f"  Excel : {resolution.excel_value!r} ({resolution.excel_class})")
        print(f"  TSSR  : {resolution.tssr_value!r} ({resolution.tssr_class})")
        print(f"  Action: {resolution.action.value}")
        print(f"  Final : {resolution.final_owner!r}")
        print(f"  Msg   : {resolution.message}")

    # 3. Compose the final dict
    final = {
        # -- Identity from Excel --
        "site_id":              plaid,
        "site_name":            excel_site.get("site_name", ""),
        "region":               excel_site.get("region", ""),
        "site_address":         excel_site.get("site_address", ""),
        "site_coords":          excel_site.get("site_coords", ""),

        # -- Geography --
        "province":             excel_site.get("province", ""),
        "municipality":         excel_site.get("municipality", ""),
        "barangay":             excel_site.get("barangay", ""),

        # -- Coordinates (raw) --
        "latitude":             excel_site.get("latitude", ""),
        "longitude":            excel_site.get("longitude", ""),

        # -- Site Owner: resolved --
        "site_owner":           resolution.final_owner,
        "towerco":              resolution.final_owner,           # alias
        "towerco_excel_raw":    resolution.excel_value,
        "towerco_tssr_raw":     resolution.tssr_value,
        "towerco_action":       resolution.action.value,
        "towerco_message":      resolution.message,
        "towerco_needs_review": resolution.needs_review,

        # -- Site key location (Excel Column M, fallback Q) --
        "site_key_location":    excel_site.get("site_key_location", ""),

        # -- Field Officer --
        "fo_name":              excel_site.get("fo_name", ""),
        "fo_mobile":            excel_site.get("fo_mobile", ""),
        "site_contact_person":  excel_site.get("site_contact_person", ""),
        "site_contact_mobile":  excel_site.get("site_contact_mobile", ""),

        # -- Extra Excel fields --
        "wireline":             excel_site.get("wireline", ""),
        "bcf":                  excel_site.get("bcf", ""),
        "territory":            excel_site.get("territory", ""),
        "assign_hub":           excel_site.get("assign_hub", ""),
        "engineer_ah":          excel_site.get("engineer_ah", ""),
        "engineer_anm1":        excel_site.get("engineer_anm1", ""),
        "engineer_anm1_id":     excel_site.get("engineer_anm1_id", ""),
        "anm_head":             excel_site.get("anm_head", ""),
        "roh":                  excel_site.get("roh", ""),
    }

    # 4. Merge in raw TSSR fields (they should NOT overwrite the
    #    resolved site_owner / towerco fields)
    PROTECTED = {
        "site_owner", "towerco",
        "towerco_excel_raw", "towerco_tssr_raw",
        "towerco_action", "towerco_message", "towerco_needs_review",
    }
    for k, v in raw_tssr.items():
        if k in PROTECTED:
            continue
        # Only set if the final dict doesn't already have it
        # (Excel wins for identity fields)
        if k not in final:
            final[k] = v

    return final


# =============================================================
# CLI test
# =============================================================

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python tssr_assembler.py <masterlist.xlsx> "
              "<raw_tssr.json> <PLAID>")
        sys.exit(1)

    xlsx_path  = sys.argv[1]
    tssr_path  = sys.argv[2]
    plaid      = sys.argv[3] if len(sys.argv) > 3 else "MIN447"

    ml = SiteMasterlist(xlsx_path)

    with open(tssr_path, "r", encoding="utf-8") as f:
        raw_tssr = json.load(f)

    final = assemble_site(plaid, ml, raw_tssr)
    print("\n--- FINAL SITE DICT ---")
    print(json.dumps(final, indent=2, ensure_ascii=False))
