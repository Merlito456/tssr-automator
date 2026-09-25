# excel_loader.py
"""
Load and query the MINDANAO site masterlist Excel.
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Optional
import pandas as pd


SHEET_NAME = "GLOBE SITE MASTERLIST"
HEADER_ROW = 0  # 0-indexed; Excel row 1


# Canonical column names (post-normalization)
COLUMNS = {
    "plaid":          "PLAID",
    "site":           "SITE",
    "wireline":       "WIRELINE_NAME",
    "bcf":            "BCF_NAME",
    "region":         "REGION",
    "province":       "PROVINCE",
    "municipality":   "MUNICIPALITY",
    "barangay":       "BARANGAY",
    "territory":      "TERRITORY",
    "latitude":       "LATITUDE",
    "longitude":      "LONGITUDE",
    "site_add":       "SITE_ADD",
    "assign_hub":     "ASSIGN_HUB",
    "towerco":        "TOWERCO",
    "new_area":       "NEW ASSIGN_AREA",
    "new_area_name":  "NEW ASSIGN_AREA NAME",
    "new_hub":        "NEW ASSIGN_HUB",
    "engineer_ah":    "NEW ENGINEER_AH",
    "engineer_anm1":  "NEW ENGINEER_ANM1",
    "engineer_anm1_id": "NEW ENGINEER_ANM1 ID NUMBER",
    "contact_number": "CONTACT NUMBER",
    "anm_head":       "NEW ANM HEAD",
    "roh":            "NEW ROH",
}


class SiteMasterlist:
    """Wrapper around the Excel masterlist."""

    def __init__(self, xlsx_path: str):
        self.path = Path(xlsx_path)
        if not self.path.exists():
            raise FileNotFoundError(f"Excel not found: {xlsx_path}")

        self.df = pd.read_excel(
            self.path,
            sheet_name=SHEET_NAME,
            header=HEADER_ROW,
            dtype=str,        # keep everything as string, convert later
        ).fillna("")

        # Normalize column headers — strip and collapse whitespace
        self.df.columns = [
            re.sub(r"\s+", " ", str(c).strip()) for c in self.df.columns
        ]

        # Normalize PLAID values
        plaid_col = COLUMNS["plaid"]
        self.df[plaid_col] = self.df[plaid_col].str.strip().str.upper()

    def get_site(self, plaid: str) -> Optional[dict]:
        """Return a dict of site data for the given PLAID, or None."""
        plaid = plaid.strip().upper()
        rows = self.df[self.df[COLUMNS["plaid"]] == plaid]
        if rows.empty:
            return None

        row = rows.iloc[0]

        def val(key: str) -> str:
            col = COLUMNS.get(key)
            if not col or col not in self.df.columns:
                return ""
            return str(row.get(col, "")).strip()

        return {
            "site_id":       val("plaid"),
            "site_name":     val("site"),
            "region":        val("region"),
            "province":      val("province"),
            "municipality":  val("municipality"),
            "barangay":      val("barangay"),
            "latitude":      val("latitude"),
            "longitude":     val("longitude"),
            "site_add":      val("site_add"),
            "towerco":       val("towerco"),
            "assign_hub":    val("assign_hub"),
            "site_key_location": val("assign_hub") or val("new_hub"),
            "fo_name":       val("engineer_anm1"),
            "fo_mobile":     val("contact_number"),
            "site_contact_person": val("engineer_anm1"),
            "site_contact_mobile": val("contact_number"),
            "territory":     val("territory"),
        }

    def list_plaids(self) -> list[str]:
        """Return sorted list of all PLAIDs — useful for autocomplete."""
        return sorted(self.df[COLUMNS["plaid"]].dropna().unique().tolist())


def compose_address(site: dict) -> str:
    """Barangay, Municipality, Province."""
    parts = [
        site.get("barangay", "").strip(),
        site.get("municipality", "").strip(),
        site.get("province", "").strip(),
    ]
    return ", ".join(p for p in parts if p)


def compose_coords(site: dict) -> str:
    """'6.63979, 124.06598' — 5 decimal places."""
    lat = site.get("latitude", "").strip()
    lon = site.get("longitude", "").strip()
    if not lat or not lon:
        return ""
    try:
        return f"{float(lat):.5f}, {float(lon):.5f}"
    except ValueError:
        return f"{lat}, {lon}"  # fall back to raw strings
