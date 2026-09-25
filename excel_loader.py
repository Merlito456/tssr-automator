# excel_loader.py
"""
Load and query the MINDANAO site masterlist Excel.
Self-diagnosing — prints what it finds if the default names don't match.
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Optional
import pandas as pd


# ─── CONFIG — we'll auto-detect if these don't match ─────────
DEFAULT_SHEET = "GLOBE SITE MASTERLIST"
HEADER_ROW = 0


def _normalize_col(name: str) -> str:
    """Strip + collapse whitespace. 'WIRELINE_NAME ' → 'WIRELINE_NAME'."""
    return re.sub(r"\s+", " ", str(name).strip())


def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
    """Find sheet by exact match, then case-insensitive, then first."""
    if target in xl.sheet_names:
        return target

    # Case-insensitive
    for s in xl.sheet_names:
        if s.strip().upper() == target.strip().upper():
            return s

    # Not found — raise with helpful message
    raise ValueError(
        f"Sheet '{target}' not found. Available sheets: {xl.sheet_names}"
    )


class SiteMasterlist:
    """Wrapper around the Excel masterlist."""

    def __init__(self, xlsx_path: str, sheet_name: str = DEFAULT_SHEET):
        self.path = Path(xlsx_path)
        if not self.path.exists():
            raise FileNotFoundError(f"Excel not found: {xlsx_path}")

        # Discover actual sheet name
        xl = pd.ExcelFile(self.path)
        actual_sheet = _find_sheet(xl, sheet_name)
        self.sheet_name = actual_sheet

        # Load dataframe
        self.df = pd.read_excel(
            self.path,
            sheet_name=actual_sheet,
            header=HEADER_ROW,
            dtype=str,
        ).fillna("")

        # Normalize headers — strip trailing/leading spaces
        self.df.columns = [_normalize_col(c) for c in self.df.columns]

        # Diagnostic dump (only shows on first run)
        print(f"✅ Loaded sheet: {actual_sheet}")
        print(f"   Rows: {len(self.df)}")
        print(f"   Columns ({len(self.df.columns)}): {list(self.df.columns)}")

        # Normalize PLAID
        plaid_col = self._find_column("PLAID")
        if plaid_col:
            self.df[plaid_col] = (
                self.df[plaid_col].astype(str).str.strip().str.upper()
            )

    # ─────────────────────────────────────────────────────────
    # Column resolution — flexible to naming variations
    # ─────────────────────────────────────────────────────────

    def _find_column(self, target: str) -> Optional[str]:
        """
        Find a column by exact match, then by normalized comparison.
        Handles trailing spaces, case differences.
        """
        target_norm = _normalize_col(target).upper()

        # Exact
        if target in self.df.columns:
            return target

        # Normalized match
        for c in self.df.columns:
            if _normalize_col(c).upper() == target_norm:
                return c

        # Fallback — substring match
        for c in self.df.columns:
            if target_norm in _normalize_col(c).upper():
                return c

        return None

    # ─────────────────────────────────────────────────────────
    # Public API
    # ─────────────────────────────────────────────────────────

    def get_site(self, plaid: str) -> Optional[dict]:
        plaid = str(plaid).strip().upper()
        plaid_col = self._find_column("PLAID")
        if not plaid_col:
            return None

        rows = self.df[self.df[plaid_col] == plaid]
        if rows.empty:
            return None
        row = rows.iloc[0]

        def val(field_name: str) -> str:
            col = self._find_column(field_name)
            if not col:
                return ""
            return str(row.get(col, "")).strip()

        # Site address: Barangay + Municipality + Province
        address_parts = [
            val("BARANGAY"),
            val("MUNICIPALITY"),
            val("PROVINCE"),
        ]
        site_address = ", ".join(p for p in address_parts if p)

        # Coordinates
        lat, lon = val("LATITUDE"), val("LONGITUDE")
        try:
            coords = f"{float(lat):.5f}, {float(lon):.5f}" if lat and lon else ""
        except ValueError:
            coords = f"{lat}, {lon}".strip(", ")

        # FO = column S ("NEW ENGINEER_ANM1")
        # FO number = column U ("CONTACT NUMBER")
        fo_name   = val("NEW ENGINEER_ANM1")
        fo_mobile = val("CONTACT NUMBER")

        return {
            "site_id":   val("PLAID"),
            "site_name": val("SITE"),
            "region":    val("REGION"),
            "province":  val("PROVINCE"),
            "municipality": val("MUNICIPALITY"),
            "barangay":  val("BARANGAY"),
            "latitude":  lat,
            "longitude": lon,
            "site_add":  site_address,
            "site_coords": coords,
            "towerco":   val("TOWERCO"),
            "assign_hub": val("ASSIGN_HUB"),
            "site_key_location": val("ASSIGN_HUB") or val("NEW ASSIGN_HUB"),
            "fo_name":      fo_name,
            "fo_mobile":    fo_mobile,
            "site_contact_person": fo_name,
            "site_contact_mobile": fo_mobile,
            "territory": val("TERRITORY"),
        }

    def list_plaids(self) -> list[str]:
        plaid_col = self._find_column("PLAID")
        if not plaid_col:
            return []
        return sorted(
            self.df[plaid_col].dropna().astype(str)
              .str.strip().str.upper()
              .unique().tolist()
        )
