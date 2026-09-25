# excel_loader.py
"""
Load and query the MINDANAO site masterlist Excel.

Sheet: GLOBE SITE MASTERLIST
Header row: 1 (0-indexed row 0)

Columns (A–W):
    A  PLAID
    B  SITE
    C  WIRELINE_NAME
    D  BCF_NAME
    E  REGION
    F  PROVINCE
    G  MUNICIPALITY
    H  BARANGAY
    I  TERRITORY
    J  LATITUDE
    K  LONGITUDE
    L  SITE_ADD
    M  ASSIGN_HUB                ← site_key_location source
    N  TOWERCO
    O  NEW ASSIGN_AREA
    P  NEW ASSIGN_AREA NAME
    Q  NEW ASSIGN_HUB            ← fallback for site_key_location
    R  NEW ENGINEER_AH
    S  NEW ENGINEER_ANM1         ← FO NAME
    T  NEW ENGINEER_ANM1 ID NUMBER
    U  CONTACT NUMBER            ← FO NUMBER
    V  NEW ANM HEAD
    W  NEW ROH

Field Officer rules:
    FO NAME   = column S ("NEW ENGINEER_ANM1")
    FO NUMBER = column U ("CONTACT NUMBER")

Site key location rules:
    PRIMARY   = column M ("ASSIGN_HUB")
    FALLBACK  = column Q ("NEW ASSIGN_HUB")
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pandas as pd


# ═════════════════════════════════════════════════════════════
# Configuration
# ═════════════════════════════════════════════════════════════

DEFAULT_SHEET = "GLOBE SITE MASTERLIST"
HEADER_ROW = 0  # 0-indexed; Excel row 1


# Canonical field → list of likely column header names (aliases)
COLUMN_ALIASES = {
    "plaid":             ["PLAID", "PLA ID", "SITE ID", "SITE_ID"],
    "site":              ["SITE", "SITE NAME", "SITE_NAME"],
    "wireline":          ["WIRELINE_NAME", "WIRELINE NAME"],
    "bcf":               ["BCF_NAME", "BCF NAME"],
    "region":            ["REGION"],
    "province":          ["PROVINCE"],
    "municipality":      ["MUNICIPALITY", "CITY/MUNICIPALITY", "CITY"],
    "barangay":          ["BARANGAY"],
    "territory":         ["TERRITORY"],
    "latitude":          ["LATITUDE", "LAT"],
    "longitude":         ["LONGITUDE", "LONG", "LON"],
    "site_add":          ["SITE_ADD", "SITE ADDRESS", "ADDRESS"],
    "assign_hub":        ["ASSIGN_HUB", "ASSIGN HUB"],
    "towerco":           ["TOWERCO", "TOWER CO", "TOWER_CO"],
    "new_area":          ["NEW ASSIGN_AREA", "NEW ASSIGN AREA"],
    "new_area_name":     ["NEW ASSIGN_AREA NAME", "NEW ASSIGN AREA NAME"],
    "new_hub":           ["NEW ASSIGN_HUB", "NEW ASSIGN HUB"],
    "engineer_ah":       ["NEW ENGINEER_AH", "NEW ENGINEER AH"],
    "engineer_anm1":     ["NEW ENGINEER_ANM1", "NEW ENGINEER ANM1"],
    "engineer_anm1_id":  ["NEW ENGINEER_ANM1 ID NUMBER", "NEW ENGINEER ANM1 ID NUMBER"],
    "contact_number":    ["CONTACT NUMBER", "CONTACT NO", "CONTACT_NO"],
    "anm_head":          ["NEW ANM HEAD", "NEW ANM_HEAD"],
    "roh":               ["NEW ROH", "NEW_ROH"],
}


# ═════════════════════════════════════════════════════════════
# Helpers
# ═════════════════════════════════════════════════════════════

def _normalize_col(name) -> str:
    """Strip + collapse internal whitespace. 'WIRELINE_NAME ' → 'WIRELINE_NAME'."""
    return re.sub(r"\s+", " ", str(name).strip())


def _norm_upper(name) -> str:
    """Normalize for comparison: uppercase, single spaces."""
    return _normalize_col(name).upper()


# ═════════════════════════════════════════════════════════════
# Main class
# ═════════════════════════════════════════════════════════════

class SiteMasterlist:
    """Wrapper around the MINDANAO site masterlist Excel."""

    def __init__(self, xlsx_path: str, sheet_name: str = DEFAULT_SHEET):
        self.path = Path(xlsx_path)
        if not self.path.exists():
            raise FileNotFoundError(f"Excel not found: {xlsx_path}")

        # Discover actual sheet name
        xl = pd.ExcelFile(self.path)
        self.sheet_name = self._find_sheet(xl, sheet_name)

        # Load dataframe — keep everything as string
        self.df = pd.read_excel(
            self.path,
            sheet_name=self.sheet_name,
            header=HEADER_ROW,
            dtype=str,
        ).fillna("")

        # Normalize column headers
        self.df.columns = [_normalize_col(c) for c in self.df.columns]

        # Build canonical-column lookup
        self._col_map: dict[str, str] = {}
        for canon, aliases in COLUMN_ALIASES.items():
            found = self._find_column(aliases)
            if found:
                self._col_map[canon] = found

        # Normalize PLAID values
        plaid_col = self._col_map.get("plaid")
        if plaid_col:
            self.df[plaid_col] = (
                self.df[plaid_col]
                .astype(str)
                .str.strip()
                .str.upper()
            )

        # ─── Diagnostics ───
        print(f"✅ Loaded sheet: '{self.sheet_name}'")
        print(f"   Rows: {len(self.df)}")
        print(f"   Columns ({len(self.df.columns)}): {list(self.df.columns)}")
        print(f"   Mapped canonical fields: {list(self._col_map.keys())}")

        # Warn on unmapped columns (helps catch header typos)
        missing = [
            canon for canon in COLUMN_ALIASES
            if canon not in self._col_map
        ]
        if missing:
            print(f"   ⚠️ Unmapped fields: {missing}")

        # Special check for site_key_location columns
        if "assign_hub" not in self._col_map:
            print("   ⚠️ Column M (ASSIGN_HUB) not found — "
                  "site_key_location will rely on NEW ASSIGN_HUB")

    # ─────────────────────────────────────────────────────────
    # Internal helpers
    # ─────────────────────────────────────────────────────────

    @staticmethod
    def _find_sheet(xl: pd.ExcelFile, target: str) -> str:
        """Find sheet by exact, then case-insensitive match."""
        if target in xl.sheet_names:
            return target

        target_norm = _norm_upper(target)
        for s in xl.sheet_names:
            if _norm_upper(s) == target_norm:
                return s

        raise ValueError(
            f"Sheet '{target}' not found. "
            f"Available sheets: {xl.sheet_names}"
        )

    def _find_column(self, candidates: list[str]) -> Optional[str]:
        """Find a column by trying each candidate alias."""
        col_lookup = {_norm_upper(c): c for c in self.df.columns}

        for candidate in candidates:
            key = _norm_upper(candidate)
            if key in col_lookup:
                return col_lookup[key]

        # Fallback — substring match on first candidate
        if candidates:
            first = _norm_upper(candidates[0])
            for norm, orig in col_lookup.items():
                if first in norm or norm in first:
                    return orig

        return None

    def _value(self, row, canon: str) -> str:
        """Get a canonical field's value from a row. Returns '' if not found."""
        col = self._col_map.get(canon)
        if not col:
            return ""
        return str(row.get(col, "")).strip()

    # ─────────────────────────────────────────────────────────
    # Public API
    # ─────────────────────────────────────────────────────────

    def get_site(self, plaid: str) -> Optional[dict]:
        """
        Return a fully composed site dict for the given PLAID, or None.
        All Nokia TSSR text fields are pre-composed here.
        """
        plaid = str(plaid).strip().upper()
        plaid_col = self._col_map.get("plaid")
        if not plaid_col:
            return None

        rows = self.df[self.df[plaid_col] == plaid]
        if rows.empty:
            return None
        row = rows.iloc[0]

        # ─── Site address: Barangay, Municipality, Province ───
        barangay     = self._value(row, "barangay")
        municipality = self._value(row, "municipality")
        province     = self._value(row, "province")

        address_parts = [barangay, municipality, province]
        site_address = ", ".join(p for p in address_parts if p)

        # ─── Coordinates: 5 decimal places ───
        lat = self._value(row, "latitude")
        lon = self._value(row, "longitude")
        site_coords = ""
        if lat and lon:
            try:
                site_coords = f"{float(lat):.5f}, {float(lon):.5f}"
            except (ValueError, TypeError):
                site_coords = f"{lat}, {lon}"

        # ─── Field Officer (column S and column U) ───
        fo_name   = self._value(row, "engineer_anm1")
        fo_mobile = self._value(row, "contact_number")

        # ─── Site key location (Column M, fallback Column Q) ───
        site_key_location = (
            self._value(row, "assign_hub")        # Column M
            or self._value(row, "new_hub")        # Fallback: Column Q
        )

        # ─── TCO name (Column N) ───
        towercо = self._value(row, "towerco")

        return {
            # ── Identity ──
            "site_id":            self._value(row, "plaid"),
            "site_name":          self._value(row, "site"),
            "region":             self._value(row, "region"),
            "site_address":       site_address,
            "site_coords":        site_coords,

            # ── Geography (individual fields) ──
            "province":           province,
            "municipality":       municipality,
            "barangay":           barangay,

            # ── Coordinates (raw) ──
            "latitude":           lat,
            "longitude":          lon,

            # ── Towerco ──
            "towerco":            towercо,

            # ── Site key location ──
            "site_key_location":  site_key_location,

            # ── Field Officer ──
            "fo_name":            fo_name,
            "fo_mobile":          fo_mobile,
            "site_contact_person": fo_name,
            "site_contact_mobile": fo_mobile,

            # ── Extra fields ──
            "wireline":           self._value(row, "wireline"),
            "bcf":                self._value(row, "bcf"),
            "territory":          self._value(row, "territory"),
            "assign_hub":         self._value(row, "assign_hub"),
            "engineer_ah":        self._value(row, "engineer_ah"),
            "engineer_anm1":      self._value(row, "engineer_anm1"),
            "engineer_anm1_id":   self._value(row, "engineer_anm1_id"),
            "anm_head":           self._value(row, "anm_head"),
            "roh":                self._value(row, "roh"),
        }

    def list_plaids(self) -> list[str]:
        """Return sorted list of all PLAIDs."""
        plaid_col = self._col_map.get("plaid")
        if not plaid_col:
            return []
        return sorted(
            self.df[plaid_col]
            .dropna()
            .astype(str)
            .str.strip()
            .str.upper()
            .unique()
            .tolist()
        )

    def search_sites(self, query: str, limit: int = 50) -> list[tuple[str, str]]:
        """Search by PLAID or site name substring."""
        plaid_col = self._col_map.get("plaid")
        site_col  = self._col_map.get("site")
        if not plaid_col:
            return []

        q = query.strip().upper()
        results = []

        for _, row in self.df.iterrows():
            plaid = str(row.get(plaid_col, "")).strip().upper()
            name  = str(row.get(site_col, "")).strip() if site_col else ""

            if q in plaid or q in name.upper():
                results.append((plaid, name))
                if len(results) >= limit:
                    break

        return results

    # ─────────────────────────────────────────────────────────
    # Diagnostics
    # ─────────────────────────────────────────────────────────

    def debug_info(self) -> dict:
        """Return diagnostic info about the loaded workbook."""
        return {
            "path":           str(self.path),
            "sheet":          self.sheet_name,
            "rows":           len(self.df),
            "columns_raw":    list(self.df.columns),
            "columns_mapped": self._col_map,
            "missing_fields": [
                canon for canon in COLUMN_ALIASES
                if canon not in self._col_map
            ],
            "sample_plaids":  self.list_plaids()[:5],
        }


# ═════════════════════════════════════════════════════════════
# Standalone utilities
# ═════════════════════════════════════════════════════════════

def compose_address(site: dict) -> str:
    """Barangay, Municipality, Province."""
    parts = [
        site.get("barangay", "").strip(),
        site.get("municipality", "").strip(),
        site.get("province", "").strip(),
    ]
    return ", ".join(p for p in parts if p)


def compose_coords(site: dict) -> str:
    """5-decimal coordinate pair."""
    lat = site.get("latitude", "").strip()
    lon = site.get("longitude", "").strip()
    if not lat or not lon:
        return ""
    try:
        return f"{float(lat):.5f}, {float(lon):.5f}"
    except (ValueError, TypeError):
        return f"{lat}, {lon}"


# ═════════════════════════════════════════════════════════════
# CLI test — run: python excel_loader.py data/xxx.xlsx [PLAID]
# ═════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("Usage: python excel_loader.py <path-to-xlsx> [PLAID]")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    print(f"Loading: {xlsx_path}")

    ml = SiteMasterlist(xlsx_path)

    print("\n─── Diagnostics ───")
    print(json.dumps(ml.debug_info(), indent=2, default=str))

    if len(sys.argv) >= 3:
        plaid = sys.argv[2]
        print(f"\n─── Looking up PLAID: {plaid} ───")
        site = ml.get_site(plaid)
        if site:
            print(json.dumps(site, indent=2))
        else:
            print(f"❌ PLAID '{plaid}' not found")
    else:
        print("\n─── First 5 PLAIDs ───")
        for p in ml.list_plaids()[:5]:
            print(f"  {p}")
