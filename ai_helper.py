# ai_helper.py
"""
AI-assisted field extraction via external LLM (Gemini / ChatGPT).
User copies a prompt → feeds to LLM → pastes JSON back into the app.
"""
from __future__ import annotations
import json
import re


# ═════════════════════════════════════════════════════════════
# The prompt — copy/paste friendly, strict JSON output
# ═════════════════════════════════════════════════════════════

EXTRACTION_PROMPT = """You are a telecom site survey analyst extracting data from an Ericsson TSSR (Technical Site Survey Report).

I will provide you with the Ericsson TSSR PDF. Extract the following fields and return them as a single JSON object.

═══════════════════════════════════════════════════════════════
CRITICAL RULES
═══════════════════════════════════════════════════════════════
1. Return ONLY the JSON object. No explanations, no markdown fences, no preamble, no trailing text.
2. Use "" (empty string) if a field is genuinely not found — do NOT invent values.
3. Use ONLY the exact values listed in the CHOICES below. Case-sensitive.
4. If the TSSR uses a synonym, map it to the closest CHOICE.
5. For fields with empty list [] — only include values you actually found.
6. Booleans must be true or false (lowercase, JSON style).

═══════════════════════════════════════════════════════════════
FIELD-BY-FIELD GUIDE (where to look, what to accept)
═══════════════════════════════════════════════════════════════

"site_class"
  WHERE: "Site Class", "Tower Class", or "Structure Classification" in the site details section
  CHOICES: "C1", "C2", "C3", "C4", "C5", "Custom", ""
  EXAMPLE: If TSSR says "Site Class: C3" -> "C3"
  If not found, use "".

"room_access"
  WHERE: "Room Access", "Access to Site", "Site Access" narrative
  CHOICES: "Outdoor site", "Indoor site", "No room access", "Street cabinet", "Shelter", ""
  EXAMPLE: "N/A - Outdoor site / No room access" -> "Outdoor site"
  If not found, use "".

"cabin_location"
  WHERE: "Cabin Location", "Equipment Location", or equipment placement narrative
  CHOICES: "Ground Level", "Rooftop", "On Tower", "Pole-mounted", "Shelter", "Indoor Room", ""
  EXAMPLE: "Ground level cabin location" -> "Ground Level"
  If not found, use "".

"flood_history"
  WHERE: "Flood History", "Flood Prone", or site condition narrative
  CHOICES: "None", "Low", "Moderate", "High", "Unknown", ""
  EXAMPLE: "None" -> "None"
  If not found, use "".

"hauling_remarks"
  WHERE: "Hauling Remarks", "Hauling Notes", "Access Notes"
  FREE TEXT — copy verbatim. Max 120 characters.
  If not found, use "".

"site_profile"
  WHERE: "Site Profile", "Towerco Profile", "Site Category"
  CHOICES: "GT Wireless", "Rural", "Urban", "Suburban", "Highway", "Coastal", ""
  NOTE: This is NOT the same as Site Type. Do NOT use "City" or "Greenfield" here.
  EXAMPLE: "GT Wireless" -> "GT Wireless"
  If not found, use "".

"site_key_location"
  WHERE: "Site Key Location", "Key Location", "Access Keys"
  FREE TEXT — copy verbatim. Max 120 characters.
  EXAMPLE: "E-LOCK/RECT. KEY & CABINET KEY AT GT HUB"

"site_owner"
  WHERE: "Site Owner" checkboxes on the site details page
  CHOICES (pick exactly ONE): "Globe", "Private", "Government", "TCO"
  EXAMPLE: If "TCO" is checked -> "TCO"
  If multiple are checked, pick the one marked with X or checkmark.

"site_security"
  WHERE: "Security" checkboxes
  CHOICES (multi-select, list only the CHECKED ones):
    "E-LOCK", "Caretaker", "Manned", "Roving Security", "Others"
  EXAMPLE: If E-LOCK is checked -> ["E-LOCK"]
  If no security is checked -> []

"work_permit"
  WHERE: "Work Permit" checkboxes OR "Site Access Requirement" narrative
  CHOICES (pick exactly ONE): "RAAWA", "Others"
  EXAMPLE: If narrative mentions "RAAWA permit required" -> "RAAWA"
  If not found, use "".

"access_requirement"
  WHERE: "Site Access Requirement" narrative field
  FREE TEXT — copy verbatim. Max 200 characters.
  EXAMPLE: "RAAWA, HSWP AND APPROVED PHILTOWER TICKET" or "Office Hours"
  If not found, use "".

"site_type"
  WHERE: "Site Type" or narrative description of the site structure
  CHOICES (pick exactly ONE): "Greenfield/Outdoor", "Street Cabinet", "Indoor", "Others"
  EXAMPLE: "Outdoor site" or "Greenfield" -> "Greenfield/Outdoor"
  If not found, use "".

"site_accessible"
  WHERE: "Accessible to Vehicle" or "Not accessible to Vehicle" or access narrative
  BOOLEAN: true if vehicle can reach the site, false if foot trail/boat only
  EXAMPLE: "Site is accessible to vehicles" -> true

"no_bridge"
  WHERE: "No. of Bridge" field
  FREE TEXT. If empty or N/A, use "N/A".

"foot_trail"
  WHERE: "Foot Trail" field
  FREE TEXT. If empty or N/A, use "N/A".

"bridge_ton"
  WHERE: "Bridge (ton)" field
  FREE TEXT. If empty or N/A, use "N/A".

"foot_bridge"
  WHERE: "Foot Bridge" field
  FREE TEXT. If empty or N/A, use "N/A".

"distance_m"
  WHERE: "Distance (m)" field
  FREE TEXT. If empty or N/A, use "N/A".

"by_boat"
  WHERE: "By Boat" field
  FREE TEXT. If empty or N/A, use "N/A".

═══════════════════════════════════════════════════════════════
REQUIRED JSON SCHEMA (return exactly this structure)
═══════════════════════════════════════════════════════════════
{
  "site_class": "",
  "room_access": "",
  "cabin_location": "",
  "flood_history": "",
  "hauling_remarks": "",
  "site_profile": "",
  "site_key_location": "",
  "site_owner": "",
  "site_security": [],
  "work_permit": "",
  "access_requirement": "",
  "site_type": "",
  "site_accessible": true,
  "no_bridge": "",
  "foot_trail": "",
  "bridge_ton": "",
  "foot_bridge": "",
  "distance_m": "",
  "by_boat": ""
}

═══════════════════════════════════════════════════════════════
EXAMPLE OUTPUT (for a typical site)
═══════════════════════════════════════════════════════════════
{
  "site_class": "C3",
  "room_access": "Outdoor site",
  "cabin_location": "Ground Level",
  "flood_history": "None",
  "hauling_remarks": "N/A - No hauling required",
  "site_profile": "GT Wireless",
  "site_key_location": "E-LOCK/RECT. KEY & CABINET KEY AT GT HUB",
  "site_owner": "TCO",
  "site_security": ["E-LOCK", "Caretaker"],
  "work_permit": "RAAWA",
  "access_requirement": "RAAWA, HSWP AND APPROVED PHILTOWER TICKET",
  "site_type": "Greenfield/Outdoor",
  "site_accessible": true,
  "no_bridge": "N/A",
  "foot_trail": "N/A",
  "bridge_ton": "N/A",
  "foot_bridge": "N/A",
  "distance_m": "N/A",
  "by_boat": "N/A"
}

═══════════════════════════════════════════════════════════════
NOW EXTRACT FROM THIS TSSR:
═══════════════════════════════════════════════════════════════
"""


# ═════════════════════════════════════════════════════════════
# Expected fields + allowed choices
# ═════════════════════════════════════════════════════════════

EXPECTED_FIELDS = {
    "site_class":         "str",
    "room_access":        "str",
    "cabin_location":     "str",
    "flood_history":      "str",
    "hauling_remarks":    "str",
    "site_profile":       "str",
    "site_key_location":  "str",
    "site_owner":         "str",
    "site_security":      "list",
    "work_permit":        "str",
    "access_requirement": "str",
    "site_type":          "str",
    "site_accessible":    "bool",
    "no_bridge":          "str",
    "foot_trail":         "str",
    "bridge_ton":         "str",
    "foot_bridge":        "str",
    "distance_m":         "str",
    "by_boat":            "str",
}


FIELD_CHOICES = {
    "site_class":     {"C1", "C2", "C3", "C4", "C5", "Custom", ""},
    "room_access":    {"Outdoor site", "Indoor site", "No room access",
                       "Street cabinet", "Shelter", ""},
    "cabin_location": {"Ground Level", "Rooftop", "On Tower",
                       "Pole-mounted", "Shelter", "Indoor Room", ""},
    "flood_history":  {"None", "Low", "Moderate", "High", "Unknown", ""},
    "site_profile":   {"GT Wireless", "Rural", "Urban", "Suburban",
                       "Highway", "Coastal", ""},
    "site_owner":     {"Globe", "Private", "Government", "TCO", ""},
    "site_security":  {"E-LOCK", "Caretaker", "Manned",
                       "Roving Security", "Others"},
    "work_permit":    {"RAAWA", "Others", ""},
    "site_type":      {"Greenfield/Outdoor", "Street Cabinet",
                       "Indoor", "Others", ""},
}


# ═════════════════════════════════════════════════════════════
# Public API
# ═════════════════════════════════════════════════════════════

def build_prompt() -> str:
    """Return the prompt string shown to the user."""
    return EXTRACTION_PROMPT


def extract_json_from_response(text: str) -> dict | None:
    """Robustly extract JSON from an AI response."""
    if not text:
        return None

    text = text.strip()

    # Strip markdown fences
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Find first { ... } block
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass

    return None


def _match_choice(value: str, choices: set) -> str:
    """Case-insensitive match against allowed choices."""
    if value in choices:
        return value
    for c in choices:
        if c.lower() == value.lower():
            return c
    return ""


def validate_and_normalize(data: dict) -> tuple[dict, list[str]]:
    """
    Validate AI JSON, coerce types, enforce choices.
    Returns (clean_dict, warnings_list).
    """
    warnings = []
    clean = {}

    if not isinstance(data, dict):
        return {}, ["Response is not a JSON object."]

    for key, expected in EXPECTED_FIELDS.items():
        val = data.get(key)

        # Missing field
        if val is None:
            clean[key] = [] if expected == "list" else \
                         False if expected == "bool" else ""
            warnings.append(f"Missing field: `{key}`")
            continue

        # ─── Type coercion ───
        if expected == "str":
            clean[key] = str(val).strip()

        elif expected == "list":
            if isinstance(val, list):
                clean[key] = [str(v).strip() for v in val if str(v).strip()]
            elif isinstance(val, str):
                clean[key] = [val.strip()] if val.strip() else []
            else:
                clean[key] = []
                warnings.append(f"`{key}` expected a list.")

        elif expected == "bool":
            if isinstance(val, bool):
                clean[key] = val
            elif isinstance(val, str):
                clean[key] = val.strip().lower() in ("true", "yes", "1")
            else:
                clean[key] = bool(val)

        # ─── Enforce choices for single-value string fields ───
        if key in FIELD_CHOICES and expected == "str":
            allowed = FIELD_CHOICES[key]
            if clean[key] not in allowed:
                matched = _match_choice(clean[key], allowed)
                if matched:
                    warnings.append(
                        f"`{key}`: normalized '{clean[key]}' -> '{matched}'"
                    )
                    clean[key] = matched
                else:
                    warnings.append(
                        f"`{key}`: '{clean[key]}' not in allowed choices — "
                        f"reset to empty."
                    )
                    clean[key] = ""

        # ─── Enforce choices for list fields ───
        if key in FIELD_CHOICES and expected == "list":
            allowed = FIELD_CHOICES[key]
            original = list(clean[key])
            clean[key] = [v for v in clean[key] if v in allowed]
            removed = set(original) - set(clean[key])
            if removed:
                warnings.append(
                    f"`{key}`: removed invalid values {sorted(removed)}"
                )

    return clean, warnings
