# ai_helper.py
"""
AI-assisted field extraction via external LLM (Gemini / ChatGPT).
User copies a prompt → feeds to LLM → pastes JSON back into the app.

NOTE: work_permit, access_requirement, and site_key_location are NOT
extracted by AI. They are derived by the app from the masterlist +
towerco rules (see towercо_rules.py).
"""
from __future__ import annotations
import json
import re


EXTRACTION_PROMPT = """You are a telecom site survey analyst extracting data from an Ericsson TSSR (Technical Site Survey Report).

I will provide you with the Ericsson TSSR PDF. Extract the following fields and return them as a single JSON object.

═══════════════════════════════════════════════════════════════
CRITICAL RULES
═══════════════════════════════════════════════════════════════
1. Return ONLY the JSON object. No explanations, no markdown fences, no preamble, no trailing text.
2. Use the exact values from the CHOICES / DEFAULT lines. Case-sensitive.
3. If the TSSR uses a synonym, map it to the closest CHOICE.
4. For list fields, only include values you actually found.
5. Booleans must be true or false (lowercase, JSON style).
6. When a field has a DEFAULT, use it if the TSSR doesn't explicitly state that field.

═══════════════════════════════════════════════════════════════
FIELD-BY-FIELD GUIDE
═══════════════════════════════════════════════════════════════

"site_class"
  WHERE: "Site Class", "Tower Class", or "Structure Classification"
  CHOICES: "C1", "C2", "C3", "C4", "C5", "Custom"
  DEFAULT: "C3"  <-- Most Philippine sites are C3. Use "C3" if not found.
  EXAMPLE: "Site Class: C2" -> "C2"

"room_access"
  WHERE: "Room Access", "Access to Site", or "Site Access" narrative
  CHOICES: "Outdoor site", "Indoor site", "No room access", "Street cabinet", "Shelter"
  EXAMPLE: "N/A - Outdoor site / No room access" -> "Outdoor site"
  If not found, use "".

"cabin_location"
  WHERE: "Cabin Location", "Equipment Location", or equipment placement narrative
  CHOICES: "Ground Level", "Rooftop", "On Tower", "Pole-mounted", "Shelter", "Indoor Room"
  EXAMPLE: "Ground level cabin" -> "Ground Level"
  If not found, use "".

"flood_history"
  WHERE: "Flood History", "Flood Prone", or site condition narrative
  CHOICES: "None", "Low", "Moderate", "High", "Unknown", "N/A"
  DEFAULT: "N/A"  <-- Use "N/A" if the TSSR doesn't mention flooding.
  EXAMPLE: "Flood History: None" -> "None"

"hauling_remarks"
  WHERE: "Hauling Remarks", "Hauling Notes", "Access Notes"
  FREE TEXT — copy verbatim. Max 120 chars.
  If not found, use "N/A".

"site_profile"
  WHERE: "Site Profile", "Towerco Profile", "Site Category"
  CHOICES: "GT Wireless", "Rural", "Urban", "Suburban", "Highway", "Coastal"
  NOTE: Do NOT use "City" or "Greenfield" here — those belong to site_type.
  EXAMPLE: "Site Profile: GT Wireless" -> "GT Wireless"
  If not found, use "".

"site_owner"
  WHERE: "Site Owner" checkboxes
  CHOICES (pick exactly ONE): "Globe", "Private", "Government", "TCO"
  EXAMPLE: If "TCO" is checked -> "TCO"
  If multiple are checked, pick the one with an X or checkmark.
  If not found, use "".

"site_security"
  WHERE: "Security" checkboxes
  CHOICES (multi-select list, only the CHECKED ones):
    "E-LOCK", "Caretaker", "Manned", "Roving Security", "Others"
  EXAMPLE: E-LOCK checked -> ["E-LOCK"]
  If none checked, use [].

"site_type"
  WHERE: "Site Type" or narrative description of the structure
  CHOICES (pick exactly ONE): "Greenfield/Outdoor", "Street Cabinet", "Indoor", "Others"
  EXAMPLE: "Outdoor site" -> "Greenfield/Outdoor"
  If not found, use "".

"site_accessible"
  WHERE: "Accessible to Vehicle" OR "Not accessible to Vehicle" OR access narrative
  BOOLEAN: true if vehicle can reach the site, false if foot trail/boat only.
  EXAMPLE: "Site is accessible to vehicles" -> true

"no_bridge"
  WHERE: "No. of Bridge" field. If empty or N/A, use "N/A".

"foot_trail"
  WHERE: "Foot Trail" field. If empty or N/A, use "N/A".

"bridge_ton"
  WHERE: "Bridge (ton)" field. If empty or N/A, use "N/A".

"foot_bridge"
  WHERE: "Foot Bridge" field. If empty or N/A, use "N/A".

"distance_m"
  WHERE: "Distance (m)" field. If empty or N/A, use "N/A".

"by_boat"
  WHERE: "By Boat" field. If empty or N/A, use "N/A".

═══════════════════════════════════════════════════════════════
DO NOT EXTRACT THESE FIELDS
═══════════════════════════════════════════════════════════════
The following fields are derived by the application, NOT by you.
Return "" for all of them:

  "site_key_location"     — comes from the masterlist (Column M)
  "work_permit"           — derived from towercо rules
  "access_requirement"    — derived from towercо rules

═══════════════════════════════════════════════════════════════
REQUIRED JSON SCHEMA — return EXACTLY this structure
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
  "access_requirement": "",
  "work_permit": "",
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
EXAMPLE OUTPUT for a typical site
═══════════════════════════════════════════════════════════════
{
  "site_class": "C3",
  "room_access": "Outdoor site",
  "cabin_location": "Ground Level",
  "flood_history": "N/A",
  "hauling_remarks": "N/A - No hauling required",
  "site_profile": "GT Wireless",
  "site_key_location": "",
  "site_owner": "TCO",
  "site_security": ["E-LOCK"],
  "access_requirement": "",
  "work_permit": "",
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
    "access_requirement": "str",
    "work_permit":        "str",
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
    "flood_history":  {"None", "Low", "Moderate", "High",
                       "Unknown", "N/A", ""},
    "site_profile":   {"GT Wireless", "Rural", "Urban", "Suburban",
                       "Highway", "Coastal", ""},
    "site_owner":     {"Globe", "Private", "Government", "TCO", ""},
    "site_security":  {"E-LOCK", "Caretaker", "Manned",
                       "Roving Security", "Others"},
    "site_type":      {"Greenfield/Outdoor", "Street Cabinet",
                       "Indoor", "Others", ""},
}


def build_prompt() -> str:
    return EXTRACTION_PROMPT


def extract_json_from_response(text: str) -> dict | None:
    if not text:
        return None
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass
    return None


def _match_choice(value: str, choices: set) -> str:
    if value in choices:
        return value
    for c in choices:
        if c.lower() == value.lower():
            return c
    return ""


def validate_and_normalize(data: dict) -> tuple[dict, list[str]]:
    warnings = []
    clean = {}
    if not isinstance(data, dict):
        return {}, ["Response is not a JSON object."]

    for key, expected in EXPECTED_FIELDS.items():
        val = data.get(key)
        if val is None:
            clean[key] = [] if expected == "list" else \
                         False if expected == "bool" else ""
            warnings.append(f"Missing field: `{key}`")
            continue

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

        if key in FIELD_CHOICES and expected == "list":
            allowed = FIELD_CHOICES[key]
            original = list(clean[key])
            clean[key] = [v for v in clean[key] if v in allowed]
            removed = set(original) - set(clean[key])
            if removed:
                warnings.append(
                    f"`{key}`: removed invalid values {sorted(removed)}"
                )

    # Force skip fields to empty — they're derived elsewhere
    for skip in ("work_permit", "access_requirement", "site_key_location"):
        if skip in clean:
            clean[skip] = ""

    return clean, warnings
