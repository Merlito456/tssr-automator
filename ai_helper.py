# ai_helper.py
"""
AI-assisted field extraction via external LLM.
User copies prompt → feeds to Gemini → pastes JSON back.
"""
from __future__ import annotations
import json
import re


EXTRACTION_PROMPT = """You are a telecom site survey analyst.

I will provide you with an Ericsson TSSR (Technical Site Survey Report) PDF or its text content.

Your task: extract the following fields and return them as a single JSON object.

RULES:
- Return ONLY the JSON object. No explanations, no markdown fences, no preamble.
- If a field is not found, use "" (empty string) for text, false for booleans.
- For "site_owner", pick exactly one: "Globe", "Private", "Government", "TCO"
- For "site_security", pick all that apply from: "E-LOCK", "Caretaker", "Manned", "Roving Security", "Others"
- For "site_type", pick exactly one: "Greenfield/Outdoor", "Street Cabinet", "Indoor", "Others"
- For "work_permit", pick exactly one: "RAAWA", "Others"
- Booleans must be true or false (lowercase).

REQUIRED JSON SCHEMA:
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

Now, here is the Ericsson TSSR content:
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


def extract_json_from_response(text: str) -> dict | None:
    """Robustly extract JSON from an AI response."""
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


def validate_and_normalize(data: dict) -> tuple[dict, list[str]]:
    """Validate AI JSON, coerce types, return (clean, warnings)."""
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

    return clean, warnings


def build_prompt() -> str:
    return EXTRACTION_PROMPT
