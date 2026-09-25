# load_calc_helper.py
"""
AI-assisted load calculation extraction.
User copies prompt -> feeds to LLM -> pastes JSON back.
"""
from __future__ import annotations
import json
import re


LOAD_CALC_PROMPT = """You are a telecom DC power analyst reading an Ericsson TSSR (Technical Site Survey Report).

I will give you the TSSR PDF. Extract ONLY the rectifier / DC power data and return ONE JSON object. No markdown, no preamble, no trailing text.

═══════════════════════════════════════════════════════════════
SECTION 1 — EXISTING RECTIFIER (C7.1)
═══════════════════════════════════════════════════════════════
"rectifier_brand"          : e.g. "Eltek", "Emerson", "Huawei"   (string)
"max_modules"              : integer (total slots) e.g. 8
"module_rating_w"          : integer watts per module e.g. 2000
"battery_brand"            : e.g. "WOLONG"                        (string)
"battery_voltage"          : integer V e.g. 48
"battery_banks"            : integer e.g. 4
"battery_capacity_ah"      : integer Ah per cell e.g. 150
"present_load_a"           : float A e.g. 100.0
"modules_in_operation"     : integer e.g. 7
"actual_float_voltage_v"   : float V e.g. 53.83

═══════════════════════════════════════════════════════════════
SECTION 2 — PROPOSED ADDITIONAL LOADS (C7.2)
═══════════════════════════════════════════════════════════════
"proposed_loads"           : array of objects, one per load row.
                             Each item:
                             {
                               "load_no":      integer (1,2,3,...),
                               "equipment":    string  e.g. "NOKIA MF-2",
                               "power_w":      float   watts,
                               "current_a":    float   amps,
                               "cable_awg":    string  (may be ""),
                               "breaker_a":    string  e.g. "16" or "2 x 63AT"
                             }

═══════════════════════════════════════════════════════════════
RULES
═══════════════════════════════════════════════════════════════
1. Return ONLY the JSON object.
2. Use exact numbers from the TSSR. If a value is missing, use 0.
3. For Nokia MF-2 OLT default: power_w=400, current_a=7.4, breaker_a="16".
4. For Nokia FX-4 OLT default: power_w=515, current_a=7.5, breaker_a="16".
5. If the TSSR shows "N/A" or blank, use 0 for numbers, "" for strings.

═══════════════════════════════════════════════════════════════
REQUIRED JSON SCHEMA
═══════════════════════════════════════════════════════════════
{
  "rectifier_brand": "",
  "max_modules": 0,
  "module_rating_w": 0,
  "battery_brand": "",
  "battery_voltage": 0,
  "battery_banks": 0,
  "battery_capacity_ah": 0,
  "present_load_a": 0,
  "modules_in_operation": 0,
  "actual_float_voltage_v": 0,
  "proposed_loads": [
    {
      "load_no": 1,
      "equipment": "NOKIA MF-2",
      "power_w": 400,
      "current_a": 7.4,
      "cable_awg": "",
      "breaker_a": "16"
    }
  ]
}

═══════════════════════════════════════════════════════════════
NOW EXTRACT FROM THIS TSSR:
═══════════════════════════════════════════════════════════════
"""


EXPECTED = {
    "rectifier_brand":        "str",
    "max_modules":            "int",
    "module_rating_w":        "int",
    "battery_brand":          "str",
    "battery_voltage":        "int",
    "battery_banks":          "int",
    "battery_capacity_ah":    "int",
    "present_load_a":         "float",
    "modules_in_operation":   "int",
    "actual_float_voltage_v": "float",
    "proposed_loads":         "list",
}


def build_prompt() -> str:
    return LOAD_CALC_PROMPT


def extract_json(text: str) -> dict | None:
    if not text:
        return None
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return None


def validate(data: dict) -> tuple[dict, list[str]]:
    warnings = []
    clean = {}
    if not isinstance(data, dict):
        return {}, ["Response is not a JSON object."]

    for key, typ in EXPECTED.items():
        val = data.get(key)
        if val is None:
            clean[key] = [] if typ == "list" else (0 if typ != "str" else "")
            warnings.append(f"Missing field: `{key}`")
            continue

        try:
            if typ == "int":
                clean[key] = int(val)
            elif typ == "float":
                clean[key] = float(val)
            elif typ == "str":
                clean[key] = str(val).strip()
            elif typ == "list":
                if not isinstance(val, list):
                    clean[key] = []
                    warnings.append(f"`{key}` expected a list")
                else:
                    clean[key] = val
        except (ValueError, TypeError):
            clean[key] = 0 if typ != "str" else ""
            warnings.append(f"`{key}` could not be coerced to {typ}")

    return clean, warnings


def compute_sufficiency(d: dict) -> dict:
    """
    Mirror the Excel formulas:
      Existing Rectifier Capacity = modules_in_operation * (module_rating_w / 48)
      Existing Battery Capacity   = battery_banks * battery_capacity_ah
      Proposed Load (A)           = sum(current_a for each proposed load)
      Total Full Load Current     = present_load_a + proposed_load_a
      Available Rectifier Capacity= capacity - total_load - (10% * battery_capacity)
      Percent Utilization         = total_load / capacity
      BBUT (hours)                = battery_capacity / total_load
    """
    module_a = d.get("module_rating_w", 0) / 48.0 if d.get("module_rating_w") else 0
    cap_a    = d.get("modules_in_operation", 0) * module_a
    batt_ah  = d.get("battery_banks", 0) * d.get("battery_capacity_ah", 0)
    prop_a   = sum(l.get("current_a", 0) for l in d.get("proposed_loads", []))
    total_a  = d.get("present_load_a", 0) + prop_a
    avail_a  = cap_a - total_a - 0.10 * batt_ah
    pct_util = (total_a / cap_a) if cap_a else 0
    bbut_h   = (batt_ah / total_a) if total_a else 0

    return {
        "existing_rectifier_capacity_a": cap_a,
        "existing_battery_capacity_ah":  batt_ah,
        "proposed_load_a":               prop_a,
        "total_full_load_a":             total_a,
        "available_rectifier_capacity_a": avail_a,
        "percent_utilization":           pct_util,
        "bbut_hours":                    bbut_h,
    }
