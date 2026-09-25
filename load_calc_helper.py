# load_calc_helper.py
"""
AI-assisted load calculation extraction for the C7 block.

The deployment is ALWAYS a single Nokia MF-2 OLT, so the
proposed-load schedule is fixed here rather than extracted by AI.
"""
from __future__ import annotations

import json
import re


# ─────────────────────────────────────────────────────────────
# Fixed proposed load — Nokia MF-2 OLT
# ─────────────────────────────────────────────────────────────

NOKIA_MF2_LOAD = {
    "load_no":   1,
    "equipment": "NOKIA MF-2",
    "power_w":   400,
    "current_a": 7.4,
    "cable_awg": "",
    "breaker_a": "16",
}


# ─────────────────────────────────────────────────────────────
# Prompt (no more proposed_loads array)
# ─────────────────────────────────────────────────────────────

LOAD_CALC_PROMPT = """You are a telecom DC power analyst reading an Ericsson TSSR.

Extract ONLY the EXISTING rectifier / battery data and return ONE JSON object.
No markdown, no preamble, no trailing text.

The proposed load is ALWAYS a single Nokia MF-2 OLT and is filled by the app,
so do NOT return any proposed-load fields.

Fields:
"rectifier_brand"          (string) e.g. "Eltek", "Emerson", "Huawei"
"max_modules"              (int) total rectifier slots
"module_rating_w"          (int) watts per module
"battery_brand"            (string)
"battery_voltage"          (int) volts
"battery_banks"            (int)
"battery_capacity_ah"      (int) per cell
"present_load_a"           (float) amps
"modules_in_operation"     (int)
"actual_float_voltage_v"   (float) volts

Rules:
1. Return ONLY the JSON object.
2. Use exact numbers from the TSSR. If a value is missing, use 0 or "".
3. Do NOT include a "proposed_loads" field.

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
  "actual_float_voltage_v": 0
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
            clean[key] = 0 if typ != "str" else ""
            warnings.append(f"Missing field: `{key}`")
            continue
        try:
            if typ == "int":
                clean[key] = int(val)
            elif typ == "float":
                clean[key] = float(val)
            elif typ == "str":
                clean[key] = str(val).strip()
        except (ValueError, TypeError):
            clean[key] = 0 if typ != "str" else ""
            warnings.append(f"`{key}` could not be coerced to {typ}")

    # Always attach the fixed Nokia MF-2 load
    clean["proposed_loads"] = [dict(NOKIA_MF2_LOAD)]

    return clean, warnings


def compute_sufficiency(d: dict) -> dict:
    """
    Sufficiency math, hardcoded for a single Nokia MF-2 OLT load.
    """
    module_a = (d.get("module_rating_w") or 0) / 48.0
    cap_a    = (d.get("modules_in_operation") or 0) * module_a
    batt_ah  = (d.get("battery_banks") or 0) * (d.get("battery_capacity_ah") or 0)

    # Fixed proposed load
    prop_a   = NOKIA_MF2_LOAD["current_a"]          # 7.4 A

    total_a  = (d.get("present_load_a") or 0) + prop_a
    avail_a  = cap_a - total_a - 0.10 * batt_ah
    pct_util = (total_a / cap_a) if cap_a else 0
    bbut_h   = (batt_ah / total_a) if total_a else 0

    return {
        "existing_rectifier_capacity_a":  cap_a,
        "existing_battery_capacity_ah":   batt_ah,
        "proposed_load_a":                prop_a,
        "total_full_load_a":              total_a,
        "available_rectifier_capacity_a": avail_a,
        "percent_utilization":            pct_util,
        "bbut_hours":                     bbut_h,
    }
