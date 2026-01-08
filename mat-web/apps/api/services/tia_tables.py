import json
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional


DATA_PATH = Path(__file__).resolve().parent.parent / "data" / "tia_tables.json"


@lru_cache
def _load_tables() -> Dict[str, Any]:
    with DATA_PATH.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def derive_inputs(inputs: Dict[str, Any]) -> Dict[str, Any]:
    derived: Dict[str, Any] = {}
    tia = str(inputs.get("TIA") or inputs.get("C18") or "").strip()
    suffix = _tia_suffix(tia)

    exposure = str(inputs.get("Exposure") or inputs.get("C24") or "").strip()
    if exposure:
        row = _lookup_exposure(suffix, exposure)
        if row:
            derived.update(
                _numeric_map(
                    {
                        "zg.US": row[1],
                        "zg.SI": row[2],
                        "alpha": row[3],
                        "Kzmin": row[4],
                        "Kc": row[5],
                    }
                )
            )
        if exposure.strip().upper() == "BC":
            derived.update(
                {
                    "alpha_c": 9.8,
                    "Kzmin_c": 0.85,
                    "zg.US_c": 2460,
                }
            )

    if suffix:
        derived["Kz_Factor"] = 2.41 if suffix == "I" else 2.01

    risk = str(inputs.get("Risk") or inputs.get("C21") or "").strip()
    asce = str(inputs.get("ASCE") or inputs.get("C17") or "").strip()
    if risk:
        row = _lookup_risk(risk)
        if row:
            if asce and asce.upper() != "ASCE 7-05":
                derived["I.wind"] = 1
            else:
                wind = row[3]
                if isinstance(wind, (int, float)):
                    derived["I.wind"] = wind

            if risk.upper() != "I":
                derived.update(
                    _numeric_map(
                        {
                            "I.ice": row[1],
                            "I.earthquake": row[2],
                        }
                    )
                )

    return derived


def _lookup_risk(risk: str) -> Optional[List[Any]]:
    tables = _load_tables().get("tables", {})
    table = tables.get("TIA-222-H", {}).get("T2.3", {})
    return _match_row(table.get("values", []), risk)


def _lookup_exposure(suffix: Optional[str], exposure: str) -> Optional[List[Any]]:
    tables = _load_tables().get("tables", {})
    version = "TIA-222-I" if suffix == "I" else "TIA-222-H"
    table = tables.get(version, {}).get("T2.4", {})
    return _match_row(table.get("values", []), exposure)


def _match_row(values: List[List[Any]], key: str) -> Optional[List[Any]]:
    target = str(key).strip().upper()
    for row in values:
        if not row:
            continue
        label = row[0]
        if label is None:
            continue
        if str(label).strip().upper() == target:
            return row
    return None


def _tia_suffix(tia: str) -> Optional[str]:
    upper = tia.upper()
    if not upper:
        return None
    if "222-I" in upper or upper.endswith("-I") or upper == "I":
        return "I"
    if "222-G" in upper or upper.endswith("-G") or upper == "G":
        return "G"
    if "222-H" in upper or upper.endswith("-H") or upper == "H":
        return "H"
    return upper[-1]


def _numeric_map(values: Dict[str, Any]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for key, value in values.items():
        if isinstance(value, (int, float)):
            out[key] = value
    return out
