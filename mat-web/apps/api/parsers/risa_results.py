import csv
from typing import Any, Dict, List, Optional


def parse_risa_results(text: str) -> Dict[str, Any]:
    lines = text.splitlines()
    in_section = False
    expected: Optional[int] = None
    rows: List[Dict[str, Any]] = []

    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        if line.startswith("["):
            in_section = line.upper().startswith("[ENVELOPE AISC")
            expected = None
            continue
        if not in_section:
            continue
        if expected is None:
            try:
                expected = int(float(line))
            except ValueError:
                expected = 0
            continue

        parsed = _parse_csv_line(line)
        if not parsed:
            continue
        row = _map_row(parsed)
        if row:
            rows.append(row)
        if expected and len(rows) >= expected:
            break

    return {
        "section": "ENVELOPE AISC 15TH (360-16): LRFD STEEL CODE CHECKS",
        "count": len(rows),
        "rows": rows,
    }


def _parse_csv_line(line: str) -> List[str]:
    reader = csv.reader([line], quotechar="'", skipinitialspace=True)
    for row in reader:
        return [item.strip() for item in row]
    return []


def _map_row(values: List[str]) -> Optional[Dict[str, Any]]:
    if len(values) < 3:
        return None
    # First value is a row index.
    row_index = _to_float([values[0]], 0)
    data = values[1:]
    mapped = {
        "row_index": row_index,
        "member": _to_value(data, 0),
        "shape": _to_value(data, 1),
        "code_check": _to_float(data, 2),
        "loc": _to_float(data, 3),
        "lc": _to_float(data, 4),
        "shear_check": _to_float(data, 5),
        "shear_loc": _to_float(data, 6),
        "shear_dir": _to_value(data, 7),
        "shear_lc": _to_float(data, 8),
        "phi_pnc": _to_float(data, 9),
        "phi_pnt": _to_float(data, 10),
        "phi_mny": _to_float(data, 11),
        "phi_mnz": _to_float(data, 12),
        "cb": _to_float(data, 13),
        "eqn_1": _to_value(data, 14),
        "eqn_2": _to_value(data, 15),
        "eqn_3": _to_value(data, 16),
        "eqn_4": _to_value(data, 17),
        "connection_check": _to_float(data, 18),
        "raw": data,
    }
    return mapped


def _to_value(values: List[str], idx: int) -> Optional[str]:
    if idx >= len(values):
        return None
    value = values[idx]
    if value == "":
        return None
    return value


def _to_float(values: List[str], idx: int) -> Optional[float]:
    if idx >= len(values):
        return None
    value = values[idx]
    if value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None
