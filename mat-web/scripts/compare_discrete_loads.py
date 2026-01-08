import argparse
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


def col_to_idx(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def idx_to_col(idx: int) -> str:
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(rem + 65) + letters
    return letters


def cell_to_coord(cell: str) -> tuple[str, int]:
    match = re.match(r"([A-Z]+)([0-9]+)", cell)
    return match.group(1), int(match.group(2))


def read_shared_strings(z: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    values = []
    for si in root.findall("ns:si", ns):
        text = "".join([t.text or "" for t in si.findall(".//ns:t", ns)])
        values.append(text)
    return values


def parse_sheet(z: zipfile.ZipFile, path: str) -> dict[str, tuple[str | None, str]]:
    root = ET.fromstring(z.read(path))
    ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    cells: dict[str, tuple[str | None, str]] = {}
    for row in root.findall("ns:sheetData/ns:row", ns):
        for c in row.findall("ns:c", ns):
            ref = c.get("r")
            t = c.get("t")
            v = c.find("ns:v", ns)
            if v is None:
                continue
            cells[ref] = (t, v.text or "")
    return cells


def get_cell(cells: dict[str, tuple[str | None, str]], shared: list[str], ref: str):
    t, val = cells.get(ref, (None, None))
    if val is None:
        return None
    if t == "s":
        return shared[int(val)] if val.isdigit() else ""
    try:
        num = float(val)
        if num.is_integer():
            return int(num)
        return num
    except ValueError:
        return val


def is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def extract_discrete_rows(cells, shared) -> list[dict]:
    rows = []
    for r in range(4, 200):
        model = get_cell(cells, shared, f"B{r}")
        if is_missing(model):
            continue
        row = {}
        for col in [
            "A",
            "B",
            "C",
            "D",
            "E",
            "F",
            "G",
            "H",
            "I",
            "J",
            "K",
            "L",
            "M",
            "N",
            "O",
            "P",
            "Q",
            "T",
            "U",
            "V",
            "W",
            "X",
            "Y",
            "AB",
            "AC",
            "AD",
            "AE",
            "AF",
            "AG",
            "AJ",
            "AK",
            "AL",
            "AM",
            "AN",
            "AO",
            "AR",
            "AS",
            "AT",
            "AU",
        ]:
            row[col] = get_cell(cells, shared, f"{col}{r}")
        row["_row"] = r
        rows.append(row)
    return rows


def extract_discrete_expected(cells, shared, row: int) -> dict:
    mapping = {
        "front_caaa": "AV",
        "side_caaa": "AW",
        "ice_front_caaa": "AX",
        "ice_side_caaa": "AY",
        "kz": "BA",
        "kd": "BB",
        "qz": "BC",
        "qiz": "BF",
    }
    out = {}
    for key, col in mapping.items():
        out[key] = get_cell(cells, shared, f"{col}{row}")
    return out


def extract_code_inputs(cells, shared) -> dict:
    keys = [
        "C17",
        "C18",
        "C21",
        "C22",
        "C24",
        "C27",
        "C28",
        "C30",
        "H16",
        "H44",
        "H40",
        "H43",
        "M16",
        "M17",
        "M18",
        "M19",
        "M22",
        "R16",
        "BA39",
        "BB39",
        "BC39",
        "BD39",
        "BA40",
        "BB40",
        "BC40",
        "BD40",
    ]
    out = {}
    for key in keys:
        out[key] = get_cell(cells, shared, key)
    return out


def extract_code_placement(cells, shared) -> list[dict]:
    rows = []
    for r in range(45, 72):
        label = get_cell(cells, shared, f"AM{r}")
        if is_missing(label):
            continue
        rows.append(
            {
                "label": label,
                "elevation": get_cell(cells, shared, f"AP{r}"),
                "positions": get_cell(cells, shared, f"AQ{r}"),
                "azimuths": get_cell(cells, shared, f"AR{r}"),
                "edge_distance": get_cell(cells, shared, f"AS{r}"),
                "front_force": get_cell(cells, shared, f"AT{r}"),
                "side_force": get_cell(cells, shared, f"AU{r}"),
            }
        )
    return rows


def extract_geometry(cells, shared) -> dict:
    members = []
    for r in range(6, 500):
        label = get_cell(cells, shared, f"B{r}")
        if is_missing(label):
            continue
        i_node = get_cell(cells, shared, f"M{r}")
        j_node = get_cell(cells, shared, f"N{r}")
        length = get_cell(cells, shared, f"O{r}")
        orientation = get_cell(cells, shared, f"P{r}")
        members.append(
            {
                "label": label,
                "i_node_label": i_node,
                "j_node_label": j_node,
                "length": length,
                "orientation_deg": orientation,
            }
        )

    nodes = []
    for r in range(6, 2000):
        label = get_cell(cells, shared, f"S{r}")
        if is_missing(label):
            continue
        nodes.append(
            {
                "label": label,
                "x": get_cell(cells, shared, f"T{r}"),
                "y": get_cell(cells, shared, f"U{r}"),
                "z": get_cell(cells, shared, f"V{r}"),
            }
        )

    return {"members": members, "nodes": nodes}


def compare_numeric(label, expected, actual, tol=1e-3) -> str | None:
    if expected is None and actual is None:
        return None
    try:
        exp = float(expected)
        act = float(actual)
    except (TypeError, ValueError):
        if expected == actual:
            return None
        return f"{label}: expected {expected}, got {actual}"
    if abs(exp - act) > tol:
        return f"{label}: expected {exp:.6f}, got {act:.6f}"
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare discrete loads outputs vs workbook.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("r3d", type=Path)
    args = parser.parse_args()

    with zipfile.ZipFile(args.workbook) as z:
        shared = read_shared_strings(z)
        cells_discrete = parse_sheet(z, "xl/worksheets/sheet6.xml")
        cells_code = parse_sheet(z, "xl/worksheets/sheet3.xml")
        cells_geometry = parse_sheet(z, "xl/worksheets/sheet4.xml")

    rows = extract_discrete_rows(cells_discrete, shared)
    expected = {row["_row"]: extract_discrete_expected(cells_discrete, shared, row["_row"]) for row in rows}
    code_inputs = extract_code_inputs(cells_code, shared)
    placement_expected = extract_code_placement(cells_code, shared)
    geometry = extract_geometry(cells_geometry, shared)

    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "apps" / "api"))
    from services.discrete_outputs import compute_outputs
    from services.tia_tables import derive_inputs as derive_tia_inputs

    if not geometry.get("members"):
        from parsers.r3d import parse_r3d

        geometry = parse_r3d(args.r3d.read_text(encoding="utf-8", errors="ignore"))
    derived = derive_tia_inputs(code_inputs)
    code_inputs = {**derived, **code_inputs}

    ice_thickness = code_inputs.get("H44")
    try:
        ice_thickness = float(ice_thickness)
    except (TypeError, ValueError):
        ice_thickness = 0.0
    outputs = compute_outputs(rows, [], [], ice_thickness, code_inputs, geometry)
    row_outputs = outputs.get("row_outputs", [])

    mismatches = []
    for idx, row in enumerate(rows):
        row_num = row["_row"]
        exp = expected.get(row_num, {})
        act = row_outputs[idx] if idx < len(row_outputs) else {}
        label = f"Row {row_num} ({row.get('B')})"
        for key in ["front_caaa", "side_caaa", "ice_front_caaa", "ice_side_caaa", "qz", "qiz"]:
            exp_val = exp.get(key)
            act_val = act.get(key)
            msg = compare_numeric(f"{label} {key}", exp_val, act_val)
            if msg:
                mismatches.append(msg)

    placement_calc = outputs.get("code_tables", {}).get("placement", [])
    placement_mismatches = []
    if len(placement_calc) != len(placement_expected):
        placement_mismatches.append(
            f"Placement count mismatch: expected {len(placement_expected)}, got {len(placement_calc)}"
        )
    for idx, row in enumerate(placement_expected):
        label = row.get("label")
        calc = placement_calc[idx] if idx < len(placement_calc) else None
        if not calc:
            placement_mismatches.append(f"Placement {label}: missing in calculated outputs")
            continue
        if str(calc.get("label")) != str(label):
            placement_mismatches.append(
                f"Placement row {idx + 1}: expected label {label}, got {calc.get('label')}"
            )
        for key, calc_key in [
            ("elevation", "elevation"),
            ("positions", "positions_text"),
            ("azimuths", "azimuths_text"),
            ("edge_distance", "edge_distance"),
            ("front_force", "member_front_force"),
            ("side_force", "member_side_force"),
        ]:
            exp_val = row.get(key)
            act_val = calc.get(calc_key)
            if key in {"positions", "azimuths", "edge_distance"}:
                if str(exp_val).strip() != str(act_val).strip():
                    placement_mismatches.append(
                        f"Placement {label} {key}: expected {exp_val}, got {act_val}"
                    )
            else:
                msg = compare_numeric(f"Placement {label} {key}", exp_val, act_val)
                if msg:
                    placement_mismatches.append(msg)

    print(f"Discrete rows compared: {len(rows)}")
    print(f"Mismatches: {len(mismatches)}")
    for msg in mismatches:
        print(f"- {msg}")

    print(f"Placement rows compared: {len(placement_expected)}")
    print(f"Placement mismatches: {len(placement_mismatches)}")
    for msg in placement_mismatches:
        print(f"- {msg}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
