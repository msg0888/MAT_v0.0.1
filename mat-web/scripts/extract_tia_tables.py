import argparse
import json
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


def get_range(
    cells: dict[str, tuple[str | None, str]],
    shared: list[str],
    start: str,
    end: str,
) -> list[list[object | None]]:
    c1, r1 = cell_to_coord(start)
    c2, r2 = cell_to_coord(end)
    c1i, c2i = col_to_idx(c1), col_to_idx(c2)
    out = []
    for r in range(r1, r2 + 1):
        row_vals = []
        for ci in range(c1i, c2i + 1):
            ref = f"{idx_to_col(ci)}{r}"
            t, val = cells.get(ref, (None, None))
            if val is None:
                row_vals.append(None)
                continue
            if t == "s":
                text = shared[int(val)] if val.isdigit() else ""
                row_vals.append(text)
            else:
                try:
                    num = float(val)
                    if num.is_integer():
                        num = int(num)
                    row_vals.append(num)
                except ValueError:
                    row_vals.append(val)
        out.append(row_vals)
    return out


def extract_tables(workbook: Path) -> dict:
    with zipfile.ZipFile(workbook) as z:
        shared = read_shared_strings(z)
        cells_h = parse_sheet(z, "xl/worksheets/sheet13.xml")
        cells_i = parse_sheet(z, "xl/worksheets/sheet14.xml")

        return {
            "source": workbook.name,
            "tables": {
                "TIA-222-H": {
                    "T2.3": {
                        "range": "J18:M21",
                        "columns": ["risk", "i_ice", "i_earthquake", "i_wind"],
                        "values": get_range(cells_h, shared, "J18", "M21"),
                    },
                    "T2.4": {
                        "range": "J28:O30",
                        "columns": ["exposure", "zg_us", "zg_si", "alpha", "kzmin", "kc"],
                        "values": get_range(cells_h, shared, "J28", "O30"),
                    },
                    "T2.5": {
                        "range": "J35:M37",
                        "values": get_range(cells_h, shared, "J35", "M37"),
                    },
                    "T2.6": {
                        "range": "J42:P50",
                        "values": get_range(cells_h, shared, "J42", "P50"),
                    },
                    "T2.11": {
                        "range": "J67:R72",
                        "values": get_range(cells_h, shared, "J67", "R72"),
                    },
                    "T2.12": {
                        "range": "J83:R88",
                        "values": get_range(cells_h, shared, "J83", "R88"),
                    },
                },
                "TIA-222-I": {
                    "T2.4": {
                        "range": "J28:O31",
                        "columns": ["exposure", "zg_us", "zg_si", "alpha", "kzmin", "kc"],
                        "values": get_range(cells_i, shared, "J28", "O31"),
                    }
                },
            },
        }


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    default_output = root / "apps" / "api" / "data" / "tia_tables.json"

    parser = argparse.ArgumentParser(description="Extract TIA tables from MAT workbook.")
    parser.add_argument("workbook", type=Path, help="Path to MAT xlsm workbook")
    parser.add_argument("--output", type=Path, default=default_output, help="Output JSON path")
    args = parser.parse_args()

    data = extract_tables(args.workbook)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
