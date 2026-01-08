from typing import Any, Dict, List, Tuple


SECTOR_ORDER = ["alpha", "beta", "gamma", "delta"]

COLUMN_ALIASES = {
    "A": ["A", "manufacturer"],
    "B": ["B", "model"],
    "C": ["C", "condition"],
    "D": ["D", "type"],
    "E": ["E", "shape"],
    "F": ["F", "attach"],
    "G": ["G", "vertical_offset"],
    "H": ["H", "elevation"],
    "I": ["I", "horizontal_offset"],
    "J": ["J", "shield_front"],
    "K": ["K", "shield_side"],
    "L": ["L", "alpha_azimuth"],
    "M": ["M", "alpha_qty"],
    "N": ["N", "alpha_pos_top"],
    "O": ["O", "alpha_loc_top"],
    "P": ["P", "alpha_pos_btm"],
    "Q": ["Q", "alpha_loc_btm"],
    "T": ["T", "beta_azimuth"],
    "U": ["U", "beta_qty"],
    "V": ["V", "beta_pos_top"],
    "W": ["W", "beta_loc_top"],
    "X": ["X", "beta_pos_btm"],
    "Y": ["Y", "beta_loc_btm"],
    "AB": ["AB", "gamma_azimuth"],
    "AC": ["AC", "gamma_qty"],
    "AD": ["AD", "gamma_pos_top"],
    "AE": ["AE", "gamma_loc_top"],
    "AF": ["AF", "gamma_pos_btm"],
    "AG": ["AG", "gamma_loc_btm"],
    "AJ": ["AJ", "delta_azimuth"],
    "AK": ["AK", "delta_qty"],
    "AL": ["AL", "delta_pos_top"],
    "AM": ["AM", "delta_loc_top"],
    "AN": ["AN", "delta_pos_btm"],
    "AO": ["AO", "delta_loc_btm"],
    "AR": ["AR", "height"],
    "AS": ["AS", "width"],
    "AT": ["AT", "depth"],
    "AU": ["AU", "weight"],
}

SECTOR_COLUMNS = {
    "alpha": {"az": "L", "qty": "M", "pos_top": "N", "loc_top": "O", "pos_btm": "P", "loc_btm": "Q"},
    "beta": {"az": "T", "qty": "U", "pos_top": "V", "loc_top": "W", "pos_btm": "X", "loc_btm": "Y"},
    "gamma": {"az": "AB", "qty": "AC", "pos_top": "AD", "loc_top": "AE", "pos_btm": "AF", "loc_btm": "AG"},
    "delta": {"az": "AJ", "qty": "AK", "pos_top": "AL", "loc_top": "AM", "pos_btm": "AN", "loc_btm": "AO"},
}


def apply_import_rows(rows: List[Dict[str, Any]], imported: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    merged = list(rows)
    for row in imported:
        merged.append(row)
    return merged


def apply_cfd_cci(
    rows: List[Dict[str, Any]],
    cfd_rows: List[Dict[str, Any]],
    cci_rows: List[Dict[str, Any]],
    ice_thickness: float,
) -> List[Dict[str, Any]]:
    from services.cfd_cci import lookup_cfd_value, lookup_cci_value

    updated = []
    for row in rows:
        shape = str(_get_value(row, "E") or "").strip().upper()
        manufacturer = _get_value(row, "A")
        model = _get_value(row, "B")
        if shape == "CFD":
            row = _apply_cfd(row, cfd_rows, manufacturer, model)
        elif shape == "CCI":
            row = _apply_cci(row, cci_rows, manufacturer, model, ice_thickness)
        updated.append(row)
    return updated


def _apply_cfd(
    row: Dict[str, Any],
    cfd_rows: List[Dict[str, Any]],
    manufacturer: Any,
    model: Any,
) -> Dict[str, Any]:
    from services.cfd_cci import lookup_cfd_value

    normalized = row.get("normalized", {})
    cfd = {}
    for sector, data in normalized.items():
        az = data.get("lookup_azimuth")
        if az is None:
            continue
        value = lookup_cfd_value(cfd_rows, manufacturer, model, az)
        if value is not None:
            cfd[sector] = value
    if cfd:
        row = dict(row)
        row["cfd"] = cfd
    return row


def _apply_cci(
    row: Dict[str, Any],
    cci_rows: List[Dict[str, Any]],
    manufacturer: Any,
    model: Any,
    ice_thickness: float,
) -> Dict[str, Any]:
    from services.cfd_cci import lookup_cci_value

    normalized = row.get("normalized", {})
    cci = {}
    for sector, data in normalized.items():
        az = data.get("lookup_azimuth")
        if az is None:
            continue
        value = lookup_cci_value(cci_rows, manufacturer, model, az, ice_thickness)
        if value is not None:
            cci[sector] = value
    if cci:
        row = dict(row)
        row["cci"] = cci
    return row


def normalize_rows(
    rows: List[Dict[str, Any]],
    mount_azimuths: Dict[str, Any],
    qty_sectors: int,
) -> List[Dict[str, Any]]:
    normalized_rows = []
    sector_count = max(1, min(qty_sectors, len(SECTOR_ORDER)))

    for row in rows:
        if not _row_used(row):
            continue
        normalized = {}
        for idx, sector in enumerate(SECTOR_ORDER[:sector_count]):
            az_key = SECTOR_COLUMNS[sector]["az"]
            az_val = _get_value(row, az_key)
            mount_val = mount_azimuths.get(sector)
            if _is_missing(az_val) or _is_missing(mount_val):
                continue
            try:
                az = float(az_val)
                mount = float(mount_val)
            except (TypeError, ValueError):
                continue
            normalized_az = (az - mount) % 360.0
            normalized[sector] = {
                "app_azimuth": az,
                "mount_azimuth": mount,
                "normalized_azimuth": normalized_az,
                "lookup_azimuth": normalized_az,
            }
        row_out = dict(row)
        if normalized:
            row_out["normalized"] = normalized
        normalized_rows.append(row_out)
    return normalized_rows


def validate_rows(
    rows: List[Dict[str, Any]],
    qty_sectors: int,
) -> Dict[str, Any]:
    errors: List[str] = []
    warnings: List[str] = []
    sector_count = max(1, min(qty_sectors, len(SECTOR_ORDER)))
    required_sectors = SECTOR_ORDER[:sector_count]

    for idx, row in enumerate(rows, start=1):
        if not _row_used(row):
            continue
        row_id = _row_id(row, idx)
        _check_required(row, row_id, errors, warnings, required_sectors)
        _check_attach(row, row_id, warnings, required_sectors)

    return {"valid": len(errors) == 0, "errors": errors, "warnings": warnings}


def _row_id(row: Dict[str, Any], idx: int) -> str:
    label = _get_value(row, "B")
    if _is_missing(label):
        label = f"row {idx}"
    return str(label).strip()


def _check_required(
    row: Dict[str, Any],
    row_id: str,
    errors: List[str],
    warnings: List[str],
    required_sectors: List[str],
) -> None:
    base_required = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "AR", "AS", "AT", "AU"]
    for col in base_required:
        if _is_missing(_get_value(row, col)):
            warnings.append(f"{row_id}: missing required input {col}")

    for sector in required_sectors:
        cols = SECTOR_COLUMNS[sector]
        for key in ("az", "qty"):
            col = cols[key]
            if _is_missing(_get_value(row, col)):
                warnings.append(f"{row_id}: missing {sector} {col}")


def _check_attach(
    row: Dict[str, Any],
    row_id: str,
    warnings: List[str],
    required_sectors: List[str],
) -> None:
    attach_val = _get_value(row, "F")
    try:
        attach = int(float(attach_val))
    except (TypeError, ValueError):
        attach = None

    for sector in required_sectors:
        cols = SECTOR_COLUMNS[sector]
        has_bottom = not _is_missing(_get_value(row, cols["pos_btm"])) or not _is_missing(
            _get_value(row, cols["loc_btm"])
        )
        if attach == 2:
            for key in ("pos_btm", "loc_btm"):
                if _is_missing(_get_value(row, cols[key])):
                    warnings.append(f"{row_id}: missing {sector} {cols[key]} for attach=2")
        elif attach == 1 and has_bottom:
            warnings.append(f"{row_id}: attach=1 but bottom fields present for {sector}")


def _row_used(row: Dict[str, Any]) -> bool:
    for value in row.values():
        if not _is_missing(value):
            return True
    return False


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _get_value(row: Dict[str, Any], key: str) -> Any:
    aliases = COLUMN_ALIASES.get(key, [key])
    for alias in aliases:
        if alias in row:
            return row.get(alias)
    return None
