import math
from typing import Any, Dict, List, Optional, Tuple

from services.cfd_cci import lookup_cfd_value, lookup_cci_value, normalize_angle


ANGLE_STEPS = [0, 30, 60, 90, 120, 150, 180, 210, 240, 270, 300, 330]
SECTOR_COLUMNS = {
    "alpha": ("L", "M", "N", "O", "P", "Q"),
    "beta": ("T", "U", "V", "W", "X", "Y"),
    "gamma": ("AB", "AC", "AD", "AE", "AF", "AG"),
    "delta": ("AJ", "AK", "AL", "AM", "AN", "AO"),
}


def compute_outputs(
    rows: List[Dict[str, Any]],
    cfd_rows: List[Dict[str, Any]],
    cci_rows: List[Dict[str, Any]],
    ice_thickness: float,
    code_inputs: Optional[Dict[str, Any]] = None,
    geometry: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    code_inputs = code_inputs or {}
    geometry = geometry or {}
    global_wind_pressure = _pick_value(code_inputs, ["qz", "C40", "H40"])
    global_ice_pressure = _pick_value(code_inputs, ["qiz", "H43"])
    member_map, node_map = _geometry_maps(geometry)
    length_unit = _pick_str(code_inputs, ["RISA3D.Unit.Length", "R16"])
    if not length_unit:
        length_unit = _pick_str(geometry.get("units", {}), ["length"])
    convert_length = None
    if length_unit:
        unit = length_unit.strip().lower()
        convert_length = 1.0 if unit in {"in", "inch", "inches"} else 12.0

    wind_table = []
    ice_table = []
    placement_table = []
    row_outputs = []

    for row in rows:
        if not _row_used(row):
            continue
        manufacturer = row.get("A")
        model = row.get("B")
        shape = str(row.get("E") or "").strip().upper()
        load_type = str(row.get("D") or "").strip().upper()
        weight = _to_float(row.get("AU"))
        shield_front = _to_float(row.get("J")) or 1.0
        shield_side = _to_float(row.get("K")) or 1.0

        front_caaa, side_caaa = _front_side_caaa(
            row,
            shape,
            load_type,
            cci_rows,
            0.0,
            code_inputs,
        )
        ice_front_caaa, ice_side_caaa = _front_side_caaa(
            row,
            shape,
            load_type,
            cci_rows,
            ice_thickness,
            code_inputs,
        )

        wind_caaa = _directional_values(
            use_cfd=shape == "CFD",
            manufacturer=manufacturer,
            model=model,
            cfd_rows=cfd_rows,
            front_caaa=front_caaa,
            side_caaa=side_caaa,
            azimuth=0.0,
        )
        ice_caaa = _directional_values(
            use_cfd=False,
            manufacturer=manufacturer,
            model=model,
            cfd_rows=cfd_rows,
            front_caaa=ice_front_caaa,
            side_caaa=ice_side_caaa,
            azimuth=0.0,
        )

        shielded_wind_caaa = _directional_values(
            use_cfd=shape == "CFD",
            manufacturer=manufacturer,
            model=model,
            cfd_rows=cfd_rows,
            front_caaa=front_caaa,
            side_caaa=side_caaa,
            azimuth=_first_azimuth(row),
            shield_front=shield_front,
            shield_side=shield_side,
        )
        shielded_ice_caaa = _directional_values(
            use_cfd=False,
            manufacturer=manufacturer,
            model=model,
            cfd_rows=cfd_rows,
            front_caaa=ice_front_caaa,
            side_caaa=ice_side_caaa,
            azimuth=_first_azimuth(row),
            shield_front=shield_front,
            shield_side=shield_side,
        )

        row_qz, row_qiz = _row_pressures(row, shape, code_inputs)
        wind_pressure = row_qz if row_qz is not None else global_wind_pressure
        ice_pressure = row_qiz if row_qiz is not None else global_ice_pressure

        wind_table.append(
            {
                "model": model,
                "caaa": wind_caaa,
                "area_source": _area_source(shape),
                "wind_pressure": wind_pressure,
                "weight": weight,
            }
        )
        ice_table.append(
            {
                "model": model,
                "ice_caaa": ice_caaa,
                "area_source": _ice_area_source(shape),
                "ice_pressure": ice_pressure,
                "ice_weight": weight,
            }
        )

        placement_row = _placement_summary(
            row,
            code_inputs,
            member_map,
            node_map,
            front_caaa,
            side_caaa,
            wind_pressure,
            shape,
            convert_length,
        )
        if placement_row:
            placement_table.append(placement_row)

        row_outputs.append(
            {
                "model": model,
                "shape": shape,
                "weight": weight,
                "front_caaa": front_caaa,
                "side_caaa": side_caaa,
                "ice_front_caaa": ice_front_caaa,
                "ice_side_caaa": ice_side_caaa,
                "wind_caaa": wind_caaa,
                "ice_caaa": ice_caaa,
                "wind_caaa_shielded": shielded_wind_caaa,
                "ice_caaa_shielded": shielded_ice_caaa,
                "wind_pressure": wind_pressure,
                "ice_pressure": ice_pressure,
                "qz": row_qz,
                "qiz": row_qiz,
            }
        )

    return {
        "row_outputs": row_outputs,
        "code_tables": {
            "wind": wind_table,
            "ice": ice_table,
            "placement": placement_table,
        },
    }


def _directional_values(
    *,
    use_cfd: bool,
    manufacturer: Any,
    model: Any,
    cfd_rows: List[Dict[str, Any]],
    front_caaa: Optional[float],
    side_caaa: Optional[float],
    azimuth: float,
    shield_front: float = 1.0,
    shield_side: float = 1.0,
) -> Dict[int, Optional[float]]:
    if use_cfd:
        return {
            angle: lookup_cfd_value(cfd_rows, manufacturer, model, normalize_angle(angle + azimuth))
            for angle in ANGLE_STEPS
        }
    return {
        angle: _directional_caaa(front_caaa, side_caaa, shield_front, shield_side, angle, azimuth)
        for angle in ANGLE_STEPS
    }


def _front_side_caaa(
    row: Dict[str, Any],
    shape: str,
    load_type: str,
    cci_rows: List[Dict[str, Any]],
    ice_thickness: float,
    code_inputs: Dict[str, Any],
) -> Tuple[Optional[float], Optional[float]]:
    manufacturer = row.get("A")
    model = row.get("B")

    if load_type == "DISH":
        diameter = _to_float(row.get("AR"))
        if diameter is None:
            return None, None
        return _dish_caaa(diameter, ice_thickness, 0), _dish_caaa(diameter, ice_thickness, 1)

    if shape == "CCI":
        front = lookup_cci_value(cci_rows, manufacturer, model, 0, ice_thickness)
        side = lookup_cci_value(cci_rows, manufacturer, model, 90, ice_thickness)
        if front is not None or side is not None:
            return front, side

    front = _tia_caaa(row, 0, ice_thickness, code_inputs)
    side = _tia_caaa(row, 1, ice_thickness, code_inputs)
    return front, side


def _placement_summary(
    row: Dict[str, Any],
    code_inputs: Dict[str, Any],
    member_map: Dict[str, Dict[str, Any]],
    node_map: Dict[str, Dict[str, Any]],
    front_caaa: Optional[float],
    side_caaa: Optional[float],
    wind_pressure: Optional[float],
    shape: str,
    convert_length: Optional[float],
) -> Optional[Dict[str, Any]]:
    model = row.get("B")
    if _is_missing(model):
        return None

    qty_total = 0.0
    positions: List[str] = []
    azimuths: List[float] = []
    edge_distances: List[float] = []

    for sector, cols in SECTOR_COLUMNS.items():
        az = _to_float(row.get(cols[0]))
        qty = _to_float(row.get(cols[1]))
        pos_top = row.get(cols[2])
        pos_btm = row.get(cols[4])

        if qty is not None:
            qty_total += qty

        if not _is_missing(pos_top):
            pos_num = _to_float(pos_top)
            if pos_num is None or pos_num != 0:
                positions.append(str(pos_top).strip())

        if az is not None:
            azimuths.append(az)

        member_label = pos_top
        if _is_missing(pos_top) or (_to_float(pos_top) == 0):
            member_label = pos_btm
        edge_member = _edge_member_label(code_inputs, sector)
        edge_dir = _edge_direction(code_inputs, sector)
        edge_distance = _edge_distance_for_sector(
            sector,
            member_label,
            edge_member,
            edge_dir,
            member_map,
            node_map,
        )
        if edge_distance is not None:
            edge_distances.append(edge_distance)

    if not positions and not azimuths and qty_total == 0:
        return None

    elevation = _centerline_elevation(row, code_inputs)
    member_front_force, member_side_force = _member_forces(
        front_caaa,
        side_caaa,
        wind_pressure,
        shape,
        code_inputs,
    )

    qty_label = _format_quantity(qty_total)
    display_label = f"({qty_label}) {model}" if qty_label else str(model)

    return {
        "model": model,
        "quantity": qty_total if qty_total else None,
        "label": display_label,
        "elevation": elevation,
        "positions": positions,
        "positions_text": _format_list(positions),
        "azimuths": azimuths,
        "azimuths_text": _format_list(azimuths),
        "edge_distances": edge_distances,
        "edge_distance": _format_list(edge_distances, scale=convert_length, decimals=1),
        "member_front_force": member_front_force,
        "member_side_force": member_side_force,
    }


def _area_source(shape: str) -> str:
    if shape == "CFD":
        return "Manufacturer"
    if shape == "CCI":
        return "CCI"
    return "TIA"


def _ice_area_source(shape: str) -> str:
    if shape == "CCI":
        return "CCI"
    return "TIA"


def _tia_caaa(
    row: Dict[str, Any],
    side: int,
    ice_thickness: float,
    code_inputs: Dict[str, Any],
) -> Optional[float]:
    shape = str(row.get("E") or "").strip().upper()
    load_type = str(row.get("D") or "").strip().upper()
    height = _to_float(row.get("AR"))
    width = _to_float(row.get("AS"))
    depth = _to_float(row.get("AT"))

    if height is None or width is None:
        return None

    if load_type == "DISH":
        return _dish_caaa(height, ice_thickness, side)

    tia_version = _tia_version(row, code_inputs)
    return _caaa_formula(tia_version, height, width, depth, _shape_for_tia(shape), side, 0.0, ice_thickness)


def _dish_caaa(diameter: float, ice_thickness: float, side: int) -> float:
    area = (math.pi / 4.0) * (diameter + 2 * ice_thickness) ** 2
    if side == 1:
        area = area / 2.0
    return area / 144.0


def _shape_for_tia(shape: str) -> str:
    if shape in {"ROUND", "ROUND (FRONT)", "ROUND (BACK)"}:
        return "Round"
    if shape in {"HSS FLAT"}:
        return "Flat"
    if shape == "CFD":
        return "Flat"
    return "Flat"


def _tia_version(row: Dict[str, Any], code_inputs: Dict[str, Any]) -> str:
    tia = (
        row.get("TIA")
        or row.get("C18")
        or code_inputs.get("TIA")
        or code_inputs.get("C18")
    )
    if isinstance(tia, str):
        upper = tia.strip().upper()
        if "222-I" in upper or upper.endswith("-I") or upper == "I":
            return "I"
    return "H"


def _caaa_formula(
    tia: str,
    height: float,
    width: float,
    depth: Optional[float],
    shape: str,
    side: int,
    c_value: float,
    ice_thickness: float,
) -> float:
    if depth is None:
        depth = width

    if tia in {"I", "1"}:
        subcritical = 32
        supercritical = 64
    else:
        subcritical = 39
        supercritical = 78

    aspect_front = (height + 2 * ice_thickness) / (width + 2 * ice_thickness)
    aspect_side = (height + 2 * ice_thickness) / (depth + 2 * ice_thickness)
    aspect = aspect_front if side == 0 else aspect_side

    if shape == "Flat":
        ca = _flat_ca(aspect)
    else:
        ca = _round_ca(aspect, c_value, subcritical, supercritical)

    if side == 0:
        area = ca * (height + 2 * ice_thickness) * (width + 2 * ice_thickness)
    else:
        area = ca * (height + 2 * ice_thickness) * (depth + 2 * ice_thickness)
    return area / 144.0


def _flat_ca(aspect: float) -> float:
    if aspect <= 2.5:
        return 1.2
    if aspect == 7:
        return 1.4
    if aspect >= 25:
        return 2.0
    if 2.5 < aspect < 7:
        return 1.2 + (aspect - 2.5) * ((1.4 - 1.2) / (7 - 2.5))
    if 7 < aspect < 25:
        return 1.4 + (aspect - 7) * ((2 - 1.4) / (25 - 7))
    return 1.0


def _round_ca(aspect: float, c_value: float, subcritical: float, supercritical: float) -> float:
    c_val = c_value
    if c_val < subcritical:
        if aspect <= 2.5:
            return 0.7
        if 2.5 < aspect < 7:
            return 0.7 + (aspect - 2.5) * ((0.8 - 0.7) / (7 - 2.5))
        if aspect == 7:
            return 0.8
        if 7 < aspect < 25:
            return 0.8 + (aspect - 7) * ((1.2 - 0.8) / (25 - 7))
        return 1.2
    if subcritical <= c_val <= supercritical:
        if aspect <= 2.5:
            return 4.14 / (c_val ** 0.485)
        if 2.5 < aspect < 7:
            return (4.14 / (c_val ** 0.485)) + (aspect - 2.5) * (
                ((3.66 / (c_val ** 0.415)) - (4.14 / (c_val ** 0.485))) / (7 - 2.5)
            )
        if aspect == 7:
            return 3.66 / (c_val ** 0.415)
        if 7 < aspect < 25:
            return (3.66 / (c_val ** 0.415)) + (aspect - 7) * (
                ((46.8 / (c_val ** 1.0)) - (3.66 / (c_val ** 0.415))) / (25 - 7)
            )
        return 46.8 / (c_val ** 1.0)
    if aspect <= 2.5:
        return 0.5
    if 2.5 < aspect < 7:
        return 0.5 + (aspect - 2.5) * ((0.6 - 0.5) / (7 - 2.5))
    return 0.6


def _directional_caaa(
    front_caaa: Optional[float],
    side_caaa: Optional[float],
    shield_front: float,
    shield_side: float,
    direction: int,
    azimuth: float,
) -> Optional[float]:
    if front_caaa is None or side_caaa is None:
        return None

    rad = math.radians(direction + azimuth)
    return (front_caaa * shield_front * (math.cos(rad) ** 2)) + (
        side_caaa * shield_side * (math.sin(rad) ** 2)
    )


def _pick_value(inputs: Dict[str, Any], keys: List[str]) -> Optional[float]:
    for key in keys:
        if key in inputs:
            return _to_float(inputs.get(key))
    return None


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError:
        return None


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _row_used(row: Dict[str, Any]) -> bool:
    for value in row.values():
        if not _is_missing(value):
            return True
    return False


def _first_azimuth(row: Dict[str, Any]) -> float:
    for key in ("L", "T", "AB", "AJ"):
        value = _to_float(row.get(key))
        if value is not None:
            return value
    return 0.0


def _pick_str(inputs: Dict[str, Any], keys: List[str]) -> Optional[str]:
    for key in keys:
        if key in inputs:
            value = inputs.get(key)
            if value is None:
                continue
            text = str(value).strip()
            if text:
                return text
    return None


def _geometry_maps(
    geometry: Dict[str, Any],
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]]]:
    member_map: Dict[str, Dict[str, Any]] = {}
    for member in geometry.get("members", []) or []:
        label = member.get("label")
        if _is_missing(label):
            continue
        member_map[str(label).strip()] = member

    node_map: Dict[str, Dict[str, Any]] = {}
    for node in geometry.get("nodes", []) or []:
        label = node.get("label")
        if _is_missing(label):
            continue
        node_map[str(label).strip()] = node

    return member_map, node_map


def _row_pressures(
    row: Dict[str, Any],
    shape: str,
    code_inputs: Dict[str, Any],
) -> Tuple[Optional[float], Optional[float]]:
    vertical_offset = _to_float(row.get("G")) or 0.0
    kz = _compute_kz(vertical_offset, code_inputs)
    if kz is None:
        kz = _pick_value(code_inputs, ["Kz", "M22"])

    kzt = _pick_value(code_inputs, ["Kzt", "C27"])
    ks = _pick_value(code_inputs, ["Ks", "M16"])
    i_wind = _pick_value(code_inputs, ["I.wind", "I_wind", "Iwind"])
    v = _pick_value(code_inputs, ["V", "C22"])
    vi = _pick_value(code_inputs, ["Vi", "C30"])
    ke = _pick_value(code_inputs, ["Ke", "M18"])
    gh = _pick_value(code_inputs, ["Gh", "M19"])
    kd_global = _pick_value(code_inputs, ["Kd", "M17"])

    ke_factor = 1.0 if _tia_suffix(code_inputs) == "G" else ke
    row_kd = _row_kd(shape)

    qz = None
    if None not in (kz, kzt, ks, i_wind, v, gh) and ke_factor is not None:
        qz = 0.00256 * kz * kzt * ks * i_wind * (v ** 2) * ke_factor * row_kd * gh

    qiz = None
    kd_factor = 1.0 if shape in {"CFD", "CCI"} else kd_global
    if None not in (kz, kzt, ks, i_wind, vi, gh, kd_factor) and ke_factor is not None:
        qiz = 0.00256 * kz * kzt * ks * i_wind * (vi ** 2) * ke_factor * kd_factor * gh

    return qz, qiz


def _compute_kz(vertical_offset: float, code_inputs: Dict[str, Any]) -> Optional[float]:
    exposure = _pick_str(code_inputs, ["Exposure", "C24"])
    z = _pick_value(code_inputs, ["z", "H16"])
    if not exposure or z is None:
        return None

    z_offset = z + (vertical_offset / 12.0)
    exposure_val = exposure.strip().upper()

    if exposure_val == "BC":
        zg_us = _pick_value(code_inputs, ["zg.US", "zg.us", "zg_us"])
        zg_us_c = _pick_value(code_inputs, ["zg.US_c", "zg.us_c", "zg_us_c"])
        alpha = _pick_value(code_inputs, ["alpha", "α"])
        alpha_c = _pick_value(code_inputs, ["alpha_c", "α_c"])
        kzmin = _pick_value(code_inputs, ["Kzmin"])
        kzmin_c = _pick_value(code_inputs, ["Kzmin_c"])
        if None in (zg_us, zg_us_c, alpha, alpha_c, kzmin, kzmin_c):
            return None
        val1 = _clamp(2.41 * (z_offset / zg_us) ** (2 / alpha), kzmin, 2.41)
        val2 = _clamp(2.41 * (z_offset / zg_us_c) ** (2 / alpha_c), kzmin_c, 2.41)
        return (val1 + val2) / 2.0

    hurricane = _pick_str(code_inputs, ["Hurricane_Prone", "C28"])
    if _tia_suffix(code_inputs) == "I" and exposure_val == "D" and (hurricane or "").upper() == "YES":
        kzmin = _pick_value(code_inputs, ["Kzmin"])
        if kzmin is None:
            return None
        base = 2.01 * (z / 700.0) ** (2 / 11.5)
        return _clamp(base, kzmin, 2.01)

    zg_us = _pick_value(code_inputs, ["zg.US", "zg.us", "zg_us"])
    alpha = _pick_value(code_inputs, ["alpha", "α"])
    kzmin = _pick_value(code_inputs, ["Kzmin"])
    if None in (zg_us, alpha, kzmin):
        return None
    kz_factor = _pick_value(code_inputs, ["Kz_Factor"])
    if kz_factor is None:
        kz_factor = 2.41 if _tia_suffix(code_inputs) == "I" else 2.01
    val = kz_factor * (z_offset / zg_us) ** (2 / alpha)
    return _clamp(val, kzmin, kz_factor)


def _clamp(value: float, min_value: float, max_value: float) -> float:
    return min(max(value, min_value), max_value)


def _tia_suffix(code_inputs: Dict[str, Any]) -> Optional[str]:
    tia = code_inputs.get("TIA") or code_inputs.get("C18")
    if tia is None:
        return None
    text = str(tia).strip().upper()
    return text[-1] if text else None


def _row_kd(shape: str) -> float:
    return 1.0 if shape in {"CFD", "CCI"} else 0.95


def _centerline_elevation(row: Dict[str, Any], code_inputs: Dict[str, Any]) -> Optional[float]:
    elevation = _to_float(row.get("H"))
    if elevation is not None and elevation != 0:
        return elevation
    return _pick_value(code_inputs, ["z", "H16"])


def _member_forces(
    front_caaa: Optional[float],
    side_caaa: Optional[float],
    wind_pressure: Optional[float],
    shape: str,
    code_inputs: Dict[str, Any],
) -> Tuple[Optional[float], Optional[float]]:
    if front_caaa is None or side_caaa is None or wind_pressure is None:
        return None, None
    ka = _pick_value(code_inputs, ["Ka", "M20"])
    if ka is None:
        ka = 0.9
    kd = _row_kd(shape)
    front = front_caaa * wind_pressure * ka * kd
    side = side_caaa * wind_pressure * ka * kd
    return front, side


def _edge_member_label(code_inputs: Dict[str, Any], sector: str) -> Optional[str]:
    keys = {
        "alpha": ["edge_member_alpha", "BA39", "edge_member"],
        "beta": ["edge_member_beta", "BB39", "edge_member"],
        "gamma": ["edge_member_gamma", "BC39", "edge_member"],
        "delta": ["edge_member_delta", "BD39", "edge_member"],
    }.get(sector, [])
    return _pick_str(code_inputs, keys)


def _edge_direction(code_inputs: Dict[str, Any], sector: str) -> Optional[str]:
    keys = {
        "alpha": ["edge_direction_alpha", "BA40", "edge_direction"],
        "beta": ["edge_direction_beta", "BB40", "edge_direction"],
        "gamma": ["edge_direction_gamma", "BC40", "edge_direction"],
        "delta": ["edge_direction_delta", "BD40", "edge_direction"],
    }.get(sector, [])
    return _pick_str(code_inputs, keys)


def _edge_distance_for_sector(
    sector: str,
    member_label: Any,
    edge_member: Optional[str],
    edge_dir: Optional[str],
    member_map: Dict[str, Dict[str, Any]],
    node_map: Dict[str, Dict[str, Any]],
) -> Optional[float]:
    if _is_missing(member_label) or _is_missing(edge_member):
        return None
    member = member_map.get(str(member_label).strip())
    edge = member_map.get(str(edge_member).strip())
    if not member or not edge:
        return None

    angle = _to_float(edge.get("orientation_deg"))
    if angle is None:
        angle = _to_float(edge.get("azimuth") or edge.get("angle"))
    length = _to_float(edge.get("length"))
    edge_j_label = edge.get("j_node_label")
    if angle is None or length is None or _is_missing(edge_j_label):
        return None
    edge_node = node_map.get(str(edge_j_label).strip())
    if not edge_node:
        return None

    use_j = sector == "gamma"
    app_node_label = member.get("j_node_label") if use_j else member.get("i_node_label")
    if _is_missing(app_node_label):
        return None
    app_node = node_map.get(str(app_node_label).strip())
    if not app_node:
        return None

    pxi = _to_float(app_node.get("x"))
    pzi = _to_float(app_node.get("z"))
    mxj = _to_float(edge_node.get("x"))
    mzj = _to_float(edge_node.get("z"))
    if None in (pxi, pzi, mxj, mzj):
        return None

    theta = math.radians(angle)
    if sector in {"alpha", "gamma"}:
        ux = -math.cos(theta)
        uz = math.sin(theta)
    else:
        ux = math.cos(theta)
        uz = -math.sin(theta)

    dx = pxi - mxj
    dz = pzi - mzj
    direction = (edge_dir or "i+").strip().lower()
    offset = 0.0 if direction == "i+" else length
    distance = (dx * ux) + (dz * uz) + offset
    return round(distance, 0)


def _format_quantity(value: float) -> str:
    if value <= 0:
        return ""
    if abs(value - round(value)) < 1e-6:
        return str(int(round(value)))
    return _format_number(value, None)


def _format_list(
    values: List[Any],
    scale: Optional[float] = None,
    decimals: Optional[int] = None,
) -> str:
    items: List[str] = []
    for value in values:
        if _is_missing(value):
            continue
        num = _to_float(value)
        if num is None:
            text = str(value).strip()
            if text:
                items.append(text)
            continue
        if scale is not None:
            num *= scale
        items.append(_format_number(num, decimals))
    return ", ".join(items)


def _format_number(value: float, decimals: Optional[int]) -> str:
    if decimals is None:
        if abs(value - round(value)) < 1e-6:
            return str(int(round(value)))
        text = f"{value:.6f}".rstrip("0").rstrip(".")
        return text if text else "0"
    text = f"{value:.{decimals}f}"
    if decimals > 0:
        text = text.rstrip("0").rstrip(".")
    return text if text else "0"
