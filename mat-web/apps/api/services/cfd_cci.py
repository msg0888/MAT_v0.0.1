from typing import Any, Dict, List, Optional, Tuple


def normalize_angle(angle: float) -> float:
    return angle % 360.0


def lookup_cfd_value(
    cfd_rows: List[Dict[str, Any]],
    manufacturer: Optional[str],
    model: Optional[str],
    azimuth: float,
) -> Optional[float]:
    if manufacturer is None or model is None:
        return None
    row = _find_cfd_row(cfd_rows, manufacturer, model)
    if not row:
        return None
    angle = normalize_angle(azimuth)
    return _interpolate_cfd(row, angle)


def lookup_cci_value(
    cci_rows: List[Dict[str, Any]],
    manufacturer: Optional[str],
    model: Optional[str],
    azimuth: float,
    ice_thickness: float,
) -> Optional[float]:
    if manufacturer is None or model is None:
        return None
    row = _find_cci_row(cci_rows, manufacturer, model)
    if not row:
        return None
    front = _to_float(row.get("front"))
    side = _to_float(row.get("side"))
    if front is None or side is None:
        return None

    if ice_thickness > 0:
        front = _interpolate_ice(row, "front", ice_thickness)
        side = _interpolate_ice(row, "side", ice_thickness)
        if front is None or side is None:
            return None

    return _interpolate_front_side(front, side, azimuth)


def _find_cfd_row(
    rows: List[Dict[str, Any]],
    manufacturer: str,
    model: str,
) -> Optional[Dict[str, Any]]:
    key = _key(manufacturer, model)
    for row in rows:
        if _key(row.get("manufacturer"), row.get("model")) == key:
            return row
    return None


def _find_cci_row(
    rows: List[Dict[str, Any]],
    manufacturer: str,
    model: str,
) -> Optional[Dict[str, Any]]:
    key = _key(manufacturer, model)
    for row in rows:
        if _key(row.get("manufacturer"), row.get("model")) == key:
            return row
    return None


def _key(manufacturer: Optional[str], model: Optional[str]) -> Tuple[str, str]:
    return (str(manufacturer or "").strip().upper(), str(model or "").strip().upper())


def _interpolate_cfd(row: Dict[str, Any], angle: float) -> Optional[float]:
    angle = normalize_angle(angle)
    lo = (int(angle // 10) * 10) % 360
    hi = (lo + 10) % 360
    if angle == lo:
        return _cfd_value(row, lo)
    v1 = _cfd_value(row, lo)
    v2 = _cfd_value(row, hi)
    if v1 is None or v2 is None:
        return None
    return _lerp(v1, v2, (angle - lo) / 10.0)


def _cfd_value(row: Dict[str, Any], angle: int) -> Optional[float]:
    key = f"deg_{angle}"
    if key not in row:
        return None
    return _to_float(row.get(key))


def _interpolate_front_side(front: float, side: float, azimuth: float) -> float:
    angle = normalize_angle(azimuth)
    rad = _deg_to_rad(angle)
    c = abs(_cos(rad))
    s = abs(_sin(rad))
    return front * (c * c) + side * (s * s)


def _interpolate_ice(
    row: Dict[str, Any],
    key_base: str,
    ice_thickness: float,
) -> Optional[float]:
    if ice_thickness <= 0:
        return _to_float(row.get(key_base))

    t = ice_thickness
    t1 = 0.5
    t2 = 1.0
    t3 = 2.0

    v1 = _to_float(row.get(f"{key_base}_ice_0_5"))
    v2 = _to_float(row.get(f"{key_base}_ice_1_0"))
    v3 = _to_float(row.get(f"{key_base}_ice_2_0"))

    if t <= t1:
        if v1 is None:
            return None
        return _lerp(v1, v2, t / t2) if v2 is not None else v1
    if t <= t2:
        if v1 is None or v2 is None:
            return None
        return _lerp(v1, v2, (t - t1) / (t2 - t1))
    if t <= t3:
        if v2 is None or v3 is None:
            return None
        return _lerp(v2, v3, (t - t2) / (t3 - t2))
    if v2 is None or v3 is None:
        return None
    return _lerp(v2, v3, (t - t2) / (t3 - t2))


def _lerp(v1: float, v2: float, t: float) -> float:
    return v1 + (v2 - v1) * t


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError:
        return None


def _deg_to_rad(deg: float) -> float:
    import math

    return math.radians(deg)


def _sin(value: float) -> float:
    import math

    return math.sin(value)


def _cos(value: float) -> float:
    import math

    return math.cos(value)
