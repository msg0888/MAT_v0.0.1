from typing import Any, Dict, List


def map_tml_records(parsed: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    mapped = []

    for record in parsed.get("discrete_loads", []):
        mapped.append(_map_discrete(record))

    for record in parsed.get("dish_loads", []):
        mapped.append(_map_dish(record))

    return {"rows": mapped, "counts": {"discrete": len(parsed.get("discrete_loads", [])), "dish": len(parsed.get("dish_loads", []))}}


def _map_discrete(record: Dict[str, Any]) -> Dict[str, Any]:
    return _base_row(
        manufacturer=record.get("database"),
        model=record.get("USName") or record.get("SIName"),
        load_type=record.get("type"),
        elevation=_to_number(record.get("startHeight")),
        vertical_offset=_to_number(record.get("verticalOffset")),
        horizontal_offset=_to_number(record.get("horizontalOffset")),
        height=_to_number(record.get("height")),
        width=_to_number(record.get("width")),
        depth=_to_number(record.get("depth")),
        weight=_to_number(record.get("selfWeight")),
        source="tml_discrete",
        raw=record,
    )


def _map_dish(record: Dict[str, Any]) -> Dict[str, Any]:
    diameter = _to_number(record.get("outsideDiameter"))
    return _base_row(
        manufacturer=record.get("database"),
        model=record.get("USName") or record.get("SIName"),
        load_type="Dish",
        elevation=_to_number(record.get("heightAboveBase")),
        vertical_offset=_to_number(record.get("verticalOffset")),
        horizontal_offset=_to_number(record.get("horizontalOffset")),
        height=diameter,
        width=diameter,
        depth=None,
        weight=_to_number(record.get("selfWeight")),
        source="tml_dish",
        raw=record,
    )


def _base_row(
    manufacturer: Any,
    model: Any,
    load_type: Any,
    elevation: Any,
    vertical_offset: Any,
    horizontal_offset: Any,
    height: Any,
    width: Any,
    depth: Any,
    weight: Any,
    source: str,
    raw: Dict[str, Any],
) -> Dict[str, Any]:
    row = {
        "A": manufacturer,
        "B": model,
        "C": "",
        "D": load_type,
        "E": "",
        "F": 1,
        "G": vertical_offset,
        "H": elevation,
        "I": horizontal_offset,
        "J": 0,
        "K": 0,
        "AR": height,
        "AS": width,
        "AT": depth,
        "AU": weight,
        "source": source,
        "raw": raw,
    }
    return row


def _to_number(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return value
