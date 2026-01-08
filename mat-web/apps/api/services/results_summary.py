from typing import Any, Dict, List, Optional


def summarize_envelope(
    rows: List[Dict[str, Any]],
    geometry: Dict[str, Any],
) -> Dict[str, Any]:
    section_map = {
        member.get("label"): member.get("section_set_label")
        for member in geometry.get("members", [])
        if member.get("label")
    }
    section_info = _section_info(geometry)

    ordered_sets: List[str] = []
    grouped: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        member = row.get("member")
        section_set = row.get("section_set") or section_map.get(member)
        if not section_set:
            continue
        if section_set not in grouped:
            grouped[section_set] = []
            ordered_sets.append(section_set)
        grouped[section_set].append(row)

    summary_rows: List[Dict[str, Any]] = []
    for section_set in ordered_sets:
        members = grouped.get(section_set, [])
        max_bending, bend_member = _max_with_member(members, "code_check")
        max_shear, shear_member = _max_with_member(members, "shear_check")
        info = section_info.get(section_set, {})
        status = None
        if max_bending is not None or max_shear is not None:
            if (max_bending or 0) > 1.05 or (max_shear or 0) > 1.05:
                status = "Fail"
            else:
                status = "Pass"
        summary_rows.append(
            {
                "section_set": section_set,
                "section_name": info.get("section_name"),
                "section_type": info.get("section_type"),
                "shape_class": info.get("shape_class"),
                "bending_max": max_bending,
                "bending_member": bend_member,
                "shear_max": max_shear,
                "shear_member": shear_member,
                "status": status,
            }
        )

    return {"count": len(summary_rows), "rows": summary_rows}


def _max_with_member(
    rows: List[Dict[str, Any]],
    key: str,
) -> tuple[Optional[float], Optional[str]]:
    best_val: Optional[float] = None
    best_member: Optional[str] = None
    for row in rows:
        val = row.get(key)
        if val is None:
            continue
        try:
            num = float(val)
        except (TypeError, ValueError):
            continue
        if best_val is None or num > best_val:
            best_val = num
            best_member = row.get("member")
    return best_val, best_member


def _section_info(geometry: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    info: Dict[str, Dict[str, Any]] = {}
    for group in geometry.get("section_sets", {}).values():
        for entry in group or []:
            label = entry.get("label")
            if not label:
                continue
            info[label] = {
                "section_name": entry.get("section_name"),
                "section_type": entry.get("section_type"),
                "shape_class": entry.get("shape_class"),
            }
    return info
