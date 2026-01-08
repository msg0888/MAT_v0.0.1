import shlex
from typing import Dict, List, Optional


def parse_r3d(text: str) -> dict:
    sections = _extract_sections(text)
    nodes = _parse_nodes(sections.get("[NODES]", []))
    section_sets = {
        "hr": _parse_section_sets(sections.get("[.HR_STEEL_SECTION_SETS]", [])),
        "cf": _parse_section_sets(sections.get("[.CF_STEEL_SECTION_SETS]", [])),
        "al": _parse_section_sets(sections.get("[.ALUMINUM_SECTION_SETS]", [])),
        "general": _parse_section_sets(sections.get("[.GENERAL_SECTION_SETS]", [])),
    }

    units = _parse_units(sections.get("[UNITS]", []))
    axis = _parse_axis(sections.get("[.SOLUTION_PARAMETERS]", []))
    design_code = _parse_design_code(sections.get("[.DESIGN_CODES]", []))
    version = _parse_version(sections.get("[VERSION_NUMBER]", []))

    members = _parse_members(
        sections.get("[.MEMBERS_MAIN_DATA]", []),
        nodes,
        section_sets,
        axis,
    )

    missing_shape_data = any(
        member.get("shape_warning") for member in members if member is not None
    )

    return {
        "members": members,
        "nodes": nodes,
        "section_sets": section_sets,
        "units": units,
        "axis": axis,
        "design_code": design_code,
        "risa_version": version,
        "flags": {"missing_shape_data": missing_shape_data},
    }


def _extract_sections(text: str) -> Dict[str, List[str]]:
    sections: Dict[str, List[str]] = {}
    current: Optional[str] = None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("["):
            end_idx = line.find("]")
            if end_idx == -1:
                continue
            header = line[: end_idx + 1]
            upper = header.upper()
            if upper.startswith("[.END") or upper.startswith("[END"):
                current = None
                continue
            current = header
            sections.setdefault(current, [])
            continue
        if current is not None:
            sections[current].append(line)

    return sections


def _parse_table(lines: List[str]) -> List[dict]:
    rows = []
    for line in lines:
        tokens = _tokenize(line)
        if not tokens:
            continue
        rows.append(
            {
                "label": tokens[0].strip(),
                "label_raw": tokens[0],
                "raw": tokens,
            }
        )
    return rows


def _parse_section_sets(lines: List[str]) -> List[dict]:
    section_sets = []
    for line in lines:
        tokens = _tokenize(line)
        if not tokens:
            continue
        section_type = tokens[1].strip() if len(tokens) > 1 else None
        section_name = tokens[2].strip() if len(tokens) > 2 else None
        shape_class, shape_note = _shape_class(section_type)
        shape_warning = section_type is None
        section_sets.append(
            {
                "label": tokens[0].strip(),
                "label_raw": tokens[0],
                "section_type": section_type,
                "section_name": section_name,
                "shape_class": shape_class,
                "shape_note": shape_note,
                "shape_warning": shape_warning,
                "raw": tokens,
            }
        )
    return section_sets


def _parse_nodes(lines: List[str]) -> List[dict]:
    nodes = []
    for line in lines:
        tokens = _tokenize(line)
        if len(tokens) < 4:
            continue
        nodes.append(
            {
                "label": tokens[0].strip(),
                "label_raw": tokens[0],
                "x": _to_float(tokens[1]),
                "y": _to_float(tokens[2]),
                "z": _to_float(tokens[3]),
                "raw": tokens,
            }
        )
    return nodes


def _parse_members(
    lines: List[str],
    nodes: List[dict],
    section_sets: Dict[str, List[dict]],
    axis: Optional[str],
) -> List[dict]:
    members = []
    for line in lines:
        tokens = _tokenize(line)
        if len(tokens) < 5:
            continue
        i_index = _to_int(tokens[3])
        j_index = _to_int(tokens[4])
        section_index = _to_int(tokens[7]) if len(tokens) > 7 else None
        type_code = _to_int(tokens[10]) if len(tokens) > 10 else None

        i_node = _node_by_index(nodes, i_index)
        j_node = _node_by_index(nodes, j_index)
        length = _member_length(i_node, j_node)
        section_record = _section_set_record(type_code, section_index, section_sets)
        section_label = section_record.get("label") if section_record else None
        section_type = section_record.get("section_type") if section_record else None
        section_name = section_record.get("section_name") if section_record else None
        shape_class = section_record.get("shape_class") if section_record else None
        shape_note = section_record.get("shape_note") if section_record else None
        shape_warning = section_record.get("shape_warning") if section_record else True

        dx = _axis_component(i_node, j_node, "X")
        dy = _axis_component(i_node, j_node, "Y")
        dz = _axis_component(i_node, j_node, "Z")
        orientation = _member_orientation_deg(axis, dx, dy, dz)

        members.append(
            {
                "label": tokens[0].strip(),
                "label_raw": tokens[0],
                "i_node_index": i_index,
                "j_node_index": j_index,
                "i_node_label": i_node["label"] if i_node else None,
                "j_node_label": j_node["label"] if j_node else None,
                "section_set_index": section_index,
                "section_set_label": section_label,
                "section_type": section_type,
                "section_name": section_name,
                "type_code": type_code,
                "length": length,
                "dx": dx,
                "dy": dy,
                "dz": dz,
                "orientation_deg": orientation,
                "shape_class": shape_class,
                "shape_note": shape_note,
                "shape_warning": shape_warning,
                "raw": tokens,
            }
        )
    return members


def _tokenize(line: str) -> List[str]:
    try:
        tokens = shlex.split(line, posix=True)
    except ValueError:
        tokens = line.split()
    return [tok.rstrip(";") for tok in tokens if tok.strip()]


def _first_tokens(lines: List[str]) -> List[str]:
    for line in lines:
        tokens = _tokenize(line)
        if tokens:
            return tokens
    return []


def _token_int(tokens: List[str], idx: int) -> Optional[int]:
    if idx >= len(tokens):
        return None
    try:
        return int(float(tokens[idx]))
    except ValueError:
        return None


def _to_float(value: str) -> Optional[float]:
    try:
        return float(value)
    except ValueError:
        return None


def _to_int(value: str) -> Optional[int]:
    try:
        return int(float(value))
    except ValueError:
        return None


def _node_by_index(nodes: List[dict], idx: Optional[int]) -> Optional[dict]:
    if idx is None:
        return None
    if 1 <= idx <= len(nodes):
        return nodes[idx - 1]
    if 0 <= idx < len(nodes):
        return nodes[idx]
    return None


def _member_length(i_node: Optional[dict], j_node: Optional[dict]) -> Optional[float]:
    if not i_node or not j_node:
        return None
    if i_node["x"] is None or j_node["x"] is None:
        return None
    if i_node["y"] is None or j_node["y"] is None:
        return None
    if i_node["z"] is None or j_node["z"] is None:
        return None
    dx = j_node["x"] - i_node["x"]
    dy = j_node["y"] - i_node["y"]
    dz = j_node["z"] - i_node["z"]
    return (dx * dx + dy * dy + dz * dz) ** 0.5


def _section_set_record(
    type_code: Optional[int],
    section_index: Optional[int],
    section_sets: Dict[str, List[dict]],
) -> Optional[dict]:
    if section_index is None:
        return None
    if type_code == 1:
        sets = section_sets.get("hr", [])
    elif type_code == 2:
        sets = section_sets.get("cf", [])
    elif type_code == 6:
        sets = section_sets.get("al", [])
    else:
        sets = section_sets.get("general", [])

    if not sets:
        return None

    if 0 <= section_index < len(sets):
        return sets[section_index]
    if 1 <= section_index <= len(sets):
        return sets[section_index - 1]
    return None


def _axis_component(
    i_node: Optional[dict],
    j_node: Optional[dict],
    axis: str,
) -> Optional[float]:
    if not i_node or not j_node:
        return None
    if axis == "X":
        if i_node["x"] is None or j_node["x"] is None:
            return None
        return j_node["x"] - i_node["x"]
    if axis == "Y":
        if i_node["y"] is None or j_node["y"] is None:
            return None
        return j_node["y"] - i_node["y"]
    if axis == "Z":
        if i_node["z"] is None or j_node["z"] is None:
            return None
        return j_node["z"] - i_node["z"]
    return None


def _member_orientation_deg(
    axis: Optional[str],
    dx: Optional[float],
    dy: Optional[float],
    dz: Optional[float],
) -> Optional[float]:
    if dx is None or dy is None or dz is None:
        return None
    vertical = _vertical_axis(axis)
    if vertical == "X":
        h1, h2 = dy, dz
    elif vertical == "Y":
        h1, h2 = dx, dz
    else:
        h1, h2 = dx, dy

    if h1 == 0 and h2 == 0:
        return 0.0

    import math

    angle = math.degrees(math.atan2(h2, h1))
    if angle < 0:
        angle += 360.0
    return angle


def _vertical_axis(axis: Optional[str]) -> str:
    if axis and len(axis) >= 1:
        upper = axis.upper()
        if upper[0] in {"X", "Y", "Z"}:
            return upper[0]
    return "Z"


def _shape_class(section_type: Optional[str]) -> tuple[Optional[str], Optional[str]]:
    if not section_type:
        return None, None
    upper = section_type.strip().upper()
    if upper in {"BAR", "HSS PIPE A1085", "HSS PIPE", "PIPE"}:
        return "Round", None
    if upper in {"SQUARETUBE", "SQUARETUBE A1085"}:
        return "HSS Flat", None
    if upper == "NONE":
        return "Unknown", "Section set type is None; shape needs input."
    return "Flat", None


def _parse_units(lines: List[str]) -> dict:
    tokens = _first_tokens(lines)
    length_code = _token_int(tokens, 1)
    force_code = _token_int(tokens, 4)
    linear_force_code = _token_int(tokens, 5)
    moment_code = _token_int(tokens, 7)
    area_code = _token_int(tokens, 9)

    return {
        "length": _map_length(length_code),
        "force": _map_force(force_code),
        "linear_force": _map_linear_force(linear_force_code),
        "moment": _map_moment(moment_code),
        "area_load": _map_area_load(area_code),
    }


def _parse_axis(lines: List[str]) -> Optional[str]:
    tokens = _first_tokens(lines)
    axis_code = _token_int(tokens, 6)
    return {1: "XZY", 2: "YXZ", 3: "ZYX"}.get(axis_code)


def _parse_design_code(lines: List[str]) -> Optional[str]:
    tokens = _first_tokens(lines)
    code = _token_int(tokens, 0)
    return {
        0: "None",
        406: "AISC 15th (360-16): ASD",
        405: "AISC 15th (360-16): LRFD",
        404: "AISC 14th (360-10): ASD",
        403: "AISC 14th (360-10): LRFD",
        304: "AISC 13th (360-05): ASD",
        303: "AISC 13th (360-05): LRFD",
        3: "AISC 3rd: LRFD",
        2: "AISC 2nd: LRFD",
        9: "AISC 9th: ASD",
    }.get(code)


def _parse_version(lines: List[str]) -> Optional[str]:
    tokens = _first_tokens(lines)
    if not tokens:
        return None
    return tokens[0].replace(";", "")


def _map_length(code: Optional[int]) -> Optional[str]:
    return {
        0: "ft",
        1: "in",
        2: "m",
        3: "cm",
        4: "mm",
    }.get(code)


def _map_force(code: Optional[int]) -> Optional[str]:
    return {
        0: "kip",
        1: "lbf",
        2: "kN",
        3: "N",
        5: "kg",
    }.get(code)


def _map_linear_force(code: Optional[int]) -> Optional[str]:
    return {
        0: "klf",
        1: "kli",
        2: "plf",
        3: "pli",
        4: "kN/m",
        5: "kN/cm",
        6: "kN/mm",
        7: "N/m",
        8: "N/cm",
        9: "N/mm",
        13: "kg/m",
        14: "kg/cm",
        15: "kg/mm",
    }.get(code)


def _map_moment(code: Optional[int]) -> Optional[str]:
    return {
        0: "kip-ft",
        1: "kip-in",
        2: "lbf-ft",
        3: "lbf-in",
        4: "kN-m",
        5: "kN-cm",
        6: "kN-mm",
        7: "N-m",
        8: "N-cm",
        9: "N-mm",
        13: "kg-m",
        14: "kg-cm",
        15: "kg-mm",
    }.get(code)


def _map_area_load(code: Optional[int]) -> Optional[str]:
    return {
        0: "ksf",
        1: "ksi",
        2: "psf",
        3: "psi",
        4: "MPa",
        5: "kPa",
        6: "Pa",
        8: "kg/m^2",
        9: "kg/mm^2",
    }.get(code)
