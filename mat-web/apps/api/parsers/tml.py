from typing import Any, Dict, List
import xml.etree.ElementTree as ET


def parse_tml(content: bytes) -> Dict[str, List[Dict[str, Any]]]:
    root = _parse_xml(content)

    return {
        "discrete_loads": _parse_records(root, "discreteLoad"),
        "dish_loads": _parse_records(root, "dish"),
        "feedline_loads": _parse_records(root, "feedline"),
    }


def _parse_xml(content: bytes) -> ET.Element:
    try:
        return ET.fromstring(content)
    except ET.ParseError:
        text = content.decode("utf-8", errors="ignore")
        return ET.fromstring(text)


def _parse_records(root: ET.Element, tag: str) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for elem in _find_all(root, tag):
        record: Dict[str, Any] = {}
        for child in list(elem):
            key = _strip_ns(child.tag)
            record[key] = (child.text or "").strip()
        records.append(record)
    return records


def _find_all(root: ET.Element, tag: str) -> List[ET.Element]:
    matches: List[ET.Element] = []
    for elem in root.iter():
        if _strip_ns(elem.tag) == tag:
            matches.append(elem)
    return matches


def _strip_ns(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag
