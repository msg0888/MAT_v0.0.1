from typing import Any, Dict, List, Optional


def parse_arc(text: str) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    current: Dict[str, Any] = {}

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        lower = line.lower()
        if lower.startswith("usname="):
            if current:
                entries.append(current)
            current = {"us_name": line[7:].strip()}
            continue
        if lower.startswith("siname="):
            current["si_name"] = line[7:].strip()
            continue
        if lower.startswith("values="):
            values = line[7:].strip().split()
            current["values"] = values
            current["values_count"] = len(values)
            continue

    if current:
        entries.append(current)

    return entries
