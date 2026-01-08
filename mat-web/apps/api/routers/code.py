from typing import Any, Dict, List

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from services.tia_tables import derive_inputs as derive_tia_inputs
from storage import projects as project_store

router = APIRouter(prefix="/code", tags=["code"])


class CodeInputs(BaseModel):
    project_id: str
    inputs: Dict[str, Any] = Field(default_factory=dict)


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _require(inputs: Dict[str, Any], keys: List[str], errors: List[str]) -> None:
    for key in keys:
        if _is_missing(inputs.get(key)):
            errors.append(f"Missing required input: {key}")


def _range_keys(col_start: str, col_end: str, row_start: int, row_end: int) -> List[str]:
    keys = []
    for row in range(row_start, row_end + 1):
        keys.append(f"{col_start}{row}")
        if col_end != col_start:
            keys.append(f"{col_end}{row}")
    return keys


def _derive_wind_speed(inputs: Dict[str, Any]) -> Dict[str, Any]:
    code_type = str(inputs.get("C19", "")).strip().upper()
    standard = str(inputs.get("D19", "")).strip().upper()
    if code_type == "TIA" and standard == "SA":
        return {"wind_speed_label": "Service Wind Speed, Vs", "wind_speed_default": 60}
    return {"wind_speed_label": "Maintenance Wind Speed, Vm", "wind_speed_default": 30}


def validate_code_inputs(inputs: Dict[str, Any]) -> Dict[str, Any]:
    errors: List[str] = []

    _require(inputs, _range_keys("C", "D", 2, 7), errors)
    _require(inputs, _range_keys("C", "D", 16, 19), errors)
    _require(inputs, _range_keys("C", "D", 21, 22), errors)
    _require(inputs, _range_keys("C", "D", 30, 31), errors)
    _require(inputs, _range_keys("C", "D", 41, 44), errors)
    _require(inputs, _range_keys("H", "I", 16, 21), errors)
    _require(inputs, _range_keys("H", "I", 23, 25), errors)
    _require(inputs, ["CenterPoint"], errors)

    code_version = str(inputs.get("C18", "")).strip()
    if code_version == "ANSI/TIA-222-I":
        _require(inputs, _range_keys("C", "D", 24, 28), errors)
        _require(inputs, _range_keys("C", "D", 33, 37), errors)
    else:
        _require(inputs, _range_keys("C", "D", 24, 27), errors)
        _require(inputs, _range_keys("C", "D", 33, 35), errors)

    if str(inputs.get("H24", "")).strip().upper() == "ROOFTOP":
        _require(inputs, _range_keys("H", "I", 27, 32), errors)

    derived = _derive_wind_speed(inputs)

    wind_speed_value = inputs.get("C23")
    if _is_missing(wind_speed_value):
        derived["wind_speed_value"] = derived["wind_speed_default"]
        derived["wind_speed_note"] = "Applied default based on Code inputs."

    return {"valid": len(errors) == 0, "errors": errors, "derived": derived}


@router.post("/inputs")
def set_code_inputs(payload: CodeInputs):
    try:
        project = project_store.load_project(payload.project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    result = validate_code_inputs(payload.inputs)
    tia_derived = derive_tia_inputs(payload.inputs)
    derived = dict(result["derived"])
    derived.update(tia_derived)
    project.setdefault("code", {})["inputs"] = payload.inputs
    project["code"]["derived"] = derived
    project["code"]["valid"] = result["valid"]
    project["code"]["errors"] = result["errors"]
    project_store.save_project(project)

    return result
