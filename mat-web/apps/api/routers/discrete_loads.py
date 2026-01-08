from typing import Any, Dict, List, Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field

from services.discrete_loads import apply_cfd_cci, normalize_rows, validate_rows
from services.discrete_outputs import compute_outputs
from storage import projects as project_store

router = APIRouter(prefix="/discrete-loads", tags=["discrete-loads"])


class DiscreteLoadsPayload(BaseModel):
    project_id: str
    qty_sectors: int = Field(default=3, ge=1, le=4)
    mount_azimuths: Dict[str, Any] = Field(default_factory=dict)
    rows: List[Dict[str, Any]] = Field(default_factory=list)
    ice_thickness: Optional[float] = None


@router.post("")
def set_discrete_loads(payload: DiscreteLoadsPayload):
    try:
        project = project_store.load_project(payload.project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    validation = validate_rows(payload.rows, payload.qty_sectors)
    normalized = normalize_rows(payload.rows, payload.mount_azimuths, payload.qty_sectors)
    cfd_rows = project.get("cfd_db", [])
    cci_rows = project.get("cci_db", [])
    ice_thickness = payload.ice_thickness
    if ice_thickness is None:
        ice_thickness = _to_float(project.get("code", {}).get("inputs", {}).get("H44"))
    normalized = apply_cfd_cci(normalized, cfd_rows, cci_rows, ice_thickness or 0.0)

    project["discrete_loads"] = {
        "qty_sectors": payload.qty_sectors,
        "mount_azimuths": payload.mount_azimuths,
        "rows": payload.rows,
        "normalized": normalized,
        "validation": validation,
        "ice_thickness": ice_thickness,
    }
    project_store.save_project(project)

    return {
        "project_id": payload.project_id,
        "validation": validation,
        "normalized_count": len(normalized),
    }


@router.post("/calc/{project_id}")
def calculate_discrete_loads(project_id: str):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    rows = project.get("discrete_loads", {}).get("normalized", [])
    ice_thickness = project.get("discrete_loads", {}).get("ice_thickness", None)
    cfd_rows = project.get("cfd_db", [])
    cci_rows = project.get("cci_db", [])
    code_inputs = project.get("code", {}).get("inputs", {})
    code_derived = project.get("code", {}).get("derived", {})
    code_inputs = {**code_derived, **code_inputs}
    if ice_thickness is None:
        ice_thickness = _to_float(code_inputs.get("H44"))

    geometry = project.get("geometry", {})
    outputs = compute_outputs(rows, cfd_rows, cci_rows, ice_thickness or 0.0, code_inputs, geometry)
    project.setdefault("discrete_loads", {})["outputs"] = outputs
    project_store.save_project(project)

    return {
        "project_id": project_id,
        "row_outputs": len(outputs.get("row_outputs", [])),
        "code_tables": {
            "wind": len(outputs.get("code_tables", {}).get("wind", [])),
            "ice": len(outputs.get("code_tables", {}).get("ice", [])),
            "placement": len(outputs.get("code_tables", {}).get("placement", [])),
        },
    }


@router.get("/{project_id}")
def get_discrete_loads(project_id: str):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    return project.get("discrete_loads", {})


def _to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError:
        return None
