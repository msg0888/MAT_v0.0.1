from pathlib import Path
from typing import Any, Dict

from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from parsers.arc import parse_arc
from parsers.tml import parse_tml
from services.discrete_loads import apply_import_rows
from services.tml_mapper import map_tml_records
from storage import projects as project_store

router = APIRouter(prefix="/imports", tags=["imports"])


@router.post("/arc")
async def import_arc(
    arc_file: UploadFile = File(...),
):
    content = await arc_file.read()
    try:
        text = content.decode("utf-8", errors="ignore")
    except UnicodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid file encoding") from exc

    entries = parse_arc(text)
    manufacturer = Path(arc_file.filename or "").stem

    return {
        "manufacturer": manufacturer,
        "entries": entries,
        "count": len(entries),
        "file": arc_file.filename,
    }


@router.post("/tml")
async def import_tml(
    project_id: str = Form(...),
    tml_file: UploadFile = File(...),
):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    content = await tml_file.read()
    parsed = parse_tml(content)
    mapped = map_tml_records(parsed)

    project.setdefault("imports", {})["tml"] = {
        "counts": mapped["counts"],
        "file": tml_file.filename,
    }

    current = project.get("discrete_loads", {}).get("rows", [])
    merged = apply_import_rows(current, mapped["rows"])
    project.setdefault("discrete_loads", {})["rows"] = merged
    project_store.save_project(project)

    return {
        "project_id": project_id,
        "counts": mapped["counts"],
        "rows_added": len(mapped["rows"]),
        "total_rows": len(merged),
    }


@router.post("/cfd-db")
async def import_cfd_db(
    project_id: str = Form(...),
    cfd_file: UploadFile = File(...),
):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    content = await cfd_file.read()
    text = content.decode("utf-8", errors="ignore")
    rows = _parse_csv(text)
    project["cfd_db"] = rows
    project_store.save_project(project)
    return {"project_id": project_id, "count": len(rows)}


@router.post("/cci-db")
async def import_cci_db(
    project_id: str = Form(...),
    cci_file: UploadFile = File(...),
):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    content = await cci_file.read()
    text = content.decode("utf-8", errors="ignore")
    rows = _parse_csv(text)
    project["cci_db"] = rows
    project_store.save_project(project)
    return {"project_id": project_id, "count": len(rows)}


def _parse_csv(text: str) -> list[dict]:
    import csv
    from io import StringIO

    reader = csv.DictReader(StringIO(text))
    rows = []
    for row in reader:
        rows.append({k.strip(): v for k, v in row.items()})
    return rows
