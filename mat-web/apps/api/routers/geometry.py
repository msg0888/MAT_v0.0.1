from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from parsers.r3d import parse_r3d
from storage import projects as project_store

router = APIRouter(prefix="/geometry", tags=["geometry"])


@router.post("/import")
async def import_geometry(
    project_id: str = Form(...),
    r3d_file: UploadFile = File(...),
):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    content = await r3d_file.read()
    try:
        text = content.decode("utf-8", errors="ignore")
    except UnicodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid file encoding") from exc

    result = parse_r3d(text)

    proj_dir = project_store.project_dir(project_id)
    inputs_dir = proj_dir / "inputs"
    inputs_dir.mkdir(parents=True, exist_ok=True)
    file_name = r3d_file.filename or "model.r3d"
    file_path = inputs_dir / file_name
    file_path.write_bytes(content)

    project["geometry"] = result
    project.setdefault("files", {})["model_r3d"] = str(file_path)
    project_store.save_project(project)

    counts = {
        "members": len(result.get("members", [])),
        "nodes": len(result.get("nodes", [])),
        "section_sets": {
            key: len(val) for key, val in result.get("section_sets", {}).items()
        },
    }

    return {
        "project_id": project_id,
        "counts": counts,
        "units": result.get("units"),
        "axis": result.get("axis"),
        "design_code": result.get("design_code"),
        "risa_version": result.get("risa_version"),
        "flags": result.get("flags", {}),
        "file": str(file_path),
    }
