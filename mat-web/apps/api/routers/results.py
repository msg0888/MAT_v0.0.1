from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from parsers.risa_results import parse_risa_results
from services.results_summary import summarize_envelope
from storage import projects as project_store

router = APIRouter(prefix="/results", tags=["results"])


@router.post("/import")
async def import_results(
    project_id: str = Form(...),
    results_file: UploadFile = File(...),
):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    content = await results_file.read()
    try:
        text = content.decode("utf-8", errors="ignore")
    except UnicodeError as exc:
        raise HTTPException(status_code=400, detail="Invalid file encoding") from exc

    parsed = parse_risa_results(text)
    section_map = {
        member.get("label"): member.get("section_set_label")
        for member in project.get("geometry", {}).get("members", [])
        if member.get("label")
    }
    for row in parsed.get("rows", []):
        row["section_set"] = section_map.get(row.get("member"))
    summary = summarize_envelope(parsed.get("rows", []), project.get("geometry", {}))

    proj_dir = project_store.project_dir(project_id)
    inputs_dir = proj_dir / "inputs"
    inputs_dir.mkdir(parents=True, exist_ok=True)
    file_name = results_file.filename or "results.txt"
    file_path = inputs_dir / file_name
    file_path.write_bytes(content)

    project.setdefault("results", {})["envelope"] = parsed
    project["results"]["summary"] = summary
    project.setdefault("files", {})["results_flat"] = str(file_path)
    project_store.save_project(project)

    return {
        "project_id": project_id,
        "count": parsed.get("count", 0),
        "file": str(file_path),
    }


@router.get("/{project_id}")
def get_results(project_id: str):
    try:
        project = project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc

    return project.get("results", {})
