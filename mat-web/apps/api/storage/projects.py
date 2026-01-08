import json
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

ROOT_DIR = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT_DIR / "data"
PROJECTS_DIR = DATA_DIR / "projects"


def _ensure_dirs() -> None:
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)


def _project_dir(project_id: str) -> Path:
    return PROJECTS_DIR / project_id


def _project_file(project_id: str) -> Path:
    return _project_dir(project_id) / "project.json"


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def create_project(name: str) -> dict:
    _ensure_dirs()
    project_id = uuid4().hex
    project_dir = _project_dir(project_id)
    project_dir.mkdir(parents=True, exist_ok=True)
    project = {
        "id": project_id,
        "name": name,
        "created_at": _utc_now(),
        "updated_at": _utc_now(),
        "code": {},
        "geometry": {},
        "files": {},
    }
    save_project(project)
    return project


def load_project(project_id: str) -> dict:
    project_file = _project_file(project_id)
    if not project_file.exists():
        raise FileNotFoundError(f"Project not found: {project_id}")
    return json.loads(project_file.read_text(encoding="utf-8"))


def save_project(project: dict) -> None:
    _ensure_dirs()
    project["updated_at"] = _utc_now()
    project_file = _project_file(project["id"])
    project_file.write_text(json.dumps(project, indent=2), encoding="utf-8")


def project_dir(project_id: str) -> Path:
    _ensure_dirs()
    path = _project_dir(project_id)
    path.mkdir(parents=True, exist_ok=True)
    return path
