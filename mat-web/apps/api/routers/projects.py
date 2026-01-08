from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from storage import projects as project_store

router = APIRouter(prefix="/projects", tags=["projects"])

class ProjectCreate(BaseModel):
    name: str

@router.post("")
def create_project(payload: ProjectCreate):
    return project_store.create_project(payload.name)


@router.get("/{project_id}")
def get_project(project_id: str):
    try:
        return project_store.load_project(project_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
