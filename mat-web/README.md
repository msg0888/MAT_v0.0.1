MAT Web (local-first)

This repo will mirror MAT v3.3.0 workflows with a web UI and a Python API.

Structure
- apps/api  : FastAPI backend
- apps/web  : React UI (pending Node install)
- docs      : MVP scope and data model

Run API (local)
- cd apps/api
- python -m venv .venv
- .venv\Scripts\activate
- pip install -r requirements.txt
- uvicorn main:app --reload

Geometry import (example)
- POST /geometry/import (multipart form)
  - project_id: <id>
  - r3d_file: <file>

Code inputs (example)
- POST /code/inputs (json)
  - project_id: <id>
  - inputs: { "C2": "...", "D2": "...", "CenterPoint": "N1", ... }

Discrete Loads (example)
- POST /discrete-loads (json)
  - project_id: <id>
  - qty_sectors: 3
  - mount_azimuths: { "alpha": 0, "beta": 120, "gamma": 240 }
  - rows: [ { "A": "...", "B": "...", "C": "...", "L": 0, "M": 1, ... } ]
- POST /discrete-loads/calc/{project_id}

Imports
- POST /imports/arc (multipart form)
  - arc_file: <file>
- POST /imports/tml (multipart form)
  - project_id: <id>
  - tml_file: <file>
- POST /imports/cfd-db (multipart form)
  - project_id: <id>
  - cfd_file: <csv>
- POST /imports/cci-db (multipart form)
  - project_id: <id>
  - cci_file: <csv>
