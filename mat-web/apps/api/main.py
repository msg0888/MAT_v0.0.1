from fastapi import FastAPI

from routers import code, discrete_loads, geometry, health, imports, projects, results

app = FastAPI(title="MAT Web API", version="0.1.0")

app.include_router(health.router)
app.include_router(projects.router)
app.include_router(geometry.router)
app.include_router(code.router)
app.include_router(discrete_loads.router)
app.include_router(imports.router)
app.include_router(results.router)
