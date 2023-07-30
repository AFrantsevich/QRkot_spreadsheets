from fastapi import FastAPI

from api.routers import main_router
from app.core.init_db import create_first_superuser
from core.config import settings

app = FastAPI(title=settings.app_title, description=settings.description)

app.include_router(main_router)


@app.on_event('startup')
async def startup():
    await create_first_superuser()
