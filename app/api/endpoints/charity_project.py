from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.api.validators import (
    check_charity_project_duplicate,
    check_charity_project_exists_and_invested_amount)
from app.core.db import get_async_session
from app.core.user import current_superuser
from app.crud.charity_project import charity_project_crud
from app.schemas.charity_project import (CharityProjectCreate,
                                         CharityProjectDB,
                                         CharityProjectUpdate)
from services.investing import check_donations_for_charity

router = APIRouter()


@router.post('/',
             response_model=CharityProjectDB,
             response_model_exclude_none=True,
             dependencies=[Depends(current_superuser)]
             )
async def create_new_charity_project(
        charity_project: CharityProjectCreate,
        session: AsyncSession = Depends(get_async_session)
):
    await check_charity_project_duplicate(charity_project.name, session)
    return await charity_project_crud.create(
        await check_donations_for_charity(charity_project, session),
        session)


@router.delete(
    '/{project_id}',
    response_model=CharityProjectDB,
    response_model_exclude_none=False,
    dependencies=[Depends(current_superuser)]
)
async def remove_charity_project(
        project_id: int,
        session: AsyncSession = Depends(get_async_session)
):
    return await charity_project_crud.remove(
        await check_charity_project_exists_and_invested_amount(
            project_id, session, flag=True),
        session)


@router.patch(
    '/{project_id}',
    response_model=CharityProjectDB,
    response_model_exclude_none=False,
    dependencies=[Depends(current_superuser)],
)
async def partially_update_charity_project(
        project_id: int,
        obj_in: CharityProjectUpdate,
        session: AsyncSession = Depends(get_async_session),
):
    if obj_in.name is not None:
        await check_charity_project_duplicate(obj_in.name, session)
    return await charity_project_crud.update(
        await check_charity_project_exists_and_invested_amount(
            project_id, session, obj_in),
        obj_in,
        session)


@router.get(
    '/',
    response_model=List[CharityProjectDB],
    response_model_exclude_none=True,
)
async def get_all_charity_projects(
        session: AsyncSession = Depends(get_async_session),
):
    return await charity_project_crud.get_multi(session)
