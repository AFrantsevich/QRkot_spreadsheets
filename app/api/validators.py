from http import HTTPStatus

from fastapi import HTTPException
from sqlalchemy.ext.asyncio import AsyncSession

from app.crud.charity_project import charity_project_crud
from app.models.charity_project import CharityProject


async def check_charity_project_duplicate(
        charity_project_name: str,
        session: AsyncSession,
) -> None:
    charity_project_id = await charity_project_crud.get_charity_project_id_by_name(
        charity_project_name, session)
    if charity_project_id is not None:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Проект с таким именем уже существует!'
        )


async def check_charity_project_exists_and_invested_amount(
        charity_project_id: int,
        session: AsyncSession,
        obj_in: CharityProject = None,
        flag: bool = False,
) -> CharityProject:
    charity_project = await charity_project_crud.get(charity_project_id, session)
    if charity_project is None:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail='Проект не найден!'
        )
    if charity_project.invested_amount != 0 and flag:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='В проект были внесены средства, не подлежит удалению!'
        )
    if charity_project.fully_invested:
        raise HTTPException(
            status_code=HTTPStatus.BAD_REQUEST,
            detail='Закрытый проект нельзя редактировать!'
        )
    if obj_in and obj_in.full_amount:
        if obj_in.full_amount < charity_project.invested_amount:
            raise HTTPException(
                status_code=HTTPStatus.BAD_REQUEST,
                detail='Нельзя вносить'
            )
    return charity_project


async def check_charity_close_project(
        charity_project_id: int,
        session: AsyncSession,
) -> None:
    charity_project = await charity_project_crud.get(charity_project_id, session)
    if charity_project is None:
        raise HTTPException(
            status_code=HTTPStatus.NOT_FOUND,
            detail='Проект не найден!'
        )
