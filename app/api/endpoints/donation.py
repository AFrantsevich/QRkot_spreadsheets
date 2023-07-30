from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_superuser, current_user
from app.crud.donation import donation_crud
from app.models.user import User
from app.schemas.donation import (DonationCreate, DonationSuperUserDB,
                                  DonationUserDB)
from app.services.investing import check_projects_for_charity

router = APIRouter()


@router.post('/',
             response_model=DonationUserDB,
             response_model_exclude_none=True,
             )
async def create_new_donation(
        donation: DonationCreate,
        session: AsyncSession = Depends(get_async_session),
        user: User = Depends(current_user)
):
    return await donation_crud.create(
        await check_projects_for_charity(donation, session),
        session,
        user)


@router.get(
    '/my',
    response_model=List[DonationUserDB],
    response_model_exclude={'user_id'},
)
async def get_all_user_donation(
        session: AsyncSession = Depends(get_async_session),
        user: User = Depends(current_user)
):
    return await donation_crud.get_by_user(session, user=user)


@router.get(
    '/',
    response_model=List[DonationSuperUserDB],
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)]
)
async def get_all_charity_projects(
        session: AsyncSession = Depends(get_async_session),
):
    return await donation_crud.get_multi(session)
