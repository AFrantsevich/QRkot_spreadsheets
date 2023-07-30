from datetime import datetime

from sqlalchemy import asc, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.models.charity_project import CharityProject
from app.models.donation import Donation
from app.schemas.charity_project import CharityProjectCreate
from app.schemas.donation import DonationCreate


async def check_projects_for_charity(
        donation: DonationCreate,
        session: AsyncSession,
) -> DonationCreate:
    invested_amount_for_charity = donation.full_amount
    flag = False
    charity_project_for_charity = await session.execute(
        select(CharityProject).where(
            CharityProject.fully_invested == flag
        ).order_by(CharityProject.create_date)
    )

    charity_project_for_charity = charity_project_for_charity.scalars().all()
    for project in charity_project_for_charity:
        project_avail_for_charity = project.full_amount - project.invested_amount
        if invested_amount_for_charity < project_avail_for_charity:
            project.invested_amount += invested_amount_for_charity
            donation.invested_amount += invested_amount_for_charity
            donation.fully_invested = True
            donation.close_date = datetime.now()
            break
        project.invested_amount = project.full_amount
        invested_amount_for_charity -= project_avail_for_charity
        donation.invested_amount += project_avail_for_charity
        project.fully_invested = True
        project.close_date = datetime.now()
    return donation


async def check_donations_for_charity(
        charity_project: CharityProjectCreate,
        session: AsyncSession,
) -> CharityProjectCreate:
    need_charity_for_project = charity_project.full_amount
    flag = False
    donation_for_charity = await session.execute(
        select(Donation).where(
            Donation.fully_invested == flag
        ).order_by(asc(Donation.invested_amount))
    )

    donation_for_charity = donation_for_charity.scalars().all()
    for donation in donation_for_charity:
        avail_amount = donation.full_amount - donation.invested_amount
        if need_charity_for_project <= avail_amount:
            charity_project.invested_amount = charity_project.full_amount
            donation.invested_amount += need_charity_for_project
            charity_project.fully_invested = True
            charity_project.close_date = datetime.now()
            break
        charity_project.invested_amount += avail_amount
        donation.invested_amount += avail_amount
        donation.fully_invested = True
        donation.close_date = datetime.now()
    return charity_project
