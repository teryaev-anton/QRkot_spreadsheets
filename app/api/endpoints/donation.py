from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_superuser, current_user
from app.crud.donation import donation_crud
from app.models import User
from app.schemas.donation import (
    DonationCreate, DonationDBFull, DonationDBPartial
)
from app.services.distribute_for_projects import distribute_for_projects

router = APIRouter()


@router.get(
    '/',
    response_model=List[DonationDBFull],
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)]
)
async def get_all_donations(
    session: AsyncSession = Depends(get_async_session)
):
    """
    Получение списка всех пожертвований.
    Доступно только суперюзеру.
    """
    all_donations = await donation_crud.get_multi(session)
    return all_donations


@router.post(
    '/',
    response_model=DonationDBPartial,
    response_model_exclude_none=True
)
async def create_donation(
    donation: DonationCreate,
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user)
):
    """
    Создание пожертвования.
    Доступно только зарегистрированным пользователям.
    """
    donation = await donation_crud.create(donation, session, user)
    await distribute_for_projects(session)
    await session.refresh(donation)
    return donation


@router.get(
    '/my',
    response_model=List[DonationDBPartial],
    response_model_exclude_none=True
)
async def get_user_donations(
    session: AsyncSession = Depends(get_async_session),
    user: User = Depends(current_user)
):
    """
    Получение списка своих пожертвований.
    Доступно только зарегистрированным пользователям.
    """
    donations = await donation_crud.get_by_user(user, session)
    return donations
