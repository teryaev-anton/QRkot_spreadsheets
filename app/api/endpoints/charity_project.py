from datetime import datetime
from typing import List

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.db import get_async_session
from app.core.user import current_superuser
from app.crud.charity_project import charity_project_crud
from app.schemas.charity_project import (CharityProjectCreate,
                                         CharityProjectDB,
                                         CharityProjectUpdate)
from app.api.validators import (check_charity_project_before_edit,
                                check_name_duplicate)
from app.services.distribute_for_projects import distribute_for_projects

router = APIRouter()


@router.get(
    '/',
    response_model=List[CharityProjectDB],
    response_model_exclude_none=True
)
async def get_all_charity_projects(
    session: AsyncSession = Depends(get_async_session)
):
    """Получение списка всех проектов"""
    all_projects = await charity_project_crud.get_multi(session)
    return all_projects


@router.post(
    '/',
    response_model=CharityProjectDB,
    response_model_exclude_none=True,
    dependencies=[Depends(current_superuser)]
)
async def create_charity_project(
    charity_project: CharityProjectCreate,
    session: AsyncSession = Depends(get_async_session)
):
    """
    Создание нового проекта.
    Доступно только суперюзеру.
    """
    await check_name_duplicate(charity_project.name, session)
    new_project = await charity_project_crud.create(charity_project, session)
    await distribute_for_projects(session)
    await session.refresh(new_project)
    return new_project


@router.patch(
    '/{project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)]
)
async def update_charity_project(
        project_id: int,
        obj_in: CharityProjectUpdate,
        session: AsyncSession = Depends(get_async_session),
):
    """
    Изменение (обновление) проекта.
    Доступно только суперюзеру.
    """
    if obj_in.full_amount:
        full_amount = obj_in.full_amount
    else:
        full_amount = None
    existing_project = await check_charity_project_before_edit(
        project_id, session, full_amount=full_amount
    )

    if obj_in.name:
        await check_name_duplicate(obj_in.name, session)

    if full_amount == existing_project.invested_amount:
        existing_project.full_amount = full_amount
        setattr(existing_project, 'fully_invested', True)
        setattr(existing_project, 'close_date', datetime.now())
    project = await charity_project_crud.update(
        existing_project, obj_in, session
    )
    return project


@router.delete(
    '/{project_id}',
    response_model=CharityProjectDB,
    dependencies=[Depends(current_superuser)]
)
async def delete_charity_project(
        project_id: int,
        session: AsyncSession = Depends(get_async_session)
):
    """
    Удаление проекта.
    Доступно только суперюзеру.
    """
    existing_project = await check_charity_project_before_edit(
        project_id, session, delete=True
    )
    deleted_project = await charity_project_crud.remove(
        existing_project, session
    )
    return deleted_project
