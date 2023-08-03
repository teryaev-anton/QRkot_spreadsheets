from typing import Dict, List, Optional

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from .base import CRUDBase
from app.models import CharityProject


class CRUDCharityProject(CRUDBase):
    """Класс для CRUD-операций для CharityProject"""
    async def get_project_id_by_name(
            self,
            project_name: str,
            session: AsyncSession
    ) -> Optional[int]:
        """Поиск в БД проекта по имени. Возвращает ID проекта"""
        project_id = await session.execute(
            select(CharityProject.id).where(
                CharityProject.name == project_name
            )
        )
        return project_id.scalars().first()

    async def get_projects_by_completion_rate(
        self,
        session: AsyncSession
    ) -> List[Dict[str, str]]:
        """Получение списка проектов отсортированных по времени сбора"""
        projects = await session.execute(
            select([CharityProject]).where(CharityProject.fully_invested)
        )
        projects = projects.scalars().all()
        project_list = []
        for project in projects:
            project_list.append({
                'name': project.name,
                'period': project.close_date - project.create_date,
                'description': project.description
            })
        project_list = sorted(project_list, key=lambda x: x['period'])
        return project_list


charity_project_crud = CRUDCharityProject(CharityProject)
