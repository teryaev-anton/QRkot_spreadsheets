from typing import List, Optional

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
    ) -> List[CharityProject]:
        """Получение списка проектов отсортированных по времени сбора"""
        projects = await session.execute(
            select([CharityProject]).where(CharityProject.fully_invested)
        )
        return sorted(
            projects.scalars().all(),
            key=lambda obj: obj.close_date - obj.create_date
        )


charity_project_crud = CRUDCharityProject(CharityProject)
