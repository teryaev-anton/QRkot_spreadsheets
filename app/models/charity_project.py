from sqlalchemy import Column, String, Text

from .base import ProjectDonationBaseModel


class CharityProject(ProjectDonationBaseModel):
    """Модель для проекта"""
    name = Column(String(100), unique=True, nullable=False)
    description = Column(Text, nullable=False)
