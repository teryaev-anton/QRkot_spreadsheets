from sqlalchemy import Column, ForeignKey, Integer, Text

from .base import ProjectDonationBaseModel


class Donation(ProjectDonationBaseModel):
    """Модель для пожертвования"""
    user_id = Column(Integer, ForeignKey('user.id'))
    comment = Column(Text)
