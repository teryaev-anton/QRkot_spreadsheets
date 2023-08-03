from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Extra, Field, PositiveInt


class CharityProjectBase(BaseModel):
    """Базовый класс для схем проекта"""
    name: Optional[str] = Field(None, max_length=100)
    description: Optional[str]
    full_amount: PositiveInt

    class Config:
        extra = Extra.forbid
        min_anystr_length = 1


class CharityProjectCreate(CharityProjectBase):
    """Схема для создания проекта"""
    name: str = Field(..., max_length=100)
    description: str


class CharityProjectUpdate(CharityProjectBase):
    """Схема для изменения (обновления) проекта"""
    full_amount: Optional[PositiveInt]


class CharityProjectDB(CharityProjectBase):
    """Схема для получения проекта из БД"""
    id: int
    invested_amount: int
    fully_invested: bool
    create_date: datetime
    close_date: datetime = Field(None)

    class Config:
        orm_mode = True
