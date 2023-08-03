from fastapi import APIRouter

from app.api.endpoints import (
    charity_project_router, donation_router, google_router, user_router
)

main_router = APIRouter()
main_router.include_router(
    charity_project_router,
    prefix='/charity_project',
    tags=['Проекты для поддержки котиков']
)
main_router.include_router(
    donation_router, prefix='/donation', tags=['Пожертвования']
)
main_router.include_router(user_router)

main_router.include_router(
    google_router,
    prefix='/google',
    tags=['Google']
)
