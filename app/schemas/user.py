from fastapi_users import schemas


class UserCreate(schemas.BaseUserCreate):
    """схема для создания пользователя"""
    pass


class UserRead(schemas.BaseUser[int]):
    """схема для получения пользователя (кроме пароля)"""
    pass


class UserUpdate(schemas.BaseUserUpdate):
    """схема для обновления пользователя"""
    pass