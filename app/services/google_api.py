from datetime import datetime
from typing import List

from aiogoogle import Aiogoogle

from app.core.config import settings


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    """Cоздание гугл-таблицы"""
    now_date_time = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheets_body = {
        'properties': {'title': f'QRKot_Отчет на {now_date_time}',
                       'locale': 'ru_RU'},
        'sheets': [{
            'properties': {'sheetType': 'GRID',
                           'sheetId': 0,
                           'title': 'Топ проектов по времени сбора',
                           'gridProperties': {'rowCount': 100,
                                              'columnCount': 3}
                           }
        }]
    }
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheets_body)
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
    spreadsheetid: str,
    wrapper_services: Aiogoogle
) -> None:
    """Выдача прав личному аккаунту Google"""
    permissions_body = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': settings.email
    }
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permissions_body,
            fields='id'
        )
    )


async def spreadsheets_update_value(
    spreadsheetid: str,
    projects: List,
    wrapper_services: Aiogoogle
) -> None:
    """Обновление гугл-таблицы"""
    now_date_time = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = [
        ['Отчет от', now_date_time],
        ['Топ проектов по скорости закрытия'],
        ['Название проекта', 'Время сбора', 'Описание']
    ]
    for project in projects:
        new_row = [
            str(project['name']),
            str(project['period']),
            str(project['description'])
        ]
        table_values.append(new_row)
    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    all_lines = len(table_values)
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range=f'A1:C{all_lines}',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
