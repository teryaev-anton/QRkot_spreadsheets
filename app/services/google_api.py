from copy import deepcopy
from datetime import datetime
from typing import List

from aiogoogle import Aiogoogle

from app.core.config import settings

FORMAT = '%Y/%m/%d %H:%M:%S'
TITLE = 'QRKot_Отчет на {now_date_time}'
ROW_COUNT = 100
COLUMN_COUNT = 3
SHEET_ID = 0
SHEET_TITLE = 'Топ проектов по времени сбора'
SPREADSHEET_BODY = dict(
    properties=dict(
        title=TITLE.format(now_date_time=datetime.now().strftime(FORMAT)),
        locale='ru_RU',
    ),
    sheets=[dict(properties=dict(
        sheetType='GRID',
        sheetId=SHEET_ID,
        title=SHEET_TITLE,
        gridProperties=dict(
            rowCount=ROW_COUNT,
            columnCount=COLUMN_COUNT,
        )
    ))]
)
TABLE_HEADER = [
    ['Отчет от', datetime.now().strftime(FORMAT)],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    """Cоздание гугл-таблицы"""
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheets_body = deepcopy(SPREADSHEET_BODY)
    spreadsheets_body['properties']['title'] = TITLE.format(
        now_date_time=datetime.now().strftime(FORMAT)
    )
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheets_body)
    )
    spreadsheet_id = response['spreadsheetId']
    return spreadsheet_id


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
    service = await wrapper_services.discover('sheets', 'v4')
    table_header = deepcopy(TABLE_HEADER)
    table_header[0][1] = datetime.now().strftime(FORMAT)
    table_values = [
        *table_header,
        *[list(map(str, [
            attr.name, attr.close_date - attr.create_date, attr.description
        ])) for attr in projects],
    ]
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
