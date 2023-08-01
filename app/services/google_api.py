from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings

from aiogoogle.excs import HTTPError

FORMAT = "%Y/%m/%d %H:%M:%S"

ROW_COUNT = 100
COLUMN_COUNT = 11
NOW_DATE_TIME = datetime.now().strftime(FORMAT)

TABLE_VALUES = [
    ['Отчет от', NOW_DATE_TIME],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]


def spreadsheet_body():
    return {
        'properties': {'title': f'Отчет от {NOW_DATE_TIME}',
                       'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист1',
                                   'gridProperties': {'rowCount': ROW_COUNT,
                                                      'columnCount': COLUMN_COUNT}}}]
    }


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    service = await wrapper_services.discover('sheets', 'v4')

    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body())
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
        spreadsheetid: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {'type': 'user',
                        'role': 'writer',
                        'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permissions_body,
            fields="id"
        ))


async def spreadsheets_update_value(
        spreadsheetid: str,
        projects: list,
        wrapper_services: Aiogoogle
) -> None:
    service = await wrapper_services.discover('sheets', 'v4')
    [TABLE_VALUES.append([str(project.name),
                          str(project.close_date - project.create_date),
                          str(project.description)])
     for project in projects]
    update_body = {
        'majorDimension': 'ROWS',
        'values': TABLE_VALUES
    }
    try:
        await wrapper_services.as_service_account(
            service.spreadsheets.values.update(
                spreadsheetId=spreadsheetid,
                range='A1:E30',
                valueInputOption='USER_ENTERED',
                json=update_body
            ))
    except HTTPError:
        raise HTTPError('Ошибка заполнения таблицы')
