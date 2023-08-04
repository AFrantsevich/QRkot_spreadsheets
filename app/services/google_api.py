from copy import deepcopy

from aiogoogle import Aiogoogle

from app.core.config import settings

from app.services.constants import (TABLE_VALUES,
                                    RANGE_SPREADSHEET,
                                    TYPE_OF_INTERPRETER,
                                    ROW_COUNT)

from app.services.tools import get_spreadsheet_body


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    service = await wrapper_services.discover('sheets', 'v4')

    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=get_spreadsheet_body())
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
    table_values = deepcopy(TABLE_VALUES)
    [table_values.append([str(project.name),
                          str(project.close_date - project.create_date),
                          str(project.description)])
     for project in projects]
    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    if len(table_values) > ROW_COUNT:
        raise ValueError('Количество новых записей больше количества строк')

    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range=RANGE_SPREADSHEET,
            valueInputOption=TYPE_OF_INTERPRETER,
            json=update_body
        ))
