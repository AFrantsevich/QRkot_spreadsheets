from aiogoogle import Aiogoogle

from app.core.config import settings

from aiogoogle.excs import HTTPError

from app.services.constants import (TABLE_VALUES,
                                    RANGE_SPREADSHEET,
                                    TYPE_OF_INTERPRETER)

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
                range=RANGE_SPREADSHEET,
                valueInputOption=TYPE_OF_INTERPRETER,
                json=update_body
            ))
    except HTTPError:
        raise HTTPError('Ошибка заполнения таблицы')
