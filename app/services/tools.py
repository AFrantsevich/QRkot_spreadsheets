from app.services.constants import NOW_DATE_TIME, ROW_COUNT, COLUMN_COUNT


def get_spreadsheet_body():
    return {
        'properties': {'title': f'Отчет от {NOW_DATE_TIME}',
                       'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист1',
                                   'gridProperties': {'rowCount': ROW_COUNT,
                                                      'columnCount': COLUMN_COUNT}}}]
    }
