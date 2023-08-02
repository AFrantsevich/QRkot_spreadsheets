from datetime import datetime


FORMAT = "%Y/%m/%d %H:%M:%S"

ROW_COUNT = 100
COLUMN_COUNT = 11
NOW_DATE_TIME = datetime.now().strftime(FORMAT)
TYPE_OF_INTERPRETER = 'USER_ENTERED'
RANGE_SPREADSHEET = 'A1:E30'

TABLE_VALUES = [
    ['Отчет от', NOW_DATE_TIME],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]
