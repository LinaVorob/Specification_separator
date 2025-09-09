from pathlib import Path
from typing import Tuple

FINAL_FILE_NAME: str = "_ОТРЕДАКТИРОВАННЫЙ.xlsx"
LOG_FILE_NAME: str = "{}_ЛОГ.txt"
ABSOLUTE_SHEET_NAME: str = "Абсолютная спецификация"
RELATIVE_SHEET_NAME: str = "Относительная спецификация"
PROJECT_PATH: Path = Path(__file__).parent.parent
NECESSARY_COLUMNS: Tuple[str, ...] = ('уровень', 'наименование', 'обозначение', 'имя_рабочего_файла',
                                      'раздел', 'способ изготовления', 'материал', 'заказ на стороне',
                                      'количество', 'примечание')
WORK_FORMATS: Tuple[str, ...] = ('.xlsx', '.xls')
NOT_NEED_COLUMNS: Tuple[str, ...] = ('PART NUMBER',)