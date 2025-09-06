from pathlib import Path
from typing import Tuple

FINAL_FILE_NAME: str = "{}_ОТРЕДАКТИРОВАННЫЙ.xlsx"
LOG_FILE_NAME: str = "{}_ЛОГ.txt"
ABSOLUTE_SHEET_NAME: str = "Абсолютная спецификация"
RELATIVE_SHEET_NAME: str = "Относительная спецификация"
PROJECT_PATH: Path = Path(__file__).parent.parent
NECESSARY_COLUMNS: Tuple[str, ...] = ('ITEM NO.', 'НАИМЕНОВАНИЕ', 'ОБОЗНАЧЕНИЕ', 'ИМЯ_РАБОЧЕГО_ФАЙЛА', 'PART NUMBER',
                                 'РАЗДЕЛ', 'СПОСОБ ИЗГОТОВЛЕНИЯ', 'МАТЕРИАЛ', 'ПРИМЕЧАНИЯ', 'ЗАКАЗ НА СТОРОНЕ')
WORK_FORMATS: Tuple[str, ...] = ('.xlsx', '.xls')
