"""
Excel File Handler Module

Provides functionality for listing and interacting with Excel files in the current directory.
"""
import copy
from dataclasses import asdict
from pathlib import Path
from typing import Tuple, Union, List, Dict

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import NECESSARY_COLUMNS, FINAL_FILE_NAME, LOG_FILE_NAME, RELATIVE_SHEET_NAME, ABSOLUTE_SHEET_NAME
from exceptions import IncorrectColumns, IncorrectRow
from logger import LoggerFile
from models import DetailTypes, AssemblyUnit, SpecificationEntity


class ExcelOutput:
    """
    Класс для обработки выходных Excel-файлов.

    Отвечает за создание и форматирование таблиц Excel: относительная и абсолютная ведомости.
    """
    def __init__(self):
        self.relative_content: pd.DataFrame = None
        self.absolute_content: pd.DataFrame = None
        self.logger = None

    def clear_content(self) -> None:
        """
        Очистка содержимого данных.

        Сбрасывает значения атрибутов `relative_content` и `absolute_content`.
        """
        self.relative_content: pd.DataFrame = None
        self.absolute_content: pd.DataFrame = None

    def write_excel_file(self, file_name: Path, details: Dict[str, Union[AssemblyUnit, SpecificationEntity]]) -> None:
        """
        Создаёт и сохраняет Excel-файл с данными из переданного словаря деталей.

        Args:
            file_name (Path): Путь к исходному файлу, на основе которого будет создан новый Excel-файл.
            details (Dict[str, Union[AssemblyUnit, SpecificationEntity]]): Словарь деталей.
        """
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        name = f'{file_name.name.rsplit('.', 1)[0]}{FINAL_FILE_NAME}'
        absolute_content = self.create_absolute_sheet(details)
        with pd.ExcelWriter(name, mode="w", engine="openpyxl") as writer:
            self.relative_content.to_excel(writer, sheet_name=RELATIVE_SHEET_NAME, index=False)
            absolute_content.to_excel(writer, sheet_name=ABSOLUTE_SHEET_NAME, index=False)
            self.set_recalculation_count(writer)
            self.set_formating_for_relative(writer.book[RELATIVE_SHEET_NAME])
            self.set_formating_for_absolute(writer.book[ABSOLUTE_SHEET_NAME])
        self.logger.info(f'Создан файл: {name}')

    def set_recalculation_count(self, writer) -> None:
        """
        Устанавливает формулы для автоматического расчёта количества в приборе.

        Вставляет строку с количеством приборов и применяет формулы для подсчёта общего количества деталей.

        Args:
            writer: Объект ExcelWriter для записи данных в Excel-файл.
        """
        wb: Workbook = writer.book
        worksheet: Worksheet = wb[ABSOLUTE_SHEET_NAME]
        worksheet.insert_rows(0, 1)
        worksheet['A1'] = 'Количество приборов:'
        worksheet['B1'] = 1
        for row_number, row in enumerate(worksheet.rows):
            if row_number < 2:
                continue

            formula = f'=B1*{row[7].value}'
            worksheet[f'H{row_number + 1}'] = formula


    def create_absolute_sheet(self, details: Dict[str, Union[AssemblyUnit, SpecificationEntity]]) -> pd.DataFrame:
        """
        Создаёт таблицу абсолютной спецификации из переданных данных.

        Args:
            details (Dict[str, Union[AssemblyUnit, SpecificationEntity]]): Словарь деталей.

        Returns:
            pd.DataFrame: Отформатированная таблица абсолютной спецификации.
        """
        absolute_content = pd.DataFrame([asdict(detail) for detail in details.values()])
        absolute_content = absolute_content.drop(columns=['number', 'components', 'amount'])
        absolute_content['detail_type'] = absolute_content['detail_type'].map(lambda x: x.value)
        absolute_content = absolute_content.sort_values(by=['detail_type', 'name']).reset_index(drop=True)
        new_order = ["name", 'code', 'work_file', 'detail_type', 'making_type', 'material', 'is_order', 'count_in_device', 'comment']
        absolute_content = absolute_content[new_order]
        absolute_content.rename(
            columns={
                "detail_type": "Раздел",
                "name": "Наименование",
                "making_type": "Способ изготовления",
                "material": "Материал",
                "comment": "Примечание",
                "is_order": "Заказ на стороне",
                "count_in_device": "Количество в приборе",
                "work_file": "Имя_рабочего_файла",
                "code": "Обозначение"
            }, inplace=True
        )
        return absolute_content

    def format_sheet(self, worksheet: Worksheet, title_row_number: int = 1) -> None:
        """
        Применяет форматирование к листу Excel.

        Устанавливает ширину столбцов, включает перенос текста в ячейках и делает заголовок жирным.

        Args:
            worksheet (Worksheet): Лист Excel для форматирования.
            title_row_number (int): Номер строки заголовка.
        """
        # Set column widths
        column_widths = {
            'A': 30,
            'B': 30,
            'C': 25,
            'D': 25,
            'E': 20,
            'F': 20,
            'G': 15,
            'H': 15,
            'I': 15,
        }
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

        # Wrap text in all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)

        # Make the first row bold
        for cell in worksheet[title_row_number]:
            cell.font = Font(bold=True)

    def set_formating_for_absolute(self, worksheet: Worksheet = None):
        """
        Применяет форматирование к листу абсолютной спецификации.

        Args:
            worksheet (Worksheet): Лист Excel для форматирования.
        """
        self.format_sheet(worksheet, 2)

    def set_formating_for_relative(self, worksheet: Worksheet = None):
        """
        Применяет форматирование к листу относительной спецификации.

        Args:
            worksheet (Worksheet): Лист Excel для форматирования.
        """
        self.format_sheet(worksheet)
        details = [assembly_unit for assembly_unit in self.absolute_content if isinstance(assembly_unit, AssemblyUnit)]
        for assembly_unit in details:
            self.create_group(worksheet, assembly_unit)

    def create_group(self, worksheet: Worksheet, assembly_unit: AssemblyUnit, outline_level: int = 1) -> None:
        """
        Группирует компоненты сборки на листе Excel.

        Args:
            worksheet (Worksheet): Лист Excel для группировки.
            assembly_unit (AssemblyUnit): Объект сборки.
            outline_level (int): Уровень вложенности.
        """
        if assembly_unit.components:
            row_number_first = '.'.join([str(number_item) for number_item in assembly_unit.components[0].number])
            row_number_last = '.'.join([str(number_item) for number_item in assembly_unit.components[-1].number])
            row_start = -1
            row_end = -1
            for row in worksheet.iter_rows(min_row=2, max_col=1):
                if row[0].value == row_number_first:
                    row_start = row[0].row
                elif row[0].value == row_number_last:
                    row_end = row[0].row
                if row_start != -1 and row_end != -1:
                    if row_start < row_end:
                        worksheet.row_dimensions.group(row_start, row_end, hidden=True, outline_level=outline_level)
                        break

            for component in assembly_unit.components:
                if component.detail_type == DetailTypes.assembly_unit:
                    self.create_group(worksheet, component, outline_level + 1)


class ExcelInput:
    """
    Класс для обработки входных Excel-файлов.

    Отвечает за чтение, проверку и предварительную обработку данных из Excel-файлов.
    """

    def __init__(self):
        self.file_contents = None
        self.logger = None
        self.empty_index = list()
        self.models = list()
        self.counter_unique_models = dict()

    def read_excel_file(self, file_name: Path) -> pd.DataFrame:
        """
        Читает данные из Excel-файла.

        Args:
            file_name (Path): Путь к Excel-файлу.

        Returns:
            pd.DataFrame: Содержимое файла в виде DataFrame.
        """

        engine = 'xlrd'
        if file_name.suffix == ".xlsx":
            engine = 'openpyxl'
        self.logger = LoggerFile(log_file=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        self.logger.info(f'Read file {file_name.name}')
        self.file_contents = pd.read_excel(str(file_name), engine=engine, na_values=[' ', '\t', '\n', '', '-'])
        self.file_contents['Количество в приборе'] = np.nan
        self.replace_spaces_to_none()
        self.delete_word_break()
        self.check_columns_kit()

    def delete_word_break(self) -> None:
        """
        Удаляет символы переноса строки из заголовков столбцов.

        Перезаписывает заголовки, заменяя символы `\n` на пустую строку.
        """
        self.logger.debug('Delete word breaks from headers')
        for column in self.file_contents.columns:
            if '\n' in column:
                correct_name = column.replace('\n', '')
                self.file_contents = self.file_contents.rename(columns={column: correct_name})

    def fix_row(self, index: int, row: pd.Series) -> None:
        """
        Выполняет исправление строки.

        Args:
            index (int): Индекс строки.
            row (pd.Series): Строка DataFrame.
        """
        self.check_row_number(row)
        self.file_contents.iloc[index] = self.check_word_break(row)
        if self.check_empty_row(row):
            self.empty_index.append(index)

    def work_with_rows(self) -> None:
        """
        Обрабатывает все строки в DataFrame.

        Вызывает методы проверки и исправления строк.
        """
        self.file_contents.iloc[:, 0] = self.file_contents.iloc[:, 0].astype(str)

        for index, row in self.file_contents.iterrows():
            try:
                self.fix_row(index, row)
                self.collect_model(row)
            except IncorrectRow as err:
                self.logger.error(str(err))
                continue
        else:
            self.file_contents = self.file_contents.drop(self.empty_index, errors='ignore').reset_index(drop=True)

    def check_row_number(self, row: pd.Series) -> None:
        """
        Проверяет номер строки.

        Заменяет запятые на точки в первом столбце.

        Args:
            row (pd.Series): Строка DataFrame.
        """
        self.logger.debug('Check row number')
        row.iloc[0] = row.iloc[0].replace(',', '.')

    def check_word_break(self, row: pd.Series) -> pd.Series:
        """
        Удаляет символы переноса строки из ячеек строки.

        Args:
            row (pd.Series): Строка DataFrame.

        Returns:
            pd.Series: Строка с очищенными ячейками.
        """
        for index, cell in row.items():
            self.logger.debug(f'Check word break in cell {index}')
            if not pd.isna(cell):
                cell = str(cell)
                row.loc[index] = cell.replace('\n', ' ').replace('\n', ' ').replace('\t', ' ').replace('\r',
                                                                                                       ' ').replace(
                    '\xa0', ' ').replace('  ', ' ')
        return row

    def check_empty_row(self, row: pd.Series) -> bool:
        """
        Проверяет, является ли строка пустой.

        Args:
            row (pd.Series): Строка DataFrame.

        Returns:
            bool: True, если строка пустая, иначе False.
        """
        self.logger.debug('Check empty row')
        return row.iloc[1:-1].isna().all()

    def check_columns_kit(self) -> None:
        """
        Проверяет, содержит ли файл все необходимые столбцы.

        Raises:
            IncorrectColumns: Если не все необходимые столбцы присутствуют.
        """
        self.logger.debug('Check columns\' kit')
        columns_from_file = [column.lower() for column in self.file_contents.columns]
        if not all([column in columns_from_file for column in NECESSARY_COLUMNS]):
            self.logger.error('File has not all necessary columns')
            raise IncorrectColumns

    def get_data(self) -> pd.DataFrame:
        """
        Возвращает содержимое файла в виде DataFrame.

        Returns:
            pd.DataFrame: Данные Excel-файла.
        """
        return self.file_contents

    def delete_columns(self, columns: Tuple[str, ...]) -> None:
        """
        Удаляет указанные столбцы из DataFrame.

        Args:
            columns (Tuple[str, ...]): Кортеж с названиями удаляемых столбцов.
        """
        self.logger.debug('Delete columns')
        self.file_contents = self.file_contents.drop(columns=list(columns), errors='ignore').reset_index(drop=True)

    def replace_spaces_to_none(self):
        """
        Заменяет ячейки, содержащие только пробелы, на NaN.
        """
        self.file_contents = self.file_contents.replace(r'^\s*$', np.nan, regex=True)

    def collect_model(self, row: pd.Series) -> None:
        """
        Собирает модель из строки данных.

        Args:
            row (pd.Series): Строка DataFrame.

        Raises:
            IncorrectRow: Если тип детали не указан.
        """
        if pd.isna(row.iloc[4]):
            raise IncorrectRow('Не указан тип детали')
        if row.iloc[4].lower() == DetailTypes.assembly_unit.value.lower() or row.iloc[4].lower() == DetailTypes.assembly_unit_2.value.lower():
            model = AssemblyUnit(
                number=[int(number) for number in row.iloc[0].split('.')],
                components=list(),
                amount=float(row.iloc[8]),
                name=row.iloc[1].strip(),
                code=row.iloc[2],
                work_file=row.iloc[3],
                material=row.iloc[6],
                making_type=row.iloc[5],
                comment=row.iloc[9],
                is_order=row.iloc[7],
                count_in_device=float(row.iloc[8])
            )
        else:
            model = SpecificationEntity(
                number=[int(number) for number in row.iloc[0].split('.')],
                name=row.iloc[1].strip(),
                detail_type=DetailTypes.get_type(row.iloc[4]),
                code=row.iloc[2],
                work_file=row.iloc[3],
                making_type=row.iloc[5],
                material=row.iloc[6],
                comment=row.iloc[9],
                is_order=row.iloc[7],
                amount=float(row.iloc[8]),
                count_in_device=float(row.iloc[8])
            )


        result = self.find_assembly(self.models, model)
        if not result:
            self.counter_unique_models[model.name] = copy.deepcopy(model)
            self.models.append(model)

    def find_assembly(self, models_collection: List[Union[AssemblyUnit,SpecificationEntity]], model: Union[AssemblyUnit, SpecificationEntity], amount_upper: float = 1) -> bool:
        """
        Рекурсивный метод для поиска детали в сборке.

        Args:
            models_collection (List): Список объектов AssemblyUnit или SpecificationEntity.
            model (Union[AssemblyUnit, SpecificationEntity]): Деталь, которую ищем.
            amount_upper (float): Количество верхнего уровня, необходимое для расчета общего количества.

        Returns:
            bool: True, если деталь найдена, иначе False.
        """
        for component in models_collection:
            if isinstance(component, AssemblyUnit):
                if component.is_detail_in_assembly(model):
                    component.components.append(model)
                    if self.counter_unique_models.get(model.name) is None:
                        self.counter_unique_models[model.name] = copy.deepcopy(model)
                        self.counter_unique_models[model.name].count_in_device = component.count_in_device * amount_upper * model.amount
                    else:
                        self.counter_unique_models[model.name].count_in_device = self.counter_unique_models[model.name].count_in_device + component.count_in_device * amount_upper * model.amount
                    return True
                else:
                    if self.find_assembly(component.components, model, amount_upper=component.count_in_device):
                        return True
                    else:
                        continue
        return False

    def clear_content(self) -> None:
        """
        Очищает содержимое данных.

        Сбрасывает значения атрибутов `file_contents`, `empty_index`, `models` и `counter_unique_models`.
        """
        self.file_contents = None
        self.empty_index = list()
        self.models = list()
        self.counter_unique_models = dict()

