"""
Excel File Handler Module

Provides functionality for listing and interacting with Excel files in the current directory.
"""
import copy
from dataclasses import asdict
from idlelib.iomenu import errors
from pathlib import Path
from typing import Tuple, Union, List, Dict

import numpy as np
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from config.config import NECESSARY_COLUMNS, FINAL_FILE_NAME, LOG_FILE_NAME, RELATIVE_SHEET_NAME, ABSOLUTE_SHEET_NAME
from exceptions import IncorrectColumns, IncorrectRow
from logger import LoggerFile
from models import DetailTypes, AssemblyUnit, SpecificationEntity


class ExcelOutput:
    """Class for handling Excel file operations in the current working directory."""

    def __init__(self):
        self.relative_content: pd.DataFrame = None
        self.absolute_content: pd.DataFrame = None
        self.logger = None

    def write_excel_file(self, file_name: Path, file_original_contents: pd.DataFrame, details: Dict[str, Union[AssemblyUnit, SpecificationEntity]]) -> None:
        self.logger = LoggerFile(name=LOG_FILE_NAME.format(file_name.name.split('.')[0])).get_logger()
        name = f'{file_name.name.rsplit('.', 1)[0]}{FINAL_FILE_NAME}'
        absolute_content = self.create_absolute_sheet(details)
        with pd.ExcelWriter(name, mode="w", engine="openpyxl") as writer:
            file_original_contents.to_excel(writer, sheet_name=RELATIVE_SHEET_NAME, index=False)
            absolute_content.to_excel(writer, sheet_name=ABSOLUTE_SHEET_NAME, index=False)
            self.set_recalculation_count(writer)

    def set_recalculation_count(self, writer):
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


class ExcelInput:
    def __init__(self):
        self.file_contents = None
        self.fixed_content = None
        self.logger = None
        self.empty_index = list()
        self.models = list()
        self.counter_unique_models = dict()

    def read_excel_file(self, file_name: Path) -> pd.DataFrame:
        """
        Read Excel file
        :param file_name: file name
        :return: DataFrame with file contents
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
        Delete word breaks from headers
        """
        self.logger.debug('Delete word breaks from headers')
        for column in self.file_contents.columns:
            if '\n' in column:
                correct_name = column.replace('\n', '')
                self.file_contents = self.file_contents.rename(columns={column: correct_name})

    def get_count_of_parts(self):
        ...

    def fix_row(self, index: int, row: pd.Series) -> None:
        self.check_row_number(row)
        self.file_contents.iloc[index] = self.check_word_break(row)
        if self.check_empty_row(row):
            self.empty_index.append(index)

    def work_with_rows(self) -> None:
        """
        Do all fixes for file
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
        Check a row number is not contains commas
        :param row: row of file
        """
        self.logger.debug('Check row number')
        row.iloc[0] = row.iloc[0].replace(',', '.')

    def check_word_break(self, row: pd.Series) -> pd.Series:
        """
        Check a row does not contain word breaks
        :param row: row of file
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
        Check a row is empty
        :param row: row of file
        :return: True if row is empty, False otherwise
        """
        self.logger.debug('Check empty row')
        return row.iloc[1:-1].isna().all()

    def check_columns_kit(self) -> None:
        """
        Check if file has all necessary columns
        """
        self.logger.debug('Check columns\' kit')
        columns_from_file = [column.lower() for column in self.file_contents.columns]
        if not all([column in columns_from_file for column in NECESSARY_COLUMNS]):
            self.logger.error('File has not all necessary columns')
            raise IncorrectColumns

    def get_data(self) -> pd.DataFrame:
        """
        Get data from file
        :return: fixed file contents
        """
        return self.file_contents

    def delete_columns(self, columns: Tuple[str, ...]) -> None:
        """
        Delete columns from file
        :param columns: columns to delete
        """
        self.logger.debug('Delete columns')
        self.file_contents = self.file_contents.drop(columns=list(columns), errors='ignore').reset_index(drop=True)

    def replace_spaces_to_none(self):
        """
        Replace spaces to None in file contents
        """
        self.file_contents = self.file_contents.replace(r'^\s*$', np.nan, regex=True)

    # def fix_row_number_sequence(self):
    #     """
    #     Fix row number sequence. Fix an error, when row number after nine set to one
    #     """
    #     for index in range(1, len(self.file_contents)):
    #         if index > 0:
    #             if self.file_contents.iloc[index - 1, 0][:-1] == self.file_contents.iloc[index, 0][:-1]:
    #                 self.file_contents.iloc[index, 0] = self.file_contents.iloc[index - 1, 0].replace('0', '1')
    #
    #                 if self.file_contents.iloc[index - 1, 0].endswith('9') and not self.file_contents.iloc[index, 0].endswith('0'):

    def collect_model(self, row: pd.Series) -> None:
        """
        Converte data into model and collect it from file
        :param row: row of file
        """
        if pd.isna(row.iloc[4]):
            raise IncorrectRow('Не указан тип детали')
        if row.iloc[4].lower() == DetailTypes.assembly_unit.value.lower() or row.iloc[4].lower() == DetailTypes.assembly_unit_2.value.lower():
            model = AssemblyUnit(
                number=[int(number) for number in row.iloc[0].split('.')],
                components=list(),
                amount=float(row.iloc[8]),
                name=row.iloc[1].strip(),
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

