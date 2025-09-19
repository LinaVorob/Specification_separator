from dataclasses import dataclass
from enum import Enum
from typing import List, Union, Optional
from exceptions import IncorrectRow

class DetailTypes(str, Enum):
    """
    Типы деталей
    """

    detail = 'Детали'
    assembly_unit = 'Сборочные единицы'
    assembly_unit_2 = 'Сборочные изделия'
    other = 'Прочие изделия'
    standard = 'Стандартные изделия'
    material = 'Материалы'

    @classmethod
    def get_type(cls, type_name: str) -> "DetailTypes":
        for detail_type in cls:
            if detail_type.value.lower() == type_name.lower():
                return detail_type
        else:
           raise IncorrectRow(f'Неизвестный тип детали {type_name}')


@dataclass
class SpecificationEntity:
    """
    Модель детали
    """

    number: List[int]
    name: str
    detail_type: DetailTypes
    code: Optional[str] = None
    work_file: Optional[str] = None
    making_type: Optional[str] = None
    material: Optional[str] = None
    comment: Optional[str] = None
    is_order: bool = False
    amount: float = 0.0
    count_in_device: float = 1.0


@dataclass
class AssemblyUnit:
    """
    Модель сборочной единицы
    """

    number: List[int]
    name: str
    code: Optional[str] = None
    work_file: Optional[str] = None
    making_type: Optional[str] = None
    material: Optional[str] = None
    comment: Optional[str] = None
    is_order: bool = False
    components: List[Union[SpecificationEntity, "AssemblyUnit"]] | None = None
    detail_type: DetailTypes = DetailTypes.assembly_unit
    amount: float = 0.0
    count_in_device: float = 1.0


    def is_detail_in_assembly(self, detail: [SpecificationEntity, "AssemblyUnit"]) -> bool:
        if len(detail.number) - len(self.number) == 1:
            for i in range(len(self.number)):
                if self.number[i] != detail.number[i]:
                    break
            else:
                return True
        return False