from dataclasses import dataclass
from enum import Enum
from typing import List, Union, Optional

from exceptions import IncorrectRow


class DetailTypes(str, Enum):
    detail = 'детали'
    assembly_unit = 'сборочные единицы'
    other = 'прочие изделия'
    standard = 'стандартные изделия'
    material = 'материалы'

    @classmethod
    def get_type(cls, type_name: str) -> "DetailTypes":
        for detail_type in cls:
            if detail_type.value == type_name.lower():
                return detail_type
        else:
           raise IncorrectRow(f'Неизвестный тип детали {type_name}')

class DetailNumeration:

    def __init__(self, number: str):
        number_parts = number.split('.')
        if len(number_parts) < 6:
            number_parts.extend([None] * (6 - len(number_parts)))
        self.main_number: int = int(number_parts[0]) if number_parts[0] else None
        self.second_level_number: int | None = int(number_parts[1]) if number_parts[1] else None
        self.third_level_number: int | None = int(number_parts[2]) if number_parts[2] else None
        self.forth_level_number: int | None = int(number_parts[3]) if number_parts[3] else None
        self.fifth_level_number: int | None = int(number_parts[4]) if number_parts[4] else None
        self.sixth_level_number: int | None = int(number_parts[5]) if number_parts[5] else None

    def __len__(self):
        return len([item for item in self.__dict__.values() if item is not None])


@dataclass
class SpecificationEntity:
    number: List[int] #DetailNumeration
    name: str
    detail_type: DetailTypes
    code: Optional[str] = None
    work_file: Optional[str] = None
    making_type: Optional[str] = None
    material: Optional[str] = None
    comment: Optional[str] = None
    is_order: bool = False
    amount: float = 0.0
    count_in_device: float = 0.0


@dataclass
class AssemblyUnit:
    number: List[int]
    name: str
    components: List[Union[SpecificationEntity, "AssemblyUnit"]] | None = None
    detail_type: DetailTypes = DetailTypes.assembly_unit
    amount: float = 0.0
    count_in_device: float = 0.0


    def is_detail_in_assembly(self, detail: [SpecificationEntity, "AssemblyUnit"]) -> bool:
        if len(detail.number) - len(self.number) == 1:
            for i in range(len(self.number)):
                if self.number[i] != detail.number[i]:
                    break
            else:
                return True
        return False