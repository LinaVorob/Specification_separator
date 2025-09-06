from enum import Enum
from pathlib import Path


class DetailTypes(str, Enum):
    detail = 'детали'
    assembly_unit = 'сборочные единицы'
    other = 'прочие изделия'
    standard = 'стандартные изделия'
    material = 'материалы'