class IncorrectColumns(Exception):
    pass


class IncorrectRow(Exception):
    def __init__(self, msg):
        super().__init__(msg)
        self.msg = msg

    def __str__(self):
        return self.msg

class EmptyRow(IncorrectRow):
    def __init__(self, row_number: int):
        super().__init__(f"№ {row_number}: Row is empty")

class IncorrectData(IncorrectRow):
    def __init__(self, row_number: int):
        super().__init__(f"№ {row_number}: Row contains incorrect data")