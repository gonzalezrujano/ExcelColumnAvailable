"""
Calculator Excel columns keys availables

A               | B              | C              |
test

Input
column_names: ["test2", "test3"]
last_column: "A"

Output
column_availables: [
    {'column_name': 'test2', 'column_key': 'B'},
    {'column_name': 'test3', 'column_key': 'C'}
]
"""


class ExcelColumnAvailables(object):
    def __init__(self, column_names=None, last_column=None):
        self.column_names = column_names
        self.last_column = last_column
        self.next_column = None

        if not isinstance(self.column_names, list):
            raise Exception("Error: column_names params must be an string list")

        if not self.is_valid_column_excel():
            raise Exception("Error: last_column must be a valid excel column")

        self.calculate_next_column()

    def is_valid_column_excel(self):
        if self.last_column is None:
            return True

        for idx, char in enumerate(self.last_column):
            ascii_code = ord(char)
            if not (65 <= ascii_code <= 90):
                return False

        return True

    def calculate_next_column(self):
        if self.last_column is None:
            return "A"

        next_column = str()
        up_next_char = False

        for idx, char in enumerate(self.last_column[::-1]):
            ascii_code = ord(char)
            if (up_next_char and idx != 0) or (idx == 0):
                if ascii_code < 90:
                    new_char = chr(ascii_code + 1)
                    next_column = new_char + next_column
                    up_next_char = False
                else:
                    new_char = "A"
                    next_column = new_char + next_column
                    up_next_char = True
            else:
                next_column = char + next_column

        if up_next_char:
            next_column = "A" + next_column

        self.next_column = next_column

    def get_columns(self):
        columns_availables = list()

        for column_name in self.column_names:
            columns_availables.append({
                'column_name': column_name,
                'column_key': self.next_column
            })
            self.last_column = self.next_column
            self.calculate_next_column()

        return columns_availables
