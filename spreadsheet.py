import csv
from openpyxl import Workbook, load_workbook
from formula_engine import FormulaEngine, FormulaError


class Spreadsheet:
    def __init__(self, rows=100, cols=26):
        self.rows, self.cols = rows, cols
        self.cells = {}
        self.engine = FormulaEngine(self._raw_value)

    def get_raw(self, address):
        return self.cells.get(address.upper(), "")

    def set(self, address, value):
        address = address.upper().replace("$", "")
        if value == "":
            self.cells.pop(address, None)
        else:
            self.cells[address] = str(value)

    def _raw_value(self, address):
        return self.get_raw(address)

    def value(self, address, stack=None):
        raw = self.get_raw(address)
        if not isinstance(raw, str) or not raw.startswith("="):
            return raw
        stack = set() if stack is None else stack
        address = address.upper()
        if address in stack:
            return "#CYCLE!"
        stack.add(address)
        old_get = self.engine.get_cell
        self.engine.get_cell = lambda ref: self.value(ref, stack)
        try:
            return self.engine.evaluate(raw)
        except (FormulaError, ZeroDivisionError, ValueError, TypeError):
            return "#ERROR!"
        finally:
            self.engine.get_cell = old_get
            stack.discard(address)

    def display(self, address):
        value = self.value(address)
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def save_xlsx(self, path):
        book = Workbook()
        sheet = book.active
        sheet.title = "Лист1"
        for address, value in self.cells.items():
            sheet[address] = value
        book.save(str(path))

    @classmethod
    def load_xlsx(cls, path):
        book = load_workbook(str(path), data_only=False)
        result = cls()
        sheet = book.active
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    result.set(cell.coordinate, cell.value)
        book.close()
        return result

    @classmethod
    def load_csv(cls, path):
        result = cls()
        with open(path, "r", encoding="utf-8-sig", newline="") as source:
            sample = source.read(4096)
            source.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\\t")
            except csv.Error:
                dialect = csv.excel
            for row_index, row in enumerate(csv.reader(source, dialect), start=1):
                for col_index, value in enumerate(row, start=1):
                    if value:
                        result.set(f"{_column_name(col_index)}{row_index}", value)
        return result


def _column_name(number):
    name = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name
