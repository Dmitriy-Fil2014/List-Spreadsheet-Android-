import csv
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from formula_engine import FormulaEngine, FormulaError


class Spreadsheet:
    def __init__(self, rows=100, cols=26):
        self.rows, self.cols = rows, cols
        self.cells = {}
        self.styles = {}
        self.merged_ranges = []
        self.active_sheet = "Лист1"
        self.sheet_names = ["Лист1"]
        self._sheet_states = {"Лист1": (self.cells, self.styles, self.merged_ranges)}
        self.engine = FormulaEngine(self._raw_value)

    def select_sheet(self, name):
        if name not in self._sheet_states:
            raise KeyError(name)
        self._sheet_states[self.active_sheet] = (self.cells, self.styles, self.merged_ranges)
        self.active_sheet = name
        self.cells, self.styles, self.merged_ranges = self._sheet_states[name]

    def get_raw(self, address):
        return self.cells.get(address.upper().replace("$", ""), "")

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

    def format_cells(self, addresses, **style):
        allowed = {"bold", "italic", "underline", "font_color", "fill_color",
                   "horizontal", "vertical", "number_format", "wrap_text"}
        for address in addresses:
            current = self.styles.setdefault(address.upper(), {})
            current.update({k: v for k, v in style.items() if k in allowed})

    def merge(self, cell_range):
        if cell_range not in self.merged_ranges:
            self.merged_ranges.append(cell_range)

    def autofill(self, source_range, target_range):
        source = _expand_range(source_range)
        target = _expand_range(target_range)
        if not source:
            return
        for index, address in enumerate(target):
            origin = source[index % len(source)]
            self.set(address, self.get_raw(origin))

    def sort_range(self, cell_range, key_column=1, reverse=False):
        addresses = _expand_range(cell_range)
        rows = {}
        for address in addresses:
            col, row = _split_address(address)
            rows.setdefault(row, {})[col] = self.get_raw(address)
        ordered = sorted(rows, key=lambda r: _sort_key(rows[r].get(key_column, "")), reverse=reverse)
        columns = sorted({c for row in rows.values() for c in row})
        for new_row, old_row in zip(sorted(rows), ordered):
            for col in columns:
                self.set(f"{_column_name(col)}{new_row}", rows[old_row].get(col, ""))

    def filter_range(self, cell_range, column=1, predicate=None):
        predicate = predicate or (lambda value: bool(value))
        addresses = _expand_range(cell_range)
        rows = {}
        for address in addresses:
            col, row = _split_address(address)
            rows.setdefault(row, {})[col] = self.get_raw(address)
        return [r for r in sorted(rows) if predicate(rows[r].get(column, ""))]

    def save_xlsx(self, path):
        book = Workbook()
        book.remove(book.active)
        self._sheet_states[self.active_sheet] = (self.cells, self.styles, self.merged_ranges)
        for sheet_name in self.sheet_names:
            cells, styles, merged_ranges = self._sheet_states[sheet_name]
            sheet = book.create_sheet(title=sheet_name[:31])
            for address, value in cells.items():
                cell = sheet[address]
                cell.value = value
                style = styles.get(address, {})
                cell.font = Font(bold=style.get("bold", False), italic=style.get("italic", False),
                                 underline="single" if style.get("underline") else None,
                                 color=style.get("font_color"))
                if style.get("fill_color"):
                    cell.fill = PatternFill("solid", fgColor=style["fill_color"])
                cell.alignment = Alignment(horizontal=style.get("horizontal"),
                                           vertical=style.get("vertical"),
                                           wrap_text=style.get("wrap_text", False))
                if style.get("number_format"):
                    cell.number_format = style["number_format"]
            for cell_range in merged_ranges:
                sheet.merge_cells(cell_range)
        book.save(str(path))

    @classmethod
    def load_xlsx(cls, path):
        book = load_workbook(str(path), data_only=False)
        result = cls()
        result._sheet_states = {}
        result.sheet_names = list(book.sheetnames)
        for source_sheet in book.worksheets:
            current = cls()
            for row in source_sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        current.set(cell.coordinate, cell.value)
                        if cell.has_style:
                            current.styles[cell.coordinate] = {
                                "bold": cell.font.bold, "italic": cell.font.italic,
                                "underline": bool(cell.font.underline),
                                "font_color": _color_rgb(cell.font.color),
                                "fill_color": _color_rgb(cell.fill.fgColor),
                                "horizontal": cell.alignment.horizontal,
                                "vertical": cell.alignment.vertical,
                                "wrap_text": cell.alignment.wrap_text,
                                "number_format": cell.number_format if cell.number_format != "General" else None}
            current.merged_ranges = [str(r) for r in source_sheet.merged_cells.ranges]
            result._sheet_states[source_sheet.title] = (current.cells, current.styles, current.merged_ranges)
        result.active_sheet = book.active.title
        result.cells, result.styles, result.merged_ranges = result._sheet_states[result.active_sheet]
        result.rows = max(book.active.max_row, 100)
        result.cols = max(book.active.max_column, 26)
        book.close()
        return result

    @classmethod
    def load_xlsx(cls, path):
        book = load_workbook(str(path), data_only=False)
        result = cls()
        sheet = book.active
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    result.set(cell.coordinate, cell.value)
                    if cell.has_style:
                        result.styles[cell.coordinate] = {
                            "bold": cell.font.bold, "italic": cell.font.italic,
                            "underline": bool(cell.font.underline),
                            "font_color": _color_rgb(cell.font.color),
                            "fill_color": _color_rgb(cell.fill.fgColor),
                            "horizontal": cell.alignment.horizontal,
                            "vertical": cell.alignment.vertical,
                            "wrap_text": cell.alignment.wrap_text,
                            "number_format": cell.number_format if cell.number_format != "General" else None}
        result.merged_ranges = [str(r) for r in sheet.merged_cells.ranges]
        result.rows = max(sheet.max_row, 100)
        result.cols = max(sheet.max_column, 26)
        book.close()
        return result

    @classmethod
    def load_xls(cls, path):
        try:
            import xlrd
        except ImportError as exc:
            raise RuntimeError("Для импорта .xls требуется пакет xlrd") from exc
        book = xlrd.open_workbook(str(path), formatting_info=False)
        sheet = book.sheet_by_index(0)
        result = cls(rows=max(sheet.nrows, 100), cols=max(sheet.ncols, 26))
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                value = sheet.cell_value(r, c)
                if value != "":
                    result.set(f"{_column_name(c + 1)}{r + 1}", value)
        return result

    @classmethod
    def load_csv(cls, path):
        result = cls()
        with open(path, "r", encoding="utf-8-sig", newline="") as source:
            sample = source.read(4096)
            source.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            for row_index, row in enumerate(csv.reader(source, dialect), start=1):
                for col_index, value in enumerate(row, start=1):
                    if value:
                        result.set(f"{_column_name(col_index)}{row_index}", value)
        return result


def _color_rgb(color):
    if color is None or color.type != "rgb" or not color.rgb:
        return None
    return "FF" + str(color.rgb)[-6:]


def _column_name(number):
    name = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _split_address(address):
    import re
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", address.upper())
    if not match:
        raise ValueError("Invalid cell address: " + address)
    col = 0
    for char in match.group(1):
        col = col * 26 + ord(char) - 64
    return col, int(match.group(2))


def _expand_range(value):
    import re
    match = re.fullmatch(r"([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)", value.upper())
    if not match:
        raise ValueError("Invalid range: " + value)
    c1, r1 = _split_address(match.group(1) + match.group(2))
    c2, r2 = _split_address(match.group(3) + match.group(4))
    return [f"{_column_name(c)}{r}" for r in range(min(r1, r2), max(r1, r2) + 1)
            for c in range(min(c1, c2), max(c1, c2) + 1)]


def _sort_key(value):
    try:
        return (0, float(value))
    except (TypeError, ValueError):
        return (1, str(value).casefold())
