from kivy.app import App
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from pathlib import Path

from spreadsheet import Spreadsheet

KV = """
<RootView>:
    orientation: "vertical"
    spacing: dp(6)
    padding: dp(8)
    canvas.before:
        Color:
            rgba: 0.95, 0.96, 0.98, 1
        Rectangle:
            pos: self.pos
            size: self.size
    BoxLayout:
        size_hint_y: None
        height: dp(48)
        Label:
            text: "Лист Таблицы"
            bold: True
            font_size: "20sp"
            color: 0.12, 0.17, 0.25, 1
            halign: "left"
            text_size: self.size
        Button:
            text: "Открыть файл"
            size_hint_x: None
            width: dp(116)
            on_release: root.open_file()
        Button:
            text: "Сохранить"
            size_hint_x: None
            width: dp(96)
            on_release: root.save_file()
    BoxLayout:
        size_hint_y: None
        height: dp(42)
        spacing: dp(5)
        Label:
            text: root.selected_cell
            size_hint_x: None
            width: dp(52)
            color: 0.15, 0.2, 0.28, 1
        TextInput:
            id: formula
            text: root.formula_text
            multiline: False
            hint_text: "Значение или формула, например =SUM(A1:A5)"
            on_text_validate: root.commit_value(self.text)
        Button:
            text: "✓"
            size_hint_x: None
            width: dp(48)
            on_release: root.commit_value(formula.text)
    ScrollView:
        do_scroll_x: True
        do_scroll_y: True
        GridLayout:
            id: sheet_grid
            cols: 11
            size_hint: None, None
            width: self.minimum_width
            height: self.minimum_height
            row_default_height: dp(42)
            row_force_default: True
            col_default_width: dp(104)
            col_force_default: True
            spacing: dp(1)
    Label:
        text: root.status
        size_hint_y: None
        height: dp(24)
        color: 0.32, 0.38, 0.46, 1
        font_size: "12sp"
"""

class RootView(BoxLayout):
    selected_cell = StringProperty("A1")
    formula_text = StringProperty("")
    status = StringProperty("Готово · Выберите ячейку и введите значение или формулу")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.sheet = Spreadsheet(rows=100, cols=26)
        self._refresh_grid()

    def _refresh_grid(self):
        grid = self.ids.sheet_grid
        grid.clear_widgets()
        headers = [""] + [self._col_name(i) for i in range(1, 11)]
        for title in headers:
            grid.add_widget(Label(text=title, bold=True, color=(0.25, 0.31, 0.4, 1),
                                  size_hint=(None, None), size=(dp(104), dp(42))))
        for row in range(1, 101):
            grid.add_widget(Label(text=str(row), color=(0.35, 0.4, 0.48, 1),
                                  size_hint=(None, None), size=(dp(104), dp(42))))
            for col in range(1, 11):
                address = f"{self._col_name(col)}{row}"
                btn = Button(text=self.sheet.display(address), font_size="13sp",
                             background_normal="", background_color=(1, 1, 1, 1),
                             color=(0.12, 0.16, 0.22, 1),
                             size_hint=(None, None), size=(dp(104), dp(42)))
                btn.bind(on_release=lambda instance, a=address: self.select_cell(a))
                grid.add_widget(btn)

    @staticmethod
    def _col_name(n):
        result = ""
        while n:
            n, rem = divmod(n - 1, 26)
            result = chr(65 + rem) + result
        return result

    def select_cell(self, address):
        self.selected_cell = address
        self.formula_text = self.sheet.get_raw(address)
        self.ids.formula.text = self.formula_text

    def commit_value(self, value):
        self.sheet.set(self.selected_cell, value)
        self.formula_text = value
        self._refresh_grid()
        self.status = f"{self.selected_cell} обновлена"

    def save_file(self):
        try:
            path = Path(App.get_running_app().user_data_dir) / "таблица.xlsx"
            self.sheet.save_xlsx(path)
            self.status = f"Сохранено в приложении: {path.name}"
        except Exception as exc:
            self._message("Не удалось сохранить", str(exc))

    def open_file(self):
        try:
            from plyer import filechooser
            filechooser.open_file(
                on_selection=self._on_file_selected,
                filters=["*.xlsx", "*.csv"]
            )
            self.status = "Выберите файл Excel (.xlsx) или CSV"
        except Exception as exc:
            self._message("Не удалось открыть выбор файла", str(exc))

    def _on_file_selected(self, selection):
        if not selection:
            self.status = "Открытие файла отменено"
            return
        try:
            path = str(selection[0])
            suffix = Path(path).suffix.lower()
            if suffix == ".xlsx":
                self.sheet = Spreadsheet.load_xlsx(path)
            elif suffix == ".csv":
                self.sheet = Spreadsheet.load_csv(path)
            else:
                self._message("Формат не поддерживается", "Выберите файл .xlsx или .csv")
                return
            self._refresh_grid()
            self.select_cell("A1")
            self.status = f"Открыт файл: {Path(path).name}"
        except Exception as exc:
            self._message("Не удалось открыть файл", str(exc))

    def _message(self, title, message):
        Popup(title=title, content=Label(text=message), size_hint=(.88, .4)).open()


class ListSpreadsheetApp(App):
    def build(self):
        Builder.load_string(KV)
        return RootView()


if __name__ == "__main__":
    ListSpreadsheetApp().run()
