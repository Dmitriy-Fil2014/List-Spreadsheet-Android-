import json
from pathlib import Path
from kivy.app import App
from kivy.clock import Clock
from kivy.animation import Animation
from kivy.graphics import Color, RoundedRectangle
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.utils import platform
from kivy.properties import StringProperty
from kivy.storage.jsonstore import JsonStore
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.recyclegridlayout import RecycleGridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from spreadsheet import Spreadsheet

COLORS={"light":{"bg":(0.95,.97,.96,1),"text":(.1,.16,.14,1),"muted":(.35,.43,.39,1),"primary":(.04,.53,.31,1)},
"dark":{"bg":(.07,.1,.09,1),"text":(.91,.96,.93,1),"muted":(.65,.72,.68,1),"primary":(.2,.78,.49,1)}}

class CellButton(RecycleDataViewBehavior,Button):
    address=StringProperty("")
    role=StringProperty("cell")
    def refresh_view_attrs(self,rv,index,data):
        self.address=data.get("address",""); self.role=data.get("role","cell")
        result = super().refresh_view_attrs(rv,index,data)
        self.background_normal = ""
        self.background_down = ""
        self.background_color = data.get("background_color", (1,1,1,1))
        self.color = data.get("color", (.12,.16,.22,1))
        self.bold = data.get("bold", False)
        self.italic = data.get("italic", False)
        self.underline = data.get("underline", False)
        self.font_size = "13sp"
        return result
    def on_release(self):
        if self.role=="cell" and self.address: App.get_running_app().root.select_cell(self.address)

class SheetGrid(RecycleView):
    COLS = 27
    ROWS = 151
    CELL_W = dp(92)
    CELL_H = dp(40)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.viewclass = CellButton
        self.size_hint = (None, None)
        self.width = self.CELL_W * self.COLS
        self.height = self.CELL_H * self.ROWS
        layout = RecycleGridLayout(cols=self.COLS, default_size=(self.CELL_W, self.CELL_H), default_size_hint=(None, None), size_hint=(None, None), width=self.width, height=self.height, spacing=dp(1))
        layout.bind(minimum_height=layout.setter("height"), minimum_width=layout.setter("width"))
        self.add_widget(layout)

class RootView(BoxLayout):
    selected_cell=StringProperty("A1")
    def __init__(self,**kwargs):
        super().__init__(orientation="vertical",padding=dp(10),spacing=dp(8),**kwargs)
        self.sheet=Spreadsheet(rows=100,cols=26)
        self.store=JsonStore(str(Path(App.get_running_app().user_data_dir)/"preferences.json"))
        self.language="ru"; self.theme="system"; self.translations={}
    def tr(self,key): return self.translations.get(key,key)
    def _locale(self,code):
        try: return json.loads((Path(__file__).parent/"locales"/(code+".json")).read_text(encoding="utf-8"))
        except (OSError,ValueError): return {}
    def _build_ui(self):
        self.show_library()

    def _documents(self):
        return self.store.get("documents").get("items", []) if self.store.exists("documents") else []

    def show_library(self):
        if getattr(self, "current_document", None) and hasattr(self, "status_label"):
            self.save_file()
        self.clear_widgets()
        Window.clearcolor = COLORS["light"]["primary"]
        page = BoxLayout(orientation="vertical", padding=dp(16), spacing=dp(12))
        header = BoxLayout(size_hint_y=None, height=dp(54), spacing=dp(8))
        header.add_widget(Label(text="Мои таблицы", bold=True, font_size="23sp", color=(1,1,1,1), halign="left"))
        header.add_widget(self._button("⚙", self.open_settings, dp(48)))
        page.add_widget(header)
        create = Button(text="+  Создать таблицу", size_hint_y=None, height=dp(54),
                        background_normal="", background_color=(1,1,1,1), color=(.05,.25,.65,1),
                        bold=True, font_size="17sp")
        create.bind(on_release=lambda *_: self.create_table())
        page.add_widget(create)
        docs = self._documents()
        if not docs:
            page.add_widget(Label(text="Пока ничего нет\\nСоздайте первую таблицу", color=(1,1,1,.9),
                                  halign="center", valign="middle", font_size="17sp"))
        else:
            scroll = ScrollView()
            items = BoxLayout(orientation="vertical", spacing=dp(8), size_hint_y=None)
            items.bind(minimum_height=items.setter("height"))
            for doc in docs:
                row = BoxLayout(size_hint_y=None, height=dp(58), spacing=dp(6))
                open_btn = Button(text="▦  " + doc.get("name","Таблица"), background_normal="",
                                  background_color=(1,1,1,.96), color=(.08,.16,.3,1), halign="left")
                open_btn.bind(on_release=lambda _, d=doc: self.open_table(d))
                row.add_widget(open_btn)
                del_btn = Button(text="×", size_hint_x=None, width=dp(46), background_normal="",
                                 background_color=(.82,.18,.22,1), color=(1,1,1,1))
                del_btn.bind(on_release=lambda _, d=doc: self.delete_table(d))
                row.add_widget(del_btn)
                items.add_widget(row)
            scroll.add_widget(items)
            page.add_widget(scroll)
        self.add_widget(page)

    def create_table(self):
        docs = self._documents()
        name = "Таблица " + str(len(docs) + 1)
        filename = "sheet_" + str(int(__import__("time").time() * 1000)) + ".xlsx"
        doc = {"name": name, "file": filename}
        docs.append(doc)
        self.store.put("documents", items=docs)
        (Path(App.get_running_app().user_data_dir) / "documents").mkdir(parents=True, exist_ok=True)
        self.sheet = Spreadsheet(rows=150, cols=26)
        self.current_document = doc
        self.selected_cell = "A1"
        self.open_editor()

    def open_table(self, doc):
        path = Path(App.get_running_app().user_data_dir) / "documents" / doc["file"]
        try:
            self.sheet = Spreadsheet.load_xlsx(path) if path.exists() else Spreadsheet(rows=100, cols=26)
            self.current_document = doc
            self.open_editor()
        except Exception as exc:
            self.message("Не удалось открыть таблицу", str(exc))

    def delete_table(self, doc):
        docs = [d for d in self._documents() if d.get("file") != doc.get("file")]
        self.store.put("documents", items=docs)
        path = Path(App.get_running_app().user_data_dir) / "documents" / doc["file"]
        try:
            if path.exists(): path.unlink()
        except OSError: pass
        self.show_library()

    def open_editor(self):
        self._build_editor_ui()

    def _build_editor_ui(self):
        self.clear_widgets()
        top=BoxLayout(size_hint_y=None,height=dp(50),spacing=dp(8))
        title=Label(text="Лист Таблицы",bold=True,font_size="20sp",color=COLORS["light"]["text"],halign="left")
        top.add_widget(title); top.add_widget(self._button("Мои таблицы",self.show_library,dp(112))); top.add_widget(self._button("Сохранить",self.save_file,dp(92))); top.add_widget(self._button(self.tr("open"),self.open_file,dp(88)))
        top.add_widget(self._button("⋮",self.open_menu,dp(44))); self.add_widget(top)
        bar=BoxLayout(size_hint_y=None,height=dp(44),spacing=dp(6))
        self.cell_label=Label(text=self.selected_cell,size_hint_x=None,width=dp(48),color=COLORS["light"]["text"])
        bar.add_widget(self.cell_label)
        self.formula=TextInput(multiline=False,hint_text=self.tr("formula_hint"))
        self.formula.bind(on_text_validate=lambda *_:self.commit_value(self.formula.text))
        bar.add_widget(self.formula); bar.add_widget(self._button("✓",lambda:self.commit_value(self.formula.text),dp(44)))
        self.add_widget(bar)
        tools=BoxLayout(size_hint_y=None,height=dp(38),spacing=dp(4))
        for label,fn in [("B",lambda:self.apply_format(bold=True)),("I",lambda:self.apply_format(italic=True)),
                         ("U",lambda:self.apply_format(underline=True)),
                         ("Цвет",lambda:self.apply_format(font_color="FF087A4E")),
                         ("Фон",lambda:self.apply_format(fill_color="FFDFF5E8")),
                         ("Центр",lambda:self.apply_format(horizontal="center",vertical="center")),
                         ("Объед.",self.merge_next_cell),("A→Z",self.sort_sheet),
                         ("Фильтр",self.filter_sheet),("Автозап.",self.fill_down),("Σ SUM",lambda:self.insert_formula("=SUM(A1:A5)")),("AVG",lambda:self.insert_formula("=AVERAGE(A1:A5)")),("IF",lambda:self.insert_formula("=IF(A1>0,1,0)")),("Синий",lambda:self.apply_format(font_color="FF1565C0")),("Жёлтый",lambda:self.apply_format(fill_color="FFFFF2A8")),("Зелёный",lambda:self.apply_format(fill_color="FFC8F7D2"))]:
            tools.add_widget(self._button(label,fn,dp(58)))
        scroller=ScrollView(size_hint_y=None, height=dp(44), do_scroll_x=True, do_scroll_y=False)
        tools.size_hint=(None,1); tools.width=dp(58)*len(tools.children)
        scroller.add_widget(tools); self.add_widget(scroller)
        self.grid=SheetGrid()
        grid_scroll=ScrollView(size_hint=(1,1), do_scroll_x=True, do_scroll_y=True,
                               bar_width=dp(5), scroll_type=["bars", "content"])
        grid_scroll.add_widget(self.grid)
        self.add_widget(grid_scroll)
        self.status_label=Label(text=self.tr("ready"),size_hint_y=None,height=dp(24),font_size="12sp")
        self.add_widget(self.status_label); self._apply_theme(); self.refresh_grid()
    def _button(self,text,fn,width):
        b=Button(text=text,size_hint_x=None,width=width,background_normal="",
                 background_down="",background_color=(0,0,0,0),
                 color=(1,1,1,1),bold=True)
        with b.canvas.before:
            Color(*COLORS["light"]["primary"])
            b._bg_rect = RoundedRectangle(pos=b.pos,size=b.size,radius=[dp(12)])
        b.bind(pos=lambda w,v:setattr(b._bg_rect,"pos",v),
               size=lambda w,v:setattr(b._bg_rect,"size",v))
        b.bind(on_release=lambda *_:fn())
        return b
    def refresh_grid(self):
        data=[{"text":"","role":"header","background_color":(.84,.90,.86,1)}]
        data += [{"text":self._col(c),"role":"header","background_color":(.84,.90,.86,1)}
                 for c in range(1,27)]
        for row in range(1,151):
            data.append({"text":str(row),"role":"header","background_color":(.88,.92,.89,1)})
            for col in range(1,27):
                addr=f"{self._col(col)}{row}"
                style=self.sheet.styles.get(addr,{})
                selected=(addr==self.selected_cell)
                data.append({"text":self.sheet.display(addr),"address":addr,"role":"cell",
                    "bold":style.get("bold",False),"italic":style.get("italic",False),
                    "underline":style.get("underline",False),
                    "color":self._rgba(style.get("font_color"),(.12,.16,.22,1)),
                    "background_color":((.82,.94,.86,1) if selected else
                        self._rgba(style.get("fill_color"),(1,1,1,1)))})
        self.grid.data=data
    @staticmethod
    def _rgba(value, fallback):
        if not value:
            return fallback
        raw = str(value).replace("#", "")
        if len(raw) == 8:
            raw = raw[-6:]
        if len(raw) != 6:
            return fallback
        try:
            return tuple(int(raw[i:i+2], 16)/255 for i in (0,2,4)) + (1,)
        except ValueError:
            return fallback
    @staticmethod
    def _col(n):
        s=""
        while n: n,r=divmod(n-1,26); s=chr(65+r)+s
        return s
    def select_cell(self,addr):
        self.selected_cell=addr
        self.cell_label.text=addr
        self.formula.text=self.sheet.get_raw(addr)
        self.refresh_grid()
        self.formula.focus = True
        self.formula.select_all()

    def edit_cell(self, addr):
        body=BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(10))
        editor=TextInput(text=self.sheet.get_raw(addr), multiline=False, size_hint_y=None,
                         height=dp(52), hint_text="Введите значение или формулу (=SUM(A1:A5))")
        body.add_widget(editor)
        actions=BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(8))
        cancel=Button(text="Отмена")
        apply=Button(text="Готово", background_normal="", background_color=COLORS["light"]["primary"],
                     color=(1,1,1,1))
        actions.add_widget(cancel)
        actions.add_widget(apply)
        body.add_widget(actions)
        popup=Popup(title="Ячейка " + addr, content=body, size_hint=(.94,None), height=dp(190),
                    auto_dismiss=False)
        cancel.bind(on_release=lambda *_:popup.dismiss())
        def save_value(*_):
            self.sheet.set(addr, editor.text)
            self.formula.text=editor.text
            self.refresh_grid()
            self.status_label.text=addr + " обновлена"
            popup.dismiss()
        apply.bind(on_release=save_value)
        editor.bind(on_text_validate=save_value)
        popup.open()
        Clock.schedule_once(lambda *_:setattr(editor,"focus",True),.2)
    def commit_value(self,value):
        self.sheet.set(self.selected_cell,value); self.refresh_grid()
        self.status_label.text=f"{self.selected_cell} обновлена"
    def merge_next_cell(self):
        import re
        match=re.fullmatch(r"([A-Z]+)([0-9]+)",self.selected_cell)
        if not match:return
        next_col=self._col(self._col_number(match.group(1))+1)
        self.sheet.merge(f"{self.selected_cell}:{next_col}{match.group(2)}")
        self.status_label.text="Объединение будет сохранено в XLSX"
    @staticmethod
    def _col_number(value):
        result=0
        for char in value: result=result*26+ord(char)-64
        return result
    def apply_format(self, **style):
        self.sheet.format_cells([self.selected_cell], **style)
        self.refresh_grid()
        self.status_label.text = "Формат применён к " + self.selected_cell
    def sort_sheet(self):
        try:
            self.sheet.sort_range("A1:J20", key_column=1)
            self.refresh_grid()
            self.status_label.text = "Диапазон A1:J20 отсортирован по столбцу A"
        except Exception as exc:
            self.message(self.tr("error"), str(exc))
    def filter_sheet(self):
        rows = self.sheet.filter_range("A1:J100", column=1)
        self.status_label.text = "Непустые строки в столбце A: " + str(len(rows))
    def fill_down(self):
        self.sheet.autofill("A1:A1", "A2:A20")
        self.refresh_grid()
        self.status_label.text = "Значение A1 скопировано в A2:A20"
    def open_file(self):
        try:
            from plyer import filechooser
            filechooser.open_file(on_selection=self._on_file_selected,filters=["*.xlsx","*.xls","*.csv"])
            self.status_label.text=self.tr("choose_file")
        except Exception as exc: self.message(self.tr("error"),str(exc))
    def _on_file_selected(self,selection):
        if not selection:return
        path=str(selection[0]); ext=Path(path).suffix.lower()
        try:
            if ext==".xlsx": self.sheet=Spreadsheet.load_xlsx(path)
            elif ext==".xls": self.sheet=Spreadsheet.load_xls(path)
            elif ext==".csv": self.sheet=Spreadsheet.load_csv(path)
            else: raise ValueError(self.tr("supported_formats"))
            self.refresh_grid(); self.select_cell("A1"); self.status_label.text=f"{self.tr('opened')}: {Path(path).name}"
        except Exception as exc:self.message(self.tr("open_error"),str(exc))
    def save_file(self):
        try:
            folder = Path(App.get_running_app().user_data_dir) / "documents"
            folder.mkdir(parents=True, exist_ok=True)
            doc = getattr(self, "current_document", None)
            if not doc:
                self.message("Сохранение", "Сначала создайте или откройте таблицу.")
                return
            path = folder / doc["file"]
            self.sheet.save_xlsx(path)
            self.status_label.text = "Сохранено: " + doc["name"]
        except Exception as exc:
            self.message(self.tr("error"), str(exc))
    def open_menu(self):
        content=BoxLayout(orientation="vertical",padding=dp(10),spacing=dp(8)); popup=Popup(title=self.tr("menu"),content=content,size_hint=(.9,.5))
        for label,fn in [(self.tr("settings"),self.open_settings),
            ("Листы",self.open_sheets),(self.tr("privacy"),lambda:self.show_document("PRIVACY_POLICY.md",self.tr("privacy"))),(self.tr("terms"),lambda:self.show_document("TERMS_OF_USE.md",self.tr("terms")))]:
            b=Button(text=label,size_hint_y=None,height=dp(48)); b.bind(on_release=lambda _,action=fn:(popup.dismiss(),action())); content.add_widget(b)
        popup.open()
    def open_sheets(self):
        body=BoxLayout(orientation="vertical",padding=dp(10),spacing=dp(6))
        popup=Popup(title="Листы книги",content=body,size_hint=(.85,.7))
        for name in self.sheet.sheet_names:
            button=Button(text=name,size_hint_y=None,height=dp(46))
            button.bind(on_release=lambda _,n=name:(self.sheet.select_sheet(n),self.refresh_grid(),self.select_cell("A1"),popup.dismiss()))
            body.add_widget(button)
        popup.open()
    def show_document(self,filename,title):
        p=Path(__file__).parent/filename; text=p.read_text(encoding="utf-8") if p.exists() else self.tr("document_missing")
        label=Label(text=text,size_hint_y=None,halign="left",valign="top",color=COLORS["light"]["text"],font_size="14sp")
        label.bind(width=lambda w,v:setattr(w,"text_size",(v-dp(24),None)))
        label.bind(texture_size=lambda w,v:setattr(w,"height",v[1]+dp(24)))
        scroll=ScrollView(); scroll.add_widget(label); Popup(title=title,content=scroll,size_hint=(.94,.88)).open()
    def show_consent(self):
        body = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(10))
        intro = Label(
            text=self.tr("consent_intro") + "\n\nВы можете ознакомиться с документами сейчас или позже в меню приложения.",
            size_hint_y=None, height=dp(82), halign="left", valign="middle",
            color=COLORS["light"]["text"]
        )
        intro.bind(width=lambda w, v: setattr(w, "text_size", (v, None)))
        body.add_widget(intro)

        docs = BoxLayout(orientation="vertical", spacing=dp(8))
        for key, filename in [("privacy", "PRIVACY_POLICY.md"), ("terms", "TERMS_OF_USE.md")]:
            button = Button(
                text=self.tr(key) + "  ·  Посмотреть",
                size_hint_y=None, height=dp(48),
                background_normal="", background_color=COLORS["light"]["primary"],
                color=(1, 1, 1, 1)
            )
            button.bind(on_release=lambda _, f=filename, k=key: self.show_document(f, self.tr(k)))
            docs.add_widget(button)
        body.add_widget(docs)

        note = Label(
            text="Читать документы перед продолжением необязательно.",
            size_hint_y=None, height=dp(34), font_size="12sp",
            color=COLORS["light"]["muted"]
        )
        body.add_widget(note)

        continue_button = Button(
            text=self.tr("agree"), size_hint_y=None, height=dp(50),
            background_normal="", background_color=COLORS["light"]["primary"],
            color=(1, 1, 1, 1)
        )
        body.add_widget(continue_button)
        popup = Popup(
            title=self.tr("consent_title"), content=body,
            size_hint=(.94, .58), auto_dismiss=False
        )
        continue_button.bind(on_release=lambda *_: (
            self.store.put("consent", accepted=True), popup.dismiss()
        ))
        popup.open()
    def open_settings(self):
        body=BoxLayout(orientation="vertical",padding=dp(12),spacing=dp(10))
        help_button=Button(text="Как пользоваться формулами? · Узнать больше",size_hint_y=None,height=dp(48))
        help_button.bind(on_release=lambda *_: self.show_formula_help())
        body.add_widget(help_button)
        body.add_widget(Label(text=self.tr("language"))); lang=Spinner(text=self.language,values=("ru","en"),size_hint_y=None,height=dp(44)); body.add_widget(lang)
        body.add_widget(Label(text=self.tr("theme"))); theme=Spinner(text=self.theme,values=("system","light","dark"),size_hint_y=None,height=dp(44)); body.add_widget(theme)
        save=Button(text=self.tr("save"),size_hint_y=None,height=dp(46)); reset=Button(text=self.tr("reset_consent"),size_hint_y=None,height=dp(46))
        body.add_widget(save); body.add_widget(reset); popup=Popup(title=self.tr("settings"),content=body,size_hint=(.9,.64))
        save.bind(on_release=lambda *_:(setattr(self,"language",lang.text),setattr(self,"theme",theme.text),self.store.put("settings",language=self.language,theme=self.theme),setattr(self,"translations",self._locale(self.language)),popup.dismiss(),self.show_library()))
        reset.bind(on_release=lambda *_:(self.store.delete("consent") if self.store.exists("consent") else None,popup.dismiss(),self.show_consent()))
        popup.open()
    def show_formula_help(self):
        guide = ("Формулы начинаются со знака =. Введите формулу в строку над таблицей и нажмите ✓.\\n\\n"
            "Ссылки: =A1+B1. Диапазон: =SUM(A1:A10).\\n\\n"
            "Математика: SUM, AVERAGE, MIN, MAX, COUNT, MEDIAN, STDEV, VAR, ABS, ROUND, SQRT, POWER.\\n"
            "Условия: IF, AND, OR, NOT, COUNTIF.\\n"
            "Текст: LEN, UPPER, LOWER, TRIM, LEFT, RIGHT, MID, CONCAT, CONCATENATE.\\n\\n"
            "Примеры: =SUM(B2:B8) — сумма; =AVERAGE(C1:C5) — среднее; "
            "=IF(A1>10,1,0) — условие; =ROUND(B1,2) — округление.\\n\\n"
            "Аргументы можно разделять запятыми или точками с запятой.")
        label=Label(text=guide,size_hint_y=None,halign="left",valign="top",color=COLORS["light"]["text"],font_size="14sp")
        label.bind(width=lambda w,v:setattr(w,"text_size",(v-dp(24),None)))
        label.bind(texture_size=lambda w,v:setattr(w,"height",v[1]+dp(24)))
        scroll=ScrollView(); scroll.add_widget(label)
        Popup(title="Справка по формулам",content=scroll,size_hint=(.94,.86)).open()

    def _apply_theme(self):
        mode = self.theme if self.theme in ("light", "dark") else "light"
        if self.theme == "system" and platform == "android":
            try:
                from jnius import autoclass
                Configuration = autoclass("android.content.res.Configuration")
                activity = autoclass("org.kivy.android.PythonActivity").mActivity
                ui_mode = activity.getResources().getConfiguration().uiMode
                mode = "dark" if (ui_mode & Configuration.UI_MODE_NIGHT_MASK) == Configuration.UI_MODE_NIGHT_YES else "light"
            except Exception:
                mode = "light"
        colors = COLORS[mode]
        Window.clearcolor = colors["bg"]
        if hasattr(self, "cell_label"):
            self.cell_label.color = colors["text"]
        if hasattr(self, "status_label"):
            self.status_label.color = colors["muted"]
        if hasattr(self, "formula"):
            self.formula.foreground_color = colors["text"]
            self.formula.background_color = colors["bg"]
    def message(self,title,message): Popup(title=title,content=Label(text=message),size_hint=(.9,.4)).open()

class ListSpreadsheetApp(App):
    def build(self):
        root=RootView()
        if root.store.exists("settings"):
            s=root.store.get("settings"); root.language=s.get("language","ru"); root.theme=s.get("theme","system"); root.translations=root._locale(root.language)
        else: root.translations=root._locale("ru")
        root._build_ui()
        if not root.store.exists("consent") or not root.store.get("consent").get("accepted", False):
            Clock.schedule_once(lambda *_: root.show_consent(), 0)
        return root

if __name__=="__main__": ListSpreadsheetApp().run()
