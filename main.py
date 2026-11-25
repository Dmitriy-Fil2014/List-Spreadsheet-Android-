import json
from pathlib import Path
from kivy.app import App
from kivy.clock import Clock
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
        return super().refresh_view_attrs(rv,index,data)
    def on_release(self):
        if self.role=="cell" and self.address: App.get_running_app().root.select_cell(self.address)

class SheetGrid(RecycleView):
    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        self.viewclass = CellButton
        layout=RecycleGridLayout(cols=11,default_size=(dp(104),dp(42)),default_size_hint=(None,None),
            size_hint=(None,None),width=dp(104)*11,spacing=dp(1))
        layout.bind(minimum_height=layout.setter("height")); self.add_widget(layout)

class RootView(BoxLayout):
    selected_cell=StringProperty("A1")
    def __init__(self,**kwargs):
        super().__init__(orientation="vertical",padding=dp(10),spacing=dp(8),**kwargs)
        self.sheet=Spreadsheet(rows=100,cols=26)
        self.store=JsonStore(str(Path(App.get_running_app().user_data_dir)/"preferences.json"))
        self.language="ru"; self.theme="system"; self.translations={}
        self._build_ui(); self._show_splash()
    def tr(self,key): return self.translations.get(key,key)
    def _locale(self,code):
        try: return json.loads((Path(__file__).parent/"locales"/(code+".json")).read_text(encoding="utf-8"))
        except (OSError,ValueError): return {}
    def _build_ui(self):
        self.clear_widgets()
        top=BoxLayout(size_hint_y=None,height=dp(50),spacing=dp(8))
        title=Label(text="Лист Таблицы",bold=True,font_size="20sp",color=COLORS["light"]["text"],halign="left")
        top.add_widget(title); top.add_widget(self._button(self.tr("open"),self.open_file,dp(112)))
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
                         ("U",lambda:self.apply_format(underline=True)),("A→Z",self.sort_sheet),
                         ("Фильтр",self.filter_sheet),("Автозап.",self.fill_down)]:
            tools.add_widget(self._button(label,fn,dp(58)))
        self.add_widget(tools)
        self.grid=SheetGrid(); self.add_widget(self.grid)
        self.status_label=Label(text=self.tr("ready"),size_hint_y=None,height=dp(24),font_size="12sp")
        self.add_widget(self.status_label); self._apply_theme(); self.refresh_grid()
    def _button(self,text,fn,width):
        b=Button(text=text,size_hint_x=None,width=width,background_normal="",background_color=COLORS["light"]["primary"])
        b.bind(on_release=lambda *_:fn()); return b
    def _show_splash(self):
        from kivy.uix.image import Image
        from kivy.uix.anchorlayout import AnchorLayout
        splash=BoxLayout(orientation="vertical",padding=dp(24),spacing=dp(16))
        splash.add_widget(Label())
        anchor=AnchorLayout(anchor_x="center",anchor_y="center",size_hint_y=None,height=dp(190))
        anchor.add_widget(Image(source=str(Path(__file__).parent/"assets"/"icon.png"),size_hint=(None,None),size=(dp(180),dp(180))))
        splash.add_widget(anchor)
        splash.add_widget(Label(text="Лист Таблицы",bold=True,font_size="24sp",size_hint_y=None,height=dp(44)))
        self.clear_widgets(); self.add_widget(splash); Clock.schedule_once(lambda *_:self._finish_splash(),1.0)
    def _finish_splash(self):
        self._build_ui()
        if not self.store.exists("consent") or not self.store.get("consent").get("accepted",False): self.show_consent()
    def refresh_grid(self):
        data=[{"text":"","role":"header","background_color":(.88,.93,.90,1)}]
        data += [{"text":self._col(c),"role":"header","background_color":(.88,.93,.90,1)} for c in range(1,11)]
        for row in range(1,501):
            data.append({"text":str(row),"role":"header","background_color":(.91,.94,.92,1)})
            for col in range(1,11):
                addr=f"{self._col(col)}{row}"
                style = self.sheet.styles.get(addr, {})
                data.append({"text":self.sheet.display(addr),"address":addr,"role":"cell",
                             "bold":style.get("bold",False),"italic":style.get("italic",False),
                             "underline":style.get("underline",False),
                             "color":self._rgba(style.get("font_color"),(0.12,0.16,0.22,1)),
                             "background_color":self._rgba(style.get("fill_color"),(1,1,1,1))})
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
        self.selected_cell=addr; self.cell_label.text=addr; self.formula.text=self.sheet.get_raw(addr)
    def commit_value(self,value):
        self.sheet.set(self.selected_cell,value); self.refresh_grid()
        self.status_label.text=f"{self.selected_cell} обновлена"
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
        path=Path(App.get_running_app().user_data_dir)/"таблица.xlsx"
        try:self.sheet.save_xlsx(path); self.status_label.text=f"{self.tr('saved')}: {path.name}"
        except Exception as exc:self.message(self.tr("error"),str(exc))
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
        label=Label(text=text,size_hint_y=None,halign="left",valign="top",color=COLORS["light"]["text"])
        label.bind(width=lambda w,v:setattr(w,"text_size",(v-dp(24),None)))
        label.bind(texture_size=lambda w,v:setattr(w,"height",v[1]+dp(24)))
        scroll=ScrollView(); scroll.add_widget(label); Popup(title=title,content=scroll,size_hint=(.94,.88)).open()
    def show_consent(self):
        body=BoxLayout(orientation="vertical",padding=dp(10),spacing=dp(8)); body.add_widget(Label(text=self.tr("consent_intro"),size_hint_y=None,height=dp(56)))
        read={"privacy":False,"terms":False}; checks={}
        for key,filename in [("privacy","PRIVACY_POLICY.md"),("terms","TERMS_OF_USE.md")]:
            row=BoxLayout(size_hint_y=None,height=dp(48)); btn=Button(text=self.tr(key+"_read")); check=CheckBox(disabled=True,size_hint_x=None,width=dp(48)); checks[key]=check
            def open_doc(_,k=key,f=filename):
                self.show_document(f,self.tr(k)); read[k]=True; checks[k].disabled=False
            btn.bind(on_release=open_doc); row.add_widget(btn); row.add_widget(check); body.add_widget(row)
        agree=Button(text=self.tr("agree"),disabled=True,size_hint_y=None,height=dp(48)); body.add_widget(agree)
        def validate(*_): agree.disabled=not(read["privacy"] and read["terms"] and all(c.active for c in checks.values()))
        for c in checks.values(): c.bind(active=validate)
        popup=Popup(title=self.tr("consent_title"),content=body,size_hint=(.94,.56),auto_dismiss=False)
        agree.bind(on_release=lambda *_:(self.store.put("consent",accepted=True),popup.dismiss()))
        popup.open()
    def open_settings(self):
        body=BoxLayout(orientation="vertical",padding=dp(12),spacing=dp(10))
        body.add_widget(Label(text=self.tr("language"))); lang=Spinner(text=self.language,values=("ru","en"),size_hint_y=None,height=dp(44)); body.add_widget(lang)
        body.add_widget(Label(text=self.tr("theme"))); theme=Spinner(text=self.theme,values=("system","light","dark"),size_hint_y=None,height=dp(44)); body.add_widget(theme)
        save=Button(text=self.tr("save"),size_hint_y=None,height=dp(46)); reset=Button(text=self.tr("reset_consent"),size_hint_y=None,height=dp(46))
        body.add_widget(save); body.add_widget(reset); popup=Popup(title=self.tr("settings"),content=body,size_hint=(.9,.64))
        save.bind(on_release=lambda *_:(setattr(self,"language",lang.text),setattr(self,"theme",theme.text),self.store.put("settings",language=self.language,theme=self.theme),setattr(self,"translations",self._locale(self.language)),popup.dismiss(),self._build_ui()))
        reset.bind(on_release=lambda *_:(self.store.delete("consent") if self.store.exists("consent") else None,popup.dismiss(),self.show_consent()))
        popup.open()
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
        return root

if __name__=="__main__": ListSpreadsheetApp().run()
