import os

import click
import win32com.client
import subprocess
import sys
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
from kivy.properties import StringProperty
from kivymd.app import MDApp
from kivymd.uix.button import MDFlatButton, MDRaisedButton
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
from kivy.core.window import Window
from kivy.utils import get_color_from_hex
from kivymd.uix.dialog import MDDialog
import pandas as pd


class ContentNavigationDrawer(Screen):
    pass


class Principal(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data_fim = None
        self.data_ini = None
        self.session = None

        try:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = self.SapGuiAuto.GetScriptingEngine
            self.connection = application.Children(0)
            self.session = self.connection.Children(0)

        except:
            self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
            subprocess.Popen(self.path)
            time.sleep(3)

            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")

            if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                return

            application = self.SapGuiAuto.GetScriptingEngine
            self.connection = application.OpenConnection("Gas Brasiliano - ECC 5.0 - E5P", True)
            time.sleep(1)
            self.session = self.connection.Children(0)

            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "200"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "loliveira"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "lpo;5159"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)

    def periodo(self):
        self.data_ini = self.ids.dt_ini.text.replace('/', '.')
        self.data_fim = self.ids.dt_fim.text.replace('/', '.')
        self.data_formatada = datetime.strptime(self.ids.dt_ini.text, '%d/%m/%Y')

        return self.data_ini, self.data_fim, self.data_formatada


class TcodeFB08(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.session = None

    def tcode_fb08(self):
        self.session = self.manager.get_screen('principal').session
        self.data = self.manager.get_screen('principal').periodo()
        self.arquivo = f'FB08 {self.data[2].month}-{self.data[2].year}.xlsx'
        print(self.arquivo)
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "1120170000"
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = self.data[0]
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = self.data[1]
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/DESPESA SEG"
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 12
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.getcwd()
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.arquivo
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]").sendVKey(3)
        self.session.findById("wnd[0]").sendVKey(3)

        self.ids.dir_pasta.text = os.path.join(os.getcwd(), self.arquivo)
        print(self.ids.dir_pasta.text)

    def estornar(self):
        self.session = self.manager.get_screen('principal').session
        dados = pd.read_excel(self.ids.dir_pasta.text, sheet_name=0, dtype=str)
        dados = dados[dados['Tipo de documento'] == 'SA']
        dados = dados[dados['Nº documento'].notnull()]
        for documento in dados['Nº documento'].unique():
            print(documento)
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "fb08"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtRF05A-BELNS").text = documento
            self.session.findById("wnd[0]/usr/ctxtUF05A-STGRD").text = "02"
            self.session.findById("wnd[0]/usr/ctxtBSIS-BUDAT").text = f"01.{self.data[2].month}.{self.data[2].year}"
            self.session.findById("wnd[0]/usr/txtBSIS-MONAT").text = self.data[2].month
            self.session.findById("wnd[0]/usr/ctxtRF05A-VOIDR").setFocus()
            self.session.findById("wnd[0]/usr/ctxtRF05A-VOIDR").caretPosition = 0
            self.session.findById("wnd[0]").sendVKey(11)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(3)
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


class WindowManager(ScreenManager):
    pass


class SapScript(MDApp):
    # Window.maximize()
    tamanho_tela = Window.size

    def build(self):
        return Builder.load_file('Script.kv')


SapScript().run()
