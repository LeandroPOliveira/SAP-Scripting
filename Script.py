import os

import click
import win32com.client
import subprocess
import sys
import time
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


class ContentNavigationDrawer(Screen):
    pass


class Principal(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
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


class TcodeFB08(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.mes = '08'


    def tcode_fb08(self):

        self.arquivo = f'FB08-{self.mes}.xlsx'

        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text = "1120170000"
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = "01.11.2022"
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = "30.11.2022"
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 10
        self.session.findById("wnd[0]").sendVKey(8)
        self.session.findById("wnd[0]/usr/lbl[48,5]").setFocus()
        self.session.findById("wnd[0]/usr/lbl[48,5]").caretPosition = 7
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById("wnd[0]/tbar[1]/btn[32]").press()
        self.session.findById(
            "wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,4]").text = "15"
        self.session.findById(
            "wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,4]").setFocus()
        self.session.findById(
            "wnd[1]/usr/tabsTS_LINES/tabpLI01/ssubSUB810:SAPLSKBH:0810/tblSAPLSKBHTC_WRITE_LIST/txtGT_WRITE_LIST-OUTPUTLEN[2,4]").caretPosition = 2
        self.session.findById("wnd[1]").sendVKey(0)
        self.session.findById("wnd[0]/usr").verticalScrollbar.position = 46
        self.session.findById("wnd[0]/usr/lbl[48,5]").caretPosition = 8
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById("wnd[0]/usr").verticalScrollbar.position = 0
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.getcwd()
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.arquivo
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/usr/lbl[48,26]").setFocus()
        self.session.findById("wnd[0]/usr/lbl[48,26]").caretPosition = 13
        self.session.findById("wnd[0]").sendVKey(3)
        self.session.findById("wnd[0]").sendVKey(3)

        self.ids.dir_pasta.text = os.path.join(os.getcwd(), self.arquivo)


class WindowManager(ScreenManager):
    pass


class SapScript(MDApp):
    # Window.maximize()
    tamanho_tela = Window.size

    def build(self):
        return Builder.load_file('Script.kv')


SapScript().run()
