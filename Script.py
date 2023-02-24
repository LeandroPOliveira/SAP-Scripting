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
from tika import parser


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
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = os.environ['usuario']
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = os.environ['senha']
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
        # self.data = self.manager.get_screen('principal').periodo()

    def tcode_fb08(self):
        self.session = self.manager.get_screen('principal').session
        self.data = self.manager.get_screen('principal').periodo()
        self.arquivo = f'FB08 {self.data[2].month}-{self.data[2].year}.xlsx'
        print(self.arquivo)
        self.session.findById("wnd[0]").maximize()
        self.session.createSession()
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
            self.session.findById("wnd[0]/usr/ctxtBKPF-BUKRS").text = "GBD1"
            self.session.findById("wnd[0]/usr/txtRF05A-GJAHS").text = "2022"
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


class TcodeFBL3N(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        # self.arquivo_energia = os.path.join(os.getcwd(), f'Energia {self.data[2].month}-{self.data[2].year}.xlsx')

    def tcode_fbl3n(self):
        self.session = self.manager.get_screen('principal').session
        self.data = self.manager.get_screen('principal').periodo()
        self.arquivo_energia = f'Energia {self.data[2].month}-{self.data[2].year}.xlsx'

        self.session.findById("wnd[0]").maximize()
        self.session.createSession()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
        self.session.findById("wnd[0]").sendVKey(0)
        # self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,0]").text = "6160322020"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,1]").text = "6160122020"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,2]").text = "6151122020"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,2]").setFocus()
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,2]").caretPosition = 0
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = self.data[0]
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = self.data[1]
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/DESPESA SEG"
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 12
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.getcwd()
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.arquivo_energia
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]").sendVKey(3)
        self.session.findById("wnd[0]").sendVKey(3)

    def rateio_energia(self):
        razao = pd.read_excel(self.arquivo_energia, sheet_name=0)
        razao['Nota'] = razao['Texto'].str.slice(10, 19)

        dados = [[], [], [], []]
        meses = {1: '01 - JANEIRO', 2: '02 - FEVEREIRO', 3: '03 - MARÇO', 4: '04 - ABRIL', 5: '05 - MAIO',
                 6: '06 - JUNHO', 7: '07 - JULHO', 8: '08 - AGOSTO', 9: '09 - SETEMBRO', 10: '10 - OUTUBRO',
                 11: '11 - NOVEMBRO', 12: '12 - DEZEMBRO'}

        dir_energia = f'G:\GECOT\\NOTAS FISCAIS DIGITALIZADAS\\{self.data[2].year}\\{meses[self.data[2].month]}\ENERGIA ELÉTRICA'
        for nota in os.listdir(dir_energia):

            if nota.endswith('.pdf'):
                conta = parser.from_file(os.path.join(dir_energia, nota))
                linha_conta = conta['content'].splitlines()
                outros_deb = 0
                for index, row in enumerate(linha_conta):
                    if 'Série C' in row:
                        dados[3].append(linha_conta[index].split(' ')[1]) if linha_conta[index].split(' ')[1] \
                                                                             not in dados[3] else None
                    if 'CNPJ' in row:
                        dados[0].append(linha_conta[index - 4])
                        dados[1].append(linha_conta[index - 2][10:].split('-')[0].strip())
                    if 'DÉBITOS' in row:
                        outros_deb = float(linha_conta[index + 2].split(' ')[6].replace(',', '.'))
                    if 'Total a Pagar (R$)' in row:
                        vr_total = linha_conta[index + 1].strip().replace('.', '')
                        try:
                            vr_total = float(vr_total.replace(',', '.'))
                        except ValueError:
                            vr_total = 0.00
                        imposto = (vr_total - outros_deb) * 0.0925
                        vr_a_pagar = vr_total - imposto
                        dados[2].append(round(vr_a_pagar, 2))

        dados = pd.DataFrame(dados).T
        dados.columns = ['Endereco', 'Cidade', 'Valor', 'Nota']

        dados_a_completar = pd.merge(razao, dados[['Nota', 'Endereco', 'Cidade']], on=['Nota'], how='left')

        # dados.to_excel('energia.xlsx', index=False)
        dados_a_completar.to_excel('energia.xlsx')


class ExtrairSAP(Screen):
    def extrair(self):
        self.session = self.manager.get_screen('principal').session
        self.data = self.manager.get_screen('principal').periodo()
        self.arquivo_pis = f'PIS {self.data[2].month}-{self.data[2].year}.xlsx'

        self.session.createSession()
        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
        self.session.findById("wnd[0]").sendVKey(0)
        # self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "6120102001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "6122102001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "6123102001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "6124102001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "6126102001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "6127302001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,6]").text = "6121012001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = "6121112001"
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,2]").setFocus()
        self.session.findById(
            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-"
            "SLOW_I[1,2]").caretPosition = 0
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/radX_AISEL").select()
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = self.data[0]
        self.session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = self.data[1]
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").text = "Z/ESTORNO"
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
        self.session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 12
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/lbl[20,5]").setFocus()
        self.session.findById("wnd[0]/usr/lbl[20,5]").caretPosition = 4
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById("wnd[0]/tbar[1]/btn[41]").press()
        self.session.findById("wnd[0]/usr/lbl[9,5]").setFocus()
        self.session.findById("wnd[0]/usr/lbl[9,5]").caretPosition = 6
        self.session.findById("wnd[0]").sendVKey(2)
        self.session.findById("wnd[0]/tbar[1]/btn[41]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.getcwd()
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.arquivo_pis
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]").sendVKey(3)
        self.session.findById("wnd[0]").sendVKey(3)


class WindowManager(ScreenManager):
    pass


class SapScript(MDApp):
    # Window.maximize()
    tamanho_tela = Window.size

    def build(self):
        return Builder.load_file('Script.kv')


SapScript().run()
