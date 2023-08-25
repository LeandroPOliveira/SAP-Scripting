# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd
from time import sleep
import os


# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('OASV.xlsx', dtype=str)

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    arquivo = open(f'OASV.vbs', 'w')
    # Adicionar os dados ao script
    arquivo.write(f'''

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "OASV"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRA01B-BLDAT").text = "{row['data']}"
session.findById("wnd[0]/usr/ctxtRA01B-BUDAT").text = "{row['data']}"
session.findById("wnd[0]/usr/txtRA01B-MONAT").text = "{row['mes']}"
session.findById("wnd[0]/usr/ctxtRA01B-BLART").text = "{row['doc']}"
session.findById("wnd[0]/usr/ctxtRA01B-LDGRP").setFocus
session.findById("wnd[0]/usr/ctxtRA01B-LDGRP").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-HKONT[0,0]").text = "{row['conta1']}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-HKONT[0,1]").text = "{row['conta2']}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/cmbRA01B-SHKZG[2,0]").key = "H"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/cmbRA01B-SHKZG[2,1]").key = "S"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-DMBTR[3,0]").text = "{row['valor'].replace('.', ',')}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-DMBTR[3,1]").text = "{row['valor'].replace('.', ',')}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-SGTXT[7,0]").text = "{row['hist1']}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-SGTXT[7,1]").text = "{row['hist2']}"
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-DMBTR[3,2]").setFocus
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110/txtRA01B-DMBTR[3,2]").caretPosition = 0
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tblSAPMA03BTCTRL_0110").verticalScrollbar.position = 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

    ''')

# Fechar o arquivo de script
    arquivo.close()

    os.startfile('OASV.vbs')
    sleep(4)


