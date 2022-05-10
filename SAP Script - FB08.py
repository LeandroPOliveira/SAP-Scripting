# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('estorno.vbs', 'w')

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('estorno.xlsx')

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
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
session.findById("wnd[0]/tbar[0]/okcd").text = "fb08"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtRF05A-BELNS").text = "{row['Data']}"
session.findById("wnd[0]/usr/ctxtUF05A-STGRD").text = "02"
session.findById("wnd[0]/usr/ctxtBSIS-BUDAT").text = "01.04.2022"
session.findById("wnd[0]/usr/txtBSIS-MONAT").text = "4"
session.findById("wnd[0]/usr/ctxtRF05A-VOIDR").setFocus
session.findById("wnd[0]/usr/ctxtRF05A-VOIDR").caretPosition = 0
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    ''')

# Fechar o arquivo de script
arquivo.close()