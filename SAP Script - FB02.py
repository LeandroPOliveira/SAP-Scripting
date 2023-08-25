# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('FB02.vbs', 'w')

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('FB02.xlsx')

arquivo.write('''
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
''')

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    # Adicionar os dados ao script
    arquivo.write(f'''


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "fb02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtRF05L-BELNR").text = {row['Doc']}
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").currentCellColumn = "SGTXT"
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "{row['Hist']}"
session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").setFocus
session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = 11
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
    ''')

# Fechar o arquivo de script
arquivo.close()