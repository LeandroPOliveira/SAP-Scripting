# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('IW32.vbs', 'w')

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('IW32.xlsx')

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
session.findById("wnd[0]/tbar[0]/okcd").text = "iw32"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = "{row['ordem']}"
session.findById("wnd[0]/tbar[1]/btn[30]").press
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtCOBRB-KONTY[0,1]").text = "ccs"
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,1]").text = "{row['centro']}"
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-PROZS[3,1]").setFocus
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-PROZS[3,1]").caretPosition = 0
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES").verticalScrollbar.position = 0
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 3
    ''')

# Fechar o arquivo de script
arquivo.close()