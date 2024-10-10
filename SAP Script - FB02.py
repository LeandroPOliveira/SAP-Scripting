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
session.findById("wnd[0]/tbar[0]/okcd").text = "FB03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtRF05L-BELNR").text = "{row['doc']}"
session.findById("wnd[0]/usr/ctxtRF05L-BUKRS").text = "{row['emp']}"
session.findById("wnd[0]/usr/txtRF05L-GJAHR").text = "{row['exe']}"
session.findById("wnd[0]/usr/txtRF05L-GJAHR").setFocus
session.findById("wnd[0]/usr/txtRF05L-GJAHR").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\loliveira\Desktop\\testes\Ativo de contrato"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "{row['for']} {row['num']}.XLSX"
session.findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3
    ''')

# Fechar o arquivo de script
arquivo.close()