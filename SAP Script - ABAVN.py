# Importar o pacote pandas para trabalhar com arquivos excel
import os

import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('ABAVN.vbs', 'w')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('ABAVN.xlsx', dtype=str)
dados.fillna('', inplace=True)
print(dados['data'])

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
''')

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    # Adicionar os dados ao script
        arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ABAVN"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subOBJECT:SAPLAMDP:0300/ctxtRAIFP2-ANLN1").text = "{row['imob']}"
session.findById("wnd[0]/usr/subOBJECT:SAPLAMDP:0300/ctxtRAIFP2-ANLN2").text = "{row['sub']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDP:0501/subSUBSCREEN1:SAPLAMDP:0200/ctxtRAIFP1-BLDAT").text = "{row['data']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDP:0501/subSUBSCREEN2:SAPLAMDP:0201/ctxtRAIFP1-BUDAT").text = "{row['data_lanc']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAMDP:0501/subSUBSCREEN3:SAPLAMDP:0202/ctxtRAIFP1-BZDAT").text = "{row['data_lanc']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDP:0506/subSUBSCREEN1:SAPLAMDP:0206/txtRAIFP2-SGTXT").text = "{row['texto']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDP:0506/subSUBSCREEN1:SAPLAMDP:0206/txtRAIFP2-SGTXT").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAMDP:0506/subSUBSCREEN1:SAPLAMDP:0206/txtRAIFP2-SGTXT").caretPosition = 21
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAMDP:0507/subSUBSCREEN1:SAPLAMDP:0203/txtRAIFP2-MONAT").text = "{row['period']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAMDP:0507/subSUBSCREEN1:SAPLAMDP:0203/txtRAIFP2-MONAT").caretPosition = 1
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDP:0401/txtRAIFP2-ANBTR").text = "{row['valor'].replace('.',',')}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDP:0401/txtRAIFP2-ANBTR").caretPosition = 8
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01").select
session.findById("wnd[0]/tbar[1]/btn[9]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 2,"HKONT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "2"
session.findById("wnd[0]").sendVKey 14
session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "11440"
session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

''')

# Fechar o arquivo de script
arquivo.close()


# session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDP:0401/radRAIFP2-XANEU").select
# session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAMDP:0401/radRAIFP2-XANEU").setFocus

# session.findById("wnd[0]/tbar[1]/btn[9]").press
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 2,"HKONT"
# session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "2"
# session.findById("wnd[0]").sendVKey 14
# session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "11330"
# session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").caretPosition = 5
# session.findById("wnd[1]/tbar[0]/btn[0]").press
# session.findById("wnd[0]/tbar[0]/btn[11]").press
# session.findById("wnd[0]").sendVKey 3