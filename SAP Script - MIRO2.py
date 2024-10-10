# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('MIRO2.vbs', 'w')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('MIRO2.xlsx')

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
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = "{row['data1']}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").text = "00{row['refe']}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BUDAT").text = "{row['data2']}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").text = "{str(row['valor']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-WAERS").text = "{row['brl']}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-SGTXT").text = "NFSE 00{row['texto']} LEC BRASIL"

session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").text = "{row['folha']}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").setFocus
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY").select
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI").select
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE").text = "zs"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE").setFocus
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[21]").press
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").select
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/txtJ_1BDYDOC-AUTHCOD").text = "{row['chave']}"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/ctxtJ_1BDYDOC-AUTHDATE").text = "{row['datachave']}"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/ctxtJ_1BDYDOC-AUTHTIME").text = "{row['hora']}"
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/ctxtJ_1BDYDOC-AUTHTIME").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subTIMESTAMP:SAPLJ1BB2:2803/ctxtJ_1BDYDOC-AUTHTIME").caretPosition = 8
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT").select
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QSSHB[2,0]").text = "{str(row['baseiss']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QSSHB[2,2]").text = "{str(row['baseiss']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QSSHB[2,3]").text = "{str(row['baseiss']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").text = "{str(row['pcc']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,2]").text = "{str(row['ir']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,3]").text = "{str(row['iss']).replace('.',',')}"
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").setFocus
session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[2]").sendVKey 0

''')

# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").text = "{row['folha']}"
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").setFocus
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").caretPosition = 10

# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN").text = "{row['folha']}"
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN").setFocus
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN").caretPosition = 10
#
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QSSHB[2,0]").text = "{str(row['baseiss']).replace('.',',')}"
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").text = "{str(row['iss']).replace('.',',')}"
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").setFocus
# session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/txtACWT_ITEM-WT_QBSHB[3,0]").caretPosition = 5