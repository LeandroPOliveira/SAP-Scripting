# Importar o pacote pandas para trabalhar com arquivos excel
import os

import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('AS01.vbs', 'w')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('AS01.xlsx', dtype=str)
dados.fillna('', inplace=True)


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
session.findById("wnd[0]/tbar[0]/okcd").text = "AS01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtANLA-ANLKL").text = "{row['classe']}"
session.findById("wnd[0]/usr/txtRA02S-NASSETS").text = "1"
session.findById("wnd[0]/usr/txtRA02S-NASSETS").setFocus
session.findById("wnd[0]/usr/txtRA02S-NASSETS").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/chkRA02S-XHIST").selected = true
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPLAIST:1141/chkANLA-INKEN").selected = true
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXT50").text = "{row['descricaon']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXA50").text = "{row['descricaoa']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLH-ANLHTXT").text = "{row['descricaol']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-MENGE").text = "{row['quantidade'].replace('.',',')}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/ctxtANLA-MEINS").text = "{row['um']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").text = "{row['inideprec']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").caretPosition = 10
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-GSBER").text = "0001"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTL").text = "11440"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTLV").text = "11440"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTLV").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1145/ctxtANLZ-KOSTLV").caretPosition = 5
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-ORD41").text = "{row['local']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-ORD42").text = "{row['sistema']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-ORD43").text = "(G)"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-GDLGRP").text = "{row['dn']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-IZWEK").text = "02"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-IZWEK").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1160/ctxtANLA-IZWEK").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1170/tblSAPLAISTTC_EQUI").getAbsoluteRow(0).selected = true
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1170/tblSAPLAISTTC_EQUI/chkRA02S-EQUI_WF[0,0]").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1170/btnDEQUI").press
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPLAIST:1181/chkRA02S-XNEU_AM").selected = true
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:SAPLATAB:0202/subAREA2:SAPLAIST:1182/ctxtANLA-POSNR").text = "{row['pep']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:SAPLATAB:0202/subAREA2:SAPLAIST:1182/ctxtANLA-POSNR").caretPosition = 21
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB06").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08").select
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,0]").text = "{row['dt1']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,4]").text = "{row['dt2']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,5]").text = "{row['dt3']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,5]").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPLAIST:1190/tblSAPLAISTTC_ANLB/ctxtANLB-AFABG[6,5]").caretPosition = 10
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]/tbar[0]/btn[3]").press
''')

# Fechar o arquivo de script
arquivo.close()

os.startfile('AS01.vbs')