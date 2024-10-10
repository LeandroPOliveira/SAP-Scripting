import pandas as pd
import os
import numpy as np
import pyperclip
from time import sleep

dados = pd.read_excel('CJIC.xlsx')

for i, v in enumerate(dados['numero'].unique()):
    if i > 0:
        print(i)
        del dados['copiar']
    dados['copiar'] = np.where(dados['numero'] == v, dados['diagrama'], '')

    dados['copiar'].to_clipboard(index=False, header=False)
    print(dados['copiar'])
    arquivo = open(f'CJIC.vbs', 'w')

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
session.findById("wnd[0]/tbar[0]/okcd").text = "CJIC"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = "RSG.24.001.024.1.05.2"
session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/IMOB_OBJETO"
session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus
session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 12
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"POBID"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "POBID"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/tbar[0]/btn[24]").press
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = ""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[13]").press
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,0]").text = "{v}"
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,0]").setFocus
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,0]").caretPosition = 0
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]").sendVKey 3
''')

    arquivo.close()

    os.startfile('CJIC.vbs')
    sleep(3)

