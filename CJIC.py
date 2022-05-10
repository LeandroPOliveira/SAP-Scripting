# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('CJIC.vbs', 'a')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('CJIC.xlsx', dtype=str)



# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    # Adicionar os dados ao script
        arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "CJIC"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = "RSG.22.001.022.1.01.1"
session.findById("wnd[0]/usr/ctxtCN_ACTVT-LOW").text = ""
session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/IMOB_OBJETO"
session.findById("wnd[0]/usr/ctxtP_DISVAR").setFocus
session.findById("wnd[0]/usr/ctxtP_DISVAR").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"POBID"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "POBID"
session.findById("wnd[0]/tbar[1]/btn[29]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "7010083 0010"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectAll
session.findById("wnd[0]/tbar[1]/btn[13]").press
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,0]").text = "601201-0"
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtCOBRB-URZUO[7,0]").text = "4"
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-AQZIF[4,0]").setFocus
session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/txtCOBRB-AQZIF[4,0]").caretPosition = 0
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
''')

# Fechar o arquivo de script
arquivo.close()
