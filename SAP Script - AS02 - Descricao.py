# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('descricao.vbs', 'a')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('teste - descricao.xlsx')

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    # Adicionar os dados ao script
    arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "as02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtANLA-ANLN1").text = "{row['Imob']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").text = "{row['Sub']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").setFocus
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXT50").text = "{row['Descricao']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLA-TXA50").text = "{row['livre']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLAIST:1140/txtANLH-ANLHTXT").text = "{row['adicional']}"
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
''')
