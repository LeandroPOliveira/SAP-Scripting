# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('Script1.vbs', 'a')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('teste.xlsx')

# iterar sobre as linhas do arquivo excel e buscar os dados necessários para o script
for index, row in dados.iterrows():
    if row['Data'][6:10] != row['Incorp'][6:10]:
    # Adicionar os dados ao script
        arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "AS02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtANLA-ANLN1").text = "{row['Imob']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").text = "{row['Sub']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").setFocus
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").text = "{row['Data']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").caretPosition = 10
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 3
''')
    else:
        arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "AS02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtANLA-ANLN1").text = "{row['Imob']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").text = "{row['Sub']}"
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").setFocus
session.findById("wnd[0]/usr/ctxtANLA-ANLN2").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").text = "{row['Data']}"
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").setFocus
session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPLAIST:1142/ctxtANLA-AKTIV").caretPosition = 10
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 3
        ''')

# Fechar o arquivo de script
arquivo.close()

