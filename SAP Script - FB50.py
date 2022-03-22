# Importar o pacote pandas para trabalhar com arquivos excel
import pandas as pd

# Abrir arquivo de script gerado pelo SAP
arquivo = open('FB50.vbs', 'a')  # modo 'a' de append, insere novos dados no arquivo sem excluir os que estavam

# Abrir arquivo com os dados a serem lançados
dados_cab = pd.read_excel('FB50.xlsx', nrows=4, usecols=[0, 1], header=None)

arquivo.write(f'''
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "fb50"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSTRIP_HEAD/tabpTAB1/ssubHEAD:SAPMF05A:1010/ctxtACGL_HEAD-BLDAT").text ="{dados_cab.iloc[0][1]}"
session.findById("wnd[0]/usr/tabsTABSTRIP_HEAD/tabpTAB1/ssubHEAD:SAPMF05A:1010/ctxtACGL_HEAD-BUDAT").text ="{dados_cab.iloc[1][1]}"
session.findById("wnd[0]/usr/tabsTABSTRIP_HEAD/tabpTAB1/ssubHEAD:SAPMF05A:1010/txtACGL_HEAD-XBLNR").text ="{dados_cab.iloc[2][1]}"
session.findById("wnd[0]/usr/tabsTABSTRIP_HEAD/tabpTAB1/ssubHEAD:SAPMF05A:1010/txtACGL_HEAD-BKTXT").text ="{dados_cab.iloc[3][1]}"''')


# Abrir arquivo com os dados a serem lançados
dados = pd.read_excel('FB50.xlsx', skiprows=5)
dados.fillna('', inplace=True)
# print(dados)
for index, row in dados.iterrows():
    arquivo.write(f'''
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,{index}]").text = "{row['Cta.Razão']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/cmbACGL_ITEM-SHKZG[3,{index}]").key = {'"S"' if row['D/C'] == 'D' else '"H"'}
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,{index}]").text = "{row['Mont.em moeda doc.']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-SGTXT[11,{index}]").text = "{row['Texto']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-KOSTL[17,{index}]").text = "{row['Centro custo']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-AUFNR[18,{index}]").text = "{row['Ordem']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PRCTR[27,{index}]").text = "{row['Centro lucro']}"
session.findById("wnd[0]/usr/ssubITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-PROJK[29,{index}]").text = "{row['Elemen.PEP']}"
''')

arquivo.close()