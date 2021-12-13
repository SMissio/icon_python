import xlsxwriter as opcoesDOXL

import os

nomeArquivo = 'C:\\Users\\SylviaMissio\\Desktop\\RPA1\\xlsx\\FormatacaoCondicionalIcones.xlsx'
plan1 = opcoesDOXL.Workbook(nomeArquivo)

sheetDados = plan1.add_worksheet("Dados")


inserirDados = [
["Coluna 1", "Coluna 2", "Coluna 3", "Coluna 4"],
[100,102,156,138],
[98,85,157,177],
[88,21,105,74],
[106,58,73,170],
[18,30,77,170],
]

sheetDados.write('A1',"Exemplo formatação condicional com icones")

for lin, range in enumerate(inserirDados):
    sheetDados.write_row (lin+2,1,range)
    
sheetDados.conditional_format('B4:E4',{'type':'icon_set',
                                      'icon_style':'3_traffic_lights'})

sheetDados.conditional_format('B5:E5',{'type':'icon_set',
                                      'icon_style':'3_traffic_lights',
                                      'reverse_icons': True})

sheetDados.conditional_format('B6:E6',{'type':'icon_set',
                                      'icon_style':'3_arrows'})

sheetDados.conditional_format('B7:E7',{'type':'icon_set',
                                      'icon_style':'4_arrows'})

sheetDados.conditional_format('B8:E8',{'type':'icon_set',
                                      'icon_style':'5_ratings'})



plan1.close()

os.startfile(nomeArquivo)