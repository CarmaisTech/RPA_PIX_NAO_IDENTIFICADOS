from .clickApi import click2 as c
import pandas as pd
import pyautogui as p
import xlrd


def filtroPix(name):

    file = f"C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xls"
    # Convertendo para xlsx
    wb = xlrd.open_workbook(f"C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xls", encoding_override='latin1')
    df = pd.read_excel(wb)
    df.to_excel(f"C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx", index = False)

    file = f"C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx"
    dados_excel = pd.read_excel(file)    

    tabela1 = pd.DataFrame(dados_excel)

    # print(tabela1['Historico'][2])
    contador = len(tabela1['Historico'])
    print('--------------------------------------')

    for i in range(contador):

        if "PIX" not in tabela1['Historico'][i] :
            tabela1 = tabela1.drop([i], axis=0 )


    print('--------------------------------------')
    # print(tabela1)
    tabela1.to_excel(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx', index = 0)

    df = pd.read_excel(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx')
    print(df)

    if len(df.index>0):
        return True
    else:
        return False
   