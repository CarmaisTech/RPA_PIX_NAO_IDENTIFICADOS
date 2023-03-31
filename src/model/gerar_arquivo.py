from .clickApi import click2 as c
import pandas as pd
from pdf2image import convert_from_path # Converte PDF em PNG
import PyPDF2# Manipula PDF's
import pyautogui as p
import xlrd
import logging

def gerar_arquivo(name):

    preview = p.locateCenterOnScreen('C:/RPA/arquivos/images/preview.png', confidence = 0.95)
    if preview != None:
        c(preview.x, preview.y )
    p.sleep(1)

    logging.info(f'Gerando Previsualizacao para relatorio da {name}')
    print(f'Gerando Previsualizacao para relatorio da {name}')

    tela_preview = p.locateCenterOnScreen('C:/RPA/arquivos/images/tela_preview_rgc.png', confidence = 0.95)
    while tela_preview:

        p.sleep(0.5)
        tela_preview = p.locateCenterOnScreen('C:/RPA/arquivos/images/tela_preview_rgc.png', confidence = 0.95)
    p.sleep(1)    

    nenhum_relatorio = p.locateCenterOnScreen('C:/RPA/arquivos/images/nenhum_relatorio2.png', confidence = 0.97)
    if nenhum_relatorio == None:

        logging.info(f'Existe relatorio para {name}')
        print(f'Existe relatorio para {name}')   
        logging.info(f'Gerar arquivo xls')
        print(f'Gerar arquivo xls')  

        p.sleep(1)

        imprimir = p.locateCenterOnScreen('C:/RPA/arquivos/images/imprimir.png', confidence = 0.95)
        while imprimir == None:
            p.sleep(0.5)
            imprimir = p.locateCenterOnScreen('C:/RPA/arquivos/images/imprimir.png', confidence = 0.95)
            print('Procurando imprimir')
        if imprimir != None:
            c(imprimir.x, imprimir.y )

            p.sleep(2)

        arquivo = p.locateCenterOnScreen('C:/RPA/arquivos/images/arquivo.png', confidence = 0.97)
        while arquivo == None:
            p.sleep(0.5)
            arquivo = p.locateCenterOnScreen('C:/RPA/arquivos/images/arquivo.png', confidence = 0.97)
            print('Procurando arquivo')
        p.sleep(0.5)
        if arquivo != None:
            c(arquivo.x, arquivo.y)


        p.sleep(1)
        
        excel_2000 = p.locateCenterOnScreen('C:/RPA/arquivos/images/excel_reward.png', confidence = 0.95)
        if excel_2000 != None:
            c(excel_2000.x, excel_2000.y )

        p.sleep(1)
        ok = p.locateCenterOnScreen('C:/RPA/arquivos/images/ok.png', confidence = 0.95)
        if ok != None:
            c(ok.x, ok.y )

            p.sleep(1)


        p.sleep(1)
        p.press('delete')
        p.typewrite(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xls')
        p.sleep(1)
        p.press('enter')

        fechar_relatorio = p.locateCenterOnScreen('C:/RPA/arquivos/images/fechar_relatorio.png', confidence = 0.95)
        if fechar_relatorio != None:
                
            c(fechar_relatorio.x , fechar_relatorio.y )

        p.sleep(1)
        logging.info(f'Arquivo {name} foi gerado.')
        print(f'Arquivo {name} foi gerado.')

        return True

    else:

        logging.info(f'Relatorio {name} sem dados.')
        print(f'Relatorio {name} sem dados.')

        ok = p.locateCenterOnScreen('C:/RPA/arquivos/images/ok.png', confidence = 0.70)
        if ok != None:
            c(ok.x , ok.y )
            p.sleep(1)

        return False