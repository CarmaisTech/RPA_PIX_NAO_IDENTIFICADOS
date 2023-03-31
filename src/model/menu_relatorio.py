import pandas as pd
import pyautogui as p
from .clickApi import click2 as c
from .loginDealer import login_dealer_contas_receber
from datetime import date
import math
import os

def menu_relatorio():

    relatorios = p.locateCenterOnScreen('C:/RPA/arquivos/images/relatorios.png', confidence = 0.95)
    if relatorios != None:
        c(relatorios.x , relatorios.y )
    
    p.sleep(1)
    seta_relatorios_ead = p.locateCenterOnScreen('C:/RPA/arquivos/images/seta_relatorios_ead.png', confidence = 0.95)
    if seta_relatorios_ead != None:
        #c(seta_relatorios_ead.x , seta_relatorios_ead.y )
        p.click(seta_relatorios_ead)
    
    p.sleep(2)

    saldo_credito = p.locateCenterOnScreen('C:/RPA/arquivos/images/saldo_credito.png', confidence = 0.95)
    if saldo_credito != None:
        c(saldo_credito.x , saldo_credito.y )
    


    p.sleep(1)

    # # data_atual = date.today()
    # # ano_total= data_atual.strftime('%Y')
    # # ano = ano_total[2:]
    # data_inicial = '01/01/21'
    # data_final = '31/12/21'
    # # data_final = data_atual.strftime('%d%m') + ano 
    
    data_atual = date.today()
    data_em_texto_barr = data_atual.strftime('%d%m%Y')
    dia = data_em_texto_barr[:2]
    mes = data_em_texto_barr[2:4]
    ano = data_em_texto_barr[6:]
    ano_anterior = int(ano) - 1

    data_inicial = dia + mes + str(ano_anterior)
    data_final = dia + mes + ano
    # data_final = '28/02/22'

    data_ini = p.locateCenterOnScreen('C:/RPA/arquivos/images/data_inicial2.png', confidence = 0.95)
    if data_ini != None:         
        c(data_ini.x+45 , data_ini.y )
        p.sleep(1)
       
        
        p.sleep(0.5)
        p.press('delete', presses=8)
        p.press('home')
        p.sleep(0.5)
        p.typewrite(data_inicial)
        print(data_inicial)
        p.sleep(0.5)
        p.press('tab')
   
        
        p.typewrite(data_final)
        print(data_final)
    