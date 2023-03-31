import pyautogui as p
from src.model.clickApi import click2 as c
from src.model.enviar_email import envia_email
from src.model.filtro_pix import filtroPix
from src.model.gerar_arquivo import gerar_arquivo
from src.model.loginDealer import login_dealer_contas_receber
from src.model.escolher_empresas import escolha_empresas
from src.model.menu_relatorio import  menu_relatorio
# from src.model.gerar_relatorio import gerar_relatorio
import os
import logging
import traceback


# Lendo o número do processo no arquivo pid_bot_running.txt
with open('C:/RPA/Credenciais/pid_bot_running.txt', 'r') as file:
    pid_bot_anterior = file.readlines()[0] # Lendo o conteudo do arquivo - Numero do processo do bot anterior
    os.system(f'taskkill /im {pid_bot_anterior} /F') # Forcando o encerramento do bot anterior

# Gravando o número do processo do bot atual no arquivo pid_bot_running.txt
with open('C:/RPA/Credenciais/pid_bot_running.txt', 'w') as file:
    file.write(str(os.getpid())) # sobrescrevendo o numero do PID

# Zerando o arquivo de log
with open("C:\\RPA\\RPA_PIX_NAO_IDENTIFICADO\\log.txt",'w') as f:
    pass

# - Configurando o Logging
logging.basicConfig(filename="C:\\RPA\\RPA_PIX_NAO_IDENTIFICADO\\log.txt",level=logging.INFO, format=' %(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%d/%m/%Y %I:%M:%S %p')
logging.info('Inicio do Programa')

list_empresas = [
    'NOVALUZ BS', 
    'NOVALUZ SD', 
    'NOVALUZ WS',
    'NOVALUZ SUL',
    'CONTERRANEA MATRIZ',
    'CONTERRANEA MOSSORO',
    'CONTERRANEA MACAIBA',
    'NOSSAMOTO MATRIZ',
    'NOSSAMOTO CAUCAIA',
    'NOSSAMOTO BATURITE',
    'NOSSAMOTO SIQUEIRA',
    'JANGADA VEICULOS',
    'NISSAN MATRIZ',
    'NISSAN FILIAL',
    'SANAUTO',
    'VOUGA',
    'CDA',
    'JANGADA AUTOMOTIVE',
    'NATIVA'
]

print('Encerrando processos do Dealer...')
logging.info('Encerrando processos do Dealer...') 

os.system("taskkill /im sof.exe")
os.system("taskkill /im ead.exe")
os.system("taskkill /im scr.exe")
os.system("taskkill /im scp.exe")
os.system("taskkill /im doro.exe")

print('Processos encerrados.')
logging.info('Processos encerrados.')

try:

    print('Efetuando login no Dealer Contas a Receber...')
    logging.info('Efetuando login no Dealer Contas a Receber...')
    login_dealer_contas_receber()

    print('Preparando extração de Relatórios...')
    logging.info('Preparando extração de Relatórios...') 
    menu_relatorio()
    p.sleep(1)

    for name in list_empresas:

        p.sleep(1)
        print(f'Selecionando Empresas {name}...')
        logging.info(f'Selecionando Empresas {name}...')      
        escolha_empresas(name)
        p.sleep(1)
        print('Extraindo relatorio...')
        logging.info('Extraindo relatorio...')
        tem_arquivo = gerar_arquivo(name)
        if tem_arquivo:
            print(f'Existe relatorio para a empresa - {name}')
            logging.info(f'Existe relatorio para a empresa - {name}')            
            tem_pix = filtroPix(name)
            if tem_pix:
                print(f'Existe PIX para a empresa - {name}')
                logging.info(f'Existe PIX para a empresa - {name}')                 
                envia_email(name)
            else:
                print(f'Nao existe PIX para a empresa - {name}')
                logging.info(f'Nao existe PIX para a empresa - {name}')
                os.remove(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xls')
                os.remove(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx')                
        
    fechar = p.locateCenterOnScreen('C:/RPA/arquivos/images/fechar.png', confidence = 0.95)
    if fechar != None:
            
        c(fechar.x , fechar.y )

    p.sleep(1)

    fechar2 = p.locateCenterOnScreen('C:/RPA/arquivos/images/fechar2.png', confidence = 0.98)
    if fechar2 != None:
            
        c(fechar2.x , fechar2.y )


except Exception:

    msg_error = traceback.format_exc()
    logging.info(f'{msg_error}')
    print(f'{msg_error}')


finally:

    p.sleep(2)
    # Encerrando o Dealer
    os.system("taskkill /im scr.exe")
    print('Concluído.')
    logging.info('Concluído.')
    # Encerrando o Dealer
    os.system("taskkill /im scr.exe")
       