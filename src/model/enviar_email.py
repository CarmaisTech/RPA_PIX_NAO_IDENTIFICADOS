import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
import pandas as pd
import os
import logging


credenciais = open('C:\\RPA\\credenciais\\credenciais_gmail.txt','r')
chaves = credenciais.readlines()
credenciais.close()

sender = chaves[0][:-1]
password = chaves[1]

# - Tuplas com os emails que deverão receber o relatório de cada empresa.
GRUPO_FIXO = ('\
    ,daniel@carmais.com.br,\
    jorge.nascimento@carmais.com.br,\
')

# FIAT
CDA_emails =   ('ricardo.pontes@carmais.com.br,') 
VOUGA_emails = ('ricardo.pontes@carmais.com.br,') 

# HONDA CARROS
NOVALUZ_BS_emails =  ('carlos.pedroso@carmais.com.br,')
NOVALUZ_WS_emails =  ('carlos.pedroso@carmais.com.br,')
NOVALUZ_SD_emails =  ('carlos.pedroso@carmais.com.br,')
NOVALUZ_SUL_emails = ('carlos.pedroso@carmais.com.br,')

# HONDA MOTOS
NOSSAMOTO_BATURITE_emails = ('tiago.lima@carmais.com.br,')
NOSSAMOTO_SIQUEIRA_emails = ('tiago.lima@carmais.com.br,')
NOSSAMOTO_MATRIZ_emails =   ('tiago.lima@carmais.com.br,')
NOSSAMOTO_CAUCAIA_emails =  ('tiago.lima@carmais.com.br,')

# GM - CHEVROLET
SANAUTO_emails = ('marlon.silva@carmais.com.br,')

# NISSAN
NISSAN_MATRIZ_emails = ('eduardo.weimar@carmais.com.br,')
NISSAN_FILIAL_emails = ('eduardo.weimar@carmais.com.br,')

# RENAULT
JANGADA_VEICULOS_emails = ('eriveltom.rocha@carmais.com.br,')

# SCANIA
CONTERRANEA_MATRIZ_emails =  ('janilson.almeida@carmais.com.br,')
CONTERRANEA_MACAIBA_emails = ('janilson.almeida@carmais.com.br,')
CONTERRANEA_MOSSORO_emails = ('janilson.almeida@carmais.com.br,')

# JANGADA AUTOMOTIVE
JANGADA_AUTOMOTIVE_emails = ('eduardo.weimar@carmais.com.br,')

NATIVA_emails = ('alan.estevao@carmais.com.br,')

TESTE_emails =  ('victor.queiroz@carmais.com.br')


# - Dicionário (Associando cada empresa a respectiva tupla com os emails)
dict_envio = {
        'CDA': CDA_emails,
        'VOUGA': VOUGA_emails,
        'NOVALUZ BS': NOVALUZ_BS_emails,
        'NOVALUZ WS': NOVALUZ_WS_emails,
        'NOVALUZ SD': NOVALUZ_SD_emails,
        'NOVALUZ SUL': NOVALUZ_SUL_emails,
        'NOSSAMOTO BATURITE': NOSSAMOTO_BATURITE_emails,
        'NOSSAMOTO SIQUEIRA': NOSSAMOTO_SIQUEIRA_emails,
        'NOSSAMOTO CAUCAIA': NOSSAMOTO_CAUCAIA_emails, 
        'NOSSAMOTO MATRIZ': NOSSAMOTO_MATRIZ_emails,
        'SANAUTO': SANAUTO_emails,
        'NISSAN MATRIZ': NISSAN_MATRIZ_emails,
        'NISSAN FILIAL': NISSAN_FILIAL_emails,
        'JANGADA VEICULOS': JANGADA_VEICULOS_emails,
        'CONTERRANEA MATRIZ': CONTERRANEA_MATRIZ_emails,
        'CONTERRANEA MACAIBA': CONTERRANEA_MACAIBA_emails,
        'CONTERRANEA MOSSORO': CONTERRANEA_MOSSORO_emails,
        'JANGADA AUTOMOTIVE': JANGADA_AUTOMOTIVE_emails,
        'NATIVA': NATIVA_emails,
        'teste': TESTE_emails,
}
       


def envia_email(name):

    ### Obtendo a data atual em formato de string
    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d_%m_%Y')
    data_em_texto_barr = data_atual.strftime('%d/%m/%Y')
    # inicio_mes = data_em_texto_barr[2:] 

    # Gerando a tabela em HTML apartir do arquivo XLSX
    
    df2 = pd.read_excel(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx')
    
    # print(df2.head().style.set_table_styles([dict(selector='th', props=[('text-align', 'left')]),
    #                                 dict(selector='td', props=[('text-align', 'left')])]))
    tabela = df2.to_html(justify = 'center')

    ### Enviar o pdf por email

    # Lista de Envio
    # receiver = 'victor.queiroz@carmais.com.br'
    receiver = dict_envio[name] + GRUPO_FIXO

    # dict_envio[name]
    # Configure o objeto MIME
    message = MIMEMultipart('related')
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = 'RELATÓRIO DE PIX NÃO ENVIADOS POR FORMULÁRIO - ' + name + ' ' + data_em_texto

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    message.attach(msgAlternative) # Mensagem Alternativa está sendo jogada no message (root-principal)

    msgText = MIMEText('RELATÓRIO DE PIX NÃO ENVIADOS POR FORMULÁRIO - ' + name)
    msgAlternative.attach(msgText) # msgText é jogada na mensagem alternativa

    msgText = MIMEText(
        f'<b>RELATÓRIO DE PIX NÃO ENVIADOS POR FORMULÁRIO - {name} {data_em_texto_barr}</b><br><p>\
            Estes recebimentos de PIX, aparecem no relatório de Saldo de Créditos não identificados e não foi utilizado\
                o formulário de pix para enviar os dados ao financeiro.{tabela}</br>', 'html')
    msgAlternative.attach(msgText)
    
    #use gmail with port
    session = smtplib.SMTP('smtp.gmail.com', 587)

    #enable security
    session.starttls()

    ######### ANTIGA FORMA DE ENVIAR EMAIL - APPS MENOS SEGURO ####################
    # login with mail_id and password
    session.login(sender, password)
    
    text = message.as_string()
    session.sendmail(sender, receiver.split(","), text)
    session.quit()

    # Excluindo os arquivos PNG
    
    os.remove(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xls')
    os.remove(f'C:\\RPA\\arquivos\\pix_nao_enviado_{name}.xlsx')

    logging.info('Email Enviado e arquivo'+ " " + str(name) +" " + 'excluído do diretório.')
    print('Email Enviado e arquivo'+ " " + str(name) +" " + 'excluído do diretório.')
    