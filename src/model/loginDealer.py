import pyautogui as p
import os
from .clickApi import click2 as c

def login_dealer_contas_receber():
    # Buscar as credenciais de acesso ao Dealernet

    credenciais = open('C:\\RPA\\credenciais\\login_dealer.txt','r')
    chaves = credenciais.readlines()
    credenciais.close()

    login = chaves[0][:-1]
    password = chaves[1]

    p.PAUSE=1
    os.startfile("C:\\Conces\\scr\\scr.exe")
    p.sleep(3)


    modulo_ja_aberto = p.locateCenterOnScreen('C:/RPA/arquivos/images/modulo_ja_aberto.png', confidence = 0.95)
    if modulo_ja_aberto != None:
        p.click(290,325) # Clique Forçado em Usar o módulo já aberto
        
    p.sleep(2)

    usuario = p.locateCenterOnScreen('C:/RPA/arquivos/images/usuario.png', confidence = 0.95)
    if usuario != None:
        
        c(usuario.x+70 , usuario.y )
        #for i in range(10):
        p.dragTo(usuario.x+165, usuario.y, duration=0.25)
        p.press('delete')
        p.write(login)




    senha = p.locateCenterOnScreen('C:/RPA/arquivos/images/senha.png', confidence = 0.95)
    if senha != None:
        c(senha.x+80 , senha.y )
        p.write(password)


    ok = p.locateCenterOnScreen('C:/RPA/arquivos/images/ok.png', confidence = 0.95)
    if ok != None:
        c(ok.x, ok.y )

    p.sleep(2)

    bd_sem_backup = p.locateCenterOnScreen('C:/RPA/arquivos/images/bd_sem_backup.png', confidence = 0.95)
    if bd_sem_backup != None:
        ok = p.locateCenterOnScreen('C:/RPA/arquivos/images/ok3.png', confidence = 0.95)
        if ok != None:
            c(ok.x, ok.y )

    p.sleep(1)

    cancelar = p.locateCenterOnScreen('C:/RPA/arquivos/images/cancelar.png', confidence = 0.95)
    while cancelar != None:
        p.sleep(0.5)
        cancelar = p.locateCenterOnScreen('C:/RPA/arquivos/images/cancelar.png', confidence = 0.95)
    p.sleep(1)
    