from .clickApi import click2 as c
import pyautogui as p


def escolha_empresas(name):

    p.sleep(1)
    empresa = p.locateCenterOnScreen('C:/RPA/arquivos/images/empresa.png', confidence = 0.95)
    if empresa != None:
        
        c(empresa.x , empresa.y-8 )
    p.sleep(1)
    marcar = p.locateCenterOnScreen('C:/RPA/arquivos/images/marcar.png', confidence = 0.95)
    if marcar != None:
        
        c(marcar.x , marcar.y )
        p.sleep(1)
        c(marcar.x , marcar.y ) 

    print(name)
    p.sleep(1)
    desmarcar = p.locateCenterOnScreen('C:/RPA/arquivos/images/desmarcar.png', confidence = 0.95)
    if desmarcar != None:
        
        c(desmarcar.x , desmarcar.y )

    p.sleep(2)



    mudar_empresa = p.locateCenterOnScreen(f'C:/RPA/arquivos/images/mudar_empresa_sdni/{name}.png', confidence = 0.95)
    while mudar_empresa == None:
        p.sleep(1)
        seta2 = p.locateCenterOnScreen('C:/RPA/arquivos/images/seta2.png', confidence = 0.95)
        c(seta2.x , seta2.y )
        mudar_empresa = p.locateCenterOnScreen(f'C:/RPA/arquivos/images/mudar_empresa_sdni/{name}.png', confidence = 0.95)
    c(mudar_empresa.x , mudar_empresa.y)
    p.sleep(1)


    ok = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\images\ok2.png', confidence = 0.95)
    if ok != None:
        c(ok.x, ok.y )

    p.sleep(1)
