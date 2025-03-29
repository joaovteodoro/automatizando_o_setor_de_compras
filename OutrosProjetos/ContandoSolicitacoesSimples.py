import pyautogui
import time

#FILTRANDO AS SDES DO USUÁRIO NO SISTEMA
pyautogui.moveTo (555, 400, 2)
pyautogui.click()
pyautogui.write('joaov')
pyautogui.press('enter')

#FILTRANDO AS SDES VERMELHAS
pyautogui.moveTo (1500, 345, 2)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo (1500, 405, 2)
pyautogui.press('enter')
time.sleep(2)

#FILTRANDO A DATA DE PESQUISA
pyautogui.moveTo (380, 345, 2)
pyautogui.doubleClick()
pyautogui.write('01062024')
pyautogui.press('enter')
time.sleep(10)

#CONTAR AS SDES VERMELHAS
pyautogui.moveTo (410, 460, 1)
ContagemVermelha = 0
TemSolicitacao = False
cor = pyautogui.pixel(465, 457)
if  cor == (255,0,0): #verigica se a bolinha PC é vermelha
    pyautogui.doubleClick()
    time.sleep(5)
    TemSolicitacao = True

while TemSolicitacao:

    ContagemVermelha +=1 
    im1 = pyautogui.screenshot(region=(340,230, 800,800))
    pyautogui.press('down')
    time.sleep(7)
    im2 = pyautogui.screenshot(region=(340,230, 800,800))

    if im1 == im2:
        break
         
print(ContagemVermelha)


#FILTRANDO AS SDES AMARELAS
pyautogui.moveTo (1500, 345, 2)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo (1500, 430, 2)
pyautogui.press('enter')
time.sleep(2)


#CONTAR AS SDES AMARELAS
pyautogui.moveTo (410, 460, 2)
ContagemAmarela = 0
TemSolicitacao = False
cor = pyautogui.pixel(465, 457)
if  cor == (255,255,0): #verifica se a bolinha PC é amarela
    pyautogui.doubleClick()
    time.sleep(5)
    TemSolicitacao = True
    
while TemSolicitacao:

    ContagemAmarela +=1 
    im1 = pyautogui.screenshot(region=(340,230, 800,800))
    pyautogui.press('down')
    time.sleep(7)
    im2 = pyautogui.screenshot(region=(340,230, 800,800))

    if im1 == im2:
        break
         
print(ContagemAmarela)

print(f'Temos {ContagemVermelha} SDEs vermelhas e {ContagemAmarela} SDEs amarelas, totalizando {ContagemVermelha+ContagemAmarela}')