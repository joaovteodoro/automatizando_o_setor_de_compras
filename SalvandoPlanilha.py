import pyautogui
import time
import os

#FILTRANDO AS SDES DO USUÁRIO NO SISTEMA:
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
time.sleep(30)

#CASO HAJA ALGUMA PLANILHA NA PASTA, DELETA
if os.path.exists('C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeVermelha.xlsx'):
    os.remove('C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeVermelha.xlsx')

if os.path.exists('C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeAmarela.xlsx'):
    os.remove('C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeAmarela.xlsx')  


#ABRINDO PLANILHA
pyautogui.moveTo (500, 885, 2)
pyautogui.click()
time.sleep(5)
pyautogui.keyDown('alt')
pyautogui.keyDown('shift')
pyautogui.press('tab')
pyautogui.keyUp('shift')
pyautogui.keyUp('alt')
time.sleep(2)

#SALVAR PLANILHA
pyautogui.keyDown('alt')
pyautogui.press('F4')
pyautogui.keyUp('alt')
pyautogui.press('enter')
pyautogui.write('sdeVermelha')
pyautogui.moveTo (600, 60, 2)
pyautogui.click()
time.sleep(1)
pyautogui.write('C:\\Users\\joaov\\Desktop\\PROGRAMA')
pyautogui.press('enter')
pyautogui.moveTo (750, 590, 2)
pyautogui.click()

#FILTRANDO AS SDES AMARELAS
time.sleep(2)
pyautogui.moveTo (1500, 345, 2)
pyautogui.click()
time.sleep(2)
pyautogui.moveTo (1500, 430, 2)
pyautogui.press('enter')
time.sleep(5)

#ABRINDO PLANILHA
pyautogui.moveTo (500, 885, 2)
pyautogui.click()
time.sleep(5)
pyautogui.keyDown('alt')
pyautogui.keyDown('shift')
pyautogui.press('tab')
pyautogui.keyUp('shift')
pyautogui.keyUp('alt')
time.sleep(2)

#SALVAR PLANILHA
pyautogui.keyDown('alt')
pyautogui.press('F4')
pyautogui.keyUp('alt')
pyautogui.press('enter')
pyautogui.write('sdeAmarela')
pyautogui.moveTo (600, 60, 2)
pyautogui.click()
time.sleep(1)
pyautogui.write('C:\\Users\\joaov\\Desktop\\PROGRAMA')
pyautogui.press('enter')
pyautogui.moveTo (750, 590, 2)
pyautogui.click()

##############################################################################

