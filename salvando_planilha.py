import os
import time
import pyautogui

# FILTRANDO O USUÁRIO NO SISTEMA
def filtrar_usuario(usuario):
    pyautogui.moveTo(555, 400, 2)
    pyautogui.click()
    pyautogui.write(usuario)
    pyautogui.press("enter")

# FILTRANDO AS SDES
def filtrar_sdes(coordenadas):
    x, y, duracao = coordenadas
    pyautogui.moveTo(1500, 345, 2)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(x, y, duracao)
    pyautogui.press("enter")
    time.sleep(2)

# FILTRANDO A DATA DE PESQUISA
def filtrar_data():
    pyautogui.moveTo(380, 345, 2)
    pyautogui.doubleClick()
    pyautogui.write("01062024")
    time.sleep(5)
    pyautogui.press("enter")
    time.sleep(30)

# CASO HAJA ALGUMA PLANILHA NA PASTA, DELETA
def excluindo_planilhas_desatualizadas():
    if os.path.exists("C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeVermelha.xlsx"):
        os.remove("C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeVermelha.xlsx")

    if os.path.exists("C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeAmarela.xlsx"):
        os.remove("C:\\Users\\joaov\\Desktop\\PROGRAMA\\sdeAmarela.xlsx")

# ABRINDO PLANILHA
def abrir_planilha():
    pyautogui.moveTo(500, 885, 2)
    pyautogui.click()
    time.sleep(5)
    pyautogui.keyDown("alt")
    pyautogui.keyDown("shift")
    pyautogui.press("tab")
    pyautogui.keyUp("shift")
    pyautogui.keyUp("alt")
    time.sleep(2)

# SALVAR PLANILHA
def salvar_planilha(nome_planilha):
    pyautogui.keyDown("alt")
    pyautogui.press("F4")
    pyautogui.keyUp("alt")
    pyautogui.press("enter")
    pyautogui.write(nome_planilha)
    pyautogui.moveTo(600, 60, 2)
    pyautogui.click()
    time.sleep(1)
    pyautogui.write("C:\\Users\\joaov\\Desktop\\PROGRAMA")
    pyautogui.press("enter")
    pyautogui.moveTo(750, 590, 2)
    pyautogui.click()


##############################################################################

usuario = "joaov"
coordenadas_vermelhas = (1500, 405, 2)
coordenadas_amarelas = (1500, 430, 2)

filtrar_usuario(usuario)

filtrar_sdes(coordenadas_vermelhas)
filtrar_data()
excluindo_planilhas_desatualizadas()
abrir_planilha()
salvar_planilha("planilha_sdes_vermelhas")

filtrar_sdes(coordenadas_amarelas)
abrir_planilha()
salvar_planilha("planilha_sdes_amarelas")