import pyautogui
from time import sleep
import ctypes
import sys
import openpyxl
import os
from datetime import datetime

wb_nfe = openpyxl.load_workbook('informacoes_nfe.xlsx')
sheet_nfe = wb_nfe['Sheet1']

pyautogui.click(792,745,duration=2)
sleep(1)
#CLICAR EM MOVIMENTOS
pyautogui.click(171,37,duration=2)
sleep(1)
#CLICAR EM REGISTRO DE ENTRADAS
pyautogui.click(213,62,duration=1)
#CLICAR EM DOCUMENTO
pyautogui.click(418,63,duration=1)
sleep(3)
#TECLAR PAGE DOWN
pyautogui.click(835,438,duration=2)
sleep(5)
pyautogui.press('pagedown')

#COMECAR A FUNCAO ITERANDO SOBRE AS LINHAS DO EXCEL
for linha in sheet_nfe.iter_rows(min_row=3, max_row=3):

    numero_parcelas = linha[3].value
    data_vcto = linha[4].value
    numero_nota = linha[5].value
    #TECLAR F5
    pyautogui.press('F5')
    sleep(2)
    pyautogui.press('F2')
    sleep(1)
    pyautogui.hotkey('ctrl', 'p')
    sleep(1)
    pyautogui.write(str(numero_nota))
    #TECLAR ENTER
    pyautogui.press('ENTER')
    sleep(2)
    pyautogui.press('enter')
    sleep(4)
    try:
        button_nota_A_B_location = pyautogui.locateOnScreen('botao_a.png')
        print('achei')
    except pyautogui.ImageNotFoundException:
            button_nota_A_B_location = None
            print("Botão A NÃO encontrado!")
            sleep(2)
            pyautogui.click(760,435, duration=2)
   

    
    


