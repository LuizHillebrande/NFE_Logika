import pyautogui
from time import sleep
import ctypes
import sys
import openpyxl
import os
from datetime import datetime


wb_nfe = openpyxl.load_workbook('informacoes_nfe.xlsx')
sheet_nfe = wb_nfe['nfe']

# no cmd digitar
#python
#from mouseinfo import mouseInfo
#mouseInfo()
#da enter

#AUTOMATIZA NOTAS DE ENTRADA NO SISTEMA EXACTUS COM BASE EM PLANILHA EXCEL XML
#PASSO A PASSO
#ABRIR A EXACTUS
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
pyautogui.press('pagedown')


#COMECAR A FUNCAO ITERANDO SOBRE AS LINHAS DO EXCEL
for linha in sheet_nfe.iter_rows(min_row=2, max_row=4):

    #TECLAR F5
    pyautogui.press('F5')
    sleep(2)

    #TECLAR ENTER
    pyautogui.press('ENTER')
    sleep(2)

    #passando parametros do excel
    cfop_da_nota = linha[2].value.strip().replace('[', '').replace(']', '').replace("'", '')
    numero_parcelas = linha[3].value.strip()
    data_vcto = linha[4].value.strip()
    if isinstance(data_vcto, datetime):
        data_formatada = datetime.strptime(data_vcto, "%Y%m%d").strftime("%d/%m/%Y")
    else:
        data_formatada = datetime.strptime(str(data_vcto), '%Y-%m-%d').strftime('%d/%m/%Y')



    #CLICAR EM CODIGO FISCAL
    codigo_fiscal = pyautogui.click(401,306,duration=2)
    sleep(2)
    pyautogui.hotkey('ctrl', 'p')
    sleep(1)
    pyautogui.click(519,339,duration=1)
    sleep(1)
    pyautogui.click(532,362,duration=1)
    sleep(1)
    pyautogui.click(627,390,duration=1)
    sleep(1)
    if cfop_da_nota =='6102':
        pyautogui.write(str(2102))
        pyautogui.press('ENTER')
        sleep(1)
        pyautogui.press('F3')
        sleep(1)
    elif cfop_da_nota =='5102':
        pyautogui.write(str(1102))
        pyautogui.press('ENTER')
        sleep(1)
        pyautogui.press('F3')
        sleep(1)
    elif cfop_da_nota =='6101':
        pyautogui.write(str(2101))
    elif cfop_da_nota =='5101':
        pyautogui.write(str(1101))
    elif cfop_da_nota =='5949':
        pyautogui.write(str(1949))
    elif cfop_da_nota =='5106':
          cfop_da_nota =='5102'
    elif cfop_da_nota == '5403':
          cfop_da_nota =='1403'
    elif cfop_da_nota =='5401':
          cfop_da_nota =='1402'
    elif cfop_da_nota =='5405':
          cfop_da_nota =='1403'
        
    pyautogui.press('ENTER')


    #CLICAR EM ALIQUOTA, LIMPAR
    #CLICAR EM BASE DE CALCULO, LIMPAR
    #CLICAR EM IMPOSTO, LIMPAR
    #CLICAR EM ISENTAS/NT, LIMPAR
    #CLICAR EM VALOR AUXILAR, LIMPAR
    #CLICAR EM BASE IMP. RETIDO, LIMPAR
    #CLICAR EM IMPOSTO RETIDO, LIMPAR

    pyautogui.click(764,300)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,319)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,343)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,367)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,414)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,440)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    pyautogui.click(770,462)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    for l in range(10):
            pyautogui.press('right')
    sleep(1)
    for k in range(10):
        pyautogui.press('backspace')
    sleep(1)

    
    pyautogui.press('ENTER')
    sleep(8)
    pyautogui.press('ENTER')
    sleep(5)

    #PREENCHER NUMERO DE PARCELAS
    pyautogui.click(754,177,duration=2)
    for l in range(4):
            pyautogui.press('right')
    sleep(1)
    for k in range(4):
        pyautogui.press('backspace')
    pyautogui.write(numero_parcelas)
    sleep(1)
    
    #CLICAR EM DIA PADRAO DO VENCIMENTO

    pyautogui.click(998,176)
    sleep(1)
    pyautogui.write('30')

    pyautogui.click(362,273)
    sleep(1)
    for _ in range(10):
            pyautogui.press('right')
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')
    pyautogui.write(data_formatada)
    sleep(1)
    pyautogui.press('ENTER')
    sleep(2)

    pyautogui.press('pagedown')