import pyautogui
from time import sleep
import ctypes
import sys
import openpyxl
import os
from datetime import datetime


wb_nfe = openpyxl.load_workbook('informacoes_nfe.xlsx')
sheet_nfe = wb_nfe['Sheet1']

# no cmd digitar
#python
#from mouseinfo import mouseInfo
#mouseInfo()
#da enter



def limpa_campo():
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

def preencher_parcelas():
        pyautogui.click(754,177,duration=2)
        for l in range(4):
                pyautogui.press('right')
        sleep(1)
        for k in range(4):
            pyautogui.press('backspace')
        pyautogui.write(str(numero_parcelas))
        sleep(1)
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

def preencher_parcelas_multiplas():
    # Preencher o número de parcelas
    pyautogui.press('TAB')  # Primeiro TAB para o campo de número de parcelas
    for l in range(4):      # Move o cursor para a direita
        pyautogui.press('right')
    sleep(1)
    for k in range(4):      # Limpa o campo
        pyautogui.press('backspace')
    pyautogui.write(str(numero_parcelas))
    sleep(1)
    
    pyautogui.press('TAB')  # Move para o campo "dias" ou similar
    sleep(1)
    pyautogui.write('30')   # Preenche o valor de dias
    sleep(1)
    
    for _ in range(3):
        pyautogui.press('TAB')  # Move para o campo de vencimento
        sleep(1)
    
    # Preenche a primeira data de vencimento
    for _ in range(10):
        pyautogui.press('right')  # Move para a direita no campo de data
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')  # Limpa o campo
    pyautogui.write(datas_vencimento[0])  # Preenche com a primeira data
    sleep(1)
    
    # Agora preenche as próximas datas de vencimento (se houver mais parcelas)
    if len(datas_vencimento) > 1:  # Se houver mais de uma data de vencimento (parcelas múltiplas)
        for i in range(1, len(datas_vencimento)):  # Começa no índice 1 para a segunda parcela
            # Dá 2 TABs extras para ir para o campo da próxima parcela
            pyautogui.press('TAB')  
            sleep(1)
            pyautogui.press('TAB')
            sleep(1)

            # Limpa e preenche o campo da data de vencimento
            for _ in range(10):
                pyautogui.press('right')  # Move para a direita no campo de data
            sleep(1)
            for j in range(10):
                pyautogui.press('backspace')  # Limpa o campo
            pyautogui.write(datas_vencimento[i])  # Preenche com a data da próxima parcela
            sleep(1)

    pyautogui.press('ENTER')  # Confirma as alterações
    sleep(2)

    pyautogui.press('pagedown')  # Realiza um page down caso necessário para continuar
#COMECAR A FUNCAO ITERANDO SOBRE AS LINHAS DO EXCEL
for linha in sheet_nfe.iter_rows(min_row=67, max_row=67):

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
    pyautogui.press('ENTER')
    sleep(2)

    #passando parametros do excel
    cfop_da_nota = linha[2].value.strip().replace('[', '').replace(']', '').replace("'", '')
    print(cfop_da_nota)
    print(data_vcto)
    datas_vencimento_composta = linha[6].value
    

    if isinstance(cfop_da_nota, str) and ',' in cfop_da_nota: 
        # Caso o CFOP seja uma string com múltiplos valores separados por vírgula
        cfops = cfop_da_nota.split(',')
        for cfop in cfops:
            cfop = cfop.strip()
            #LOGICA PARA QUANDO FOR MAIS DE 1 CFOP

            pyautogui.press('ENTER')
            sleep(2)

            for tab in range(5):
                  pyautogui.press('TAB')

            pyautogui.press('ENTER')
            sleep(2)

    
            pyautogui.press('TAB')
            sleep(1)
            pyautogui.press('F2')
            sleep(2)
            pyautogui.hotkey('ctrl', 'p')   
            sleep(1)
            pyautogui.click(740,340,duration=2)
            sleep(1)    
            pyautogui.click(560,358,duration=2)
            sleep(1)
            for i in range(2):
                 pyautogui.press('TAB')
            sleep(1)
            if cfop =='6102':
                pyautogui.write(str(2102))
            elif cfop =='5102':
                pyautogui.write(str(1102))
            elif cfop =='6101':
                pyautogui.write(str(2101))
            elif cfop =='5101':
                pyautogui.write(str(1101))
            elif cfop =='5949':
                pyautogui.write(str(1949))
            elif cfop =='5106':
                pyautogui.write(str(1403))
            elif cfop == '5403':
                pyautogui.write(str(1403))
            elif cfop =='5401':
                pyautogui.write(str(1403))
            elif cfop =='5405':
                pyautogui.write(str(1403))
            elif cfop =='5104':
                pyautogui.write(str(1102))
            else:
                 pyautogui.write(str(1102))

            pyautogui.press('ENTER')
            sleep(1)
            pyautogui.press('F3')
            sleep(1)
            pyautogui.press('ENTER')
            sleep(3)

            if cfop in ['5403', '5401', '5405']:
                    for l in range (2):
                        pyautogui.press('TAB')
                        sleep(1)
                        pyautogui.write(str(980))
                        for r in range(6):
                             pyautogui.press('TAB')
                             sleep(1)
                             limpa_campo()
                             sleep(1)
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()  
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()
                        pyautogui.press('TAB')
                        sleep(1)
                        limpa_campo()

                        pyautogui.press('TAB')
                        pyautogui.press('TAB')
                        sleep(1)
                        pyautogui.press('ENTER')
            else:
                for _ in range(8):
                     pyautogui.press('TAB')  
                     sleep(1)
                     limpa_campo()
                     sleep(1)
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()  
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()
                pyautogui.press('TAB')
                sleep(1)
                limpa_campo()

                pyautogui.press('TAB')
                pyautogui.press('TAB')
                sleep(1)
                pyautogui.press('ENTER')

                    

    else:             
        # Se for um único valor
        cfops = [cfop_da_nota]
        pyautogui.press('ENTER')


        # Caso a data de vencimento seja vazia ou inválida, usa uma data padrão
    if data_vcto is None or data_vcto == '':
        # Se a data composta de vencimento for uma string, dividimos pelas vírgulas
        datas_vencimento_composta = linha[6].value  # A data composta de vencimento
        
        if isinstance(datas_vencimento_composta, str):
            # Limpa a string para remover os colchetes e as aspas simples
            datas_vencimento = [data.strip("[]' ") for data in datas_vencimento_composta.split(',')]
            
            # Tenta converter a primeira data
            try:
                # A primeira data de vencimento
                data_formatada = datetime.strptime(datas_vencimento[0], '%Y-%m-%d').strftime('%d/%m/%Y')
            except ValueError:
                print(f"Erro ao tentar formatar a data: {datas_vencimento[0]}. Usando data padrão.")
                data_formatada = '01/01/1900'  # Data padrão ou outro valor
        else:
            # Se a data composta não for uma string, define a data padrão
            print("Data composta de vencimento não encontrada ou está mal formatada. Usando data padrão.")
            data_formatada = '01/01/1900'  # Data padrão ou outro valor
    else:
        # Se data_vcto não for None, usamos o valor normal
        if isinstance(data_vcto, datetime):
            # Se data_vcto já for um objeto datetime, apenas formate
            data_formatada = data_vcto.strftime('%d/%m/%Y')
        elif isinstance(data_vcto, str):
            try:
                # Tentando converter para o formato %Y-%m-%d
                data_formatada = datetime.strptime(data_vcto, '%Y-%m-%d').strftime('%d/%m/%Y')
            except ValueError:
                # Caso o valor seja inválido, define uma data padrão
                print(f"Erro ao tentar formatar a data: {data_vcto}. Usando data padrão.")
                data_formatada = '01/01/1900'  # Data padrão
        else:
            # Caso data_vcto seja outro tipo inesperado
            print(f"Tipo inesperado para data_vcto: {data_vcto}. Usando data padrão.")
            data_formatada = '01/01/1900'  # Valor padrão



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
    elif cfop_da_nota =='5102':
        pyautogui.write(str(1102))
    elif cfop_da_nota =='6101':
        pyautogui.write(str(2101))
    elif cfop_da_nota =='5101':
        pyautogui.write(str(1101))
    elif cfop_da_nota =='5949':
        pyautogui.write(str(1949))
    elif cfop_da_nota =='5106':
          pyautogui.write(str(1102))
    elif cfop_da_nota == '5403':
          pyautogui.write(str(1403))
    elif cfop_da_nota =='5401':
          pyautogui.write(str(1403))
    elif cfop_da_nota =='5405':
          pyautogui.write(str(1403))
    elif cfop_da_nota =='5104':
            pyautogui.write(str(1102))
    else:
        pyautogui.write(str(1102))

    pyautogui.press('ENTER')
    sleep(1)
    pyautogui.press('F3')
    sleep(1)
        
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
    if isinstance(datas_vencimento_composta, str) and ',' in datas_vencimento_composta:
        datas_vencimento = [data.strip() for data in datas_vencimento_composta.split(',')]  # Para várias datas
        preencher_parcelas_multiplas()
    else:
        datas_vencimento = [datas_vencimento_composta]
        preencher_parcelas()  # Para o caso de apenas uma parcela