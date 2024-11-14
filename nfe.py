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

def limpa_parcelas():
    for _ in range(3):
            pyautogui.press('right')
    sleep(1)
    for j in range(3):
        pyautogui.press('backspace')
    for l in range(3):
            pyautogui.press('right')
    sleep(1)
    for k in range(3):
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
pyautogui.click(835,438,duration=2)
sleep(5)
pyautogui.press('pagedown')

def preencher_parcelas():
    data_formatada = datas_vencimento[0]  # Assume que a data já está formatada corretamente no Excel
    pyautogui.press('TAB')
    sleep(1)
    limpa_parcelas()
    sleep(1)
    pyautogui.write(str(numero_parcelas))
    sleep(1)
    pyautogui.press('TAB')
    sleep(1)
    pyautogui.write('30')
    sleep(1)
    for _ in range(3):
        pyautogui.press('tab')
    limpa_parcelas()
    sleep(1)
    pyautogui.write(data_vcto)  # Escreve a data diretamente
    print(data_formatada)
    sleep(1)
    pyautogui.press('ENTER')
    sleep(2)
    

def preencher_parcelas_multiplas():
    # Preencher o número de parcelas
    pyautogui.press('TAB')  # Primeiro TAB para o campo de número de parcelas
    limpa_parcelas()
    pyautogui.write(str(numero_parcelas))
    sleep(1)
    
    pyautogui.press('TAB')  # Move para o campo "dias" ou similar
    sleep(1)
    pyautogui.write('30')   # Preenche o valor de dias
    sleep(2)
    
    for _ in range(3):
        pyautogui.press('TAB')  # Move para o campo de vencimento
        sleep(1)
    
    # Preenche a primeira data de vencimento
    data_formatada = datas_vencimento[0]  # Assume que a data já está formatada corretamente no Excel
    for _ in range(10):
        pyautogui.press('right')  # Move para a direita no campo de data
    sleep(1)
    for j in range(10):
        pyautogui.press('backspace')  # Limpa o campo
    pyautogui.write(data_formatada)  # Preenche com a primeira data formatada
    sleep(2)
    
    # Agora preenche as próximas datas de vencimento (se houver mais parcelas)
    if len(datas_vencimento) > 1:  # Se houver mais de uma data de vencimento (parcelas múltiplas)
        for i in range(1, len(datas_vencimento)):  # Começa no índice 1 para a segunda parcela
            for _ in range(2):
                pyautogui.press('TAB')

            # Limpa e preenche o campo da data de vencimento
            for _ in range(10):
                pyautogui.press('right')  # Move para a direita no campo de data
            sleep(1)
            for j in range(10):
                pyautogui.press('backspace')  # Limpa o campo
            # Preenche com a próxima data
            data_formatada = datas_vencimento[i]  # A data já vem formatada do Excel
            pyautogui.write(data_formatada)  # Preenche com a data formatada
            print(data_formatada)  # Para verificar a data formatada
            sleep(1)

    pyautogui.press('ENTER')  
    sleep(4)

#COMECAR A FUNCAO ITERANDO SOBRE AS LINHAS DO EXCEL
for linha in sheet_nfe.iter_rows(min_row=2, max_row=3):

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
    pyautogui.press('ENTER')
    sleep(2)

    

    #passando parametros do excel
    cfop_da_nota = linha[2].value.strip().replace('[', '').replace(']', '').replace("'", '')
    print("data do vcto simples:",data_vcto)
    datas_vencimento_composta = linha[6].value              
    try:
        button_nota_A_B_location = pyautogui.locateOnScreen('botao_a.png')
        button_nota_B_location = pyautogui.locateOnScreen('botao_b.png')
        button_nota_C_location = pyautogui.locateOnScreen('botao_c.png')
        button_nota_D_location = pyautogui.locateOnScreen('botao_d.png')
    except pyautogui.ImageNotFoundException:
        button_nota_A_B_location = None
        button_nota_B_location = None
        button_nota_C_location = None
        button_nota_D_location = None
    if button_nota_A_B_location and isinstance(cfop_da_nota, str) and ',' in cfop_da_nota: 
        # Caso o CFOP seja uma string com múltiplos valores separados por vírgula
        cfops = cfop_da_nota.split(',')
        for tab in range(5):
                  pyautogui.press('TAB')
        for cfop in cfops:
            cfop = cfop.strip()
            #LOGICA PARA QUANDO FOR MAIS DE 1 CFOP

            pyautogui.press('ENTER')
            sleep(4)


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
                        pyautogui.click(444,393,duration=2)
                        sleep(1)
                        pyautogui.write(str(980))
                        sleep(1)
                        for r in range(6):
                             pyautogui.press('TAB')
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
                        sleep(4)
                        pyautogui.click(702,438,duration=2)
                        sleep(8)
                        
                      
                        
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
                sleep(5)
                pyautogui.click(702,438,duration=2)
                sleep(5)
    
        sleep(10)
        if isinstance(datas_vencimento_composta, str) and ',' in datas_vencimento_composta:
            datas_vencimento = [data.strip() for data in datas_vencimento_composta.split(',')]  # Para várias datas
            pyautogui.press('ENTER')
            sleep(5)
            pyautogui.press('ESC')
            sleep(5)
            preencher_parcelas_multiplas()
        else:
            datas_vencimento = [datas_vencimento_composta]
            pyautogui.press('ENTER')
            sleep(5)
            pyautogui.press('ESC')
            sleep(5)
            preencher_parcelas()  # Para o caso de apenas uma parcela

    elif button_nota_A_B_location and isinstance(cfop_da_nota, str) and ',' not in cfop_da_nota:
        cfop = cfop_da_nota.strip()
        while True:
            sleep(2)
            for tab in range(5):
                  pyautogui.press('TAB')

            pyautogui.press('ENTER')
            sleep(4)


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
                        pyautogui.click(444,393,duration=2)
                        sleep(1)
                        pyautogui.write(str(980))
                        sleep(1)
                        for r in range(6):
                             pyautogui.press('TAB')
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
                        sleep(4)
                        pyautogui.click(702,438,duration=2)
                        sleep(8)
                        
                      
                        
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
                sleep(5)
                pyautogui.click(702,438,duration=2)
                sleep(5)
    
                sleep(10)
                if isinstance(datas_vencimento_composta, str) and ',' in datas_vencimento_composta:
                    datas_vencimento = [data.strip() for data in datas_vencimento_composta.split(',')]  # Para várias datas
                    pyautogui.press('ENTER')
                    sleep(5)
                    pyautogui.press('ESC')
                    sleep(5)
                    preencher_parcelas_multiplas()
                else:
                    datas_vencimento = [datas_vencimento_composta]
                    pyautogui.press('ENTER')
                    sleep(5)
                    pyautogui.press('ESC')
                    sleep(5)
                    preencher_parcelas()  # Para o caso de apenas uma parcela

                
                if not button_nota_A_B_location and button_nota_B_location and button_nota_C_location and button_nota_D_location:
                    break  

    else:             
        #se so tiver 1 cfop
        cfops = [cfop_da_nota]


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
                    data_formatada = [
                        datetime.strptime(data, '%Y-%m-%d').strftime('%d/%m/%Y') for data in datas_vencimento
                    ]
                except ValueError:
                    data_formatada = '01/01/1900'  # Data padrão ou outro valor
            else:
                # Se a data composta não for uma string, define a data padrão
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
                    data_formatada = '01/01/1900'  # Data padrão
            else:
                # Caso data_vcto seja outro tipo inesperado
                data_formatada = '01/01/1900'  # Valor padrão



        #CLICAR EM CODIGO FISCAL DANDO TAB
        for _ in range(3):
            pyautogui.press('TAB')
        sleep(2)
        pyautogui.press('F2')
        sleep(1)
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
        sleep(1)


        #CLICAR EM ALIQUOTA, LIMPAR
        #CLICAR EM BASE DE CALCULO, LIMPAR
        #CLICAR EM IMPOSTO, LIMPAR
        #CLICAR EM ISENTAS/NT, LIMPAR
        #CLICAR EM VALOR AUXILAR, LIMPAR
        #CLICAR EM BASE IMP. RETIDO, LIMPAR
        #CLICAR EM IMPOSTO RETIDO, LIMPAR
        for _ in range(10):
            pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)
               
        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        pyautogui.press('TAB')
        limpa_campo()
        sleep(1)

        #limpando os 2 campos depois de aliquota etc.. (podem nao ser clicaveis, pode bugar)
        pyautogui.click(1004,366,duration=1)
        limpa_campo()
        sleep(1)

        pyautogui.click(1000,418,duration=1)
        limpa_campo()
        sleep(1)


        pyautogui.click(982,661,duration=1) #clicando em salvar, pq os tab aq pode nao funcionar
        sleep(5)

        #clicando em "parametro contabil foi alterado etc"
        pyautogui.click(684,437,duration=1)
        sleep(3)

        pyautogui.press('ENTER')
        sleep(5)


        #PREENCHER NUMERO DE PARCELAS
        if isinstance(datas_vencimento_composta, str) and ',' in datas_vencimento_composta:
            datas_vencimento = [data.strip() for data in datas_vencimento_composta.split(',')]  # Para várias datas
            preencher_parcelas_multiplas()
        else:
            datas_vencimento = [datas_vencimento_composta]

            preencher_parcelas()  # Para o caso de apenas uma parcela

        pyautogui.press('pagedown')

        #FIM DO PROGRAMA

