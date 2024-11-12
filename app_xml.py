import xmltodict
import os
import json
import pandas as pd  # Para salvar os dados em Excel
from datetime import datetime

# Função para gerar a data padrão (30/mês atual/ano atual)
def data_padrao():
    hoje = datetime.today()
    return f"30/{hoje.month:02d}/{hoje.year}"

# Função para pegar as informações e armazenar em uma estrutura para exportação
def pegar_infos(name_files):
    dados = []  # Lista para armazenar os dados que serão exportados

    with open(f'nfe/{name_files}', "rb") as arquivo_xml:
        try:
            # Parse do XML
            dicionario_arquivo = xmltodict.parse(arquivo_xml)
            # print(json.dumps(dicionario_arquivo,indent=4))

            # Acessando as informações necessárias
            info_nfe = dicionario_arquivo["NFeLog"]["procNFe"]["NFe"]["infNFe"]
            destinatario = info_nfe["dest"]["xNome"]
            numero_nota = info_nfe["ide"]["nNF"]

            # Inicializando a estrutura de dados para o arquivo
            dados_arquivo = {
                "Arquivo": name_files,
                "Destinatário": destinatario,
                "CFOPs": [],
                "Número de Parcelas": 1,  # Inicializa com 1 por padrão
                "Data de Vencimento": None,
                "Numero da Nota": numero_nota
            }

            # Verifica se "det" é uma lista ou um único dicionário
            det = info_nfe.get("det")
            cfop_set = set()  # Usando set para garantir CFOPs únicos

            if det:
                if isinstance(det, list):  # Se "det" for uma lista
                    for item in det:
                        cfop_set.add(item["prod"]["CFOP"])  # Adiciona o CFOP no set
                elif isinstance(det, dict):  # Se "det" for um único dicionário
                    cfop_set.add(det["prod"]["CFOP"])  # Adiciona o CFOP no set

            # Armazenando os CFOPs únicos
            dados_arquivo["CFOPs"] = list(cfop_set)

            # Verificação da chave "cobr" antes de acessar "dup"
            cobr = info_nfe.get("cobr", None)  # Garantir que cobr seja um dicionário ou None
            dup = None  # Inicializando 'dup' com None

            if cobr:  # Verifica se 'cobr' existe (não é None)
                dup = cobr.get("dup", None)  # 'dup' pode ser uma lista ou um dicionário, se não existir, dup será None

            # Verificando se 'dup' existe ou não
            if dup:
                if isinstance(dup, list):  # Se 'dup' for uma lista (várias parcelas)
                    datas_vencimento = []  # Lista para armazenar todas as datas de vencimento
                    for parcela in dup:  # Iterando sobre as parcelas
                        # Acessando as informações da parcela diretamente, já que parcela é um dicionário
                        data_vcto = parcela.get("dVenc", data_padrao())  # Se não houver 'dVenc', usa a data padrão
                        datas_vencimento.append(data_vcto)  # Adicionando a data à lista

                    # Se houver pelo menos uma data de vencimento, armazene-a
                    if datas_vencimento:
                        dados_arquivo["Datas de Vencimento"] = datas_vencimento

                    # Atualizando o número de parcelas com base no tamanho da lista 'dup'
                    dados_arquivo["Número de Parcelas"] = len(dup)  # O número de parcelas é igual ao número de itens em 'dup'

                elif isinstance(dup, dict):  # Se 'dup' for um dicionário único (uma única parcela)
                    # Acessando a data de vencimento ou usando a data padrão
                    data_vcto = dup.get("dVenc", data_padrao())  # Se não houver 'dVenc', usa a data padrão
                    dados_arquivo["Data de Vencimento"] = data_vcto
                    # Se for uma parcela única, então o número de parcelas é 1
                    dados_arquivo["Número de Parcelas"] = 1
            else:
                # Caso não haja 'dup', usa a data padrão e assume que há 1 parcela
                dados_arquivo["Data de Vencimento"] = data_padrao()
                dados_arquivo["Número de Parcelas"] = 1

            # Adiciona os dados do arquivo na lista de dados
            dados.append(dados_arquivo)

        except Exception as msg:
            # Em caso de erro, coleta os detalhes para depuração
            print(f'Erro no arquivo {name_files}: {msg}')
            print("Conteúdo do XML com erro:")
            print(json.dumps(dicionario_arquivo, indent=4))
            return None  # Retorna None caso haja erro

    return dados  # Retorna a lista com os dados coletados


# Lista os arquivos dentro da pasta "nfe"
lista_files = os.listdir("nfe")
todas_informacoes = []

for files in lista_files:
    dados_arquivo = pegar_infos(files)
    if dados_arquivo:
        todas_informacoes.extend(dados_arquivo)

# Salva os dados coletados em um arquivo Excel
if todas_informacoes:
    df = pd.DataFrame(todas_informacoes)
    df.to_excel("informacoes_nfe.xlsx", index=False)  # Salva em Excel
    print("Dados exportados para 'informacoes_nfe.xlsx'.")
else:
    print("Nenhum dado foi coletado.")
