import xmltodict
import os
import json
import pandas as pd  # Para salvar os dados em Excel

# Função para pegar as informações e armazenar em uma estrutura para exportação
def pegar_infos(name_files):
    dados = []  # Lista para armazenar os dados que serão exportados

    with open(f'nfe/{name_files}', "rb") as arquivo_xml:
        try:
            # Parse do XML
            dicionario_arquivo = xmltodict.parse(arquivo_xml)
            #print(json.dumps(dicionario_arquivo,indent=4))

            
            # Acessando as informações necessárias
            info_nfe = dicionario_arquivo["NFeLog"]["procNFe"]["NFe"]["infNFe"]
            destinatario = info_nfe["dest"]["xNome"]
            numero_parcelas = info_nfe.get("cobr", {}).get("dup", {}).get("nDup", "1")
            
            # Inicializando a estrutura de dados para o arquivo
            dados_arquivo = {
                "Arquivo": name_files,
                "Destinatário": destinatario,
                "CFOPs": [],
                "Número de Parcelas": numero_parcelas,
                "Data de Vencimento": None
                
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
            cobr = info_nfe.get("cobr")
            if cobr:
                dup = cobr.get("dup")
                if dup:
                    data_vcto = dup.get("dVenc")
                    if data_vcto:
                        dados_arquivo["Data de Vencimento"] = data_vcto

            # Adiciona os dados do arquivo na lista de dados
            dados.append(dados_arquivo)

        except Exception as msg:
            # Em caso de erro, coleta os detalhes para depuração
            print(f'Erro na {msg}')
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
