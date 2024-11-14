import xmltodict
import os
import json
import pandas as pd  # Para salvar os dados em Excel
from datetime import datetime

# Caminho do arquivo XML
arquivo_xml = "nfe/35241006853754000170550000016977401654601893.xml"

# Abre o arquivo XML e imprime seu conteúdo
with open(arquivo_xml, "rb") as arquivo:
    try:
        # Fazendo o parsing do XML para um dicionário
        dicionario_arquivo = xmltodict.parse(arquivo)
        
        # Imprime o conteúdo do dicionário de maneira formatada
        print("Conteúdo do XML (estruturado como dicionário):")
        print(json.dumps(dicionario_arquivo, indent=4))  # Formato bonito para análise
        
    except Exception as e:
        print(f"Erro ao ler o XML: {e}")
