import os

# Caminho da pasta onde estão os arquivos XML
# Caminho da pasta onde estão os arquivos XML
pasta_nfe = r'\\Hdlogika\k\Backup_19_08_2021\Unidade_D\K\LEATICIA\Luiz Fernando\NFE_Automation\NFE_Logika\nfe'


# Listar todos os arquivos na pasta
arquivos = os.listdir(pasta_nfe)

# Iterar sobre os arquivos e excluir os arquivos .xml
for arquivo in arquivos:
    if arquivo.endswith('.xml'):
        # Caminho completo do arquivo
        caminho_arquivo = os.path.join(pasta_nfe, arquivo)
        
        try:
            os.remove(caminho_arquivo)
            print(f'Arquivo {arquivo} excluído com sucesso!')
        except Exception as e:
            print(f'Erro ao excluir {arquivo}: {e}')
