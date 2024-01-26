
import os
import shutil

usuario = 'alex'
senha = 'desafiosrpa'
link = 'https://jornadarpa.com.br/demandas/login.html'

def mover_arquivos(pasta_origem, pasta_destino):
        # Garante que a pasta de destino exista, se não existir, ela será criada
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        # Lista todos os arquivos no diretório de origem
        files = os.listdir(pasta_origem)

        # Itera sobre os arquivos e move aqueles com extensão '.png' para o diretório de destino
        for file in files:
            if file.endswith('.png'):
                source_path = os.path.join(pasta_origem, file)
                destination_path = os.path.join(pasta_destino, file)
                shutil.move(source_path, destination_path)
                print(f"Arquivo '{file}' movido para '{pasta_destino}'.")