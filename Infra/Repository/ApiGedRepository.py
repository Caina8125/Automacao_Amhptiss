import json
import time
import requests

def ObterFaturasRobô(pasta,nomeArquivo,token):

    urlPost = 'https://amhpged.amhp.com.br/api/AtendimentoArquivo/listar-arquivos/'

    headers = {
        'Authorization': f'bearer {token}'
    }

    data = {
        'idAtendimento ': 1
    }

    files = {
        'file': (nomeArquivo, open(pasta, 'rb'))
    }

    response = requests.post(urlPost, headers=headers, data=data, files=files, verify=False)
    time.sleep(2)
    files['file'][1].close()
    print("")
    print("GED =>", response)
    print("")
    print(f"O arquivo: {nomeArquivo} foi armazenado na S3")
    print("")