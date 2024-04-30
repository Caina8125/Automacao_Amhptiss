import json
import time
import requests

def auth():
    global token
    global proxies

    urlAuth = 'https://amhptiss.amhp.com.br/api/Auth'
    proxies = {
        "http": f"http://lucas.timoteo:81259018@10.0.0.230:3128/",
        "https": f"http://lucas.timoteo:81259018@10.0.0.230:3128/"
    }

    usuario_login = {
        "Usuario": "lucas.timoteo",
        "Senha" : "Caina8125"
    }

    post = requests.post( urlAuth, usuario_login, proxies=proxies)
    time.sleep(2)
    content = json.loads(post.content)
    time.sleep(1)
    token = content['AccessToken']
    time.sleep(1)
    print('Token =>',token)
    print('')
    print('')



def apiFaturamento(login,senha, tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    urlApi = f'http://10.0.0.142:9000/api/faturamento/{tipoNegociacao}/processo-remessa-sem-recurso?convenioId={convenioId}&statusProcesso={statusProcesso}&dataInicio={dataInicio}&dataFim={dataFim}'

    proxies = {
        "http": f"http://lucas.timoteo:81259018@10.0.0.230:3128",
        "https": f"http://lucas.timoteo:81259018@10.0.0.230:3128"
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'api-key': '8EC08ED9-94DC-413E-9A29-BC211A9BEA30'
    }

    data = {
        'tiponegociacao': tipoNegociacao,
        'convenioId': convenioId,
        'statusProcesso': statusProcesso,
        'dataInicio': dataInicio,
        'dataFim': dataFim
    }

    response = requests.get(url=urlApi, headers=headers, data=data, verify=False)
    content = json.loads(response.content)
    print(content)

tokenTeste = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI0YmYzOWM4MS03OWIxLTZjNWQtZTA1My0wMTAwMDA3ZjIyMTUiLCJ1bmlxdWVfbmFtZSI6Imx1Y2FzLnRpbW90ZW8iLCJqdGkiOiI5ZTMyMThhMC1jODUyLTQwNjMtYTgyNi1kMzAxZjIyODY5NDYiLCJuYmYiOjE3MTQ0OTk4OTksImlhdCI6MTcxNDQ5OTg5OSwiaWQiOiI0MzIxOTQ5NCIsIm5vbWUiOiJMdWNhcyBDYWluYSBGZXJyZWlyYSBUaW1vdGVvIiwicGVzc29hX2Zpc2ljYSI6InRydWUiLCJyb2xlIjpbIkZ1bmNpb25hcmlvQU1IUCIsIkFkbWluaXN0cmFkb3IiLCJGYXR1cmFtZW50b0FkbWluaXN0cmF0aXZvIiwiSW5mb3JtYXRpdm9zIiwiTmVnb2NpYWNhbyIsIk5lZ29jaWFjYW9EaXJldG8iLCJPcGVyYWNpb25hbCIsIlNHQ0ZBVFVSQU1FTlRPIiwiU0dDSU5GT1JNQVRJVk8iLCJTR0NTSVRFIl0sImV4cCI6MTcxNDUzNTg5OSwiaXNzIjoiQU1IUCIsImF1ZCI6Imh0dHBzOi8vYW1ocHRpc3MuYW1ocC5jb20uYnIifQ.0tY9PUuFqafPAn2rssLR8Uy4mWcvvruqpvqWGhK1TXM"

apiFaturamento("","","normal",27,400,"2024/04/01","2024/04/30",tokenTeste)





