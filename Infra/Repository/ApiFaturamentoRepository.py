import json
import time
import requests

def apiFaturamento(login,senha, tipoNegociacao):
    urlApi = f'http://10.0.0.142:9000/api/faturamento/{tipoNegociacao}'

    proxies = {
        "http": f"http:// + {login}:{senha}@10.0.0.230:3128/",
        "https": f"http://{login}:{senha}@10.0.0.230:3128/"
    }





