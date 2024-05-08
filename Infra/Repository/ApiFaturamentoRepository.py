import json
import time
import requests

def apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    urlApi = f'http://10.0.0.142:9000/api/faturamento/{tipoNegociacao}/processo-remessa-sem-recurso?convenioId={convenioId}&statusProcesso={statusProcesso}&dataInicio={dataInicio}&dataFim={dataFim}'

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
    return content