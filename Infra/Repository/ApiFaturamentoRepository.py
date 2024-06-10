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

def api_fatura(tipo_negociacao, convenio_id, processo_id, token):
    url_api = f'http://10.0.0.142:9000/api/faturamento/{tipo_negociacao}/obter-processo-completo/{convenio_id}/{processo_id}'

    data = {
        'tiponegociacao': tipo_negociacao,
        'convenioId': convenio_id,
        'processoId': processo_id,
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'api-key': '8EC08ED9-94DC-413E-9A29-BC211A9BEA30'
    }

    response = requests.get(url=url_api, headers=headers, data=data, verify=False)
    content = json.loads(response.content)
    return content

def post_status_envio_operadora(self, tipo_negociacao, remessa_id, envio_operadora, token):
    url_api = f'http://10.0.0.142:9000/api/faturamento/{tipo_negociacao}/envio-operadora?remessaId={remessa_id}&envioOperadora={envio_operadora}'

    data = {
        'tiponegociacao': tipo_negociacao,
        'remessaId': remessa_id,
        'envioOperadora': envio_operadora,
    }

    headers = {
        'Authorization': f'Bearer {token}',
        'api-key': '8EC08ED9-94DC-413E-9A29-BC211A9BEA30'
    }

    response = requests.get(url=url_api, headers=headers, data=data, verify=False)
    content = json.loads(response.content)
    return content
