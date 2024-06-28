from Infra.Repository.ApiFaturamentoRepository import apiFaturamento, get_processos_por_remessa

def obterListaFaturasPorRemessa(lista_remessa, tipo_negociacao, token, convenio_id):
    faturas_dict_list = []
    for remessa in lista_remessa:
        lista = get_processos_por_remessa(tipo_negociacao, remessa, token, convenio_id)
        faturas_dict_list.extend(lista)
        
    listaFat = [[item['processoId'], item['remessaId'], item['apelido'],item['usuarioLanca'], item['protocoloItem'], item['envioGuiaPortal'], '', '', f'\\\\10.0.0.239\\guiasscaneadas\\2024\GEAP\\{item["processoId"]}.pdf']  for item in faturas_dict_list]

    return listaFat

def obterListaFaturas(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)

    if convenioId == 381:
        lista.extend(apiFaturamento(tipoNegociacao, 457, statusProcesso, dataInicio, dataFim,token))

    elif convenioId == 225:
        lista.extend(apiFaturamento(tipoNegociacao, 225, 405, dataInicio, dataFim,token))

    listaFat = [[item['processoId'], item['remessaId'], item['apelido'],item['usuarioLanca'], item['protocoloItem'], item['envioGuiaPortal'], '', '']  for item in lista]

    return listaFat

def get_faturas_data(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)

    if convenioId == 381:
        lista.extend(apiFaturamento(tipoNegociacao, 457, statusProcesso, dataInicio, dataFim,token))

    listaFat = [[item['processoId'], str(item['protocoloItem']), item['remessaId'],item['quantidadeGuias'], item['valorTotal'], '']  for item in lista if item['protocoloItem'] != None]

    return listaFat