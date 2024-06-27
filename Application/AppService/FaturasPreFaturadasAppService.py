from Infra.Repository.ApiFaturamentoRepository import apiFaturamento

def obterListaFaturas(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)

    if convenioId == 381:
        lista.extend(apiFaturamento(tipoNegociacao, 457, statusProcesso, dataInicio, dataFim,token))

    listaFat = [[item['processoId'], item['remessaId'], item['apelido'],item['usuarioLanca'], item['protocoloItem'], item['envioGuiaPortal'], '', '']  for item in lista]

    return listaFat

def get_faturas_data(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)

    if convenioId == 381:
        lista.extend(apiFaturamento(tipoNegociacao, 457, statusProcesso, dataInicio, dataFim,token))

    listaFat = [[item['processoId'], str(item['protocoloItem']), item['remessaId'],item['quantidadeGuias'], item['valorTotal'], '']  for item in lista if item['protocoloItem'] != None]

    return listaFat