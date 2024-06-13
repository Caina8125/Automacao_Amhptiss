from Infra.Repository.ApiFaturamentoRepository import apiFaturamento

def obterListaFaturas(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)
    listaFat = [[iten['processoId'], iten['remessaId'], iten['apelido'],iten['usuarioLanca'], iten['protocoloItem'], iten['envioGuiaPortal'], '', '']  for iten in lista]

    return listaFat