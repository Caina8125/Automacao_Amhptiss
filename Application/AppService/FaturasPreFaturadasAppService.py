from Infra.Repository.ApiFaturamentoRepository import apiFaturamento
import pandas as pd

def obterListaFaturas(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token):
    lista = apiFaturamento(tipoNegociacao, convenioId, statusProcesso, dataInicio, dataFim,token)
    listaFat = [[iten['processoId'], iten['remessaId'], iten['apelido'],iten['usuarioLanca']]  for iten in lista]

    return listaFat