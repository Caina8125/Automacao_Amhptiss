from Infra.Repository.ApiFaturamentoRepository import api_fatura

def obterListaGuias(tipoNegociacao, convenio_id, processo_id,token):
    lista = api_fatura(tipoNegociacao, convenio_id, processo_id, token)
    lista_guia = list(set([guia['atendimentoAmhptissId'] for guia in lista['atendimentosProcesso']]))

    return lista_guia

def obterNotaFiscalFatura(tipoNegociacao, convenio_id, processo_id,token):
    lista = api_fatura(tipoNegociacao, convenio_id, processo_id, token)
    nf: str | None = lista['atendimentosProcesso'][0]['notaFiscal']
    return nf