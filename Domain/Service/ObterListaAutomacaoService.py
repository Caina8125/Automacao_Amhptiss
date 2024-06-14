import Infra.Repository.ListaAutomacaoRepository as listaAutomacao


def aplicarValidacaoUsuarioSetor(setor):
    match setor:
        case "Faturamento":
            lista = listaAutomacao.obterListaFaturamento()
        case "Glosa":
            lista = listaAutomacao.obterListaGlosa()
        case "Financeiro":
            lista = listaAutomacao.obterListaFinanceiro()
        case "NotaFiscal":
            lista = listaAutomacao.obterListaTesouraria()
    
    return lista