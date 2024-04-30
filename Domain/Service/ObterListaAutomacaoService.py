import Infra.Repository.ListaAutomacaoRepository as listaAutomacao


def aplicarValidacaoUsuarioSetor(setor):
    match setor:
        case "FATURAMENTO":
            lista = listaAutomacao.obterListaFaturamento()
            return lista
        case "GLOSA":
            lista = listaAutomacao.obterListaGlosa()
            return lista
        case "FINANCEIRO":
            lista = listaAutomacao.obterListaFinanceiro()
            return lista
        case "NOTA_FISCAL":
            lista = listaAutomacao.obterListaTesouraria()
            return lista