import Infra.Repository.ListaAutomacaoRepository as listaAutomacao


def aplicarValidacaoUsuarioSetor(setor):
    match setor:
        case "Faturamento":
            lista = listaAutomacao.obterListaFaturamento()
            return lista
        case "Glosa":
            lista = listaAutomacao.obterListaGlosa()
            return lista
        case "Financeiro":
            lista = listaAutomacao.obterListaFinanceiro()
            return lista
        case "Nota_Fiscal":
            lista = listaAutomacao.obterListaTesouraria()
            return lista