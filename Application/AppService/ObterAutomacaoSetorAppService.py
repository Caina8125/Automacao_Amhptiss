import Domain.Service.ObterListaAutomacaoService as obterListas

def obterListaSetor(setor):
    lista = obterListas.aplicarValidacaoUsuarioSetor(setor)
    return lista