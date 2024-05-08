import Infra.Repository.AutenticacaoRepository as AuthRepository
import Domain.Service.AutenticacaoService as AutenticacaoService

def Autenticar(login, senha, setor):
    content = AuthRepository.auth(login,senha)
    roules = content['UsuarioToken']["Claims"]
    token = content['AccessToken']
    verificaSetor = AutenticacaoService.validarSetor(roules, setor)
    validacao = []
    validacao.append(verificaSetor)
    validacao.append(token)
    return validacao

