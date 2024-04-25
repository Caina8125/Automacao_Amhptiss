import Infra.Repository.AutenticacaoRepository as AuthRepository
import Domain.Service.AutenticacaoService as AutenticacaoService

def Autenticar(login, senha, setor):
    roules = AuthRepository.auth(login,senha)
    verificaSetor = AutenticacaoService.validarSetor(roules, setor)
    return verificaSetor

