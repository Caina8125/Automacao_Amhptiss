from Infra.Repository.ApiFaturamentoRepository import post_status_envio_operadora

def atualizar_status_envio_operadora(tipo_negociacao, processo_id, envio, token):
    response = post_status_envio_operadora(tipo_negociacao, processo_id, envio, token)
    if response == 204:
        return True
    
    else:
        return False