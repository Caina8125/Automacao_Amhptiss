from Infra.Repository.ApiFaturamentoRepository import post_status_envio_operadora

def atualizar_status_envio_operadora(self, tipo_negociacao, remessa_id, envio_operadora, token):
    response = post_status_envio_operadora(tipo_negociacao, remessa_id, envio_operadora, token)
    if response == 200:
        return True
    
    else:
        return False