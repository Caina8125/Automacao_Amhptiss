import os
import tkinter.messagebox

from pandas import DataFrame

# def GuiasEscaneadasBrb():
#     pathEcaneadasBrb = r"\\10.0.0.239\guiasscaneadas\2024\BRB"
#     import os

def BuscarGuiasEscaneadas(listas,caminho, cod_convenio):
    nome_pasta = caminho

    caminhoFaturas = []
    statusBusca = []
    arquivo = []

    for index, linha in listas.iterrows():
        if cod_convenio == 225:
            arquivoFatura = linha["Fatura"]
        if cod_convenio == 433:
            arquivoFatura = f'LOTE {linha["Protocolo"]}.pdf'.replace('.0', '')
        else:
            arquivoFatura = linha["Fatura"]+'.pdf'

        # Verifica se a pasta existe
        if not os.path.isdir(nome_pasta):
            tkinter.messagebox.showerror("", f"A pasta especificada para buscar as faturas não existe")

        # Verifica se o arquivo existe na pasta
        caminho_arquivo = os.path.join(nome_pasta, arquivoFatura)

        if linha['Envio Operadora'] == 'S':
            listas = listas.drop(index)
            continue

        if cod_convenio == 225:
            if os.path.isdir(caminho_arquivo):
                caminhoFaturas.append(caminho_arquivo)
                statusBusca.append("Fatura Encontrada")
                arquivo.append(arquivoFatura)
            else:
                caminhoFaturas.append("Não Encontrado")
                statusBusca.append("Fatura Não Encontrada")
                arquivo.append("Não Encontrado")
                
        else:
            if os.path.isfile(caminho_arquivo):
                caminhoFaturas.append(caminho_arquivo)
                statusBusca.append("Fatura Encontrada")
                arquivo.append(arquivoFatura)
            else:
                caminhoFaturas.append("Não Encontrado")
                statusBusca.append("Fatura Não Encontrada")
                arquivo.append("Não Encontrado")
    
    listas['Status Envio'] = statusBusca
    listas['Arquivo'] = arquivo
    listas['Caminho Fatura'] = caminhoFaturas

    return listas