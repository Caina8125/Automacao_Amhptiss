import os
import tkinter.messagebox
import pandas as pd

# def GuiasEscaneadasBrb():
#     pathEcaneadasBrb = r"\\10.0.0.239\guiasscaneadas\2024\BRB"
#     import os

def BuscarGuiasEscaneadas(listas,caminho,variavel):
    nome_pasta = caminho

    caminhoFaturas = []
    statusBusca = []
    arquivo = []

    for index, linha in listas.iterrows():
        arquivoFatura = variavel+linha["Fatura"]+'.pdf'

        # Verifica se a pasta existe
        if not os.path.isdir(nome_pasta):
            tkinter.messagebox.showerror("Erro Autenticação", f"A pasta especificada para buscar as faturas não existe")

        # Verifica se o arquivo existe na pasta
        caminho_arquivo = os.path.join(nome_pasta, arquivoFatura)

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