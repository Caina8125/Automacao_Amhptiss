import pandas as pd
from tkinter.filedialog import askopenfilename

df_planilha = pd.read_excel(askopenfilename())

df_planilha = df_planilha.loc[df_planilha["Recupera ?"] == "Sim"]
df_planilha['Recursado no Portal'] = ''
df_planilha['Lote'] = ''

colunas_a_serem_utilizadas = [
    "Fatura Inicial", "Protocolo Glosa", "Paciente", "Fatura", "Lote", "Controle Inicial", "Controle",
    "Amhptiss", "Autorização", "Matrícula", "Nro. Guia", "Guia Prestador", "Tipo Guia", "Realizado", "Procedimento", "Descrição", "Valor Cobrado",
    "Valor Glosa", "Valor Recursado", "Recurso Glosa", "Motivo Glosa", "Usuário Glosa", "Recupera ?", "Recursado no Portal"
    ]

df_planilha = df_planilha.loc[:, colunas_a_serem_utilizadas]

lista_de_protocolos = [
    f'{value}'.replace('.0', '')
    for value in list(set(df_planilha['Fatura Inicial'].values.tolist()))
    ]

for protocolo in lista_de_protocolos:
    df_processo = df_planilha.loc[(df_planilha["Fatura Inicial"] == int(protocolo))]

    if df_processo.empty:
        df_processo = df_planilha.loc[(df_planilha["Fatura Inicial"] == protocolo)]

    df_processo = df_processo.sort_values(by=['Paciente', 'Nro. Guia','Procedimento'])
    df_processo.to_excel(f"\\\\10.0.0.239\\automacao_glosa\\Automação Amil\\Planilhas\\{protocolo}.xlsx", sheet_name="Recurso", index=False)

print('Finalizado!')