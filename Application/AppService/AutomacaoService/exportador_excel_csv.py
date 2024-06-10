from pandas import DataFrame
from tkinter import filedialog
from pandas import read_excel
from unidecode import unidecode
import sys

def substitui_valor(valor: str):
    return str(valor.replace('_', ''))

def converte_data(valor):
    return valor.strftime('%d/%m/%Y')

def converte_para_lower(valor: str):
    return valor.lower()

def converte_para_int(valor):
    int(str(valor).replace('.0', ''))

def converte_str(valor:str):
    return unidecode(valor.upper())

planilha = 'Planilha'
while 'xlsx' not in planilha:
    planilha = filedialog.askopenfilename()

    if not planilha:
        sys.exit()

df_planilha: DataFrame = read_excel(planilha, header=6)
df_planilha = df_planilha.drop('CID', axis=1)
df_planilha = df_planilha.drop('Indicação Clínica', axis=1)
df_planilha = df_planilha.dropna(0)
df_planilha['matricula_empresa'] = ''
df_planilha['cnpj_convenio'] = ''
df_planilha['nom_convenio'] = ''
df_planilha['uf_conselho_medico'] = 'DF'
df_planilha['carater_atendimento'] = '1-Eletiva'
df_planilha["Cliente"] = df_planilha["Cliente"].apply(converte_str)
df_planilha["Prof. executante"] = df_planilha["Prof. executante"].apply(converte_str)
df_planilha["Carteirinha"] = df_planilha["Carteirinha"].apply(substitui_valor)
df_planilha["Dt. Sessão"] = df_planilha["Dt. Sessão"].apply(converte_data)
df_planilha["Atendimento"] = df_planilha["Atendimento"].apply(converte_para_lower)
df_planilha["ID"] = df_planilha["ID"].astype(int)
df_planilha["CRP"] = df_planilha["CRP"].astype(int)
df_planilha["Cód. Procedimento"] = df_planilha["Cód. Procedimento"].astype(int)
colunas_a_serem_utilizadas = [
    'matricula_empresa', 'ID', 'Atendimento', 'Dt. Sessão', 'Cliente', 'cnpj_convenio', 'nom_convenio', 'Carteirinha',
    'Cód. Procedimento', 'CRP', 'Prof. executante', 'CBO', 'Conselho Médico', 'uf_conselho_medico', 'carater_atendimento'
    ]

novas_colunas = ['matricula_empresa', 'num_guia', 'tipo_guia_tiss', 'dat_atendimento', 'nom_paciente', 'cnpj_convenio', 'nom_convenio', 'num_matricula',
                'cod_procedimento', 'num_crm', 'nom_medico', 'cbo_medico', 'conselho_medico', 'uf_conselho_medico', 'carater_atendimento']

df_planilha = df_planilha.loc[:, colunas_a_serem_utilizadas]
df_planilha.columns = novas_colunas
df_planilha.to_csv('relacao_guias.csv', ';', index=False, float_format='%.0f')