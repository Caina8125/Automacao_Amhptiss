from pandas import read_excel


def get_lista_clinicas(convenio: str):
    try:
        logins_list = []
        path = f"\\\\10.0.0.239\\automacao_integralis\\{convenio}\\Logins e Senhas.xlsx"
        df_planilha_login = read_excel(path).sort_values(by=['Clinica'])

        for _, linha in df_planilha_login.iterrows():
            logins_list.append(
                {
                    'nome': linha['Clinica'],
                    'user': f'{linha["Login"]}'.replace('.0', ''),
                    'password': f'{linha["Senha"]}'.replace('.0', '')
                }
            )
        
        return logins_list
    
    except:
        pass