from datetime import datetime
import os
from tkinter.messagebox import showerror
from pandas import DataFrame
from Application.AppService.AutomacaoService.classes.WebSitesContext.next_cloud_maida_context import NextCloudMaidaContext
from selenium.common.exceptions import InvalidArgumentException
from selenium.webdriver.remote.webelement import WebElement
from time import sleep

from Application.AppService.EnvioOperadoraService import atualizar_status_envio_operadora
from Infra.Core.core import Core

class NextCloudMaidaEnvio(NextCloudMaidaContext):

    def login(self):
        self.driver.find_element(*self.input_usuario).send_keys(self.usuario)
        sleep(2)
        self.driver.find_element(*self.input_senha).send_keys(self.senha)
        sleep(2)
        self.driver.find_element(*self.input_entrar).click()
        sleep(2)

    def caminho(self):
        self.driver.find_element(*self.span_convenio).click()
        sleep(1)
        self.driver.find_element(*self.span_envio_faturamento).click()
        sleep(1)

    def inicia_automacao(self, **kwargs):
        self.df_treeview: DataFrame = kwargs.get('df_treeview')
        token: str = kwargs.get('token')
        data_atual = datetime.now()
        new_dir_name = "output/SIS_" + data_atual.strftime('%d_%m_%Y_%H_%M_%S')
        try:
            if self.convenio == 160:
                self.preparar_arquivos_com_nota(new_dir_name)
                paths_list = [os.path.abspath(os.path.join(new_dir_name, file)) for file in os.listdir(new_dir_name)]

            else:
                paths_list = [path for path in self.df_treeview['Caminho'].values.tolist() if "Não Encontrado" not in path]

            self.init_driver(py_auto_gui=True)
            self.login()
            self.caminho()

            self.driver.find_element(*self.input_type_file).send_keys('\n'.join(paths_list))
            sleep(2)

            self.df_treeview['Status Fatura'] = self.df_treeview['Status Fatura'].replace('Fatura Encontrada', 'Parcialmente Enviada')

            while 'poucos segundos' in self.driver.find_element(*self.body).text:
                sleep(1.5)

            sleep(2)

            #Este if não irá conferir os arquivos conflitantes do SIS. A conferência dos arquivos do SIS é feita de maneira diferente mais abaixo
            if 'Quais arquivos você quer manter?' in self.driver.find_element(*self.body).text:
                div_conflict_list = self.driver.find_elements(*self.div_conflict)

                if self.convenio != 160:
                    self.conferir_arquivos_conflitantes(div_conflict_list)

                self.driver.find_element(*self.btn_cancelar).click()
                sleep(1)

            df_reduzido = self.df_treeview[['Fatura', 'Status Fatura', 'Caminho', 'Nome Arquivo']]
            files_dict_list = self.get_file_names_list(df_reduzido)

            if self.convenio != 160:
                self.conferencia_arquivos(files_dict_list, token)
            else:
                self.conferencia_arquivos_sis(files_dict_list, token)

        except Exception as e:
            showerror('', f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__}:\n{e}")
        
        try:
            self.driver.quit()
        except:
            pass
        
        self.df_treeview.drop('Nota Fiscal', axis=1)

        return self.df_treeview
    
    def conferir_arquivos_conflitantes(self, div_conflict_list: list[WebElement], token: str):
        for div_conflict in div_conflict_list:
            lote_conflito = div_conflict.find_element(*self.div_filename).text
            self.df_treeview.loc[self.df_treeview['Nome Arquivo'] == lote_conflito, 'Status Fatura'] = 'Já anexada anteriormente'
            fatura = self.df_treeview.loc[self.df_treeview['Nome Arquivo'] == lote_conflito, 'Fatura'].values.tolist()[0]
            atualizar_status_envio_operadora('normal', fatura, "S", token)
    
    def get_file_names_list(self, df_reduzido: DataFrame):
        if self.convenio != 160:
            files_names_list = [
                {
                    'fatura': value[0],
                    'path': os.path.basename(value[3])
                }
                for value in df_reduzido.values.tolist()
                if "Não Encontrado" not in value[1] and 'Já anexada anteriormente' not in value[1]
            ]
        else:
            files_names_list = [
                {
                    'fatura': value[0],
                    'paths': [os.path.basename(string_value) for string_value in value[2].split(', ')]
                }
                for value in df_reduzido.values.tolist()
                if "Não Encontrado" not in value[1] and 'Já anexada anteriormente' not in value[1]
            ]
        
        return files_names_list
    
    def conferencia_arquivos(self, files_names_list, token):
        for file in files_names_list:
            self.driver.find_element(*self.input_pesquisa).send_keys(file['path'])
            sleep(1)

            tbody_content = self.driver.find_element(*self.tbody_filelist).text

            if file['path'] in tbody_content.replace('\n', ''):
                atualizar_status_envio_operadora('normal', file['fatura'], "S", token)
                self.df_treeview.loc[self.df_treeview['Fatura'] == file['fatura'], 'Status Fatura'] = 'Enviada'

            self.driver.find_element(*self.input_pesquisa).clear()

    def conferencia_arquivos_sis(self, files_names_list, token):
        arquivos_enviados = True

        for file in files_names_list:
            for path in file['paths']:
                self.driver.find_element(*self.input_pesquisa).send_keys(path)
                sleep(1)

                tbody_content = self.driver.find_element(*self.tbody_filelist).text

                if path not in tbody_content.replace('\n', ''):
                    arquivos_enviados = False
                
                self.driver.find_element(*self.input_pesquisa).clear()

            if arquivos_enviados:
                atualizar_status_envio_operadora('normal', file['fatura'], "S", token)
                self.df_treeview.loc[self.df_treeview['Fatura'] == file['fatura'], 'Status Fatura'] = 'Enviada'

            arquivos_enviados = True

    def mover_arquivos(self, path_arquivo, remessa, nota_fiscal, new_files_names_list, new_dir_name):
        pdf_path, xml_path = Core.obter_caminhos_nf(remessa, nota_fiscal)

        if pdf_path == None or xml_path == None:
            return False

        paths_dict = {
            path_arquivo: new_files_names_list[0],
            pdf_path: new_files_names_list[1],
            xml_path: new_files_names_list[2] 
        }

        return Core.criar_pasta_e_armazenar_arquivos(new_dir_name, paths_dict)
    
    def preparar_arquivos_com_nota(self, new_dir_name):
        for _, linha in self.df_treeview.iterrows():
            remessa = f"{linha['Remessa']}"
            protocolo = f"{linha['Protocolo']}"
            path_arquivo = linha['Caminho']
            nota = linha['Nota Fiscal']

            new_files_names_list = [f'PEG {protocolo}.pdf', f'PEG {protocolo} NF.pdf', f'PEG {protocolo} NF.xml']
            
            if not self.mover_arquivos(path_arquivo, remessa, nota, new_files_names_list, new_dir_name):
                self.df_treeview.loc[self.df_treeview['Protocolo'] == int(protocolo), 'Status Fatura'] = 'Nota fiscal não encontrada'
            
            else:
                self.df_treeview.loc[self.df_treeview['Protocolo'] == int(protocolo), 'Caminho'] = (', ').join([os.path.join(new_dir_name, file) for file in new_files_names_list])