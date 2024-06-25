import os
from tkinter.messagebox import showerror
from pandas import DataFrame
from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.common.exceptions import InvalidArgumentException
from selenium.webdriver.common.by import By
from time import sleep

from Application.AppService.EnvioOperadoraService import atualizar_status_envio_operadora

class NextCloudMaida(PageElement):
    input_usuario = By.ID, 'user'
    input_senha = By.ID, 'password'
    input_entrar = By.ID, 'submit-form'
    span_envio_faturamento = By.XPATH, '/html/body/div[3]/div[2]/div[2]/table/tbody/tr[1]/td[2]/a/span[1]/span'
    a_btn_mais = By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[2]/a'
    input_type_file = By.ID, 'file_upload_start'
    input_pesquisa = By.ID, 'searchbox'
    tbody_filelist = By.ID, 'fileList'
    div_conflict = By.CLASS_NAME, 'conflict'
    div_filename = By.CLASS_NAME, 'filename'
    btn_cancelar = By.XPATH, '/html/body/div[8]/div[2]/button[1]'

    def __init__(self, url: str, convenio: int, usuario: str, senha: str) -> None:
        super().__init__(url)
        self.convenio = convenio
        self.usuario = usuario
        self.senha = senha
        self.get_xpath_span_convenio()

    def get_xpath_span_convenio(self):
        match self.convenio:
            case 381:
                self.span_convenio = By.PARTIAL_LINK_TEXT, 'PMDF - FATURAMENTO'
            case 457:
                self.span_convenio = By.PARTIAL_LINK_TEXT, 'PMDF - FATURAMENTO'
            case 433:
                self.span_convenio = By.PARTIAL_LINK_TEXT, 'GDF SAUDE'
            case 160:
                self.span_convenio = By.PARTIAL_LINK_TEXT, 'AMHPDF - SENADO'

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

    def gerar_lista_paths(self):
        if self.convenio == 160:
            ...
        else:
            ...

    def inicia_automacao(self, **kwargs):
        df_treeview: DataFrame = kwargs.get('df_treeview')
        token: str = kwargs.get('token')
        try:
            paths_list = [path for path in df_treeview['Caminho'].values.tolist() if "Não Encontrado" not in path]

            self.init_driver(py_auto_gui=True)
            self.login()
            self.caminho()

            self.driver.find_element(*self.input_type_file).send_keys('\n'.join(paths_list))
            sleep(2)

            df_treeview['Status Fatura'] = df_treeview['Status Fatura'].replace('Fatura Encontrada', 'Parcialmente Enviada')

            while 'poucos segundos' in self.driver.find_element(*self.body).text:
                sleep(1.5)

            sleep(2)

            if 'Quais arquivos você quer manter?' in self.driver.find_element(*self.body).text:
                div_conflict_list = self.driver.find_elements(*self.div_conflict)

                for div_conflict in div_conflict_list:
                    lote_conflito = div_conflict.find_element(*self.div_filename).text
                    df_treeview.loc[df_treeview['Nome Arquivo'] == lote_conflito, 'Status Fatura'] = 'Já anexada anteriormente'
                    fatura = df_treeview.loc[df_treeview['Nome Arquivo'] == lote_conflito, 'Fatura'].values.tolist()[0]
                    atualizar_status_envio_operadora('normal', fatura, "S", token)

                self.driver.find_element(*self.btn_cancelar).click()
                sleep(1)    

            df_reduzido = df_treeview[['Fatura', 'Status Fatura', 'Caminho']]
            files_names_list = [
                {
                    'fatura': value[0],
                    'path': os.path.basename(value[2])
                }
                for value in df_reduzido.values.tolist()
                if "Não Encontrado" not in value[1] and 'Já anexada anteriormente' not in value[1]
            ]

            for file in files_names_list:
                self.driver.find_element(*self.input_pesquisa).send_keys(file['path'])
                sleep(1)

                tbody_content = self.driver.find_element(*self.tbody_filelist).text

                if file['path'] in tbody_content.replace('\n', ''):
                    atualizar_status_envio_operadora('normal', file['fatura'], "S", token)
                    df_treeview.loc[df_treeview['Fatura'] == file['fatura'], 'Status Fatura'] = 'Enviada'

                self.driver.find_element(*self.input_pesquisa).clear()

        except Exception as e:
            showerror('', f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__}:\n{e}")
        
        try:
            self.driver.quit()
        except:
            pass
        
        return df_treeview

            