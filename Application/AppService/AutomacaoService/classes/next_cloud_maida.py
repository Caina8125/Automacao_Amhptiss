import os
from pandas import DataFrame
from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.common.exceptions import InvalidArgumentException
from selenium.webdriver.common.by import By
from time import sleep

class NextCloudMaida(PageElement):
    input_usuario = By.ID, 'user'
    input_senha = By.ID, 'password'
    input_entrar = By.ID, 'submit-form'
    span_envio_faturamento = By.XPATH, '/html/body/div[3]/div[2]/div[2]/table/tbody/tr[1]/td[2]/a/span[1]/span'
    a_btn_mais = By.XPATH, '/html/body/div[3]/div[2]/div[2]/div[1]/div[2]/a'
    input_type_file = By.ID, 'file_upload_start'
    input_pesquisa = By.ID, 'searchbox'

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

    def inicia_automacao(self, **kwargs):
        df_treeview: DataFrame = kwargs.get('df_treeview')
        paths_list = [path for path in df_treeview['Caminho'].values.tolist() if "Não Encontrado" not in path]
        files_names_list = [os.path.basename(file) for file in paths_list]

        self.init_driver(py_auto_gui=True)
        self.login()
        self.caminho()

        self.driver.find_element(*self.input_type_file).send_keys('\n'.join(paths_list))
        sleep(2)

        while 'poucos segundos' in self.driver.find_element(*self.body).text:
            sleep(1)

        for file in files_names_list:
            self.driver.find_element(*self.pes)

        return df_treeview