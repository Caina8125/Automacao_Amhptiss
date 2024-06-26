from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.webdriver.common.by import By

class NextCloudMaidaContext(PageElement):
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