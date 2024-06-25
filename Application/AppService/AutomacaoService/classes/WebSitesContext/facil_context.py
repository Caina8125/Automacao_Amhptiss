from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.webdriver.common.by import By

class FacilContext(PageElement):
    input_usuario = By.ID, 'login-entry'
    input_senha = By.ID, 'password-entry'
    btn_entrar = By.ID, 'BtnEntrar'
    url_envio_xml = '/GuiasTISS/RecebeXMLTISS/LoteXml'
    btn_pesquisar = By.PARTIAL_LINK_TEXT, 'Pesquisar'
    option_100_item = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]/div[101]/div[1]/div[2]/div/div/select/option[7]'
    div_body_panel = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]'
    div_faturas_enviadas = By.CLASS_NAME, 'enviado'
    a_proxima_pag_xml = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]/div[101]/div[1]/div[1]/div/nav/ul/li[8]/a'

    def __init__(self, url, codigo_converio) -> None:
        super().__init__(url)
        self.set_tipo_acesso(codigo_converio)
        
    def set_tipo_acesso(self, codigo_convenio):
        match codigo_convenio:
            case 216:
                self.option_tipo_acesso = By.XPATH, '/html/body/main/div[2]/div[1]/form/table/tbody/tr[2]/td[2]/select/option[6]'
                # TODO fechar modal = By.XPATH, ''