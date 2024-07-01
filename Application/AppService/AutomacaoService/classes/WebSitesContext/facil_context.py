from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.webdriver.common.by import By

class FacilContext(PageElement):
    input_usuario = By.ID, 'login-entry'
    input_senha = By.ID, 'password-entry'
    btn_entrar = By.ID, 'BtnEntrar'
    i_perfil = By.XPATH, '/html/body/header/nav/div/div[2]/ul[2]/li[5]/a/i'
    url_envio_xml = '/GuiasTISS/RecebeXMLTISS/LoteXml'
    btn_pesquisar = By.XPATH, '/html/body/main/div[1]/div[1]/div[2]/div[2]/button[1]'
    option_100_item = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]/div[11]/div[1]/div[2]/div/div/select/option[7]'
    div_body_panel = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]'
    div_faturas_enviadas = By.CLASS_NAME, 'enviado'
    a_proxima_pag_xml = By.XPATH, '/html/body/main/div[1]/div[2]/div/div[2]/div[101]/div[1]/div[1]/div/nav/ul/li[8]/a'
    div_col_xml = By.CLASS_NAME, 'col-md-4'
    i_clip = By.CLASS_NAME, 'fa-paperclip'
    modal_envio_arquivo = By.XPATH, '/html/body/main/div[1]/div[4]/div/div'
    input_type_file = By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[2]/div/div/div/div[2]/div[1]/input-file/div/div/div[1]/div/div/input'
    fechar_modal = By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[3]/button'
    adicionar_arquivo = By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[2]/div/div/div/div[2]/div[1]/input-file/div/div/div[1]/div/div/div[1]/div[2]/button'
    span_arquivo = By.XPATH, '/html/body/main/div[1]/div[4]/div/div/div[2]/div/div/div/div[2]/div[2]/div/div/table/tbody/tr/td[1]/span'
    clip1 = By.XPATH, '/html/body/main/div/div[2]/div/div[2]/div[1]/div[1]/div/div[5]/button[1]'
    btn_ignorar = By.ID, 'btnIngorar'

    def __init__(self, url, codigo_convenio) -> None:
        super().__init__(url)
        self.set_atributos_convenio(codigo_convenio)
        
    def set_atributos_convenio(self, codigo_convenio):
        match codigo_convenio:
            case 216:
                self.option_tipo_acesso = By.XPATH, '/html/body/main/div[2]/div[1]/form/table/tbody/tr[2]/td[2]/select/option[6]'
                self.qtd_paginas = 3

    