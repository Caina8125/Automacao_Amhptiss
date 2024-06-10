from datetime import date
from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from seleniumwire import webdriver

class Geap(PageElement):
    acessar_portal = By.XPATH, '/html/body/div[3]/div[3]/div[1]/form/div[1]/div[1]/div/a'
    usuario = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[1]/div/div[1]/div/input'
    input_senha = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[2]/div/div[1]/div[1]/input'
    entrar = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[2]/button/div[2]/div/div'
    fechar = By.XPATH, '/html/body/div[4]/div[2]/div/div[3]/button'
    portal_tiss = (By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[1]/div[1]/div/div')
    alerta = By.XPATH,' /html/body/div[2]/div/center/a'
    a_tiss = By.XPATH, '/html/body/div[1]/nav/ul/li[3]'
    anexo_conta = By.XPATH, '/html/body/div[1]/nav/ul/li[3]/ul/li[7]/a'
    select_mes_referencia = By.ID, 'dropReferencia'
    select_nr = By.ID, 'dropNr'
    select_conta = By.ID, 'dropConta'
    select_tipo_anexo = By.ID, 'dropTipoAnexo'
    text_area_decricao = By.ID, 'txtAnexoDesc'
    btn_processar = By.ID, 'btnFinalizar'
    input_file = By.XPATH, '/html/body/input[2]'

    def __init__(self, url: str, cpf: str, senha: str) -> None:
        super().__init__(url)
        self.cpf = cpf
        self.senha = senha

    def exe_login(self):
        sleep(4)
        try:
            self.driver.find_element(*self.fechar).click()
        except:
            pass
        sleep(2)
        try:
            self.driver.implicitly_wait(15)
            self.driver.find_element(*self.acessar_portal).click()
            self.driver.implicitly_wait(4)
            self.driver.find_element(*self.usuario).send_keys(self.cpf)
            self.driver.find_element(*self.input_senha).click()
            sleep(2)
            self.driver.find_element(*self.input_senha).send_keys(self.senha)
            sleep(2)
            self.driver.find_element(*self.entrar).click()
            self.driver.find_element(*self.entrar).click()
            sleep(2)
            self.driver.find_element(*self.portal_tiss)

        except Exception as e:
            self.driver.implicitly_wait(180)
            self.driver.find_element(*self.portal_tiss)
            sleep(2)
            self.driver.implicitly_wait(15)


    def exe_caminho(self):
        sleep(5)
        self.driver.implicitly_wait(3)
        try:
            lista = [element for element in self.driver.find_elements(By.TAG_NAME, 'i') if element.text == 'close']
            sleep(1)
            for _ in range(0, len(lista)):
                for element in lista:
                    try:
                        element.click()
                    except:
                        pass
        except:
            pass
        self.driver.find_element(*self.portal_tiss).click()
        sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.implicitly_wait(4)
        try:
            self.driver.find_element(*self.alerta).click()
        except:
            print('Alerta não apareceu')
        self.driver.implicitly_wait(15)
        sleep(2)
        # self.driver.find_element(*self.a_tiss).click()
        # sleep(2)
        self.driver.find_element(*self.anexo_conta).click()
        sleep(2)

    def inicia_automacao(self, **kwargs):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        setor = kwargs.get('setor')
        dados_remessa = kwargs.get('dados_faturas')
        mes_atual = date.today().strftime('%m/%Y')

        select_mes_referencia_element = self.driver.find_element(*self.select_mes_referencia)
        select_nr_element = self.driver.find_element(*self.select_nr)
        select_conta = self.driver.find_element(*self.select_conta)
        select_tipo_anexo = self.driver.find_element(*self.select_tipo_anexo)
        sleep(1.5)

        option_mes_atual = next((
            option
            for option in select_mes_referencia_element.find_elements(By.TAG_NAME, 'option')
            if option.get_attribute('value') == mes_atual
        ), None)

        for fatura in dados_remessa['faturas']:
            n_processo = fatura['fatura']
            protocolo = fatura['protocolo']
            guias = fatura['lista_guias']

            option_fatura = next((
                option
                for option in select_nr_element.find_elements(By.TAG_NAME, 'option')
                if option.text == protocolo
            ), None)

            if option_fatura == None:
                #TODO não lançar que essa fatura foi enviada
                ...
            
            for guia_data in guias:
                option_mes_atual.click()
                sleep(2)
                option_fatura.click()
                guia = f"{guia_data['guia']}".replace('.0', '')
                caminho = guia_data['caminho_guia']
                sleep(2)
                option_guia = next((
                    option
                    for option in select_conta.find_elements(By.TAG_NAME, 'option')
                    if option.text == guia
                ), None)

                if option_guia == None:
                    raise Exception(f'Guia {guia} não foi encontrada no portal! Por favor, verificar se o número diverge com o que está no portal.')

                option_guia.click()
                sleep(2)

                
                option_tipo_anexo = next((
                    option
                    for option in select_tipo_anexo.find_elements(By.TAG_NAME, 'option')
                    if 'COMPROVANTE DE COMPARECIMENTO (ASSINATURA)' in option.text
                ), None)

                # TODO caso for none

                option_tipo_anexo.click()
                sleep(2)

                self.driver.find_element(*self.text_area_decricao).send_keys("Guia de faturamento.")
                sleep(2)
                self.driver.find_element(*self.input_file).send_keys(caminho)
                sleep(1.5)
                self.driver.find_element(*self.btn_processar).click()
                sleep(1.5)