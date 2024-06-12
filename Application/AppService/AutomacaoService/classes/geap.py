from datetime import date
from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from time import sleep

class Geap(PageElement):
    acessar_portal = By.XPATH, '/html/body/div[3]/div[3]/div[1]/form/div[1]/div[1]/div/a'
    usuario = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[1]/div/div[1]/div/input'
    input_senha = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[2]/div/div[1]/div[1]/input'
    entrar = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[2]/button/div[2]/div/div'
    fechar = By.XPATH, '/html/body/div[4]/div[2]/div/div[3]/button'
    portal_tiss = (By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[1]/div[1]/div/div')
    alerta = By.XPATH,' /html/body/div[2]/div/center/a'
    a_tiss = By.XPATH, '/html/body/div[1]/nav/ul/li[3]'
    anexo_conta = By.XPATH, '/html/body/div[1]/nav/ul/li[3]/ul/li[5]/a'
    select_mes_referencia = By.ID, 'dropReferencia'
    select_nr = By.ID, 'dropNr'
    select_conta = By.ID, 'dropConta'
    select_tipo_anexo = By.ID, 'dropTipoAnexo'
    text_area_decricao = By.ID, 'txtAnexoDesc'
    btn_processar = By.ID, 'btnFinalizar'
    input_file = By.XPATH, '/html/body/input[2]'
    btn_remover_anexos = By.ID, 'clear-dropzone'

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

    def click_option(self, select_tuple, text):
        select = self.driver.find_element(*select_tuple)
        option_element = next((
            option
            for option in select.find_elements(By.TAG_NAME, 'option')
            if text in option.text or option.text == text
        ), None)
        if option_element:
            option_element.click()
            return True
        else:
            return False

    def inicia_automacao(self, **kwargs):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        setor = kwargs.get('setor')
        dados_remessa = kwargs.get('dados_faturas')
        mes_atual = date.today().strftime('%m/%Y')
        sleep(1.5)

        for fatura in dados_remessa['faturas']:
            n_processo = fatura['fatura']
            protocolo = fatura['protocolo']
            guias = fatura['lista_guias']
            self.click_option(self.select_mes_referencia, mes_atual)

            if not self.click_option(self.select_nr, protocolo):
                #TODO não lançar que essa fatura foi enviada
                ...
            
            for guia_data in guias:
                self.click_option(self.select_mes_referencia, mes_atual)
                sleep(2)
                self.click_option(self.select_nr, protocolo)
                guia = f"{guia_data['guia']}".replace('.0', '')
                caminho = guia_data['caminho_guia']
                sleep(2)
                self.click_option(self.select_conta, guia)

                if not self.click_option(self.select_conta, guia):
                    raise Exception(f'Guia {guia} não foi encontrada no portal! Por favor, verificar se o número diverge com o que está no portal.')
                
                sleep(2)

                self.click_option(self.select_tipo_anexo, 'COMPROVANTE DE COMPARECIMENTO (ASSINATURA)')
                sleep(2)

                self.driver.find_element(*self.text_area_decricao).send_keys("Guia de faturamento.")
                sleep(2)
                self.driver.find_element(*self.input_file).send_keys(caminho)
                sleep(2)
                self.driver.find_element(*self.btn_processar).click()
                sleep(3)
                if 'Ocorreu um erro ao salvar os dados' in self.driver.find_element(*self.body).text:
                    self.driver.find_element(*self.text_area_decricao).clear()
                    sleep(2)
                    self.driver.find_element(*self.btn_remover_anexos).click()
                    sleep(2)