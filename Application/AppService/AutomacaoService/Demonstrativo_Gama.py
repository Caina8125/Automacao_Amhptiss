import pandas as pd
from selenium.webdriver.common.by import By
import time
import tkinter
from Application.AppService.AutomacaoService.page_element import PageElement

class BaixarDemonstrativosGama(PageElement):
    usuario_input = (By.XPATH, '//*[@id="username"]')
    senha_input = (By.XPATH, '//*[@id="password"]')
    entrar = (By.XPATH, '//*[@id="submitPrestador"]')
    gama = (By.XPATH, '//*[@id="TB_ajaxContent"]/div/div[1]/div/div/a/img')
    extrato = (By.XPATH, '//*[@id="header"]/div[2]/div/div[2]/ul/li[4]/a')
    demonstrativo_tiss = (By.XPATH, '//*[@id="header"]/div[2]/div/div[2]/ul/li[4]/ul/li[2]/a')
    periodo_30_dias = (By.XPATH, '//*[@id="cmbPeriodo"]/option[2]')
    consultar = (By.XPATH, '//*[@id="btnConsultarExtratoPeriodo"]')
    detalhar_demonstrativo = (By.XPATH, '/html/body/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[4]/form/a')
    botao_aceitar_mensagem = (By.XPATH, '//*[@id="TB_ajaxContent"]/div/div/div[2]/div/a')
    tabela = (By.XPATH, '/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table')
    baixar_xml = (By.XPATH, '//*[@id="demonstrativoLotePagamento"]/div[2]/a')
    fechar = (By.XPATH, '//*[@id="demonstrativoLotePagamento"]/a')

    def __init__(self, url, usuario, senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        self.driver.find_element(*self.entrar).click()

    def fechar_comunicado(self):

        while True:
            try:
                self.driver.implicitly_wait(5)
                self.driver.find_element(*self.botao_aceitar_mensagem).click()
                time.sleep(1.5)

            except:
                break


    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.gama).click()
        self.fechar_comunicado()
        time.sleep(2)
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.extrato).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativo_tiss).click()
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.periodo_30_dias).click()
        time.sleep(2)
        self.driver.find_element(*self.consultar).click()
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.detalhar_demonstrativo).click()
        time.sleep(4)

    def inica_automacao(self, **kwargs):
        self.driver.implicitly_wait(30)
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        tabela = self.driver.find_element(*self.tabela)
        codigo_html_tabela = tabela.get_attribute('outerHTML')
        data_frame_tabela = pd.read_html(codigo_html_tabela)[0]
        quantidade_de_demonstrativos = len(data_frame_tabela) + 1

        for i in range(1, quantidade_de_demonstrativos):
            valor_glosa = self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table/tbody/tr[{i}]/td[7]').text
            time.sleep(2)

            if not valor_glosa == 'R$ 0,00':
                numero_do_lote_clicado = False

                while numero_do_lote_clicado == False:
                    try:
                        self.driver.find_element(By.XPATH, f'/html/body/div[2]/div[1]/div/div[2]/div[1]/div[7]/table/tbody/tr[{i}]/td[1]/a').click()
                        numero_do_lote_clicado = True
                    except:
                        self.driver.execute_script('scrollBy(0,50)')

                time.sleep(3)
                self.driver.implicitly_wait(60)
                time.sleep(1.5)
                self.driver.find_element(*self.baixar_xml).click()
                time.sleep(2)
                self.driver.find_element(*self.fechar).click()
                time.sleep(2)

            else:
                continue

        tkinter.messagebox.showinfo( 'Demonstrativos Gama' , f"Downloads concluídos!" )
        self.driver.quit()