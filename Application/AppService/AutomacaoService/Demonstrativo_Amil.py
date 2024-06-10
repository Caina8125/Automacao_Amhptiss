import os
import pandas as pd
from selenium.webdriver.common.by import By
import time
from Application.AppService.AutomacaoService.page_element import PageElement

class BaixarDemonstrativoAmil(PageElement):
    fechar = (By.ID, 'finalizar-walktour')
    usuario_input = (By.ID, 'login-usuario')
    senha_input = (By.ID, 'login-senha')
    entrar = (By.XPATH, '/html/body/as-main-app/as-login-container/div[1]/div/as-login/div[2]/form/fieldset/button')
    amil_logo = (By.XPATH, '/html/body/as-main-app/as-comunicado-detalhe/as-base-layout/as-header-container/header/div/h1/a')
    acesso_sis_amil = (By.XPATH, '/html/body/as-main-app/as-comunicado-detalhe/as-base-layout/section/div/as-navbar/nav/div[2]/form/button')
    menu = (By.XPATH, '/html/body/div[2]/div[3]/div')
    portal_prestador = (By.XPATH, '/html/body/div[2]/h1/a')
    consultas_e_relatorios = (By.XPATH, '/html/body/div[2]/div/h2[3]/a')
    dem_tiss_3 = (By.XPATH, '/html/body/div[2]/div/div[3]/h3[2]/a')
    mes_anterior = (By.XPATH, '/html/body/fieldset[2]/center/table/tbody/tr/td/table/tbody/tr/td/form[1]/table/tbody/tr[2]/td[1]/b/a')
    table = (By.XPATH, '/html/body/form[1]/table/tbody/tr[5]/td/table')
    imprimir = (By.XPATH, '/html/body/div/span[2]/img')
    xml = (By.XPATH, '/html/body/div/span[3]/img')
    voltar = (By.XPATH, '/html/body/div/span[1]/img')

    def __init__(self, url, usuario, senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.fechar).click()
        time.sleep(2)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(5)
    
    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.amil_logo)
        self.driver.execute_script('scrollBy(0,1000)')
        self.driver.execute_script('scrollBy(0,1000)')
        self.driver.execute_script('scrollBy(0,1000)')
        time.sleep(2)
        self.driver.find_element(*self.acesso_sis_amil).click()
        time.sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        time.sleep(2)
        self.driver.find_element(*self.menu).click()
        time.sleep(2)
        self.driver.switch_to.frame('menu')
        self.driver.find_element(*self.portal_prestador).click()
        time.sleep(2)
        self.driver.find_element(*self.consultas_e_relatorios).click()
        time.sleep(2)
        self.driver.find_element(*self.dem_tiss_3).click()
        time.sleep(2)
        self.driver.switch_to.default_content()
        time.sleep(2)
        global titulo_pagina
        titulo_pagina = self.driver.find_element(By.XPATH, '/html/head/title').get_attribute('innerHTML')
        self.driver.switch_to.frame('principal')
        time.sleep(2)
        self.driver.find_element(*self.mes_anterior).click()
        time.sleep(2)

    def baixar_demonstrativo(self):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()
        
        tabela = self.driver.find_element(*self.table).get_attribute('outerHTML')
        tabela = pd.read_html(tabela)[0]
        tamanho_tabela = len(tabela)

        for i in range(2, tamanho_tabela + 1):
            xpath_valor_glosado = (By.XPATH, f"/html/body/form[1]/table/tbody/tr[5]/td/table/tbody/tr[{i}]/td[6]/font")
            xpath_nmr_lote = (By.XPATH, f"/html/body/form[1]/table/tbody/tr[5]/td/table/tbody/tr[{i}]/td[2]/font/a")
            valor_glosado = self.driver.find_element(*xpath_valor_glosado).text
            
            if valor_glosado != "0,00":
                numero_processo = self.driver.find_element(*xpath_nmr_lote).text
                self.driver.find_element(*xpath_nmr_lote).click()
                time.sleep(2)
                self.driver.switch_to.default_content()
                time.sleep(2)
                self.driver.switch_to.frame('toolbar')
                time.sleep(2)
                self.driver.find_element(*self.imprimir).click()
                time.sleep(3)
                endereco = r"\\10.0.0.239\automacao_financeiro\AMIL"
                novo_nome = f"{endereco}\\{numero_processo}.pdf"
                contador = 0

                while contador < 20:
                    try:
                        os.rename(f"{endereco}\\{titulo_pagina}.pdf", novo_nome)
                        break
                    
                    except:
                        print("Download ainda não foi feito")
                        contador += 1

                        if contador == 18:
                            time.sleep(10)
                        
                        else:
                            time.sleep(2)

                self.driver.find_element(*self.xml).click()
                time.sleep(5)
                self.driver.switch_to.default_content()
                time.sleep(2)
                self.driver.switch_to.frame('principal')