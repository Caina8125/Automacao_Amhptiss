import tkinter.messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os
from datetime import datetime
from xml.dom import minidom
from Application.AppService.page_element import PageElement

class BaixarDemonstrativoCassi(PageElement):
    usuario_input = (By.XPATH, '//*[@id="cpfOuCnpj"]')
    senha_input = (By.XPATH, '//*[@id="senha"]')
    acessar = (By.XPATH, '//*[@id="loginGeral"]')
    finalizar = (By.XPATH, '//*[@id="step-0"]/nav/button')
    demonstrativo_tiss = (By.XPATH, '/html/body/div[1]/aside/section/div/div/div[1]/div[1]/ul/li[11]/a')
    demonstrativo_de_analises = (By.XPATH, '/html/body/div[1]/aside/section/div/div/div[1]/div[1]/ul/li[11]/ul/li[1]/a')
    data_inicial_input = (By.XPATH, '//*[@id="DataInicial"]')
    data_final_input = (By.XPATH, '//*[@id="DataFinal"]')
    consultar = (By.XPATH, '/html/body/div[1]/div[5]/section/div/form/fieldset/div[4]/div/button')
    xpath_tabela = (By.XPATH, '/html/body/div[1]/div[5]/section/div/fieldset/div/table')
    download_xml = (By.XPATH, '//*[@id="formExportar"]/button[2]')
    download_pdf = (By.XPATH, '//*[@id="formExportar"]/button[1]')
    voltar = (By.XPATH, '//*[@id="btnVoltar"]')
    xpath_corpo_da_pagina = (By.XPATH, '/html/body')

    def __init__(self, url, usuario, senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        self.driver.implicitly_wait(30)
        time.sleep(1.5)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(1.5)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(1.5)
        self.driver.find_element(*self.acessar).click()
        time.sleep(1.5)

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.finalizar).click()
        time.sleep(4)
        self.driver.get('https://servicosonline.cassi.com.br/Prestador/RecursoRevisaoPagamento/TISS/DemonstrativoAnaliseContas/Index') 

    def inicia_automacao(self, **kwargs):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        data_inicial = kwargs.get('data_inicial')
        data_final = kwargs.get('data_final')

        self.driver.find_element(*self.data_inicial_input).send_keys(data_inicial)
        time.sleep(2)
        self.driver.find_element(*self.data_final_input).send_keys(data_final)
        self.driver.find_element(*self.data_inicial_input).send_keys(Keys.ESCAPE)
        time.sleep(2)
        self.driver.find_element(*self.consultar).click()
        time.sleep(2)
        corpo_pagina = self.driver.find_element(*self.xpath_corpo_da_pagina).text

        if "Não foram encontrados resultados para a pesquisa" in corpo_pagina:
            tkinter.messagebox.showinfo( 'Demonstrativo Cassi' , 'Não foram encontrados resultados para a pesquisa' )

        else:
            tabela = self.driver.find_element(*self.xpath_tabela).get_attribute('outerHTML')
            df_tabela = pd.read_html(tabela)[0]
            quantidade_demonstrativos = len(df_tabela) + 1

            for index, linha in df_tabela.iterrows():
                numero_do_protocolo_df = str(linha['Número do Protocolo']).replace(".0","")

                for i in range(1, quantidade_demonstrativos):
                    numero_protocolo_portal = str(self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr[{i}]/td[1]').text).replace(".0", "")

                    if numero_do_protocolo_df != numero_protocolo_portal:
                        continue
                    
                    time.sleep(2)
                    self.driver.find_element(By.XPATH, f'/html/body/div[1]/div[5]/section/div/fieldset/div/table/tbody/tr[{i}]/td[3]/form/input[3]').click()
                    time.sleep(2)
                    self.driver.find_element(*self.download_xml).click()
                    time.sleep(2)
                    count = 0

                    while count < 20:
                        try:
                            with open(f"\\\\10.0.0.239\\automacao_financeiro\\CASSI\\{numero_do_protocolo_df}.xml", "r", encoding="utf-8") as f:
                                valor_total_glosa = 0.00
                                xml = minidom.parse(f)
                                glosa = xml.getElementsByTagName("ans:valorGlosaGeral")

                                for tag in glosa:
                                    valor_total_glosa = tag.firstChild.data
                            break

                        except:
                            count += 1

                            if count == 18:
                                time.sleep(10)

                            else:
                                time.sleep(2)

                    f.close()    
                    
                    if valor_total_glosa != "0":
                        self.driver.find_element(*self.download_pdf).click()
                        data_atual = datetime.now().strftime('%d_%m_%Y')
                        endereco = r"\\10.0.0.239\automacao_financeiro\CASSI"
                        novo_nome = f"{endereco}\\{numero_do_protocolo_df}.pdf"
                        contador = 0

                        while contador < 20:
                            try:
                                os.rename(f"{endereco}\\Demonstrativo de Analise de Conta {data_atual}.pdf", novo_nome)
                                break
                            
                            except:
                                print("Download ainda não foi feito")
                                contador += 1

                                if contador == 18:
                                    time.sleep(10)
                                
                                else:
                                    time.sleep(2)

                    time.sleep(2)
                    self.driver.find_element(*self.voltar).click()
                    time.sleep(2)
                    self.driver.refresh()
                    break

            tkinter.messagebox.showinfo( 'Demonstrativos CASSI' , f"Downloads concluídos!" )
            self.driver.quit()