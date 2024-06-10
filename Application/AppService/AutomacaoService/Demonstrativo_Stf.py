import os
import time
import shutil
import tkinter
import pandas as pd
import tkinter.messagebox
from selenium import webdriver
from tkinter import filedialog
from seleniumwire import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Application.AppService.AutomacaoService.Pidgin import financeiroDemo
from Application.AppService.AutomacaoService.page_element import PageElement

class BaixarDemonstrativosStf(PageElement):
    email = (By.XPATH, '//*[@id="Email"]')
    senha = (By.XPATH, '//*[@id="Senha"]')
    logar = (By.XPATH, '//*[@id="btnLogin"]')
    demonstrativo        = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[27]/a/span[1]')
    analise_conta        = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[27]/ul/li[3]/a/span')
    selecionar_convenio  = (By.XPATH, '//*[@id="s2id_OperadorasCredenciadas_HandleOperadoraSelected"]/a/span[2]/b')
    opcao_stf            = (By.XPATH, '/html/body/div[14]/ul/li[7]')
    inserir_protocolo    = (By.XPATH, '//*[@id="Protocolo"]')
    baixar_demonstrativo = (By.XPATH, '//*[@id="btn-Baixar_Demonstrativo"]')
    baixar_xml           = (By.XPATH, '//*[@id="btn-Baixar_XML"]')
    fechar_botao         = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')
    fechar_alerta        = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')
    erro                 = False
    botao_fechar: tuple  = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')
    fechar_lote: tuple   = (By.XPATH, '/html/body/div[3]/div[2]/div/div[2]/div/bc-detalhes-lote/div[3]/div/div[1]/div[2]/a[1]/span')
    proximo_botao: tuple = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')

    def exe_login(self, email, senha):
        # caminho().Alert()
        self.driver.find_element(*self.email).send_keys(email)
        self.driver.find_element(*self.senha).send_keys(senha)
        self.driver.find_element(*self.logar).click()

class BaixarDemonstrativosStf(PageElement):
    

    def exe_caminho(self):
        time.sleep(1)
        self.driver.find_element(*self.demonstrativo).click()
        time.sleep(1)
        self.driver.find_element(*self.analise_conta).click()
        time.sleep(2)
        BaixarDemonstrativosStf(driver, url).Alert()
        self.driver.find_element(*self.selecionar_convenio).click()
        time.sleep(1)
        self.driver.find_element(*self.opcao_stf).click()
        time.sleep(1)

    def buscar_demonstrativo(self):
        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        global count, quantidade_de_faturas, faturas_com_erro
        count = 0
        quantidade_de_faturas = len(df)
        faturas_com_erro = []

        for index, linha in df.iterrows():

            erro = False

            global fatura
            try:
                protocolo =  f"{linha['Nº do Protocolo']}".replace(".0","")
            except:
                protocolo =  f"{linha['Nº do Protocolo']}"
            try:
                fatura =  f"{linha['Nº Fatura']}".replace(".0","")
            except:
                fatura =  f"{linha['Nº Fatura']}"

            self.driver.find_element(*self.inserir_protocolo).send_keys(protocolo)
            time.sleep(1)
            endereco = r"\\10.0.0.239\automacao_financeiro\STF\Renomear"
            arquivo_na_pasta = os.listdir(f"{endereco}")

            for arquivo in arquivo_na_pasta:
                if '.pdf' in arquivo:
                    endereco_arquivo = f'{endereco}\\{arquivo}'
                    shutil.move(endereco_arquivo, r"\\10.0.0.239\automacao_financeiro\STF\Não Renomeados")

            self.driver.find_element(*self.baixar_demonstrativo).click()
            time.sleep(6)
            self.driver.find_element(*self.baixar_xml).click()
            time.sleep(8)

            for i in range(10):
                pasta = r"\\10.0.0.239\automacao_financeiro\STF\Renomear"
                nomes_arquivos = os.listdir(pasta)
                if len(nomes_arquivos) == 0:
                    break
                # time.sleep(2)
                
                for nome in nomes_arquivos:
                    if '.pdf' in nome:
                        nomepdf  = os.path.join(pasta, nome)
                        renomear = r"\\10.0.0.239\automacao_financeiro\STF\Renomear" +f"\\{fatura}"  +  ".pdf"
                        arqDest  = r"\\10.0.0.239\automacao_financeiro\STF" + f"\\{fatura}"  +  ".pdf"
                        
                        try:
                            os.rename(nomepdf,renomear)
                            shutil.move(renomear,arqDest)
                            time.sleep(2)
                            print("Arquivo renomeado e guardado com sucesso")
                            break

                        except Exception as e:
                            print(e)
                            print("Download ainda não foi feito/Arquivo não renomeado")
                            time.sleep(2)
                    else:
                        arqDest_xml = r"\\10.0.0.239\automacao_financeiro\STF" + f"\\{nome}"
                        nomexml     = os.path.join(pasta, nome)
                        try:
                            shutil.move(nomexml,arqDest_xml)
                            time.sleep(2)
                            print("Arquivo renomeado e guardado com sucesso")
                            break

                        except Exception as e:
                            print(e)
                            print("Download ainda não foi feito/Arquivo não renomeado")
                            time.sleep(2)

                if i == 9:
                    faturas_com_erro.append(fatura)
                    erro = True

            time.sleep(2)
            self.driver.find_element(*self.inserir_protocolo).clear()
            time.sleep(1)

            if erro == False:
                count += 1

        if count == quantidade_de_faturas:
            tkinter.messagebox.showinfo( 'Demonstrativos Câmara' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

        else:
            tkinter.messagebox.showinfo( 'Demonstrativos Câmara' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(faturas_com_erro) }." )            
        
    def Alert(self):
        self.driver.implicitly_wait(3)
        while True:
            try:
                self.driver.find_element(*self.proximo_botao).click()
                time.sleep(0.5)
            except:
                break
        try:
            self.driver.find_element(*self.botao_fechar).click()
        except:
            return