import os
import time
import shutil
import tkinter
import pandas as pd
import tkinter.messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Application.AppService.AutomacaoService.page_element import PageElement


class DemonstrativoCamed(PageElement):
    email = (By.XPATH, '//*[@id="Email"]')
    senha_input = (By.XPATH, '//*[@id="Senha"]')
    logar = (By.XPATH, '//*[@id="btnLogin"]')
    demonstrativo        = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[27]/a/span[1]')
    analise_conta        = (By.XPATH, '/html/body/div[3]/div[1]/div/ul/li[27]/ul/li[3]/a/span')
    selecionar_convenio  = (By.XPATH, '//*[@id="s2id_OperadorasCredenciadas_HandleOperadoraSelected"]/a/span[2]/b')
    opcao_camed         = (By.XPATH, '/html/body/div[14]/ul/li[3]/div')
    inserir_protocolo    = (By.XPATH, '//*[@id="Protocolo"]')
    baixar_demonstrativo = (By.XPATH, '//*[@id="btn-Baixar_Demonstrativo"]')
    baixar_xml           = (By.XPATH, '//*[@id="btn-Baixar_XML"]')
    fechar_botao         = (By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')
    fechar_alerta        = (By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')

    def __init__(self, url, usuario, senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        # caminho().Alert()
        self.driver.find_element(*self.email).send_keys(self.usuario)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        self.driver.find_element(*self.logar).click()

    def exe_caminho(self):
        time.sleep(1)
        self.driver.find_element(*self.demonstrativo).click()
        time.sleep(1)
        self.driver.find_element(*self.analise_conta).click()
        time.sleep(2)
        self.Alert()
        self.driver.find_element(*self.selecionar_convenio).click()
        time.sleep(1)
        self.driver.find_element(*self.opcao_camed).click()
        time.sleep(1)

    def inicia_automacao(self, **kwargs):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        planilha = kwargs.get('arquivo')
        
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
            endereco = r"\\10.0.0.239\automacao_financeiro\CAMED\Renomear"
            arquivo_na_pasta = os.listdir(f"{endereco}")

            for arquivo in arquivo_na_pasta:
                if '.pdf' in arquivo:
                    endereco_arquivo = f'{endereco}\\{arquivo}'
                    shutil.move(endereco_arquivo, r"\\10.0.0.239\automacao_financeiro\CAMED\Não Renomeados")

            self.driver.find_element(*self.baixar_demonstrativo).click()
            time.sleep(6)
            self.driver.find_element(*self.baixar_xml).click()
            time.sleep(8)

            for i in range(10):
                pasta = r"\\10.0.0.239\automacao_financeiro\CAMED\Renomear"
                nomes_arquivos = os.listdir(pasta)
                if len(nomes_arquivos) == 0:
                    break
                # time.sleep(2)
                
                for nome in nomes_arquivos:
                    if '.pdf' in nome:
                        nomepdf  = os.path.join(pasta, nome)
                        renomear = r"\\10.0.0.239\automacao_financeiro\CAMED\Renomear" +f"\\{fatura}"  +  ".pdf"
                        arqDest  = r"\\10.0.0.239\automacao_financeiro\CAMED" + f"\\{fatura}"  +  ".pdf"
                        
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
                        arqDest_xml = r"\\10.0.0.239\automacao_financeiro\CAMED" + f"\\{nome}"
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
        try:
            modal = WebDriverWait(self.driver, 3.0).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div')))
            self.driver.find_element(*self.fechar_botao).click()

            while True:
                try:
                    proximo_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '//*[@id="bcInformativosModal"]/div/div/div[3]/button[2]')))
                    proximo_botao.click()
                except:
                    break

            try:
                fechar_botao = WebDriverWait(self.driver, 0.2).until(EC.presence_of_element_located((By.XPATH, '/html/body/bc-modal-evolution/div/div/div/div[3]/button[3]')))
                fechar_botao.click()
            except:
                print("Não foi possível encontrar o botão de fechar.")
                pass

        except:
            print("Não teve Modal")
            pass