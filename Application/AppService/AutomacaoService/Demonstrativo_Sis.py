import tkinter.messagebox
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
import shutil
from Application.AppService.AutomacaoService.page_element import PageElement

class BaixarDemonstrativosSis(PageElement):
    usuario_input = (By.XPATH, '//*[@id="UserName"]')
    senha_input = (By.XPATH, '//*[@id="Password"]')
    acessar = (By.XPATH, '//*[@id="LoginButton"]')
    demonstrativos = (By.XPATH, '//*[@id="sidebar_demonstrativos"]')
    demonstrativo_de_glosa = (By.XPATH, '//*[@id="ctl00_SidebarMenu"]/li[4]/ul/li[3]/a')
    numero_do_protocolo = (By.XPATH, '//*[@id="ctl00_Main_DefaultDataInputForm_PageControl_GERAL_GERAL_NUMERODOPROTOCOLO"]')
    botao_ok = (By.XPATH, '//*[@id="ctl00_Main_DefaultDataInputForm_toolbar"]/a[1]')
    imprimir = (By.XPATH, '//*[@id="Print"]/table/tbody/tr/td[2]')
    imprimir_pdf = (By.XPATH, '/html/body/form/div[3]/div[3]/div[3]/div/span/div/div[12]/div/div[1]/table/tbody/tr/td[2]')
    botao_demonstrativo_de_glosa = (By.XPATH, '//*[@id="top-CMD_DemonstrativodeGlosa"]')

    def __init__(self, url, usuario,senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        self.driver.find_element(*self.acessar).click()

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.demonstrativos).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativo_de_glosa).click()
        time.sleep(2)

    def inicia_automacao(self, **kwargs):
        self.init_driver()
        self.open()
        self.exe_login()
        self.exe_caminho()

        planilha = kwargs.get('arquivo')

        df = pd.read_excel(planilha, header=5)
        df = df.iloc[:-1]
        df = df.dropna()
        df['Concluído'] = ''
        count = 0
        quantidade_de_faturas = len(df)
        lista_diretorio = os.listdir(r"\\10.0.0.239\automacao_financeiro\SIS")
        lista_de_nomes_sem_extensao = [nome.replace('.pdf', '') for nome in lista_diretorio]
        lista_faturas_com_erro = []

        for tentativa in range(1, 6):
            erro_portal = False
            print(f'Tentativa {tentativa}')

            try:

                for index, linha in df.iterrows():
                    numero_fatura = str(linha['Nº Fatura']).replace('.0', '')
                    numero_protocolo_planilha = str(linha['Nº do Protocolo']).replace('.0', '')

                    if df['Concluído'][index] == "Sim":
                        continue

                    if numero_fatura in lista_de_nomes_sem_extensao:
                        count += 1
                        continue

                    self.driver.implicitly_wait(30)
                    self.driver.find_element(*self.botao_demonstrativo_de_glosa).click()
                    time.sleep(2)
                    self.driver.find_element(*self.numero_do_protocolo).send_keys(numero_protocolo_planilha)
                    time.sleep(2)
                    self.driver.find_element(*self.botao_ok).click()
                    time.sleep(2)
                    self.driver.find_element(*self.imprimir).click()
                    time.sleep(2)
                    endereco = r"\\10.0.0.239\automacao_financeiro\SIS\Renomear"
                    arquivo_na_pasta = os.listdir(f"{endereco}")

                    for arquivo in arquivo_na_pasta:
                        endereco_arquivo = f'{endereco}\\{arquivo}'
                        shutil.move(endereco_arquivo, r"\\10.0.0.239\automacao_financeiro\SIS\Não Renomeados")

                    self.driver.find_element(*self.imprimir_pdf).click()
                    time.sleep(4)
                    arquivo_renomeado = False
                    endereco = r"\\10.0.0.239\automacao_financeiro\SIS\Renomear"
                    
                    for i in range(10):
                        novo_nome = f"{endereco}\\{numero_fatura}.pdf"
                        arquivo_na_pasta = os.listdir(f"{endereco}")
                        pasta_nova = f"\\\\10.0.0.239\\automacao_financeiro\\SIS\\{numero_fatura}.pdf"

                        for arquivo in arquivo_na_pasta:
                            nome_antigo = f"{endereco}\\{arquivo}"

                        try:
                            os.rename(nome_antigo, novo_nome)
                            print("Renomeado")
                            shutil.move(novo_nome, pasta_nova)
                            print("Arquivo na pasta")
                            arquivo_renomeado = True
                            break
                        
                        except Exception as e:
                            print(e)
                            print("Download ainda não foi feito/Arquivo não renomeado")
                            time.sleep(2)
                            
                            try:
                                self.driver.implicitly_wait(5)
                                mensagem = self.driver.find_element(By.XPATH, '//*[@id="ctl00_Main_DEMONSTRATIVODEANLISEDECONTA_MsgUser_message"]').text

                                if mensagem == 'Código do protocolo não encontrado' or mensagem == 'Identifica¿¿¿¿o do benefici¿¿rio n¿¿o consistente':
                                    print(mensagem)
                                    lista_faturas_com_erro.append(numero_fatura)
                                    break



                                else:
                                    continue
                        
                            except:

                                if i == 9:
                                    erro_portal = True
                                    self.driver.quit()

                    if arquivo_renomeado == True:
                        count += 1
                        print(f"Download da fatura {numero_fatura} concluído com sucesso")
                    
                    else:

                        if numero_fatura in lista_faturas_com_erro:
                            print(f"Download da fatura {numero_fatura} não foi feito ou o arquivo não foi renomeado.")
                        
                        else:
                            print(f"Download da fatura {numero_fatura} não foi feito ou o arquivo não foi renomeado.")
                            lista_faturas_com_erro.append(numero_fatura)


                    df.loc[index, 'Concluído'] = 'Sim'

                    self.driver.implicitly_wait(5)
                    self.driver.get("https://intra4p.senado.leg.br/WebAppPortal/Prestador/Demonstrativo/DemonstrativoGlosa.aspx?i=K_DEMONSTRATIVODEGLOSA&m=MENU_DEMONSTRATIVOS_AGRUPADO")
                    time.sleep(2)

                    print("-------------------------------------------------------------------------------")

                if count == quantidade_de_faturas:
                    tkinter.messagebox.showinfo( 'Demonstrativos SIS' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

                else:
                    tkinter.messagebox.showinfo( 'Demonstrativos SIS' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(lista_faturas_com_erro) }." )

                self.driver.quit()
                
                break

            except Exception as error:
                print(f"{error.__class__.__name__}: {error}")

                if erro_portal == True:
                    print("Portal sem resposta, tente novamente mais tarde")
                    break
                
                self.driver.get("https://intra4p.senado.leg.br/WebAppPortal/Prestador/Demonstrativo/DemonstrativoGlosa.aspx?i=K_DEMONSTRATIVODEGLOSA&m=MENU_DEMONSTRATIVOS_AGRUPADO")