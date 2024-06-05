import pandas as pd
from selenium.webdriver.common.by import By
import time
import os
import tkinter
import shutil
from Application.AppService.page_element import PageElement

class BaixarDemonstrativoCasembrapa(PageElement):
    usuario_input = (By.XPATH, '//*[@id="username"]')
    senha_input = (By.XPATH, '//*[@id="password"]')
    entrar = (By.XPATH, '//*[@id="submit-login"]')
    salutis = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    websaude = (By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[2]')
    credenciados = (By.XPATH, '//*[@id="divTreeNavegation"]/div[8]/span[2]')
    demonstrativos = (By.XPATH, '//*[@id="divTreeNavegation"]/div[12]/span[2]')
    demonstrativos_de_analise = (By.XPATH, '//*[@id="divTreeNavegation"]/div[13]/span[2]')
    numero_dos_lotes_prestador = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/textarea')
    numero_dos_protocolos = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/textarea')
    salutis = (By.XPATH, '//*[@id="menuButtons"]/td[1]')
    numero_dos_lotes_prestador = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/textarea')
    numero_dos_lotes_operadora1 = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/textarea')
    numero_dos_lotes_operadora2 = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/textarea')
    numero_dos_protocolos = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/textarea')
    numero_dos_protocolos2 = (By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[9]/td[2]/table/tbody/tr/td[1]/textarea')
    executar = (By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]')
    imprimir = (By.XPATH, '//*[@id="bt_1892814041"]/table/tbody/tr/td[2]')
    imprimir_em_texto = (By.XPATH, '//*[@id="buttonsContainer_1"]/td[4]/span[2]')
    botao_voltar = (By.XPATH, '//*[@id="buttonsCell"]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')
    erro = (By.XPATH, '/html/body/div[10]')
    erro_ok = (By.XPATH, '//*[@id="confirm"]')

    def __init__(self, url, usuario, senha) -> None:
        super().__init__(url)
        self.usuario = usuario
        self.senha = senha

    def exe_login(self):
        self.driver.implicitly_wait(30)
        self.driver.find_element(*self.usuario_input).send_keys(self.usuario)
        time.sleep(2)
        self.driver.find_element(*self.senha_input).send_keys(self.senha)
        time.sleep(2)
        self.driver.find_element(*self.entrar).click()
        time.sleep(2)

    def exe_caminho(self):
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.salutis).click()
        time.sleep(2)
        self.driver.find_element(*self.websaude).click()
        time.sleep(2)
        self.driver.find_element(*self.credenciados).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativos).click()
        time.sleep(2)
        self.driver.find_element(*self.demonstrativos_de_analise).click()

    def refazer_caminho(self):
        self.driver.get(self.url)
        self.exe_login()
        self.driver.implicitly_wait(30)
        time.sleep(2)
        self.driver.find_element(*self.salutis).click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, '/html/body/div[8]/div[2]/div[13]/span[2]').click()

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
        contador_vezes = 1
        count = 0
        quantidade_de_faturas = len(df)
        lista_diretorio = os.listdir(r"\\10.0.0.239\automacao_financeiro\CASEMBRAPA")
        lista_de_nomes_sem_extensao = [
            nome.replace('.pdf', '') for nome in lista_diretorio
            ]
        erro_portal = False
        lista_faturas_com_erro = []

        for tentativa in range(1, 6):
            print(f'Tentativa {tentativa}')

            try:
                for index, linha in df.iterrows():
                    numero_fatura = str(linha["Nº Fatura"]).replace('.0', '')
                    numero_protocolo = str(linha["Nº do Protocolo"]).replace('.0', '')

                    if df['Concluído'][index] == "Sim":
                        continue

                    if numero_fatura in lista_de_nomes_sem_extensao:
                        count += 1
                        continue

                    print(numero_fatura)
                    self.driver.implicitly_wait(30)
                    time.sleep(2)
                    self.driver.switch_to.frame('inlineFrameTabId1')
                    time.sleep(1)
                    self.driver.find_element(*self.numero_dos_lotes_prestador).clear()
                    time.sleep(1)

                    if contador_vezes == 1:
                        self.driver.find_element(*self.numero_dos_lotes_operadora1).clear()
                        time.sleep(0.5)
                        self.driver.find_element(*self.numero_dos_protocolos).clear()
                        time.sleep(1)
                        self.driver.find_element(*self.numero_dos_protocolos).click()
                        time.sleep(1)
                        self.driver.find_element(*self.numero_dos_protocolos).send_keys(numero_protocolo)

                    else:
                        self.driver.find_element(*self.numero_dos_lotes_operadora2).clear()
                        time.sleep(0.5)
                        self.driver.find_element(*self.numero_dos_protocolos2).clear()
                        time.sleep(1)
                        self.driver.find_element(*self.numero_dos_protocolos2).click()
                        time.sleep(1)
                        self.driver.find_element(*self.numero_dos_protocolos2).send_keys(numero_protocolo)

                    time.sleep(2)
                    self.driver.switch_to.default_content()
                    time.sleep(1)

                    for i in range(1, 6):
                        try:
                            self.driver.find_element(*self.executar).click()
                            time.sleep(0.5)
                        except:
                            break

                    fatura_nao_encontrada = False

                    while True:

                        try:
                            self.driver.implicitly_wait(5)
                            self.driver.find_element(*self.imprimir_em_texto)
                            break

                        except:
                            try:
                                self.driver.switch_to.default_content()
                                self.driver.find_element(*self.erro)
                                time.sleep(2)
                                self.driver.find_element(*self.erro_ok).click()
                                df.loc[index, 'Concluído'] = 'Sim'
                                lista_faturas_com_erro.append(numero_fatura)
                                fatura_nao_encontrada = True
                                break
                            except:
                                pass
                        
                    if fatura_nao_encontrada == True:
                        continue
        
                    time.sleep(2)
                    endereco = r"\\10.0.0.239\automacao_financeiro\CASEMBRAPA\Renomear"
                    arquivo_na_pasta = os.listdir(f"{endereco}")

                    for arquivo in arquivo_na_pasta:
                        endereco_arquivo = f'{endereco}\\{arquivo}'
                        shutil.move(endereco_arquivo, r"\\10.0.0.239\automacao_financeiro\CASEMBRAPA\Não Renomeados")

                    self.driver.find_element(*self.imprimir).click()
                    time.sleep(2)

                    arquivo_renomeado = False
                    
                    for i in range(10):
                        novo_nome = f"{endereco}\\{numero_fatura}.pdf"
                        arquivo_na_pasta = os.listdir(f"{endereco}")
                        pasta_nova = f"\\\\10.0.0.239\\automacao_financeiro\\CASEMBRAPA\\{numero_fatura}.pdf"

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
                            if i == 9:
                                erro_portal = True
                                self.driver.quit()

                    if arquivo_renomeado == True:
                        count += 1
                        print(f"Download da fatura {numero_fatura} concluído com sucesso")
                    
                    else:
                        print(f"Download da fatura {numero_fatura} não foi feito ou o arquivo não foi renomeado.")
                        lista_faturas_com_erro.append(numero_fatura)

                    df.loc[index, 'Concluído'] = 'Sim'
                    print("-------------------------------------------------------------------------------")

                    self.driver.find_element(*self.botao_voltar).click()
                    contador_vezes += 1

                if count == quantidade_de_faturas:
                    tkinter.messagebox.showinfo( 'Demonstrativos Casembrapa' , f"Downloads concluídos: {count} de {quantidade_de_faturas}." )

                else:
                    tkinter.messagebox.showinfo( 'Demonstrativos Casembrapa' , f"Downloads concluídos: {count} de {quantidade_de_faturas}. Conferir fatura(s): {', '.join(lista_faturas_com_erro) }." )

                break

            except Exception as error:
                print(f"{error.__class__.__name__}: {error}")                  

                if erro_portal == True:
                    tkinter.messagebox.showerror("Automação", f"Ocorreu uma exceção não tratada. \n {error.__class__.__name__} - {error}")
                    break
                
                self.exe_caminho()
                contador_vezes = 1