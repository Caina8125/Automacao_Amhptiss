from datetime import date, datetime
from tkinter.messagebox import showerror, showwarning

from pandas import DataFrame
from Application.AppService.AutomacaoService.page_element import PageElement
from selenium.common.exceptions import InvalidArgumentException
from selenium.webdriver.common.by import By
from time import sleep

from Application.AppService.EnvioOperadoraService import atualizar_status_envio_operadora

class Geap(PageElement):
    acessar_portal = By.XPATH, '/html/body/div[3]/div[3]/div[1]/form/div[1]/div[1]/div/a'
    usuario = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[1]/div/div[1]/div/input'
    input_senha = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[1]/div/label[2]/div/div[1]/div[1]/input'
    entrar = By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div[2]/button/div[2]/div/div'
    fechar = By.XPATH, '/html/body/div[4]/div[2]/div/div[3]/button'
    portal_tiss = (By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[1]/div[2]/div/div/div')
    alerta = By.XPATH,' /html/body/div[2]/div/center/a'
    a_tiss = By.XPATH, '/html/body/div[1]/nav/ul/li[3]/a'
    anexo_conta = By.XPATH, '/html/body/div[1]/nav/ul/li[3]/ul/li[7]/a'
    select_mes_referencia = By.ID, 'dropReferencia'
    select_nr = By.ID, 'dropNr'
    select_conta = By.ID, 'dropConta'
    select_tipo_anexo = By.ID, 'dropTipoAnexo'
    text_area_decricao = By.ID, 'txtAnexoDesc'
    btn_processar = By.ID, 'btnFinalizar'
    input_file = By.XPATH, '/html/body/input[2]'
    btn_remover_anexos = By.ID, 'clear-dropzone'
    div_msg = By.ID, 'toast-container'

    def __init__(self, url: str, cpf: str, senha: str, tipo_negociacao: str) -> None:
        super().__init__(url)
        self.cpf = cpf
        self.senha = senha
        self.tipo_negociacao = tipo_negociacao

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
            sleep(2)
            self.driver.find_element(*self.portal_tiss)

        except Exception as e:
            self.driver.implicitly_wait(180)
            self.driver.find_element(*self.portal_tiss)
            sleep(2)
            self.driver.implicitly_wait(15)


    def exe_caminho(self):
        self.driver.find_element(*self.portal_tiss)
        sleep(3)
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
        self.driver.implicitly_wait(180)
        # sleep(2)
        # self.driver.get('https://geap.topsaude.com.br/PortalCredenciado/HomePortalCredenciado/Home/AreaLogada#PORCRED50')
        # sleep(3)
        # self.driver.find_element(*self.a_tiss).click()
        # sleep(2)
        self.driver.find_element(*self.anexo_conta).click()
        self.driver.implicitly_wait(15)
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
        df_treeview = kwargs.get('df_treeview')
        lista_relatorio_guia = []
        data_atual = datetime.now()
        try:
            self.init_driver(py_auto_gui=True)
            self.open()
            self.exe_login()
            self.exe_caminho()

            token = kwargs.get('token')
            dados_remessa = kwargs.get('dados_faturas')
            sleep(1.5)

        except Exception as e:
            showerror('', f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__}:\n{e}")
            self.gerar_relatorio(lista_relatorio_guia, data_atual)
            return df_treeview

        element_select_mes = self.driver.find_element(*self.select_mes_referencia)
        mes_options_list = [option for option in element_select_mes.find_elements(By.TAG_NAME, 'option') if 'Selecione a Referência' not in option.text]

        for option in mes_options_list:
            try:
                option.click()
                sleep(1)

                for fatura in dados_remessa['faturas']:
                    n_processo = int(fatura['fatura'])
                    protocolo = fatura['protocolo']
                    guias = fatura['lista_guias']
                    option.click()

                    if not self.click_option(self.select_nr, protocolo):
                        continue
                    
                    for i, guia_data in enumerate(guias):
                        option.click()
                        self.click_option(self.select_nr, protocolo)
                        guia = f"{guia_data['guia']}".replace('.0', '')
                        caminho = guia_data['caminho_guia']
                        sleep(1)
                        self.click_option(self.select_conta, guia)

                        if not self.click_option(self.select_conta, guia):
                            continue
                        
                        sleep(1)

                        self.click_option(self.select_tipo_anexo, 'RELATÓRIO DE AUDITORIA')
                        sleep(1)

                        self.driver.find_element(*self.text_area_decricao).send_keys("Guia de faturamento.")
                        sleep(1)
                        try:
                            self.driver.find_element(*self.input_file).send_keys(caminho)
                        except InvalidArgumentException as e:
                            self.driver.find_element(*self.text_area_decricao).clear()
                            sleep(2)
                            continue
                        sleep(2)
                        self.driver.find_element(*self.btn_processar).click()
                        sleep(4)
                        msg = self.driver.find_element(*self.div_msg).text

                        if 'Ocorreu um erro ao salvar os dados' in msg:
                            self.driver.find_element(*self.text_area_decricao).clear()
                            sleep(1)
                            self.driver.find_element(*self.btn_remover_anexos).click()
                            sleep(1)
                            guias[i]['guia_enviada'] = False
                            lista_relatorio_guia.append([n_processo, protocolo, guia, msg.replace('x\n', '').replace('\n','. ')])
                            continue

                        
                        guias[i]['guia_enviada'] = True
                        lista_relatorio_guia.append([n_processo, protocolo, guia, msg.replace('x\n', '').replace('\n','. ')])

                    bool_list = [value['guia_enviada'] for value in guias]

                    if False not in bool_list:
                        atualizar_status_envio_operadora(self.tipo_negociacao, n_processo, "S", token)
                        df_treeview.loc[df_treeview['Fatura'] == n_processo, 'Status Fatura'] = 'Enviada'
                    elif True not in bool_list:
                        atualizar_status_envio_operadora(self.tipo_negociacao, n_processo, "N", token)
                        df_treeview.loc[df_treeview['Fatura'] == n_processo, 'Status Fatura'] = 'Não Enviada'
                    elif True in bool_list and False in bool_list:
                        atualizar_status_envio_operadora(self.tipo_negociacao, n_processo, "P", token)
                        df_treeview.loc[df_treeview['Fatura'] == n_processo, 'Status Fatura'] = 'Enviada Parcialmente'

            except Exception as e:
                showerror('', f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__}:\n{e}")
                break
            
        self.gerar_relatorio(lista_relatorio_guia, data_atual)
        self.driver.quit()
        return df_treeview
    
    def gerar_relatorio(self, lista_relatorio, data):
        df_relatorio = DataFrame(lista_relatorio)
        if df_relatorio.empty:
            showwarning('', 'Nenhum relatório a ser gerado.')
            return
        df_relatorio.columns = ['Fatura', 'Protocolo', 'Guia', 'Log Portal']
        df_relatorio.to_excel('Output\\geap_' + data.strftime('%d_%m_%Y_%H_%M_%S') + '.xlsx', index=False)