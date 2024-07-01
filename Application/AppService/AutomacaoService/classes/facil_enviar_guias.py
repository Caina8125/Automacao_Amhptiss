from time import sleep
from tkinter.messagebox import showerror
from pandas import DataFrame
from selenium.webdriver.common.keys import Keys

from Application.AppService.AutomacaoService.classes.WebSitesContext.facil_context import FacilContext
from Application.AppService.EnvioOperadoraService import atualizar_status_envio_operadora

class FacilEnviarGuias(FacilContext):

    def __init__(self, url, codigo_convenio, usuario, senha) -> None:
        super().__init__(url, codigo_convenio)
        self.usuario = usuario
        self.senha = senha

    def login(self):
        self.click(self.option_tipo_acesso, 1)
        self.send_keys(self.input_usuario, self.usuario, 1)
        self.send_keys(self.input_senha, self.senha, 1)
        self.click(self.btn_entrar, 6)  

    def inicia_automacao(self, **kwargs):
        try:
            token: str = kwargs.get('token')
            self.df_treeview: DataFrame = kwargs.get('df_treeview')
            self.init_driver(py_auto_gui=True)
            self.open()
            self.login()
            self.click(self.btn_ignorar, 1)
            self.driver.get(self.url + self.url_envio_xml)
            sleep(2)
            self.click(self.btn_pesquisar, 2)
            self.click(self.option_100_item, 2)
            self.driver.find_element(*self.clip1).send_keys(Keys.CONTROL + Keys.HOME)
            self.click(self.i_perfil, 2)
            
            for _ in range(0,self.qtd_paginas):
                element_body_panel = self.driver.find_element(*self.div_body_panel)
                div_faturas_enviadas_list = element_body_panel.find_elements(*self.div_faturas_enviadas)

                for div_fatura_enviada in div_faturas_enviadas_list:
                    xml_name = div_fatura_enviada.find_element(*self.div_col_xml).text
                    numero_fatura = xml_name.split('_')[0].removeprefix('0000000000000')

                    row = self.df_treeview[self.df_treeview['Fatura'] == int(numero_fatura)]

                    if row.empty:
                        continue

                    self.get_element_visible(web_element=div_fatura_enviada.find_element(*self.i_clip))
                    self.driver.find_element(*self.adicionar_arquivo)
                    sleep(1)

                    if ".pdf" in self.driver.find_element(*self.modal_envio_arquivo).text:
                        atualizar_status_envio_operadora('normal', numero_fatura, "S", token)
                        self.df_treeview.loc[self.df_treeview['Fatura'] == int(numero_fatura), 'Status Fatura'] = 'Enviada'
                        self.click(self.fechar_modal, 1)
                        continue

                    self.send_keys(self.input_type_file, row['Caminho'][0], 2)
                    self.click(self.adicionar_arquivo, 2)
                    try:
                        nome_arquivo = self.driver.find_element(*self.span_arquivo).text
                    except:
                        nome_arquivo = None

                    if nome_arquivo != None:
                        atualizar_status_envio_operadora('normal', numero_fatura, "S", token)
                        self.df_treeview.loc[self.df_treeview['Fatura'] == int(numero_fatura), 'Status Fatura'] = 'Enviada'

                    self.click(self.fechar_modal, 1)

                self.get_element_visible(self.a_proxima_pag_xml)
                self.click(self.i_perfil, 0)
                self.driver.find_element(*self.clip1).send_keys(Keys.CONTROL + Keys.HOME)
                sleep(2)
            
        except Exception as e:
            showerror('', f"Ocorreu uma exceção não tratada.\n{e.__class__.__name__}:\n{e}")
        
        try:
            self.driver.quit()
        except:
            pass

        return self.df_treeview