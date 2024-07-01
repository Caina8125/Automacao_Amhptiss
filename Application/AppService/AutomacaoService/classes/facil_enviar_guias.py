from pandas import DataFrame
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
        token: str = kwargs.get('token')
        df_treeview: DataFrame = kwargs.get('df_treeview')
        self.init_driver()
        self.open()
        self.login()
        self.driver.get(self.url + self.url_envio_xml)
        self.click(self.btn_pesquisar, 2)
        self.click(self.option_100_item, 2)
        self.click(self.i_perfil, 0)
        
        for _ in range(0,self.qtd_paginas):
            element_body_panel = self.driver.find_element(*self.div_body_panel)
            div_faturas_enviadas_list = element_body_panel.find_elements(*self.div_faturas_enviadas)

            for div_fatura_enviada in div_faturas_enviadas_list:
                xml_name = div_fatura_enviada.find_element(*self.div_col_xml).text
                numero_fatura = xml_name.split('_')[0].removeprefix('0000000000000')

                row = df_treeview[df_treeview['Fatura'] == numero_fatura]

                if row.empty:
                    continue

                div_fatura_enviada.find_element(*self.i_clip).click()

                if f"{numero_fatura}.pdf" in self.driver.find_element(*self.modal_envio_arquivo).text:
                    atualizar_status_envio_operadora('normal', numero_fatura, "S", token)
                    self.df_treeview.loc[self.df_treeview['Fatura'] == numero_fatura, 'Status Fatura'] = 'Enviada'
                    continue

                self.send_keys(self.input_type_file, row['Caminho'][0], 2)

                if f"{numero_fatura}.pdf" in self.driver.find_element(*self.modal_envio_arquivo).text:
                    atualizar_status_envio_operadora('normal', numero_fatura, "S", token)
                    self.df_treeview.loc[self.df_treeview['Fatura'] == numero_fatura, 'Status Fatura'] = 'Enviada'

            self.get_element_visible(self.a_proxima_pag_xml)
            self.click(self.i_perfil, 0)
        
        return df_treeview