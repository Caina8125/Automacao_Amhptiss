from datetime import datetime
from tkinter.filedialog import askopenfilenames
import PyPDF2
from pandas import DataFrame

from Infra.Core.core import Core

class GDFConferenciaLotes():

    def inicia_automacao(self, **kwargs):
        files_path_list = kwargs.get('files_path_list')
        self.df_treeview: DataFrame = kwargs.get('df_treeview')
        lines_list = self.ler_pdfs(files_path_list)
        df_lines = self.convert_to_df(lines_list)
        self.comparar_valores(df_lines)

        return self.df_treeview

    def ler_pdfs(self, files_path_list: list[str]):
        pdf_lines_list = []

        for file_path in files_path_list:
            with open(file_path, 'rb') as arquivo_pdf:
                leitor_pdf: PyPDF2.PdfReader = PyPDF2.PdfReader(arquivo_pdf)
                for pagina_numero in range(len(leitor_pdf.pages)):
                    pagina = leitor_pdf.pages[pagina_numero]
                    page_text = pagina.extract_text()
                    page_text = page_text.replace('Prezado Prestador,\n', '')
                    page_text = page_text.replace('PROTOCOLO DE RECEBIMENTO - GDF SAÚDE\n', '')
                    page_text = page_text.replace('Confirmamos recebimento da documentação de cobrança referente \n', '')
                    page_text = page_text.replace('aos LOTE(S) abaixo relacionados.', '')
                    page_text = page_text.replace('DATA\n', '')
                    page_text = page_text.replace('SQ', '')
                    page_text = page_text.replace('LOTE', '')
                    page_text = page_text.replace('TIPO DE GUIA', '')
                    page_text = page_text.replace(' QUANTIDADE DE \n', '')
                    page_text = page_text.replace('GUIAS', '')
                    page_text = page_text.replace('VALOR APRESENTADO', '')
                    page_text = page_text.replace('RECEBIDO POR', '')

                    lines_list = page_text.split('\n')

                    if Core.is_date(lines_list[0], '%d/%m/%Y'):
                        lines_list.pop(0)

                    if lines_list[0] == ' ':
                        lines_list.pop(0)
                    
                    pdf_lines_list.extend(lines_list)

                pdf_lines_list = pdf_lines_list[:-2]
                arquivo_pdf.close()
        return pdf_lines_list
    
    def convert_to_df(self, lines_list: list[str]):
        df = DataFrame([value.replace('  ', ' ').replace('Guia ', '').split(' ') for value in lines_list])
        df = df[[1, 3, 5]]
        try:
            df.columns = ["lote", "quantidade_guias", "valor_total"]
        except ValueError:
            raise ValueError('Erro na leitura do PDF!')
        
        df["valor_total"] = df['valor_total'].apply(Core.formatar)
        
        return df
        
    def comparar_valores(self, df_lines: DataFrame):
        dict_cols = {j: i for i, j in enumerate(self.df_treeview.columns)}

        for row in self.df_treeview.values:
            protocolo = f"{row[dict_cols['Protocolo']]}"
            quantidade = f"{row[dict_cols['Quantidade']]}"
            valor = float(row[dict_cols['Valor Total']])
            status = f"{row[dict_cols['Status']]}"

            if status == 'OK':
                continue

            value = df_lines[df_lines['lote'] == protocolo]

            if value.empty:
                self.df_treeview.loc[self.df_treeview['Protocolo'] == int(protocolo), 'Status'] = "Lote não encontrado"
                continue

            for _, linha in value.iterrows():
                lote = linha['lote']
                valor_lote = float(linha['valor_total'])
                if quantidade != linha['quantidade_guias']:
                    self.df_treeview.loc[self.df_treeview['Protocolo'] == int(lote), 'Status'] = "Quantidade de guias divergente"
                    continue
                elif valor != valor_lote:
                    self.df_treeview.loc[self.df_treeview['Protocolo'] == int(lote), 'Status'] = "Valor divergente"
                    continue
                self.df_treeview.loc[self.df_treeview['Protocolo'] == int(lote), 'Status'] = "OK"