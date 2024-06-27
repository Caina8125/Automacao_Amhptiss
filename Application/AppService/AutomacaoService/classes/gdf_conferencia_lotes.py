from datetime import datetime
from tkinter.filedialog import askopenfilenames
import PyPDF2
from pandas import DataFrame

# from Infra.Core.core import Core


class Core:
        
    @staticmethod    
    def is_date(string, date_format='%Y-%m-%d'):
        try:
            datetime.strptime(string, date_format)
            return True
        except ValueError:
            return False

class GDFConferenciaLotes():

    def inicia_automacao(self, **kwargs):
        files_path_list = kwargs.get('files_path_list')
        lines_list = self.ler_pdfs(files_path_list)
        df_lines = self.convert_to_df(lines_list)

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
        df = DataFrame([value.replace('Guia ', '').split(' ') for value in lines_list])
        df = df[[1, 3, 5]]
        try:
            df.columns = ["lote", "quantidade_guias", "valor_total"]
        except ValueError:
            raise ValueError('Erro na leitura do PDF!')
        print(df)


if __name__ == "__main__":
    files = askopenfilenames()
    GDFConferenciaLotes().inicia_automacao(files_path_list=files)