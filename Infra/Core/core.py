import os
import shutil

class Core:

    @staticmethod
    def obter_caminhos_nf(remessa: str, nota_fiscal:str):
        """
        Obtém os caminhos completos dos arquivos PDF e XML da nota fiscal, se existirem.

        Args:
            remessa (str): O nome do diretório da remessa.
            nota_fiscal (str): O nome da nota fiscal (sem extensão).

        Returns:
            tuple: Caminhos completos dos arquivos PDF e XML da nota fiscal, se encontrados.
                   (pdf_path, xml_path)
            tuple: (None, None) se os arquivos ou diretório não forem encontrados.
        """
        base_dir = r"\\10.0.0.239\financeiro - faturamento\IMPRESSÃO DE NFE\SIS"
        remessa_dir = os.path.join(base_dir, remessa)

        if not os.path.isdir(remessa_dir):
            return None, None

        try:
            files_list = os.listdir(remessa_dir)
        except OSError as e:
            return None, None

        pdf_filename = next((file for file in files_list if file.endswith(nota_fiscal + '-nfse.pdf')), None)
        xml_filename = next((file for file in files_list if file.endswith(nota_fiscal + '-nfse.xml')), None)

        pdf_path = os.path.join(remessa_dir, pdf_filename) if pdf_filename else None
        xml_path = os.path.join(remessa_dir, xml_filename) if xml_filename else None

        return (pdf_path if pdf_path and os.path.isfile(pdf_path) else None,
                xml_path if xml_path and os.path.isfile(xml_path) else None)
    
    @staticmethod
    def criar_pasta_e_armazenar_arquivos(destino_dir: str, arquivos: list[str]):
        """
        Cria um diretório e armazena os arquivos especificados nele.

        Args:
            destino_dir (str): O caminho do diretório onde os arquivos serão armazenados.
            arquivos (list): Lista de caminhos de arquivos a serem copiados.

        Returns:
            bool: True se os arquivos foram copiados com sucesso, False caso contrário.
        """
        try:
            os.makedirs(destino_dir, exist_ok=True)
            for arquivo in arquivos:
                if arquivo and os.path.isfile(arquivo):
                    shutil.copy2(arquivo, destino_dir)
            return True
        
        except Exception as e:
            return False