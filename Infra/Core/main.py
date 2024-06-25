from datetime import datetime
from tkinter.messagebox import showerror, showinfo
from core import Core

path_arquivo = r"\\10.0.0.239\guiasscaneadas\2024\SIS\2309815.pdf"
pdf_path, xml_path = Core.obter_caminhos_nf('48915', '0035148')

paths_list = [path_arquivo, pdf_path, xml_path]

data_atual = datetime.now()
new_dir_name = "output/SIS_" + data_atual.strftime('%d_%m_%Y_%H_%M_%S')
if Core.criar_pasta_e_armazenar_arquivos(new_dir_name, paths_list):

    showinfo('', 'Ok')

else:
    showerror('' ,"Não deu certo")