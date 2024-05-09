import configparser
import shutil

class Atualiza:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read('Config.ini')

        self.pathLocal = self.ObterPathLocal()
        self.pathAtualiza = self.ObterPathAtualiza()

        self.configAtualiza = configparser.ConfigParser()
        self.configAtualiza.read(self.pathAtualiza)

        self.CompararVersoes()

    def ObterPathAtualiza(self):
        self.pastaAtualiza = self.config['ConfigPath']['PathAtualizaIni']
        return self.pastaAtualiza
     
    def ObterPathLocal(self):
        self.pastaLocal = self.config['ConfigPath']['PathLocalIni']
        return self.pastaLocal
    
    def atualizar_pasta(self,origem, destino):
        # Copia a árvore de diretórios da pasta de origem para a pasta de destino
        shutil.rmtree(destino)  # Remove a pasta de destino se ela já existir
        shutil.copytree(origem, destino)

    def CompararVersoes(self):
        # Obter valores de seções
        self.VersãoAtualiza = self.configAtualiza['ConfigVension']['version']
        self.VersãoLocal = self.config['ConfigVension']['version']

        if self.VersãoAtualiza != self.VersãoLocal:
            self.origem = self.config['Folders']['PathAtualiza']
            self.destino = self.config['Folders']['PathLocal']
            self.atualizar_pasta(self.origem, self.destino)

Atualiza()





















        # # Modificar valores
        # config['Seção1']['Chave1'] = 'Novo valor'

        # # Adicionar uma nova seção ou chave
        # if 'NovaSeção' not in config:
        #     config.add_section('NovaSeção')
        # config['NovaSeção']['NovaChave'] = 'Valor'

        # # Escrever de volta ao arquivo .ini
        # with open('arquivo.ini', 'w') as configfile:
        #     config.write(configfile)
