import os
import shutil
import configparser
import time
import semantic_version

class PublishExe:
    def __init__(self):
        self.pastasAuxiliaresDev = []
        self.pathPastasAtualiza = []

        self.config = configparser.ConfigParser()
        self.config.read(r'Config\Config.ini')

        self.pathConfigAtualiza = self.config['ConfigPath']['PathAtualizaIni']
        self.pathConfigDev = self.config['ConfigPath']['PathLocalIni']
        self.pathArquivoConfigDev = self.config['ConfigPath']['PathConfigLocal']

        self.origem = self.config['PathDesenvolvimento']['PathLocal']
        self.destino = self.config['PathDesenvolvimento']['PathAtualiza']

        self.pastaDevPresentation = self.config['FoldersDesenvolvimento']['Presentation']
        self.pastaDevApplication = self.config['FoldersDesenvolvimento']['Application']
        self.pastaDevDomain = self.config['FoldersDesenvolvimento']['Domain']
        self.pastaDevInfra = self.config['FoldersDesenvolvimento']['Infra']
        self.pastaDevIni = self.config['ConfigPath']['PathLocalIni']

        self.pastaAtualizaPresentation = self.config['FoldersAtualiza']['Presentation']
        self.pastaAtualizaApplication = self.config['FoldersAtualiza']['Application']
        self.pastaAtualizaDomain = self.config['FoldersAtualiza']['Domain']
        self.pastaAtualizaInfra = self.config['FoldersAtualiza']['Infra']
        self.pastaAtualizaIni = self.config['FoldersAtualiza']['ConfigIni']

        self.pastasAuxiliaresDev.append(self.pastaDevPresentation)
        self.pastasAuxiliaresDev.append(self.pastaDevApplication)
        self.pastasAuxiliaresDev.append(self.pastaDevDomain)
        self.pastasAuxiliaresDev.append(self.pastaDevInfra)
        self.pastasAuxiliaresDev.append(self.pastaDevIni)

        self.pathPastasAtualiza.append(self.pastaAtualizaPresentation)
        self.pathPastasAtualiza.append(self.pastaAtualizaApplication)
        self.pathPastasAtualiza.append(self.pastaAtualizaDomain)
        self.pathPastasAtualiza.append(self.pastaAtualizaInfra)
        self.pathPastasAtualiza.append(self.pastaAtualizaIni)

        self.publishExecutavel()
        self.SubirPastas()
        self.alterarVersaoLocal()
        self.alterarVersaoAtualiza()

    def publishExecutavel(self):
        if not os.path.isdir(self.destino):
            os.makedirs(self.destino)
        
        print("Excluindo Automação antiga")
        shutil.rmtree(self.destino)  # Remove a pasta de destino se ela já existir
        print(f"subindo atualização para {self.destino}")
        time.sleep(5)
        shutil.copytree(self.origem, self.destino)


    def SubirPastas(self):
        i=0
        for pasta in self.pastasAuxiliaresDev:
            print("Excluindo pastas antiga")

            if not os.path.isdir(self.pathPastasAtualiza[i]):
                os.makedirs(self.pathPastasAtualiza[i])

            shutil.rmtree(self.pathPastasAtualiza[i])  # Remove a pasta de destino se ela já existir
            print(f"subindo pastas auxiliares para {self.pathPastasAtualiza[i]}")
            shutil.copytree(pasta, self.pathPastasAtualiza[i])
            i+=1
        
    def alterarVersaoAtualiza(self):
        configAtualiza = configparser.ConfigParser()
        configAtualiza.read(self.pathConfigAtualiza)

        versao = semantic_version.Version(configAtualiza['ConfigVersion']['Version'])

        incrementoVersao = str(versao.next_patch())

        configAtualiza['ConfigVersion']['version'] = incrementoVersao

        with open(self.pathConfigAtualiza, 'w') as configfile:
            configAtualiza.write(configfile)

        print(f"Versão do Sistema alterado para {incrementoVersao}")

    def alterarVersaoLocal(self):
        self.versao = semantic_version.Version(self.config['ConfigVersion']['Version'])
        self.incrementoVersao = str(self.versao.next_patch())
        self.config['ConfigVersion']['Version'] = self.incrementoVersao

        with open(self.pathArquivoConfigDev, 'w') as configfile:
            self.config.write(configfile)

        print(f"Versão do Sistema alterado para {self.incrementoVersao}")
        

PublishExe()




