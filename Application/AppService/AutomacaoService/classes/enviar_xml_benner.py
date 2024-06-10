from selenium.webdriver.chrome.options import Options
from seleniumwire.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from tkinter.filedialog import askopenfilenames

from classes.benner import Benner

def enviar_xml_benner(user, password) -> None:
    URL: str = r'https://portalconectasaude.com.br/Account/Login?ReturnUrl=%2FHome%2FIndex'

    DIRETORIO: tuple = askopenfilenames()

    PROXY: dict = {
    'proxy': {
            'http': f'http://{user}:{password}@10.0.0.230:3128',
            'https': f'http://{user}:{password}@10.0.0.230:3128'
        }
    }

    chrome_options: Options = Options()          
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--kiosk-printing')


    EMAIL: str = 'negociacao.gerencia@amhp.com.br'
    SENHA: str = 'Amhp@0073'

    try:
        servico = Service(ChromeDriverManager().install())
        DRIVER = Chrome(service=servico, seleniumwire_options=PROXY, options=chrome_options)
    except:
        DRIVER = Chrome(seleniumwire_options=PROXY, options=chrome_options)

    envio_xml_benner = Benner(DRIVER, URL, EMAIL, SENHA)
    envio_xml_benner.exec_envio_xml(DIRETORIO)