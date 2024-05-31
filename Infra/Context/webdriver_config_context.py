import time
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showerror, showinfo
import pyautogui
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from seleniumwire import webdriver


def config(convenio_context) -> None:
    try:
        user = ''
        password = ''
        showinfo('', 'Selecione a pasta com as planilhas.')
        diretorio_planilhas = askdirectory()
        showinfo('', 'Selecione a pasta com os anexos')
        diretorio_anexos = askdirectory()

        url = 'https://wwwt.connectmed.com.br/conectividade/prestador/home.htm'

        proxies = {
                'http': f'http://{user}:{password}@10.0.0.230:3128',
                'https': f'http://{user}:{password}@10.0.0.230:3128'
            }

        options = {
            'proxy' : proxies
        }

        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        # chrome_options.add_argument('--ignore-certificate-errors')
        # chrome_options.add_argument('--ignore-ssl-errors')
        try:
            servico = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=servico, options = chrome_options)
        except:
            driver = webdriver.Chrome(options = chrome_options)

        usuario = "glosaamhp"
        senha = "D8nEUqb!"

        convenio_context.iniciar()

    except Exception as e :
        showerror('Automação', f'Ocorreu uma exceção não trata\n{e.__class__.__name__}:\n{e}')
        driver.quit()