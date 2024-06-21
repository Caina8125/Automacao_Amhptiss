from abc import ABC, abstractmethod
import pyautogui
from selenium.webdriver.remote.webelement import WebElement
import seleniumwire.webdriver as wire_driver
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By
from time import sleep

class PageElement(ABC):
    body: tuple = (By.XPATH, '/html/body')
    a: tuple = (By.TAG_NAME, 'a')
    p: tuple = (By.TAG_NAME, 'p')
    h1: tuple = (By.TAG_NAME, 'h1')
    h2: tuple = (By.TAG_NAME, 'h2')
    h3: tuple = (By.TAG_NAME, 'h3')
    h4: tuple = (By.TAG_NAME, 'h4')
    h5: tuple = (By.TAG_NAME, 'h5')
    h6: tuple = (By.TAG_NAME, 'h6')
    table: tuple = (By.TAG_NAME, 'table')

    def __init__(self, url) -> None:
        self.url = url

    def init_driver(self, py_auto_gui=False):
        if not py_auto_gui:
            options: dict = {
            'proxy' : {
                'http': 'http://faturamento.fat:87316812#hg12@@10.0.0.230:3128',
                'https': 'https://faturamento.fat:87316812#hg12@@10.0.0.230:3128'
                }
            }

            chrome_options: Options = Options()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument('--ignore-certificate-errors')
            chrome_options.add_argument('--ignore-ssl-errors')

            try:
                servico: Service = Service(ChromeDriverManager().install())
                self.driver = wire_driver.Chrome(service=servico, seleniumwire_options= options, options = chrome_options)
            except:
                self.driver = wire_driver.Chrome(seleniumwire_options= options, options = chrome_options) 
        
        chrome_options: Options = Options()
        chrome_options.add_argument("--start-maximized")

        try:
            servico: Service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=servico, options = chrome_options)
        except:
            self.driver = webdriver.Chrome(options = chrome_options)
            
        self.open()
        sleep(4)
        pyautogui.write('faturamento.fat')
        pyautogui.press("TAB")
        sleep(0.5)
        pyautogui.write("87316812#hg12@")
        pyautogui.press("enter")
        sleep(4)

    def open(self) -> None:
        self.driver.get(self.url)

    def get_attribute_value(self, element: tuple, atributo: str) -> str | None:
        return self.driver.find_element(*element).get_attribute(atributo)

    def confirma_valor_inserido(self, element: tuple, valor: str) -> None:
        """Este método verifica se um input recebeu os valores que foram enviados.
           Caso não tenha recebido, tenta enviar novamente até 10x."""
        try:
            self.driver.find_element(*element).clear()
            valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
            count: int = 0

            while valor_inserido == '':
                self.driver.find_element(*element).send_keys(valor)
                sleep(0.5)
                valor_inserido: str = self.driver.find_element(*element).get_attribute('value')
                count += 1

                if count == 10:
                    raise Exception("Element not interactable")

        except Exception as e:
            raise Exception(e)
        
    def get_element_visible(self, element: tuple | None = None,  web_element: WebElement | None = None) -> bool:
        """Este método observa se irá ocorrer ElementClickInterceptedException. Caso ocorra
        irá dar um scroll até 10x na página conforme o comando passado até achar o click do elemento"""
        for i in range(10):
            try:
                if element != None:
                    self.driver.find_element(*element).click()
                    return True
                
                elif web_element != None:
                    web_element.click()
                    return True
            
            except ElementClickInterceptedException as e:
                if i == 10:
                    return False
                
                self.driver.execute_script('scrollBy(0,100)')
                continue

    def get_click(self, element: tuple, valor: str) -> None:
        for i in range(10):
            self.driver.find_element(*element).click()
            sleep(3)
            content: str = self.driver.find_element(*self.body).text

            if valor in content:
                break
            
            else:
                if i == 10:
                    raise Exception('Element not interactable')
                
                sleep(2)
                continue

    def click(self, element, tempo):
        self.driver.find_element(*element).click()
        sleep(tempo)

    def send_keys(self, element, texto, tempo):
        self.driver.find_element(*element).send_keys(texto)
        sleep(tempo)

    def clear(self, element, tempo):
        self.driver.find_element(*element).clear()
        sleep(tempo)

    def elemento_existe(self, element: tuple):
        try: 
            self.driver.find_element(*element)
            return True
        except:
            return False
        
    @abstractmethod    
    def inicia_automacao(self, **kwargs): ...