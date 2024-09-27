# substitua o site desse scrapy pelo o de seu interesse. pois como se trata de um sistema da Força Aérea Brasileira, só as maquinas de lá tem acesso devido proxy. caso tente executar esse código não tera sucesso por esse motivo.




import scrapy
import logging
from selenium.webdriver.remote.remote_connection import LOGGER
from selenium.common.exceptions import *
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from scrapy.selector import Selector
from selenium.webdriver.support.select import Select
from scrapy.crawler import CrawlerProcess
from time import sleep
from scrapy.settings import Settings
from selenium.webdriver.chrome.service import Service
import pandas as pd
from openpyxl.workbook import Workbook



def iniciar_driver():
    chrome_options = Options()
    LOGGER.setLevel(logging.WARNING)
    arguments = ['--lang=pt-BR', '--window-size=1920,1080', '--incognito', '--ignore-certificate-errors', '--allow-running-insecure-content', '--headless']

    caminho_chromedriver = "C:\\Users\\rebello\\Downloads\\chromedriver-win64\\chromedriver.exe" # não necessário, pois ja vai automatico nas maquinas pessoas, fiz dessa forma devido o proxy do meu setor que não deixar fazer da forma mais atualizada
    service = Service(caminho_chromedriver)

    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1,
    })
    driver = webdriver.Chrome(service=service, options=chrome_options)

    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait


class ProductScraperSpider(scrapy.Spider):
    name = 'siloms'
    def start_requests(self):
        urls = ['http://transporte.siloms.intraer:8080/modulotransporte/servlet/htra00007']
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse, meta={'proximo_url': url})

    def parse(self, response):
        driver, wait = iniciar_driver()
        driver.get('http://transporte.siloms.intraer:8080/modulotransporte/servlet/htra00007')

        janela_inicial = driver.current_window_handle
        sleep(1)

        try:
            usuario = wait.until(EC.presence_of_element_located((By.XPATH, '//form/div/div/input[@id="username"]')))
            usuario.send_keys('18871343794')
                        
            senha = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@id="password"]')))
            senha.send_keys('Rebello123@')
                        
            entrar_na_conta = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="btn btn-primary btn-lg btn-block"]')))
            entrar_na_conta.click()
        except TimeoutException:
            print("Tempo de espera esgotado para elementos na nova aba.")
        except NoSuchElementException as e:
            print(f"Elemento não encontrado: {e}")

        driver.switch_to.window(janela_inicial)
        sleep(2)

        # Clicar no botão de avançar
        avancar = driver.find_element(By.XPATH,'//button[@id="proceed-button"]')
        sleep(1)
        avancar.click()

        # Confirmar o alerta
        alerta2 = driver.switch_to.alert
        sleep(2)
        alerta2.accept()

        driver.get('http://transporte.siloms.intraer:8080/modulotransporte/servlet/htra00007')

        todos_volumes = driver.find_element(By.XPATH,'//select[@id="vGRIDSHOW"]')
        Select(todos_volumes).select_by_index(2)

        disponiveis = driver.find_element(By.XPATH,'//select[@id="vCST_VOL_LOG"]')
        Select(disponiveis).select_by_index(2)

        # Selecionar todas as opções de PCAN e varrer os dados
        pcans = driver.find_element(By.XPATH,'//select[@id="vID_PTC_DESTINO"]')
        pcans_options = Select(pcans).options

        all_data = {}

        for i, option in enumerate(pcans_options):
            if i == 0:  # Pula a primeira opção se for um valor padrão, como "Selecione"
                continue

            Select(pcans).select_by_index(i)
            sleep(2)

            botao_pesquisar = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@name="BTN_REFRES"]')))
            botao_pesquisar.click()
            sleep(8)

            response_webdriver = Selector(text=driver.page_source)
            rows = response_webdriver.xpath('//table/tbody/tr[@style="font-family:Verdana;font-size:8pt;"]')

            if rows:
                data = []
                for produto in rows:
                    data.append({
                        'volume': produto.xpath("./td[3]/span/a/text()").get(),
                        'unid_Origem': produto.xpath("./td[11]/span/text()").get(),
                        'unid_Destin': produto.xpath("./td[15]/span/text()").get(),
                        'peso': produto.xpath("./td[27]/span/text()").get()
                    })
                
                all_data[option.text] = data

        driver.quit()

        # Criar arquivo Excel com uma aba para cada PCAN
        with pd.ExcelWriter('dados_carga.xlsx') as writer:
            for sheet_name, data in all_data.items():
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        


# Definições de configurações adicionais
settings = {
    'USER_AGENT': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36',
    'ROBOTSTXT_OBEY': False,
    'CONCURRENT_REQUESTS': 32,
    'DOWNLOAD_DELAY': 5,
    'FEEDS': {
        'carga.csv': {
            'format': 'csv',
            'encoding': 'utf8',
        },
    },
    'REQUEST_FINGERPRINTER_IMPLEMENTATION': '2.7',
}

process = CrawlerProcess(settings)
process.crawl(ProductScraperSpider)
process.start()
