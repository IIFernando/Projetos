import pandas as pd
from selenium import webdriver
from time import sleep as slp
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options

print('Opções Automação',
      '[1] Automação Guias AM',
      '[2] Automação Guias MS',
      '[3] Automação Guias RN',
      '[4] Automação Guias PE',
      '[5] Aceitar Termo de Fiel Depositário - Carga GRANEL', sep='\n')

opcao = input('Informe o número do processo: ')

def FielAM():

    lib = int(input('Quantos termos seram liberados: '))

    servico = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=servico)

    browser.implicitly_wait(125)
    browser.get('http://www.sefaz.am.gov.br/')
    browser.maximize_window()

    slp(60)
    for r in range(1, lib):
        browser.find_element('xpath', '//*[@id="cpfProcuracao"]').send_keys('a')
        browser.find_element('xpath', '//*[@id="registros"]/tbody/tr[6]/td[4]/a/img').click()
        alert = Alert(browser)
        alert.accept()
        slp(5)
        browser.refresh()
        slp(10)

def GuiaMS():
    caminhoA = input(r'Informe o caminho da planilha: ')
    caminhoG = input(r'informe o caminho para salvar as Guias: ')
    dVenc = input('O vencimento será alterado S/N? ').upper()

    df = pd.read_excel(caminhoA, dtype=str, sheet_name='Guias faltantes')

    # Configurações de download
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs',
                                           {'plugins.plugins_list': [{'enabled': False, 'name': 'Chrome PDF Viewer'}],
                                            'download.default_directory': caminhoG,
                                            'download.prompt_for_download': False,
                                            'download.directory_upgrade': True,
                                            'plugins.always_open_pdf_externally': True,
                                            'safebrowsing.enabled': True})

    servico = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=servico, chrome_options=chrome_options)
    browser.maximize_window()
    browser.get('https://servicos.efazenda.ms.gov.br/daemsabertopublico/emissaotermosfronteira?abertoAoPublico=True')

    if dVenc == 'S':
        data = input('Informe o vencimento: ')
    else:
        print('Guias teram vencimento na data de HOJE.')

    for row in df.itertuples():

        n = row.TVF
        c = row.CNPJ

        tipoTVF = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tipoTVF"]')))
        tipoTVF.click()

        termo = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="numero"]')))
        termo.send_keys(n)

        posto = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="posto"]')))
        posto.send_keys('150')

        opcaoCnpj = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="cnpj"]')))
        opcaoCnpj.click()

        cnpj = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Documento"]')))
        cnpj.click()

        cnpj = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Documento"]')))
        cnpj.send_keys(c)

        if data is not None:
            pagamento = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.ID, 'DataPagamento')))
            pagamento.clear()
            pagamento.send_keys(data)
        else:
            print('Guia vence HOJE.')

        consulta = WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="consultar"]')))
        consulta.click()
        slp(5)

        # Tratativa para os erros de termos já baixados
        try:
            mensagem = browser.find_element('xpath', '//*[@id="menssagem"]')
            browser.refresh()
            slp(1)
            continue

        except NoSuchElementException:
            print('Termo OK.')

        # Tratativa para caso o registro não carregue.
        try:
            registro = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="gridFront"]/tbody')))
            registro.click()

        except NoSuchElementException:
            print(f'Registro do termo {n} não localizado.')
            browser.back()
            slp(1)
            continue

        browser.find_element('xpath', '//*[@id="daemsBody"]/div[14]/button[1]').click()
        slp(5)

        browser.back()
        slp(1)

        browser.refresh()
        slp(1)

def GuiaRN():
    caminhoA = input(r'Informe o caminho da planilha: ')
    caminhoG = input(r'informe o caminho para salvar as Guias: ')
    dataVenc = input('Informe o vencimento:')

    # Configurações do Chrome
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs',
                                           {'plugins.plugins_list': [{'enabled': False, 'name': 'Chrome PDF Viewer'}],
                                            'download.default_directory': caminhoG,
                                            'download.prompt_for_download': False,
                                            'download.directory_upgrade': True,
                                            'plugins.always_open_pdf_externally': True,
                                            'safebrowsing.enabled': True})

    # Configuração do driver
    servico = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=servico, chrome_options=chrome_options)
    browser.implicitly_wait(60)
    browser.maximize_window()

    # Leitura dos dados
    df = pd.read_excel(caminhoA, dtype=str)

    cont = 0

    def go_back():
        browser.back()
        slp(1)

    browser.get('https://uvt2.set.rn.gov.br/#/services/icms-antecipado')

    for row in df.itertuples():

        n = row.NOTA
        c = row.CNPJ

        Identificador = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/section/div/form/div/div[1]/div[1]/div/input')))
        Identificador.send_keys(c)

        vencimento = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/section/div/form/div/div[1]/div[2]/div/div/input')))
        vencimento.send_keys(dataVenc)

        nota = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/section/div/div[4]/div[2]/form/div/div[1]/div[1]/div/input')))
        nota.send_keys(n)

        cnpj = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/section/div/div[4]/div[2]/form/div/div[2]/div/input')))
        cnpj.send_keys(c)

        enviar = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/section/div/div[6]/div/input')))
        enviar.click()

        with open(caminhoG + '\Log_RN.csv', 'a', newline='', encoding='UTF-16') as arquivo:

            vDifal = browser.find_element('xpath',
                                          '/html/body/section/div/div[3]/div/div[3]/div/div/input').get_attribute(
                'value')

            if vDifal != '0,00':
                browser.find_element('xpath', '/html/body/section/div/div[3]/div/div[4]/div/div[2]/label/input').click()
                slp(1)
                browser.find_element('xpath', '/html/body/section/div/div[4]/input').click()
                slp(1)
                browser.find_element('xpath', '/html/body/div[1]/div/div/div[3]/button').click()
                slp(1)

                go_back()
                go_back()

                cont += 1

                arquivo.write(n + '_' + '_' + str(vDifal) + '_' + str('Gerada'))
                arquivo.write(str('\n'))

            elif vDifal == '0,00':
                arquivo.write(n + '_' + str('Não Gerada'))
                arquivo.write(str('\n'))

                cont += 1
                browser.back()
                slp(1)

    browser.quit()

def GuiaPE():
    caminhoA = input(r'Informe o caminho da planilha: ')
    caminhoG = input(r'informe o caminho para salvar as Guias: ')
    dVenc = input('Informe a data de vencimento: ')

    df = pd.read_excel(caminhoA, dtype=str)

    # Configurações de download
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs',
                                           {'plugins.plugins_list': [{'enabled': False, 'name': 'Chrome PDF Viewer'}],
                                            'download.default_directory': caminhoG,
                                            'download.prompt_for_download': False,
                                            'download.directory_upgrade': True,
                                            'plugins.always_open_pdf_externally': True,
                                            'safebrowsing.enabled': True})

    servico = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=servico, chrome_options=chrome_options)

    browser.get('https://efisco.sefaz.pe.gov.br/sfi_trb_cmt/PRConsultarNFTermoFielDepositario')
    browser.maximize_window()
    slp(5)

    with open(caminhoG + '\Log_PE.csv', 'a', newline='', encoding='UTF-16') as arquivo:

        for c in df['REGISTRO']:
            browser.find_element('xpath', '//*[@id="primeiro_campo"]').clear()
            browser.find_element('xpath', '//*[@id="primeiro_campo"]').send_keys(c)
            browser.find_element('xpath', '//*[@id="btt_localizar"]').click()
            browser.find_element('xpath', '//*[@id="btt_detalhar"]').click()
            chaveA = browser.find_element('xpath',
                    '/html/body/form/table/tbody/tr[2]/td/div/table/tbody/tr/td/input').get_attribute('value')
            browser.back()
            slp(3)

            browser.find_element('xpath', '//*[@id="bt_emitirgnre"]').click()
            browser.find_element('xpath', '//*[@id="DtPagamentoGNRE"]').click()
            browser.find_element('xpath', '//*[@id="DtPagamentoGNRE"]').send_keys(dVenc)

            vDifal = browser.find_element('xpath', '//*[@id="VL_DIFAL_ABERTO_' + c + '"]').get_attribute('value')
            rSocial = browser.find_element('xpath', '//*[@id="PESSOA_NM_RAZAO_SOCIAL"]').get_attribute('value')
            cnjp = browser.find_element('xpath', '//*[@id="NuDocumentoEmitente"]').get_attribute('value')
            nNota = browser.find_element('xpath', '//*[@id="table_tabeladados"]/tbody/tr[3]/td[3]').get_attribute(
                'class')

            if vDifal != '0,00':
                browser.find_element('xpath', '//*[@id="btt_confirmaremitirgnre"]').click()
                browser.find_element('xpath', '/html/body/div/form/div[2]/div/div/div/div[2]/p[2]/input').click()
                browser.back()
                browser.back()

            else:
                vDifal = browser.find_element('xpath', '/html/body/form/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td[7]')
                valor = vDifal.text
                browser.find_element('xpath', '//*[@id="VL_DIFAL_ABERTO_' + c + '"]').send_keys(valor)
                browser.find_element('xpath', '//*[@id="btt_confirmaremitirgnre"]').click()
                browser.find_element('xpath', '/html/body/div/form/div[2]/div/div/div/div[2]/p[2]/input').click()
                slp(1)
                browser.back()
                browser.back()
                continue

            arquivo.write(
                c + '_' + str(cnjp) + '_' + str(rSocial) + '_' + str(nNota) + '_' + str(chaveA) + '_' + str(vDifal))
            arquivo.write(str('\n'))

        browser.quit()

def GuiaAM():
    caminhoA = input(r'Informe o caminho da planilha: ')
    caminhoG = input(r'informe o caminho para salvar as Guias: ')

    df = pd.read_excel(caminhoA, dtype=str)

    # Configurações de download
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs',
                                           {'plugins.plugins_list': [{'enabled': False, 'name': 'Chrome PDF Viewer'}],
                                            'download.prompt_for_download': False,
                                            'download.directory_upgrade': True,
                                            'plugins.always_open_pdf_externally': True,
                                            'safebrowsing.enabled': True})

    servico = Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=servico, chrome_options=chrome_options)

    browser.maximize_window()

    with open(caminhoG + '\Log_AM.csv', 'a', newline='', encoding='UTF-16') as arquivo:

        for row in df.itertuples():

            c = row.CHAVE

            browser.get('http://online.sefaz.am.gov.br/daravulso/darFormDesembaracoAvulsoNovo.asp')

            browser.find_element('xpath', '//*[@id="chave"]').clear()
            browser.find_element('xpath', '//*[@id="chave"]').send_keys(c)
            slp(1)

            browser.find_element('xpath', '//*[@id="search"]').click()
            slp(1)

            try:
                imprimir = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="print"]')))
                imprimir.click()
                slp(1)

                html = browser.page_source
                html = html.replace('\n', '|')
                soup = BeautifulSoup(html, "html.parser")
                text = soup.get_text()
                slp(1)

                arquivo.write(c + '_' + str(text))
                arquivo.write(str('\n'))
                slp(1)

                browser.execute_script("document.body.style.zoom='90%'")
                browser.set_window_size(800, 1200)
                browser.save_screenshot(c + '.png')
                slp(1)

            except:
                arquivo.write(c + '_' + str('Guia não gerada.'))
                arquivo.write(str('\n'))

    browser.quit()

if opcao == '2':
    GuiaMS()

elif opcao == '3':
    GuiaRN()

elif opcao == '4':
    GuiaPE()

elif opcao == '1':
    GuiaAM()

elif opcao == '5':
    FielAM()