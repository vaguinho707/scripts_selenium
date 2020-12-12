from selenium.webdriver import Chrome
from selenium.webdriver import ChromeOptions
from selenium.common.exceptions import StaleElementReferenceException
import time
import os
import glob


#Definição de classes

class SiteOrtolookAdm:
    """Classe do site Administrativo da Ortolook"""

    def __init__(self):
        self.site = Chrome()

    def clicar(self,xpath):
        while True:
            try:
                self.acao = self.site.find_element_by_xpath(xpath)
                self.acao.click()
                break
            except StaleElementReferenceException:
                time.sleep(2)
                continue

    def formulario(self,xpath,texto):
        while True:
            try:
                self.formulariovar = self.site.find_element_by_xpath(xpath)
                self.formulariovar.clear()
                time.sleep(1)
                self.formulariovar.send_keys(texto)
                break
            except StaleElementReferenceException:
                time.sleep(2)
                continue

    def defineDownloadPath(self, downloadPath=None,headless=True):
        options = ChromeOptions()

        if downloadPath is not None:
            prefs = {}
            prefs["profile.default_content_settings.popups"]=0
            prefs["download.default_directory"]=downloadPath
            options.add_experimental_option("prefs", prefs)
        return options

    def apagarArquivosAntigos(self,caminhoPasta):
        arquivosPasta = glob.glob(caminhoPasta)
        for arquivo in arquivosPasta:
            os.remove(arquivo)

    def renomearArquivos(self, pastaDownload, nomesArquivosClinicas):
        arquivosPasta = glob.glob(pastaDownload)
        os.chdir(pastaDownload[:-2])
        aux = 0
        for arquivo in arquivosPasta:
            os.rename(arquivo,nomesArquivosClinicas[aux])
            aux += 1

    def getKey(dictCaminhos,val):
        for key, value in dictCaminhos.items():
            if val == value:
                return key

    def esperarLoading(webdriver):
        acharString = webdriver.find_element_by_xpath('//*[@id="j_idt23_complete"]').value_of_css_property("style.display")
        return bool(acharString)


if __name__ == "__main__":
    caminhos = {
        'username': '//*[@id="username"]',
        'password': '//*[@id="password"]',
        'entrar': '//*[@id="j_idt12"]',
        'relatorios': '/html/body/aside/ul/li[9]/a',
        'financeiro': '/html/body/aside/ul/li[9]/ul/li[1]/a',
        'pagamentos': '/html/body/aside/ul/li[9]/ul/li[1]/ul/li[1]/a',
        'tipoPagamento': '/html/body/aside/ul/li[9]/ul/li[1]/ul/li[1]/ul/li/a',
        'tiposVisualizacao': '/html/body/section/div/div/div/form/table/tbody/tr[1]/td[2]/div[1]',
        'baixar': '/html/body/div[5]/div/ul/li[2]',
        'formatosArquivo': '/html/body/section/div/div/div/form/table/tbody/tr[1]/td[2]/div[2]',
        'formatoXLSX': '/html/body/div[7]/div/ul/li[2]',
        'listarFranquias': '//*[@id="frm-print-relatorio:idGrupoConta"]',
        'ortolook': '//*[@id="frm-print-relatorio:idGrupoConta_1"]',
        'listarClinicas': '//*[@id="frm-print-relatorio:idConta"]',
        'alcantara': '//*[@id="frm-print-relatorio:idConta_1"]',
        'botafogo': '//*[@id="frm-print-relatorio:idConta_2"]',
        'caxias1': '//*[@id="frm-print-relatorio:idConta_3"]',
        'caxias2': '//*[@id="frm-print-relatorio:idConta_4"]',
        'leblon': '//*[@id="frm-print-relatorio:idConta_5"]',
        'madureira': '//*[@id="frm-print-relatorio:idConta_6"]',
        'niteroi': '//*[@id="frm-print-relatorio:idConta_7"]',
        'novaIguacu': '//*[@id="frm-print-relatorio:idConta_8"]',
        'saoGoncalo': '//*[@id="frm-print-relatorio:idConta_9"]',
        'dataInicio': '//*[@id="frm-print-relatorio:j_idt95_input"]',
        'dataFim': '//*[@id="frm-print-relatorio:j_idt97_input"]',
        'gerarRelatorio': '//*[@id="frm-print-relatorio:j_idt100"]'
    }
    #Início do Código Principal

    URL = r'http://adm.dttsistemas.com.br/pages/autenticacao.xhtml'
    downloadPath = r"C:\Users\Rafael\Documents\Atualizacao BD - Recebimento"
    browser = SiteOrtolookAdm()
    options = browser.defineDownloadPath(downloadPath)

    browser.site = Chrome(chrome_options = options)
    browser.site.get(URL)

    #Deletando todas as planilhas antigas da pasta
#    caminhoPasta = r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\*'
    browser.apagarArquivosAntigos(downloadPath + r"\*")


    #Leitura das datas de início e fim no arquivo txt
    caminhoDatasTxt = r"C:\Users\Rafael\DatasRecebimento.txt"
    with open(caminhoDatasTxt,"r") as arquivoDatas:
        datas = arquivoDatas.readlines()
        inicio = datas[0][:8]
        fim = datas[1][:8]

    #Preenchimento e email e senha
    email = '***@gmail.com'
    senha = 'xxx'

    browser.formulario(caminhos['username'],email)
    browser.formulario(caminhos['password'],senha)

    #Clicar em "ENTRAR"
    browser.clicar(caminhos['entrar'])

    #Clicar em "Relatórios"
    browser.clicar(caminhos['relatorios'])

    #Clicar em "Financeiro"
    browser.clicar(caminhos['financeiro'])

    #Clicar em "Pagamentos"
    browser.clicar(caminhos['pagamentos'])
    time.sleep(1)


    #Clicar em "Tipo de Pagamento"
    browser.clicar(caminhos['tipoPagamento'])
    time.sleep(1)


    #Clicar para listar tipos de visualização
    browser.clicar(caminhos['tiposVisualizacao'])

    #Selecionar "Baixar"
    browser.clicar(caminhos['baixar'])

    #Clicar para listar formatos de arquivo
    browser.clicar(caminhos['formatosArquivo'])

    #Selecionar formato XLSX
    browser.clicar(caminhos['formatoXLSX'])

    #Clicar para listar Franquias
    browser.clicar(caminhos['listarFranquias'])

    #Selecionar "Ortolook"
    browser.clicar(caminhos['ortolook'])

    #Clicar para Listar Clínicas
    browser.clicar(caminhos['listarClinicas'])

    #Selecionar Alcântara
    browser.clicar(caminhos['alcantara'])
    time.sleep(1)

    #Inserir datas de início e fim
    browser.formulario(caminhos['dataInicio'],inicio)
    browser.formulariovar.submit()
    time.sleep(1)

    browser.formulario(caminhos['dataFim'],fim)
    browser.formulariovar.submit()
    time.sleep(1)

    #Clicar em "Gerar Relatório"
    browser.clicar(caminhos['gerarRelatorio'])

    restanteDasClinicas = [caminhos['botafogo'], caminhos['caxias1'], caminhos['caxias2'], caminhos['leblon'], caminhos['madureira'], caminhos['niteroi'], caminhos['novaIguacu'], caminhos['saoGoncalo']]

    #loop iterativo para percorrer clinicas e baixar seus arquivos
    for item in restanteDasClinicas:
        #Clicar para Listar Clínicas
        browser.clicar(caminhos['listarClinicas'])

        #Selecionar Clínica
        browser.clicar(item)
        time.sleep(1)

        #Clicar em "Gerar Relatório"
        browser.clicar(caminhos['gerarRelatorio'])


    time.sleep(40)

#    pastaDownload = r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\*'
    nomesArquivosClinicas = ['ALCANTARA.xlsx', 'BOTAFOGO.xlsx','CAXIAS.xlsx', 'CAXIAS 2.xlsx', 'LEBLON.xlsx', 'MADUREIRA.xlsx', 'NITEROI.xlsx', 'NOVA IGUACU.xlsx', 'SAO GONCALO.xlsx']
    browser.renomearArquivos(downloadPath + r"\*", nomesArquivosClinicas)
