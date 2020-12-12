from selenium.webdriver import Chrome
from selenium.common.exceptions import StaleElementReferenceException
from Download_Planilhas_de_Receita import SiteOrtolookAdm
import time
import os

#Início do Código Principal
if __name__ == "__main__":

#Dicionário com os xPaths dos elementos que serão acessados no site.
    caminhos = {
        'username': '//*[@id="username"]',
        'password': '//*[@id="password"]',
        'entrar': '//*[@id="j_idt12"]',
        'relatorios': '/html/body/aside/ul/li[9]/a',
        'produtividade': '/html/body/aside/ul/li[9]/ul/li[2]/a',
        'analiticoDentista': '/html/body/aside/ul/li[9]/ul/li[2]/ul/li/a',
        'tiposVisualizacao': '//*[@id="frm-print-relatorio:j_idt75"]',
        'baixar': '//*[@id="frm-print-relatorio:j_idt75_1"]',
        'formatosArquivo': '//*[@id="frm-print-relatorio:j_idt77"]',
        'formatoXLSX': '//*[@id="frm-print-relatorio:j_idt77_1"]',
        'listarFranquias': '//*[@id="frm-print-relatorio:idGrupoConta"]',
        'ortolook': '//*[@id="frm-print-relatorio:idGrupoConta_1"]',
        'listarClinicas': '//*[@id="frm-print-relatorio:idConta"]',
        'ALCANTARA': '//*[@id="frm-print-relatorio:idConta_1"]',
        'BOTAFOGO': '//*[@id="frm-print-relatorio:idConta_2"]',
        'CAXIAS': '//*[@id="frm-print-relatorio:idConta_3"]',
        'CAXIAS 2': '//*[@id="frm-print-relatorio:idConta_4"]',
        'LEBLON': '//*[@id="frm-print-relatorio:idConta_5"]',
        'MADUREIRA': '//*[@id="frm-print-relatorio:idConta_6"]',
        'NITEROI': '//*[@id="frm-print-relatorio:idConta_7"]',
        'NOVA IGUACU': '//*[@id="frm-print-relatorio:idConta_8"]',
        'SAO GONCALO': '//*[@id="frm-print-relatorio:idConta_9"]',
        'dataInicio': '//*[@id="frm-print-relatorio:j_idt95_input"]',
        'dataFim': '//*[@id="frm-print-relatorio:j_idt97_input"]',
        'gerarRelatorio': '//*[@id="frm-print-relatorio:j_idt100"]'
    }


    URL = 'http://adm.eauhekdkfaja.com.br/pages/autenticacao.xhtml'
    downloadPath = r'C:\Users\Rafael\Documents\Planilhas - Dentistas'
    browser = SiteOrtolookAdm()
    options = browser.defineDownloadPath(downloadPath)

    browser.site = Chrome(chrome_options = options)
    browser.site.get(URL)

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
    time.sleep(1)

    #Clicar em "Produtividade"
    browser.clicar(caminhos['produtividade'])

    #Clicar em "Analítico - Dentista"
    browser.clicar(caminhos['analiticoDentista'])



    #Clicar para listar Franquias
    browser.clicar(caminhos['listarFranquias'])

    #Selecionar "Ortolook"
    browser.clicar(caminhos['ortolook'])

    time.sleep(1)
    browser.site.back()

    #Clicar para listar tipos de visualização
    browser.clicar(caminhos['tiposVisualizacao'])

    #Selecionar "Baixar"
    browser.clicar(caminhos['baixar'])

    #Clicar para listar formatos de arquivo
    browser.clicar(caminhos['formatosArquivo'])

    #Selecionar formato XLSX
    browser.clicar(caminhos['formatoXLSX'])

    #Clicar para Listar Clínicas
    browser.clicar(caminhos['listarClinicas'])

    #Selecionar Alcântara
    browser.clicar(caminhos['ALCANTARA'])
    time.sleep(0.5)

    #Inserir datas de início e fim
    browser.formulario(caminhos['dataInicio'],inicio)
    browser.formulariovar.submit()

    browser.formulario(caminhos['dataFim'],fim)
    browser.formulariovar.submit()

    #Clicar em "Gerar Relatório"
    browser.clicar(caminhos['gerarRelatorio'])
    time.sleep(3)

    os.chdir(downloadPath)

    renameStatus = False
    while (renameStatus == False):
        try:
            time.sleep(1)
            os.rename('analitico-dentista.xlsx','ALCANTARA.xlsx')
            renameStatus = True
        except:
            pass


    restanteDasClinicas = [caminhos['BOTAFOGO'], caminhos['CAXIAS'], caminhos['CAXIAS 2'], caminhos['LEBLON'], caminhos['MADUREIRA'], caminhos['NITEROI'], caminhos['NOVA IGUACU'], caminhos['SAO GONCALO']]

    #loop iterativo para percorrer clinicas e baixar seus arquivos
    for item in restanteDasClinicas:
        #Clicar para Listar Clínicas
        browser.clicar(caminhos['listarClinicas'])

        #Selecionar Clínica
        browser.clicar(item)
        time.sleep(0.6)

        #Clicar em "Gerar Relatório"
        browser.clicar(caminhos['gerarRelatorio'])
        time.sleep(3)

       	nomeClinica = SiteOrtolookAdm.getKey(caminhos,item)
       	print(nomeClinica)

        renameStatus = False

        while (renameStatus == False):
            try:
                time.sleep(1)
       	        os.rename('analitico-dentista.xlsx', nomeClinica + '.xlsx')
                renameStatus = True
            except:
                pass
