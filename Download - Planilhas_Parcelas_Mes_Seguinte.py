from selenium.webdriver import Chrome
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from Download_Planilhas_de_Receita import SiteOrtolookAdm
from selenium.common.exceptions import NoSuchElementException
import time
import os
import calendar
import shutil


#Início do Código Principal
if __name__ == "__main__":

#Dicionário com os xPaths dos elementos que serão acessados no site.
    caminhos = {
    	'CNPJ': '//*[@id="cnpj"]',
        'username': '//*[@id="email"]',
        'password': '//*[@id="password"]',
        'entrar': '//*[@id="j_idt18"]',
        'selecionarDia': '//*[@id="frm-dataproc:j_idt43"]/span',
        'clientes': '//*[@id="frm-menu:j_idt40"]/ul/li[1]/ul/li[1]',
        'cadastro': '//*[@id="frm-menu:j_idt40"]/ul/li[1]/a/span[2]',
        'pesquisarClientes': '//*[@id="j_idt249:j_idt262"]',
        'clientesXLS': '//*[@id="j_idt249:j_idt267"]',
        'consultas': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/a/span[2]',
        'consultasClientes': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/ul/li[1]',
        'agenda': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/ul/li[1]/ul/li[1]/a',
        'inicioAgenda': '//*[@id="j_idt245:data-consulta_input"]',
        'fimAgenda': '//*[@id="j_idt245:data-fim_input"]',
        'pesquisarAgenda': '//*[@id="j_idt245:j_idt264"]',
        'agendaXLS': '//*[@id="j_idt245:j_idt268"]',
        'financeiro': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/ul/li[2]/a',
        'contratos': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/ul/li[2]/ul/li[1]/a',
        'inicioContratos':'//*[@id="j_idt244:data-inicio_input"]',
        'fimContratos': '//*[@id="j_idt244:data-fim_input"]',
        'pesquisarContratos': '//*[@id="j_idt244:j_idt257"]',
        'contratosXLS': '//*[@id="j_idt244:j_idt261"]',
        'sair':'//*[@id="frm-menu:j_idt40"]/ul/li[7]/a/span[2]',
        'sairDoSistema': '//*[@id="frm-menu:j_idt40"]/ul/li[7]/ul/li/a',
        'parcelas': '//*[@id="frm-menu:j_idt40"]/ul/li[4]/ul/li[2]/ul/li[2]/a',
        'inicioParcelas': '//*[@id="j_idt245:data-inicio-vencimento_input"]',
        'fimParcelas': '//*[@id="j_idt245:data-fim-vencimento_input"]',
        'pesquisarParcelas': '//*[@id="j_idt245:j_idt262"]',
        'parcelasXLS': '//*[@id="j_idt245:j_idt267"]',
        'loading': '//*[@id="j_idt23_complete"]',
        'erroAoProcessarSuaSolicitacao': '//*[@id="j_idt243_content"]/h2'
    }

    dictCNPJClinicas = {
		'ALCANTARA': '32814593000116',
		'BOTAFOGO': '34445861000103',
		'CAXIAS': '30530509000107',
		'CAXIAS 2': '34298866000142',
		'LEBLON': '00000000000102',
		'MADUREIRA': '35660946000169',
		'NITEROI':'28697866000112',
		'NOVA IGUACU':'05806879000187',
		'SAO GONCALO': '06271306000169'
	}


    arrayCNPJClinicas = [dictCNPJClinicas['ALCANTARA'], dictCNPJClinicas['BOTAFOGO'],dictCNPJClinicas['CAXIAS'], dictCNPJClinicas['CAXIAS 2'], dictCNPJClinicas['LEBLON'], dictCNPJClinicas['MADUREIRA'], dictCNPJClinicas['NITEROI'], dictCNPJClinicas['NOVA IGUACU'], dictCNPJClinicas['SAO GONCALO']]


    URL = 'http://www.dttsistemas.com.br/pages/autenticacao.xhtml'
    downloadPath = r'C:\Users\Rafael\Documents\Planilhas - Clientes'
    browser = SiteOrtolookAdm()
    options = browser.defineDownloadPath(downloadPath)
    os.chdir(downloadPath)


    browser.site = Chrome(chrome_options = options)
    browser.site.get(URL)

    browser.apagarArquivosAntigos(downloadPath + r"\*")

    #Leitura das datas de início e fim no arquivo txt
    caminhoDatasTxt = r"C:\Users\Rafael\DatasRecebimento.txt"
    with open(caminhoDatasTxt,"r") as arquivoDatas:
        datas = arquivoDatas.readlines()
        inicio = datas[0][:8]
        inicio.replace(datas[2:7],str((int(datas[2:4])+1)%12) + '202')
        fim = datas[1][:8]

    for clinica in arrayCNPJClinicas:
        try:

            #Preenchimento e email e senha
            email = 'odontologiafinanceiro@gmail.com'
            senha = 'saoconra'

            cnpj = browser.site.find_element_by_xpath(caminhos['CNPJ'])
            for i in range(18):
                cnpj.send_keys(Keys.ARROW_LEFT)
            time.sleep(2)

            cnpj.send_keys(clinica)
            browser.formulario(caminhos['username'],email)
            browser.formulario(caminhos['password'],senha)
            time.sleep(3)

            #Clicar em "ENTRAR"
            browser.clicar(caminhos['entrar'])
            time.sleep(3)
            # diaAbertoOuFechado = browser.site.find_element_by_xpath("//*[text()= "+ fim[1:2] + "]")
            # diaAbertoOuFechado = diaAbertoOuFechado.value_of_css_property("Background")

            # coresStatusDia = {
            #     'vermelho': '#FF0000',
            #     'amarelo': '#FFFF00',
            #     'verde': '#008000'
            # }
            # print(diaAbertoOuFechado)

            #selecionarDia = browser.find_element_by_xpath('//*[text()= "Selecionar Dia"]')

            #Clicar no botao "Selecionar Dia..."
            botaoSelecionarDia = browser.site.find_element_by_xpath(caminhos['selecionarDia'])
            action = ActionChains(browser.site)
            action.move_to_element(botaoSelecionarDia)
            action.click()
            action.perform()

            nomeClinica = SiteOrtolookAdm.getKey(dictCNPJClinicas, clinica)


            time.sleep(3)

            ultimoDiaMes = calendar.monthrange(int(inicio[4:8]), int(inicio[2:4])+1)[1]


            action = ActionChains(browser.site)
            botaoCadastro = browser.site.find_element_by_xpath(caminhos['consultas'])
            action.move_to_element(botaoCadastro).perform()
            botaoCadastro = browser.site.find_element_by_xpath(caminhos['financeiro'])
            action.move_to_element(botaoCadastro).perform()
            browser.clicar(caminhos['parcelas'])

            inicioParcelas = browser.site.find_element_by_xpath(caminhos['inicioParcelas'])
            for i in range(10):
                inicioParcelas.send_keys(Keys.ARROW_LEFT)
            time.sleep(2)

            inicioParcelas.send_keys(inicio)

            fimParcelas = browser.site.find_element_by_xpath(caminhos['fimParcelas'])
            for i in range(10):
                fimParcelas.send_keys(Keys.ARROW_LEFT)
            time.sleep(2)

            fimParcelas.send_keys(str(ultimoDiaMes) + str(((int(inicio[2:4])+1))%12) + inicio[4:] )


            browser.clicar(caminhos['pesquisarParcelas'])

            while ((browser.site.find_element_by_xpath(caminhos['loading']).value_of_css_property('display')) == 'none'):
                time.sleep(1)
                try:
                    avisoErroAoProcessar = browser.site.find_element_by_xpath(caminhos['erroAoProcessarSuaSolicitacao'])
                except:
                    pass
                else:
                    break


            time.sleep(3)

            browser.clicar(caminhos['parcelasXLS'])
            time.sleep(5)

            os.rename('Parcelas.xls', 'Parcelas - ' + nomeClinica + '.xls')


            action = ActionChains(browser.site)
            botaoCadastro = browser.site.find_element_by_xpath(caminhos['sair'])
            action.move_to_element(botaoCadastro).perform()
            browser.clicar(caminhos['sairDoSistema'])

        except:
            time.sleep(20)
            action = ActionChains(browser.site)
            botaoCadastro = browser.site.find_element_by_xpath(caminhos['sair'])
            action.move_to_element(botaoCadastro).perform()
            browser.clicar(caminhos['sairDoSistema'])
