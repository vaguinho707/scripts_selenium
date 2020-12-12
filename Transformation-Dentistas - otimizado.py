import pandas as pd

class Transformation:
#    def __init__(self):
#        dicionario = {}
    def escreverClinica(dictClinica, clinica):
        for item in dictClinica['Clinica']:
            dictClinica['Clinica'][item] = clinica
        return dictClinica

    def escreverDataMarcada(dicts):
        for dictClinica in dicts:
            dataMarcada = dictClinica['Unnamed: 1'][5][-10:]
            for celula in dictClinica['Unnamed: 1']:
                if (dictClinica['Unnamed: 1'][celula][:4] == 'Data'):
                    dataMarcada = dictClinica['Unnamed: 1'][celula][-10:]
                dictClinica['Data Marcada'][celula] = dataMarcada
        return dicts

    def escreverDentista(dicts):
        for dictClinica in dicts:
            dentista = dictClinica['Unnamed: 1'][8][10:]
            for celula in dictClinica['Unnamed: 1']:
                if (dictClinica['Unnamed: 1'][celula][:8] == 'Dentista'):
                    dentista = dictClinica['Unnamed: 1'][celula][10:]
                dictClinica['Dentista'][celula] = dentista
        return dicts

    def adicionarColunas(dataframe):
        for item in dataframe:
            item['Clinica'] = ''
            item['Data Marcada'] = ''
            item['Dentista'] = ''
        return dataframe

    def toStr(dicts, coluna):
        for dicionario in dicts:
            for celula in dicionario[coluna]:
                dicionario[coluna][celula] = str(dicionario[coluna][celula])
        return dicts

    def apagarLinhasInuteis(dicts):
        for dictClinica in range(0,9):
            dicts[dictClinica] = {k:v 
                for k,v in dicts[dictClinica].items()
                    if not((v['Unnamed: 9']) == 'nan'  or v['Unnamed: 9'] =='Início')}
        return dicts

    #Transpor "matriz" dicionario para facilitar manipulação.
    def transporMatrizDicionario(dicts,dfs):
        for aux in range(0,9):
            dfs[aux] = dfs[aux].from_dict(dicts[aux])
            dicts[aux] = dfs[aux].to_dict('index')
        return dicts


dfAlcantara = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\ALCANTARA.xlsx')
dfBotafogo = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\BOTAFOGO.xlsx')
dfCaxias1 = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\CAXIAS.xlsx')
dfCaxias2 = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\CAXIAS 2.xlsx')
dfLeblon = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\LEBLON.xlsx')
dfMadureira = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\MADUREIRA.xlsx')
dfNiteroi = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\NITEROI.xlsx')
dfNovaIguacu = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\NOVA IGUACU.xlsx')
dfSaoGoncalo = pd.read_excel (r'C:\Users\Rafael\Documents\Planilhas - Dentistas\SAO GONCALO.xlsx')

dfs = [dfAlcantara, dfBotafogo, dfCaxias1, dfCaxias2, dfLeblon, dfMadureira, dfNiteroi, dfNovaIguacu, dfSaoGoncalo]


dfs = Transformation.adicionarColunas(dfs)


dictAlcantara = dfAlcantara.to_dict()
dictBotafogo = dfBotafogo.to_dict()
dictCaxias1 = dfCaxias1.to_dict()
dictCaxias2 = dfCaxias2.to_dict()
dictLeblon = dfLeblon.to_dict()
dictMadureira = dfMadureira.to_dict()
dictNiteroi = dfNiteroi.to_dict()
dictNovaIguacu = dfNovaIguacu.to_dict()
dictSaoGoncalo = dfSaoGoncalo.to_dict()

dicts = [dictAlcantara, dictBotafogo, dictCaxias1, dictCaxias2, dictLeblon, dictMadureira, dictNiteroi, dictNovaIguacu, dictSaoGoncalo]

dicts = Transformation.toStr(dicts, 'Unnamed: 1')
dicts = Transformation.toStr(dicts, 'Unnamed: 9')


dicts[0] = Transformation.escreverClinica(dictAlcantara, 'ALCANTARA')
dicts[1] = Transformation.escreverClinica(dictBotafogo, 'BOTAFOGO')
dicts[2] = Transformation.escreverClinica(dictCaxias1, 'CAXIAS')
dicts[3] = Transformation.escreverClinica(dictCaxias2, 'CAXIAS 2')
dicts[4] = Transformation.escreverClinica(dictLeblon, 'LEBLON')
dicts[5] = Transformation.escreverClinica(dictMadureira, 'MADUREIRA')
dicts[6] = Transformation.escreverClinica(dictNiteroi, 'NITEROI')
dicts[7] = Transformation.escreverClinica(dictNovaIguacu, 'NOVA IGUACU')
dicts[8] = Transformation.escreverClinica(dictSaoGoncalo, 'SAO GONCALO')

#Copiando a Data Marcada para a coluna referente
dicts = Transformation.escreverDataMarcada(dicts)

#Copiando a Data Marcada para a coluna referente
dicts = Transformation.escreverDentista(dicts)

#Transpondo a matriz para facilitar a manipulação
dicts = Transformation.transporMatrizDicionario(dicts,dfs)


#Apagando linhas que não têm serventia.
dicts = Transformation.apagarLinhasInuteis(dicts)


#Transpondo a matriz de volta para a orientação inicial
dicts = Transformation.transporMatrizDicionario(dicts,dfs)


dfAlcantara = dfAlcantara.from_dict(dicts[0])
dfBotafogo = dfBotafogo.from_dict(dicts[1])
dfCaxias1 = dfCaxias1.from_dict(dicts[2])
dfCaxias2 = dfCaxias2.from_dict(dicts[3])
dfLeblon = dfLeblon.from_dict(dicts[4])
dfMadureira = dfMadureira.from_dict(dicts[5])
dfNiteroi = dfNiteroi.from_dict(dicts[6])
dfNovaIguacu = dfNovaIguacu.from_dict(dicts[7])
dfSaoGoncalo = dfSaoGoncalo.from_dict(dicts[8])

dfs = [dfAlcantara, dfBotafogo, dfCaxias1, dfCaxias2, dfLeblon, dfMadureira, dfNiteroi, dfNovaIguacu, dfSaoGoncalo]

df = pd.concat(dfs, ignore_index=True)

dictCompleto = df.to_dict()
dictCompleto.pop('Unnamed: 0')
dictCompleto.pop('Unnamed: 2')
dictCompleto.pop('Unnamed: 4')
dictCompleto.pop('Unnamed: 5')
dictCompleto.pop('Unnamed: 6')
dictCompleto.pop('Unnamed: 7')
dictCompleto.pop('Unnamed: 8')
dictCompleto.pop('Unnamed: 10')
dictCompleto.pop('Unnamed: 12')

#Renomeando os Labels das colunas
new_keys = ["Cod. Cliente","Cliente","Início","Nome do Procedimento","Clinica","Data Marcada","Dentista"]

dictCompleto = dict(zip(new_keys,dictCompleto.values()))

df = df.from_dict(dictCompleto)


df.to_excel(r"C:\Users\Rafael\Planilha - Dentistas.xlsx",index=False)
