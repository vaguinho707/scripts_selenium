import pandas as pd

dfAlcantara = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\ALCANTARA.xlsx')
dfBotafogo = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\BOTAFOGO.xlsx')
dfCaxias1 = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\CAXIAS.xlsx')
dfCaxias2 = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\CAXIAS 2.xlsx')
dfLeblon = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\LEBLON.xlsx')
dfMadureira = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\MADUREIRA.xlsx')
dfNiteroi = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\NITEROI.xlsx')
dfNovaIguacu = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\NOVA IGUACU.xlsx')
dfSaoGoncalo = pd.read_excel (r'C:\Users\Rafael\Documents\Atualizacao BD - Recebimento\SAO GONCALO.xlsx')

frames = [dfAlcantara, dfBotafogo, dfCaxias1, dfCaxias2, dfMadureira, dfLeblon, dfNiteroi, dfNovaIguacu, dfSaoGoncalo]
df = pd.concat(frames,ignore_index=True)


#adicionando as colunas de Tipo de Pagamento, Data do Pagamento, Clínica e Status do Pagamento
df['Unnamed: 19'] = ''
df['Unnamed: 20'] = ''
df['Unnamed: 21'] = ''
df['Unnamed: 22'] = ''
df['Unnamed: 23'] = ''


dictionary = df.to_dict()

#Mudando os tipos dos values das keys Unnamed 1,3,5,15,20,21 e 22 para string 
for celula in dictionary['Unnamed: 1']:
    dictionary['Unnamed: 1'][celula] = str(dictionary['Unnamed: 1'][celula])
    
for celula in dictionary['Unnamed: 3']:
    dictionary['Unnamed: 3'][celula] = str(dictionary['Unnamed: 3'][celula])

for celula in dictionary['Unnamed: 5']:
    dictionary['Unnamed: 5'][celula] = str(dictionary['Unnamed: 5'][celula])
    
for celula in dictionary['Unnamed: 15']:
    dictionary['Unnamed: 15'][celula] = str(dictionary['Unnamed: 15'][celula])

for celula in dictionary['Unnamed: 20']:
    dictionary['Unnamed: 20'][celula] = str(dictionary['Unnamed: 20'][celula])

for celula in dictionary['Unnamed: 21']:
    dictionary['Unnamed: 21'][celula] = str(dictionary['Unnamed: 21'][celula])

for celula in dictionary['Unnamed: 22']:
    dictionary['Unnamed: 22'][celula] = str(dictionary['Unnamed: 22'][celula])
    
 
    
#copiando o Nome da Clínica para a coluna referente

nomeClinica = dictionary['Unnamed: 3'][1]
#print(nomeClinica)
for celula in dictionary['Unnamed: 21']:
    if (dictionary['Unnamed: 3'][celula] == 'ALCÂNTARA') or (dictionary['Unnamed: 3'][celula] == 'BOTAFOGO') or (dictionary['Unnamed: 3'][celula] == 'CAXIAS') or (dictionary['Unnamed: 3'][celula] == 'CAXIAS 2') or (dictionary['Unnamed: 3'][celula] == 'MADUREIRA') or (dictionary['Unnamed: 3'][celula] == 'NITERÓI') or (dictionary['Unnamed: 3'][celula] == 'NOVA IGUAÇU') or (dictionary['Unnamed: 3'][celula] == 'LEBLON') or (dictionary['Unnamed: 3'][celula] == 'SÃO GONÇALO'):
        nomeClinica = dictionary['Unnamed: 3'][celula]
    dictionary['Unnamed: 21'][celula] = nomeClinica
    
#Copiando a Data do Pagamento para a coluna referente
dataPagamento = dictionary['Unnamed: 1'][5][-10:]

for celula in dictionary['Unnamed: 1']:
    if (dictionary['Unnamed: 1'][celula][:4] == 'Data'):
        dataPagamento = dictionary['Unnamed: 1'][celula][-10:]
    dictionary['Unnamed: 20'][celula] = dataPagamento

    

#Cálculo e preenchimento da coluna Status de Pagamento
    
for linha in dictionary['Unnamed: 20']:
    if ((dictionary['Unnamed: 20'][linha][3:5]) > (dictionary['Unnamed: 15'][linha][3:5])):
        dictionary['Unnamed: 22'][linha] = "RECUPERACAO"

    if ((dictionary['Unnamed: 20'][linha][3:5]) == (dictionary['Unnamed: 15'][linha][3:5])):
        dictionary['Unnamed: 22'][linha] = "ATUAL"

    if ((dictionary['Unnamed: 20'][linha][3:5]) < (dictionary['Unnamed: 15'][linha][3:5])):
        dictionary['Unnamed: 22'][linha] = "ANTECIPACAO"

    
#Copiando o nome tipo de pagamento para a coluna referente
auxIloc=7

for linha in dictionary['Unnamed: 1']:
    if (dictionary['Unnamed: 1'][linha] != "Cartão") and (dictionary['Unnamed: 1'][linha] != "Dinheiro") and (dictionary['Unnamed: 1'][linha]!= "Cheque"):
        if (dictionary['Unnamed: 1'][linha][:4] != "Total") and (dictionary['Unnamed: 1'][linha][:] != "") and (dictionary['Unnamed: 1'][linha][:2] != "nan"):
            dictionary['Unnamed: 19'][linha] = dictionary['Unnamed: 1'][auxIloc] 
    else:
        auxIloc=linha



#Transpondo a matriz para facilitar a manipulação
df=df.from_dict(dictionary)
dictionary = df.to_dict('index')


#Deletando as linhas inúteis do dicionário (criando um novo dicionário apenas com )
dicionarioFiltrado = {k:v 
    for k,v in dictionary.items()
        if not((v['Unnamed: 5'] == 'nan' or v['Unnamed: 5']=='Cliente'))}



#Transpondo a matriz de volta para a orientação inicial
df=df.from_dict(dicionarioFiltrado)
dicionarioFiltrado = df.to_dict('index')


#Deletando as colunas inúteis
del dicionarioFiltrado['Unnamed: 0'];
del dicionarioFiltrado['Unnamed: 2'];
del dicionarioFiltrado['Unnamed: 3'];
del dicionarioFiltrado['Unnamed: 4'];
del dicionarioFiltrado['Unnamed: 6'];
del dicionarioFiltrado['Unnamed: 7'];
del dicionarioFiltrado['Unnamed: 8'];
del dicionarioFiltrado['Unnamed: 9'];
del dicionarioFiltrado['Unnamed: 10'];
del dicionarioFiltrado['Unnamed: 12'];
del dicionarioFiltrado['Unnamed: 13'];
del dicionarioFiltrado['Unnamed: 14'];





#Renomeando os Labels das colunas
new_keys = ["Cod. Cliente","Cliente","Produto","Vencimento","Parcela","Valor Pago","Status","Tipo de Pagamento","Data do Pagamento","Clinica","Status do Pagamento","Quantificador"]

dicionarioFiltrado = dict(zip(new_keys,dicionarioFiltrado.values()))

df=df.from_dict(dicionarioFiltrado)


df.to_excel("Receitas.xlsx",index=False)




