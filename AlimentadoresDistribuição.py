import pandas as pd 
import numpy as np 
import matplotlib.pyplot as plt 
from random import sample
from scipy import stats



# Como o índice das cargas ligadas em cada ramal não apresenta uma ordem definida
# primeiramente,foi indicado para cada ponto do ramal,quais os indices de suas cargas 

consumidores = [[39,26,38,24,15,14,32,43,42,29,25,20,6,9,40,23,34,19],
            [28,10,5,31,41],
            [12,4,17,27,11,16,3,7,22,45,30,8,2,18],    
            [35,33,21,37,44,36,13,46],
            [47,48,49,50,51,52,53],
            [54],
            [55,56,57,58]
            ]            


num_ramais= len(consumidores)
ramais=dict() # guarda os dados dos consumidores dentro de cada ramal e os dados do sistema
ramais2=dict() #guarda o  dados de cada ramal proprimamente dito 
for i in range(0,num_ramais):
    
    ramais['Ramal 0{}'.format(i+1)]={'consumidores no ramal':consumidores[i]}
    
############# Agora vamos importar os dados da tabela csv para cada um dos ramais da rede ##############

dados = pd.read_csv('data.csv').T.to_numpy() # Tive que transpor os dados, senão seria uma matriz pra 
                                             # cada horário, pegando todos os consumidores
                                             
carga_instalada=pd.read_csv('potencias.csv').T.to_numpy() # Nessa planilha foi colocada uma carga instalada
                                                          # fantasma de 1000 no indice 0, isso foi feito 
                                                          # somente para os indices ficarem igualados com
                                                          # a das outras 

carga_instalada=carga_instalada[1]


######################## Dando para o algoritmo outros dados importantes ######################

num_consumidores=len(dados) - 1     # -1 pq um será o indice das horas 
 
num_medidas=len(dados[0])

fp=0.9                              #Fator de potência de todas as cargas = 0.9 tal como pedido 

P_trafo = 630                       #[kVA] --> Potência do trafo 

delta_t = 0.25                          #a cada um quarto de uma hora tem-se uma nova medida
tempo_medicao = num_medidas*delta_t

Demanda_sistema=np.zeros(num_medidas)
Dmnc_sistema = 0 #Demanda máxima não coincidente 
CarInstSistema=0
ConsumoSistema=0

# Aqui faz-se o calculo das grandezas referentes aos consumidores
for j in range(0,len(ramais)):
    
    writer = pd.ExcelWriter('Ramal 0{}.xlsx'.format(j+1),engine = 'openpyxl')
    
    Demanda_ramal = np.zeros(num_medidas) # ---->  Essa grandeza será usada para se calcular a 
                                          #demanda diversificada de cada ramal  
                                          
    Dmnc=0                                # ----> Essa grandeza será usada a demanda máxima não
                                          # não coincidente para cada ramal
                                          
    CarInstRama=0    
    
    ConsumoRamal=0                                  
    for i in consumidores[j]:
            
            # Pegando a curva de carga em cada consumidor:
            ramais['Ramal 0{}'.format(j+1)].update({'consumidor {}'.format(i):{'Curva de carga':dados[i]
            }})
            
            # Pegando o número de consumidores:
            ramais['Ramal 0{}'.format(j+1)]['N Consumidores'] = len(consumidores[j])
            
            # Pegando a potência instalada em cada consumidor:
            ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Potencia instalada']=carga_instalada[i]
            
            # Pegando a demanda máxima: 
            ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda maxima'] = max(ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Curva de carga'])
            
            #Pegando a Demanda máxima diária de cada consumidor:
                
            a=0
            cdc=ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Curva de carga']
            for t in range(0,7): # foi medido durante 7 dias, por isso range(0,7)
               ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda maxima Dia {}'.format(a+1)]=max(cdc[a*96:(a*96)+96]) # Todo dia tem-se 96 medições diferentes 
               ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Instante Demanda maxima Dia {}'.format(a+1)]=np.argmax(cdc[a*96:(a*96)+96])
               a+=1
            
            # Calculo do fator de demanda:
            if ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Potencia instalada'] == 0:
                
                ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Fator de Demanda'] = 0
            else:
                
                ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Fator de Demanda'] = ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda maxima']/(ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Potencia instalada'])

            #Essa parte é pra checar se tem defeitos nos dados, e retornou que tem
                #if ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Fator de Demanda'] > 1 :
                    #print(i)
            
            # Para calcularmos o fator de carga, primeiramente calcula-se a demanda média:
                
            # Primeiramente, calcula-se o consumo total de cada consumidor em [kWh]:                
            ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Consumo'] = delta_t*sum(ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Curva de carga'])

            # Depois, a demanda média em [kW]:
            ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda Media']=ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Consumo']/tempo_medicao  
                
            # Agora, pode-se calcular o fator de carga de cada consumidor:
                
            if ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda Media'] == 0:
                
                ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Fator de Carga'] = 0
                
            else:
                
                ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Fator de Carga']= ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda Media']/ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda maxima']
            
            #No final do for, teremos a curva de demanda de todos os ramais: 
            Demanda_ramal += ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Curva de carga']
            
            #Além disso, atualiza-se também a demanda máxima não coincidente:
            Dmnc += ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Demanda maxima']
            
            #Calculo da Carga Instalada no Ramal:
            CarInstRama+=ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Potencia instalada']
    
            #Calculo do consumo total em cada Ramal:
            ConsumoRamal+=ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)]['Consumo']
    
            
            # Salvando o consumidor na planilha do ramal:
            salva = pd.DataFrame(ramais['Ramal 0{}'.format(j+1)]['consumidor {}'.format(i)])
            salva.to_excel(writer, sheet_name = 'consumidor {}'.format(i+1))         
    
    #Agora colocaremos a curva de demanda dos ramais dentro de cada dicionário 
    ramais2.update({'Ramal 0{}'.format(j+1):{'Demanda Diversificada': Demanda_ramal}})
    
    #Depois, utiliza-se a mesma variável para calcular a Demanda do sistema:
    Demanda_sistema += Demanda_ramal 
    
    #Agora calcularemos a Demanda máxima diversificada de cada um dos ramais:
    ramais2['Ramal 0{}'.format(j+1)]['Demanda Maxima Diversificada'] = max(ramais2['Ramal 0{}'.format(j+1)]['Demanda Diversificada'])
            
    #E então aloca-se a Demanda máxima não coincidente de cada um dos ramais:
    ramais2['Ramal 0{}'.format(j+1)]['Demanda Máxima Não Coincidente'] = Dmnc
    
    #Calculo do fator de diversidade:
    ramais2['Ramal 0{}'.format(j+1)]['Fator de Diversidade'] = ramais2['Ramal 0{}'.format(j+1)]['Demanda Máxima Não Coincidente']/ramais2['Ramal 0{}'.format(j+1)]['Demanda Maxima Diversificada']
    
    #Calculo do Fator de Demanda dos Ramais:
    ramais2['Ramal 0{}'.format(j+1)]['Fator de Demanda']= ramais2['Ramal 0{}'.format(j+1)]['Demanda Maxima Diversificada']/CarInstRama
    
    #Depois, utiliza-se a mesma variável para calcular a Demanda Máxima não coincidente do sistema:
    Dmnc_sistema += Dmnc    
    
    #Consumo Total do Ramal em [kWh]:
    ramais2['Ramal 0{}'.format(j+1)]['Consumo']= ConsumoRamal
    
    #Consumo do Sistema:
    ConsumoSistema += ConsumoRamal
    
    #Calculando a Demanda Média:
    ramais2['Ramal 0{}'.format(j+1)]['Demanda Media'] = ramais2['Ramal 0{}'.format(j+1)]['Consumo']/tempo_medicao 

    #Calculando o Fator de Carga de Cada Ramal:
    ramais2['Ramal 0{}'.format(j+1)]['Fator de Carga'] = ramais2['Ramal 0{}'.format(j+1)]['Demanda Media']/ramais2['Ramal 0{}'.format(j+1)]['Demanda Maxima Diversificada']   
    
    # Carga Instalada do sistema:
    CarInstSistema+=CarInstRama
        
    #Potando os gráficos de Curva de Carga de Cada Ramal
    plt.xticks(rotation=90)
    plt.plot(ramais2['Ramal 0{}'.format(j+1)]['Demanda Diversificada'],label = 'Ramal 0{}'.format(j+1),color = 'c')
    plt.title("Curva de carga do Ramal 0{}".format(j+1))
    plt.xlabel("Hora")
    plt.ylabel("Carga [kW]")
    plt.show()
    
    writer.save() 
            
#Agora colocaremos a curva de demanda do sistema dentro do dicionário:
ramais.update({'Sistema':{'Demanda Diversificada': Demanda_sistema}})          
              
#Plotando o gráfico de curva de carga do sistema: 
plt.xticks(rotation=90)
plt.plot(ramais['Sistema']['Demanda Diversificada'],label = 'Sistema',color = 'r')
plt.title("Curva de carga do Sistema")
plt.xlabel("Hora")
plt.ylabel("Carga [kW]")
plt.show()

# Agora calcularemos a Demanda Máxima diversificada do sistema:
ramais['Sistema']['Demanda Máxima Diversificada']=max(ramais['Sistema']['Demanda Diversificada'])
    
# Alocando a demanda máxima não coincidente:
ramais['Sistema']['Demanda Máxima Não Coincidente'] =Dmnc_sistema

# Fator de diversidade do sistema:
ramais['Sistema']['Fator de Diversidade']=ramais['Sistema']['Demanda Máxima Não Coincidente']/ramais['Sistema']['Demanda Máxima Diversificada']

#Fator de Demanda:
ramais['Sistema']['Fator de Demanda'] =  ramais['Sistema']['Demanda Máxima Diversificada'] /CarInstSistema

#Consumo do Sistema:
ramais['Sistema']['Consumo'] = ConsumoSistema

#Demanda Média:
ramais['Sistema']['Demanda Media'] = ramais['Sistema']['Consumo']/tempo_medicao

#Fator de Carga:
ramais['Sistema']['Fator de Carga'] = ramais['Sistema']['Demanda Media']/ramais['Sistema']['Demanda Máxima Diversificada'] 

writer2 = pd.ExcelWriter('Sistema.xlsx',engine = 'openpyxl')

dados1=pd.DataFrame(ramais['Sistema'])
dados1.to_excel(writer2,sheet_name='Sistema')
writer2.save()

#Tbm irei salvar os dados de cada ramal dentro da planilha sistema
for i in range(0,len(consumidores)):
    dadosn=pd.DataFrame(ramais2['Ramal 0{}'.format(i+1)])
    dadosn.to_excel(writer2,sheet_name='Ramal 0{}'.format(i+1))
    writer2.save()            

################### Dimensionando a demanda total do alimentador ########################


# Nessa parte, foram pegas 15 combinações diferentes para cada um dos itens
# da tabela, calculada, para cada combinação o fator de diversidade, a partir
# desses valores foi feita uma média

# Primeiramente,pegamos o maior ramal da lista consumidores, para, a partir dele
# Fazermos a previsão de carga dos outros ramais

#Pega o indice do maior ramal
maior = np.argmax(len(consumidores[i]) for i in range(0,len(consumidores))) 


class Escolha: # Feito para sempre escolher permutações diferentes
    
    def __init__(self,permut,lista_permut):
        self.escolher_amostras(permut,lista_permut)
    
    def escolher_amostras(self,permut,lista_permut):
        amostra =sample(consumidores[maior],permut)
        if amostra in lista_permut:
            self.escolher_amostras(permut,lista_permut)
        else:
            lista_permut.append(amostra)
            return amostra


tabela = dict() #Essa será a tabela de valores diversificados 
n_amostras=15
permut=2 #Primeiramente, pegaremos combinações de 2

#Nesse while  não colocamos para a tabela nem permutações de um nem permutações de
#todos(que é só uma)

while permut < len(consumidores[maior]):
    lista_permut = list()
    Fator = np.zeros(n_amostras)
    for i in range(0,n_amostras):
        Escolha(permut,lista_permut)
        amostra=lista_permut[-1]
        Ddiver = 0 # Curva de Demanda Diversificada
        Dmnc = 0   # Demanda máxima não coincidente 
        for k in amostra:
            Ddiver += ramais['Ramal 0{}'.format(maior +1)]['consumidor {}'.format(k)]['Curva de carga']
            Dmnc+=ramais['Ramal 0{}'.format(maior+1)]['consumidor {}'.format(k)]['Demanda maxima']
        Ddiver_max=max(Ddiver)
        Fator[i] = Dmnc/Ddiver_max
    tabela.update({permut:round(sum(Fator)/n_amostras,4)}) # Coloca a média das amostras no dicionário 
    permut+=1
    
#Colocando na Tabela indice 1 e o indice len(consumidores)
tabela.update({1:1})
tabela.update({len(consumidores[0]):ramais2['Ramal 0{}'.format(maior+1)]['Fator de Diversidade']})

#Plotando o gráfico:

plote=[tabela[i] for i in range(1,len(consumidores[maior])+1)] 
    
plt.xticks(rotation=90)
plt.plot(plote,label= 'Fator de Diversidade',color = 'r')
plt.title("Fator de Diversidade Calculado")
plt.xlabel("Número de Consumidores")
plt.ylabel("Fator")
plt.show()


# Comparando com o resultado obtido em cada ramal 

erro_abs= list() # Erro absuluto 
erro_rel=list() # Erro relativo
De_max_diver=list()

# Lista com os argumentos dos ramais
argumentos = [i for i in range(0,len(consumidores))]

#Retirando o maior para depois podermos comparar com os outros:
argumentos.remove(maior)

for i in argumentos: # Consumidores =  Lista onde tem os ramais 
    tamanho = ramais['Ramal 0{}'.format(i+1)]['N Consumidores']
    Diversidade_Ramal = tabela[tamanho]
    #Calculando a Demanda máxima diversificada usando a tabela 
    De_max_diver.append(ramais2['Ramal 0{}'.format(i+1)]['Demanda Máxima Não Coincidente']/Diversidade_Ramal)
    
    #Comparando o valor calculado com o real medido, e colocando no gráfico pra 
    erro_abs.append(abs(ramais2['Ramal 0{}'.format(i+1)]['Demanda Maxima Diversificada'] - De_max_diver[-1]))

    erro_rel.append(erro_abs[-1]/ramais2['Ramal 0{}'.format(i+1)]['Demanda Maxima Diversificada'])
 

erro_abs=np.array(erro_abs)
erro_rel=np.array(erro_rel)
    
# Plotando o gráfico da demanda diversificada calculada vs real
plt.plot([i+1 for i in argumentos],De_max_diver,'o',label= 'D',color = 'r')
plt.plot([i+1 for i in argumentos],[ramais2['Ramal 0{}'.format(i+1)]['Demanda Maxima Diversificada'] for i in argumentos],'o',label= 'D',color = 'b')
plt.title("Demanda Diversificada Calculada e Real ")
plt.xlabel("Número do Ramal")
plt.ylabel("Demanda")
plt.show()

# Plotando o gráfico do erro absoluto 
plt.plot([i+1 for i in argumentos],erro_abs,'o',label= 'Erro Fator de Diversidade',color = 'r')
plt.title("Erro Fator de Diversidade Calculado")
plt.xlabel("Número do Ramal")
plt.ylabel("Erro Absoluto")
plt.show()

# Plotando o gráfico do erro relativo 
plt.plot([i+1 for i in argumentos],erro_rel,'o',label= 'Erro Fator de Diversidade',color = 'r')
plt.title("Erro Fator de Diversidade Calculado")
plt.xlabel("Número do Ramal")
plt.ylabel("Erro Absoluto")
plt.show()

EQM = (sum((erro_abs)**2)/len(erro_abs))**1/2


############### Agora dimensionando consumidor a consumidor pela pesquisa de carga ############

#Para essa regressão linear, também usaremos o maior ramal

#Lista com os dados de X (consumo em kWh) e Y (demanda máxima em kW) para a regressão linear

eixo_x = list()
eixo_y= list()


# Lista com os dados de X (consumo em kWh) e Y (demanda máxima em kW) para a regressão linear
reg_consumo = []
reg_demanda = []

#Salvará os erros quadraticos médios
eqms=dict()

for i in consumidores[maior]:

    eixo_x.append(ramais['Ramal 0{}'.format(maior+1)]['consumidor {}'.format(i)]['Consumo'])
    eixo_y.append(ramais['Ramal 0{}'.format(maior+1)]['consumidor {}'.format(i)]['Demanda maxima'])

eixo_x=np.array(eixo_x)
eixo_y=np.array(eixo_y)

# Fazendo a regressão linear
regressao=stats.linregress(eixo_x,eixo_y)  

# Plotando os gráficos referentes a regressão linear:

plt.plot(eixo_x, eixo_y, 'o', label='Dados medidos')
plt.plot(eixo_x,regressao.intercept + regressao.slope*eixo_x,'r', label='Regressao')
plt.legend()
plt.title("Regressão feita no Ramal 0{}".format(maior+1))
plt.xlabel("Consumo [kWh]")
plt.ylabel("Demanda Máxima [kW]")
plt.show()    

#Obtenção dos valores de consumo e da demanda prevista para os outros ramais

for i in argumentos:
    print("tira o x")
    
    eixo_x_estudado = list()
    eixo_y_estudado = list()
    
    for j in consumidores[i]:
        eixo_x_estudado.append(ramais['Ramal 0{}'.format(i+1)]['consumidor {}'.format(j)]['Consumo'])
        eixo_y_estudado.append(ramais['Ramal 0{}'.format(i+1)]['consumidor {}'.format(j)]['Demanda maxima'])
    
    eixo_x_estudado=np.array(eixo_x_estudado)
    eixo_y_estudado=np.array(eixo_y_estudado)
    
    demanda_prevista = eixo_x_estudado*regressao.slope + regressao.intercept
    
    plt.plot(eixo_x_estudado, demanda_prevista, 'o', label='Demanda prevista')
    plt.plot(eixo_x_estudado, eixo_y_estudado, 'o', label='Dados medidos')
    plt.legend()
    plt.title("Previsão de Demanda - Consumidores do Ramal 02")
    plt.xlabel("Consumo [kWh]")
    plt.ylabel("Demanda Máxima [kW]")
    plt.show()
        

    erro_abs = abs(eixo_y_estudado - demanda_prevista) #Erro Absoluto
    erro_rel=erro_abs/demanda_prevista #Erro relativo
    erro_abs = np.array(erro_abs)
    erro_rel=np.array(erro_rel)
    #vai sobre escrever os eqm
    eqm =  (sum((erro_abs)**2)/len(consumidores[i]))**1/2
    
    plt.plot(consumidores[i],erro_abs,'o')
    plt.title("Erro Absoluto entre Demanda Real e Estimada do Ramal 0{}".format(i+1))
    plt.xlabel("Índice do Consumidor")
    plt.ylabel("Erro")
    plt.show()
    
    #plt.xticks(xgenerico, consumidores[1])
    plt.plot(consumidores[i],erro_rel,'o')
    plt.title("Erro Relativo entre Demanda Real e Estimada do Ramal 0{}".format(i+1))
    plt.xlabel("Índice do Consumidor")
    plt.ylabel("Erro")
    plt.show()
    
    eqms['Ramal 0{}'.format(i+1)]= eqm
