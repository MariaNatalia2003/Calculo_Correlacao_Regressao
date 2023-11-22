#instalação das bibliotecas
# pip install numpy matplotlib
# pip install pandas matplotlib
# pip install openpyxl

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def calcula_correlacao_regressao_dados_excel(arquivo_excel, coluna_x, coluna_y):
    # Carrega os dados do Excel usando o pandas
    dados = pd.read_excel(arquivo_excel)
    #print(dados.columns)
    
    # Obtém os valores das colunas especificadas
    x = dados[coluna_x].values
    y = dados[coluna_y].values
    
    # Calcula o Coeficiente de Correlação de Pearson
    correlacao = np.corrcoef(x, y)[0, 1]
    
    # Calcula a Reta de Regressão
    coeficientes_regressao = np.polyfit(x, y, 1)
    a, b = coeficientes_regressao
    
    # Cria a Reta de Regressão
    linha_regressao = np.polyval(coeficientes_regressao, x)
    
    # Plotagem dos dados e da Reta de Regressão
    plt.scatter(x, y, label='Dados')
    plt.plot(x, linha_regressao, color='red', label=f'Reta de Regressão: y = {a:.4f}x + {b:.4f}')
    plt.xlabel(coluna_x)
    plt.ylabel(coluna_y)
    plt.legend()
    plt.title(f'Coeficiente de Correlação: {correlacao:.4f}')
    plt.show()

# Uso com um arquivo Excel chamado 'dados.xlsx', onde as colunas de interesse são 'X' e 'Y'
#calcula_correlacao_regressao_dados_excel('dados.xlsx', 'X', 'Y')
#calcula_correlacao_regressao_dados_excel('dados2.xlsx', 'Consumo de energia (kWh) ', 'Demanda de pico (kW)')
calcula_correlacao_regressao_dados_excel('dados3.xlsx', 'Peso', 'Altura')