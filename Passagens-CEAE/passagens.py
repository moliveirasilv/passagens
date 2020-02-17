#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Nov  6 20:14:38 2019

@author: math
"""

import os 
import pandas as pd

def get_data(diretorio,tab):
    lista_dados_volta = []
    os.chdir(diretorio)
    for f in os.listdir():
        try:
            df = pd.read_excel(f, tab, header=None)
            df.dropna(how='all', inplace=True)
            df.reset_index(inplace=True)
            df.drop(columns=['index'], inplace=True)
            lista_dados_volta.append(df)
        except:
            print('Error')
    return lista_dados_volta

def gross_data(diretorio,tab):
    start_index = [i for i in range(0,78,11)]
    end_index = [i for i in range(5,78,11)]
    index = zip(start_index,end_index)
    dados = get_data(diretorio,tab)
    dfs = []
    
    for s,e in index:
        for i in range(0,len(dados)):
            df = dados[i].loc[s:e]
            dfs.append(df)                
    for df in dfs:
        df.reset_index(inplace=True)
           
    return dfs


def dicti(diretorio,tab): 
    dfs = gross_data(diretorio,tab)
    dados = []
    for i in range(0,len(dfs)):
        dicti = {}
        dicti['Mundi']= {dfs[i][5][0]:{dfs[i][10][0]:{dfs[i][1][0]: {'TAM':{'M':dfs[i][1][3],'T':dfs[i][2][3],'N':dfs[i][3][3]},
                           'GOL':{'M':dfs[i][1][4],'T':dfs[i][2][4],'N':dfs[i][3][4]},
                           'AZUL':{'M':dfs[i][1][5],'T':dfs[i][2][5],'N':dfs[i][3][5]}
                           }}}}
        dicti['Decolar'] = {dfs[i][5][0]:{dfs[i][10][0]:{ dfs[i][1][0]: {'TAM':{'M':dfs[i][4][3],'T':dfs[i][5][3],'N':dfs[i][6][3]},
                              'GOL':{'M':dfs[i][4][4],'T':dfs[i][5][4],'N':dfs[i][6][4]},
                              'AZUL':{'M':dfs[i][4][5],'T':dfs[i][5][4],'N':dfs[i][6][4]}
                              }}}}
        dicti['ViajaNet'] = {dfs[i][5][0]:{dfs[i][10][0]:{dfs[i][1][0]: { 'TAM':{'M':dfs[i][7][3],'T':dfs[i][8][3],'N':dfs[i][9][3]},
                               'GOL':{'M':dfs[i][7][4],'T':dfs[i][8][4],'N':dfs[i][9][4]},
                               'AZUL':{'M':dfs[i][7][5],'T':dfs[i][8][5],'N':dfs[i][9][5]}
                              }}}}
        dicti['Submarino Viagens'] = {dfs[i][5][0]:{dfs[i][10][0]:{dfs[i][1][0]: {'TAM':{'M':dfs[i][10][3],'T':dfs[i][11][3],'N':dfs[i][12][3]},
                                       'GOL':{'M':dfs[i][10][4],'T':dfs[i][11][4],'N':dfs[i][12][4]},
                                       'AZUL':{'M':dfs[i][10][5],'T':dfs[i][11][5],'N':dfs[i][12][5]}
                                  }}}}
        dados.append(dicti)    
    return dados


def dados_passagens(diretorio,tab):
    """diretorio: o dito cujo, tab: a planilha no arquivo xlsx"""
    dados = pd.DataFrame(dicti(diretorio,tab)) 
    d = {(i,j,k,l,m,n,o): dados[i][j][k][l][m][n][o]
           for i in dados.keys() 
           for j in dados[i].keys()
           for k in dados[i][j].keys()
           for l in dados[i][j][k].keys()
           for m in dados[i][j][k][l].keys()
           for n in dados[i][j][k][l][m].keys()
           for o in dados[i][j][k][l][m][n].keys()
                                                   }    
    
    mux = pd.MultiIndex.from_tuples(d.keys())
    multi_in = pd.DataFrame(list(d.values()), index=mux)
    raw_data  = multi_in.droplevel(1) # dataframe multiindexado,
    
    passagens = raw_data.reset_index()
    passagens.columns = ['Sites','Data da Coleta','Data da Viagem','Destino','Empresa','Turno','Precos']
    return passagens
    
