# -*- coding: utf-8 -*-
"""
Created on Tue Nov  7 15:28:48 2023

@author: 2018459
"""


#%% Biblioteca
import pandas as pd
from xml.dom import minidom
import glob
from decimal import Decimal


#%% Diretório
# arquivo_xml = r'C:\Users\2018459\OneDrive - CPFL Energia S A\Documentos\Projetos\2023\Anexo 1 - Tratamento das Reclamações\Dados\2023\Paulista\Envio\APLSCR00063_SAC_012023_S001 - Copia.xml'
# Leitura do arquivo XML
# file = minidom.parse(arquivo_xml)

# Arquivo de apoio
arquivo_apoio = r'C:\Users\2018459\OneDrive - CPFL Energia S A\Documentos\Projetos\2023\Anexo 1 - Tratamento das Reclamações\Dados\Apoio_tipologia.xlsx'

# Lista para guardar os elementos das tag
municipios = []
codigo_forma_contato = []
codigo_tipologia = []
qtde_recebidas = []
qtde_procedentes = []
qtde_improcedentes = []
prazo_procedentes = []
prazo_improcedentes = []

# Variaveis
lista_distribuidora = []
lista_arquivo = []
lista_erro = []
lista_codigo = []
lista_linha = []

# Tabela para exportar o log de erros
df_log_erros = pd.DataFrame(columns=['distribuidora','arquivo','erro','valor','linha'])


#%% Leitura
# Leitura do arquivo de apoio com as tipologias
df_tipologia = pd.read_excel(arquivo_apoio,sheet_name = 'Tipologia',usecols='C:D',dtype='str')
df_forma_contato = pd.read_excel(arquivo_apoio,sheet_name = 'Forma de Contato',dtype='str')
df_municipios = pd.read_excel(arquivo_apoio,sheet_name = 'Municipios',dtype='str',usecols = [6,7])
df_municipios_paulista = df_municipios[df_municipios['DISTRIBUIDORA'] == '63']
df_municipios_piratininga = df_municipios[df_municipios['DISTRIBUIDORA'] == '2937']
df_municipios_santacruz = df_municipios[df_municipios['DISTRIBUIDORA'] == '69']
df_municipios_rge = df_municipios[df_municipios['DISTRIBUIDORA'] == '396']

# Separação do arquivo por tag
def separa_tag(file):
    global municipio_tag, tipologia_tag, qtde_recebidas_tag, qtde_procedentes_tag, qtde_improcedentes_tag, prazo_procedentes_tag, prazo_improcedentes_tag
    municipio_tag = file.getElementsByTagName('Municipio')
    tipologia_tag = file.getElementsByTagName('Tipologia')
    qtde_recebidas_tag = file.getElementsByTagName('Quantidade_recebidas')
    qtde_procedentes_tag = file.getElementsByTagName('Quantidade_procedentes')
    qtde_improcedentes_tag = file.getElementsByTagName('Quantidade_improcedentes')
    prazo_procedentes_tag = file.getElementsByTagName('Prazo_tratamento_procedentes')
    prazo_improcedentes_tag = file.getElementsByTagName('Prazo_tratamento_improcedentes')


# Seleciona os elementos dentro da tag e adiciona na lista
def seleciona_elementos(municipio_tag, tipologia_tag, qtde_recebidas_tag, qtde_procedentes_tag, qtde_improcedentes_tag, prazo_procedentes_tag, prazo_improcedentes_tag,distribuidora,arquivo,lista_distribuidora,lista_arquivo,lista_erro,lista_codigo,lista_linha):
    i = 0
    # Municipio
    for elem in municipio_tag:
        municipios.append(elem.attributes['CODIGO_MUNICIPIO'].value)

    
    # Tipologia
    for elem in tipologia_tag:
        codigo_forma_contato.append(elem.attributes['CODIGO_FORMA_CONTATO'].value)
        codigo_tipologia.append(elem.attributes['CODIGO_TIPOLOGIA'].value)


    #Qtde Recebida
    for elem in qtde_recebidas_tag:     
        try:
            qtde_recebidas.append(elem.firstChild.data)
            i += 1    
        except:
            print('Valor nulo na linha:',i, elem)
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Valor nulo Qtde Recebida')
            lista_codigo.append('Valor nulo')
            lista_linha.append(i+1)
    i = 0
      
    #Qtde Procedente
    for elem in qtde_procedentes_tag:
        try:
            qtde_procedentes.append(elem.firstChild.data)
            i += 1
        except:
            print('Valor nulo na linha:',i, elem)
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Valor nulo Qtde Procedente')
            lista_codigo.append('Valor nulo')
            lista_linha.append(i+1)
    i = 0
        
    #Qtde Improcedente
    for elem in qtde_improcedentes_tag:
        try:
            qtde_improcedentes.append(elem.firstChild.data)
            i += 1
        except:
            print('Valor nulo na linha:',i, elem)
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Valor nulo Qtde Improcedente')
            lista_codigo.append('Valor nulo')
            lista_linha.append(i+1)
    i = 0
        
    #Prazo Procedente
    for elem in prazo_procedentes_tag:
        try:
            prazo_procedentes.append(elem.firstChild.data)
            i += 1
        except:
            print('Valor nulo na linha:',i, elem)
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Valor nulo Prazo Procedente')
            lista_codigo.append('Valor nulo')
            lista_linha.append(i+1)
    i = 0
        
    #Prazo Improcedente
    for elem in prazo_improcedentes_tag:
        try:        
            prazo_improcedentes.append(elem.firstChild.data)
            i += 1
        except:
            print('Valor nulo na linha:',i, elem)
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Valor nulo Prazo Improcedente')
            lista_codigo.append('Valor nulo')
            lista_linha.append(i+1)
    i = 0
        
        

#%% Validações
# Tipologia
def valida_tipologia(codigo_tipologia,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo):
    for index,tipologia in enumerate(codigo_tipologia):
        if tipologia not in list(df_tipologia['CÓDIGO']):
            print('Tipologia inválida:',tipologia, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Tipologia')
            lista_codigo.append(tipologia)
            lista_linha.append(index+1)
  
            
# Código Forma de Contato
def valida_forma_contato(codigo_forma_contato,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo):
    for index,forma_contato in enumerate(codigo_forma_contato):
        if forma_contato not in list(df_forma_contato['CÓDIGO']):
            print('Código Forma de Contato inválida:',forma_contato, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Forma de Contato')
            lista_codigo.append(forma_contato)
            lista_linha.append(index+1)
    
        
# Municípios
def valida_municipio(municipios,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo):
    if distribuidora == 'Paulista':
        for index,municipio in enumerate(municipios):
            if municipio not in list(df_municipios_paulista['MUNICIPIO']):
                print('Município fora da área de concessão da Paulista:',municipio, 'Linha:',index+1)
                
                lista_distribuidora.append(distribuidora)
                lista_arquivo.append(arquivo.split('\\')[-1])
                lista_erro.append('Municipio')
                lista_codigo.append(municipio)
                lista_linha.append(index+1)
                
    elif distribuidora == 'Piratininga':
        for index,municipio in enumerate(municipios):
            if municipio not in list(df_municipios_piratininga['MUNICIPIO']):
                print('Município fora da área de concessão da Piratininga:',municipio, 'Linha:',index+1)
                
                lista_distribuidora.append(distribuidora)
                lista_arquivo.append(arquivo.split('\\')[-1])
                lista_erro.append('Municipio')
                lista_codigo.append(municipio)
                lista_linha.append(index+1)
                
    elif distribuidora == 'Santa Cruz':
        for index,municipio in enumerate(municipios):
            if municipio not in list(df_municipios_santacruz['MUNICIPIO']):
                print('Município fora da área de concessão da Santa Cruz:',municipio, 'Linha:',index+1)
                
                lista_distribuidora.append(distribuidora)
                lista_arquivo.append(arquivo.split('\\')[-1])
                lista_erro.append('Municipio')
                lista_codigo.append(municipio)
                lista_linha.append(index+1)
                
    elif distribuidora == 'RGE':
        for index,municipio in enumerate(municipios):
            if municipio not in list(df_municipios_rge['MUNICIPIO']):
                print('Município fora da área de concessão da RGE:',municipio, 'Linha:',index+1)
                
                lista_distribuidora.append(distribuidora)
                lista_arquivo.append(arquivo.split('\\')[-1])
                lista_erro.append('Municipio')
                lista_codigo.append(municipio)
                lista_linha.append(index+1)
    
# Número
def valida_numero(qtde_recebidas,qtde_procedentes,qtde_improcedentes,prazo_procedentes,prazo_improcedentes,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo):
    # Quantidade Recebidas
    # Número inteiro
    for index,qtde in enumerate(qtde_recebidas):
        try:
            int(qtde)
            
        except:
            print('Valor não é inteiro:',qtde, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Número Inteiro')
            lista_codigo.append(qtde)
            lista_linha.append(index+1)
            
            
    # Quantidade Procedentes
    # Número inteiro
    for index,qtde in enumerate(qtde_procedentes):
        try:
            int(qtde)
            
        except:
            print('Valor não é inteiro:',qtde, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Número Inteiro')
            lista_codigo.append(qtde)
            lista_linha.append(index+1)
            
            
    # Quantidade Improcedentes
    # Número inteiro
    for index,qtde in enumerate(qtde_improcedentes):
        try:
            int(qtde)
            
        except:
            print('Valor não é inteiro:',qtde, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Número Inteiro')
            lista_codigo.append(qtde)
            lista_linha.append(index+1)
            
            
    # Prazo Procedente
    # Número decimal
    for index,qtde in enumerate(prazo_procedentes):
        qtde = qtde.replace(',','.')
        # Checa se o número possui 2 casas decimais
        if Decimal(qtde).as_tuple().exponent == -2:
            continue
            
        else:
            print('Valor não tem 2 casas decimais:',qtde, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Número Decimal Prazo Procedente')
            lista_codigo.append(qtde)
            lista_linha.append(index+1)
            
            
    # Prazo Improcedente
    # Número decimal
    for index,qtde in enumerate(prazo_improcedentes):
        qtde = qtde.replace(',','.')
        # Checa se o número possui 2 casas decimais
        if Decimal(qtde).as_tuple().exponent == -2:
            continue
            
        else:
            print('Valor não tem 2 casas decimais:',qtde, 'Linha:',index+1)
            
            lista_distribuidora.append(distribuidora)
            lista_arquivo.append(arquivo.split('\\')[-1])
            lista_erro.append('Número Decimal Prazo Improcedente')
            lista_codigo.append(qtde)
            lista_linha.append(index+1)


#%% Roda as funções
distribuidoras = ['Paulista','Piratininga','Santa Cruz','RGE']
for distribuidora in distribuidoras:
    print(distribuidora)
    arquivos_xml = glob.glob(fr'C:\Users\2018459\OneDrive - CPFL Energia S A\Documentos\Projetos\2023\Anexo 1 - Tratamento das Reclamações\Dados\2023\{distribuidora}\Envio\*')

    for arquivo in arquivos_xml:
        print(arquivo)
        
        # Leitura do arquivo XML
        file = minidom.parse(arquivo)
        
        # Executa todas as funções
        separa_tag(file)
        seleciona_elementos(municipio_tag, tipologia_tag, qtde_recebidas_tag, qtde_procedentes_tag, qtde_improcedentes_tag, prazo_procedentes_tag, prazo_improcedentes_tag,distribuidora,arquivo,lista_distribuidora,lista_arquivo,lista_erro,lista_codigo,lista_linha)
        valida_tipologia(codigo_tipologia,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo)
        valida_forma_contato(codigo_forma_contato,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo)
        valida_municipio(municipios,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo)
        valida_numero(qtde_recebidas,qtde_procedentes,qtde_improcedentes,prazo_procedentes,prazo_improcedentes,distribuidora,lista_erro,lista_codigo,lista_linha,lista_arquivo,arquivo)
        
        # Limpa as listas com os elementos 
        municipios = []
        codigo_forma_contato = []
        codigo_tipologia = []
        qtde_recebidas = []
        qtde_procedentes = []
        qtde_improcedentes = []
        prazo_procedentes = []
        prazo_improcedentes = []
        
        print('\n')



# Monta o dataframe com o log de erros
df_log_erros['distribuidora'] = pd.Series(lista_distribuidora)
df_log_erros['arquivo'] = pd.Series(lista_arquivo)
df_log_erros['erro'] = pd.Series(lista_erro)
df_log_erros['valor'] = pd.Series(lista_codigo)
df_log_erros['linha'] = pd.Series(lista_linha)

# Exportação do log de erros
df_log_erros.to_csv(r'C:\Users\2018459\OneDrive - CPFL Energia S A\Documentos\Projetos\2023\Anexo 1 - Tratamento das Reclamações\Dados\2023\Log de Erros\log_erros.csv',sep=';',decimal=',',index=False,encoding='ANSI')


