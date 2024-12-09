import numpy as np
import pandas as pd
import tempfile
import os

# Código criado por Milton Rocha
# https://medium.com/@milton-rocha

def feriados(override: bool = False) -> np.ndarray:
    """
    Função que faz o download dos feriados do site da Anbima e os compila em um formato a ser utilizado

      Resposta:
        np.array(feriados, dtype = 'datetime64[D]')
    """

    # O arquivo fornecido por download possui algumas marcas de fornecimento dos feriados que serão tratadas
    # O arquivo finaliza antes da linha que possui "Fonte: ANBIMA" como valor
    arq_temp = f'{tempfile.gettempdir()}/fer_anbima.parquet'

    # Checa se já existe um arquivo de feriados Anbima gerado na pasta temporária, caso não, cria o arquivo
    if not os.path.isfile(arq_temp) and not override:
        feriados = pd.read_excel(r'https://www.anbima.com.br/feriados/arqs/feriados_nacionais.xls')
        feriados = feriados['Data'][
                   : feriados[feriados['Data'] == 'Fonte: ANBIMA'].index[0]].values  # Acha a linha de footer
        feriados = pd.DataFrame({'Feriados ANBIMA': feriados.astype('datetime64[D]')})  # Cria um dataframe com os dados
        feriados.to_parquet(arq_temp)  # Exporta o DataFrame para .parquet
    else:
        feriados = pd.read_parquet(arq_temp)  # Caso o arquivo já exista no diretório temporário, lê o .parquet

    # A função irá retornar um np.array contendo todas as datas de feriado disponíveis em formato 'datetime64[D]'
    return feriados.values.astype('datetime64[D]').flatten()


def networkdays(data_inicial, data_final) -> np.ndarray:
    """
    Função equivalente ao =NETWORKDAYS() do Excel

      Variáveis:
        data_inicial : str, list, np.ndarray, pd.Series
        	É a data ou o conjunto de datas que serão utilizadas para cálculo do início dos dias úteis

        data_final   : str, list, np.ndarray, pd.Series
        	É a data ou o conjunto de datas que serão utilizadas para cálculo do fim dos dias úteis

      Reposta:
      	du  : np.ndarray
			Resposta do número de dias úteis entre os conjuntos de datas fornecidos
    """
    # Primeiramente criam-se dois np.array para cada uma das variáveis
    data_inicial, data_final = np.array(data_inicial).flatten(), np.array(data_final).flatten()
    #  Caso exista diferença entre o tamanho das variáveis, irá trabalhar com a que tem maior tamanho (v)
    # e repetirá a data da outra variável len(v) vezes
    if data_inicial.shape[0] != data_final.shape[0]:
        if data_inicial.shape[0] == 1:
            data_inicial = np.repeat(data_inicial, data_final.shape[0])
        elif data_final.shape[0] == 1:
            data_final = np.repeat(data_final, data_inicial.shape[0])
        elif data_inicial.shape[0] > 1 and data_final.shape[0] > 1:
            # Caso específico do usuário fornecendo as duas variáveis com mais de 1 elemento mas sem tamanho igual
            raise ValueError(
                f'O código não aceita data_inicial com tamanho > 1 [{data_inicial.shape[0]} elementos] ao mesmo tempo que data_final tem tamanho > 1 [{data_final.shape[0]} elementos]')

    # As variáveis de datas serão apenas um np.array convertido para o formato de 'datetime64[D]' (formato dos feriados)
    data_inicial, data_final = np.array(np.array([data_inicial, data_final])).astype('datetime64[D]')

    # A resposta é a contagem de dias úteis considerando a função anteriormente feita (feriados)
    return np.busday_count(data_inicial, data_final, holidays=feriados())


# #Exemplos de cálculos de DU:
# print(networkdays('2024-11-01', '2025-03-30')[0])  # array([251])
# print(networkdays(['2022-01-01'], ['2022-12-31']) ) # array([251])
# networkdays(['2022-01-01', '2022-02-01'], ['2022-06-01', '2022-12-31'])  # array([103, 230])
# networkdays([f'2022-{i:02d}-01' for i in range(1, 12)],
#             ['2022-12-31'])  # array([251, 230, 211, 189, 170, 148, 127, 106,  83,  62,  42])
# networkdays([f'2021-{i:02d}-01' for i in range(1, 12)], [f'2022-{i:02d}-01' for i in range(1,
#                                                                                            12)])  # array([251, 252, 253, 252, 251, 252, 252, 251, 252, 252, 252])

# Exemplo de erro:
# networkdays(['2022-01-01', '2022-02-01', '2022-03-01'], ['2022-06-01', '2022-12-31'])
# ValueError: O código não aceita dataInicial com tamanho > 1 [3 elementos] ao mesmo tempo que dataFinal tem tamanho > 1 [2 elementos]
