def contrato_para_data_vencimento(dados_com_contrato: dict):
    '''
    Função para converter os contratos de forma a serem mais fáceis
    de compreender sua data de vencimento.

    Args: 
        dados_com_contrato : (dict) - dicionário com dados extraidos

    Return:
        dados_com_contrato : (dict) - dicionario com dados modificados

    
    '''
    dicionario_vencimento = {
        'F': 'Jan',
        'G': 'Fev',
        'H': 'Mar',
        'J': 'Abr',
        'K': 'Mai',
        'M': 'Jun',
        'N': 'Jul',
        'Q': 'Ago',
        'U': 'Set',
        'V': 'Out',
        'X': 'Nov',
        'Z': 'Dez'
    }
    for contrato, dados_extraidos in dados_com_contrato.items():
        if contrato == 'VENCTO \xa0' or contrato == 'Contrato':
            dados_com_contrato[contrato] = f'{dicionario_vencimento[dados_extraidos[-3]]}-20{dados_extraidos[-2:]}'

    return dados_com_contrato
