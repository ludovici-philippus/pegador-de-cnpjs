import pandas as pd
import openpyxl
import re
from unicodedata import normalize
from urllib.request import urlopen
import json
from brutils import remove_symbols_cnpj
from time import sleep
from datetime import datetime

data = []

CONTACTED_CNPJS = set([

])

CNPJS_TO_CONTACT = set([

]) - CONTACTED_CNPJS

def send_data_to_excel(data):
    today = datetime.today().strftime('%Y-%m-%d')
    df = pd.DataFrame(data, columns=[
                        'CNPJ', 
                        'Nome Fantasia', 
                        'Razão Social',
                        'CNAE', 
                        'Sócios',
                        'Endereço',
                        'E-mail',
                        'Telefone'])

    df.to_excel(f"companies-{today}.xlsx", sheet_name='new_sheet_name')

def get_stockholder_names(stockholders):
    names = []
    for stockholder in stockholders:
        names.append(f"{stockholder['nome']} - {stockholder['qualificacao_socio']['descricao']}")
    return names

def get_location(business):
    return f"{business['tipo_logradouro']} {business['logradouro']} {business['numero']} - {business['bairro']} ({business['cep']})"

def get_data_from_api(cnpjs):
    API = 'https://publica.cnpj.ws/cnpj'
    data_to_be_appended = []
    iteration = 1
    for cnpj in cnpjs:
        if iteration % 3 == 0: sleep(60)
        unformated_cnpj = remove_symbols_cnpj(cnpj)
        result = urlopen(f"{API}/{unformated_cnpj}")
        result = json.load(result)
        data_to_be_appended = [
            result['estabelecimento']['cnpj'], 
            result['estabelecimento']['nome_fantasia'], 
            result['razao_social'], 
            result['estabelecimento']['atividade_principal']['subclasse'],
            get_stockholder_names(result['socios']),
            get_location(result['estabelecimento']),
            result['estabelecimento']['email'],
            f"{result['estabelecimento']['ddd1']} {result['estabelecimento']['telefone1']}"
        ]
        data.append(data_to_be_appended)
        iteration += 1

API = 'https://publica.cnpj.ws/cnpj/'
get_data_from_api(CNPJS_TO_CONTACT)

send_data_to_excel(data)
print(data)
