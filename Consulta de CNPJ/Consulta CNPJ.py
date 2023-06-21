import pandas as pd
import requests
import json
from time import sleep as slp

df = pd.read_excel(input(r'Informe o caminho do arquivo: '), dtype=str)

print('Consultando...')

with open('ConsultaCNPJ.csv' , 'a', newline= '', encoding='ISO-8859-1') as arquivo:
    
    for c in df['CNPJ']:
        
        browser = requests.get('https://www.receitaws.com.br/v1/cnpj/' + c)
        slp(3)
        
        resp = json.loads(browser.text)
        
        cep = resp['cep']
        nome = resp['nome']
        logradouro = resp['logradouro']
        numero = resp['numero']
        tipo = resp['tipo']
        email = resp['email']
        situacao = resp['situacao']
        bairro = resp['bairro']
        municipio = resp['municipio']
        uf = resp['uf']
        #abertura = resp['abertura']
        #atividade = resp['atividade_principal']
        #status = resp['status']
        #capital = resp['capital_social']
        
        arquivo.write(c + '_' + str(nome) + '_' + str(tipo) + '_' + str(cep) + '_' + str(logradouro) + '_' + str(numero) + 
                      '_' + str(bairro) + '_' + str(municipio) + '_' + str(uf) + '_' + str(situacao) + '_' + str(email))
        #               + '_' + str(abertura) + '_' + str(atividade) + '_' + str(status) + '_' + str(capital))
        
        arquivo.write(str('\n'))
        slp(20)

print('Finalizado.')