# instalar pandas
import pandas as pd
# instalar openpyxl (pandas e openpyxl servem para a integração com os arquivos excel funcionar)

# instalar twilio (integração do python com sms)
from twilio.rest import Client

# Your Account SID from twilio.com/console
account_sid = "AC1692bd35ebd1aaac61ab89ec520181a2"
# Your Auth Token from twilio.com/console
auth_token  = "7febc77fb0095b33a78d738fdbc52015"
client = Client(account_sid, auth_token)

# abrir os 6 arquivos em excel
lista_meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho']

# para cada arquivo:
for mes in lista_meses:
    tabela_vendas = pd.read_excel(f'{mes}.xlsx')
    # verificar se algum valor na coluna Vendas daquele arquivo é maior que 55.000
    if (tabela_vendas['Vendas'] > 55000).any(): # se for maior que 55.000:
        vendedor = tabela_vendas.loc[tabela_vendas['Vendas'] > 55000, 'Vendedor'].values[0]
        vendas = tabela_vendas.loc[tabela_vendas['Vendas'] > 55000, 'Vendas'].values[0]
        #print(f'Encontrei alguem no mes de {mes}. Vendedor: {vendedor}, Vendas: {vendas}')
        message = client.messages.create( # envia um sms com o nome, o mês e as vendas do vendedor
            to='+5586999188680', 
            from_='+17576371080',
            body=f'Encontrei alguem no mes de {mes}. Vendedor: {vendedor}, Vendas: {vendas}')
        print(message.sid)
        