

# importar a base de dados
import pandas as pd
import win32com.client as win32
tabela = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# calcular faturamento total por loja
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# quantidade de produtos vendidos por loja
quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# ticket médio (faturamento/qtd produtos vendidos) em cada loja
ticket = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})

# enviar um e-mail com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'kyleoliveirasilva@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>
<p>Segue aqui o relatório de vendas por loja.</p><br>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de produtos vendidos por loja:<p>
{quantidade.to_html(formatters={'Quantidade': 'R${:,.2f}'.format})}

<p>Ticket médio por loja:<p>
{ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição</p>
<p>Att.,</p>
<p>Kayle</p>
'''

mail.Send()
