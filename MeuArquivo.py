import pandas as pd
import win32com.client as win32


# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizar a base de dados
pd.set_option('display.max_columns' , None)
print(tabela_vendas)

# Faturamento loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja 
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Enviar um email com o relatório 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'iagooreis27@hotmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Lira</p>
'''

mail.Send()

print('Email Enviado')