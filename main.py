import pandas as pd
import win32com.client as win32

#importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualizar a base de dados
pd.set_option('display.max_columns', None)


# faturamento por loja

faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()



# quantidade de produtos vendidos por loja

quantidade_produtos = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()


# ticket medio por produto em cada loja

ticket_medio = (faturamento['Valor Final'] / quantidade_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})


# enviar o email com um relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gabriel1902008@hotmail.com'
mail.Subject = 'Relatorio de Vendas'
mail.HTMLBody = f"""
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade_produtos.to_html()}

<p>Ticket médio de produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Att,Guilherme<p>
"""

mail.Send()

print('Email enviado')