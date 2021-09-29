import pandas as pd
import win32com.client as win32

# importa a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#visualiza a base de dados
pd.set_option('display.max_columns',None)


# Faturamento por loja
Faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()


# Quantidade de produtos vendidos por lojas
Quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()


# ticket médio pro produto vendidos por lojas
ticket_medio = (Faturamento['Valor Final'] / Quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'ticket Medio'})

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Cashofcla160@gmail.com'
mail.Subject ='Relatório de vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p> Segue o Relatório de vendas por cada Loja .</p>

<p>Faturamento:</p>
 {Faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<h2>Quantidade Vendida:</h2>
{Quantidade.to_html(formatters={'Quantidade': 'R${:,.2f}'.format})}

<h2>Ticket Médio dos Produtos em cada Loja:</h2>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}
<p>Qualquer Dúvida fico a Diposição</p>
<p>Att..,</p>
<p>Luiz Henrique</p>  
'''
mail.send()

