# Importar base de dados
import pandas as pd
import win32com.client as win32


# Funções
def linha(tam=50):
    print('-' * tam)


def titulo(texto):
    linha()
    print(f'{texto:^50}')
    linha()


# Visualizar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)
titulo('TABELA DE VENDAS COMPLETA')
print(tabela_vendas)

# Visualizar Faturamento
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
titulo('TABELA DE FATURAMENTO DE CADA LOJA')
print(faturamento)

# Quantidade de produtos vendidos por Loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
titulo('TABELA DE QUANTIDADE DE PRODUTOS POR LOJA')
print(quantidade)

# Ticket médio de cada Loja
ticket_medio = faturamento['Valor Final']/ quantidade['Quantidade']
ticket_medio = ticket_medio.to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
titulo('TABELA DE TICKET MÉDIO DE CADA LOJA')
print(ticket_medio)
linha()

# Enviar e-mail com relatório de vendas de cada loja
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'joaon.c.jv@gmail.com'
email.Subject = 'Relatório de Vendas - Projeto Python'
email.HTMLBody = f'''

<p>Prezado cliente,</p>

<p>Segue o relatório com os dados seguimentado por loja.</p>

<p>Faturamento Total por Loja</p>

{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida por Loja</p>
{quantidade.to_html()}

<p>Ticket Médio por Loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

'''
email.Send()
print('e-mail enviado com sucesso!')