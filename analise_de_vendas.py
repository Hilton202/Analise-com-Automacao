import pandas as pd
import win32com.client as win32

try:
    # Importing the sales data
    dados_vendas = pd.read_excel('vendas_lojas.xlsx')

    # Faturamento por Loja
    faturamento_loja = dados_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

    # Quantidade de produtos vendidos por Loja
    produto_vendido = dados_vendas[['ID Loja', 'Quantidade', 'Valor Final']].groupby('ID Loja').sum()

    # Ticket médio por produto em cada loja
    ticket_medio = (faturamento_loja['Valor Final'] / produto_vendido['Quantidade']).to_frame()

    # Send an email with the report
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ''  # coloque o email aqui 
    mail.Subject = 'Relatório de Vendas por Loja'
    mail.HTMLBody = f''' 
        <p>Prezados,</p>

        <p>Segue o Relatório de Vendas por cada Loja.</p>

        <p>Faturamento:</p>
        {faturamento_loja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

        <p>Quantidade Vendida:</p>
        {produto_vendido.to_html()}

        <p>Ticket Médio dos Produtos em cada Loja:</p>
        {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

        <p>Qualquer dúvida estou à disposição.</p>

        <p>Att.,</p>
        <p>Lira</p> 
        '''
    mail.Send()

    print('Email Enviado')

except Exception as e:
    print(f"An error occurred: {e}")
