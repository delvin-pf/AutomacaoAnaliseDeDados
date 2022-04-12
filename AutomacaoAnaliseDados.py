# Automação de procesos
# Analise de dados e envio de e-mails para gerentes de lojas e diretoria

import pandas as pd
import pathlib
import yagmail

emails_df = pd.read_excel('Bases de Dados\Emails.xlsx')
lojas_df = pd.read_csv('Bases de Dados\Lojas.csv', sep=';', encoding = 'latin-1')
vendas_df = pd.read_excel('Bases de Dados\Vendas.xlsx')
vendas_df = vendas_df.merge(lojas_df, on='ID Loja')

# Criando valor em dicionario para cada loja
dic = {}
for loja in lojas_df['Loja']:
    dic[loja] = vendas_df.loc[vendas_df['Loja']==loja]

# Definir o dia - data
data = vendas_df['Data'].max()

# Criando backup 
caminho = pathlib.Path(r'Backup Arquivos Lojas')
pastasExistentes = caminho.iterdir()
listaPastas = [pasta.name for pasta in pastasExistentes]

for loja in dic:
    if loja not in listaPastas:
        (caminho/loja).mkdir()
    nmArquivo = f'{data.day}_{data.month}_{loja}.xlsx'
    localArquivo = caminho/loja/nmArquivo
    dic[loja].to_excel(localArquivo) 

# Definir valor dos indicadores
metaFatDia = 1000
metaFatAno = 1650000
metaDiverDia = 4
metaDiverAno = 120
metaTicketMedio = 500


for loja in dic:
    # Analise de dados e envio de e-mail para cada gerente
    faturamentoAno = dic[loja]['Valor Final'].sum()
    faturamentoDia = dic[loja].loc[dic[loja]['Data']==data, 'Valor Final'].sum()
    diversidadeAno = len(dic[loja]['Produto'].unique())
    diversidadeDia = len(dic[loja].loc[dic[loja]['Data']==data, 'Produto'].unique())

    valorPorVenda = dic[loja].groupby('Código Venda').sum()
    ticketMedioAno = valorPorVenda['Valor Final'].mean()
    valorPorVenda = dic[loja][dic[loja]['Data']==data].groupby('Código Venda').sum()
    ticketMedioDia = valorPorVenda['Valor Final'].mean()

    nmGerente = emails_df.loc[emails_df['Loja']==loja, 'Gerente'].values[0]

    listaCor = ['red', 'red', 'red', 'red', 'red', 'red']
    listaCor[0] = 'green' if faturamentoDia >= metaFatDia else 'red'
    listaCor[1] = 'green' if diversidadeDia >= metaDiverDia else 'red'
    listaCor[2] = 'green' if ticketMedioDia >= metaTicketMedio else 'red'
    listaCor[3] = 'green' if faturamentoAno >= metaFatAno else 'red'
    listaCor[4] = 'green' if diversidadeAno >= metaDiverAno else 'red'
    listaCor[5] = 'green' if ticketMedioAno >= metaTicketMedio else 'red'

    corpo = f'''
    <p>Bom dia, {nmGerente} </p>
    <p>O resultado de ontem <strong>({data.day}/{data.month})</strong> da loja <strong>{loja}</strong> foi de:</p>

    <table style="width:60%; border:1px solid #dddddd; border-collapse: collapse;">
      <tr style="border: 1px solid #dddddd;">
        <th style="text-align: center">Indicador</th>
        <th style="text-align: center">Valor Dia</th>
        <th style="text-align: center">Meta Dia</th>
        <th style="text-align: center">Cenario Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamentoDia:,.2f}</td>
        <td style="text-align: center">R${metaFatDia:,.2f}</td>
        <td style="text-align: center"><font color={listaCor[0]}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{diversidadeDia}</td>
        <td style="text-align: center">{metaDiverDia}</td>
        <td style="text-align: center"><font color={listaCor[1]}>◙</font></td>
      </tr>
      <tr>
       <td>Ticket Medio</td>
        <td style="text-align: center">R${ticketMedioDia:,.2f}</td>
        <td style="text-align: center">R${metaTicketMedio:,.2f}</td>
        <td style="text-align: center"><font color={listaCor[2]}>◙</font></td>
      </tr>
    </table >
    <br>
    <table style="width:60%; border:1px solid #dddddd; border-collapse: collapse;">
      <tr style="border: 1px solid #dddddd;">
        <th style="text-align: center">Indicador</th>
        <th style="text-align: center">Valor Ano</th>
        <th style="text-align: center">Meta Ano</th>
        <th style="text-align: center">Cenario Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamentoAno:,.2f}</td>
        <td style="text-align: center">R${metaFatAno:,.2f}</td>
        <td style="text-align: center"><font color={listaCor[3]}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{diversidadeAno}</td>
        <td style="text-align: center">{metaDiverAno}</td>
        <td style="text-align: center"><font color={listaCor[4]}>◙</font></td>
      </tr style="border: 1px solid #dddddd;">
      <tr>
       <td>Ticket Medio</td>
        <td style="text-align: center">R${ticketMedioAno:,.2f}</td>
        <td style="text-align: center">RS{metaTicketMedio:,.2f}</td>
        <td style="text-align: center"><font color={listaCor[5]}>◙</font></td>
      </tr>
    </table>
    <p>Segue em anexo a planilhas com todos os dados para mais detalhes</p>
    <p>Att, Delvin </p>
    '''
    with open('login.txt', 'r')  as login: # Arquivo com e-mail e senha.
        email, senha = login.readlines()
    
    email = yagmail.SMTP(email, senha)
    corpo = corpo.replace("\n", "")
    email.send(
        to = emails_df.loc[emails_df['Loja']==loja, 'E-mail'].values[0],
        subject = f'OnePage Dia {data.day}/{data.month} - Loja {loja}',
        contents = corpo,
        attachments = str(pathlib.Path.cwd() / caminho / loja / f'{data.day}_{data.month}_{loja}.xlsx')
       )
    print(f'E-mail da loja {loja} enviado com suceso ')


#Criar ranking 
fatLojasAno_df = vendas_df.groupby('Loja')[['Loja', 'Valor Final']].sum()
fatLojasAno_df = fatLojasAno_df.sort_values(by='Valor Final', ascending=False)
fatLojasAno_df.to_excel(f'Backup Arquivos Lojas\\{data.day}_{data.month}_Ranking Anual.xlsx')

fatLojasDia_df = vendas_df[vendas_df['Data']==data].groupby('Loja')[['Loja', 'Valor Final']].sum()
fatLojasDia_df = fatLojasDia_df.sort_values(by='Valor Final', ascending=False)
fatLojasDia_df.to_excel(f'Backup Arquivos Lojas\\{data.day}_{data.month}_Ranking Diario.xlsx')

# Criar e enviar email pra Diretoria
corpo = f'''
Prezados, bom dia

Melhor loja do dia: {fatLojasDia_df.index[0]}, com faturamento R${fatLojasDia_df.iloc[0, 0]:,.2f}
Pior loja do dia: {fatLojasDia_df.index[-1]}, com faturamento R${fatLojasDia_df.iloc[-1, 0]:,.2f}

Melhor loja do ano: {fatLojasAno_df.index[0]}, faturamento R${fatLojasAno_df.iloc[0, 0]:,.2f}
Pior loja do ano: {fatLojasAno_df.index[-1]}, faturamento R${fatLojasAno_df.iloc[-1, 0]:,.2f}

Segue em anexo os rankings do dia e do ano de todas as lojas

Att.,

Python - Analise de Dados
'''
email.send(
    to = emails_df.loc[emails_df['Loja']=='Diretoria', 'E-mail'].values[0],
    subject = f'Ranking dia {data.day}/{data.month} - Todas as lojas',
    contents = corpo,
    attachments = [str(pathlib.Path.cwd() / caminho / f'{data.day}_{data.month}_Ranking Anual.xlsx'), 
                   str(pathlib.Path.cwd() / caminho / f'{data.day}_{data.month}_Ranking Diario.xlsx')]
)

print("Email da Diretoria Enviado")
