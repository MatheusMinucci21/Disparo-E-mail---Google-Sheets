import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime, timedelta

# Define o escopo das credenciais e cria uma instância de cliente para interagir com o Google Sheets
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    'CAMINHO DO ARQUIVO json', scope)
client = gspread.authorize(credentials)

spreadsheet_url = 'URL DA PLANILHA SHEETS'
spreadsheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]
spreadsheet = client.open_by_key(spreadsheet_id)

worksheet = spreadsheet.worksheet('ABA DA PLANILHA')

data = worksheet.get_all_values()

# Transforma a lista de listas em um DataFrame do Pandas
df = pd.DataFrame(data[1:], columns=data[0]) 

# Selecionando Colunas desejadas
df = df[['ID', 'NF', 'VOL DOS PEDIDOS', 'DATA DE EMISSÃO NF', 'VALOR DA NOTA', 'CNPJ', 'FORNECEDOR', 'Enviar o E-mail?', 'E-mail Fornecedor']]

# Filtrando o que quero enviar
df_filtro = df.loc[df['Enviar o E-mail?'] == 'Enviar', :]

# Agrupando e-mails por fornecedor
emails_por_fornecedor = df_filtro.groupby('FORNECEDOR').agg(
    emails=('E-mail Fornecedor', list),
    nfs=('NF', list),
    valores=('VALOR DA NOTA', list),
)

# Calculando data limite
data_limite = (datetime.now() + timedelta(days=4)).strftime("%d/%m/%Y")

# Loop para enviar e-mails
for fornecedor, info in emails_por_fornecedor.iterrows():
    emails = info['emails']
    nfs = info['nfs']
    valores = info['valores']
    
    Corpo_email = (
        f"Boa Tarde!\n\n"
        f"Prezados {fornecedor},\n\n"
        f"Segue dados em anexo com a(s) NF(s) {', '.join(nfs)}, no(s) valor(es) de {', '.join(valores)} a serem coletadas aqui no CD 1 Extrema-MG, favor coletar as mesmas até {data_limite}.\n\n"
        f"Peço que nos informem quando a coleta será realizada, e os seguintes dados:\n\n"
        f"Informações pessoais do motorista:\n"
        f"Nome completo:\n"
        f"RG:\n\n"
        f"Informações do veículo:\n"
        f"Modelo:\n"
        f"Ano:\n"
        f"Placa:\n"
        f"Carreta:\n\n"
        f"Por gentileza, enviar junto ao motorista uma ordem de coleta para agilizar o processo.\n\n"
        f"Lembrando nosso horário de coleta:\n"
        f"SEG a QUI: 08h30 às 17:00\n"
        f"Sexta: 08h30 às 16:00\n\n"
        f"Informo que é uma mensagem automática e será cessada quando houver retorno.\n"
    )
    
    # Cria um objeto SMTP
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        # Conecta ao servidor SMTP
        smtp.starttls()  # Habilita SSL/TLS
        smtp.login('E-MAIL SERVER', 'SENHA SERVER TEMPORARIA')
        
        # Cria um objeto MIMEMultipart para o email
        for email in emails:
            msg = MIMEMultipart()
            msg['From'] = 'EMAIL DO DISPARO'
            msg['To'] = email
            msg['Subject'] = f"Devolução Fornecedor NF{', '.join(nfs)} fornecedor {fornecedor} TESTE!!!"
            msg.attach(MIMEText(Corpo_email, 'plain'))
            
            # Envia o email
            smtp.sendmail(msg['From'], email, msg.as_string())
            print(f"Enviando email para: {email}")

