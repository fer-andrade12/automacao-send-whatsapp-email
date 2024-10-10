import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib.parse
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do email
smtp_server = 'smtp-mail.outlook.com'
smtp_port = 587  # Para TLS
email_usuario = 'emailr@outlook.com'  # Insira seu email
email_senha = '####'  # Insira sua senha ou senha de aplicativo

# Carregar a planilha
planilha = pd.read_excel(r'C:\pasta\pasta1.xlsx')

# Iniciar o driver do Selenium
driver = webdriver.Chrome()  # Certifique-se de que o chromedriver está no PATH

# Função para enviar mensagem via WhatsApp
def enviar_mensagem_whatsapp(telefone, mensagem):
    print(f"Enviando mensagem para o número {telefone} via WhatsApp...")
    mensagem_codificada = urllib.parse.quote(mensagem)  # Codifica a mensagem para URL
    url = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem_codificada}"
    driver.get(url)

    # Aguardar que a página do WhatsApp Web carregue e o botão de enviar fique clicável
    try:
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]')))
        botao_enviar = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        botao_enviar.click()
        print("Mensagem enviada com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar mensagem via WhatsApp: {e}")

# Função para enviar email
def enviar_email(destinatario_email, mensagem):
    msg = MIMEMultipart()
    msg['From'] = email_usuario
    msg['To'] = destinatario_email
    msg['Subject'] = 'Mensagem Automática'

    # Adicionar o corpo da mensagem
    msg.attach(MIMEText(mensagem, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as servidor:
            servidor.starttls()  # Inicia a conexão segura
            servidor.login(email_usuario, email_senha)
            servidor.sendmail(email_usuario, destinatario_email, msg.as_string())
            print(f"Email enviado para {destinatario_email} com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar email para {destinatario_email}: {e}")

# Iterar sobre cada linha da planilha
for index, row in planilha.iterrows():
    nome = row['NOME']
    telefone = row['CONTATO']
    email = row['EMAIL']

    # Verificar e corrigir o formato do número de telefone
    if not telefone.startswith('+'):
        telefone = '+55' + telefone  # Adiciona o código do Brasil (+55) caso não tenha

    # Remover parênteses, espaços e traços
    telefone = telefone.replace('(', '').replace(')', '').replace(' ', '').replace('-', '')

    # Mensagem personalizada
    mensagem = f"Olá {nome}, esta recebendo uma mensagem automática da automatização.\n\n" \
               "Lista de cursos disponíveis:\n" \
               "📚 Curso de Python\n" \
               "💻 Curso de Java\n" \
               "📊 Curso de Excel\n" \
               "🎨 Curso de Design Gráfico\n" \
               "🔒 Curso de Segurança da Informação"

    # Enviar mensagem via WhatsApp
    enviar_mensagem_whatsapp(telefone, mensagem)
    
    # Enviar email
    enviar_email(email, mensagem)

    # Aguardar para evitar bloqueio de spam pelo WhatsApp
    time.sleep(10)  # 5 segundos entre cada mensagem para evitar bloqueio

# Fechar o driver do Selenium
driver.quit()
