#Importação de Blibliotecas
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe
from dotenv import load_dotenv
import os
import smtplib
from email.message import EmailMessage
import json
#Conexão API google Sheets
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds_dict = json.loads(st.secrets['GOOGLE_CREDENTIALS'])
creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
client = gspread.authorize(creds)

#Puxando variáveis de ambiente
EMAILS_ID = st.secrets['EMAILS_ID']
sheet = client.open_by_key(EMAILS_ID).sheet1

# Ler os dados
df_emails = get_as_dataframe(sheet, evaluate_formulas=True)

st.set_page_config(layout='wide',page_icon='news')
st.title('Newsletter Comissão T-38')
mensagem_basica = st.text_input('Insira Aqui a Mensagem que deseja enviar')
send_button = st.button('Enviar emails')

if send_button and mensagem_basica:
    email_remetente = 'pedro.ponte.9126@ga.ita.br'
    senha_app = st.secrets['SENHA_APP']

    for _,linha in df_emails.iterrows():
        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(email_remetente,senha_app)
                msg = EmailMessage()
                msg['Subject'] = 'Newsletter Comissão T38'
                msg['From'] = email_remetente
                msg['To'] = linha['Email']

                mensagem_completa = f'Olá {linha['Nome']}! \n\n {mensagem_basica}'
                msg.set_content(mensagem_completa)
                smtp.send_message(msg)
            st.success('Email_enviado!')
        except Exception as e:
            st.error(f"Erro ao enviar e-mail de {linha['Nome']}\n Erro: {e}")



