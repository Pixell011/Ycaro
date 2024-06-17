import imaplib
import email
from email.header import decode_header
import pandas as pd

# Conectar ao servidor IMAP
username = "gasilva@itapevarec.com.br"
password = ""
imap_url = "imap.outlook.com.br"

# Autenticar
mail = imaplib.IMAP4_SSL(imap_url)
mail.login(username, password)

# Selecionar a caixa de entrada
mail.select("inbox")
