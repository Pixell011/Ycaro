import imaplib
import email
from email.header import decode_header
import os
import pandas as pd
import spacy

# Configurar NLP
nlp = spacy.load("pt_core_news_sm")

# Função para categorizar o assunto do e-mail
def categorize_subject(subject):
    doc = nlp(subject)
    for token in doc:
        if token.text.lower() in ["fatura", "conta"]:
            return "Financeiro"
        elif token.text.lower() in ["relatório", "resumo"]:
            return "Relatórios"
    return "Outros"

# Função para salvar anexos
def save_attachment(msg, download_folder="anexos"):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            filePath = os.path.join(download_folder, fileName)
            if not os.path.isfile(filePath):
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
            print(f"Salvo {fileName} em {download_folder}")
            return filePath
    return None

# Configurar conexão com o e-mail
username = "gasilva@itapevarec.com.br"
password = "Bad43694"

# Criar pasta para salvar anexos
if not os.path.isdir("anexos"):
    os.makedirs("anexos")

# Conectar ao servidor IMAP do Outlook usando SSL
mail = imaplib.IMAP4_SSL("outlook.office365.com")
mail.login(username, password)
mail.select("inbox")

# Buscar e-mails não lidos com anexos PDF
status, messages = mail.search(None, '(UNSEEN)')
email_ids = messages[0].split()

email_data = []

for e_id in email_ids:
    res, msg = mail.fetch(e_id, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode()
            sender = msg.get("From")
            date = msg.get("Date")
            
            # Salvar anexos
            attachment_path = save_attachment(msg)
            
            # Categorizar e-mail
            category = categorize_subject(subject)
            
            # Adicionar dados à lista
            email_data.append([subject, sender, date, category, attachment_path])

# Exportar dados para Excel
df = pd.DataFrame(email_data, columns=["Assunto", "Remetente", "Data", "Categoria", "Caminho do Anexo"])
df.to_excel("emails_categorizados.xlsx", index=False)

print("Dados exportados para emails_categorizados.xlsx")

# Logout do e-mail
mail.logout()
