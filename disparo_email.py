import os

import base64
import re
from datetime import datetime

import msal
import pandas as pd
import requests

from dotenv import load_dotenv

load_dotenv()

# ============================================================
# CONFIGURAÇÕES
# ============================================================

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_TESTE = os.getenv("EMAIL_TESTE")
EMAIL_COMERCIAL = os.getenv("EMAIL_COMERCIAL")

ARQUIVO_EXCEL = os.getenv("ARQUIVO_EXCEL")
PLANILHA = os.getenv("PLANILHA")
IMAGEM_PROMO = os.getenv("IMAGEM_PROMO")
IMAGEM_ASSINATURA = os.getenv("IMAGEM_ASSINATURA")
TEMPLATE_EMAIL = os.getenv("TEMPLATE_EMAIL")

#Validação de variáveis de ambiente
if not CLIENT_ID or not TENANT_ID or not EMAIL_TESTE:
    raise RuntimeError(
        "CLIENT_ID, TENANT_ID e EMAIL_TESTE não foram configurados no arquivo .env."
    )

if not ARQUIVO_EXCEL or not PLANILHA or not IMAGEM_PROMO or not IMAGEM_ASSINATURA:
    raise RuntimeError(
        "Alguma das variáveis de ambiente para os arquivos da campanha não foi configurada no arquivo .env."
    )

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send"]

URL_ENVIO = "https://graph.microsoft.com/v1.0/me/sendMail"


# ============================================================
# ASSUNTO
# ============================================================

ASSUNTO = "Uma Nova forma de levar sabor aos seus produtos"


# ============================================================
# CORPO DO E-MAIL
# ============================================================
with open(TEMPLATE_EMAIL, "r", encoding="utf-8") as arquivo:
    HTML_EMAIL = arquivo.read()

HTML_EMAIL = HTML_EMAIL.replace(
    "{EMAIL_COMERCIAL}",
    EMAIL_COMERCIAL
)
# ============================================================
# VALIDAÇÃO DE E-MAIL
# ============================================================

def email_valido(email):
    padrao = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    return bool(re.match(padrao, email))


# ============================================================
# LER EXCEL
# ============================================================

print("\nLendo a planilha...")

df = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=PLANILHA
)

if "Email" not in df.columns:
    raise Exception(
        "A coluna 'Email' não foi encontrada na planilha Base Disparo."
    )


emails = (
    df["Email"]
    .dropna()
    .astype(str)
    .str.strip()
    .str.lower()
)

emails = emails[emails != ""]
emails = emails[emails.apply(email_valido)]
emails = emails.drop_duplicates().tolist()


print(f"\nTotal de registros na Base Disparo: {len(df)}")
print(f"E-mails válidos e únicos para envio: {len(emails)}")

print("\nPrimeiros destinatários:")

for email in emails[:10]:
    print(f"  - {email}")

if len(emails) > 10:
    print(f"  ... e mais {len(emails) - 10}")

# ============================================================
# CARREGAR IMAGENS
# ============================================================

print("\nCarregando imagens...")

with open(IMAGEM_PROMO, "rb") as arquivo:
    imagem_promocao_base64 = base64.b64encode(
        arquivo.read()
    ).decode("utf-8")


with open(IMAGEM_ASSINATURA, "rb") as arquivo:
    imagem_assinatura_base64 = base64.b64encode(
        arquivo.read()
    ).decode("utf-8")


# ============================================================
# AUTENTICAÇÃO MICROSOFT
# ============================================================

print("\nAutenticando na Microsoft...")

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

contas = app.get_accounts()

resultado = None

if contas:
    resultado = app.acquire_token_silent(
        SCOPES,
        account=contas[0]
    )

if not resultado:
    resultado = app.acquire_token_interactive(
        scopes=SCOPES
    )

if "access_token" not in resultado:
    print("\nErro ao obter o token:")
    print(resultado.get("error"))
    print(resultado.get("error_description"))
    raise SystemExit


access_token = resultado["access_token"]

print("Login realizado com sucesso!")

# ============================================================
# CABEÇALHOS
# ============================================================

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}
def enviar_email(destinatario):
    dados = {
        "message": {
            "subject": ASSUNTO,

            "body": {
                "contentType": "HTML",
                "content": HTML_EMAIL
            },

            "toRecipients": [
                {
                    "emailAddress": {
                        "address": destinatario
                    }
                }
            ],

            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": "Fumaca-em-Po-Promo.jpeg",
                    "contentType": "image/jpeg",
                    "contentBytes": imagem_promocao_base64,
                    "isInline": True,
                    "contentId": "imagem_promocao"
                },
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": "Assinatura-Gustavo.png",
                    "contentType": "image/png",
                    "contentBytes": imagem_assinatura_base64,
                    "isInline": True,
                    "contentId": "imagem_assinatura"
                }
            ]
        },

        "saveToSentItems": True
    }

    return requests.post(
        URL_ENVIO,
        headers=headers,
        json=dados,
        timeout=30
    )

# ============================================================
# TESTE EXCLUSIVO
# ============================================================

print("\n" + "=" * 60)
print("ENVIANDO E-MAIL DE TESTE")
print("=" * 60)

print(f"\nDestinatário do teste: {EMAIL_TESTE}")

resposta_teste = enviar_email(EMAIL_TESTE)

if resposta_teste.status_code == 202:

    print("\nTESTE ENVIADO COM SUCESSO!")
    print("Confira o e-mail recebido antes de continuar.")

else:

    print("\nERRO NO E-MAIL DE TESTE.")
    print("Status:", resposta_teste.status_code)
    print("Resposta:", resposta_teste.text)

    raise SystemExit
confirmacao = input(
    "\nDigite ENVIAR para iniciar o disparo: "
).strip().upper()

if confirmacao != "ENVIAR":
    print("\nDisparo cancelado.")
    raise SystemExit

# ============================================================
# DISPARO
# ============================================================

enviados = []
erros = []

total = len(emails)

print("\nIniciando disparo...\n")


for indice, email in enumerate(emails, start=1):

    try:

        resposta = enviar_email(email)

        if resposta.status_code == 202:

            enviados.append(email)

            print(
                f"[{indice}/{total}] OK - {email}"
            )

        else:

            erros.append({
                "email": email,
                "status": resposta.status_code,
                "erro": resposta.text
            })

            print(
                f"[{indice}/{total}] ERRO - {email} "
                f"({resposta.status_code})"
            )

    except Exception as erro:

        erros.append({
            "email": email,
            "status": "EXCEPTION",
            "erro": str(erro)
        })

        print(
            f"[{indice}/{total}] ERRO - {email} "
            f"({erro})"
        )

# ============================================================
# LOG
# ============================================================

data_hora = datetime.now().strftime(
    "%Y-%m-%d_%H-%M-%S"
)

arquivo_log = f"log_disparo_{data_hora}.csv"

linhas_log = []


for email in enviados:

    linhas_log.append({
        "email": email,
        "status": "Enviado",
        "data_hora": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    })


for erro in erros:

    linhas_log.append({
        "email": erro["email"],
        "status": f"Erro {erro['status']}",
        "data_hora": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    })


pd.DataFrame(linhas_log).to_csv(
    arquivo_log,
    index=False,
    encoding="utf-8-sig"
)


# ============================================================
# RESUMO
# ============================================================

print("\n" + "=" * 60)
print("DISPARO FINALIZADO")
print("=" * 60)

print(f"Total de destinatários: {total}")
print(f"Enviados com sucesso: {len(enviados)}")
print(f"Com erro: {len(erros)}")

print(f"\nLog salvo em:")
print(arquivo_log)